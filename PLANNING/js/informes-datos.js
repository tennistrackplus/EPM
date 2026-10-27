/**
 * ============================================================
 * DRACO PLANNING — INFORMES · DATOS
 * (actualización de tablas dentro de un informe)
 * ============================================================
 * Un bloque "Actualización de tablas" de un informe muestra y permite
 * editar una tabla de parametrización con el MISMO comportamiento que la
 * pantalla de Actualización de tablas (js/table-updates.js), pero pintada
 * con los estilos del informe.
 *
 * La página del informe corre en un iframe aislado (sin sesión de Draco),
 * así que NO toca la base de datos: pide las operaciones por mensajes y este
 * módulo, que corre en generador.html / visor.html con la sesión del
 * usuario, las ejecuta:
 *
 *   tu.list      vistas de actualización del proyecto (para el editor)
 *   tu.describe  variables y campos de una vista (para el editor)
 *   tu.load      valida las variables, aplica el filtro y devuelve las filas
 *   tu.preview   valida y calcula altas/modificaciones/bajas SIN grabar
 *   tu.commit    valida, calcula y GRABA
 *   tu.export    descarga CSV / Excel
 *   tu.parse     lee un fichero CSV / Excel y devuelve sus filas
 *
 * Reutilización: en vez de copiar la lógica, se crea por cada vista una
 * "sesión" (Object.create(TableUpdates)) con su propio estado, y se llaman
 * los mismos métodos que usa la pantalla original (buildWhere,
 * resolveValidationOptions, computeDiff, getKeyColumns...). Ninguno toca el
 * DOM, y al ser una copia con estado propio no interfiere con la pantalla.
 *
 * Seguridad:
 *  - El informe NUNCA envía SQL, nombres de tabla ni filtros: solo la
 *    referencia a la vista (ACTUALIZACION_ID) y los valores de las variables.
 *    La tabla, las columnas y el WHERE se recalculan aquí desde la definición
 *    guardada de la vista y siempre dentro del proyecto del host.
 *  - Todo valor se escapa con Provider.esc.
 *
 * Grabar usa la misma estrategia que la pantalla original: borrar el
 * subconjunto filtrado y volver a insertarlo completo. Como esa estrategia
 * no es atómica, si el INSERT falla se intenta restaurar las filas
 * originales (la pantalla original no lo hace).
 *
 * Los errores llevan un `code`: NO_AUTH, NOT_FOUND, BAD_INPUT, BAD_VARS,
 * TOO_MANY, DUP_KEYS, EMPTY_NOT_ALLOWED, FIXED_ROWS, SAVE_FAILED, NOT_AVAILABLE.
 *
 * Filas fijas y "Completar con todos los valores": la lógica vive en TableUpdates
 * (completeRows, isFixedRows...) y se reutiliza tal cual; aquí solo se aplica al cargar
 * y se hace cumplir al grabar.
 */
const InformesDatos = {
    MAX_ROWS: 50000,
    MAX_OPTIONS: 3000,
    CHUNK: 500,
    META_TTL: 300000,        // la definición, dimensiones, claves y listas de valores se reutilizan 5 min (Recargar las renueva)
    sessions: {},
    opening: {},
    pending: {},             // tablas precargadas mientras se dibuja el informe (un solo uso, 60 s)

    fail(code, message, details) {
        const err = new Error(message);
        err.code = code;
        if (details) err.details = details;
        return err;
    },

    // ---------------------------------------------------------
    // Sesiones (una por vista de actualización)
    // ---------------------------------------------------------
    async record(project, ref) {
        if (!ref || typeof ref !== "string") throw this.fail("BAD_INPUT", "Falta elegir la vista de actualización.");
        // SELECT *: así funciona aunque la columna OPCIONES_JSON (filas fijas) aún no exista en este entorno.
        const rows = await Provider.runQuery(`SELECT *
            FROM ${Provider.qualifyControl("ACTUALIZACIONES")}
            WHERE PROYECTO_ID = '${Provider.esc(project.PROYECTO_ID)}'
              AND ACTUALIZACION_ID = '${Provider.esc(ref)}'`);
        if (!rows.length) throw this.fail("NOT_FOUND", "La vista de actualización ya no existe en este proyecto.");
        return rows[0];
    },

    /**
     * Sesión de una vista. mode: "fresh" la rehace; "ttl" reutiliza la existente si tiene menos
     * de META_TTL; "any" reutiliza la existente sea cual sea su edad (grabar/exportar, para usar la
     * misma definición con la que se cargó).
     * Se hace en el menor número de viajes posible: definición y dimensiones a la vez; y después,
     * claves, listas de valores de las variables y de los campos, también a la vez. T.metaP agrupa
     * lo que no bloquea el arranque (claves y listas de los campos): quien lo necesite lo espera.
     */
    async open(project, ref, mode) {
        await Informes.guard();
        const key = project.PROYECTO_ID + ":" + ref;
        const hit = this.sessions[key];
        if (hit && (mode === "any" || (mode === "ttl" && Date.now() - hit.openedAt < this.META_TTL))) {
            await hit.metaP;
            return hit;
        }
        if (mode !== "fresh" && this.opening[key]) return this.opening[key];
        const p = this.build(project, ref, key).finally(() => { if (this.opening[key] === p) delete this.opening[key]; });
        this.opening[key] = p;
        return p;
    },

    async build(project, ref, key) {
        const T = Object.create(TableUpdates);
        T.project = project;
        T.openedAt = Date.now();
        const dimsP = T.loadDimensions();             // no depende de la vista: arranca ya
        dimsP.catch(() => {});
        const record = await this.record(project, ref);
        T.runRecord = record;
        T.runScreen = T.parseScreen(record.VARIABLES_JSON);
        T.runFields = T.safeParse(record.CAMPOS_JSON, []).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
        T.vars = T.flatVars(T.runScreen);
        T.runOptions = T.parseOptions(record);          // opciones de la actualización (filas fijas)
        // La clave se recalcula en vivo desde Parametrización, igual que en startRun().
        const keysP = T.getParamKeyColumns(record.TABLA).then(k => T.applyParamKeys(T.runFields, k));
        await dimsP;
        const optsP = this.loadFieldOptions(T);
        T.metaP = Promise.all([keysP, optsP]).then(() => T);
        T.metaP.catch(() => {});
        await T.resolveRunVarOptions();
        this.sessions[key] = T;
        return T;
    },

    /** Listas de valores de todos los campos validados, a la vez (antes, una detrás de otra). */
    async loadFieldOptions(T) {
        T.fieldOptions = {};
        await Promise.all(T.runFields.filter(f => f.validation && f.validation.type !== "NONE").map(async f => {
            const opts = await T.resolveValidationOptions(f.validation);
            // "Completar con todos los valores" genera una fila por valor: sin recortar la lista.
            const all = f.validation.showAll ? opts : opts.slice(0, this.MAX_OPTIONS);
            T.fieldOptions[f.name] = all.map(o => ({ id: String(o.id), desc: o.desc || "" }));
        }));
    },

    // ---------------------------------------------------------
    // Descripción de una vista (para el editor)
    // ---------------------------------------------------------
    async list(project) {
        await Informes.guard();
        try { await Provider.runQuery(DracoSchema.ddl("ACTUALIZACIONES")); } catch (e) { /* ver TableUpdates.loadList */ }
        const rows = await Provider.runQuery(`SELECT *
            FROM ${Provider.qualifyControl("ACTUALIZACIONES")}
            WHERE PROYECTO_ID = '${Provider.esc(project.PROYECTO_ID)}'
            ORDER BY NOMBRE`);
        return rows.map(r => ({
            id: r.ACTUALIZACION_ID,
            name: r.NOMBRE,
            description: r.DESCRIPCION || "",
            table: r.TABLA,
            nVars: TableUpdates.flatVars(TableUpdates.parseScreen(r.VARIABLES_JSON)).length,
            nFields: TableUpdates.safeParse(r.CAMPOS_JSON, []).length
        }));
    },

    /** Tabla multidimensional de la vista ({rows,cols,value} con nombres de campo) o null. */
    pivotInfo(T) {
        const pv = T.parsePivot(T.runOptions, T.runFields);
        return pv ? { rows: pv.rows.map(f => f.name), cols: pv.cols.map(f => f.name), value: pv.value.name } : null;
    },

    validationSummary(v) {
        if (!v || v.type === "NONE") return null;
        return { type: v.type, allowEmpty: v.allowEmpty !== false, showText: !!v.showText, showAll: !!v.showAll, searchHelp: v.searchHelp || "LISTBOX" };
    },

    async describe(project, ref) {
        const T = await this.open(project, ref, "fresh");
        await T.metaP;
        const vars = T.vars.map(v => {
            const val = v.validation || {};
            const opts = (T.runVarOptions[v.id] || []);
            return {
                id: v.id, name: v.name, label: v.label || v.name,
                type: v.type || "STRING", selectMode: v.selectMode || "unico",
                required: !!(val.type && val.type !== "NONE" && val.allowEmpty === false),
                validation: this.validationSummary(val),
                options: opts.slice(0, this.MAX_OPTIONS).map(o => ({ id: String(o.id), desc: o.desc || "" })),
                optionsTruncated: opts.length > this.MAX_OPTIONS
            };
        });
        const descKeys = T.getKeyColumns({ fields: T.runFields, columns: T.runFields.map(f => f.name) });
        const fields = T.runFields.map(f => ({
            name: f.name, label: f.description || f.name, key: descKeys.includes(f.name),
            filter: (f.filter && f.filter.type === "VAR") ? { type: "VAR", variable: T.varNameById(f.filter.value) }
                  : (f.filter && f.filter.type === "CONST" && f.filter.value !== "" && f.filter.value != null) ? { type: "CONST", value: String(f.filter.value) }
                  : { type: "NONE" },
            validation: this.validationSummary(f.validation)
        }));
        return { id: ref, name: T.runRecord.NOMBRE, description: T.runRecord.DESCRIPCION || "", table: T.runRecord.TABLA, title: T.runScreen.title || T.runRecord.NOMBRE, fixedRows: T.isFixedRows(T.runOptions, T.runFields), pivot: this.pivotInfo(T), vars, fields };
    },

    // ---------------------------------------------------------
    // Valores de las variables
    // ---------------------------------------------------------
    /** Normaliza lo que manda el informe a las formas que espera TableUpdates.buildWhere. */
    cleanValues(T, values) {
        const src = (values && typeof values === "object") ? values : {};
        const str = (x) => (x == null ? "" : String(x)).slice(0, 500);
        const out = {};
        T.vars.forEach(v => {
            const raw = src[v.name];
            if (v.selectMode === "multiple" || v.selectMode === "cualquiera") {
                const allowed = (T.SELOPT_MODE_OPTIONS[v.selectMode] || T.SELOPT_MODE_OPTIONS.cualquiera).map(o => o.value);
                out[v.name] = (Array.isArray(raw) ? raw : []).slice(0, 200).map(r => ({
                    sign: r && r.sign === "E" ? "E" : "I",
                    option: r && allowed.includes(r.option) ? r.option : "EQ",
                    low: str(r && r.low),
                    high: str(r && r.high)
                })).filter(r => r.low.trim() !== "");
            } else if (v.selectMode === "rango") {
                out[v.name] = { low: str(raw && raw.low), high: str(raw && raw.high) };
            } else if (v.type === "BOOLEAN") {
                out[v.name] = !!raw;
            } else {
                out[v.name] = str(raw);
            }
        });
        return out;
    },

    /** Misma validación que la pantalla de variables de la ejecución original. */
    checkVars(T, vals) {
        const selOpt = {};
        T.vars.forEach(v => { if (v.selectMode === "multiple" || v.selectMode === "cualquiera") selOpt[v.id] = vals[v.name]; });
        const res = UI.validateScreenVars(T.vars, vals, T.runVarOptions, selOpt);
        if (!res.valid) {
            throw this.fail("BAD_VARS", "Revisa las variables de esta tabla en las propiedades del bloque: " + res.errors.map(e => e.message).join(" "), res.errors);
        }
    },

    // ---------------------------------------------------------
    // Lectura de la tabla
    // ---------------------------------------------------------
    pick(row, c) {
        if (row[c] !== undefined) return row[c];
        const k = Object.keys(row).find(x => x.toUpperCase() === c.toUpperCase());
        return k !== undefined ? row[k] : "";
    },
    toObjects(cols, rows) {
        return rows.map(r => { const o = {}; cols.forEach(c => { const v = this.pick(r, c); o[c] = v == null ? "" : String(v); }); return o; });
    },

    async readRows(T, where) {
        const table = Provider.qualify(T.project.DATASET, T.runRecord.TABLA);
        const rows = await Provider.runQuery(`SELECT * FROM ${table} ${where}`);
        if (rows.length > this.MAX_ROWS) {
            throw this.fail("TOO_MANY", `La tabla tiene ${rows.length} filas con este filtro y el máximo para editarla aquí es ${this.MAX_ROWS}. Acota los valores de las variables.`);
        }
        return { table, rows };
    },

    async load(project, ref, values, force) {
        const T = await this.open(project, ref, force ? "fresh" : "ttl");
        const vals = this.cleanValues(T, values);
        this.checkVars(T, vals);                                   // se valida ANTES de consultar la tabla
        const where = T.buildWhere(T.runFields, vals);
        const [{ rows }] = await Promise.all([this.readRows(T, where), T.metaP]);   // datos y metadatos a la vez
        const columns = T.runFields.map(f => f.name);
        // Filas fijas / "Completar con todos los valores": las filas salen de los valores posibles
        const comp = T.completeRows(T.runFields, T.fieldOptions, this.toObjects(columns, rows), vals);
        const objs = this.toObjects(columns, comp.rows);
        const keyCols = T.getKeyColumns({ fields: T.runFields, columns });
        return {
            name: T.runRecord.NOMBRE, table: T.runRecord.TABLA, columns, keyCols,
            fixedRows: T.isFixedRows(T.runOptions, T.runFields),
            pivot: this.pivotInfo(T),
            fields: T.runFields.map(f => ({ name: f.name, label: f.description || f.name, key: keyCols.includes(f.name), validation: this.validationSummary(f.validation) })),
            optionsByField: T.fieldOptions,
            rows: objs.map(o => columns.map(c => o[c]))
        };
    },

    // ---------------------------------------------------------
    // Grabar (con o sin ejecutar)
    // ---------------------------------------------------------
    cleanRows(cols, rows) {
        if (!Array.isArray(rows)) throw this.fail("BAD_INPUT", "No hay filas que grabar.");
        if (rows.length > this.MAX_ROWS) throw this.fail("TOO_MANY", `Hay demasiadas filas (${rows.length}); el máximo es ${this.MAX_ROWS}.`);
        return rows.map(r => {
            const arr = Array.isArray(r) ? r : [];
            const o = { __isNew: false };
            cols.forEach((c, i) => { o[c] = arr[i] == null ? "" : String(arr[i]); });
            return o;
        });
    },

    /** Las mismas comprobaciones que TableUpdates.commitGrid(), devueltas como error en vez de toast. */
    checkRows(T, state) {
        const keyCols = T.getKeyColumns(state);
        const seen = new Set(), dup = new Set();
        state.currentRows.forEach(r => {
            const k = T.rowKeyValue(r, keyCols);
            if (!k.replace(/\u0001/g, "")) return;
            if (seen.has(k)) dup.add(k); else seen.add(k);
        });
        if (dup.size) {
            throw this.fail("DUP_KEYS", `No se puede grabar: hay ${dup.size} valor(es) de "${keyCols.join(" + ")}" repetidos. Corrige o elimina las filas duplicadas.`, { keys: [...dup] });
        }
        const bad = [];
        state.fields.forEach(f => {
            if (f.validation && f.validation.type !== "NONE" && f.validation.allowEmpty === false) {
                state.currentRows.forEach(r => { if (r[f.name] === "" || r[f.name] == null) bad.push(f.name); });
            }
        });
        if (bad.length) {
            throw this.fail("EMPTY_NOT_ALLOWED", `No se puede grabar: hay valor(es) vacío(s) en campos que no lo permiten (${[...new Set(bad)].join(", ")}).`, { fields: [...new Set(bad)] });
        }
    },

    async insertRows(table, columns, rows) {
        for (let i = 0; i < rows.length; i += this.CHUNK) {
            const chunk = rows.slice(i, i + this.CHUNK);
            const values = chunk.map(r =>
                `(${columns.map(c => (r[c] === "" || r[c] === null || r[c] === undefined) ? "NULL" : `'${Provider.esc(r[c])}'`).join(", ")})`
            ).join(", ");
            if (chunk.length) await Provider.runQuery(`INSERT INTO ${table} (${columns.map(c => Provider.ident(c, "columna")).join(", ")}) VALUES ${values}`);
        }
    },

    async save(project, ref, values, rows, doCommit) {
        const T = await this.open(project, ref, "any");
        const vals = this.cleanValues(T, values);
        this.checkVars(T, vals);
        const columns = T.runFields.map(f => Provider.ident(f.name, "columna"));   // auditoría F-07
        const current = this.cleanRows(columns, rows);

        const where = T.buildWhere(T.runFields, vals);          // siempre recalculado aquí
        const { table, rows: dbRows } = await this.readRows(T, where);
        const original = this.toObjects(columns, dbRows);

        const state = { fields: T.runFields, columns, originalRows: original, currentRows: current };
        this.checkRows(T, state);
        T.gridState = state;
        const diff = T.computeDiff();
        const keyCols = T.getKeyColumns(state);
        // Filas fijas: el servidor lo hace cumplir (no se fía de la página). Con "Completar con todos los
        // valores" solo pueden ser nuevas las combinaciones de valores; sin él, no puede haber altas ni bajas.
        if (T.isFixedRows(T.runOptions, T.runFields)) {
            const comp = T.completeRows(T.runFields, T.fieldOptions, original, vals);
            const bad = comp.universeKeys
                ? diff.added.filter(r => !comp.universeKeys.has(T.rowKeyValue(r, keyCols)))
                : diff.added.concat(diff.deleted);
            if (bad.length) {
                throw this.fail("FIXED_ROWS", "No se puede grabar: esta tabla tiene filas fijas y no admite crear ni eliminar filas.");
            }
        }
        const label = (r) => T.rowKeyLabel(r, keyCols);
        const cap = (a) => a.slice(0, 300).map(label);
        const summary = {
            changed: !!(diff.added.length || diff.modified.length || diff.deleted.length),
            added: diff.added.length, modified: diff.modified.length, deleted: diff.deleted.length,
            addedKeys: cap(diff.added), modifiedKeys: cap(diff.modified.map(m => m.after)), deletedKeys: cap(diff.deleted),
            table: T.runRecord.TABLA
        };
        if (!doCommit || !summary.changed) return summary;

        // Misma estrategia que la pantalla original: borrar el subconjunto filtrado y reinsertarlo completo.
        await Provider.runQuery(`DELETE FROM ${table} ${where || "WHERE TRUE"}`);
        try {
            await this.insertRows(table, columns, current);
        } catch (err) {
            // La pantalla original no hace esto: si el INSERT falla, se intenta dejar la tabla como estaba.
            try {
                await this.insertRows(table, columns, original);
            } catch (err2) {
                throw this.fail("SAVE_FAILED", "No se pudo grabar y tampoco restaurar los datos anteriores. Revisa la tabla " + T.runRecord.TABLA + " antes de seguir: " + err.message);
            }
            throw this.fail("SAVE_FAILED", "No se pudo grabar; se han restaurado los datos anteriores. " + err.message);
        }
        return summary;
    },

    // ---------------------------------------------------------
    // Exportar e importar ficheros
    // ---------------------------------------------------------
    async exportRows(project, ref, p) {
        const T = await this.open(project, ref, "any");
        const columns = T.runFields.map(f => f.name);
        const aoa = [columns, ...this.cleanRows(columns, p.rows).map(r => columns.map(c => r[c]))];
        const base = T.runRecord.TABLA;
        if (p.kind === "csv") {
            UI.downloadBlob(`${base}.csv`, "\uFEFF" + T.toCsv(aoa), "text/csv;charset=utf-8");
        } else {
            if (typeof XLSX === "undefined") throw this.fail("NOT_AVAILABLE", "La librería de Excel no está cargada en esta página.");
            const ws = XLSX.utils.aoa_to_sheet(aoa);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Datos");
            XLSX.writeFile(wb, `${base}.xlsx`);
        }
        return { name: `${base}.${p.kind === "csv" ? "csv" : "xlsx"}` };
    },

    /** Lee un CSV (texto) o Excel (ArrayBuffer) y devuelve sus filas como array de arrays. */
    async parseFile(p) {
        if (typeof XLSX === "undefined") throw this.fail("NOT_AVAILABLE", "La librería de Excel no está cargada en esta página.");
        if (!p || (typeof p.data !== "string" && !(p.data instanceof ArrayBuffer))) throw this.fail("BAD_INPUT", "El archivo no se ha podido leer.");
        try {
            // CSV: raw:true lee todo como texto literal. Sin eso la librería convierte "2026-01" en una
            // fecha ("1/1/26") y corrompe claves y períodos al importar (la pantalla original de
            // Actualización de tablas tiene esa misma trampa).
            const wb = (typeof p.data === "string")
                ? XLSX.read(p.data, { type: "string", raw: true })
                : XLSX.read(new Uint8Array(p.data), { type: "array" });
            const ws = wb.Sheets[wb.SheetNames[0]];
            return { aoa: XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false }) };
        } catch (err) {
            throw this.fail("BAD_INPUT", "No se pudo interpretar el archivo: " + err.message);
        }
    },

    // ---------------------------------------------------------
    // Precarga: el visor la lanza en cuanto tiene la definición del informe, mientras
    // se prepara y dibuja la página, para que la tabla esté lista (o casi) al aparecer.
    // ---------------------------------------------------------
    pkey(project, ref, values) { return project.PROYECTO_ID + "|" + ref + "|" + JSON.stringify(values || {}); },

    prefetch(project, spec) {
        const now = Date.now();
        Object.keys(this.pending).forEach(k => { if (now - this.pending[k].at > 60000) delete this.pending[k]; });
        let n = 0;
        const walk = (blocks) => (blocks || []).forEach(b => {
            if (!b || n >= 6) return;
            if (b.type === "tableUpdate" && b.ref) {
                const key = this.pkey(project, b.ref, b.values);
                if (!this.pending[key]) {
                    const pr = this.load(project, b.ref, b.values, false);
                    pr.catch(() => {});
                    this.pending[key] = { pr, at: now };
                    n++;
                }
            } else if (b.type === "tabs") (b.tabs || []).forEach(t => walk(t.blocks));
        });
        ((spec && spec.sections) || []).forEach(s => walk(s.blocks));
    },

    takePending(project, ref, values) {
        const key = this.pkey(project, ref, values), x = this.pending[key];
        if (!x) return null;
        delete this.pending[key];
        return (Date.now() - x.at < 60000) ? x.pr : null;
    },

    // ---------------------------------------------------------
    // Despacho
    // ---------------------------------------------------------
    async handle(project, op, p) {
        if (!project) throw this.fail("BAD_INPUT", "No hay ningún proyecto seleccionado.");
        p = p || {};
        try {
            switch (op) {
                case "tu.list": return await this.list(project);
                case "tu.describe": return await this.describe(project, p.ref);
                case "tu.load": {
                    const pre = p.force ? null : this.takePending(project, p.ref, p.values);
                    return pre ? await pre : await this.load(project, p.ref, p.values, !!p.force);
                }
                case "tu.preview": return await this.save(project, p.ref, p.values, p.rows, false);
                case "tu.commit": return await this.save(project, p.ref, p.values, p.rows, true);
                case "tu.export": return await this.exportRows(project, p.ref, p);
                case "tu.parse": return await this.parseFile(p);
                default: throw this.fail("NOT_AVAILABLE", "Esta función todavía no está disponible.");
            }
        } catch (err) {
            throw Informes.asError(err);
        }
    }
};

window.InformesDatos = InformesDatos;
