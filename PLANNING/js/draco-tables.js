/**
 * DracoTables — "Tablas Draco": informes tipo tabla dinámica (los del addin de Excel) dentro de un informe.
 *
 * Una Tabla Draco guarda en DRACO_CONTROL.TABLAS_DRACO:
 *   - DEFINICION_JSON: la definición del informe tal como la guarda el addin (ver exportDefinition());
 *   - SQL_TEXT: el SQL YA GENERADO al guardar en el diseñador. Al refrescar (visor, vista previa) NO se
 *     vuelve a generar: se ejecuta tal cual. Se regenera abriendo el diseñador y guardando.
 *   - CUBO / MOTOR / MODELO_FECHA: sobre qué cubo está, para qué motor es el SQL y la fecha del modelo con
 *     la que se generó (si el cubo se regenera después, la tabla avisa de que el SQL está desfasado).
 *
 * El motor es el MISMO código del addin (PLANNING/tablas-draco/, generado por tools/sync-addin.py) corriendo
 * en un iframe sin interfaz con un "Excel falso" (tablas-draco/js/host-bridge.js). Para ejecutar una tabla:
 * se arranca ese iframe con su definición, se le pasa el SQL guardado, el motor ejecuta la consulta con la
 * conexión de Planning y pinta el resultado en su hoja; de ahí se saca una foto (celdas con su formato) que la
 * página del informe dibuja como HTML. El iframe se conserva para poder expandir/contraer jerarquías sin
 * volver a consultar.
 *
 * Operaciones para la herramienta de informes (ver generador-app.js / visor-app.js): dt.list, dt.design,
 * dt.run, dt.toggle.
 */
const DracoTables = {
    TABLE: "TABLAS_DRACO",
    MAX_FIELD_BYTES: 400000,
    MAX_NAME: 120,
    MAX_DESC: 500,
    MAX_ENGINES: 6,
    ENGINE_URL: "tablas-draco/engine.html?v=20261022c",

    engines: new Map(),      // "proyecto:id" -> { frame, host, stamp, chain, lastUse }
    pending: {},             // precarga: "proyecto:id" -> { pr, at }

    // Los tests (y quien quiera) pueden sustituir cómo se obtiene el modelo de un cubo
    modelLoader: null,

    fail(code, message) { return Informes.fail(code, message); },

    // ---------------------------------------------------------------------
    // Datos (tabla de control). Mismo patrón que Informes: Informes.q() crea la tabla y repite una vez si falta.
    // ---------------------------------------------------------------------
    t() { return Provider.qualifyControl(this.TABLE); },

    /**
     * Literal de texto seguro para BigQuery y Snowflake. Provider.esc() solo protege comillas y barras: BigQuery NO admite
     * saltos de línea reales dentro de '...' ("Unclosed string literal"), y el SQL generado por el addin los lleva (CRLF).
     * Se escriben como \n \r \t (escapes que entienden los dos motores) y vuelven a ser saltos al leerlos.
     */
    lit(v) {
        return "'" + Provider.esc(v == null ? "" : v).replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/\t/g, "\\t") + "'";
    },

    /** Lista sin la definición ni el SQL (pesan). `desfasado`: el cubo cambió después de generar el SQL. */
    async list(projectId) {
        await Informes.guard();
        const rows = await Informes.q(`
            SELECT t.TABLA_DRACO_ID, t.NOMBRE, t.DESCRIPCION, t.CUBO, t.MOTOR, t.SQL_GENERADO, t.MODELO_FECHA,
                   t.FECHA_MODIFICACION, c.MODELO_YAML_FECHA AS CUBO_FECHA,
                   CASE WHEN t.SQL_TEXT IS NULL OR t.SQL_TEXT = '' THEN 0 ELSE 1 END AS TIENE_SQL
            FROM ${this.t()} t
            LEFT JOIN ${Provider.qualifyControl("CUBOS")} c ON c.PROYECTO_ID = t.PROYECTO_ID AND c.CUBOS = t.CUBO
            WHERE t.PROYECTO_ID = '${Provider.esc(projectId)}'
            ORDER BY t.NOMBRE`);
        return rows.map(r => this.toItem(r));
    },

    toItem(r) {
        // list() trae TIENE_SQL calculado; get() trae la fila entera (con SQL_TEXT)
        const tieneSql = r.TIENE_SQL !== undefined ? (String(r.TIENE_SQL) === "1" || r.TIENE_SQL === true) : !!(r.SQL_TEXT && String(r.SQL_TEXT).trim());
        return {
            id: r.TABLA_DRACO_ID, nombre: r.NOMBRE, descripcion: r.DESCRIPCION || "", cubo: r.CUBO || "",
            motor: r.MOTOR || "", sqlGenerado: r.SQL_GENERADO || null, tieneSql,
            desfasado: !!(tieneSql && r.CUBO && String(r.CUBO_FECHA || "") !== String(r.MODELO_FECHA || ""))
        };
    },

    async get(projectId, id) {
        await Informes.guard();
        const rows = await Informes.q(`
            SELECT t.*, c.MODELO_YAML_FECHA AS CUBO_FECHA
            FROM ${this.t()} t
            LEFT JOIN ${Provider.qualifyControl("CUBOS")} c ON c.PROYECTO_ID = t.PROYECTO_ID AND c.CUBOS = t.CUBO
            WHERE t.TABLA_DRACO_ID = '${Provider.esc(id)}' AND t.PROYECTO_ID = '${Provider.esc(projectId)}'`);
        if (!rows.length) throw this.fail("NOT_FOUND", "Esa tabla Draco ya no existe (puede que la hayan eliminado).");
        return rows[0];
    },

    /** Crea (sin id) o sobrescribe (con id). definition: objeto; sql: texto. Devuelve { id, nombre }. */
    async save({ id, projectId, nombre, descripcion, cubo, definition, sql, modeloFecha }) {
        if (!projectId) throw this.fail("BAD_INPUT", "Falta el proyecto.");
        nombre = String(nombre || "").trim();
        descripcion = String(descripcion || "").trim();
        if (!nombre) throw this.fail("BAD_INPUT", "La tabla Draco necesita un nombre.");
        if (nombre.length > this.MAX_NAME) throw this.fail("BAD_INPUT", `El nombre no puede pasar de ${this.MAX_NAME} caracteres.`);
        if (descripcion.length > this.MAX_DESC) throw this.fail("BAD_INPUT", `La descripción no puede pasar de ${this.MAX_DESC} caracteres.`);
        if (!definition || typeof definition !== "object") throw this.fail("BAD_INPUT", "La definición del informe no es válida.");
        sql = String(sql || "");
        if (!sql.trim()) throw this.fail("NO_SQL", "No se pudo generar el SQL del informe: comprueba que tiene campos en Filas o Columnas y un modelo elegido.");
        const json = JSON.stringify(definition);
        const bytes = (s) => new TextEncoder().encode(s).length;
        if (bytes(json) > this.MAX_FIELD_BYTES) throw this.fail("TOO_BIG", `La definición ocupa ${Math.round(bytes(json) / 1024)} KB y el máximo es ${Math.round(this.MAX_FIELD_BYTES / 1000)} KB.`);
        if (bytes(sql) > this.MAX_FIELD_BYTES) throw this.fail("TOO_BIG", `El SQL ocupa ${Math.round(bytes(sql) / 1024)} KB y el máximo es ${Math.round(this.MAX_FIELD_BYTES / 1000)} KB.`);

        await Informes.guard();
        const t = this.t(), pid = Provider.esc(projectId);
        const dup = await Informes.q(`SELECT TABLA_DRACO_ID FROM ${t} WHERE PROYECTO_ID = '${pid}' AND UPPER(NOMBRE) = UPPER('${Provider.esc(nombre)}')`);
        if (dup.some(r => r.TABLA_DRACO_ID !== id)) throw this.fail("NAME_EXISTS", "Ya existe una tabla Draco con ese nombre en este proyecto.");

        let exists = false;
        if (id) exists = (await Informes.q(`SELECT TABLA_DRACO_ID FROM ${t} WHERE TABLA_DRACO_ID = '${Provider.esc(id)}' AND PROYECTO_ID = '${pid}'`)).length > 0;
        const motor = Provider.key();
        const mf = modeloFecha === undefined || modeloFecha === null || modeloFecha === "" ? "NULL" : `'${Provider.esc(modeloFecha)}'`;
        if (exists) {
            await Informes.q(`UPDATE ${t} SET NOMBRE = ${this.lit(nombre)}, DESCRIPCION = ${this.lit(descripcion)},
                    CUBO = ${this.lit(cubo || "")}, MOTOR = ${this.lit(motor)},
                    DEFINICION_JSON = ${this.lit(json)}, SQL_TEXT = ${this.lit(sql)},
                    SQL_GENERADO = CURRENT_TIMESTAMP(), MODELO_FECHA = ${mf}, FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                WHERE TABLA_DRACO_ID = '${Provider.esc(id)}' AND PROYECTO_ID = '${pid}'`);
            this.dispose(projectId, id);
            return { id, nombre };
        }
        const newId = id || Provider.newId();
        await Informes.q(`INSERT INTO ${t}
            (TABLA_DRACO_ID, PROYECTO_ID, NOMBRE, DESCRIPCION, CUBO, MOTOR, DEFINICION_JSON, SQL_TEXT, SQL_GENERADO, MODELO_FECHA, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
            VALUES ('${Provider.esc(newId)}', '${pid}', ${this.lit(nombre)}, ${this.lit(descripcion)}, ${this.lit(cubo || "")},
                    ${this.lit(motor)}, ${this.lit(json)}, ${this.lit(sql)}, CURRENT_TIMESTAMP(), ${mf},
                    ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`);
        return { id: newId, nombre };
    },

    // ---------------------------------------------------------------------
    // Cubos y modelos semánticos (los cubos del proyecto = los "modelos" del selector del addin)
    // ---------------------------------------------------------------------
    async cubes(projectId) {
        const rows = await Provider.runQuery(`
            SELECT CUBO_ID, CUBOS, MODELO_YAML_PATH, MODELO_YAML_FECHA FROM ${Provider.qualifyControl("CUBOS")}
            WHERE PROYECTO_ID = '${Provider.esc(projectId)}' ORDER BY CUBOS`);
        return rows.map(r => ({ id: r.CUBO_ID, name: r.CUBOS, path: r.MODELO_YAML_PATH || "", fecha: r.MODELO_YAML_FECHA || "" }));
    },

    splitQualified(q) {
        const parts = String(q || "").replace(/`/g, "").split(".");
        return { project: parts[0] || "", dataset: parts[1] || "", table: parts.slice(2).join(".") || "" };
    },

    /** LookML ya parseado (js/lkml-parse.js) -> modelo en el formato de SemanticModelStore.saveModel() del addin. */
    lkmlToSemanticModel(lkml) {
        const fact = this.splitQualified(lkml.factTable);
        const fields = [];
        lkml.dimensions.forEach(dim => {
            const dp = this.splitQualified(dim.table);
            fields.push({
                enabled: true, type: "DIMENSION", name: dim.id,
                relProject: dp.project, relDataset: dp.dataset, relTable: dp.table,
                attributes: dim.attributes.map(a => ({ name: a.colId, alias: a.colId, dataType: a.type || "STRING", isKey: !!a.isKey, enabled: true })),
                hierarchies: dim.hierarchies.map(h => ({ name: h.name, levels: h.levels.map(l => ({ attribute: l.colId })) }))
            });
        });
        lkml.measures.forEach(m => {
            fields.push({ enabled: true, type: "MEASURE", name: m.column, aggregation: "SUM", format: "#,##0.00" });
        });
        return { fact, fields };
    },

    /**
     * Modelo del cubo en el formato de SemanticModelStore (el que lee el addin), sea cual sea el proveedor:
     *   1) BigQuery: el .lkml que Planning sube a GitHub al guardar el cubo (rápido, sin tocar tablas).
     *   2) Si no hay .lkml (cubo sin generar, GitHub/secreto inaccesible...) o el proveedor es Snowflake (que
     *      solo guarda un YAML en un stage y no genera .lkml): se RECONSTRUYE desde el propio cubo, con lo mismo que
     *      Planning usa para generar ese modelo (CUBOS.CAMPOS_JSON + DIMENSIONES + JERARQUIAS).
     * Cacheado por cubo y fecha del modelo. Si ninguna vía funciona, el error cuenta el motivo de cada una.
     */
    _modelCache: {},
    engine() { return Provider.key(); },
    async resolveModel(project, cube) {
        if (this.modelLoader) return this.modelLoader(project, cube);
        const key = `${project.PROYECTO_ID}:${cube.id}:${cube.fecha}`;
        if (this._modelCache[key]) return this._modelCache[key];
        const reasons = [];
        let model = null;
        if (this.engine() !== "snowflake") {      // solo BigQuery genera .lkml
            try { model = await this.modelFromLkml(project, cube); }
            catch (err) { reasons.push("GitHub: " + err.message); }
        }
        if (!model) {
            try { model = await this.modelFromCube(project, cube); }
            catch (err) { reasons.push("cubo: " + err.message); }
        }
        if (!model || !model.fields || !model.fields.length) {
            console.error(`DracoTables: no se pudo obtener el modelo del cubo «${cube.name}»:`, reasons);
            throw this.fail("NO_MODEL", `No se pudo obtener el modelo del cubo «${cube.name}». ${reasons.join(" · ") || "El cubo no tiene dimensiones ni medidas."}`);
        }
        this._modelCache[key] = model;
        return model;
    },

    /** Vía 1: el .lkml de GitHub. Devuelve null si no existe. */
    async modelFromLkml(project, cube) {
        let path = cube.path ? cube.path.replace(/\.yaml$/i, ".lkml") : null;
        if (!path && typeof SemanticModel !== "undefined") {
            const ident = Provider.toIdentifier(cube.name) || cube.id;
            path = SemanticModel.buildPath(project, ident).replace(/\.yaml$/i, ".lkml");
        }
        if (!path) throw new Error("el cubo no tiene ruta de modelo (.lkml)");
        const cfg = (typeof DracoConfig !== "undefined" && DracoConfig.semanticModelGithub) || {};
        const token = await BQ.getGithubPatFromSecretManager();
        const text = await GithubRepo.getFile(path, { url: cfg.url, branch: cfg.branch, token });
        if (text === null) throw new Error(`no existe ${path}`);
        return this.lkmlToSemanticModel(LkmlParse.parse(text));
    },

    /** Vía 2: reconstruir el modelo desde el cubo (lo mismo que Planning usa para generar su YAML/.lkml). */
    async modelFromCube(project, cube) {
        if (typeof SemanticModel === "undefined") throw new Error("SemanticModel no está cargado");
        const pid = Provider.esc(project.PROYECTO_ID);
        const rows = await Provider.runQuery(`SELECT CUBO_ID, CUBOS, DESCRIPCION, TABLA, CAMPOS_JSON FROM ${Provider.qualifyControl("CUBOS")}
            WHERE CUBO_ID = '${Provider.esc(cube.id)}' AND PROYECTO_ID = '${pid}'`);
        if (!rows.length) throw new Error("el cubo ya no existe");
        let spec;
        try { spec = JSON.parse(rows[0].CAMPOS_JSON || "{}"); } catch (e) { throw new Error("la definición del cubo no es válida"); }
        spec.dimensions = spec.dimensions || []; spec.measures = spec.measures || [];
        if (!spec.dimensions.length && !spec.measures.length) throw new Error("el cubo no tiene dimensiones ni medidas");
        const ids = spec.dimensions.map(d => d.id).filter(Boolean);
        let dimRows = [], hierRows = [];
        if (ids.length) {
            const inList = ids.map(id => `'${Provider.esc(id)}'`).join(", ");
            [dimRows, hierRows] = await Promise.all([
                Provider.runQuery(`SELECT DIMENSION_ID, DIMENSION, DESCRIPCION, TABLA, CAMPOS_JSON FROM ${Provider.qualifyControl("DIMENSIONES")} WHERE DIMENSION_ID IN (${inList})`),
                Provider.runQuery(`SELECT DIMENSION_ID, JERARQUIA, NIVELES_JSON FROM ${Provider.qualifyControl("JERARQUIAS")} WHERE DIMENSION_ID IN (${inList})`)
            ]);
        }
        const built = SemanticModel.build(project, { id: rows[0].CUBO_ID, name: rows[0].CUBOS, description: rows[0].DESCRIPCION, table: rows[0].TABLA }, spec, dimRows, hierRows);
        return this.planningModelToSemanticModel(built);
    },

    /** Objeto de SemanticModel.build() -> formato de SemanticModelStore. La dimensión se llama como su columna en la tabla de hechos. */
    planningModelToSemanticModel(m) {
        const project = (m.model && m.model.database) || "", dataset = (m.model && m.model.schema) || "";
        const fks = (m.fact && m.fact.foreign_keys) || [];
        const fields = [];
        (m.dimensions || []).forEach((d, i) => {
            const fk = fks.find(f => f.dimension === d.name) || fks[i];
            fields.push({
                enabled: true, type: "DIMENSION", name: (fk && fk.column) || d.key_column,
                relProject: project, relDataset: dataset, relTable: d.table,
                attributes: (d.attributes || []).map(a => ({ name: a.column, alias: a.column, dataType: a.type || "STRING", isKey: !!a.is_key, enabled: true })),
                hierarchies: (d.hierarchies || []).map(h => ({ name: h.name, levels: (h.levels || []).map(l => ({ attribute: l.column })) }))
            });
        });
        ((m.fact && m.fact.measures) || []).forEach(x => fields.push({ enabled: true, type: "MEASURE", name: x.column, aggregation: "SUM", format: "#,##0.00" }));
        return { fact: { project, dataset, table: m.fact && m.fact.table }, fields };
    },

    // ---------------------------------------------------------------------
    // Definición del informe (lo que se guarda) <-> estado del "Excel falso"
    // ---------------------------------------------------------------------
    /**
     * Extrae la definición del informe ACTIVO del iframe del diseñador. Se guarda lo mínimo para rehacerlo:
     * ajustes del addin (informe, filtros, modelo activo con sus campos) y hojas internas (sin el último SQL/JSON,
     * EDIT_REPORT!X1/Y1, que pesan). No se guarda el resultado pintado: se rehace al ejecutar.
     */
    async exportDefinition(win, host) {
        const settings = Object.assign({}, host.state.taskpaneSettings || {});
        const activeId = String(win.ReportStore.getActiveReportId());
        const reports = JSON.parse(settings.epm_reports || "{}");
        if (!reports[activeId]) throw this.fail("NO_REPORT", "No hay ningún informe diseñado.");
        // INFORMES VINCULADOS: la tabla es el informe PRINCIPAL y se guardan también sus vinculados (se pintan
        // juntos, ver dracoRefreshLinkGroup en commands.js). Sin vinculados, como siempre: solo el informe activo.
        const m = typeof win.dracoLinkGroupMain === "function" ? win.dracoLinkGroupMain(Number(activeId)) : null;
        const reportId = (m === null || m === undefined) ? activeId : String(m);
        const linked = (typeof win.dracoLinkedOf === "function" ? win.dracoLinkedOf(Number(reportId)) : []).map(String).filter(id => reports[id]);
        const groupIds = [reportId, ...linked].filter(id => reports[id]);
        settings.epm_reports = JSON.stringify(Object.fromEntries(groupIds.map(id => [id, reports[id]])));
        settings.epm_activeReportId = Number(reportId);
        const modelName = reports[reportId].semanticModelName || "";
        const models = JSON.parse(settings.epm_semanticModels || "{}");
        const keep = {};
        groupIds.forEach(id => { const mn = reports[id].semanticModelName || ""; if (mn && models[mn]) keep[mn] = models[mn]; });
        settings.epm_semanticModels = JSON.stringify(keep);
        settings.epm_activeSemanticModel = modelName;
        delete settings.epm_reportClipboard;

        const hidden = JSON.parse(JSON.stringify(host.state.hiddenSheets || {}));
        if (hidden.EDIT_REPORT) { delete hidden.EDIT_REPORT["0_23"]; delete hidden.EDIT_REPORT["0_24"]; }   // X1 e Y1
        const def = {
            v: 1, reportId, cubo: modelName, settings,
            linked,                                   // ids de los informes vinculados (vacío si no hay)
            hiddenSheets: hidden, hiddenSheetMerges: host.state.hiddenSheetMerges || {}, namedRanges: host.state.namedRanges || {},
            // Informes Estáticos / Híbridos: las líneas del eje Estático SON el informe (celdas =EPM_VALUE /
            // =EPM_FORMULA de la hoja). Se guardan con su formato; lo pintado desde los datos no (se rehace).
            cells: this.staticCells(host)
        };
        return def;
    },

    /** Celdas de la hoja principal con fórmula EPM_VALUE / EPM_FORMULA (definición del eje Estático). */
    staticCells(host) {
        const out = {}, cells = host.state.cells || {};
        Object.keys(cells).forEach(k => {
            const o = cells[k];
            if (o && typeof o.f === "string" && /^\s*=\s*@?\s*EPM_(VALUE|FORMULA)\s*\(/i.test(o.f)) {
                const copy = JSON.parse(JSON.stringify(o));
                delete copy.db;
                out[k] = copy;
            }
        });
        return out;
    },

    /**
     * CSS de una barra de datos ya calculada por el motor (host-bridge: o.db = {a, p, col, only}; a = eje y
     * p = extremo de la barra, en fracción del ancho de la celda). Fondo con la barra y la línea del eje.
     */
    /** HTML del icono de una celda (semáforo / señales): mismo dibujo en el diseñador y en el visor. */
    iconHtml(ic) {
        if (!ic) return "";
        const name = ic.k === "signs" ? ["Rombo rojo", "Triángulo ámbar", "Círculo verde"][ic.i] : ["Rojo", "Ámbar", "Verde"][ic.i];
        return `<span class="dx-ic dx-ic-${ic.k === "signs" ? "s" : "t"}${Math.max(0, Math.min(2, ic.i | 0))}" role="img" aria-label="${name}" title="${name}"></span>`;
    },

    barCss(db) {
        if (!db) return "";
        const pc = (x) => (Math.round(x * 1000) / 10) + "%";
        const lo = Math.min(db.a, db.p), hi = Math.max(db.a, db.p);
        const col = /^#[0-9a-f]{3,8}$/i.test(db.col || "") ? db.col : "#70D06A";
        const axis = `linear-gradient(to right, transparent calc(${pc(db.a)} - 0.5px), #1A2233 calc(${pc(db.a)} - 0.5px), #1A2233 calc(${pc(db.a)} + 0.5px), transparent calc(${pc(db.a)} + 0.5px))`;
        const bar = `linear-gradient(to right, transparent ${pc(lo)}, ${col} ${pc(lo)}, ${col} ${pc(hi)}, transparent ${pc(hi)})`;
        return `background-image:${axis},${bar};background-size:100% 70%;background-position:center;background-repeat:no-repeat`;
    },

    stateFromDefinition(def) {
        return {
            rows: 1000, cols: 200,
            cells: def.cells ? JSON.parse(JSON.stringify(def.cells)) : {}, merges: def.merges ? JSON.parse(JSON.stringify(def.merges)) : [],
            colWidths: {}, rowHeights: {},
            taskpaneSettings: JSON.parse(JSON.stringify(def.settings || {})),
            hiddenSheets: JSON.parse(JSON.stringify(def.hiddenSheets || {})),
            hiddenSheetMerges: JSON.parse(JSON.stringify(def.hiddenSheetMerges || {})),
            namedRanges: JSON.parse(JSON.stringify(def.namedRanges || {}))
        };
    },

    // ---------------------------------------------------------------------
    // Motor sin interfaz
    // ---------------------------------------------------------------------
    key(projectId, id) { return projectId + ":" + id; },

    createEngine(state) {
        const frame = document.createElement("iframe");
        frame.setAttribute("aria-hidden", "true");
        frame.style.cssText = "position:absolute;width:0;height:0;border:0;visibility:hidden;";
        const host = { state, selection: { r1: 0, c1: 0, r2: 0, c2: 0 }, cubes: [], resolveModel: async () => null, markDirty() {}, renderGrid() {} };
        frame.__dracoHost = host;
        const ready = new Promise((resolve) => frame.addEventListener("load", () => resolve()));
        frame.src = this.ENGINE_URL;
        document.body.appendChild(frame);
        return { frame, host, ready, chain: Promise.resolve(), lastUse: Date.now() };
    },

    /**
     * Descarta el motor de una tabla. Sale de la lista AL MOMENTO (la próxima ejecución crea uno nuevo), pero su iframe
     * se elimina cuando termina lo que estuviera ejecutando: si se quitara en mitad de una ejecución, el código del
     * motor seguiría corriendo sin anfitrión ("host-bridge: el iframe no tiene anfitrión") y esa ejecución fallaría.
     */
    dispose(projectId, id) {
        const k = this.key(projectId, id), e = this.engines.get(k);
        if (e) {
            this.engines.delete(k);
            Promise.resolve(e.chain).then(() => { try { e.frame.remove(); } catch (x) { /* nada */ } });
        }
        delete this.pending[k];
    },

    disposeAll() { Array.from(this.engines.keys()).forEach(k => { const [p, ...r] = k.split(":"); this.dispose(p, r.join(":")); }); },

    trimEngines() {
        if (this.engines.size <= this.MAX_ENGINES) return;
        const old = Array.from(this.engines.entries()).sort((a, b) => a[1].lastUse - b[1].lastUse)[0];
        if (old) { const [p, ...r] = old[0].split(":"); this.dispose(p, r.join(":")); }
    },

    /**
     * Texto de un valor numérico según su formato de celda de Excel (#,##0.00 · 0% · 0.0 · #,##0 "€"...), con
     * separadores a la española (1.234,50). Devuelve null si no hay nada que formatear (texto, General, @).
     */
    formatValue(v, nf) {
        if (v === undefined || v === null || v === "" || !nf || nf === "General" || nf === "@") return null;
        const raw = String(v).trim();
        if (!/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(raw)) return null;
        let n = Number(raw);
        if (!isFinite(n)) return null;
        const fmt = String(nf).split(";")[0];
        const lit = fmt.replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1").replace(/_./g, " ").replace(/\*./g, "");
        const m = lit.match(/^([^#0?,.%]*)([#0?,.]*%?[#0?,.]*)(.*)$/);
        const prefix = m ? m[1] : "", core = m ? m[2] : lit, suffix = m ? m[3] : "";
        const pct = core.indexOf("%") >= 0;
        if (pct) n *= 100;
        const dot = core.indexOf(".");
        const decimals = dot >= 0 ? (core.slice(dot + 1).match(/[0#?]/g) || []).length : 0;
        const grouping = core.slice(0, dot >= 0 ? dot : undefined).indexOf(",") >= 0;
        const neg = n < 0;
        let [ip, fp] = Math.abs(n).toFixed(decimals).split(".");
        if (grouping) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
        return (neg ? "-" : "") + prefix + ip + (fp ? "," + fp : "") + (pct ? "%" : "") + suffix;
    },

    /** Foto de lo pintado en la hoja: recorte al contenido, con el formato de cada celda. */
    snapshot(host) {
        const cells = host.state.cells || {};
        let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
        Object.keys(cells).forEach(k => {
            const o = cells[k];
            if (!o || ((o.v === undefined || o.v === "") && !o.bg)) return;
            const [r, c] = k.split("_").map(Number);
            if (r < r1) r1 = r; if (c < c1) c1 = c; if (r > r2) r2 = r; if (c > c2) c2 = c;
        });
        if (r2 < 0) return { r0: 0, c0: 0, rows: 0, cols: 0, cells: {}, merges: [] };
        const out = {};
        Object.keys(cells).forEach(k => {
            const [r, c] = k.split("_").map(Number);
            if (r < r1 || r > r2 || c < c1 || c > c2) return;
            const o = cells[k];
            if (!o) return;
            const cell = {};
            ["v", "b", "i", "u", "al", "nf", "ind", "bg", "col", "fs"].forEach(p => { if (o[p] !== undefined && o[p] !== "" && o[p] !== 0) cell[p] = o[p]; });
            const bd = (b) => (b && (b.style === undefined || b.style !== "none")) ? (b.color || "#000000") : 0;
            const edges = { t: bd(o.bt), r: bd(o.br), b: bd(o.bb), l: bd(o.bl) };
            if (edges.t || edges.r || edges.b || edges.l) cell.bd = edges;
            const t = this.formatValue(o.v, o.nf);
            if (t !== null) cell.t = t;                       // texto ya formateado (el HTML no tiene que saber de formatos de Excel)
            if (o.ic) cell.ic = { k: o.ic.k, i: o.ic.i, only: o.ic.only ? 1 : 0 };   // icono (semáforo / señales)
            if (o.db) {                                       // barra de datos (formato condicional), ya como CSS
                cell.bar = this.barCss(o.db);
                if (!cell.al) cell.al = Number(o.v) < 0 ? "left" : "right";
                if (o.db.only) cell.only = 1;
            }
            if (Object.keys(cell).length) out[`${r - r1}_${c - c1}`] = cell;
        });
        const merges = (host.state.merges || []).filter(m => m.r >= r1 && m.r <= r2 && m.c >= c1 && m.c <= c2)
            .map(m => ({ r: m.r - r1, c: m.c - c1, rowSpan: m.rowSpan, colSpan: m.colSpan }));
        return { r0: r1, c0: c1, rows: r2 - r1 + 1, cols: c2 - c1 + 1, cells: out, merges };
    },

    /**
     * Variables de una tabla guardada (las del diseñador, ver ADDIN/src/js/variablesStore.js):
     * [{ name, dim, attr, value, display }]. Para el futuro selector de variables del generador/visor.
     */
    async variables(project, id) {
        const row = await this.get(project.PROYECTO_ID, id);
        const def = JSON.parse(row.DEFINICION_JSON || "{}");
        let vars = {};
        try { vars = JSON.parse(((def.settings || {}).epm_variables) || "{}"); } catch (e) { vars = {}; }
        return Object.keys(vars).sort().map(name => Object.assign({ name }, vars[name]));
    },

    // ---------------------------------------------------------------------
    // KPIs (bloque "dracoKpi" y $$NOMBRE$$ en los textos)
    // ---------------------------------------------------------------------
    // Un KPI es una Tabla Draco con UNA medida y solo filtros. Al guardarlo, el diseñador guarda en la definición sus
    // piezas (def.kpi = { agg, field, from, joins, conds }, ver buildDracoKpiParts en el addin). Todos los KPIs de un
    // informe se calculan con UNA consulta (kpiSql):
    //   - se agrupan por tabla de hechos + joins: cada grupo lee su tabla de hechos UNA vez;
    //   - cada KPI es una columna de agregación condicional: SUM(IF(<sus filtros>, f.campo, NULL)) AS K_n;
    //   - el WHERE del grupo es el OR de los filtros de sus KPIs (solo se leen las filas que interesan);
    //   - los grupos (una fila cada uno) se unen con CROSS JOIN: el resultado es UNA fila con una columna por KPI.
    // Así 100 KPIs del mismo cubo cuestan una lectura, no 100 (UNION ALL de subconsultas leería la tabla 100 veces).
    _kpiCache: {},
    KPI_TTL: 0,          // sin caché: los KPIs se calculan al entrar y siempre que cambie su definición o sus filtros

    /** Sustituye $NOMBRE$ en un trozo de SQL con los valores de `vars` (escapando según el motor). Lanza si falta alguna. */
    resolveKpiVars(sqlPart, vars) {
        const bq = Provider.key() === "bigquery";
        const escIn = (v) => bq ? String(v).replace(/\\/g, "\\\\").replace(/'/g, "\\'") : String(v).replace(/'/g, "''");
        let out = "", inStr = false, i = 0;
        while (i < sqlPart.length) {
            const ch = sqlPart[i];
            if (ch === "'") {
                if (inStr && !bq && sqlPart[i + 1] === "'") { out += "''"; i += 2; continue; }
                inStr = !inStr; out += ch; i++; continue;
            }
            if (inStr && bq && ch === "\\") { out += sqlPart.substr(i, 2); i += 2; continue; }
            const m = ch === "$" ? /^\$([A-Za-z][A-Za-z0-9_]{0,39})\$/.exec(sqlPart.slice(i)) : null;
            if (m) {
                const k = m[1].toUpperCase();
                if (!Object.prototype.hasOwnProperty.call(vars, k)) throw new Error(`usa la variable $${k}$, que no existe`);
                const v = String(vars[k]);
                out += inStr ? escIn(v) : (/^-?\d+(\.\d+)?$/.test(v.trim()) ? v.trim() : "'" + escIn(v) + "'");
                i += m[0].length; continue;
            }
            out += ch; i++;
        }
        return out;
    },

    /**
     * SQL de todos los KPIs. items: [{ key, parts, conds (ya con las variables resueltas) }].
     * Devuelve { sql, groups: [{ sql, keys: [[col, key]] }] } (los grupos sueltos sirven para reintentar si falla).
     */
    kpiSql(items) {
        const prov = Provider.key();
        const iff = (cond, x) => !cond ? x : (prov === "snowflake" ? `IFF(${cond}, ${x}, NULL)` : prov === "fabric" ? `CASE WHEN ${cond} THEN ${x} END` : `IF(${cond}, ${x}, NULL)`);
        const AGG = { SUM: "SUM", AVG: "AVG", COUNT: "COUNT", MIN: "MIN", MAX: "MAX" };
        const byGroup = new Map();
        items.forEach((it, ix) => {
            // un KPI con agregación de excepción necesita su propia subconsulta (GROUP BY por sus dimensiones de referencia)
            const gk = it.exc ? "exc#" + ix : String(it.parts.from) + "\n" + String(it.parts.joins || "");
            if (!byGroup.has(gk)) byGroup.set(gk, []);
            byGroup.get(gk).push(it);
        });
        // joins de los filtros de la barra: se comparten en todo el grupo (LEFT JOIN 1:1 a dimensiones, no cambia el
        // resultado de los KPIs que no los usan) para seguir leyendo la tabla de hechos UNA vez
        const extraJoins = (lst) => { const m = new Map(); lst.forEach(it => (it.extraJoins || []).forEach(j => m.set(j, 1))); return Array.from(m.keys()); };
        let n = 0;
        const groups = Array.from(byGroup.values()).map(list => {
            const keys = [], cols = [], conds = [];
            let all = true;
            list.forEach(it => {
                const col = "K_" + (n++);
                const cond = it.conds.length ? it.conds.map(c => "(" + c + ")").join(" AND ") : null;
                if (cond) conds.push("(" + cond + ")"); else all = false;
                const x = iff(cond, it.parts.field);
                const agg = String(it.parts.agg || "SUM").toUpperCase();
                let e = agg === "COUNT_DISTINCT" ? `COUNT(DISTINCT ${x})` : (AGG[agg] ? `${AGG[agg]}(${x})` : `${it.parts.agg}(${x})`);
                const sc = Number(it.parts.scale);
                if (isFinite(sc) && sc > 1) e = "(" + e + ") / " + sc;      // factor de escala de la medida
                cols.push("    " + e + " AS " + col);
                keys.push([col, it.key]);
            });
            const first = list[0].parts;
            if (list[0].exc) {
                // AGREGACIÓN DE EXCEPCIÓN: 1) la medida se agrega (con su agregación normal) por las dimensiones de
                // referencia; 2) sobre esos valores se aplica la agregación de excepción (MIN, MAX, AVG...).
                const it = list[0], x = it.exc;
                const inner = cols[0].trim().replace(/ AS K_\d+$/, "");
                const cond = it.conds.length ? it.conds.map(c => "(" + c + ")").join(" AND ") : null;
                const joins = [(first.joins || "").trim()].concat(it.extraJoins || [], x.joins || []).filter(Boolean);
                const outer = x.agg === "COUNT" ? "COUNT(v)" : x.agg + "(v)";
                const sql = "SELECT\n    " + outer + " AS " + keys[0][0] + "\nFROM (\nSELECT\n    " + inner + " AS v\n" + first.from
                    + (joins.length ? "\n" + joins.join("\n") : "") + (cond ? "\nWHERE\n    " + cond : "")
                    + "\nGROUP BY " + x.cols.join(", ") + "\n) xg";
                return { sql, keys };
            }
            const where = (all && conds.length) ? "\nWHERE\n    " + Array.from(new Set(conds)).join("\n    OR ") : "";
            const ej = extraJoins(list);
            return { sql: "SELECT\n" + cols.join(",\n") + "\n" + first.from + "\n" + (first.joins || "").trim() + (ej.length ? "\n" + ej.join("\n") : "") + where, keys };
        });
        const sql = groups.length === 1 ? groups[0].sql
            : "SELECT * FROM\n" + groups.map((g, i) => "(\n" + g.sql + "\n) G" + i).join("\nCROSS JOIN\n");
        return { sql, groups };
    },

    /**
     * Valores de varios KPIs con una consulta (más una a la tabla de control para leer sus definiciones).
     * refs: ids de Tablas Draco que son KPI; variables: { NOMBRE: valor } de la página (opcional).
     * Devuelve { values: { ref: número|null }, errors: { ref: mensaje }, nombres: { ref: nombre } }.
     */
    /**
     * KPIs que viven DENTRO del informe (spec.dracoKpis): la herramienta los manda ya con sus piezas
     * ([{ id, parts, vars }], tomadas de SU copia del informe, nunca de la página). Una sola consulta.
     */
    /** Join y condición de un filtro de la barra sobre un KPI (alias fijo por dimensión: xf_<DIM>). */
    kpiFilterSql(f) {
        const sf = Provider.key() === "snowflake";
        const db = f.fld.relProject || (sf ? SF.getDatabase() : BQ.getGcpProject());
        const table = sf ? `${db}.${f.fld.relDataset}.${f.fld.relTable}` : `\`${db}.${f.fld.relDataset}.${f.fld.relTable}\``;
        const key = (f.fld.attributes || []).find(a => a.isKey) || (f.fld.attributes || [])[0];
        const ok = (x) => /^[A-Za-z_][A-Za-z0-9_]*$/.test(String(x || ""));
        if (!key || !ok(f.fld.name) || !ok(key.name) || !ok(f.att.name)) return null;
        const alias = "xf_" + String(f.fld.name).replace(/[^A-Za-z0-9_]/g, "_");
        const num = /INT|NUMERIC|FLOAT|DECIMAL|NUMBER|DOUBLE|BIGNUMERIC/i.test(f.att.dataType || "") && /^-?\d+(\.\d+)?$/.test(f.value.trim());
        const lit = num ? f.value.trim() : "'" + this.resolveKpiVars("'$V$'", { V: f.value }).slice(1, -1) + "'";
        return { join: `LEFT JOIN ${table} ${alias} ON f.${f.fld.name} = ${alias}.${key.name}`, cond: `${alias}.${f.att.name} = ${lit}` };
    },

    KPI_EXC_AGGS: ["MIN", "MAX", "AVG", "SUM", "COUNT"],

    async kpisInline(project, kpis, variables, filtersByKpi) {
        const out = { values: {}, errors: {}, nombres: {} };
        const list = (kpis || []).filter(k => k && k.id);
        if (!list.length) return out;
        const ck = project.PROYECTO_ID + "|inline|" + JSON.stringify(list.map(k => [k.id, k.parts, k.vars])) + "|" + JSON.stringify(variables || {}) + "|" + JSON.stringify(filtersByKpi || {});
        const hit = this._kpiCache[ck];
        if (hit && Date.now() - hit.at < this.KPI_TTL) return hit.res;
        const over = {};
        Object.keys(variables || {}).forEach(k => { const v = variables[k]; over[k.toUpperCase()] = (v && typeof v === "object") ? v.value : v; });
        const items = [];
        for (const k of list) {
            const parts = k.parts;
            if (!parts || !parts.from || !parts.field) { out.errors[k.id] = "Este KPI no está terminado: ábrelo en el diseñador y pulsa Guardar."; continue; }
            const vars = Object.assign({}, k.vars || {}, over);
            let conds;
            try { conds = (parts.conds || []).map(c => this.resolveKpiVars(c, vars)); }
            catch (e) { out.errors[k.id] = "El KPI " + e.message + "."; continue; }
            // filtros de la barra aplicados directamente a este KPI (validados contra el modelo de su cubo)
            const extraJoins = [];
            const fl = await this.validRuntimeFilters(project, k.cubo, (filtersByKpi || {})[k.id]);
            fl.forEach(f => { const x = this.kpiFilterSql(f); if (x) { extraJoins.push(x.join); conds.push(x.cond); } });
            // Agregación de excepción: la pone el diseñador en las piezas del KPI (Opciones de campo de la medida)
            // como { agg, cols: [columnas de referencia], joins: [JOIN que faltan] }; sin ella, nada cambia.
            const pe = parts.exc;
            const exc = (pe && this.KPI_EXC_AGGS.indexOf(String(pe.agg || "").toUpperCase()) >= 0 && Array.isArray(pe.cols) && pe.cols.length)
                ? { agg: String(pe.agg).toUpperCase(), cols: pe.cols, joins: Array.isArray(pe.joins) ? pe.joins : [] } : null;
            items.push({ key: k.id, parts, conds, extraJoins, exc });
        }
        await this.runKpiItems(items, out);
        this._kpiCache[ck] = { at: Date.now(), res: out };
        return out;
    },

    /** Ejecuta la consulta única de los KPIs y reparte los valores (si falla, grupo a grupo). */
    async runKpiItems(items, out) {
        if (!items.length) return;
        await Informes.guard();
        const { sql, groups } = this.kpiSql(items);
        const take = (row, keys) => keys.forEach(([col, key]) => {
            const v = row ? (row[col] ?? row[col.toLowerCase()]) : null;
            out.values[key] = (v === null || v === undefined || v === "") ? null : Number(v);
        });
        try {
            const res = await Provider.runQuery(sql);
            take(res[0], groups.flatMap(g => g.keys));
        } catch (err) {
            for (const g of groups) {
                try { const r = await Provider.runQuery(g.sql); take(r[0], g.keys); }
                catch (e2) { g.keys.forEach(([, key]) => { out.errors[key] = String(e2.message || e2); }); }
            }
        }
    },

    async kpis(project, refs, variables) {
        refs = Array.from(new Set((refs || []).filter(Boolean)));
        const out = { values: {}, errors: {}, nombres: {} };
        if (!refs.length) return out;
        const vkey = JSON.stringify(variables || {});
        const ck = project.PROYECTO_ID + "|" + refs.slice().sort().join(",") + "|" + vkey;
        const hit = this._kpiCache[ck];
        if (hit && Date.now() - hit.at < this.KPI_TTL) return hit.res;

        await Informes.guard();
        const rows = await Informes.q(`
            SELECT TABLA_DRACO_ID, NOMBRE, DEFINICION_JSON
            FROM ${this.t()}
            WHERE PROYECTO_ID = '${Provider.esc(project.PROYECTO_ID)}'
              AND TABLA_DRACO_ID IN (${refs.map(r => "'" + Provider.esc(r) + "'").join(", ")})`);
        const over = {};
        Object.keys(variables || {}).forEach(k => { const v = variables[k]; over[k.toUpperCase()] = (v && typeof v === "object") ? v.value : v; });
        const items = [];
        const found = new Set();
        rows.forEach(row => {
            const ref = row.TABLA_DRACO_ID; found.add(ref);
            out.nombres[ref] = row.NOMBRE;
            let def = {};
            try { def = JSON.parse(row.DEFINICION_JSON || "{}"); } catch (e) { def = {}; }
            const parts = def.kpi;
            if (!parts || !parts.from || !parts.field) { out.errors[ref] = "No es un KPI (guárdalo desde el diseñador de KPI)."; return; }
            const vars = {};
            try { const d = JSON.parse(((def.settings || {}).epm_variables) || "{}"); Object.keys(d).forEach(k => { vars[k.toUpperCase()] = d[k].value; }); } catch (e) { /* sin variables */ }
            Object.assign(vars, over);
            try { items.push({ key: ref, parts, conds: (parts.conds || []).map(c => this.resolveKpiVars(c, vars)) }); }
            catch (e) { out.errors[ref] = "El KPI " + e.message + "."; }
        });
        refs.forEach(r => { if (!found.has(r)) out.errors[r] = "Este KPI ya no existe."; });

        if (items.length) {
            const { sql, groups } = this.kpiSql(items);
            const take = (row, keys) => keys.forEach(([col, key]) => {
                const v = row ? (row[col] ?? row[col.toLowerCase()]) : null;
                out.values[key] = (v === null || v === undefined || v === "") ? null : Number(v);
            });
            try {
                const res = await Provider.runQuery(sql);
                take(res[0], groups.flatMap(g => g.keys));
            } catch (err) {
                // Un KPI roto no debe tumbar a los demás: se reintenta grupo a grupo.
                for (const g of groups) {
                    try { const r = await Provider.runQuery(g.sql); take(r[0], g.keys); }
                    catch (e2) { g.keys.forEach(([, key]) => { out.errors[key] = String(e2.message || e2); }); }
                }
            }
        }
        this._kpiCache[ck] = { at: Date.now(), res: out };
        return out;
    },

    // ---------------------------------------------------------------------
    // Selector de variable (bloque "dracoVar" del generador)
    // ---------------------------------------------------------------------
    _membersCache: {},

    async cubeByName(project, name) {
        const cube = (await this.cubes(project.PROYECTO_ID)).find(c => c.name === name);
        if (!cube) throw this.fail("NO_CUBE", `El cubo «${name}» ya no existe en este proyecto.`);
        return cube;
    },

    /** Campos (dimensión · atributo) de un cubo, con los nombres que usan las Tablas Draco (y sus variables). */
    async fields(project, cubeName) {
        const model = await this.resolveModel(project, await this.cubeByName(project, cubeName));
        const out = [];
        (model.fields || []).filter(f => f.type === "DIMENSION" && f.enabled !== false).forEach(f => {
            (f.attributes || []).filter(a => a.enabled !== false).forEach(a => out.push({
                dim: f.name, attr: a.name, label: String(f.name).toLowerCase() + " · " + String(a.alias || a.name).toLowerCase()
            }));
        });
        return out.sort((a, b) => a.label.localeCompare(b.label));
    },

    /** Valores distintos de un atributo (lista del desplegable). Caché de 10 minutos por cubo/campo. */
    async members(project, cubeName, dim, attr) {
        const k = [project.PROYECTO_ID, cubeName, dim, attr].join("|").toUpperCase();
        const hit = this._membersCache[k];
        if (hit && Date.now() - hit.at < 600000) return hit.list;
        const model = await this.resolveModel(project, await this.cubeByName(project, cubeName));
        const f = (model.fields || []).find(x => x.type === "DIMENSION" && String(x.name).toUpperCase() === String(dim).toUpperCase());
        if (!f) throw this.fail("NO_FIELD", `La dimensión «${dim}» no está en el cubo «${cubeName}».`);
        const a = (f.attributes || []).find(x => String(x.name).toUpperCase() === String(attr).toUpperCase());
        if (!a) throw this.fail("NO_FIELD", `El atributo «${attr}» no está en la dimensión «${dim}».`);
        const sf = Provider.key() === "snowflake";
        const db = f.relProject || (sf ? SF.getDatabase() : BQ.getGcpProject());
        const table = sf ? `${db}.${f.relDataset}.${f.relTable}` : `\`${db}.${f.relDataset}.${f.relTable}\``;
        const col = /^[A-Za-z_][A-Za-z0-9_]*$/.test(a.name) ? a.name : null;
        if (!col) throw this.fail("BAD_INPUT", "Nombre de atributo no válido.");
        const rows = await Provider.runQuery(`SELECT DISTINCT ${col} AS V FROM ${table} WHERE ${col} IS NOT NULL ORDER BY 1 LIMIT 5000`);
        const list = rows.map(r => String(r.V ?? r.v ?? Object.values(r)[0]));
        this._membersCache[k] = { at: Date.now(), list };
        return list;
    },

    /**
     * Variables que usan las tablas indicadas (las de un informe), agrupadas por nombre:
     * [{ name, dim, attr, value, display, tablas: [nombre], refs: [id] }]. Una tabla que ya no exista se ignora.
     */
    async varsFor(project, refs) {
        const by = {};
        for (const ref of Array.from(new Set(refs || [])).filter(Boolean)) {
            let row;
            try { row = await this.get(project.PROYECTO_ID, ref); } catch (e) { continue; }
            let vars = {};
            try { vars = JSON.parse(((JSON.parse(row.DEFINICION_JSON || "{}").settings || {}).epm_variables) || "{}"); } catch (e) { vars = {}; }
            // solo las que la tabla USA de verdad (su SQL lleva el marcador)
            const used = new Set(String(row.SQL_TEXT || "").match(/\$[A-Za-z][A-Za-z0-9_]*\$/g) || []);
            Object.keys(vars).forEach(name => {
                if (used.size && !used.has("$" + name + "$")) return;
                const v = vars[name];
                const g = by[name] || (by[name] = { name, dim: v.dim, attr: v.attr, value: v.value, display: v.display || v.value, tablas: [], refs: [] });
                g.tablas.push(row.NOMBRE); g.refs.push(ref);
            });
        }
        return Object.values(by).sort((a, b) => a.name.localeCompare(b.name));
    },

    /**
     * Ejecuta una tabla guardada con SU SQL (sin regenerarlo) y devuelve la foto de lo pintado.
     * opts.variables = { NOMBRE: valor | { value, display } } cambia el valor de las variables SOLO en esta
     * ejecución (el SQL guardado lleva los marcadores $NOMBRE$ y se sustituyen al ejecutar).
     */
    async run(project, id, force, opts) {
        try {
            return await this.runOnce(project, id, force, opts);
        } catch (err) {
            // Red de seguridad: el motor se desmontó mientras ejecutaba (p. ej. lo descartó otra petición). Se reintenta
            // una vez con un motor nuevo en vez de mostrar el error.
            if (/no tiene anfitri/i.test(String(err && err.message))) {
                console.warn("DracoTables: el motor se desmontó durante la ejecución; se reintenta con uno nuevo.");
                return await this.runOnce(project, id, "fresh", opts);
            }
            throw err;
        }
    },

    /**
     * Filtros de la barra que se aplican DIRECTAMENTE a una tabla/gráfico/KPI: [{dim, attr, value}] -> solo los que
     * existen en el modelo del cubo (el resto se ignora: ese objeto no tiene ese campo). Devuelve también el campo del
     * modelo (para construir joins de KPI).
     */
    async validRuntimeFilters(project, cubeName, filters) {
        const list = (filters || []).filter(f => f && f.dim && f.attr && f.value !== undefined && f.value !== null && String(f.value) !== "");
        if (!list.length || !cubeName) return [];
        let model;
        try { model = await this.resolveModel(project, await this.cubeByName(project, cubeName)); } catch (e) { return []; }
        const out = [];
        list.forEach(f => {
            const fld = (model.fields || []).find(x => x.type === "DIMENSION" && String(x.name).toUpperCase() === String(f.dim).toUpperCase());
            const att = fld && (fld.attributes || []).find(a => String(a.name).toUpperCase() === String(f.attr).toUpperCase());
            if (fld && att) out.push({ dim: fld.name, attr: att.name, value: String(f.value), fld, att });
        });
        return out;
    },

    async runOnce(project, id, force, opts) {
        const runVars = (opts && opts.variables) || {};
        const runFilters = (opts && opts.filters) || [];
        const k = this.key(project.PROYECTO_ID, id);
        const row = await this.get(project.PROYECTO_ID, id);
        const item = this.toItem(row);
        if (!item.tieneSql) throw this.fail("NO_SQL", "Esta tabla Draco no tiene SQL generado: ábrela en el diseñador y guárdala.");
        if (row.MOTOR && row.MOTOR !== Provider.key()) {
            throw this.fail("OTHER_ENGINE", `El SQL de esta tabla Draco se generó para ${row.MOTOR} y esta conexión es ${Provider.key()}. Ábrela en el diseñador y guárdala para regenerarlo.`);
        }
        const def = JSON.parse(row.DEFINICION_JSON || "{}");
        const stamp = String(row.FECHA_MODIFICACION || "") + "|" + String(row.SQL_GENERADO || "");
        let eng = this.engines.get(k);
        if (eng && (force === "fresh" || eng.stamp !== stamp)) { this.dispose(project.PROYECTO_ID, id); eng = null; }
        if (!eng) {
            eng = this.createEngine(this.stateFromDefinition(def));
            eng.stamp = stamp; eng.def = def; eng.sql = row.SQL_TEXT;
            this.engines.set(k, eng);
            this.trimEngines();
        }
        eng.lastUse = Date.now();
        const result = eng.chain.then(async () => {
            await eng.ready;
            const win = eng.frame.contentWindow;
            await win.Office.onReady();
            if (win.DracoVariables) win.DracoVariables.setOverrides(runVars);
            // Filtros de la barra aplicados directamente: el motor regenera el SQL con ellos (el guardado no los lleva)
            const rf = await this.validRuntimeFilters(project, row.CUBO, runFilters);
            win.__dracoRuntimeFilters = rf.length ? rf.map(f => ({ dim: f.dim, attr: f.attr, value: f.value })) : null;
            // Con informes vinculados no se usa el SQL guardado: el motor genera el de cada informe del grupo
            const group = Array.isArray(def.linked) && def.linked.length > 0;
            // Filtro «Campos del eje»: cambia los campos de Filas (o Columnas) del informe en el motor y regenera su SQL
            const axisChanged = await this.applyAxis(win, eng, def, opts && opts.axis);
            try { await win.DracoEngineApi.refresh(def.reportId, (rf.length || group || axisChanged) ? null : row.SQL_TEXT); }
            finally { win.__dracoRuntimeFilters = null; }
            const variables = win.DracoVariables ? win.DracoVariables.list() : [];
            return { id, nombre: row.NOMBRE, motor: row.MOTOR || "", sqlGenerado: row.SQL_GENERADO || null, desfasado: item.desfasado, variables, grid: this.snapshot(eng.host) };
        });
        eng.chain = result.catch(() => {});
        return result;
    },

    /**
     * Filtro «Campos del eje» de la barra: pone en el eje de Filas (axis.axis = "rows") o de Columnas ("cols") del
     * informe SOLO los campos elegidos, en ese orden ({dim, attr}); las medidas del eje se conservan al final. Un campo
     * que ya estaba en el diseño conserva su configuración. Se guarda el eje original la primera vez y se repone cuando
     * ya no hay selección. Devuelve true si el eje queda distinto del guardado (hay que regenerar el SQL).
     */
    async applyAxis(win, eng, def, axis) {
        const rep = win.ReportStore && win.ReportStore.getReport(def.reportId);
        if (!rep) return false;
        const d = rep.design || {};
        if (!eng.axisOrig) {
            if (!axis) return false;
            eng.axisOrig = { expandedRows: JSON.parse(JSON.stringify(d.expandedRows || [])), expandedCols: JSON.parse(JSON.stringify(d.expandedCols || [])) };
        }
        const orig = eng.axisOrig, fields = axis && Array.isArray(axis.fields) ? axis.fields.filter(f => f && f.dim && f.attr) : [];
        let er = orig.expandedRows, ec = orig.expandedCols;
        if (fields.length) {
            const key = (e) => String(e[0]).toUpperCase() + "|" + String(e[1]).toUpperCase();
            const known = {};
            orig.expandedRows.concat(orig.expandedCols).forEach(e => { if (e && e[0] !== "MEASURE") known[key(e)] = e; });
            const entries = fields.map(f => known[key([f.dim, f.attr])] || [f.dim, f.attr, "1", "", "", ""]);
            if (axis.axis === "cols") ec = entries.concat(orig.expandedCols.filter(e => e && e[0] === "MEASURE"));
            else er = entries.concat(orig.expandedRows.filter(e => e && e[0] === "MEASURE"));
        }
        if (JSON.stringify(d.expandedRows || []) !== JSON.stringify(er) || JSON.stringify(d.expandedCols || []) !== JSON.stringify(ec)) {
            await win.ReportStore.updateDesignFields(def.reportId, { expandedRows: JSON.parse(JSON.stringify(er)), expandedCols: JSON.parse(JSON.stringify(ec)) });
        }
        return JSON.stringify(er) !== JSON.stringify(orig.expandedRows) || JSON.stringify(ec) !== JSON.stringify(orig.expandedCols);
    },

    /** Clic en el +/- de una jerarquía: el motor repinta desde su caché, sin volver a consultar. */
    async toggle(project, id, r, c) {
        const eng = this.engines.get(this.key(project.PROYECTO_ID, id));
        if (!eng) throw this.fail("NOT_READY", "La tabla todavía no está cargada.");
        eng.lastUse = Date.now();
        const sn = this.snapshot(eng.host);
        const result = eng.chain.then(async () => {
            const win = eng.frame.contentWindow;
            await win.__fireExcelEvent("onSingleClicked", sn.r0 + r, sn.c0 + c);
            return { id, grid: this.snapshot(eng.host) };
        });
        eng.chain = result.catch(() => {});
        return result;
    },

    // ---------------------------------------------------------------------
    // Precarga: el visor la lanza en cuanto conoce el informe, mientras se dibuja la página
    // ---------------------------------------------------------------------
    prefetch(project, spec) {
        const walk = (blocks) => (blocks || []).forEach(b => {
            if (!b) return;
            if (b.type === "dracoTable" && b.ref) {
                const k = this.key(project.PROYECTO_ID, b.ref);
                if (!this.pending[k]) { const pr = this.run(project, b.ref); pr.catch(() => {}); this.pending[k] = { pr, at: Date.now() }; }
            } else if (b.type === "tabs") (b.tabs || []).forEach(t => walk(t.blocks));
        });
        ((spec && spec.sections) || []).forEach(s => walk(s.blocks));
    },

    takePending(project, id) {
        const k = this.key(project.PROYECTO_ID, id), x = this.pending[k];
        if (!x) return null;
        delete this.pending[k];
        return (Date.now() - x.at < 60000) ? x.pr : null;
    },

    // ---------------------------------------------------------------------
    // Despacho de peticiones de la herramienta / de la página
    // ---------------------------------------------------------------------
    async handle(project, op, p, ui) {
        if (!project) throw this.fail("BAD_INPUT", "No hay ningún proyecto seleccionado.");
        p = p || {};
        try {
            switch (op) {
                case "dt.list": return await this.list(project.PROYECTO_ID);
                case "dt.run": {
                    // p.variables / p.filters (opcionales): valores de variables y filtros de la barra para ESTA ejecución
                    // p.axis (opcional): campos del eje elegidos en un filtro «Campos del eje» de la barra
                    const hasVars = (p.variables && Object.keys(p.variables).length) || (p.filters && p.filters.length) || p.axis;
                    const pre = (p.force || hasVars) ? null : this.takePending(project, p.ref);
                    return pre ? await pre : await this.run(project, p.ref, p.force ? "fresh" : undefined, { variables: p.variables || {}, filters: p.filters || [], axis: p.axis || null });
                }
                case "dt.variables": return await this.variables(project, p.ref);
                case "dt.cubes": return (await this.cubes(project.PROYECTO_ID)).map(c => ({ name: c.name }));
                case "dt.fields": return await this.fields(project, p.cube);
                case "dt.members": return await this.members(project, p.cube, p.dim, p.attr);
                case "dt.varsFor": return await this.varsFor(project, p.refs || []);
                case "dt.kpis":
                    // KPIs del informe (con sus piezas, los añade la herramienta); si no vienen, el formato antiguo por id
                    return Array.isArray(p.kpis) ? await this.kpisInline(project, p.kpis, p.variables || null, p.filters || null)
                                                 : await this.kpis(project, p.refs || [], p.variables || null);
                case "dt.toggle": return await this.toggle(project, p.ref, p.r, p.c);
                case "dt.design":
                    if (!ui || typeof ui.openDesigner !== "function") throw this.fail("NOT_AVAILABLE", "El diseñador solo está disponible en el generador de informes.");
                    return await ui.openDesigner(project, p.ref || null, { kpi: !!p.kpi, def: p.def || null, nombre: p.nombre || "" });
                default: throw this.fail("NOT_AVAILABLE", "Esta función todavía no está disponible.");
            }
        } catch (err) {
            throw Informes.asError(err);
        }
    }
};

window.DracoTables = DracoTables;
