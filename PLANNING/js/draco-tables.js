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
    ENGINE_URL: "tablas-draco/engine.html?v=20261012b",

    engines: new Map(),      // "proyecto:id" -> { frame, host, stamp, chain, lastUse }
    pending: {},             // precarga: "proyecto:id" -> { pr, at }

    // Los tests (y quien quiera) pueden sustituir cómo se obtiene el modelo de un cubo
    modelLoader: null,

    fail(code, message) { return Informes.fail(code, message); },

    // ---------------------------------------------------------------------
    // Datos (tabla de control). Mismo patrón que Informes: Informes.q() crea la tabla y repite una vez si falta.
    // ---------------------------------------------------------------------
    t() { return Provider.qualifyControl(this.TABLE); },

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
            await Informes.q(`UPDATE ${t} SET NOMBRE = '${Provider.esc(nombre)}', DESCRIPCION = '${Provider.esc(descripcion)}',
                    CUBO = '${Provider.esc(cubo || "")}', MOTOR = '${Provider.esc(motor)}',
                    DEFINICION_JSON = '${Provider.esc(json)}', SQL_TEXT = '${Provider.esc(sql)}',
                    SQL_GENERADO = CURRENT_TIMESTAMP(), MODELO_FECHA = ${mf}, FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                WHERE TABLA_DRACO_ID = '${Provider.esc(id)}' AND PROYECTO_ID = '${pid}'`);
            this.dispose(projectId, id);
            return { id, nombre };
        }
        const newId = id || Provider.newId();
        await Informes.q(`INSERT INTO ${t}
            (TABLA_DRACO_ID, PROYECTO_ID, NOMBRE, DESCRIPCION, CUBO, MOTOR, DEFINICION_JSON, SQL_TEXT, SQL_GENERADO, MODELO_FECHA, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
            VALUES ('${Provider.esc(newId)}', '${pid}', '${Provider.esc(nombre)}', '${Provider.esc(descripcion)}', '${Provider.esc(cubo || "")}',
                    '${Provider.esc(motor)}', '${Provider.esc(json)}', '${Provider.esc(sql)}', CURRENT_TIMESTAMP(), ${mf},
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
        const reportId = String(win.ReportStore.getActiveReportId());
        const reports = JSON.parse(settings.epm_reports || "{}");
        if (!reports[reportId]) throw this.fail("NO_REPORT", "No hay ningún informe diseñado.");
        settings.epm_reports = JSON.stringify({ [reportId]: reports[reportId] });
        const modelName = reports[reportId].semanticModelName || "";
        const models = JSON.parse(settings.epm_semanticModels || "{}");
        const keep = {};
        if (modelName && models[modelName]) keep[modelName] = models[modelName];
        settings.epm_semanticModels = JSON.stringify(keep);
        settings.epm_activeSemanticModel = modelName;
        delete settings.epm_reportClipboard;

        const hidden = JSON.parse(JSON.stringify(host.state.hiddenSheets || {}));
        if (hidden.EDIT_REPORT) { delete hidden.EDIT_REPORT["0_23"]; delete hidden.EDIT_REPORT["0_24"]; }   // X1 e Y1
        const def = {
            v: 1, reportId, cubo: modelName, settings,
            hiddenSheets: hidden, hiddenSheetMerges: host.state.hiddenSheetMerges || {}, namedRanges: host.state.namedRanges || {}
        };
        return def;
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

    dispose(projectId, id) {
        const k = this.key(projectId, id), e = this.engines.get(k);
        if (e) { try { e.frame.remove(); } catch (x) { /* nada */ } this.engines.delete(k); }
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
            if (Object.keys(cell).length) out[`${r - r1}_${c - c1}`] = cell;
        });
        const merges = (host.state.merges || []).filter(m => m.r >= r1 && m.r <= r2 && m.c >= c1 && m.c <= c2)
            .map(m => ({ r: m.r - r1, c: m.c - c1, rowSpan: m.rowSpan, colSpan: m.colSpan }));
        return { r0: r1, c0: c1, rows: r2 - r1 + 1, cols: c2 - c1 + 1, cells: out, merges };
    },

    /** Ejecuta una tabla guardada con SU SQL (sin regenerarlo) y devuelve la foto de lo pintado. */
    async run(project, id, force) {
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
            await win.DracoEngineApi.refresh(def.reportId, row.SQL_TEXT);
            return { id, nombre: row.NOMBRE, motor: row.MOTOR || "", sqlGenerado: row.SQL_GENERADO || null, desfasado: item.desfasado, grid: this.snapshot(eng.host) };
        });
        eng.chain = result.catch(() => {});
        return result;
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
                    const pre = p.force ? null : this.takePending(project, p.ref);
                    return pre ? await pre : await this.run(project, p.ref, p.force ? "fresh" : undefined);
                }
                case "dt.toggle": return await this.toggle(project, p.ref, p.r, p.c);
                case "dt.design":
                    if (!ui || typeof ui.openDesigner !== "function") throw this.fail("NOT_AVAILABLE", "El diseñador solo está disponible en el generador de informes.");
                    return await ui.openDesigner(project, p.ref || null);
                default: throw this.fail("NOT_AVAILABLE", "Esta función todavía no está disponible.");
            }
        } catch (err) {
            throw Informes.asError(err);
        }
    }
};

window.DracoTables = DracoTables;
