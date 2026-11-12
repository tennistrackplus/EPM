/**
 * ============================================================
 * DRACO — DIARIO DE CONSULTAS (Diagnóstico)
 * ============================================================
 * Con el Diagnóstico activo (cinta ▸ Diagnóstico), CADA consulta SQL que
 * lanza el add-in —informes, filtros, buscador y reconocimiento de
 * miembros, planificación, modelo semántico, editor de cálculos...— se
 * acumula en EDIT_REPORT a partir de la columna AD, en orden de llegada
 * (la más reciente ABAJO: así cada consulta nueva es solo una fila más que
 * escribir, no reescribir el diario entero):
 *
 *   AD Nº | AE Inicio (hh:mm:ss.mmm) | AF Tipo | AG Motor | AH Duración (s) |
 *   AI Filas | AJ Procesado (MB) | AK Facturado (MB) | AL Coste est. (USD) |
 *   AM Slot·s | AN Caché / modelo de precio | AO SQL | AP Resultado | AQ Error
 *
 * Con «Inicio» + «Duración» se ve si las consultas se solapan (paralelo) o
 * van una detrás de otra.
 *
 * La fila 1 lleva los TOTALES (consultas, segundos, MB y coste) para ver
 * de un vistazo lo que ha consumido una sesión de trabajo. Las columnas
 * numéricas son números de verdad: se pueden sumar o filtrar.
 *
 * Volumen y coste:
 *  - BigQuery: tras cada consulta se piden las estadísticas del job
 *    (jobs.get: bytes procesados y facturados, caché, slots y si corre en
 *    una RESERVA). Facturado = 0 tiene tres causas y se dice cuál:
 *      · caché de BigQuery (misma consulta reciente: gratis);
 *      · proyecto con reservas/editions: no se paga por bytes sino por
 *        capacidad (slots) — el coste se muestra como «equivalente
 *        on-demand» calculado con los bytes procesados, para comparar;
 *      · consulta que no lee tablas.
 *    Coste on-demand = facturado × DracoConfig.bigqueryPricePerTiB.
 *  - Databricks: no hay coste por consulta (se paga el tiempo encendido del
 *    warehouse): se anota la duración y las filas.
 *
 * RENDIMIENTO: el diario NO se escribe mientras hay un refresco en curso
 * (se acumula en memoria y se vuelca al terminar): escribir en Excel a la
 * vez que se pinta un informe ralentizaba el pintado varias veces.
 *
 * Sin Diagnóstico no se guarda nada (ni en memoria ni en el libro). Al
 * desactivarlo se borra todo (auditoría F-15: el libro no debe llevar SQL
 * ni resultados si se comparte).
 */
(function () {
    const SHEET = "EDIT_REPORT";
    const FIRST_COL = "AD";
    const LAST_COL = "AQ";
    const MAX_ENTRIES = 200;
    const SQL_MAX = 32000;       // límite de Excel por celda: 32 767
    const RESULT_MAX = 2000;     // el resultado completo sigue en Y1 (último informe) y en la consola
    const MB = 1024 * 1024;
    const TIB = 1024 * 1024 * 1024 * 1024;

    const entries = [];
    let seq = 0;
    let flushTimer = null;
    let flushing = false;
    let dirty = false;
    let written = 0;               // filas de entradas ya escritas en la hoja
    let fullRewrite = true;        // reescribir todo (primera vez o tras recortar las más antiguas)
    const dirtyEntries = new Set(); // entradas ya escritas que han cambiado (estadísticas del job)

    function refreshBusy() {
        try { return typeof window.dracoRefreshInProgress === "function" && window.dracoRefreshInProgress(); } catch (e) { return false; }
    }

    function enabled() {
        try { return typeof window.dracoDiagOn === "function" && window.dracoDiagOn(); } catch (e) { return false; }
    }

    function engine() {
        try { return (window.Provider && Provider.label) ? Provider.label() : "BigQuery"; } catch (e) { return ""; }
    }

    function pricePerTiB() {
        const p = typeof DracoConfig !== "undefined" ? Number(DracoConfig.bigqueryPricePerTiB) : NaN;
        return isFinite(p) && p >= 0 ? p : 6.25;
    }

    /** Tipo de consulta a partir del SQL cuando quien llama no lo indica. */
    function classify(sql) {
        const s = String(sql || "").replace(/^\s*(--[^\n]*\n|\/\*[\s\S]*?\*\/)\s*/g, "");
        if (/^\s*(INSERT|MERGE|UPDATE|DELETE)\b/i.test(s)) return "Planificación (escritura)";
        if (/^\s*(SHOW|DESCRIBE)\b|INFORMATION_SCHEMA/i.test(s)) return "Metadatos";
        if (/\b(ROW_IDS|LINE_IDS|TAGGED)\b/.test(s)) return "Informe";
        if (/SELECT\s+DISTINCT/i.test(s)) return "Miembros";
        return "Consulta";
    }

    function num(re, text) {
        const m = re.exec(text);
        return m ? Number(m[1]) : null;
    }

    /** Estadísticas de una respuesta de BigQuery (texto JSON) sin convertirla entera. */
    function statsFromText(text) {
        const t = String(text || "");
        const head = t.length > 4000 ? t.slice(0, 2000) + t.slice(-2000) : t;   // los totales van al principio o al final
        return {
            rows: num(/"totalRows"\s*:\s*"?(\d+)/, head),
            processed: num(/"totalBytesProcessed"\s*:\s*"?(\d+)/, head),
            billed: num(/"totalBytesBilled"\s*:\s*"?(\d+)/, head),
            cacheHit: /"cacheHit"\s*:\s*true/.test(head) ? true : (/"cacheHit"\s*:\s*false/.test(head) ? false : null),
            affected: num(/"numDmlAffectedRows"\s*:\s*"?(\d+)/, head),
            slotMs: num(/"totalSlotMs"\s*:\s*"?(\d+)/, head),
            jobId: (/"jobId"\s*:\s*"([^"]+)"/.exec(head) || [])[1] || "",
            projectId: (/"jobReference"\s*:\s*\{[^}]*?"projectId"\s*:\s*"([^"]+)"/.exec(head) || [])[1] || "",
            location: (/"jobReference"\s*:\s*\{[^}]*?"location"\s*:\s*"([^"]+)"/.exec(head) || [])[1] || ""
        };
    }

    function fmtTime(d) {
        return d.toTimeString().slice(0, 8) + "." + String(d.getMilliseconds()).padStart(3, "0");
    }

    /** Rellena volumen, coste y nota de una entrada a partir de sus datos crudos. */
    function compute(entry) {
        const r = entry.raw;
        const isBq = entry.engine === "BigQuery";
        let billed = r.billed, note = "";
        const processed = r.processed;
        let cost = null;
        if (!isBq) {
            note = r.cacheHit === true ? "Caché de Draco" : "Databricks: se paga el tiempo encendido del warehouse";
        } else if (r.cacheHit === true) {
            billed = 0; cost = 0;
            note = r.draco ? "Caché de Draco (sin consultar)" : "Caché de BigQuery (gratis)";
        } else if (r.reservation) {
            cost = processed > 0 ? processed / TIB * pricePerTiB() : 0;
            note = "Reserva/editions («" + r.reservation + "»): se paga capacidad (slots); coste = equivalente on-demand";
        } else if ((billed === null || billed === undefined) && processed > 0) {
            billed = Math.max(Math.ceil(processed / MB) * MB, 10 * MB);   // mínimo 10 MB por consulta
            cost = billed / TIB * pricePerTiB();
            note = "No (facturado ≈ estimado)";
        } else if (billed === 0 && processed > 0) {
            cost = processed / TIB * pricePerTiB();
            note = "Facturado 0 con datos procesados: proyecto con reservas/editions; coste = equivalente on-demand";
        } else if (billed !== null && billed !== undefined) {
            cost = billed / TIB * pricePerTiB();
            note = r.cacheHit === false ? "No" : "";
        }
        const mb = (x) => x === null || x === undefined ? "" : Math.round(x / MB * 100) / 100;
        entry.processed = mb(processed);
        entry.billed = mb(billed);
        entry.cost = cost === null ? "" : Math.round(cost * 1e6) / 1e6;
        entry.slots = r.slotMs === null || r.slotMs === undefined ? "" : Math.round(r.slotMs / 10) / 100;
        entry.cache = note;
    }

    function record(e) {
        if (!enabled()) return null;
        const start = e.start || new Date();
        const entry = {
            n: ++seq,
            time: fmtTime(start),
            kind: e.kind || classify(e.sql),
            engine: e.engine || engine(),
            secs: Math.round((e.ms || 0)) / 1000,
            rows: e.rows === null || e.rows === undefined ? "" : e.rows,
            raw: { processed: e.processed, billed: e.billed, cacheHit: e.cacheHit, slotMs: e.slotMs, reservation: "", draco: !!e.dracoCache },
            sql: String(e.sql || "").slice(0, SQL_MAX),
            result: e.result ? String(e.result).slice(0, RESULT_MAX) + (String(e.result).length > RESULT_MAX ? " …(recortado)" : "") : "",
            error: e.error ? String(e.error).slice(0, 2000) : ""
        };
        compute(entry);
        entries.push(entry);
        if (entries.length > MAX_ENTRIES) {
            entries.splice(0, entries.length - MAX_ENTRIES + 50);   // quita las 50 más antiguas de golpe
            fullRewrite = true;
        }
        dirty = true;
        schedule();
        return entry;
    }

    /**
     * Estadísticas definitivas del job (jobs.get): la respuesta de jobs.query
     * no siempre trae los bytes facturados ni dice si hay reserva. Solo con
     * el Diagnóstico activo; no retrasa el informe (va en segundo plano).
     */
    async function enrichFromJob(entry, s) {
        try {
            if (!s.jobId || !window.BQ || typeof BQ.getToken !== "function") return;
            const token = BQ.getToken();
            if (!token) return;
            const project = s.projectId || "";
            if (!project) return;
            const url = "https://bigquery.googleapis.com/bigquery/v2/projects/" + encodeURIComponent(project)
                + "/jobs/" + encodeURIComponent(s.jobId) + (s.location ? "?location=" + encodeURIComponent(s.location) : "");
            const r = await fetch(url, { headers: { "Authorization": "Bearer " + token } });
            if (!r.ok) return;
            const j = await r.json();
            const st = (j && j.statistics) || {};
            const q = st.query || {};
            const n = (x) => x === undefined || x === null || x === "" ? null : Number(x);
            if (n(q.totalBytesProcessed) !== null) entry.raw.processed = n(q.totalBytesProcessed);
            if (n(q.totalBytesBilled) !== null) entry.raw.billed = n(q.totalBytesBilled);
            if (typeof q.cacheHit === "boolean") entry.raw.cacheHit = q.cacheHit;
            if (n(st.totalSlotMs) !== null) entry.raw.slotMs = n(st.totalSlotMs);
            const resv = st.reservation_id || st.reservationId
                || (Array.isArray(st.reservationUsage) && st.reservationUsage[0] && st.reservationUsage[0].name) || "";
            if (resv) entry.raw.reservation = String(resv).split("/").pop();
            compute(entry);
            dirtyEntries.add(entry);
            dirty = true;
            schedule();
        } catch (e) { /* las estadísticas son un extra: sin ellas se queda lo que trajo la respuesta */ }
    }

    /**
     * Ejecuta fn() (que lanza la consulta) y la anota en el diario.
     * opts: { kind, engine }. Si el diagnóstico está apagado, solo ejecuta.
     */
    async function track(sql, fn, opts) {
        if (!enabled()) return fn();
        const o = opts || {};
        const start = new Date();
        const t0 = performance.now();
        const eng = o.engine || engine();
        try {
            const res = await fn();
            const ms = performance.now() - t0;
            if (typeof res === "string") {
                const s = statsFromText(res);
                const entry = record({ sql, kind: o.kind, engine: eng, ms, start, rows: s.affected !== null ? s.affected : s.rows,
                    processed: s.processed, billed: s.billed, cacheHit: s.cacheHit, slotMs: s.slotMs, result: res });
                if (entry && eng === "BigQuery") enrichFromJob(entry, s);   // sin await: no retrasa el informe
            } else {
                const rows = res && Array.isArray(res.rows) ? res.rows.length : null;
                let txt = "";
                try { txt = JSON.stringify(res && res.rows ? res.rows : res); } catch (e) { /* sin texto */ }
                record({ sql, kind: o.kind, engine: eng, ms, start, rows, result: txt });
            }
            return res;
        } catch (err) {
            record({ sql, kind: o.kind, engine: eng, ms: performance.now() - t0, start, error: (err && err.message) || String(err) });
            throw err;
        }
    }

    function schedule() {
        if (flushTimer) return;
        flushTimer = setTimeout(flush, 800);
    }

    const COLS = 14;
    const FMT_ROW = ["0", "@", "@", "@", "0.000", "#,##0", "#,##0.00", "#,##0.00", "0.000000", "#,##0.00", "@", "@", "@", "@"];
    const rowOf = (e) => [e.n, e.time, e.kind, e.engine, e.secs, e.rows, e.processed, e.billed, e.cost, e.slots, e.cache, e.sql, e.result, e.error];

    async function flush() {
        flushTimer = null;
        if (flushing) { schedule(); return; }
        if (!dirty || !enabled()) return;
        if (typeof Excel === "undefined" || !Excel.run) return;   // diálogos: sin acceso al libro
        if (refreshBusy()) { schedule(); return; }                // se vuelca al acabar el refresco
        flushing = true;
        dirty = false;
        const snap = entries.slice();
        const full = fullRewrite;
        const from = full ? 0 : written;
        const changed = full ? [] : [...dirtyEntries].map(e => snap.indexOf(e)).filter(i => i >= 0 && i < from);
        fullRewrite = false;
        dirtyEntries.clear();
        try {
            const sum = (k) => snap.reduce((a, e) => a + (typeof e[k] === "number" ? e[k] : 0), 0);
            const round = (x, d) => Math.round(x * Math.pow(10, d)) / Math.pow(10, d);
            const totals = [`CONSULTAS (diagnóstico) · ${snap.length} · en orden de llegada (la última abajo)`, "", "TOTAL", "",
                round(sum("secs"), 3), "", round(sum("processed"), 2), round(sum("billed"), 2), round(sum("cost"), 6), round(sum("slots"), 2), "", "", "", ""];
            const header = ["Nº", "Inicio", "Tipo", "Motor", "Duración (s)", "Filas", "Procesado (MB)", "Facturado (MB)",
                `Coste est. (USD, ${pricePerTiB()} $/TiB)`, "Slot·s", "Caché / modelo de precio", "SQL", "Resultado (recortado)", "Error"];

            await Excel.run(async (context) => {
                const sh = context.workbook.worksheets.getItemOrNullObject(SHEET);
                sh.load("isNullObject");
                await context.sync();
                if (sh.isNullObject) return;
                const at = (r0, nRows) => sh.getRangeByIndexes(r0, 29, nRows, COLS);   // columna AD = índice 29
                if (full) {
                    // Primera vez o tras recortar: se limpia el bloque entero y se escribe todo
                    at(0, MAX_ENTRIES + 2).clear(Excel.ClearApplyTo.contents);
                    const head = at(1, 1);
                    head.numberFormat = [header.map(() => "@")];
                    head.values = [header];
                    head.format.font.bold = true;
                }
                // Totales (fila 1): siempre, una sola fila
                const tot = at(0, 1);
                tot.numberFormat = [["@", "@", "@", "@", "0.000", "@", "#,##0.00", "#,##0.00", "0.000000", "#,##0.00", "@", "@", "@", "@"]];
                tot.values = [totals];
                tot.format.font.bold = true;
                tot.format.fill.color = "#FFF2CC";
                // Entradas nuevas: un solo bloque
                if (snap.length > from) {
                    const rows = snap.slice(from).map(rowOf);
                    const blk = at(2 + from, rows.length);
                    blk.numberFormat = rows.map(() => FMT_ROW);
                    blk.values = rows;
                }
                // Entradas ya escritas cuyas estadísticas han llegado después
                changed.forEach(i => {
                    const r = at(2 + i, 1);
                    r.numberFormat = [FMT_ROW];
                    r.values = [rowOf(snap[i])];
                });
                await context.sync();
            });
            written = snap.length;
        } catch (e) {
            console.warn("[Draco] No se pudo volcar el diario de consultas:", e);
            fullRewrite = true;   // en el siguiente intento, entero
            dirty = true;
        } finally {
            flushing = false;
            if (dirty) schedule();
        }
    }

    /** Borra el diario (memoria y hoja). Se llama al desactivar el diagnóstico. */
    async function clear(context) {
        entries.length = 0;
        dirty = false;
        written = 0;
        fullRewrite = true;
        dirtyEntries.clear();
        const sh = context.workbook.worksheets.getItem(SHEET);
        const rng = sh.getRange(`${FIRST_COL}1:${LAST_COL}${MAX_ENTRIES + 2}`);
        rng.clear(Excel.ClearApplyTo.contents);
        rng.format.font.bold = false;
        rng.format.fill.clear();
    }

    window.DracoQueryJournal = { track, record, classify, clear, flush, statsFromText, RANGE: `${FIRST_COL}1:${LAST_COL}${MAX_ENTRIES + 2}` };
})();
