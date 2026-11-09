/**
 * ============================================================
 * DRACO — DIARIO DE CONSULTAS (Diagnóstico)
 * ============================================================
 * Con el Diagnóstico activo (cinta ▸ Diagnóstico), CADA consulta SQL que
 * lanza el add-in —informes, filtros, buscador y reconocimiento de
 * miembros, planificación, modelo semántico, editor de cálculos...— se
 * acumula en EDIT_REPORT a partir de la columna AD, la más reciente arriba:
 *
 *   AD Nº | AE Hora | AF Tipo | AG Motor | AH Duración (s) | AI Filas |
 *   AJ Procesado (MB) | AK Facturado (MB) | AL Coste est. (USD) | AM Caché |
 *   AN SQL | AO Resultado (recortado) | AP Error
 *
 * La fila 1 lleva los TOTALES (consultas, segundos, MB y coste) para ver
 * de un vistazo lo que ha consumido una sesión de trabajo. Las columnas
 * numéricas son números de verdad: se pueden sumar o filtrar.
 *
 * Volumen y coste:
 *  - BigQuery: bytes procesados y facturados que devuelve la propia API
 *    (con caché = 0). Si la respuesta no trae los facturados, se estiman
 *    (mínimo 10 MB, redondeo a MB) y se marcan con «≈». Coste = facturado
 *    × DracoConfig.bigqueryPricePerTiB (on-demand; con reservas/editions el
 *    coste real es la capacidad contratada, no este importe).
 *  - Databricks: no hay coste por consulta (se paga el tiempo encendido del
 *    warehouse): se anota la duración y las filas.
 *
 * Sin Diagnóstico no se guarda nada (ni en memoria ni en el libro). Al
 * desactivarlo se borra todo (auditoría F-15: el libro no debe llevar SQL
 * ni resultados si se comparte).
 */
(function () {
    const SHEET = "EDIT_REPORT";
    const FIRST_COL = "AD";
    const LAST_COL = "AP";
    const MAX_ENTRIES = 200;
    const SQL_MAX = 32000;       // límite de Excel por celda: 32 767
    const RESULT_MAX = 8000;     // el resultado completo sigue en Y1 (último informe) y en la consola
    const MB = 1024 * 1024;
    const TIB = 1024 * 1024 * 1024 * 1024;

    const entries = [];
    let seq = 0;
    let flushTimer = null;
    let flushing = false;
    let dirty = false;

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
            affected: num(/"numDmlAffectedRows"\s*:\s*"?(\d+)/, head)
        };
    }

    function record(e) {
        if (!enabled()) return;
        const isBq = e.engine === "BigQuery";
        let billed = e.billed, estimated = false;
        if (isBq && e.cacheHit === true) billed = 0;
        else if (isBq && (billed === null || billed === undefined) && e.processed > 0) {
            billed = Math.max(Math.ceil(e.processed / MB) * MB, 10 * MB);   // mínimo 10 MB por consulta
            estimated = true;
        }
        const cost = isBq && billed !== null && billed !== undefined ? billed / TIB * pricePerTiB() : null;
        const d = new Date();
        entries.unshift({
            n: ++seq,
            time: d.toTimeString().slice(0, 8),
            kind: e.kind || classify(e.sql),
            engine: e.engine || engine(),
            secs: Math.round((e.ms || 0)) / 1000,
            rows: e.rows === null || e.rows === undefined ? "" : e.rows,
            processed: e.processed === null || e.processed === undefined ? "" : Math.round(e.processed / MB * 100) / 100,
            billed: billed === null || billed === undefined ? "" : Math.round(billed / MB * 100) / 100,
            cost: cost === null ? "" : Math.round(cost * 1e6) / 1e6,
            cache: (e.cacheHit === true ? "Sí" : (e.cacheHit === false ? "No" : "")) + (estimated ? " (facturado ≈ estimado)" : ""),
            sql: String(e.sql || "").slice(0, SQL_MAX),
            result: e.result ? String(e.result).slice(0, RESULT_MAX) + (String(e.result).length > RESULT_MAX ? " …(recortado)" : "") : "",
            error: e.error ? String(e.error).slice(0, 2000) : ""
        });
        if (entries.length > MAX_ENTRIES) entries.length = MAX_ENTRIES;
        dirty = true;
        schedule();
    }

    /**
     * Ejecuta fn() (que lanza la consulta) y la anota en el diario.
     * opts: { kind, engine }. Si el diagnóstico está apagado, solo ejecuta.
     */
    async function track(sql, fn, opts) {
        if (!enabled()) return fn();
        const o = opts || {};
        const t0 = performance.now();
        const eng = o.engine || engine();
        try {
            const res = await fn();
            const ms = performance.now() - t0;
            if (typeof res === "string") {
                const s = statsFromText(res);
                record({ sql, kind: o.kind, engine: eng, ms, rows: s.affected !== null ? s.affected : s.rows,
                    processed: s.processed, billed: s.billed, cacheHit: s.cacheHit, result: res });
            } else {
                const rows = res && Array.isArray(res.rows) ? res.rows.length : null;
                let txt = "";
                try { txt = JSON.stringify(res && res.rows ? res.rows : res); } catch (e) { /* sin texto */ }
                record({ sql, kind: o.kind, engine: eng, ms, rows, result: txt });
            }
            return res;
        } catch (err) {
            record({ sql, kind: o.kind, engine: eng, ms: performance.now() - t0, error: (err && err.message) || String(err) });
            throw err;
        }
    }

    function schedule() {
        if (flushTimer) return;
        flushTimer = setTimeout(flush, 800);
    }

    async function flush() {
        flushTimer = null;
        if (flushing) { schedule(); return; }
        if (!dirty || !enabled()) return;
        if (typeof Excel === "undefined" || !Excel.run) return;   // diálogos: sin acceso al libro
        flushing = true;
        dirty = false;
        const snap = entries.slice();
        try {
            const sum = (k) => snap.reduce((a, e) => a + (typeof e[k] === "number" ? e[k] : 0), 0);
            const round = (x, d) => Math.round(x * Math.pow(10, d)) / Math.pow(10, d);
            const rows = [
                [`CONSULTAS (diagnóstico) · ${snap.length} · más reciente arriba`, "", "TOTAL", "", round(sum("secs"), 3), "",
                    round(sum("processed"), 2), round(sum("billed"), 2), round(sum("cost"), 6), "", "", "", ""],
                ["Nº", "Hora", "Tipo", "Motor", "Duración (s)", "Filas", "Procesado (MB)", "Facturado (MB)",
                    `Coste est. (USD, ${pricePerTiB()} $/TiB)`, "Caché", "SQL", "Resultado (recortado)", "Error"]
            ];
            snap.forEach(e => rows.push([e.n, e.time, e.kind, e.engine, e.secs, e.rows, e.processed, e.billed,
                e.cost, e.cache, e.sql, e.result, e.error]));
            const lastRow = MAX_ENTRIES + 2;
            while (rows.length < lastRow) rows.push(["", "", "", "", "", "", "", "", "", "", "", "", ""]);

            await Excel.run(async (context) => {
                const sh = context.workbook.worksheets.getItemOrNullObject(SHEET);
                sh.load("isNullObject");
                await context.sync();
                if (sh.isNullObject) return;
                const rng = sh.getRange(`${FIRST_COL}1:${LAST_COL}${lastRow}`);
                // Texto en las columnas de texto: que nada se interprete como fórmula o fecha
                const fmt = rows.map(() => ["0", "@", "@", "@", "0.000", "#,##0", "#,##0.00", "#,##0.00", "0.000000", "@", "@", "@", "@"]);
                fmt[0] = ["@", "@", "@", "@", "0.000", "@", "#,##0.00", "#,##0.00", "0.000000", "@", "@", "@", "@"];
                fmt[1] = fmt[1].map(() => "@");
                rng.numberFormat = fmt;
                rng.values = rows;
                const top = sh.getRange(`${FIRST_COL}1:${LAST_COL}2`);
                top.format.font.bold = true;
                sh.getRange(`${FIRST_COL}1:${LAST_COL}1`).format.fill.color = "#FFF2CC";
                await context.sync();
            });
        } catch (e) {
            console.warn("[Draco] No se pudo volcar el diario de consultas:", e);
        } finally {
            flushing = false;
            if (dirty) schedule();
        }
    }

    /** Borra el diario (memoria y hoja). Se llama al desactivar el diagnóstico. */
    async function clear(context) {
        entries.length = 0;
        dirty = false;
        const sh = context.workbook.worksheets.getItem(SHEET);
        const rng = sh.getRange(`${FIRST_COL}1:${LAST_COL}${MAX_ENTRIES + 2}`);
        rng.clear(Excel.ClearApplyTo.contents);
        rng.format.font.bold = false;
        rng.format.fill.clear();
    }

    window.DracoQueryJournal = { track, record, classify, clear, flush, statsFromText, RANGE: `${FIRST_COL}1:${LAST_COL}${MAX_ENTRIES + 2}` };
})();
