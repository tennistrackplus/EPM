/**
 * engine-api.js — API propia de Tablas Draco sobre el motor del addin (commands.js).
 * (Generado por tools/sync-addin.py desde tools/templates/; se carga DESPUÉS de commands.js.)
 *
 *   generateSql(reportId)  -> { sql, fixed, mode } Construye el SQL del informe SIN ejecutarlo (JS puro del addin:
 *                                              buildSQL / buildSQLFixed / buildSQLHybrid según el modo).
 *   refresh(reportId, sql) -> pinta el informe en la hoja. Con `sql` NO se vuelve a construir: ejecuta ese SQL
 *                             (parches "con SQL guardado" de tools/sync-addin.py) y pinta el resultado.
 *   usedRange()            -> { r1,c1,r2,c2 } de lo pintado en la hoja principal (o null).
 *
 * Todas usan funciones globales de commands.js (getEditReportGrid, loadReportDefinition, buildSQL...).
 */
(function () {
    // Sin indicador flotante (usa formas de Excel): el anfitrión ya muestra su propio "Actualizando…".
    window.BusyIndicator = window.BusyIndicator || { show: async () => {}, update: async () => {}, hide: async () => {} };

    const API = window.DracoEngineApi = {
        storedSql: null,        // lo lee commands.js (parches): si hay SQL guardado no se vuelve a generar

        activeReportId() { return activeReportIdOrNull(); },

        /** Modo del informe (mismo criterio que actualizarUnInforme(), ver dracoReportMode en commands.js):
         *  "fixed" (los dos ejes Estáticos), "hybrid" (uno) o "dynamic" (ninguno). */
        async mode(reportId) {
            return await Excel.run(async (context) => dracoReportMode(await getEditReportGrid(context, reportId)));
        },

        async isFixed(reportId) { return (await API.mode(reportId)) === "fixed"; },

        async generateSql(reportId) {
            reportId = activeReportIdOrNull();       // ver refresh()
            const mode = await API.mode(reportId);
            let sql = null;
            await Excel.run(async (context) => {
                const editReportGrid = await getEditReportGrid(context, reportId);
                const relGrid = await window.SemanticModelStore.getModelGrid("MODEL_RELATIONSHIP");
                const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");
                const atributesGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
                if (mode === "fixed" || mode === "hybrid") {
                    const resultSheetName = resultSheetNameFromGrid(editReportGrid, reportId);
                    await ensureDracoResultSheetExists(context, resultSheetName);
                    const csvGrid = await getFormulaGrid(context, resultSheetName);
                    loadReportDefinition(editReportGrid, reportId);
                    if (mode === "fixed") {
                        sql = await buildSQLFixed(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid);
                    } else {
                        const built = await buildSQLHybrid(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid, reportId);
                        sql = built.sql;
                    }
                } else {
                    loadReportDefinition(editReportGrid, reportId);
                    applyFilterMeasuresToDynamicReport();
                    const subtotalsOnTop = String(cellValue(editReportGrid, 4, 4)).trim().toUpperCase() === "X";
                    sql = buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop);
                }
            });
            return { sql: typeof sql === "string" ? sql : (sql && sql.text) || "", fixed: mode === "fixed", mode };
        },

        async refresh(reportId, sql) {
            // Una tabla Draco = UN informe: siempre el activo, con el MISMO tipo de id (número) que usa el addin al
            // atender un clic en el +/-. Si aquí se pasara "1" (texto) y allí 1 (número), el estado de jerarquías
            // contraídas se guardaría con claves distintas y el +/- no encontraría sus indicadores.
            reportId = activeReportIdOrNull();
            API.storedSql = sql || null;
            try { await actualizarUnInforme(reportId); }
            finally { API.storedSql = null; }
        },

        usedRange() {
            const h = window.frameElement && window.frameElement.__dracoHost;
            if (!h) return null;
            const keys = Object.keys(h.state.cells || {});
            if (!keys.length) return null;
            let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
            keys.forEach((k) => {
                const o = h.state.cells[k]; if (!o || (o.v === undefined || o.v === "")) return;
                const [r, c] = k.split("_").map(Number);
                if (r < r1) r1 = r; if (c < c1) c1 = c; if (r > r2) r2 = r; if (c > c2) c2 = c;
            });
            return r2 < 0 ? null : { r1, c1, r2, c2 };
        }
    };
})();
