/**
 * ReportStore
 * ------------------------------------------------------------------------
 * Permite tener VARIOS informes en el mismo libro (antes solo existía uno,
 * cuyo diseño se escribía directamente en la hoja EDIT_REPORT).
 *
 * Todo lo que antes se escribía en EDIT_REPORT como "diseño" (filtros,
 * filas, columnas, Estático/Dinámico, rangos H10/N10, opciones de campo,
 * propiedades del informe) se guarda ahora en UNA sola clave JSON dentro
 * de Office.context.document.settings — el mismo mecanismo de "roaming
 * settings" que ya usa SemanticModelStore — así que no aparece nunca como
 * hoja ni pestaña oculta.
 *
 * La hoja EDIT_REPORT se SIGUE creando y usando (ver semantic_model.js /
 * ensureTechnicalSheets), pero ahora solo para lo que de verdad tiene que
 * vivir en una celda física:
 *   - D1 / E1: hoja y celda activas al pulsar "Editar informe" (para saber
 *     dónde estaba el usuario).
 *   - X1 / Y1: último SQL y último JSON devueltos por BigQuery/Snowflake.
 *   - A5 y el resto del "picker" de miembros: comunicación por doble clic
 *     con el XLAM.
 *   - B1: flag de reconocimiento de miembros.
 *
 * Para que el resto de commands.js (que lee EDIT_REPORT con coordenadas
 * fila/columna fijas: H10, N10, H12, N12, C15:F, H15:M, N15:S) siga
 * funcionando sin reescribir toda esa lógica, getReportGrid() genera un
 * "grid" {values, startRow:0, startCol:0} con esas mismas coordenadas,
 * construido a partir del JSON. commands.js combina ese grid con la
 * lectura física de EDIT_REPORT (ver getEditReportGrid en commands.js):
 * las coordenadas de diseño las gana siempre el JSON; el resto (D1, E1,
 * X1, Y1, B1, A5...) lo gana la hoja física.
 */

(function () {

    const RS_KEY = "epm_reports";              // JSON: { [reportId]: reportObj }
    const RS_ACTIVE_KEY = "epm_activeReportId"; // id del informe activo
    const RS_SEQ_KEY = "epm_reportSeq";         // último id secuencial usado

    const DEFAULT_REPORT_PROPERTIES = {
        reportName: "",
        suppressZeroRows: false,
        suppressZeroCols: false,
        subtotalsOnTop: false,
        overwriteFormats: true,
        autoFitColumns: true
    };

    /* ---------------------------------------------------------------
     * Persistencia bruta (Office roaming settings)
     * ------------------------------------------------------------- */

    function _readStoreRaw() {
        try {
            const raw = Office.context.document.settings.get(RS_KEY);
            if (!raw) return {};
            const parsed = JSON.parse(raw);
            return (parsed && typeof parsed === "object") ? parsed : {};
        } catch (e) {
            console.error("ReportStore: JSON de epm_reports corrupto, se reinicia.", e);
            return {};
        }
    }

    function _saveAsync() {
        return new Promise((resolve, reject) => {
            Office.context.document.settings.saveAsync((asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    reject(asyncResult.error);
                } else {
                    resolve();
                }
            });
        });
    }

    async function _writeStoreRaw(storeObj) {
        Office.context.document.settings.set(RS_KEY, JSON.stringify(storeObj));
        await _saveAsync();
    }

    /* ---------------------------------------------------------------
     * Utilidades de grid (copia local, igual que en commands.js/
     * excelService.js) para poder reutilizar MODEL_HIER al expandir
     * jerarquías en filas/columnas.
     * ------------------------------------------------------------- */

    function cellValue(grid, row1, col1) {
        const r = row1 - 1 - grid.startRow;
        const c = col1 - 1 - grid.startCol;
        if (r < 0 || c < 0 || r >= grid.values.length) return "";
        const rowArr = grid.values[r];
        if (!rowArr || c >= rowArr.length) return "";
        const v = rowArr[c];
        return v === null || v === undefined ? "" : v;
    }

    function lastRowInColumnValues(grid, col1) {
        const c = col1 - 1 - grid.startCol;
        for (let r = grid.values.length - 1; r >= 0; r--) {
            const rowArr = grid.values[r];
            const v = rowArr && c >= 0 && c < rowArr.length ? rowArr[c] : "";
            if (v !== null && v !== undefined && String(v) !== "") {
                return r + 1 + grid.startRow;
            }
        }
        return grid.startRow;
    }

    function existsHierarchy(hierGrid, dimension, hierarchy) {
        const lastRow = lastRowInColumnValues(hierGrid, 1);
        for (let R = 2; R <= lastRow; R++) {
            if (String(cellValue(hierGrid, R, 4)).toUpperCase() === String(dimension).toUpperCase()
                && String(cellValue(hierGrid, R, 2)).toUpperCase() === String(hierarchy).toUpperCase()) {
                return true;
            }
        }
        return false;
    }

    /* ---------------------------------------------------------------
     * API pública: informe a informe
     * ------------------------------------------------------------- */

    function getAllReports() {
        return _readStoreRaw();
    }

    // Lista ordenada por id de creación (orden en el que se fueron
    // añadiendo), tal y como se deben mostrar en el listbox "Informe".
    function listReports() {
        const store = _readStoreRaw();
        return Object.keys(store)
            .map(id => Number(id))
            .sort((a, b) => a - b)
            .map(id => {
                const r = store[String(id)];
                return { id, name: r.name, semanticModelName: r.semanticModelName || "" };
            });
    }

    function getReport(reportId) {
        const store = _readStoreRaw();
        return store[String(reportId)] || null;
    }

    async function _writeReport(reportId, reportObj) {
        const store = _readStoreRaw();
        store[String(reportId)] = reportObj;
        await _writeStoreRaw(store);
    }

    function _nextSequence() {
        const raw = Office.context.document.settings.get(RS_SEQ_KEY);
        const next = (Number(raw) || 0) + 1;
        Office.context.document.settings.set(RS_SEQ_KEY, next);
        return next;
    }

    function _pad3(n) {
        return String(n).padStart(3, "0");
    }

    /**
     * Crea un informe nuevo y VACÍO (sin filtros/filas/columnas), lo marca
     * como activo y lo devuelve. El id interno es el secuencial (nunca se
     * reutiliza, aunque se borren informes). El nombre por defecto es
     * "Informe - 0XX".
     *
     * resultSheetName (opcional): hoja donde se va a pintar ESTE informe.
     * El add-in YA NO crea una pestaña nueva por informe: addReport() en
     * taskpane.js captura la hoja activa de Excel en el momento de pulsar
     * "Añadir informe" (ver captureActiveEditContext) y la pasa aquí, así
     * que el informe se queda anclado en la pestaña/celda donde estaba el
     * cursor. Si no se recibe (compatibilidad con llamadas antiguas), se
     * guarda vacío y el código que lee la hoja de resultados
     * (resultSheetNameFromGrid / getDracoResultSheetName en commands.js)
     * usa como último recurso la celda física compartida EDIT_REPORT!D1.
     */
    async function createReport(semanticModelName, resultSheetName) {
        const id = _nextSequence();
        const name = `Informe - ${_pad3(id)}`;
        const report = {
            id,
            name,
            semanticModelName: semanticModelName || "",
            design: {
                filters: [],
                rows: [],
                columns: [],
                expandedRows: [],
                expandedCols: [],
                rowsStatic: false,
                colsStatic: false,
                rrAddress: "",
                rcAddress: "",
                resultSheetName: resultSheetName || "",
                fieldOptions: {}
            },
            reportProperties: Object.assign({}, DEFAULT_REPORT_PROPERTIES, { reportName: name }),
            memberRecognition: false
        };
        await _writeReport(id, report);
        await _saveAsync();
        await setActiveReportId(id);
        return report;
    }

    /**
     * Hoja de resultados guardada para ESTE informe ("" si es un informe
     * antiguo, creado antes de que esto se guardara por informe: en ese
     * caso el llamante debe usar el fallback físico de EDIT_REPORT!D1, ver
     * resultSheetNameFromGrid en commands.js).
     */
    function getReportResultSheetName(reportId) {
        const report = getReport(reportId);
        return report && report.design ? (report.design.resultSheetName || "") : "";
    }

    /**
     * Fija la hoja de resultados de un informe concreto (llamado desde
     * "Editar report" en el taskpane, cuando el usuario reubica manualmente
     * dónde se pinta). Queda guardada solo para ESE informe: ya no se
     * comparte con el resto a través de una única celda física.
     */
    async function setReportResultSheetName(reportId, sheetName) {
        const report = getReport(reportId);
        if (!report) return;
        report.design = report.design || {};
        report.design.resultSheetName = sheetName || "";
        await _writeReport(reportId, report);
    }

    async function deleteReport(reportId) {
        const store = _readStoreRaw();
        delete store[String(reportId)];
        await _writeStoreRaw(store);

        if (Number(Office.context.document.settings.get(RS_ACTIVE_KEY)) === Number(reportId)) {
            const remaining = Object.keys(store).map(Number).sort((a, b) => a - b);
            await setActiveReportId(remaining[0] || null);
        }
    }

    async function renameReport(reportId, newName) {
        const report = getReport(reportId);
        if (!report) return;
        report.name = newName;
        report.reportProperties = report.reportProperties || {};
        report.reportProperties.reportName = newName;
        await _writeReport(reportId, report);
    }

    /**
     * Informe "activo": el que el taskpane está editando ahora mismo. Si el
     * marcado como activo ya no existe, se usa el primero (por orden de
     * creación) que quede, o null si no hay ninguno (taskpane vacío).
     */
    function getActiveReportId() {
        const raw = Office.context.document.settings.get(RS_ACTIVE_KEY);
        const active = raw ? Number(raw) : null;
        const store = _readStoreRaw();
        if (active && store[String(active)]) return active;

        const ids = Object.keys(store).map(Number).sort((a, b) => a - b);
        return ids.length > 0 ? ids[0] : null;
    }

    async function setActiveReportId(reportId) {
        Office.context.document.settings.set(RS_ACTIVE_KEY, reportId || "");
        await _saveAsync();
    }

    function getActiveReport() {
        const id = getActiveReportId();
        return id ? getReport(id) : null;
    }

    // Propiedades del informe activo (o los valores por defecto si todavía
    // no hay ningún informe). Sustituye a la antigua clave global
    // "draco_reportProperties": accesible en sync desde cualquier
    // contexto (taskpane o commands.html/ribbon).
    function getActiveReportProperties() {
        const report = getActiveReport();
        if (!report) return Object.assign({}, DEFAULT_REPORT_PROPERTIES);
        return Object.assign({}, DEFAULT_REPORT_PROPERTIES, report.reportProperties || {});
    }

    async function saveReportProperties(reportId, props) {
        const report = getReport(reportId);
        if (!report) return;
        report.reportProperties = Object.assign({}, DEFAULT_REPORT_PROPERTIES, report.reportProperties, props);
        if (props && props.reportName) report.name = props.reportName;
        await _writeReport(reportId, report);
    }

    async function setSemanticModel(reportId, semanticModelName) {
        const report = getReport(reportId);
        if (!report) return;
        report.semanticModelName = semanticModelName || "";
        await _writeReport(reportId, report);
    }

    /* ---------------------------------------------------------------
     * Diseño (filtros/filas/columnas/Estático/rangos/opciones de campo):
     * expande jerarquías (vía MODEL_HIER) EN EL MOMENTO DE GUARDAR, igual
     * que hacía antes excelService.saveEditReportDesign al escribir en la
     * hoja, para que getReportGrid() no tenga que repetir el trabajo cada
     * vez que se lee.
     * ------------------------------------------------------------- */

    async function saveDesign(reportId, state) {
        const report = getReport(reportId);
        if (!report) return;

        const hierGrid = await window.SemanticModelStore.getModelGrid("MODEL_HIER");

        const subtotalAndOrder = (zoneId, dimension, name, fieldOptions) => {
            const opts = (fieldOptions || {})[`${zoneId}|${dimension}|${name}`] || {};
            const subtotal = opts.showTotals ? "X" : "";
            const order = opts.sortOrder === "asc" ? "UP" : (opts.sortOrder === "desc" ? "DOWN" : "");
            return [subtotal, order];
        };

        const visibleLevelsFor = (zoneId, dimension, name, fieldOptions) => {
            const opts = (fieldOptions || {})[`${zoneId}|${dimension}|${name}`] || {};
            return Array.isArray(opts.visibleLevels) ? opts.visibleLevels : null;
        };

        const expandAxis = (list, zoneId, fieldOptions) => {
            const out = [];
            list.forEach(item => {
                if (item.isHierarchy && existsHierarchy(hierGrid, item.dimension, item.name)) {
                    const visibleLevels = visibleLevelsFor(zoneId, item.dimension, item.name, fieldOptions);
                    const lastRow = lastRowInColumnValues(hierGrid, 4);
                    for (let R = 2; R <= lastRow; R++) {
                        if (String(cellValue(hierGrid, R, 4)).toUpperCase() === item.dimension.toUpperCase()
                            && String(cellValue(hierGrid, R, 2)).toUpperCase() === item.name.toUpperCase()) {
                            const nivel = Number(cellValue(hierGrid, R, 3));
                            if (visibleLevels && !visibleLevels.includes(nivel)) continue;
                            out.push([
                                cellValue(hierGrid, R, 4), // DIMENSION
                                cellValue(hierGrid, R, 5), // ATRIBUTO
                                cellValue(hierGrid, R, 3), // NIVEL
                                cellValue(hierGrid, R, 2), // JERARQUIA
                                "", ""                     // subtotal/orden: no aplica a jerarquías
                            ]);
                        }
                    }
                } else {
                    const [subtotal, order] = subtotalAndOrder(zoneId, item.dimension, item.name, fieldOptions);
                    out.push([item.dimension, item.name, 1, "", subtotal, order]);
                }
            });
            return out;
        };

        report.design = {
            filters: (state.filters || []).map(f => ({
                dimension: f.dimension, name: f.name, isHierarchy: f.isHierarchy,
                realAttribute: f.realAttribute,
                // El filtro (multi-valor / rango / incluir-excluir) se guarda
                // como JSON en la misma celda "Valor" que antes tenía un
                // valor simple; commands.js (parseFilterValue) sabe leer
                // ambos formatos.
                value: f.filter ? JSON.stringify(f.filter) : ""
            })),
            rows: (state.rows || []).map(r => ({ dimension: r.dimension, name: r.name, isHierarchy: r.isHierarchy })),
            columns: (state.columns || []).map(c => ({ dimension: c.dimension, name: c.name, isHierarchy: c.isHierarchy })),
            expandedRows: expandAxis(state.rows || [], "rows", state.fieldOptions),
            expandedCols: expandAxis(state.columns || [], "columns", state.fieldOptions),
            rowsStatic: !!state.rowsStatic,
            colsStatic: !!state.colsStatic,
            rrAddress: state.rrAddress || "",
            rcAddress: state.rcAddress || "",
            fieldOptions: state.fieldOptions || {}
        };

        await _writeReport(reportId, report);
    }

    async function setMemberRecognition(reportId, value) {
        const report = getReport(reportId);
        if (!report) return;
        report.memberRecognition = !!value;
        await _writeReport(reportId, report);
    }

    /**
     * Genera el grid virtual {values, startRow:0, startCol:0} con EXACTAMENTE
     * las mismas coordenadas que tenía el diseño en EDIT_REPORT:
     *   H10 / N10  -> rrAddress / rcAddress
     *   H12 / N12  -> rowsStatic / colsStatic ("X"/"")
     *   C15:F..    -> filtros (Dimension, Atributo real, Valor, Jerarquía)
     *   H15:M..    -> filas expandidas (Dim, Atributo, Nivel, Jerarquía, Subtotal, Orden)
     *   N15:S..    -> columnas expandidas (Dim, Atributo, Nivel, Jerarquía, Subtotal, Orden)
     * commands.js combina este grid con la lectura física de EDIT_REPORT
     * (ver getEditReportGrid/isDesignOwnedCell) para las coordenadas que
     * SÍ siguen viviendo en la hoja (D1/E1/X1/Y1/B1/A5...).
     */
    function getReportGrid(reportId) {
        const report = getReport(reportId);
        const design = report ? (report.design || {}) : {};

        const filters = design.filters || [];
        const expandedRows = design.expandedRows || [];
        const expandedCols = design.expandedCols || [];

        const maxRow = 14 + Math.max(filters.length, expandedRows.length, expandedCols.length, 0);
        const values = [];
        for (let r = 0; r < maxRow; r++) values.push(new Array(19).fill(""));

        const set = (row1, col1, v) => {
            if (row1 - 1 < values.length) values[row1 - 1][col1 - 1] = (v === null || v === undefined) ? "" : v;
        };

        set(10, 8, design.rrAddress || "");
        set(10, 14, design.rcAddress || "");
        set(12, 8, design.rowsStatic ? "X" : "");
        set(12, 14, design.colsStatic ? "X" : "");

        filters.forEach((f, i) => {
            const row = 15 + i;
            set(row, 3, f.dimension || "");
            set(row, 4, f.realAttribute || "");
            set(row, 5, f.value || "");
            set(row, 6, f.isHierarchy ? f.name : "");
        });

        expandedRows.forEach((cols, i) => {
            const row = 15 + i;
            set(row, 8, cols[0]); set(row, 9, cols[1]); set(row, 10, cols[2]); set(row, 11, cols[3]);
            set(row, 12, cols[4]); set(row, 13, cols[5]);
        });

        expandedCols.forEach((cols, i) => {
            const row = 15 + i;
            set(row, 14, cols[0]); set(row, 15, cols[1]); set(row, 16, cols[2]); set(row, 17, cols[3]);
            set(row, 18, cols[4]); set(row, 19, cols[5]);
        });

        return { values, startRow: 0, startCol: 0 };
    }

    window.ReportStore = {
        getAllReports,
        listReports,
        getReport,
        createReport,
        deleteReport,
        renameReport,
        getActiveReportId,
        setActiveReportId,
        getActiveReport,
        getActiveReportProperties,
        saveReportProperties,
        setSemanticModel,
        saveDesign,
        setMemberRecognition,
        getReportGrid,
        getReportResultSheetName,
        setReportResultSheetName
    };

})();
