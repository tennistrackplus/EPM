/**
 * Lógica principal del TaskPane — Diseñador de informes
 * Traducción de frmReportDesigner2 (VBA) a Office.js
 */
if (typeof Office !== "undefined") {
    Office.onReady(() => {
        TaskPaneApp.init();
    });
} else {
    document.addEventListener("DOMContentLoaded", () => {
        TaskPaneApp.init();
    });
}


/* ---------------------------------------------------------------------
 * RangeAxis — traducción de RR_Load/RC_Load/RR_Refresh/RC_Refresh/
 * RR_Move/RC_Move. Mantiene los dos rangos (Filas / Columnas) y sus
 * reglas de colisión, igual que el VBA.
 * ------------------------------------------------------------------- */
const RangeAxis = {
    rr: { row: 1, col: 1, height: 1, width: 0 },   // eje Filas
    rc: { row: 1, col: 1, height: 0, width: 1 },   // eje Columnas

    loadFromAddresses(rrAddress, rcAddress, anchorAddress) {
        // Ancla de la PRIMERA vez (EDIT_REPORT!E1): solo se usa como
        // fallback cuando el eje correspondiente todavía no tiene una
        // dirección guardada (rrAddress/rcAddress vacíos), es decir, la
        // primera vez que se monta el informe. Si E1 no tiene una
        // dirección válida, se mantiene A1 (comportamiento anterior).
        const anchor = window.ReportDesignerUtils.parseAddress(anchorAddress) || { row: 1, col: 1 };

        this.rr = this._parseRange(rrAddress, { row: anchor.row, col: anchor.col, height: 1, width: 0 });
        this.rc = this._parseRange(rcAddress, { row: anchor.row, col: anchor.col, height: 0, width: 1 });
    },

    _parseRange(address, fallback) {
        if (!address) return { ...fallback };
        const parts = address.split(":");
        const p1 = window.ReportDesignerUtils.parseAddress(parts[0]);
        const p2 = window.ReportDesignerUtils.parseAddress(parts[1] || parts[0]);
        if (!p1 || !p2) return { ...fallback };
        return {
            row: p1.row,
            col: p1.col,
            height: p2.row - p1.row + 1,
            width: p2.col - p1.col + 1
        };
    },

    addressOf(axis) {
        const o = axis === "rr" ? this.rr : this.rc;
        if (!o.row || !o.col || !o.width || !o.height) return "";
        const from = window.ReportDesignerUtils.addressFromRC(o.row, o.col);
        const to = window.ReportDesignerUtils.addressFromRC(o.row + o.height - 1, o.col + o.width - 1);
        return from + ":" + to;
    },

    moveRR(dRow, dCol) {
        this.rr.row += dRow;
        this.rr.col += dCol;
        if (this.rr.row < (this.rc.row + this.rc.height)) this.rr.row = this.rc.row + this.rc.height;
        if (this.rr.col < 1) this.rr.col = 1;
    },

    moveRC(dRow, dCol) {
        this.rc.row += dRow;
        this.rc.col += dCol;
        if (this.rc.row < 1) this.rc.row = 1;
        if (this.rc.col < (this.rr.col + this.rr.width)) this.rc.col = this.rr.col + this.rr.width;
    },

    // lblRRight_Click: mover filas a la derecha empuja columnas si se solapan
    moveRRRight() {
        this.moveRR(0, 1);
        if (this.rr.col + this.rr.width > this.rc.col) this.moveRC(1, 0);
    },

    // lblcDown_Click: mover columnas abajo empuja filas si se solapan
    moveRCDown() {
        this.moveRC(1, 0);
        if (this.rc.row + this.rc.height > this.rr.row) this.moveRR(1, 0);
    },

    // MODO HÍBRIDO (un eje Estático): las líneas del eje Estático son
    // celdas fijas de la hoja (fórmulas EPM_VALUE del usuario), así que NO
    // se pueden desplazar. Cuando el otro eje (Dinámico) gana o pierde un
    // campo, su cabecera crece/encoge HACIA FUERA (Columnas hacia arriba,
    // Filas hacia la izquierda) en vez de empujar al eje Estático. Si ya
    // está pegada al borde (fila/columna 1), se hace como siempre.

    // Al añadir un campo al eje Filas: crece en anchura y empuja columnas
    onRowFieldAdded(colsStatic) {
        if (colsStatic && this.rr.col > 1) {
            this.rr.col -= 1;
            this.rr.width += 1;
            return;
        }
        this.rr.width += 1;
        this.moveRC(0, 1);
    },

    // Al quitar un campo del eje Filas
    onRowFieldRemoved(colsStatic) {
        if (colsStatic && this.rr.width > 0) {
            this.rr.width -= 1;
            this.rr.col += 1;
            return;
        }
        this.rr.width = Math.max(0, this.rr.width - 1);
        this.moveRC(0, -1);
    },

    // Al añadir un campo al eje Columnas: crece en altura y empuja filas
    onColFieldAdded(rowsStatic) {
        if (rowsStatic && this.rc.row > 1) {
            this.rc.row -= 1;
            this.rc.height += 1;
            return;
        }
        this.rc.height += 1;
        this.moveRR(1, 0);
    },

    // Al quitar un campo del eje Columnas
    onColFieldRemoved(rowsStatic) {
        if (rowsStatic && this.rc.height > 0) {
            this.rc.height -= 1;
            this.rc.row += 1;
            return;
        }
        this.rc.height = Math.max(0, this.rc.height - 1);
        this.moveRR(-1, 0);
    }
};

const TaskPaneApp = {
    draggedElementData: null,

    // Id (secuencial, interno) del informe que se está editando ahora
    // mismo en el taskpane. null = todavía no hay ningún informe creado
    // (taskpane vacío, ver showEmptyState()).
    currentReportId: null,

    // Estado del diseño (equivalente a lstFilters/lstRows/lstCols del VBA)
    state: {
        filters: [],
        rows: [],
        columns: [],
        rowsStatic: false,
        colsStatic: false,
        // Opciones por campo (mostrar totales, orden, expandir hasta nivel,
        // niveles visibles, formato de medida...), clave "zona|dim|nombre".
        fieldOptions: {}
    },

    // Propiedades generales del informe (modal "Propiedades del informe"),
    // guardadas en Office roaming settings (visibles también desde el ribbon).
    reportProperties: {
        reportName: "Report 001",
        suppressZeroRows: false,
        suppressZeroCols: false,
        subtotalsOnTop: false,
        hierarchyNodesOnTop: true,
        overwriteFormats: true,
        overwriteStaticFormats: false,
        autoFitColumns: false,
        planningReport: false
    },

    // Columnas/filas físicas extra que ya están contadas en el rango de cada eje
    // por los "ID y texto" (ver reconcileTextColumns).
    textExtra: { rows: 0, columns: 0 },

    // Campo actualmente seleccionado en el panel "Opciones de campo"
    // ({ zoneId, dimension, name, isHierarchy }) o null si no hay ninguno.
    selectedFieldForOptions: null,

    /* -------------------------------------------------------------
     * Autoguardado + autoactualización: cada cambio estructural en el
     * taskpane (añadir/quitar campo, elegir valor de filtro, marcar
     * Estático/Dinámico, mover rangos) guarda el diseño en EDIT_REPORT
     * y dispara Actualizar() (que internamente llama a jsonTo3Matrices
     * cuando el eje es Dinámico), sin que el usuario tenga que pulsar
     * "Guardar" ni el botón del ribbon.
     * ----------------------------------------------------------- */
    autoRefreshTimer: null,
    isAutoRefreshing: false,
    autoRefreshQueued: false,

    setAutoStatus(text) {
        const el = document.getElementById("autoStatus");
        if (el) el.innerText = text;
    },

    // Payload común para SaveEditReportDesign (autoguardado, guardado manual
    // y el guardado "solo guardar, sin actualizar" del botón Estático).
    collectDesignPayload() {
        return {
            filters: this.state.filters,
            rows: this.state.rows,
            columns: this.state.columns,
            rowsStatic: this.state.rowsStatic,
            colsStatic: this.state.colsStatic,
            fieldOptions: this.state.fieldOptions,
            rrAddress: RangeAxis.addressOf("rr"),
            rcAddress: RangeAxis.addressOf("rc")
        };
    },

    /**
     * "Filtrar valores seleccionados" (clic derecho en celdas del informe):
     * pone en Filtros, para cada campo {dim, attr, values}, la lista de
     * valores (sustituye la selección si el campo ya estaba en Filtros).
     */
    async setFilterValues(reportId, items) {
        if (!this.currentReportId || String(this.currentReportId) !== String(reportId)) {
            this.setAutoStatus("Abre ese informe en el panel para filtrar sus valores.");
            return;
        }
        items.forEach(({ dim, attr, values }) => {
            const filter = { mode: "list", attribute: attr, values: values.slice(), ranges: [], excludeValues: [], excludeRanges: [] };
            let entry = this.state.filters.find(f => !f.isHierarchy
                && String(f.dimension).toUpperCase() === String(dim).toUpperCase()
                && String(f.realAttribute || f.name).toUpperCase() === String(attr).toUpperCase());
            if (entry) {
                entry.filter = filter;
            } else {
                this.state.filters.push({ dimension: dim, name: attr, isHierarchy: false, realAttribute: attr, filter });
            }
        });
        const filtersContent = document.querySelector('.zone-card[data-zone="filters"] .dropzone-content');
        if (filtersContent) {
            filtersContent.innerHTML = "";
            this.state.filters.forEach(f => this.renderTag(filtersContent, "filters", f));
            if (this.renderLockedFilterTags) this.renderLockedFilterTags(filtersContent);
        }
        this.scheduleAutoUpdate();
    },

    /**
     * Clic derecho en una celda del informe (ver dracoFieldActionFromCell en
     * commands.js): aplica una acción al campo de esa celda, cambiando antes
     * de informe si hace falta. action: "totals:show|hide", "mode:ID|ID_TEXT|TEXT",
     * "options" (abre Opciones de campo: niveles y expandir de una jerarquía).
     */
    async fieldActionFromCell(target, action) {
        if (!this.currentReportId || String(this.currentReportId) !== String(target.reportId)) {
            const sel = document.getElementById("reportSelector") || document.getElementById("reportSelect");
            if (sel) sel.value = String(target.reportId);
            await this.onReportSelectorChange(String(target.reportId));
        }
        const zone = target.axis === "columns" ? "columns" : "rows";
        const list = this.listForZone(zone) || [];
        const up = (x) => String(x || "").toUpperCase();

        // Borrar filtro: elimina de la zona Filtros los filtros de la
        // dimensión de la celda (cualquier atributo o jerarquía de esa dimensión).
        if (action === "filter:clear") {
            const hits = this.state.filters.filter(f => up(f.dimension) === up(target.dim));
            if (!hits.length) { this.setAutoStatus("Este campo no tiene filtro."); return; }
            this.state.filters = this.state.filters.filter(f => hits.indexOf(f) === -1);
            this.rerenderFiltersZone();
            this.scheduleAutoUpdate();
            this.setAutoStatus("Filtro de " + target.dim + " borrado.");
            return;
        }
        const sameDim = list.filter(e => up(e.dimension) === up(target.dim));
        let entry = sameDim.find(e => e.isHierarchy && up(e.name) === up(target.attr))
            || sameDim.find(e => !e.isHierarchy && up(e.name) === up(target.attr))
            || sameDim.find(e => !e.isHierarchy && up(e.name) === up(target.levelAttr))
            || (sameDim.length === 1 ? sameDim[0] : null);
        if (!entry) { this.setAutoStatus("No se encontró el campo de la celda en el diseño."); return; }

        if (action === "options") {
            await this.selectFieldForOptions(zone, entry);
        } else if (action.indexOf("mode:") === 0) {
            const info = window.ReportStore.getDisplayTextInfo(this.currentReportId, entry.dimension, entry.name, !!entry.isHierarchy);
            const mode = action.slice(5);
            if (!info || !info.applicable) { this.setAutoStatus("Este campo no tiene texto asociado."); return; }
            if (mode !== "ID" && !info.available) { this.setAutoStatus(info.reason); return; }
            this.setDisplayMode(zone, entry, mode);
        } else if (action.indexOf("totals:") === 0) {
            if (entry.isHierarchy) { this.setAutoStatus("Una jerarquía no tiene totales: usa Opciones del campo."); return; }
            const cur = this.getFieldOptions(zone, entry);
            cur.showTotals = action === "totals:show";
            this.setFieldOptions(zone, entry, cur);
            const selF = this.selectedFieldForOptions;
            if (selF && selF.zoneId === zone && selF.dimension === entry.dimension && selF.name === entry.name) this.renderFieldOptionsBody(zone, entry);
            this.scheduleAutoUpdate();
        }
    },

    /**
     * Autoguardado del diseño + actualización de datos (700 ms después del
     * último cambio). opts.saveOnly: el cambio no altera el informe (p. ej.
     * soltar un campo en Filtros sin valor todavía): solo se guarda el
     * diseño, sin consultar ni repintar. Si mientras tanto llega otro cambio
     * que sí lo altera, se actualiza igualmente.
     */
    scheduleAutoUpdate(opts) {
        if (!this.currentReportId) return; // sin informe activo todavía: nada que guardar
        if (!(opts && opts.saveOnly)) this.autoNeedsRefresh = true;
        this.setAutoStatus("Cambios pendientes…");
        if (this.autoRefreshTimer) clearTimeout(this.autoRefreshTimer);
        this.autoRefreshTimer = setTimeout(() => this.runAutoSaveAndRefresh(), 700);
    },

    async runAutoSaveAndRefresh() {
        if (!this.currentReportId) return;
        if (this.isAutoRefreshing) {
            this.autoRefreshQueued = true;
            return;
        }
        this.isAutoRefreshing = true;
        const needRefresh = !!this.autoNeedsRefresh;
        this.autoNeedsRefresh = false;

        try {
            this.setAutoStatus("Guardando…");
            await window.ExcelService.saveEditReportDesign(this.currentReportId, this.collectDesignPayload());

            if (!needRefresh) {
                // Cambio que no altera el informe: guardado, sin actualizar
                this.setAutoStatus("");
                return;
            }

            if (this.isRefreshPaused()) {
                // Pausa activa (botón de pausa marcado): se guarda el diseño
                // igualmente, pero NO se actualizan los datos (no se llama a
                // Actualizar()/BigQuery) hasta que se desmarque la pausa.
                this.setAutoStatus("Pausado (sin actualizar)");
                return;
            }

            if (window.ReportActions && typeof window.ReportActions.actualizar === "function") {
                this.setAutoStatus("Actualizando…");
                if (window.BusyIndicator) await window.BusyIndicator.show("Actualizando");
                try {
                    await window.ReportActions.actualizar();
                } finally {
                    if (window.BusyIndicator) await window.BusyIndicator.hide();
                }
            }

            this.setAutoStatus("Actualizado ✓");
            setTimeout(() => {
                const el = document.getElementById("autoStatus");
                if (el && el.innerText === "Actualizado ✓") el.innerText = "";
            }, 2000);
        } catch (err) {
            console.error("Error en el autoguardado/autoactualización:", err);
            this.setAutoStatus("Error al actualizar");
        } finally {
            this.isAutoRefreshing = false;
            if (this.autoRefreshQueued) {
                this.autoRefreshQueued = false;
                this.scheduleAutoUpdate({ saveOnly: true }); // conserva si hacía falta actualizar
            }
        }
    },

    /**
     * Botón "Pausar actualización de datos" del TASKPANE. Es
     * completamente independiente del botón de pausa del ribbon (no
     * comparte clave de ajuste ni depende del manifest): usa su propia
     * clave en Office roaming settings solo para recordar el estado
     * entre aperturas del taskpane. Mientras está marcado,
     * runAutoSaveAndRefresh() sigue guardando el diseño pero NO llama a
     * Actualizar() (no se refrescan los datos).
     */
    isRefreshPaused() {
        try {
            const settings = Office.context && Office.context.document && Office.context.document.settings;
            return !!(settings && settings.get("draco_taskpane_refreshPaused"));
        } catch (e) {
            return false;
        }
    },

    setRefreshPausedUI(isPaused) {
        const btn = document.getElementById("btnPauseRefresh");
        if (!btn) return;
        btn.classList.toggle("toggle-active", !!isPaused);
        btn.setAttribute("aria-pressed", isPaused ? "true" : "false");
        btn.title = isPaused ? "Actualización de datos PAUSADA (clic para reanudar)" : "Pausar actualización de datos";
    },

    initRefreshPauseButton() {
        this.setRefreshPausedUI(this.isRefreshPaused());
        const btn = document.getElementById("btnPauseRefresh");
        if (!btn) return;
        btn.addEventListener("click", () => {
            try {
                const settings = Office.context.document.settings;
                const nowPaused = !this.isRefreshPaused();
                settings.set("draco_taskpane_refreshPaused", nowPaused);
                this.setRefreshPausedUI(nowPaused); // feedback inmediato, sin esperar saveAsync
                settings.saveAsync();
            } catch (e) {
                console.warn("No se pudo alternar la pausa de actualización:", e);
            }
        });
    },


    // Guarda el diseño SIN disparar Actualizar() (BigQuery). Se usa al
    // marcar/desmarcar Estático: el propio botón pide explícitamente que
    // eso no dispare un refresco automático.
    async saveDesignOnly() {
        if (!this.currentReportId) return;
        try {
            this.setAutoStatus("Guardando…");
            await window.ExcelService.saveEditReportDesign(this.currentReportId, this.collectDesignPayload());
            this.setAutoStatus("Guardado ✓");
            setTimeout(() => {
                const el = document.getElementById("autoStatus");
                if (el && el.innerText === "Guardado ✓") el.innerText = "";
            }, 1500);
        } catch (err) {
            console.error("Error al guardar (sin actualizar):", err);
            this.setAutoStatus("Error al guardar");
        }
    },

    /**
     * Botón "Refrescar todos los informes" del taskpane: recorre TODOS los
     * informes de ReportStore (no solo el activo) y los actualiza — ver
     * actualizarTodosCore() en commands.js (consultas en paralelo, pintado
     * en Excel en serie). Antes, "Actualizar" (ribbon o autoguardado) solo
     * refrescaba el informe que se estuviera editando en ese momento.
     */
    async refreshAllReports() {
        const btn = document.getElementById("btnRefreshAll");
        if (!window.ReportActions || typeof window.ReportActions.actualizarTodos !== "function") return;
        try {
            if (btn) { btn.disabled = true; btn.innerText = "Refrescando todos…"; }
            this.setAutoStatus("Refrescando todos los informes…");
            await window.ReportActions.actualizarTodos();
            this.setAutoStatus("Todos los informes actualizados ✓");
            setTimeout(() => {
                const el = document.getElementById("autoStatus");
                if (el && el.innerText === "Todos los informes actualizados ✓") el.innerText = "";
            }, 2500);
        } catch (err) {
            console.error("Error al refrescar todos los informes:", err);
            this.setAutoStatus("Error al refrescar todos los informes");
        } finally {
            if (btn) { btn.disabled = false; btn.innerText = "Refrescar todos"; }
        }
    },

    /**
     * Se ejecuta automáticamente en init(), es decir, cada vez que este
     * taskpane se CARGA de nuevo (botón del ribbon "Editar informe" ->
     * ShowTaskpane -> TaskpaneId="Taskpane"). Abrir un taskpane no toca la
     * hoja activa ni la selección de Excel, así que en este punto siguen
     * siendo las que tenía el usuario en el momento de pulsar el botón.
     * Guarda:
     *   EDIT_REPORT!D1 <- nombre de la hoja activa   (p.ej. "CSV_RESULT")
     *   EDIT_REPORT!E1 <- celda seleccionada         (p.ej. "C3")
     *
     * Nota: si el panel YA estaba abierto, Excel solo lo trae al frente
     * (no recarga la página ni vuelve a llamar a init()), así que este
     * paso no se repite en ese caso; para eso sigue estando el botón
     * manual "Editar report" (btnEditReport) dentro del propio panel.
     */
    async captureActiveEditContext() {
        try {
            let activeSheetName = "";
            let selectedCellAddress = "";

            await Excel.run(async (context) => {
                const activeSheet = context.workbook.worksheets.getActiveWorksheet();
                activeSheet.load("name");

                const selectedRange = context.workbook.getSelectedRange();
                selectedRange.load("address");

                await context.sync();

                activeSheetName = activeSheet.name;

                // selectedRange.address llega como "Hoja1!C3" (o
                // "Hoja1!C3:D5" si hay un rango de varias celdas
                // seleccionado); nos quedamos solo con la referencia de
                // celda ("C3"), la esquina superior izquierda si es rango.
                const rawAddress = selectedRange.address.split("!").pop();
                selectedCellAddress = rawAddress.split(":")[0];

                const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
                editReportSheet.getRange("D1").values = [[activeSheetName]];
                editReportSheet.getRange("E1").values = [[selectedCellAddress]];

                await context.sync();
            });

            console.log(`[Draco] Editar informe: EDIT_REPORT!D1="${activeSheetName}", EDIT_REPORT!E1="${selectedCellAddress}"`);
            return { activeSheetName, selectedCellAddress };
        } catch (err) {
            console.error("Error guardando la hoja/celda activa en EDIT_REPORT (D1/E1):", err);
            return null;
        }
    },

    /**
     * Botón "Editar report" del TASKPANE (no del ribbon: no depende del
     * manifest). Flujo: el usuario deja activa en Excel la pestaña/hoja
     * que quiere editar y pulsa este botón aquí; se guarda el nombre de
     * esa hoja en EDIT_REPORT!D1 (de donde lo leen commands.js
     * getDracoResultSheetName/resultSheetNameFromGrid en lugar del
     * literal fijo "CSV_RESULT") y se recarga el diseño existente para
     * esa hoja.
     */
    async editReportFromTaskpane() {
        const btn = document.getElementById("btnEditReport");
        try {
            if (btn) btn.disabled = true;
            this.setAutoStatus("Editando report…");

            let activeSheetName = "";
            await Excel.run(async (context) => {
                const activeSheet = context.workbook.worksheets.getActiveWorksheet();
                activeSheet.load("name");
                await context.sync();
                activeSheetName = activeSheet.name;

                const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
                editReportSheet.getRange("D1").values = [[activeSheetName]];
                await context.sync();
            });

            // Se guarda también en el propio informe (no solo en la celda
            // física compartida EDIT_REPORT!D1): así, si luego se edita OTRO
            // informe y cambia D1, este informe no "pierde" su hoja.
            if (this.currentReportId && window.ReportStore && window.ReportStore.setReportResultSheetName) {
                await window.ReportStore.setReportResultSheetName(this.currentReportId, activeSheetName);
            }

            await this.loadDesignFromSheet();
            this.setAutoStatus(`Editando "${activeSheetName}"`);
        } catch (err) {
            console.error("Error en 'Editar report':", err);
            this.setAutoStatus("Error al editar report");
            alert("No se pudo guardar la pestaña activa en EDIT_REPORT!D1: " + (err.message || err));
        } finally {
            if (btn) btn.disabled = false;
        }
    },

    // Evita reentradas mientras el propio cambio de informe (disparado
    // desde la selección de Excel) está en curso: onReportSelectorChange
    // no toca la selección de Excel, pero por si el futuro trajera algún
    // cambio que sí lo hiciera, este flag corta cualquier reentrada.
    _switchingReportFromSelection: false,
    _selectionSyncRegistered: false,

    /**
     * "Seguir el informe activo según la celda seleccionada": engancha
     * UNA sola vez (para todo el libro, no hoja por hoja) el evento
     * workbook.onSelectionChanged. Cada vez que el usuario selecciona una
     * celda (en cualquier pestaña, incluido un cambio de pestaña), se
     * mira si esa hoja es la "hoja de resultados" de algún informe
     * guardado (ReportActions.reportIdForResultSheet, misma función que ya
     * usa el picker de expandir/contraer de EDIT_REPORT!T1): si es de un
     * informe DISTINTO al que el taskpane tiene abierto ahora mismo, el
     * taskpane cambia solo a ese informe -exactamente igual que si el
     * usuario lo hubiera elegido a mano en el desplegable "Informe" (ver
     * onReportSelectorChange)-. Si la hoja no pertenece a ningún informe
     * conocido (p.ej. EDIT_REPORT o una hoja de datos cualquiera), o ya es
     * el informe actual, no se hace nada.
     */
    async registerActiveReportFollowsSelection() {
        if (this._selectionSyncRegistered) return;
        this._selectionSyncRegistered = true;

        await Excel.run(async (context) => {
            context.workbook.onSelectionChanged.add(
                (event) => this.handleSelectionChangedForReportSync(event)
            );
            await context.sync();
        });
    },

    async handleSelectionChangedForReportSync(event) {
        if (this._switchingReportFromSelection) return;
        if (!window.ReportStore || !window.ReportActions || !window.ReportActions.reportIdForResultSheet) return;

        try {
            let sheetName = "";
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(event.worksheetId);
                sheet.load("name");
                await context.sync();
                sheetName = sheet.name;
            });
            if (!sheetName) return;

            const reportId = window.ReportActions.reportIdForResultSheet(sheetName);
            if (!reportId) return;
            if (this.currentReportId && Number(reportId) === Number(this.currentReportId)) return;
            if (!window.ReportStore.getReport(reportId)) return; // por si acaso ya no existe

            this._switchingReportFromSelection = true;
            try {
                const select = document.getElementById("reportSelector");
                if (select) select.value = String(reportId);
                await this.onReportSelectorChange(String(reportId));
                this.setAutoStatus(`Informe: ${this.reportProperties.reportName || ("Informe " + reportId)} (hoja "${sheetName}")`);
            } finally {
                this._switchingReportFromSelection = false;
            }
        } catch (err) {
            console.warn("[Draco] No se pudo sincronizar el informe activo con la selección de Excel:", err);
        }
    },

    async init() {
        // Cada paso va en su propio try/catch: un fallo puntual (p.ej. al
        // leer un ajuste de Office, o al procesar una acción pendiente del
        // ribbon) NUNCA debe impedir que loadFields() rellene el árbol de
        // campos disponibles ni que se cargue el diseño existente.
        const safeStep = async (label, fn) => {
            try {
                await fn();
            } catch (err) {
                console.error(`Error en init() -> ${label}:`, err);
            }
        };

        await safeStep("applyPlanningGateUI", () => this.applyPlanningGateUI());

        if (typeof FilterModal !== "undefined" && FilterModal.init) {
            await safeStep("FilterModal.init", () => FilterModal.init());
        }
        await safeStep("bindEvents", () => this.bindEvents());
        await safeStep("initRefreshPauseButton", () => this.initRefreshPauseButton());
        await safeStep("registerPendingRibbonActionListener", () => this.registerPendingRibbonActionListener());

        // Botón del ribbon "Editar informe" (EditarInformeButton -> Show
        // Taskpane -> TaskpaneId="Taskpane"): abrir el taskpane NO cambia
        // la hoja activa ni la selección de Excel, así que aquí, nada más
        // cargar la página, siguen siendo las que tenía el usuario al
        // pulsar el botón. Las guardamos en EDIT_REPORT!D1/E1 antes de
        // cargar el diseño.
        await safeStep("captureActiveEditContext", () => this.captureActiveEditContext());

        // Selector de informes ("Informe") + selector de modelo semántico
        // (justo encima del buscador de campo): determina currentReportId.
        await safeStep("initReportAndModelSelectors", () => this.initReportAndModelSelectors());

        // Propiedades del informe (modal "Propiedades del informe"): hay
        // que leerlas DESPUÉS de initReportAndModelSelectors(), que es
        // quien fija this.currentReportId (antes se llamaba aquí mismo
        // pero justo antes de ese paso, así que su condición
        // "if (this.currentReportId && ...)" siempre fallaba y nunca
        // llegaba a leer nada de Office roaming settings: se quedaba todo
        // el rato con los valores por defecto del código, como si nunca
        // se hubieran guardado).
        await safeStep("loadReportPropertiesFromSettings", () => this.loadReportPropertiesFromSettings());

        // loadFields() es lo más importante de esta pantalla (el árbol de
        // campos de la izquierda), pero solo tiene sentido si ya hay un
        // informe activo (y un modelo semántico elegido): si todavía no se
        // ha creado ningún informe, el taskpane se queda vacío hasta que
        // el usuario pulse "Añadir informe".
        if (this.currentReportId) {
            await safeStep("loadFields", () => this.loadFields());
            await safeStep("loadDesignFromSheet", () => this.loadDesignFromSheet());
        } else {
            this.showEmptyState();
        }
        await safeStep("handlePendingRibbonAction", () => this.handlePendingRibbonAction());

        // Registro temprano de los listeners de reconocimiento de miembros
        // / indicadores +/-, para que funcionen sin necesidad de que el
        // usuario abra el panel ni pulse "Actualizar informe" antes.
        await safeStep("ensureDracoHandlersRegistered", async () => {
            if (window.ReportActions && window.ReportActions.ensureDracoHandlersRegistered) {
                await window.ReportActions.ensureDracoHandlersRegistered();
            }
        });

        // Igual que arriba, pero para los rangos de "Añadir filtro": se
        // enganchan en TODAS las hojas del libro (pueden vivir en
        // cualquiera, no solo en la de resultados de un informe).
        await safeStep("ensureDracoFilterRangeHandlersRegistered", async () => {
            if (window.ReportActions && window.ReportActions.ensureDracoFilterRangeHandlersRegistered) {
                await window.ReportActions.ensureDracoFilterRangeHandlersRegistered();
            }
        });
        // Nota: el listener de clic en A50 -> A51 NO se registra aquí.
        // Se registra en commands.js (Office.onReady), que se carga tanto
        // en este taskpane como en el runtime de comandos del ribbon, así
        // que se activa con lo que ocurra antes: abrir el panel o pulsar
        // cualquier botón del ribbon (p.ej. "Actualizar").

        // "Seguir el informe activo según la celda seleccionada" (ver la
        // función de arriba): solo tiene sentido dentro del propio
        // taskpane (necesita su desplegable "Informe" y su estado en
        // memoria), así que se registra aquí y no en commands.js/
        // Office.onReady.
        await safeStep("registerActiveReportFollowsSelection", () => this.registerActiveReportFollowsSelection());
    },

    /**
     * Los botones del ribbon "Propiedades" y "Opciones de campo" (ver
     * openReportProperties/openFieldOptions en commands.js) guardan la
     * acción pendiente en Office roaming settings y llaman a
     * Office.addin.showAsTaskpane(). Si el taskpane YA estaba abierto,
     * showAsTaskpane() no recarga la página, así que init() (y por tanto
     * handlePendingRibbonAction) no se vuelve a ejecutar solo por eso: sin
     * este listener de SettingsChanged los paneles/modal existen en el
     * HTML pero nunca llegan a mostrarse en ese caso, que es el habitual.
     */
    registerPendingRibbonActionListener() {
        try {
            const settings = Office.context && Office.context.document && Office.context.document.settings;
            if (settings && settings.addHandlerAsync) {
                settings.addHandlerAsync(Office.EventType.SettingsChanged, () => {
                    this.handlePendingRibbonAction();
                    this.setRefreshPausedUI(this.isRefreshPaused());
                });
            }
        } catch (err) {
            console.warn("No se pudo registrar el listener de SettingsChanged:", err);
        }
    },

    bindEvents() {
        const btnRefresh = document.getElementById("btnRefresh");
        if (btnRefresh) btnRefresh.addEventListener("click", () => this.loadFields());

        const fieldSearch = document.getElementById("fieldSearch");
        if (fieldSearch) fieldSearch.addEventListener("input", (e) => this.filterFields(e.target.value));

        const btnAddReport = document.getElementById("btnAddReport");
        if (btnAddReport) btnAddReport.addEventListener("click", () => this.addReport());

        const reportSelector = document.getElementById("reportSelector");
        if (reportSelector) reportSelector.addEventListener("change", (e) => this.onReportSelectorChange(e.target.value));

        const modelSelector = document.getElementById("semanticModelSelector");
        if (modelSelector) modelSelector.addEventListener("change", (e) => this.onModelSelectorChange(e.target.value));

        const btnEditReport = document.getElementById("btnEditReport");
        if (btnEditReport) btnEditReport.addEventListener("click", () => this.editReportFromTaskpane());

        const btnRefreshAll = document.getElementById("btnRefreshAll");
        if (btnRefreshAll) btnRefreshAll.addEventListener("click", () => this.refreshAllReports());

        // Checkboxes Estático / Dinámico (Checkrow / CheckCol del VBA): NO
        // disparan Actualizar() (solo guardan el flag), sombrean la zona y
        // deshabilitan el drag&drop mientras esté marcado. Además, convierten
        // AHORA MISMO (sin esperar al próximo refresco real) las celdas ya
        // pintadas de Draco_001_Rows/Draco_001_Cols entre texto plano y
        // fórmula EPM_VALUE (ver convertAxisStaticFormulas en commands.js).
        const chkRows = document.getElementById("chkAsymmetricRows");
        const chkCols = document.getElementById("chkAsymmetricCols");

        if (chkRows) {
            chkRows.addEventListener("change", async (e) => {
                this.state.rowsStatic = e.target.checked;
                this.updateStaticLabel("rows");
                this.setZoneLocked("rows", this.state.rowsStatic);
                await this.applyAxisStaticFormulas("rows", this.state.rowsStatic);
                this.saveDesignOnly();
            });
        }
        if (chkCols) {
            chkCols.addEventListener("change", async (e) => {
                this.state.colsStatic = e.target.checked;
                this.updateStaticLabel("cols");
                this.setZoneLocked("columns", this.state.colsStatic);
                await this.applyAxisStaticFormulas("columns", this.state.colsStatic);
                this.saveDesignOnly();
            });
        }

        // Flechas de movimiento de rango
        this.bindArrow("btnRRUp", () => { RangeAxis.moveRR(-1, 0); this.refreshRangeLabels(); this.scheduleAutoUpdate(); });
        this.bindArrow("btnRRDown", () => { RangeAxis.moveRR(1, 0); this.refreshRangeLabels(); this.scheduleAutoUpdate(); });
        this.bindArrow("btnRRLeft", () => { RangeAxis.moveRR(0, -1); this.refreshRangeLabels(); this.scheduleAutoUpdate(); });
        this.bindArrow("btnRRRight", () => { RangeAxis.moveRRRight(); this.refreshRangeLabels(); this.scheduleAutoUpdate(); });

        this.bindArrow("btnRCUp", () => { RangeAxis.moveRC(-1, 0); this.refreshRangeLabels(); this.scheduleAutoUpdate(); });
        this.bindArrow("btnRCDown", () => { RangeAxis.moveRCDown(); this.refreshRangeLabels(); this.scheduleAutoUpdate(); });
        this.bindArrow("btnRCLeft", () => { RangeAxis.moveRC(0, -1); this.refreshRangeLabels(); this.scheduleAutoUpdate(); });
        this.bindArrow("btnRCRight", () => { RangeAxis.moveRC(0, 1); this.refreshRangeLabels(); this.scheduleAutoUpdate(); });

        // Dropzones
        const zones = document.querySelectorAll(".zone-card");
        zones.forEach(zone => {
            zone.addEventListener("dragenter", (e) => e.preventDefault());
            zone.addEventListener("dragover", (e) => this.handleDragOver(e));
            zone.addEventListener("dragleave", (e) => this.handleDragLeave(e));
            zone.addEventListener("drop", (e) => this.handleDrop(e));
        });

        // Botón "Propiedades del informe"
        const btnProps = document.getElementById("btnReportProperties");
        if (btnProps) btnProps.addEventListener("click", () => this.openReportPropertiesModal());

        const btnCancelProps = document.getElementById("btnCancelProperties");
        if (btnCancelProps) btnCancelProps.addEventListener("click", () => this.closeReportPropertiesModal());

        const btnCloseProps = document.getElementById("closePropertiesModalBtn");
        if (btnCloseProps) btnCloseProps.addEventListener("click", () => this.closeReportPropertiesModal());

        const btnSaveProps = document.getElementById("btnSaveProperties");
        if (btnSaveProps) btnSaveProps.addEventListener("click", () => this.saveReportPropertiesFromModal());

        // "Comportamiento de dimensiones" (solo visible/aplicable para
        // informes de planificación): se muestra/oculta y se rellena al
        // vuelo según el tick y el modelo semántico elegidos, sin tener
        // que reabrir el modal.
        const propPlanningCheckbox = document.getElementById("propPlanningReport");
        if (propPlanningCheckbox) propPlanningCheckbox.addEventListener("change", () => this.toggleDimensionsBehaviorSection());

        const propModelSelect = document.getElementById("propSemanticModel");
        if (propModelSelect) propModelSelect.addEventListener("change", () => {
            this.applyPlanningGateToProperties();
            if (document.getElementById("propPlanningReport").checked) this.renderDimensionsBehaviorRows();
        });

        // Botón "Opciones de campo" (toggle del panel derecho)
        const btnFieldOptions = document.getElementById("btnFieldOptions");
        if (btnFieldOptions) btnFieldOptions.addEventListener("click", () => this.toggleFieldOptionsPanel());

        const btnDistribuirValores = document.getElementById("btnDistribuirValores");
        if (btnDistribuirValores) btnDistribuirValores.addEventListener("click", () => this.openDistributeValuesModal());
        if (window.DistributeValuesPanel) window.DistributeValuesPanel.initEvents();

        // Botón "Guardar en bucket" (sube el .xlsx activo a Cloud Storage)
        const btnSaveToBucket = document.getElementById("btnSaveToBucket");
        if (btnSaveToBucket) btnSaveToBucket.addEventListener("click", () => this.saveWorkbookToBucket());

        // Botón "Añadir filtro" DE LA BARRA DEL TASKPANE (btnAddFilterRange,
        // junto a "Añadir informe"): al estar en esta misma página, abre
        // directamente el modal embebido (sin diálogo aparte). Distinto del
        // botón del RIBBON (AnadirFiltroButton en manifest.xml), que corre
        // en el runtime aislado de comandos y por eso usa un diálogo
        // standalone (addFilterRange.html) — ver abrirAnadirFiltro en
        // commands.js.
        const btnAddFilterRange = document.getElementById("btnAddFilterRange");
        if (btnAddFilterRange) btnAddFilterRange.addEventListener("click", () => this.openAddFilterRangeModal());
        const btnCancelAddFilterRange = document.getElementById("btnCancelAddFilterRange");
        if (btnCancelAddFilterRange) btnCancelAddFilterRange.addEventListener("click", () => this.closeAddFilterRangeModal());

        const btnCloseAddFilterRange = document.getElementById("closeAddFilterRangeModalBtn");
        if (btnCloseAddFilterRange) btnCloseAddFilterRange.addEventListener("click", () => this.closeAddFilterRangeModal());

        const btnConfirmAddFilterRange = document.getElementById("btnConfirmAddFilterRange");
        if (btnConfirmAddFilterRange) btnConfirmAddFilterRange.addEventListener("click", () => this.createFilterRangeFromModal());

        // Modal de confirmación genérico (ver showConfirmDialog): los tres
        // botones solo resuelven la promesa pendiente, guardada en
        // this._confirmDialogResolve por cada llamada.
        const btnAcceptConfirmDialog = document.getElementById("btnAcceptConfirmDialog");
        if (btnAcceptConfirmDialog) btnAcceptConfirmDialog.addEventListener("click", () => this.resolveConfirmDialog(true));
        const btnCancelConfirmDialog = document.getElementById("btnCancelConfirmDialog");
        if (btnCancelConfirmDialog) btnCancelConfirmDialog.addEventListener("click", () => this.resolveConfirmDialog(false));
        const btnCloseConfirmDialog = document.getElementById("closeConfirmDialogBtn");
        if (btnCloseConfirmDialog) btnCloseConfirmDialog.addEventListener("click", () => this.resolveConfirmDialog(false));
    },

    /* -------------------------------------------------------------
     * Sustituto de window.confirm(): Excel bloquea/ignora los diálogos
     * nativos del navegador dentro del WebView del taskpane, así que un
     * confirm() ahí nunca llega a resolverse (por eso la "×" de un filtro
     * global no borraba nada). Este modal hace lo mismo con el mismo
     * patrón que el resto de modales de la app (modal-backdrop/dialog).
     * Devuelve una Promise<boolean>: true si el usuario pulsa "Aceptar",
     * false si cancela, cierra con la "×" o pulsa fuera del diálogo.
     * ----------------------------------------------------------- */
    showConfirmDialog(message, options) {
        const opts = options || {};
        const modal = document.getElementById("confirmDialogModal");
        if (!modal) return Promise.resolve(false); // por si el HTML no está (defensivo)

        const titleEl = document.getElementById("confirmDialogTitle");
        const messageEl = document.getElementById("confirmDialogMessage");
        const acceptBtn = document.getElementById("btnAcceptConfirmDialog");
        if (titleEl) titleEl.textContent = opts.title || "Confirmar";
        if (messageEl) messageEl.textContent = message || "";
        if (acceptBtn) acceptBtn.textContent = opts.confirmLabel || "Aceptar";

        // Si ya había una confirmación pendiente (no debería, pero por si
        // acaso), se cancela antes de abrir la nueva para no dejarla colgada.
        if (this._confirmDialogResolve) this._confirmDialogResolve(false);

        modal.style.display = "flex";
        return new Promise((resolve) => {
            this._confirmDialogResolve = resolve;
        });
    },

    resolveConfirmDialog(result) {
        const modal = document.getElementById("confirmDialogModal");
        if (modal) modal.style.display = "none";
        const resolve = this._confirmDialogResolve;
        this._confirmDialogResolve = null;
        if (resolve) resolve(result);
    },

    /* -------------------------------------------------------------
     * "Añadir filtro": crea, sobre la celda seleccionada AHORA MISMO en
     * Excel, un rango con nombre Draco_Filter_<DIM>_<CAMPO>_<sufijo> que
     * más adelante (clic sobre esa celda) abre el selector de valores y
     * pinta el resultado con las exclusiones en granate. Ver
     * dracoFilterRangeName/ensureDracoFilterRangeHandlersRegistered y el
     * resto de "Añadir filtro" en commands.js, y js/filterRangeStore.js
     * para los metadatos guardados.
     *
     * Se abre desde btnAddFilterRange (barra del taskpane, ver bindEvents).
     * El botón del RIBBON (AnadirFiltroButton) es un flujo aparte: usa el
     * diálogo standalone addFilterRange.html + abrirAnadirFiltro en
     * commands.js, porque ese runtime no tiene acceso a este modal.
     * ----------------------------------------------------------- */
    async openAddFilterRangeModal() {
        const modal = document.getElementById("addFilterRangeModal");
        if (!modal) return;

        const dimSelect = document.getElementById("selAddFilterDimension");
        const reportSelect = document.getElementById("selAddFilterReport");
        if (dimSelect) dimSelect.innerHTML = "<option value=\"\">Cargando…</option>";
        if (reportSelect) reportSelect.innerHTML = "";

        modal.style.display = "flex";

        // Dimensiones/campos: mismo origen de datos que el árbol de campos
        // de la izquierda (readDim2Data, sobre el modelo semántico activo).
        try {
            const result = await window.ExcelService.readDim2Data();
            const dimensions = (result && result.data) || [];
            if (dimSelect) {
                dimSelect.innerHTML = "";
                dimensions.forEach(dim => {
                    const group = document.createElement("optgroup");
                    group.label = dim.dimension;

                    dim.attributes.forEach(attr => {
                        const opt = document.createElement("option");
                        opt.value = JSON.stringify({ dim: dim.dimension, name: attr, isHierarchy: false });
                        opt.textContent = attr;
                        group.appendChild(opt);
                    });
                    dim.hierarchies.forEach(hier => {
                        const opt = document.createElement("option");
                        opt.value = JSON.stringify({ dim: dim.dimension, name: hier, isHierarchy: true });
                        opt.textContent = `${hier} (jerarquía)`;
                        group.appendChild(opt);
                    });

                    if (group.children.length) dimSelect.appendChild(group);
                });
                if (!dimSelect.options.length) {
                    dimSelect.innerHTML = "<option value=\"\">No hay campos disponibles</option>";
                }
            }
        } catch (err) {
            console.error("Error al cargar dimensiones para 'Añadir filtro':", err);
            if (dimSelect) dimSelect.innerHTML = "<option value=\"\">Error al cargar</option>";
        }

        // Informes: los ya creados en este libro (ReportStore) + "Todos".
        if (reportSelect) {
            const reports = (window.ReportStore ? window.ReportStore.listReports() : []) || [];
            reports.forEach(r => {
                const opt = document.createElement("option");
                opt.value = String(r.id);
                opt.textContent = r.name || `Informe ${r.id}`;
                if (this.currentReportId && Number(this.currentReportId) === Number(r.id)) opt.selected = true;
                reportSelect.appendChild(opt);
            });
            const optAll = document.createElement("option");
            optAll.value = "all";
            optAll.textContent = "Todos los informes";
            reportSelect.appendChild(optAll);
        }
    },

    closeAddFilterRangeModal() {
        const modal = document.getElementById("addFilterRangeModal");
        if (modal) modal.style.display = "none";
    },

    async createFilterRangeFromModal() {
        const dimSelect = document.getElementById("selAddFilterDimension");
        const reportSelect = document.getElementById("selAddFilterReport");
        const btn = document.getElementById("btnConfirmAddFilterRange");

        if (!dimSelect || !dimSelect.value) {
            alert("Elige una dimensión o jerarquía para el filtro.");
            return;
        }
        if (typeof dracoFilterRangeName !== "function") {
            alert("No se encuentra la lógica de 'Añadir filtro' (commands.js). Recarga el panel e inténtalo de nuevo.");
            return;
        }

        let fieldData;
        try {
            fieldData = JSON.parse(dimSelect.value);
        } catch (e) {
            alert("No se ha podido leer el campo seleccionado.");
            return;
        }

        const reportValue = reportSelect ? reportSelect.value : "all";
        const isAll = !reportValue || reportValue === "all";
        const reportId = isAll ? null : Number(reportValue);
        const suffix = isAll ? "all" : pad3(reportId);

        const rangeName = dracoFilterRangeName(fieldData.dim, fieldData.name, suffix);

        if (btn) { btn.disabled = true; btn.innerText = "Añadiendo…"; }

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.load("name");

                const selectedRange = context.workbook.getSelectedRange();
                selectedRange.load("address");
                await context.sync();

                // Esquina superior izquierda si hay varias celdas seleccionadas.
                const rawAddress = selectedRange.address.split("!").pop();
                const addr = rawAddress.split(":")[0];
                const cell = sheet.getRange(addr);

                // Si ya existía un rango con este mismo nombre (misma
                // dimensión/campo + mismo informe), se borra y se vuelve a
                // crear sobre la NUEVA celda seleccionada (así "Añadir
                // filtro" también sirve para mover un filtro existente).
                const existing = context.workbook.names.getItemOrNullObject(rangeName);
                existing.load("isNullObject");
                await context.sync();
                if (!existing.isNullObject) {
                    existing.delete();
                    await context.sync();
                }

                context.workbook.names.add(rangeName, cell);
                await context.sync();
            });

            await window.FilterRangeStore.set(rangeName, {
                dim: fieldData.dim,
                name: fieldData.name,
                isHierarchy: !!fieldData.isHierarchy,
                reportId: reportId,
                filter: null
            });

            this.closeAddFilterRangeModal();
            this.setAutoStatus("Filtro añadido ✓");
            await this.loadDesignFromSheet();

            // Por si el usuario ha creado el filtro en una hoja nueva que
            // todavía no tuviera enganchado el listener de clic.
            if (window.ReportActions && window.ReportActions.ensureDracoFilterRangeHandlersRegistered) {
                await window.ReportActions.ensureDracoFilterRangeHandlersRegistered();
            }
        } catch (err) {
            console.error("Error al añadir el rango de filtro:", err);
            alert("Error al añadir el filtro: " + (err.message || err));
        } finally {
            if (btn) { btn.disabled = false; btn.innerText = "Añadir"; }
        }
    },

    /**
     * Abre el selector de proyecto/bucket (saveBucket.html, ver
     * js/gcsSaveBridge.js) y, con esa elección, sube una copia del Excel
     * activo (tal cual está guardado) al bucket elegido.
     * Ver js/gcsExport.js (GCS.saveActiveWorkbookToBucketNamed).
     */
    saveWorkbookToBucket() {
        if (!window.GCS || !window.GcsSaveBridge) {
            alert("No se encontró el módulo de exportación a Cloud Storage (js/gcsExport.js o js/gcsSaveBridge.js).");
            return;
        }

        const btn = document.getElementById("btnSaveToBucket");

        window.GcsSaveBridge.openSaveBucketDialog(async ({ bucket, objectName }) => {
            if (btn) btn.disabled = true;
            this.setAutoStatus("Subiendo a Cloud Storage…");

            try {
                const result = await window.GCS.saveActiveWorkbookToBucketNamed(bucket, objectName);
                this.setAutoStatus("Guardado en el bucket ✓");
                alert(`Archivo subido correctamente a gs://${result.bucket}/${result.name}`);
                setTimeout(() => {
                    const el = document.getElementById("autoStatus");
                    if (el && el.innerText === "Guardado en el bucket ✓") el.innerText = "";
                }, 2500);
            } catch (err) {
                console.error("Error al subir el Excel a Cloud Storage:", err);
                this.setAutoStatus("Error al subir");
                alert("Error al subir el archivo a Cloud Storage: " + (err.message || err));
            } finally {
                if (btn) btn.disabled = false;
            }
        });
    },

    /* -------------------------------------------------------------
     * Eje Estático: sombrea la zona y desactiva el drag&drop (tanto
     * soltar campos nuevos como arrastrar/quitar los que ya hay).
     * ----------------------------------------------------------- */
    setZoneLocked(zoneId, locked) {
        const zone = document.querySelector(`.zone-card[data-zone="${zoneId}"]`);
        if (!zone) return;
        zone.classList.toggle("zone-locked", locked);
        zone.querySelectorAll(".dropped-tag, .measure-group").forEach(tag => { tag.draggable = !locked; });
    },

    isZoneLocked(zoneId) {
        const zone = document.querySelector(`.zone-card[data-zone="${zoneId}"]`);
        return !!(zone && zone.classList.contains("zone-locked"));
    },

    /**
     * "Pulsador" del ribbon (Propiedades / Opciones de campo): Office.js no
     * soporta un botón de ribbon con estado visual pulsado/despulsado
     * nativo (eso solo existe en el ribbon COM/VBA clásico). El acercamiento
     * soportado en add-ins web es cambiar la etiqueta del botón para reflejar
     * el estado, igual que ya se hace con "Pausar refresco" — reutiliza
     * requestRibbonLabelUpdate (definida en commands.js; mismo contexto JS
     * gracias al Shared Runtime).
     */
    updateRibbonToggleLabel(controlId, baseLabel, isOpen) {
        try {
            if (typeof requestRibbonLabelUpdate === "function") {
                requestRibbonLabelUpdate(controlId, isOpen ? baseLabel + " ✓" : baseLabel);
            }
        } catch (err) {
            console.warn(`No se pudo actualizar el estado visual del botón ${controlId}:`, err);
        }
    },

    /**
     * Reescribe in situ (Excel.run) las celdas ya pintadas del rango con
     * nombre Draco_001_Rows/Draco_001_Cols del eje indicado, alternando
     * entre texto plano y fórmula EPM_VALUE. Si todavía no existe tabla
     * pintada no hace nada (se pintará ya en el modo correcto en el
     * próximo refresco). No bloquea la UI: los errores solo se registran.
     */
    async applyAxisStaticFormulas(axis, isStatic) {
        try {
            // "Solo texto": con IDs ocultos no se puede convertir celda a celda
            // (el ID no está en la hoja): se guarda el diseño y se repinta.
            if (window.ReportActions && window.ReportActions.repaintAxisIfHiddenIds && this.currentReportId) {
                await window.ExcelService.saveEditReportDesign(this.currentReportId, this.collectDesignPayload());
                if (await window.ReportActions.repaintAxisIfHiddenIds(axis)) return;
            }
            if (window.ReportActions && window.ReportActions.convertAxisStaticFormulas) {
                await window.ReportActions.convertAxisStaticFormulas(axis, isStatic);
            }
        } catch (err) {
            console.error(`Error al convertir las celdas de "${axis}" a EPM_VALUE:`, err);
        }
    },

    bindArrow(id, handler) {
        const el = document.getElementById(id);
        if (el) el.addEventListener("click", handler);
    },

    updateStaticLabel(axis) {
        const labelId = axis === "rows" ? "lblStaticRows" : "lblStaticCols";
        const isStatic = axis === "rows" ? this.state.rowsStatic : this.state.colsStatic;
        const label = document.getElementById(labelId);
        if (!label) return;
        label.innerText = isStatic ? "Estático" : "Dinámico";
        label.style.color = isStatic ? "#1D8154" : "#7045A9";

        // Modo del informe (ver dracoReportMode en commands.js): se muestra
        // como ayuda al pasar el ratón por cualquiera de los dos interruptores.
        const r = this.state.rowsStatic, c = this.state.colsStatic;
        const mode = r && c ? "Estático (las dos líneas las defines tú en la hoja)"
            : (r || c) ? `Híbrido: ${r ? "Filas" : "Columnas"} estático (tus líneas de la hoja) y ${r ? "Columnas" : "Filas"} dinámico (sale de los datos)`
            : "Dinámico (los dos ejes salen de los datos)";
        ["lblStaticRows", "lblStaticCols"].forEach(id => {
            const el = document.getElementById(id);
            if (el && el.parentElement) el.parentElement.title = "Modo del informe: " + mode;
        });
    },

    refreshRangeLabels() {
        const rrLabel = document.getElementById("lblRRango");
        const rcLabel = document.getElementById("lblCRango");
        if (rrLabel) rrLabel.innerText = RangeAxis.addressOf("rr") || "(sin definir)";
        if (rcLabel) rcLabel.innerText = RangeAxis.addressOf("rc") || "(sin definir)";
    },

    /* -------------------------------------------------------------
     * Carga del diseño existente desde EDIT_REPORT
     * ----------------------------------------------------------- */
    async loadDesignFromSheet() {
        if (!this.currentReportId) return;
        try {
            const design = await window.ExcelService.loadEditReportDesign(this.currentReportId);

            this.state.filters = design.filters.map(f => ({
                dimension: f.dimension,
                name: f.name,
                isHierarchy: f.isHierarchy,
                realAttribute: f.realAttribute,
                // f.value es la cadena tal cual se guardó (JSON del filtro
                // nuevo, o un valor simple si el informe es de antes de
                // tener selección múltiple/rango/incluir-excluir).
                filter: (typeof window.parseFilterValue === "function") ? window.parseFilterValue(f.value) : null
            }));
            this.state.rows = design.rows.map(r => ({ dimension: r.dimension, name: r.name, isHierarchy: r.isHierarchy }));
            this.state.columns = design.columns.map(c => ({ dimension: c.dimension, name: c.name, isHierarchy: c.isHierarchy }));
            this.state.rowsStatic = design.rowsStatic;
            this.state.colsStatic = design.colsStatic;
            this.state.fieldOptions = design.fieldOptions || {};
            this.refreshCalcBadges();

            RangeAxis.loadFromAddresses(design.rrAddress, design.rcAddress, design.anchorAddress);
            // El rango guardado ya incluye las columnas de "ID y texto" del diseño.
            this.reconcileTextColumns(true);

            // Pintar checkboxes + sombreado/bloqueo si el eje ya venía Estático
            const chkRows = document.getElementById("chkAsymmetricRows");
            const chkCols = document.getElementById("chkAsymmetricCols");
            if (chkRows) chkRows.checked = this.state.rowsStatic;
            if (chkCols) chkCols.checked = this.state.colsStatic;
            this.updateStaticLabel("rows");
            this.updateStaticLabel("cols");
            this.setZoneLocked("rows", this.state.rowsStatic);
            this.setZoneLocked("columns", this.state.colsStatic);

            this.refreshRangeLabels();

            // Pintar tags existentes en cada zona
            const filtersContent = document.querySelector('.zone-card[data-zone="filters"] .dropzone-content');
            const rowsContent = document.querySelector('.zone-card[data-zone="rows"] .dropzone-content');
            const colsContent = document.querySelector('.zone-card[data-zone="columns"] .dropzone-content');

            if (filtersContent) filtersContent.innerHTML = "";
            if (rowsContent) rowsContent.innerHTML = "";
            if (colsContent) colsContent.innerHTML = "";

            this.state.filters.forEach(f => this.renderTag(filtersContent, "filters", f));
            this.renderLockedFilterTags(filtersContent);
            this.state.rows.forEach(r => this.renderTag(rowsContent, "rows", r));
            this.state.columns.forEach(c => this.renderTag(colsContent, "columns", c));

            // El bloqueo de arrastre se aplica de nuevo aquí porque renderTag
            // crea las etiquetas con draggable=true por defecto.
            this.setZoneLocked("rows", this.state.rowsStatic);
            this.setZoneLocked("columns", this.state.colsStatic);

        } catch (err) {
            console.error("Error cargando el diseño desde EDIT_REPORT:", err);
        }
    },

    /* -------------------------------------------------------------
     * Selector de informes ("Informe") + selector de modelo semántico
     * (encima del buscador de campo) + botón "Añadir informe".
     * ----------------------------------------------------------- */

    async initReportAndModelSelectors() {
        await this.populateReportSelector();
        await this.populateModelSelectorMain();
    },

    // Vacía el árbol de campos y las tres zonas: se usa mientras no hay
    // ningún informe creado todavía (justo después de abrir el taskpane
    // por primera vez, antes de pulsar "Añadir informe").
    showEmptyState() {
        const container = document.getElementById("availableFieldsContainer");
        if (container) {
            container.innerHTML = "<div style='color:#605e5c; padding:4px;'>Pulsa \"Añadir informe\" para empezar.</div>";
        }
        ["filters", "rows", "columns"].forEach(zoneId => {
            const content = document.querySelector(`.zone-card[data-zone="${zoneId}"] .dropzone-content`);
            if (content) content.innerHTML = "";
        });
    },

    async populateReportSelector() {
        const select = document.getElementById("reportSelector");
        if (!select || !window.ReportStore) return;

        const reports = window.ReportStore.listReports();
        this.currentReportId = window.ReportStore.getActiveReportId();

        select.innerHTML = "";
        if (reports.length === 0) {
            select.innerHTML = "<option value=\"\">— Sin informes —</option>";
            select.value = "";
            return;
        }
        reports.forEach(r => {
            const opt = document.createElement("option");
            opt.value = String(r.id);
            opt.textContent = r.name;
            select.appendChild(opt);
        });
        select.value = this.currentReportId ? String(this.currentReportId) : "";
    },

    // Selector de modelo semántico del editor de informes (justo encima
    // del buscador de campo). Si el informe actual todavía no tiene
    // modelo asignado y solo existe uno guardado, se rellena solo.
    async populateModelSelectorMain() {
        const select = document.getElementById("semanticModelSelector");
        if (!select) return;

        const models = (window.ExcelService && window.ExcelService.getSemanticModels)
            ? await window.ExcelService.getSemanticModels() : [];

        select.innerHTML = "";
        if (models.length === 0) {
            select.innerHTML = "<option value=\"\">— Sin modelos semánticos —</option>";
            select.disabled = true;
            return;
        }
        select.disabled = false;
        select.innerHTML = "<option value=\"\">— Selecciona un modelo semántico —</option>";
        models.forEach(name => {
            const opt = document.createElement("option");
            opt.value = name;
            opt.textContent = name;
            select.appendChild(opt);
        });

        const report = (this.currentReportId && window.ReportStore) ? window.ReportStore.getReport(this.currentReportId) : null;
        let modelName = report ? (report.semanticModelName || "") : "";

        // Un único modelo semántico disponible: se asigna por defecto.
        if (!modelName && models.length === 1 && report && window.ReportStore) {
            modelName = models[0];
            await window.ReportStore.setSemanticModel(this.currentReportId, modelName);
        }

        select.value = modelName || "";
        if (modelName && window.SemanticModelStore) {
            await window.SemanticModelStore.setActiveModelName(modelName);
        }
    },

    async onReportSelectorChange(reportIdStr) {
        if (!reportIdStr || !window.ReportStore) return;
        this.currentReportId = Number(reportIdStr);
        await window.ReportStore.setActiveReportId(this.currentReportId);

        this.loadReportPropertiesFromSettings();
        await this.populateModelSelectorMain();
        await this.loadFields();
        await this.loadDesignFromSheet();
    },

    async onModelSelectorChange(modelName) {
        if (!this.currentReportId || !window.ReportStore) return;
        await window.ReportStore.setSemanticModel(this.currentReportId, modelName);
        if (modelName && window.SemanticModelStore) {
            await window.SemanticModelStore.setActiveModelName(modelName);
        }
        await this.loadFields();
    },

    /* -------------------------------------------------------------
     * Botón "Añadir informe": crea un informe nuevo y VACÍO (nombre por
     * defecto secuencial "Informe - 0XX"), lo selecciona como activo y
     * deja el taskpane listo para diseñarlo desde cero. Sustituye al
     * antiguo botón "Guardar" (el diseño ya se autoguarda solo, ver
     * scheduleAutoUpdate/runAutoSaveAndRefresh).
     *
     * El informe NO crea una pestaña nueva: se pinta en la hoja y celda
     * donde esté el cursor en el momento de pulsar el botón (ver
     * captureActiveEditContext, que se re-ejecuta aquí para tener la hoja/
     * celda actuales, no las de cuando se abrió el taskpane).
     * ----------------------------------------------------------- */
    async addReport() {
        const btn = document.getElementById("btnAddReport");
        try {
            if (btn) { btn.disabled = true; btn.innerText = "Añadiendo…"; }

            // Re-capturar la hoja/celda activas de Excel EN ESTE MOMENTO
            // (EDIT_REPORT!D1/E1). Sin esto, E1 seguía teniendo la celda
            // capturada al abrir el taskpane (ver captureActiveEditContext,
            // que solo se ejecuta una vez en init()), así que el segundo
            // informe (y siguientes) anclaba en la celda del PRIMER
            // informe en lugar de en la celda seleccionada ahora.
            const editContext = await this.captureActiveEditContext();

            const models = (window.ExcelService && window.ExcelService.getSemanticModels)
                ? await window.ExcelService.getSemanticModels() : [];
            const defaultModel = models.length === 1 ? models[0] : "";

            // La hoja de resultados de este informe es la hoja activa
            // capturada arriba: NO se crea una pestaña nueva "Informe -
            // 0XX" por informe (ver reportStore.createReport).
            const activeSheetName = editContext ? editContext.activeSheetName : "";
            const report = await window.ReportStore.createReport(defaultModel, activeSheetName);
            this.currentReportId = report.id;

            // Estado en memoria del taskpane, vacío para el informe recién creado.
            this.state.filters = [];
            this.state.rows = [];
            this.state.columns = [];
            this.state.rowsStatic = false;
            this.state.colsStatic = false;
            this.state.fieldOptions = {};
            this.refreshCalcBadges();
            this.reportProperties = Object.assign({}, this.reportProperties, {
                reportName: report.name,
                suppressZeroRows: false,
                suppressZeroCols: false,
                subtotalsOnTop: false,
                hierarchyNodesOnTop: true,
                overwriteFormats: true,
                overwriteStaticFormats: false,
                autoFitColumns: false,
                planningReport: false
            });
            RangeAxis.loadFromAddresses("", "", "");

            await this.populateReportSelector();
            await this.populateModelSelectorMain();
            this.showEmptyState();
            await this.loadFields();
            await this.loadDesignFromSheet();
        } catch (err) {
            console.error("Error al añadir el informe:", err);
            alert("Error al añadir el informe: " + (err.message || err));
        } finally {
            if (btn) { btn.disabled = false; btn.innerText = "+ Informe"; }
        }
    },

    /* -------------------------------------------------------------
     * Lista de campos disponibles (izquierda)
     * ----------------------------------------------------------- */
    /**
     * Al volver a «Editar informes» (p. ej. desde el modelo semántico): vuelve
     * a leer los modelos y los campos de dimensiones del informe actual, por
     * si el modelo semántico ha cambiado.
     */
    async refreshSemanticFields() {
        try {
            await this.populateModelSelectorMain();
            await this.loadFields();
        } catch (e) {
            console.warn("[Draco] No se pudieron recargar los campos del modelo:", e);
        }
    },

    async loadFields() {
        const container = document.getElementById("availableFieldsContainer");
        if (!container) return;

        container.innerHTML = "<div style='color: #605e5c; padding: 4px;'>Cargando dimensiones...</div>";

        try {
            // Modelo del INFORME ACTUAL directamente (ver comentario en
            // readDim2Data): no depender del "modelo activo" global.
            const report = (this.currentReportId && window.ReportStore)
                ? window.ReportStore.getReport(this.currentReportId) : null;
            const modelName = report ? (report.semanticModelName || "") : "";

            const result = await window.ExcelService.readDim2Data(modelName || undefined);

            if (result.error) {
                container.innerHTML = `<div style='color: #a80000; padding: 4px;'>⚠️ ${draco_escapeHtml(result.error)}</div>`;
                return;
            }

            const dimensions = result.data || [];
            container.innerHTML = "";

            if (dimensions.length === 0) {
                container.innerHTML = "<div style='color: #605e5c; padding: 4px;'>No se encontraron campos</div>";
                return;
            }

            dimensions.forEach(dim => {
                const group = document.createElement("div");
                group.className = "dimension-group";

                const header = document.createElement("div");
                header.className = "dimension-header";
                // SEGURIDAD: dim.dimension puede venir de un modelo semántico
                // importado desde un repositorio Git compartido (.lkml), no
                // solo del esquema de BigQuery -> se construye con nodos y
                // textContent en vez de innerHTML.
                const caretSpan = document.createElement("span");
                caretSpan.className = "dimension-caret";
                caretSpan.textContent = "\u25be"; // ▾
                const nameSpan = document.createElement("span");
                nameSpan.textContent = dim.dimension.toLowerCase();
                header.appendChild(caretSpan);
                header.appendChild(nameSpan);
                // ƒx: cálculos entre miembros de los atributos de esta
                // dimensión (abre el editor con el primero seleccionado).
                if (dim.attributes && dim.attributes.length) {
                    header.appendChild(this.createCalcButton(dim.dimension, null, dim.attributes));
                }
                // Contraer/expandir la dimensión oculta sus atributos (y
                // jerarquías) para dejar el árbol de campos más compacto,
                // igual que las listas de campos de una tabla dinámica.
                header.addEventListener("click", () => {
                    const collapsed = group.classList.toggle("collapsed");
                    header.querySelector(".dimension-caret").textContent = collapsed ? "▸" : "▾";
                });
                group.appendChild(header);

                dim.hierarchies.forEach(hier => {
                    group.appendChild(this.createFieldElement(dim.dimension, hier, true));
                });

                dim.attributes.forEach(att => {
                    group.appendChild(this.createFieldElement(dim.dimension, att, false, dim.attributes));
                });

                container.appendChild(group);
            });
            this.refreshCalcBadges();
        } catch (err) {
            console.error("Error al cargar dimensiones:", err);
            container.innerHTML = `<div style='color: #a80000; padding: 4px;'>❌ Error: ${draco_escapeHtml(err.message || err)}</div>`;
        }
    },

    /**
     * Botón ƒx (a la derecha del campo / de la dimensión) que abre el
     * editor de cálculos entre miembros (js/calcEditor.js). Solo en
     * atributos: las jerarquías no admiten cálculos.
     */
    createCalcButton(dim, attribute, attributes) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "field-fx-btn";
        btn.textContent = "ƒx";
        btn.title = attribute ? `Cálculos de ${attribute.toLowerCase()}` : `Cálculos de ${String(dim).toLowerCase()}`;
        btn.setAttribute("aria-label", btn.title);
        btn.dataset.dim = String(dim).toUpperCase();
        if (attribute) btn.dataset.attr = String(attribute).toUpperCase();
        btn.draggable = false;
        btn.addEventListener("mousedown", (e) => e.stopPropagation());
        btn.addEventListener("click", (e) => {
            e.stopPropagation();
            e.preventDefault();
            if (!window.DracoCalcEditor) return;
            // Desde un atributo: ventana de cálculos de ese atributo.
            // Desde la dimensión: listado de los cálculos de todos sus atributos.
            window.DracoCalcEditor.open(this, dim, attribute || null, attributes);
        });
        return btn;
    },

    // Marca (ƒx en color) los atributos y dimensiones que tienen cálculos en el informe actual.
    refreshCalcBadges() {
        if (!window.DracoCalc) return;
        const fo = this.state.fieldOptions || {};
        const withCalcs = new Set(Object.keys(fo)
            .filter(k => window.DracoCalc.isCalcKey(k) && Array.isArray(fo[k]) && fo[k].length)
            .map(k => k.split("|").slice(1).join("|")));
        const dimsWithCalcs = new Set(Array.from(withCalcs).map(k => k.split("|")[0]));
        document.querySelectorAll(".field-fx-btn").forEach(b => {
            const has = b.dataset.attr
                ? withCalcs.has(b.dataset.dim + "|" + b.dataset.attr)
                : dimsWithCalcs.has(b.dataset.dim);
            b.classList.toggle("has-calcs", has);
        });
    },

    createFieldElement(dim, name, isHierarchy, dimAttributes) {
        const div = document.createElement("div");
        div.className = "field-item";
        div.draggable = true;

        const iconClass = isHierarchy ? "field-icon hierarchy-icon" : "field-icon";
        const iconSymbol = isHierarchy ? "🗂️" : "📄";

        // SEGURIDAD: "name" puede venir de un modelo semántico importado
        // (.lkml de un repo compartido), no solo del esquema de BigQuery.
        const iconSpan = document.createElement("span");
        iconSpan.className = iconClass;
        iconSpan.textContent = iconSymbol;
        const labelSpan = document.createElement("span");
        labelSpan.className = "field-label";
        labelSpan.textContent = name.toLowerCase();
        div.appendChild(iconSpan);
        div.appendChild(labelSpan);
        if (!isHierarchy) div.appendChild(this.createCalcButton(dim, name, [name]));

        const fieldData = { dim, name, isHierarchy };

        div.addEventListener("dragstart", (e) => {
            this.draggedElementData = { data: fieldData, sourceTag: null, sourceZone: null };
            div.classList.add("dragging");
            e.dataTransfer.effectAllowed = "move";
            e.dataTransfer.setData("text/plain", JSON.stringify(fieldData));
        });

        div.addEventListener("dragend", () => div.classList.remove("dragging"));

        return div;
    },

    filterFields(query) {
        const q = query.toLowerCase();
        document.querySelectorAll(".field-item").forEach(item => {
            const label = item.querySelector(".field-label").innerText.toLowerCase();
            item.style.display = label.includes(q) ? "flex" : "none";
        });
    },

    /* -------------------------------------------------------------
     * Drag & drop entre zonas + reordenar dentro de Filas/Columnas/Filtros
     * ----------------------------------------------------------- */
    handleDragOver(e) {
        const zoneId = e.currentTarget.getAttribute("data-zone");
        if (this.isZoneLocked(zoneId)) return; // eje Estático: no admite drop
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        e.currentTarget.classList.add("drag-over");

        // Reordenar EN VIVO dentro de la MISMA zona: solo aplica cuando se
        // está arrastrando una tag ya colocada (sourceTag), no un campo
        // nuevo desde la lista de la izquierda.
        const dragged = this.draggedElementData;
        if (dragged && dragged.isMeasureBlock && dragged.sourceZone === zoneId) {
            // Bloque "Σ Medidas" completo: se coloca entre las dimensiones
            // del eje (encima o debajo de cualquiera de ellas).
            const dropzoneContent = e.currentTarget.querySelector(".dropzone-content");
            if (dropzoneContent) {
                const afterElement = this.getDragAfterElement(dropzoneContent, e.clientY);
                if (afterElement == null) {
                    dropzoneContent.appendChild(dragged.sourceTag);
                } else if (afterElement !== dragged.sourceTag) {
                    dropzoneContent.insertBefore(dragged.sourceTag, afterElement);
                }
            }
        } else if (dragged && dragged.sourceTag && dragged.sourceZone === zoneId) {
            const dropzoneContent = e.currentTarget.querySelector(".dropzone-content");
            if (dropzoneContent) {
                // Una medida se reordena SOLO dentro de su propio grupo
                // "Σ Medidas" (no se compara su posición con las tags de
                // dimensiones sueltas del mismo eje); el resto de campos se
                // reordena entre sí, directamente en la dropzone.
                const container = this.isMeasureField(dragged.data)
                    ? this.ensureMeasureGroup(dropzoneContent, zoneId)
                    : dropzoneContent;
                const afterElement = this.getDragAfterElement(container, e.clientY);
                if (afterElement == null) {
                    container.appendChild(dragged.sourceTag);
                } else if (afterElement !== dragged.sourceTag) {
                    container.insertBefore(dragged.sourceTag, afterElement);
                }
            }
        }
    },

    // Devuelve la tag después de la cual habría que insertar el elemento
    // arrastrado, según la posición vertical del cursor (algoritmo estándar
    // de listas "sortable" con drag&drop nativo del navegador). Solo mira
    // los ".dropped-tag" hijos DIRECTOS de "container" (con :scope), para no
    // colarse dentro del grupo de medidas al reordenar dimensiones sueltas,
    // ni al revés.
    getDragAfterElement(container, y) {
        // El bloque "Σ Medidas" cuenta como un elemento más: una dimensión
        // puede quedar encima o debajo de él, y el bloque entre dimensiones.
        const elements = [...container.querySelectorAll(":scope > .dropped-tag:not(.dragging), :scope > .measure-group:not(.dragging)")];
        return elements.reduce((closest, child) => {
            const box = child.getBoundingClientRect();
            const offset = y - box.top - box.height / 2;
            if (offset < 0 && offset > closest.offset) {
                return { offset, element: child };
            }
            return closest;
        }, { offset: Number.NEGATIVE_INFINITY, element: null }).element;
    },

    // ¿Es "MEASURE" el campo que se está arrastrando/soltando? (data trae
    // { dim, name, isHierarchy }, igual que las entradas de state.rows/cols).
    isMeasureField(data) {
        return !!(data && String(data.dim).toUpperCase() === "MEASURE");
    },

    // Busca (o crea) el contenedor ".measure-group-content" de una zona:
    // caja propia, con su propio título "Σ Medidas", donde viven TODAS las
    // medidas de ese eje, agrupadas y separadas del resto de campos, para
    // poder quitarlas/ordenarlas entre sí de un vistazo.
    ensureMeasureGroup(dropzoneContent, zoneId) {
        let group = dropzoneContent.querySelector(":scope > .measure-group");
        if (!group) {
            group = document.createElement("div");
            group.className = "measure-group";
            group.dataset.zone = zoneId;
            group.innerHTML = `<div class="measure-group-title" title="Arrastra para mover todas las medidas">Σ Medidas</div><div class="measure-group-content"></div>`;
            // El bloque entero se arrastra (desde su título o su borde): se
            // recoloca entre las dimensiones del eje o pasa al otro eje con
            // todas sus medidas. Cada medida sigue arrastrándose sola dentro
            // del bloque (su dragstart para la propagación).
            group.draggable = !this.isZoneLocked(zoneId);
            group.addEventListener("dragstart", (e) => {
                const zone = group.dataset.zone;
                this.draggedElementData = {
                    data: { dim: "MEASURE", name: "", isHierarchy: false },
                    sourceTag: group,
                    sourceZone: zone,
                    isMeasureBlock: true
                };
                group.classList.add("dragging");
                e.dataTransfer.effectAllowed = "move";
                e.dataTransfer.setData("text/plain", "draco-measure-block");
            });
            group.addEventListener("dragend", () => group.classList.remove("dragging"));
            dropzoneContent.appendChild(group);
        }
        return group.querySelector(".measure-group-content");
    },

    // Quita el grupo "Σ Medidas" de una zona en cuanto se queda vacío (se
    // quitó la última medida de ese eje).
    cleanupMeasureGroup(dropzoneContent) {
        const group = dropzoneContent.querySelector(":scope > .measure-group");
        if (group && group.querySelector(".measure-group-content").children.length === 0) {
            group.remove();
        }
    },

    // Regla: las medidas nunca pueden quedar repartidas entre Filas y
    // Columnas a la vez (todas juntas en el mismo eje). Comprueba, para un
    // drop de una medida en "targetZoneId", si el eje CONTRARIO se
    // quedaría con alguna medida después de esta operación (contando con
    // que, si la medida viene de ese mismo eje contrario, se está
    // moviendo fuera de él).
    wouldSplitMeasuresAcrossAxes(targetZoneId, data, sourceZone) {
        if (targetZoneId !== "rows" && targetZoneId !== "columns") return false;
        if (!this.isMeasureField(data)) return false;

        const otherZone = targetZoneId === "rows" ? "columns" : "rows";
        const otherList = this.listForZone(otherZone);
        const remaining = otherList.filter(entry => {
            if (String(entry.dimension).toUpperCase() !== "MEASURE") return false;
            const isTheOneBeingMoved = sourceZone === otherZone && entry.dimension === data.dim && entry.name === data.name;
            return !isTheOneBeingMoved;
        });
        return remaining.length > 0;
    },

    // Aviso visual breve (borde rojo + texto en el estado de autoguardado)
    // cuando se rechaza un drop por la regla anterior.
    flashZoneWarning(zoneId, message) {
        const card = document.querySelector(`.zone-card[data-zone="${zoneId}"]`);
        if (card) {
            card.classList.add("zone-warning");
            setTimeout(() => card.classList.remove("zone-warning"), 900);
        }
        this.setAutoStatus(message);
        setTimeout(() => {
            const el = document.getElementById("autoStatus");
            if (el && el.innerText === message) el.innerText = "";
        }, 2500);
    },

    handleDragLeave(e) {
        e.currentTarget.classList.remove("drag-over");
    },

    handleDrop(e) {
        const dropzoneBox = e.currentTarget;
        const zoneId = dropzoneBox.getAttribute("data-zone");
        if (this.isZoneLocked(zoneId)) return; // eje Estático: no admite drop

        e.preventDefault();
        dropzoneBox.classList.remove("drag-over");

        if (!this.draggedElementData) {
            try {
                const rawData = e.dataTransfer.getData("text/plain");
                if (rawData) this.draggedElementData = { data: JSON.parse(rawData), sourceTag: null, sourceZone: null };
            } catch (err) {
                console.warn("No se pudo leer dataTransfer raw:", err);
            }
        }
        if (!this.draggedElementData) return;

        const { data, sourceTag, sourceZone, isMeasureBlock } = this.draggedElementData;
        const targetContent = dropzoneBox.querySelector(".dropzone-content");

        if (isMeasureBlock) {
            this.draggedElementData = null;
            if (sourceZone === zoneId) {
                this.syncStateOrderFromDom(zoneId, targetContent);
                return;
            }
            if (zoneId !== "rows" && zoneId !== "columns") {
                this.flashZoneWarning(zoneId, "El bloque de medidas solo puede ir a Filas o Columnas.");
                return;
            }
            if (this.isZoneLocked(sourceZone)) return;
            this.moveMeasureBlock(sourceZone, zoneId);
            return;
        }

        if (sourceZone === zoneId) {
            // Mismo eje: handleDragOver ya reordenó la tag en el DOM en
            // vivo; aquí solo hace falta reflejar ese nuevo orden en el
            // estado (this.state.rows/columns/filters) y relanzar el
            // guardado + actualización, ya que el orden de los campos
            // determina el orden de las columnas/niveles del informe.
            if (sourceTag) this.syncStateOrderFromDom(zoneId, targetContent);
            this.draggedElementData = null;
            return;
        }

        // El eje de origen también puede estar bloqueado (no se puede sacar
        // un campo de un eje Estático arrastrándolo a otro sitio).
        if (sourceZone && this.isZoneLocked(sourceZone)) {
            this.draggedElementData = null;
            return;
        }

        // Regla: nunca una medida a Filas y otra a Columnas. Si esta medida
        // dejaría alguna medida en el eje contrario, se rechaza el drop.
        if (this.wouldSplitMeasuresAcrossAxes(zoneId, data, sourceZone)) {
            this.flashZoneWarning(zoneId, "Todas las medidas deben ir en el mismo eje (Filas o Columnas).");
            this.draggedElementData = null;
            return;
        }

        // Dimensiones entre Filtros y Filas/Columnas: se COPIAN (el campo
        // sigue donde estaba y además aparece en el destino). Así un filtro
        // no se pierde al llevar la dimensión a Filas, ni al revés. Las
        // medidas, y los movimientos entre Filas y Columnas, siguen moviendo.
        const crossesFilters = !!sourceZone && ((sourceZone === "filters") !== (zoneId === "filters"));
        if (sourceTag && crossesFilters && !this.isMeasureField(data)) {
            if (this.listForZone(zoneId).some(x => x.dimension === data.dim && x.name === data.name)) {
                this.flashZoneWarning(zoneId, "Ese campo ya está en esta zona.");
            } else {
                this.addField(targetContent, zoneId, data);
            }
            this.draggedElementData = null;
            return;
        }

        if (sourceTag) {
            sourceTag.remove();
            this.removeFromState(sourceZone, data);
            if (sourceZone === "rows" || sourceZone === "columns") {
                const sourceContent = document.querySelector(`.zone-card[data-zone="${sourceZone}"] .dropzone-content`);
                if (sourceContent) this.cleanupMeasureGroup(sourceContent);
            }
        }

        this.addField(targetContent, zoneId, data);

        this.draggedElementData = null;
    },

    /**
     * Mueve TODAS las medidas de un eje al otro (arrastrando el bloque
     * "Σ Medidas"), en el mismo orden, al final del eje de destino.
     */
    moveMeasureBlock(fromZone, toZone) {
        const fromContent = document.querySelector(`.zone-card[data-zone="${fromZone}"] .dropzone-content`);
        const toContent = document.querySelector(`.zone-card[data-zone="${toZone}"] .dropzone-content`);
        if (!fromContent || !toContent) return;
        const measures = this.listForZone(fromZone)
            .filter(x => String(x.dimension).toUpperCase() === "MEASURE")
            .map(x => ({ dim: x.dimension, name: x.name, isHierarchy: x.isHierarchy }));
        if (!measures.length) return;
        measures.forEach(m => {
            const tag = [...fromContent.querySelectorAll(".dropped-tag")]
                .find(t => t.dataset.dim === m.dim && t.dataset.fieldName === m.name);
            if (tag) tag.remove();
            this.removeFromState(fromZone, m);
        });
        this.cleanupMeasureGroup(fromContent);
        measures.forEach(m => this.addField(toContent, toZone, m));
    },

    /** Vuelve a pintar la zona Filtros a partir de this.state.filters. */
    rerenderFiltersZone() {
        const filtersContent = document.querySelector('.zone-card[data-zone="filters"] .dropzone-content');
        if (!filtersContent) return;
        filtersContent.innerHTML = "";
        this.state.filters.forEach(f => this.renderTag(filtersContent, "filters", f));
        if (this.renderLockedFilterTags) this.renderLockedFilterTags(filtersContent);
    },

    /**
     * Doble clic en una dimensión de Filas/Columnas: abre el mismo diálogo
     * de filtro que en la zona Filtros y guarda la selección EN Filtros
     * (crea el filtro si no estaba; si ya estaba, lo sustituye). El campo
     * sigue en Filas/Columnas.
     */
    async openFilterForAxisField(entry) {
        if (typeof FilterModal === "undefined" || !FilterModal.open) return;
        const same = f => f.dimension === entry.dimension && f.name === entry.name;
        const existing = this.state.filters.find(same);
        const result = await FilterModal.open({
            dim: entry.dimension,
            name: entry.name,
            isHierarchy: entry.isHierarchy,
            currentFilter: existing ? existing.filter : null
        });
        if (!result) return;
        if (existing) {
            existing.filter = result;
            if (!existing.isHierarchy) existing.realAttribute = result.attribute || existing.name;
            this.rerenderFiltersZone();
            this.scheduleAutoUpdate();
        } else {
            const filtersContent = document.querySelector('.zone-card[data-zone="filters"] .dropzone-content');
            if (!filtersContent) return;
            this.addField(filtersContent, "filters", { dim: entry.dimension, name: entry.name, isHierarchy: entry.isHierarchy });
            const added = this.state.filters.find(same);
            if (added) {
                added.filter = result;
                if (!added.isHierarchy) added.realAttribute = result.attribute || added.name;
            }
            this.rerenderFiltersZone();
            this.scheduleAutoUpdate();
        }
    },

    // Reconstruye this.state.<zona> en el mismo orden en que han quedado
    // las tags en el DOM tras un arrastre de reordenación dentro del mismo
    // eje, y dispara el autoguardado + autoactualización.
    syncStateOrderFromDom(zoneId, container) {
        const list = this.listForZone(zoneId);
        const domOrder = [...container.querySelectorAll(".dropped-tag")];
        const newList = domOrder
            .map(tag => list.find(x => x.dimension === tag.dataset.dim && x.name === tag.dataset.fieldName))
            .filter(Boolean);
        // Por seguridad, si algo no casó (no debería pasar), no se pierde
        // ningún campo: se completa con los que faltasen al final.
        list.forEach(entry => {
            if (!newList.includes(entry)) newList.push(entry);
        });
        list.length = 0;
        list.push(...newList);
        this.scheduleAutoUpdate({ saveOnly: zoneId === "filters" }); // el orden de los filtros no cambia el informe
    },

    /* -------------------------------------------------------------
     * Estado + render de un campo añadido a una zona
     * ----------------------------------------------------------- */
    addField(container, zoneId, data) {
        const list = this.listForZone(zoneId);
        const already = list.find(x => x.dimension === data.dim && x.name === data.name);
        if (already) return;

        const entry = { dimension: data.dim, name: data.name, isHierarchy: data.isHierarchy };

        if (zoneId === "filters") {
            entry.realAttribute = data.isHierarchy ? "" : data.name;
            entry.filter = null; // sin selección todavía: doble clic para elegir
        }

        list.push(entry);

        // Todas las medidas de un eje comparten UNA sola fila/columna física
        // de cabecera ("Σ Medidas", ver computeAxisPaintPlan en commands.js):
        // el rango de Filas/Columnas del taskpane solo debe crecer al añadir
        // la PRIMERA medida de ese eje. Si ya había otra medida, esta nueva
        // se apila en el MISMO hueco (no en uno nuevo), así que hacer crecer
        // el rango otra vez dejaba una fila/columna de más (en blanco) sin
        // usar. Los campos normales (dimensiones) siguen creciendo el rango
        // siempre, como hasta ahora.
        const isMeasure = String(entry.dimension).toUpperCase() === "MEASURE";
        const measureCountInZone = (zoneId === "rows" || zoneId === "columns")
            ? list.filter(x => String(x.dimension).toUpperCase() === "MEASURE").length
            : 0;
        const isFirstMeasureInZone = isMeasure && measureCountInZone === 1;
        const shouldGrowRange = !isMeasure || isFirstMeasureInZone;

        if (shouldGrowRange) {
            if (zoneId === "rows") RangeAxis.onRowFieldAdded(this.state.colsStatic && !this.state.rowsStatic);
            if (zoneId === "columns") RangeAxis.onColFieldAdded(this.state.rowsStatic && !this.state.colsStatic);
        }
        this.reconcileTextColumns();
        this.refreshRangeLabels();

        this.renderTag(container, zoneId, entry);

        // Un filtro recién soltado aún no tiene valor (se ignora en el WHERE
        // hasta que se elija uno): el informe no cambia, solo se guarda el
        // diseño. Añadir campos a Filas o Columnas sí lo cambia: actualizar.
        this.scheduleAutoUpdate({ saveOnly: zoneId === "filters" && !entry.filter });
    },

    removeFromState(zoneId, data) {
        const list = this.listForZone(zoneId);
        const idx = list.findIndex(x => x.dimension === data.dim && x.name === data.name);
        const removed = idx !== -1 ? list[idx] : null;
        if (idx !== -1) list.splice(idx, 1);

        // Simétrico a addField: solo se encoge el rango si se ha quitado un
        // campo normal, o si la medida quitada era la ÚLTIMA medida de ese
        // eje. Mientras quede al menos otra medida en el eje, el hueco
        // compartido "Σ Medidas" sigue haciendo falta, así que el rango no
        // debe encogerse todavía.
        const isMeasure = removed && String(removed.dimension).toUpperCase() === "MEASURE";
        const measureCountLeftInZone = (zoneId === "rows" || zoneId === "columns")
            ? list.filter(x => String(x.dimension).toUpperCase() === "MEASURE").length
            : 0;
        const shouldShrinkRange = !isMeasure || measureCountLeftInZone === 0;

        if (shouldShrinkRange) {
            if (zoneId === "rows") RangeAxis.onRowFieldRemoved(this.state.colsStatic && !this.state.rowsStatic);
            if (zoneId === "columns") RangeAxis.onColFieldRemoved(this.state.rowsStatic && !this.state.colsStatic);
        }
        this.reconcileTextColumns();
        this.refreshRangeLabels();

        // Quitar de Filtros un filtro sin valor no cambia el informe
        this.scheduleAutoUpdate({ saveOnly: zoneId === "filters" && !!removed && !removed.filter });
    },

    listForZone(zoneId) {
        if (zoneId === "filters") return this.state.filters;
        if (zoneId === "rows") return this.state.rows;
        if (zoneId === "columns") return this.state.columns;
        return [];
    },

    renderTag(container, zoneId, entry) {
        const isMeasure = String(entry.dimension).toUpperCase() === "MEASURE";

        const tag = document.createElement("div");
        tag.className = "dropped-tag";
        if (isMeasure) tag.classList.add("measure-tag");
        tag.draggable = true;
        tag.dataset.dim = entry.dimension;
        tag.dataset.fieldName = entry.name;
        tag.dataset.isHierarchy = entry.isHierarchy;

        // Una medida se muestra solo con su nombre ("Importe"), sin el
        // prefijo de la pseudo-dimensión ("MEASURE.Importe"); el resto de
        // campos sigue mostrando "dimensión.nombre" como hasta ahora.
        const fieldLabel = isMeasure ? entry.name : `${entry.dimension}.${entry.name}`;

        // Filtro sin valor seleccionado todavía: se muestra vacío (no se
        // añade al WHERE de la consulta hasta que el usuario elija algo con
        // doble clic). Si ya hay filtro, se resume (varios valores, rango,
        // incluir/excluir…) con describeFilter(). Una MEDIDA en Filtros no
        // tiene doble clic (no hay miembros que elegir, ver más abajo), así
        // que tampoco muestra el texto "(vacío · doble clic para elegir)".
        const filterSummary = zoneId === "filters" && !isMeasure && typeof window.describeFilter === "function"
            ? window.describeFilter(entry.filter)
            : "";
        const titleText = zoneId === "filters"
            ? (isMeasure ? fieldLabel : (filterSummary ? `${fieldLabel}: ${filterSummary}` : `${fieldLabel}: (vacío · doble clic para elegir)`))
            : fieldLabel;

        // SEGURIDAD: titleText puede incluir valores de miembro reales
        // (describeFilter -> valores elegidos en memberPicker, que a su vez
        // provienen de datos de BigQuery, no de código de este
        // add-in). NUNCA insertar ese texto con innerHTML: si una fila de
        // datos contiene marcado HTML (p.ej. "<img src=x onerror=...>"),
        // se ejecutaría en el contexto del add-in con acceso a
        // localStorage (tokens OAuth/GitHub) y a la API de Office. Se
        // construyen los nodos aparte y se usa textContent para el texto
        // que no es de confianza.
        const titleSpan = document.createElement("span");
        titleSpan.className = "dropped-tag-title";
        titleSpan.textContent = titleText;

        const removeSpan = document.createElement("span");
        removeSpan.className = "dropped-tag-remove";
        removeSpan.textContent = "\u00d7"; // ×

        tag.appendChild(titleSpan);
        tag.appendChild(removeSpan);

        tag.addEventListener("dragstart", (e) => {
            e.stopPropagation();
            this.draggedElementData = {
                data: { dim: entry.dimension, name: entry.name, isHierarchy: entry.isHierarchy },
                sourceTag: tag,
                sourceZone: zoneId
            };
            tag.classList.add("dragging");
            e.dataTransfer.effectAllowed = "move";
            e.dataTransfer.setData("text/plain", JSON.stringify(this.draggedElementData.data));
        });

        tag.addEventListener("dragend", () => tag.classList.remove("dragging"));

        // Solo los filtros abren el modal de selección de valor (doble clic).
        // Una MEDIDA en Filtros no tiene miembros que elegir (no es una
        // dimensión con valores): no se abre el FilterModal para ella.
        if (zoneId === "filters" && !isMeasure) {
            tag.addEventListener("dblclick", async () => {
                if (typeof FilterModal === "undefined" || !FilterModal.open) return;

                const result = await FilterModal.open({
                    dim: entry.dimension,
                    name: entry.name,
                    isHierarchy: entry.isHierarchy,
                    currentFilter: entry.filter
                });
                if (result) {
                    entry.filter = result;
                    if (!entry.isHierarchy) entry.realAttribute = result.attribute || entry.name;
                    const summary = window.describeFilter ? window.describeFilter(result) : "";
                    tag.querySelector(".dropped-tag-title").innerText = `${fieldLabel}: ${summary}`;
                    this.scheduleAutoUpdate();
                }
            });
        }

        tag.querySelector(".dropped-tag-remove").addEventListener("click", (e) => {
            e.stopPropagation();
            tag.remove();
            this.removeFromState(zoneId, { dim: entry.dimension, name: entry.name });
            if (zoneId === "rows" || zoneId === "columns") this.cleanupMeasureGroup(container);
        });

        // Doble clic en una dimensión de Filas/Columnas: filtro sobre ella
        // (se guarda en la zona Filtros, ver openFilterForAxisField).
        if ((zoneId === "rows" || zoneId === "columns") && !isMeasure) {
            tag.title = "Doble clic: filtrar este campo";
            tag.addEventListener("dblclick", (e) => {
                if (e.target.closest(".dropped-tag-remove")) return;
                e.stopPropagation();
                this.openFilterForAxisField(entry);
            });
        }

        // Clic derecho (Filas/Columnas): Mostrar → ID / ID y descripción, si el
        // campo es la clave de una dimensión con texto, o una jerarquía.
        if ((zoneId === "rows" || zoneId === "columns") && !isMeasure) {
            tag.addEventListener("contextmenu", async (e) => {
                e.preventDefault();
                e.stopPropagation();
                const x = e.clientX, y = e.clientY;
                const info = window.ReportStore && window.ReportStore.getDisplayTextInfo
                    ? window.ReportStore.getDisplayTextInfo(this.currentReportId, entry.dimension, entry.name, !!entry.isHierarchy)
                    : null;
                let levels = null;
                if (entry.isHierarchy) {
                    try { levels = await window.ExcelService.getHierarchyLevels(entry.dimension, entry.name); }
                    catch (err) { console.error("Niveles de la jerarquía:", err); }
                }
                this.showDisplayModeMenu(x, y, zoneId, entry, info, levels);
            });
        }

        // Clic (simple) sobre el campo: abre/actualiza el panel "Opciones de
        // campo" con los controles de ESE campo (medida o dimensión).
        tag.classList.add("clickable-for-options");
        tag.addEventListener("click", (e) => {
            if (e.target.closest(".dropped-tag-remove")) return;
            this.selectFieldForOptions(zoneId, entry);
        });

        // Las medidas de Filas/Columnas se agrupan en su propia caja
        // "Σ Medidas" (creada bajo demanda), separadas del resto de campos
        // del eje; el resto de campos (y las medidas en Filtros, si las
        // hubiera) van directos a la dropzone, como hasta ahora.
        const appendTarget = (isMeasure && (zoneId === "rows" || zoneId === "columns"))
            ? this.ensureMeasureGroup(container, zoneId)
            : container;
        appendTarget.appendChild(tag);
    },

    /* -------------------------------------------------------------
     * Etiquetas "bloqueadas" en la zona Filtros: uno por cada rango con
     * nombre creado con "Añadir filtro" (FilterRangeStore) que aplique al
     * informe actual (atado a este informe, o creado para "Todos los
     * informes"). A diferencia de renderTag(), NO son arrastrables ni se
     * pueden quitar con un simple clic en la "×" — hace falta confirmar,
     * porque quitar una de estas etiquetas borra también el rango con
     * nombre en Excel, no solo la fila de diseño de este informe.
     * ----------------------------------------------------------- */
    renderLockedFilterTags(filtersContent) {
        if (!filtersContent || !window.FilterRangeStore) return;

        const all = window.FilterRangeStore.getAll();
        Object.keys(all).forEach(rangeName => {
            const meta = all[rangeName];
            if (!meta) return;

            const isAll = meta.reportId === null || meta.reportId === undefined;
            const appliesToThisReport = isAll
                || (this.currentReportId !== null && this.currentReportId !== undefined
                    && Number(meta.reportId) === Number(this.currentReportId));
            if (!appliesToThisReport) return;

            this.renderLockedFilterTag(filtersContent, rangeName, meta);
        });
    },

    renderLockedFilterTag(container, rangeName, meta) {
        const tag = document.createElement("div");
        tag.className = "dropped-tag locked-filter-tag";
        tag.draggable = false;
        tag.style.borderLeft = "3px solid #A80000";
        tag.style.background = "#FBEAEA";
        tag.dataset.rangeName = rangeName;

        const fieldLabel = `${meta.dim}.${meta.name}`;
        const filterSummary = typeof window.describeFilter === "function" ? window.describeFilter(meta.filter) : "";
        const scopeLabel = (meta.reportId === null || meta.reportId === undefined) ? "Todos los informes" : `Informe ${meta.reportId}`;
        const titleText = filterSummary
            ? `🔒 ${fieldLabel}: ${filterSummary}`
            : `🔒 ${fieldLabel}: (vacío · doble clic para elegir)`;

        // SEGURIDAD: titleText (y fieldLabel/filterSummary dentro de él)
        // puede contener valores reales de datos del cliente -> nunca por
        // innerHTML. Se construyen los nodos aparte; textContent y la
        // propiedad .title escapan automáticamente su contenido, a
        // diferencia de interpolar dentro de un atributo HTML en un
        // template literal (que además permitiría "escapar" del atributo
        // con un valor que contuviera comillas).
        const titleSpan = document.createElement("span");
        titleSpan.className = "dropped-tag-title";
        titleSpan.title = `Creado con 'Añadir filtro' (${scopeLabel}). Bloqueado: no se arrastra ni se reordena; doble clic para cambiar los valores.`;
        titleSpan.textContent = titleText;

        const removeSpan = document.createElement("span");
        removeSpan.className = "dropped-tag-remove";
        removeSpan.title = "Eliminar este filtro y su rango con nombre en Excel";
        removeSpan.textContent = "\u00d7"; // ×

        tag.appendChild(titleSpan);
        tag.appendChild(removeSpan);

        tag.addEventListener("dblclick", () => this.openLockedFilterRangePicker(rangeName, meta));

        tag.querySelector(".dropped-tag-remove").addEventListener("click", async (e) => {
            e.stopPropagation();
            const ok = await this.showConfirmDialog(
                `¿Eliminar el filtro "${fieldLabel}"?\n\nEsto borra también el rango con nombre "${rangeName}" en Excel (no solo la etiqueta).`,
                { title: "Eliminar filtro", confirmLabel: "Eliminar" }
            );
            if (!ok) return;
            await this.deleteLockedFilterRange(rangeName);
        });

        container.appendChild(tag);
    },

    async openLockedFilterRangePicker(rangeName, meta) {
        if (typeof resolveDracoFilterRangeAddress !== "function" || typeof openDracoFilterRangePicker !== "function") {
            alert("No se encuentra la lógica de 'Añadir filtro' (commands.js). Recarga el panel e inténtalo de nuevo.");
            return;
        }
        const loc = await resolveDracoFilterRangeAddress(rangeName);
        if (!loc) {
            alert(`No se encuentra el rango "${rangeName}" en el libro (¿se borró la hoja o el nombre?).`);
            return;
        }
        await openDracoFilterRangePicker(loc.addr, loc.sheetName, rangeName, meta);
        await this.loadDesignFromSheet();
    },

    async deleteLockedFilterRange(rangeName) {
        try {
            await Excel.run(async (context) => {
                const item = context.workbook.names.getItemOrNullObject(rangeName);
                item.load("isNullObject");
                await context.sync();
                if (!item.isNullObject) {
                    item.delete();
                    await context.sync();
                }
            });
            if (window.FilterRangeStore) await window.FilterRangeStore.remove(rangeName);
            await this.loadDesignFromSheet();
        } catch (err) {
            console.error("Error al eliminar el filtro bloqueado:", err);
            alert("Error al eliminar el filtro: " + (err.message || err));
        }
    },

    /* =================================================================
     * PROPIEDADES DEL INFORME (modal): nombre del informe, suprimir
     * ceros en filas/columnas, subtotales arriba, sobrescribir formatos,
     * autoajustar columnas. Se guardan en Office roaming settings, para
     * que también las pueda leer commands.js (ribbon y jsonTo3Matrices).
     * ================================================================= */
    // Lee las propiedades del informe ACTUAL (this.currentReportId) desde
    // ReportStore. Cada informe tiene las suyas (antes había una única
    // clave global "draco_reportProperties" compartida por todo el libro).
    loadReportPropertiesFromSettings() {
        try {
            if (this.currentReportId && window.ReportStore) {
                const report = window.ReportStore.getReport(this.currentReportId);
                if (report && report.reportProperties) {
                    this.reportProperties = Object.assign({}, this.reportProperties, report.reportProperties);
                }
            }
        } catch (err) {
            console.warn("No se pudieron leer las propiedades del informe:", err);
        }
    },

    async saveReportPropertiesToSettings() {
        if (!this.currentReportId || !window.ReportStore) return;
        await window.ReportStore.saveReportProperties(this.currentReportId, this.reportProperties);
        await this.populateReportSelector(); // por si ha cambiado el nombre del informe
    },

    // Rellena el listbox "Modelo semántico" del modal de Propiedades del
    // informe (informativo: el selector real vive ahora encima del
    // buscador de campo). Se preselecciona el modelo del informe actual.
    async populateSemanticModelDropdown() {
        const select = document.getElementById("propSemanticModel");
        if (!select) return;
        select.innerHTML = "<option>Cargando…</option>";
        try {
            const models = (window.ExcelService && window.ExcelService.getSemanticModels)
                ? await window.ExcelService.getSemanticModels()
                : [];
            select.innerHTML = "";
            if (models.length === 0) {
                select.innerHTML = "<option value=\"\">— Sin modelos semánticos guardados —</option>";
                return;
            }
            models.forEach(name => {
                const opt = document.createElement("option");
                opt.value = name;
                opt.textContent = name;
                select.appendChild(opt);
            });
            const report = (this.currentReportId && window.ReportStore) ? window.ReportStore.getReport(this.currentReportId) : null;
            select.value = report && report.semanticModelName ? report.semanticModelName : models[0];
        } catch (err) {
            console.warn("No se pudieron cargar los modelos semánticos guardados:", err);
            select.innerHTML = "<option value=\"\">— No disponible —</option>";
        }
    },

    async openReportPropertiesModal() {
        const modal = document.getElementById("reportPropertiesModal");
        if (!modal) return;

        await this.populateSemanticModelDropdown();

        document.getElementById("propReportName").value = this.reportProperties.reportName || "Report 001";
        document.getElementById("propSuppressZeroRows").checked = !!this.reportProperties.suppressZeroRows;
        document.getElementById("propSuppressZeroCols").checked = !!this.reportProperties.suppressZeroCols;
        document.getElementById("propSubtotalsOnTop").checked = !!this.reportProperties.subtotalsOnTop;
        document.getElementById("propHierarchyNodesOnTop").checked = this.reportProperties.hierarchyNodesOnTop !== false;
        document.getElementById("propOverwriteFormats").checked = !!this.reportProperties.overwriteFormats;
        document.getElementById("propOverwriteStaticFormats").checked = !!this.reportProperties.overwriteStaticFormats;
        document.getElementById("propAutoFitColumns").checked = !!this.reportProperties.autoFitColumns;
        document.getElementById("propPlanningReport").checked = !!this.reportProperties.planningReport;
        this.applyPlanningGateToProperties();

        await this.toggleDimensionsBehaviorSection();

        modal.style.display = "flex";
        this.updateRibbonToggleLabel("BtnPropiedadesInforme", "Propiedades", true);
    },

    // Solo se muestra si el informe está marcado como "de planificación" —
    // por ahora es solo la parte visual (marca qué dimensiones son
    // imprescindibles y el reparto para el resto), sin enganchar todavía
    // el guardado/uso real de estos valores.
    async toggleDimensionsBehaviorSection() {
        const section = document.getElementById("propDimensionsBehaviorSection");
        if (!section) return;
        const isPlanning = document.getElementById("propPlanningReport").checked;
        section.style.display = isPlanning ? "block" : "none";
        if (isPlanning) await this.renderDimensionsBehaviorRows();
    },

    async renderDimensionsBehaviorRows() {
        const container = document.getElementById("propDimensionsBehaviorContainer");
        if (!container) return;
        container.innerHTML = "<div style='font-size:11px; color:#605e5c;'>Cargando dimensiones…</div>";

        try {
            const modelName = document.getElementById("propSemanticModel").value;
            if (!modelName || !window.SemanticModelStore) {
                container.innerHTML = "<div style='font-size:11px; color:#605e5c;'>Elige primero un modelo semántico.</div>";
                return;
            }

            const grid = await window.SemanticModelStore.getModelGrid("MODEL_DIMENSION", modelName);
            const dims = [];
            const seen = new Set();
            const values = grid && grid.values ? grid.values : [];
            for (let r = 1; r < values.length; r++) { // fila 0 = cabecera
                const dim = values[r] && values[r][1]; // columna DIMENSION
                if (dim && !seen.has(dim)) { seen.add(dim); dims.push(dim); }
            }

            if (dims.length === 0) {
                container.innerHTML = "<div style='font-size:11px; color:#605e5c;'>Este modelo no tiene dimensiones.</div>";
                return;
            }

            const saved = (this.reportProperties.dimensionsBehavior) || {};
            container.innerHTML = "";
            dims.forEach(dim => {
                const cfg = saved[dim] || { required: true, mode: "null" };
                container.appendChild(this.buildDimensionBehaviorRow(dim, cfg));
            });
        } catch (err) {
            console.warn("No se pudieron cargar las dimensiones del modelo para 'Comportamiento de dimensiones':", err);
            container.innerHTML = "<div style='font-size:11px; color:#a80000;'>No se pudieron cargar las dimensiones.</div>";
        }
    },

    // Fila reutilizable "checkbox (imprescindible) + dimensión + criterio
    // de reparto" — la usan tanto "Propiedades del informe" (global) como
    // el panel de Opciones de campo de una medida (override por medida).
    // El criterio solo tiene sentido si la dimensión NO está marcada (si
    // lo está, nunca puede faltar, así que no hay nada que repartir) —
    // por eso se oculta al marcarla.
    buildDimensionBehaviorRow(dim, cfg) {
        const row = document.createElement("div");
        row.className = "dim-behavior-row";
        row.style.cssText = "display:flex; align-items:center; gap:8px; font-size:11px;";
        // SEGURIDAD: "dim" puede venir de un modelo semántico importado
        // (.lkml de un repo compartido) -> se construye con setAttribute/
        // textContent (que escapan su valor) en vez de interpolarlo dentro
        // de un template literal HTML, donde además de XSS permitiría
        // "romper" el atributo data-dim con un valor que contuviera comillas.
        const label = document.createElement("label");
        label.style.cssText = "display:flex; align-items:center; gap:6px; flex:1 1 auto; overflow:hidden;";
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.className = "dim-behavior-required";
        checkbox.setAttribute("data-dim", dim);
        checkbox.checked = !!cfg.required;
        checkbox.style.cssText = "width:auto; flex:0 0 auto;";
        const dimSpan = document.createElement("span");
        dimSpan.style.cssText = "font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;";
        dimSpan.textContent = dim;
        label.appendChild(checkbox);
        label.appendChild(dimSpan);

        const select = document.createElement("select");
        select.className = "dim-behavior-mode";
        select.setAttribute("data-dim", dim);
        select.style.cssText = "width:auto; flex:0 0 auto;" + (cfg.required ? "display:none;" : "");
        [["null", "NULL"], ["lineal", "Lineal"], ["proporcional", "Proporcional"]].forEach(([value, label2]) => {
            const opt = document.createElement("option");
            opt.value = value;
            opt.textContent = label2;
            opt.selected = cfg.mode === value;
            select.appendChild(opt);
        });

        row.appendChild(label);
        row.appendChild(select);
        checkbox.addEventListener("change", () => {
            select.style.display = checkbox.checked ? "none" : "";
        });
        return row;
    },

    openDistributeValuesModal() {
        if (window.DracoPlanningGate) {
            const report = window.ReportStore ? window.ReportStore.getReport(window.ReportStore.getActiveReportId()) : null;
            const gate = DracoPlanningGate.check(report ? report.semanticModelName : "");
            if (!gate.ok) {
                alert(gate.reason);
                return;
            }
        }
        if (window.DistributeValuesPanel) window.DistributeValuesPanel.open();
    },

    /**
     * Planificación según la sesión (DracoPlanningGate): con un token de solo
     * lectura no se muestra nada de planificación (botón "Distribuir valores",
     * casilla "Informe de planificación") y se desactiva "Guardar" en la cinta.
     * Se llama al iniciar el panel y tras cada inicio de sesión.
     */
    applyPlanningGateUI() {
        if (!window.DracoPlanningGate) return;
        const canWrite = DracoPlanningGate.canWrite();
        const btn = document.getElementById("btnDistribuirValores");
        if (btn) btn.style.display = canWrite ? "" : "none";
        DracoPlanningGate.syncRibbon();
    },

    /** Casilla "Informe de planificación" del modal de propiedades según la puerta. */
    applyPlanningGateToProperties() {
        const checkbox = document.getElementById("propPlanningReport");
        if (!checkbox || !window.DracoPlanningGate) return;
        const row = checkbox.closest("label") || checkbox.parentElement;
        const modelName = (document.getElementById("propSemanticModel") || {}).value || "";
        const gate = DracoPlanningGate.check(modelName);
        // Solo lectura: no se muestra. Modelo que no es de Draco Planning: se ve desactivada, con el motivo.
        if (row) row.style.display = DracoPlanningGate.canWrite() ? "" : "none";
        checkbox.disabled = !gate.ok;
        if (row) row.title = gate.ok ? "" : gate.reason;
        if (!gate.ok) {
            checkbox.checked = false;
            const section = document.getElementById("propDimensionsBehaviorSection");
            if (section) section.style.display = "none";
        }
    },

    closeReportPropertiesModal() {
        const modal = document.getElementById("reportPropertiesModal");
        if (modal) modal.style.display = "none";
        this.updateRibbonToggleLabel("BtnPropiedadesInforme", "Propiedades", false);
    },

    async saveReportPropertiesFromModal() {
        // "Comportamiento de dimensiones" — solo la parte visual por ahora
        // (se guarda la selección para que no se pierda al reabrir el
        // modal, pero todavía no se usa en ningún sitio: ni al pintar en
        // cian, ni al guardar planificación).
        const dimensionsBehavior = {};
        document.querySelectorAll("#propDimensionsBehaviorContainer .dim-behavior-row").forEach(row => {
            const dim = row.querySelector(".dim-behavior-required").getAttribute("data-dim");
            dimensionsBehavior[dim] = {
                required: row.querySelector(".dim-behavior-required").checked,
                mode: row.querySelector(".dim-behavior-mode").value
            };
        });

        this.reportProperties = {
            reportName: (document.getElementById("propReportName").value || "Report 001").trim(),
            suppressZeroRows: document.getElementById("propSuppressZeroRows").checked,
            suppressZeroCols: document.getElementById("propSuppressZeroCols").checked,
            subtotalsOnTop: document.getElementById("propSubtotalsOnTop").checked,
            hierarchyNodesOnTop: document.getElementById("propHierarchyNodesOnTop").checked,
            overwriteFormats: document.getElementById("propOverwriteFormats").checked,
            overwriteStaticFormats: document.getElementById("propOverwriteStaticFormats").checked,
            autoFitColumns: document.getElementById("propAutoFitColumns").checked,
            planningReport: document.getElementById("propPlanningReport").checked
                && (!window.DracoPlanningGate || DracoPlanningGate.check(document.getElementById("propSemanticModel").value).ok),
            dimensionsBehavior
        };

        const btn = document.getElementById("btnSaveProperties");
        try {
            if (btn) { btn.disabled = true; btn.innerText = "Guardando…"; }
            await this.saveReportPropertiesToSettings();
            this.closeReportPropertiesModal();

            // Al guardar las propiedades del informe, se refresca el
            // informe para que los cambios (suprimir ceros, subtotales
            // arriba, sobrescribir formatos, autoajustar columnas...) se
            // vean reflejados de inmediato.
            if (window.ReportActions && typeof window.ReportActions.actualizar === "function") {
                this.setAutoStatus("Actualizando…");
                await window.ReportActions.actualizar();
                this.setAutoStatus("Actualizado ✓");
                setTimeout(() => {
                    const el = document.getElementById("autoStatus");
                    if (el && el.innerText === "Actualizado ✓") el.innerText = "";
                }, 2000);
            }
        } catch (err) {
            console.error("Error al guardar las propiedades del informe:", err);
            alert("Error al guardar: " + (err.message || err));
        } finally {
            if (btn) { btn.disabled = false; btn.innerText = "Guardar"; }
        }
    },

    /* =================================================================
     * ACCIÓN PENDIENTE DESDE EL RIBBON: los botones "Propiedades del
     * informe" y "Opciones de campo" del ribbon dejan marcada una acción
     * en Office roaming settings y abren/traen al frente el taskpane;
     * aquí se recoge esa marca al arrancar y se abre lo que corresponda.
     * ================================================================= */
    async handlePendingRibbonAction() {
        try {
            const settings = Office.context.document.settings;
            const pending = settings.get("draco_pendingAction");
            if (!pending) return;

            settings.remove("draco_pendingAction");
            await new Promise((resolve) => settings.saveAsync(resolve));

            if (pending === "properties") {
                this.openReportPropertiesModal();
            } else if (pending === "fieldOptions") {
                this.setFieldOptionsPanelOpen(true);
            } else if (pending === "distributeValues") {
                this.openDistributeValuesModal();
            } else if (pending === "addFilterRange") {
                this.openAddFilterRangeModal();
            }
        } catch (err) {
            console.warn("No se pudo procesar la acción pendiente del ribbon:", err);
        }
    },

    /* =================================================================
     * PANEL "OPCIONES DE CAMPO" (columna derecha, solo estético: al
     * mostrarse amplía el ancho del cuerpo del taskpane, al ocultarse lo
     * reduce). Un clic en un campo de Filtros/Filas/Columnas muestra sus
     * opciones aquí; el propio panel no dispara ningún refresco.
     * ================================================================= */
    isFieldOptionsPanelOpen() {
        const container = document.querySelector(".taskpane-container");
        return !!(container && container.classList.contains("field-options-open"));
    },

    setFieldOptionsPanelOpen(open) {
        const container = document.querySelector(".taskpane-container");
        const btn = document.getElementById("btnFieldOptions");
        if (container) container.classList.toggle("field-options-open", open);
        // Más ancho con las propiedades del campo a la vista; normal al cerrarlas
        if (window.dracoSetTaskpaneWidth) window.dracoSetTaskpaneWidth(open ? "wide" : "normal");
        if (btn) btn.classList.toggle("toggle-active", open);
        if (!open) this.selectedFieldForOptions = null;
        this.updateRibbonToggleLabel("BtnOpcionesCampo", "Opciones de campo", open);
    },

    toggleFieldOptionsPanel() {
        this.setFieldOptionsPanelOpen(!this.isFieldOptionsPanelOpen());
    },

    fieldOptionsKey(zoneId, entry) {
        return `${zoneId}|${entry.dimension}|${entry.name}`;
    },

    getFieldOptions(zoneId, entry) {
        const key = this.fieldOptionsKey(zoneId, entry);
        const isMeasure = String(entry.dimension).toUpperCase() === "MEASURE";
        const defaults = isMeasure
            // aggregation "" = la del modelo semántico (no se sobrescribe).
            ? { numberFormat: "number", decimals: 2, aggregation: "" }
            : { showTotals: false, sortOrder: "none", expandToLevel: null, visibleLevels: null, displayMode: "ID" };
        return Object.assign({}, defaults, this.state.fieldOptions[key] || {});
    },

    setFieldOptions(zoneId, entry, options) {
        const key = this.fieldOptionsKey(zoneId, entry);
        this.state.fieldOptions[key] = options;
    },

    async selectFieldForOptions(zoneId, entry) {
        this.selectedFieldForOptions = { zoneId, dimension: entry.dimension, name: entry.name, isHierarchy: entry.isHierarchy };
        this.setFieldOptionsPanelOpen(true);
        await this.renderFieldOptionsBody(zoneId, entry);
    },

    async renderFieldOptionsBody(zoneId, entry) {
        const body = document.getElementById("fieldOptionsBody");
        if (!body) return;

        const isMeasure = String(entry.dimension).toUpperCase() === "MEASURE";
        const options = this.getFieldOptions(zoneId, entry);

        body.innerHTML = "";

        const nameEl = document.createElement("div");
        nameEl.className = "field-options-field-name";
        nameEl.innerText = isMeasure ? entry.name : `${entry.dimension}.${entry.name}`;
        body.appendChild(nameEl);

        if (isMeasure) {
            body.appendChild(await this.buildMeasureOptionsForm(zoneId, entry, options));
        } else if (entry.isHierarchy) {
            // Una jerarquía NO tiene "Mostrar totales" ni "Ordenar": solo
            // se puede elegir hasta qué nivel expandir y qué niveles se
            // muestran (subtotales/orden no aplican a nivel de jerarquía
            // completa, solo tendrían sentido por atributo individual).
            const levelsBox = document.createElement("div");
            levelsBox.className = "field-options-group";
            levelsBox.innerHTML = `<div class="field-options-section-title">Jerarquía</div>
                <div class="field-options-empty" id="fieldOptionsLevelsLoading">Cargando niveles…</div>`;
            body.appendChild(levelsBox);

            try {
                const levels = await window.ExcelService.getHierarchyLevels(entry.dimension, entry.name);
                this.renderHierarchyLevelControls(levelsBox, zoneId, entry, options, levels);
            } catch (err) {
                console.error("Error al leer los niveles de la jerarquía:", err);
                levelsBox.querySelector("#fieldOptionsLevelsLoading").innerText = "No se pudieron cargar los niveles.";
            }
        } else {
            body.appendChild(this.buildDimensionOptionsForm(zoneId, entry, options));
        }
    },

    /**
     * Cada "ID y texto" añade una columna (Filas) o fila (Columnas) física:
     * el rango del eje (RangeAxis) tiene que crecer/encoger igual que al
     * añadir/quitar un campo, o el texto se solapa con el otro eje.
     * reset=true: solo toma la cuenta actual como ya incluida en el rango
     * (al cargar un informe).
     */
    reconcileTextColumns(reset) {
        if (!window.ReportStore || !window.ReportStore.countTextColumns || !this.currentReportId) return;
        ["rows", "columns"].forEach(zone => {
            const n = window.ReportStore.countTextColumns(this.currentReportId, zone, this.listForZone(zone), this.state.fieldOptions);
            let delta = reset ? 0 : n - (this.textExtra[zone] || 0);
            const rowsFlag = this.state.colsStatic && !this.state.rowsStatic;
            const colsFlag = this.state.rowsStatic && !this.state.colsStatic;
            for (; delta > 0; delta--) {
                if (zone === "rows") RangeAxis.onRowFieldAdded(rowsFlag); else RangeAxis.onColFieldAdded(colsFlag);
            }
            for (; delta < 0; delta++) {
                if (zone === "rows") RangeAxis.onRowFieldRemoved(rowsFlag); else RangeAxis.onColFieldRemoved(colsFlag);
            }
            this.textExtra[zone] = n;
        });
        this.refreshRangeLabels();
    },

    /** Cambia "Solo ID" / "ID y texto" de un campo (panel de opciones o menú contextual). */
    setDisplayMode(zoneId, entry, mode) {
        const current = this.getFieldOptions(zoneId, entry);
        current.displayMode = (mode === "ID_TEXT" || mode === "TEXT") ? mode : "ID";
        this.setFieldOptions(zoneId, entry, current);
        this.reconcileTextColumns();
        const sel = this.selectedFieldForOptions;
        if (sel && sel.zoneId === zoneId && sel.dimension === entry.dimension && sel.name === entry.name) {
            this.renderFieldOptionsBody(zoneId, entry);
        }
        this.scheduleAutoUpdate();
    },

    /**
     * Menú de clic derecho sobre un campo de Filas/Columnas (mismas opciones
     * que "Opciones de campo"): Mostrar (ID / ID y descripción / Descripción),
     * Totales (atributos) o Jerarquía (niveles a mostrar y expandir hasta).
     */
    showDisplayModeMenu(x, y, zoneId, entry, info, levels) {
        this.hideDisplayModeMenu();

        const menu = document.createElement("div");
        menu.id = "displayModeMenu";
        menu.style.cssText = "position:fixed;z-index:100000;min-width:190px;max-height:80vh;overflow:auto;background:#fff;border:1px solid #cbd5e1;border-radius:6px;"
            + "box-shadow:0 4px 14px rgba(0,0,0,.18);padding:4px 0;font-size:12px;color:#0f172a;font-family:inherit;";

        const addTitle = (text) => {
            const t = document.createElement("div");
            t.textContent = text;
            t.style.cssText = "padding:5px 12px 3px;font-size:10px;font-weight:700;text-transform:uppercase;color:#64748b;";
            menu.appendChild(t);
        };
        const addSep = () => {
            const h = document.createElement("div");
            h.style.cssText = "height:1px;background:#e2e8f0;margin:4px 0;";
            menu.appendChild(h);
        };
        // onPick(); keepOpen=true deja el menú abierto (casillas de niveles)
        const addItem = (label, checked, onPick, opts) => {
            opts = opts || {};
            const item = document.createElement("div");
            const mark = () => (checked ? "\u2713 " : "\u00a0\u00a0\u00a0 ");
            item.textContent = mark() + label;
            item.style.cssText = "padding:5px 12px;cursor:" + (opts.disabled ? "not-allowed" : "pointer") + ";" + (opts.disabled ? "color:#94a3b8;" : "");
            if (opts.title) item.title = opts.title;
            if (!opts.disabled) {
                item.addEventListener("mouseenter", () => { item.style.background = "#eff6ff"; });
                item.addEventListener("mouseleave", () => { item.style.background = ""; });
                item.addEventListener("click", () => {
                    if (opts.keepOpen) {
                        checked = !checked;
                        item.textContent = mark() + label;
                    } else {
                        this.hideDisplayModeMenu();
                    }
                    onPick(checked);
                });
            }
            menu.appendChild(item);
        };

        // --- Mostrar (ID / ID y descripción / Descripción)
        if (info && info.applicable) {
            const saved = this.getFieldOptions(zoneId, entry).displayMode;
            const current = (info.available && (saved === "ID_TEXT" || saved === "TEXT")) ? saved : "ID";
            addTitle("Mostrar");
            [["ID", "ID"], ["ID_TEXT", "ID y descripción"], ["TEXT", "Descripción"]].forEach(([mode, label]) => {
                const disabled = mode !== "ID" && !info.available;
                addItem(label, current === mode, () => { if (mode !== current) this.setDisplayMode(zoneId, entry, mode); },
                    { disabled, title: disabled ? info.reason : "" });
            });
        }

        const patchOptions = (patch) => {
            const cur = Object.assign(this.getFieldOptions(zoneId, entry), patch);
            this.setFieldOptions(zoneId, entry, cur);
            const sel = this.selectedFieldForOptions;
            if (sel && sel.zoneId === zoneId && sel.dimension === entry.dimension && sel.name === entry.name) {
                this.renderFieldOptionsBody(zoneId, entry);
            }
            this.scheduleAutoUpdate();
        };

        if (entry.isHierarchy) {
            // --- Jerarquía: niveles a mostrar y expandir hasta nivel
            const opts = this.getFieldOptions(zoneId, entry);
            if (levels && levels.length) {
                if (menu.children.length) addSep();
                addTitle("Jerarquía \u203a Mostrar niveles");
                const visible = new Set(Array.isArray(opts.visibleLevels) ? opts.visibleLevels : levels.map(l => l.nivel));
                levels.forEach(l => {
                    addItem(`Nivel ${l.nivel} \u2014 ${l.attribute}`, visible.has(l.nivel), (checked) => {
                        if (checked) visible.add(l.nivel); else visible.delete(l.nivel);
                        patchOptions({ visibleLevels: levels.map(v => v.nivel).filter(n => visible.has(n)) });
                    }, { keepOpen: true });
                });
                addSep();
                addTitle("Jerarquía \u203a Expandir hasta nivel");
                const exp = opts.expandToLevel;
                addItem("(todos)", exp === null || exp === undefined || exp === "", () => patchOptions({ expandToLevel: null }));
                levels.forEach(l => {
                    addItem(`Nivel ${l.nivel} \u2014 ${l.attribute}`, String(exp) === String(l.nivel), () => patchOptions({ expandToLevel: Number(l.nivel) }));
                });
            } else {
                if (menu.children.length) addSep();
                addTitle("Jerarquía");
                addItem("No se pudieron cargar los niveles", false, () => {}, { disabled: true });
            }
        } else {
            // --- Totales (subtotales) mostrar / ocultar
            if (menu.children.length) addSep();
            addTitle("Totales");
            const showing = !!this.getFieldOptions(zoneId, entry).showTotals;
            addItem("Mostrar", showing, () => { if (!showing) patchOptions({ showTotals: true }); });
            addItem("Ocultar", !showing, () => { if (showing) patchOptions({ showTotals: false }); });
        }

        document.body.appendChild(menu);
        const r = menu.getBoundingClientRect();
        menu.style.left = Math.max(0, Math.min(x, window.innerWidth - r.width - 4)) + "px";
        menu.style.top = Math.max(0, Math.min(y, window.innerHeight - r.height - 4)) + "px";

        this._displayMenuCloser = (ev) => {
            if (ev.type === "keydown" && ev.key !== "Escape") return;
            if (ev.type === "mousedown" && menu.contains(ev.target)) return;
            this.hideDisplayModeMenu();
        };
        setTimeout(() => {
            document.addEventListener("mousedown", this._displayMenuCloser, true);
            document.addEventListener("keydown", this._displayMenuCloser, true);
            window.addEventListener("blur", this._displayMenuCloser, true);
            window.addEventListener("scroll", this._displayMenuCloser, true);
        }, 0);
    },

    hideDisplayModeMenu() {
        const m = document.getElementById("displayModeMenu");
        if (m) m.remove();
        if (this._displayMenuCloser) {
            document.removeEventListener("mousedown", this._displayMenuCloser, true);
            document.removeEventListener("keydown", this._displayMenuCloser, true);
            window.removeEventListener("blur", this._displayMenuCloser, true);
            window.removeEventListener("scroll", this._displayMenuCloser, true);
            this._displayMenuCloser = null;
        }
    },

    /**
     * "Mostrar como: Solo ID / ID y texto" del campo (clave de dimensión o
     * jerarquía). El atributo de texto se define en el modelo semántico; aquí
     * solo se elige el modo. Devuelve null si el campo no admite el modo.
     */
    buildDisplayModeBlock(zoneId, entry, options) {
        if (!window.ReportStore || !window.ReportStore.getDisplayTextInfo) return null;
        const info = window.ReportStore.getDisplayTextInfo(this.currentReportId, entry.dimension, entry.name, !!entry.isHierarchy);
        if (!info.applicable) return null;

        const label = document.createElement("label");
        label.className = "field-options-field-block";
        const span = document.createElement("span");
        span.className = "field-options-field-label";
        span.textContent = "Mostrar como";
        const select = document.createElement("select");
        select.id = "optDisplayMode";
        [["ID", "Solo ID"], ["ID_TEXT", "ID y texto"], ["TEXT", "Solo texto"]].forEach(([v, t]) => {
            const o = document.createElement("option");
            o.value = v;
            o.textContent = t;
            select.appendChild(o);
        });
        select.value = (info.available && (options.displayMode === "ID_TEXT" || options.displayMode === "TEXT")) ? options.displayMode : "ID";
        select.disabled = !info.available;
        if (!info.available) select.title = info.reason;
        select.addEventListener("change", () => this.setDisplayMode(zoneId, entry, select.value));
        label.appendChild(span);
        label.appendChild(select);
        if (!info.available) {
            const hint = document.createElement("div");
            hint.className = "field-options-empty";
            hint.textContent = info.reason;
            label.appendChild(hint);
        }
        return label;
    },

    buildDimensionOptionsForm(zoneId, entry, options) {
        const wrap = document.createElement("div");
        wrap.className = "field-options-group";

        // "Subtotales" (antes "Mostrar totales") y "Ordenar" solo aplican a
        // atributos que NO son jerarquía. Se guardan en EDIT_REPORT en las
        // columnas L/M (eje Filas) o R/S (eje Columnas) -ver
        // ExcelService.saveEditReportDesign-, aunque de momento no se usan
        // todavía para construir el SQL.
        wrap.innerHTML = `
            <div class="field-options-section-title">General</div>
            <label class="field-options-field-block">
                <span class="field-options-field-label">Subtotales</span>
                <select id="optShowTotals">
                    <option value="0" ${options.showTotals ? "" : "selected"}>Ocultar</option>
                    <option value="1" ${options.showTotals ? "selected" : ""}>Mostrar</option>
                </select>
            </label>
            <label class="field-options-field-block">
                <span class="field-options-field-label">Ordenar</span>
                <select id="optSortOrder">
                    <option value="none" ${options.sortOrder === "none" ? "selected" : ""}>Sin ordenar</option>
                    <option value="asc" ${options.sortOrder === "asc" ? "selected" : ""}>Ascendente (UP)</option>
                    <option value="desc" ${options.sortOrder === "desc" ? "selected" : ""}>Descendente (DOWN)</option>
                </select>
            </label>
        `;

        const persist = () => {
            const current = this.getFieldOptions(zoneId, entry);
            current.showTotals = wrap.querySelector("#optShowTotals").value === "1";
            current.sortOrder = wrap.querySelector("#optSortOrder").value;
            this.setFieldOptions(zoneId, entry, current);
            this.scheduleAutoUpdate();
        };

        wrap.querySelector("#optShowTotals").addEventListener("change", persist);
        wrap.querySelector("#optSortOrder").addEventListener("change", persist);

        const displayBlock = this.buildDisplayModeBlock(zoneId, entry, options);
        if (displayBlock) wrap.appendChild(displayBlock);

        return wrap;
    },

    renderHierarchyLevelControls(container, zoneId, entry, options, levels) {
        container.innerHTML = `<div class="field-options-section-title">Jerarquía</div>`;

        if (!levels || levels.length === 0) {
            const empty = document.createElement("div");
            empty.className = "field-options-empty";
            empty.innerText = "Esta jerarquía no tiene niveles definidos.";
            container.appendChild(empty);
            return;
        }

        // «Mostrar como» en una jerarquía solo si ya tiene textos por nivel
        // (el selector de texto por nivel está oculto de momento en el modelo)
        const hInfo = window.ReportStore && window.ReportStore.getDisplayTextInfo
            ? window.ReportStore.getDisplayTextInfo(this.currentReportId, entry.dimension, entry.name, true) : null;
        if (hInfo && hInfo.available) {
            const displayBlock = this.buildDisplayModeBlock(zoneId, entry, options);
            if (displayBlock) container.appendChild(displayBlock);
        }

        const expandLabel = document.createElement("label");
        expandLabel.className = "field-options-field-block";
        expandLabel.innerHTML = `<span class="field-options-field-label">Expandir hasta nivel</span>`;
        const expandSelect = document.createElement("select");
        expandSelect.id = "optExpandToLevel";
        // SEGURIDAD: l.attribute puede venir de un modelo semántico
        // importado (.lkml de un repo compartido) -> se construyen las
        // <option> con propiedades del DOM (que escapan su valor) en vez
        // de con un template literal HTML.
        const optAll = document.createElement("option");
        optAll.value = "";
        optAll.textContent = "(todos)";
        expandSelect.appendChild(optAll);
        levels.forEach(l => {
            const opt = document.createElement("option");
            opt.value = l.nivel;
            opt.title = `Nivel ${l.nivel} — ${l.attribute}`;
            opt.selected = String(options.expandToLevel) === String(l.nivel);
            opt.textContent = `Nivel ${l.nivel} — ${l.attribute}`;
            expandSelect.appendChild(opt);
        });
        expandLabel.appendChild(expandSelect);
        container.appendChild(expandLabel);

        const visibleLevels = Array.isArray(options.visibleLevels) ? options.visibleLevels : levels.map(l => l.nivel);

        const levelsTitle = document.createElement("div");
        levelsTitle.className = "field-options-section-title";
        levelsTitle.style.marginTop = "6px";
        levelsTitle.innerText = "Mostrar niveles";
        container.appendChild(levelsTitle);

        const levelsBox = document.createElement("div");
        levelsBox.className = "field-options-levels";
        levels.forEach(l => {
            const row = document.createElement("label");
            row.className = "field-options-checkbox-row field-options-level-row";
            row.title = `Nivel ${l.nivel} — ${l.attribute}`;
            // SEGURIDAD: l.attribute no es de confianza (ver arriba) -> nodos
            // construidos aparte en vez de innerHTML.
            const cb = document.createElement("input");
            cb.type = "checkbox";
            cb.dataset.nivel = l.nivel;
            cb.checked = visibleLevels.includes(l.nivel);
            const levelText = document.createElement("span");
            levelText.className = "field-options-level-text";
            levelText.textContent = `Nivel ${l.nivel} — ${l.attribute}`;
            row.appendChild(cb);
            row.appendChild(levelText);
            levelsBox.appendChild(row);
        });
        container.appendChild(levelsBox);

        const persist = () => {
            const current = this.getFieldOptions(zoneId, entry);
            const val = expandSelect.value;
            current.expandToLevel = val === "" ? null : Number(val);
            current.visibleLevels = Array.from(levelsBox.querySelectorAll("input[type=checkbox]"))
                .filter(cb => cb.checked)
                .map(cb => Number(cb.dataset.nivel));
            this.setFieldOptions(zoneId, entry, current);
            this.scheduleAutoUpdate();
        };

        expandSelect.addEventListener("change", persist);
        levelsBox.querySelectorAll("input[type=checkbox]").forEach(cb => cb.addEventListener("change", persist));
    },

    async buildMeasureOptionsForm(zoneId, entry, options) {
        const wrap = document.createElement("div");
        wrap.className = "field-options-group";

        // Formato: solo Número (con separador de miles) o Porcentaje, con
        // el nº de decimales de abajo (ver dracoNumberFormatFromOptions en
        // commands.js). El separador decimal lo pone Excel según el idioma
        // del usuario, y el factor de escala ya no existe. Los informes
        // antiguos traen numberFormat "#,##0.00"/"0.00%"/...: lo que lleva
        // "%" se muestra como Porcentaje y el resto como Número.
        const formatKind = (String(options.numberFormat || "").indexOf("%") !== -1 || options.numberFormat === "percent") ? "percent" : "number";
        // Agregaciones soportadas tal cual por el SQL dinámico (GROUPING
        // SETS) y por DAX. "" = la definida en el modelo semántico.
        const AGGREGATIONS = [
            ["", "Según el modelo"],
            ["SUM", "Suma (SUM)"],
            ["AVG", "Media (AVG)"],
            ["COUNT", "Recuento (COUNT)"],
            ["COUNT_DISTINCT", "Recuento distinto (COUNT DISTINCT)"],
            ["MIN", "Mínimo (MIN)"],
            ["MAX", "Máximo (MAX)"]
        ];
        const currentAgg = String(options.aggregation || "").toUpperCase();
        // Agregación de excepción (ver normalizeDracoExc en commands.js): "" = ninguna
        const EXC_AGGREGATIONS = [
            ["", "Ninguna"],
            ["MIN", "Mínimo (MIN)"],
            ["MAX", "Máximo (MAX)"],
            ["AVG", "Media (AVG)"],
            ["SUM", "Suma (SUM)"],
            ["COUNT", "Número de valores (COUNT)"]
        ];
        const currentExc = String(options.excAggregation || "").toUpperCase();
        const decimalsValue = Number.isFinite(Number(options.decimals)) ? Number(options.decimals) : 2;

        wrap.innerHTML = `
            <div class="field-options-section-title">Formato de número</div>
            <label>
                Formato
                <select id="optNumberFormat">
                    <option value="number" ${formatKind === "number" ? "selected" : ""}>Número (1.234)</option>
                    <option value="percent" ${formatKind === "percent" ? "selected" : ""}>Porcentaje (12%)</option>
                </select>
            </label>
            <label>
                Nº de decimales
                <input type="number" id="optDecimals" min="0" max="10" value="${decimalsValue}">
            </label>
            <label>
                Factor de escala
                <select id="optScale">
                    ${[1, 10, 100, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9].map(f =>
                        `<option value="${f}" ${Number(options.scale || 1) === f ? "selected" : ""}>${f === 1 ? "1 (sin escala)" : "÷ " + f.toLocaleString("es-ES")}</option>`).join("")}
                </select>
            </label>
            <div style="font-size:10px; color:#605e5c; margin:2px 0 6px;">
                Se aplica al actualizar si "Sobrescribir formatos" está activado
                en las propiedades del informe.
            </div>
            <div class="field-options-section-title">Agregación</div>
            <label>
                Función
                <select id="optAggregation">
                    ${AGGREGATIONS.map(([value, label]) =>
                        `<option value="${value}" ${currentAgg === value ? "selected" : ""}>${label}</option>`).join("")}
                </select>
            </label>
            <label>
                Agregación de excepción
                <select id="optExcAggregation">
                    ${EXC_AGGREGATIONS.map(([value, label]) =>
                        `<option value="${value}" ${currentExc === value ? "selected" : ""}>${label}</option>`).join("")}
                </select>
            </label>
            <div id="optExcSection" style="display:${currentExc ? "block" : "none"};">
                <div style="font-size:10px; color:#605e5c; margin:2px 0 6px;">
                    Primero se agrega la medida (con la función de arriba) por cada valor de las dimensiones
                    de referencia —por todas las demás, como siempre— y después se aplica la agregación de
                    excepción sobre esos valores. Ej.: mínimo mensual = referencia «mes» + Mínimo.
                </div>
                <div style="font-size:11px; font-weight:600; margin-bottom:4px;">Dimensiones de referencia</div>
                <div id="optExcDims" style="display:flex; flex-direction:column; gap:3px; max-height:180px; overflow-y:auto;"></div>
            </div>
            <div class="field-options-section-title">Planificación</div>
            <label class="field-options-field-block">
                <span class="field-options-field-label">Planificable</span>
                <select id="optPlanificable"><option value="0" ${options.planificable ? "" : "selected"}>No</option><option value="1" ${options.planificable ? "selected" : ""}>Sí</option></select>
            </label>
            <div id="optPlanificableSection" style="display:${options.planificable ? "block" : "none"}; margin-top:6px;">
                <label class="field-options-field-block">
                    <span class="field-options-field-label">Criterio de dimensiones</span>
                    <select id="optDimensionsBehaviorCustom">
                        <option value="0" ${options.dimensionsBehaviorCustom ? "" : "selected"}>El del informe</option>
                        <option value="1" ${options.dimensionsBehaviorCustom ? "selected" : ""}>Personalizado para esta medida</option>
                    </select>
                </label>
                <div style="font-size:10px; color:#605e5c; margin:4px 0 6px;">
                    Si no se marca, usa el criterio global del informe
                    (Propiedades del informe).
                </div>
                <div id="optDimensionsBehaviorContainer" style="display:${options.dimensionsBehaviorCustom ? "flex" : "none"}; flex-direction:column; gap:6px; max-height:180px; overflow-y:auto;"></div>
            </div>
        `;

        const persist = () => {
            const current = this.getFieldOptions(zoneId, entry);
            current.numberFormat = wrap.querySelector("#optNumberFormat").value;
            const dec = Math.floor(Number(wrap.querySelector("#optDecimals").value));
            current.decimals = Number.isFinite(dec) ? Math.min(10, Math.max(0, dec)) : 0;
            // Factor de escala: el valor se divide en la consulta (1 = sin escala)
            const sc = Number(wrap.querySelector("#optScale").value);
            if (sc > 1) current.scale = sc; else delete current.scale;
            current.aggregation = wrap.querySelector("#optAggregation").value;
            // Agregación de excepción: sin función (o sin dimensiones) no se guarda nada y el SQL no cambia
            const excAgg = wrap.querySelector("#optExcAggregation").value;
            const excDims = [];
            wrap.querySelectorAll("#optExcDims input[type=checkbox]:checked").forEach(cb => {
                excDims.push({ dim: cb.getAttribute("data-dim"), attr: cb.getAttribute("data-attr") });
            });
            if (excAgg) { current.excAggregation = excAgg; current.excDims = excDims; }
            else { delete current.excAggregation; delete current.excDims; }
            // Opciones retiradas del panel (ya no se usan en ningún sitio).
            delete current.decimalSeparator;
            delete current.thousandsSeparator;
            delete current.factor;
            current.planificable = wrap.querySelector("#optPlanificable").value === "1";
            current.dimensionsBehaviorCustom = wrap.querySelector("#optDimensionsBehaviorCustom").value === "1";

            const dimensionsBehavior = {};
            wrap.querySelectorAll("#optDimensionsBehaviorContainer .dim-behavior-row").forEach(row => {
                const dim = row.querySelector(".dim-behavior-required").getAttribute("data-dim");
                dimensionsBehavior[dim] = {
                    required: row.querySelector(".dim-behavior-required").checked,
                    mode: row.querySelector(".dim-behavior-mode").value
                };
            });
            current.dimensionsBehavior = dimensionsBehavior;

            this.setFieldOptions(zoneId, entry, current);
            this.scheduleAutoUpdate();
        };

        wrap.querySelectorAll("select, input").forEach(el => el.addEventListener("change", persist));

        // Agregación de excepción: lista de atributos del modelo (se rellena al elegir una función)
        const excBox = wrap.querySelector("#optExcDims");
        excBox.addEventListener("change", persist);
        const fillExc = async () => {
            if (excBox.getAttribute("data-filled")) return;
            excBox.setAttribute("data-filled", "1");
            await this.fillExcDimsContainer(excBox, options.excDims || []);
        };
        wrap.querySelector("#optExcAggregation").addEventListener("change", async (e) => {
            wrap.querySelector("#optExcSection").style.display = e.target.value ? "block" : "none";
            if (e.target.value) await fillExc();
        });
        if (currentExc) await fillExc();

        // "Planificable" despliega/oculta la sección de personalización.
        wrap.querySelector("#optPlanificable").addEventListener("change", (e) => {
            wrap.querySelector("#optPlanificableSection").style.display = e.target.value === "1" ? "block" : "none";
        });

        // "Personalizar..." despliega/oculta y rellena la lista de
        // dimensiones (mismo bloque reutilizable que Propiedades del
        // informe: buildDimensionBehaviorRow).
        const dimsContainer = wrap.querySelector("#optDimensionsBehaviorContainer");
        wrap.querySelector("#optDimensionsBehaviorCustom").addEventListener("change", async (e) => {
            dimsContainer.style.display = e.target.value === "1" ? "flex" : "none";
            if (e.target.value === "1") await this.fillDimensionsBehaviorContainer(dimsContainer, options.dimensionsBehavior || {});
        });

        if (options.dimensionsBehaviorCustom) {
            await this.fillDimensionsBehaviorContainer(dimsContainer, options.dimensionsBehavior || {});
        }

        return wrap;
    },

    // Agregación de excepción: una casilla por cada atributo de las dimensiones del modelo del informe
    async fillExcDimsContainer(container, saved) {
        container.innerHTML = "<div style='font-size:11px; color:#605e5c;'>Cargando dimensiones…</div>";
        try {
            const report = (this.currentReportId && window.ReportStore) ? window.ReportStore.getReport(this.currentReportId) : null;
            const modelName = report ? report.semanticModelName : "";
            if (!modelName || !window.SemanticModelStore) {
                container.innerHTML = "<div style='font-size:11px; color:#605e5c;'>Este informe no tiene modelo semántico asignado.</div>";
                return;
            }
            const grid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES", modelName);
            const values = grid && grid.values ? grid.values : [];
            const items = [], seen = new Set();
            for (let r = 1; r < values.length; r++) {
                const dim = values[r] && values[r][1], attr = values[r] && values[r][2];
                if (!dim || !attr || seen.has(dim + "|" + attr)) continue;
                seen.add(dim + "|" + attr);
                items.push({ dim: String(dim), attr: String(attr), label: String(values[r][7] || attr) });
            }
            if (!items.length) {
                container.innerHTML = "<div style='font-size:11px; color:#605e5c;'>Este modelo no tiene dimensiones.</div>";
                return;
            }
            const isOn = (it) => (saved || []).some(d => String(d.dim).toUpperCase() === it.dim.toUpperCase() && String(d.attr).toUpperCase() === it.attr.toUpperCase());
            const esc = (t) => String(t).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" }[c]));
            container.innerHTML = items.map(it =>
                `<label class="field-options-checkbox-row"><input type="checkbox" data-dim="${esc(it.dim)}" data-attr="${esc(it.attr)}" ${isOn(it) ? "checked" : ""}><span>${esc(it.dim.toLowerCase())} · ${esc(it.label)}</span></label>`).join("");
        } catch (err) {
            console.warn("No se pudieron cargar las dimensiones para la agregación de excepción:", err);
            container.innerHTML = "<div style='font-size:11px; color:#a80000;'>No se pudieron cargar las dimensiones.</div>";
        }
    },

    // Rellena un contenedor con una fila "Obligatoria + criterio" por cada
    // dimensión del modelo del informe actual — usado tanto aquí (Opciones
    // de campo de una medida) como en Propiedades del informe.
    async fillDimensionsBehaviorContainer(container, saved) {
        container.innerHTML = "<div style='font-size:11px; color:#605e5c;'>Cargando dimensiones…</div>";
        try {
            const report = (this.currentReportId && window.ReportStore) ? window.ReportStore.getReport(this.currentReportId) : null;
            const modelName = report ? report.semanticModelName : "";
            if (!modelName || !window.SemanticModelStore) {
                container.innerHTML = "<div style='font-size:11px; color:#605e5c;'>Este informe no tiene modelo semántico asignado.</div>";
                return;
            }

            const grid = await window.SemanticModelStore.getModelGrid("MODEL_DIMENSION", modelName);
            const dims = [];
            const seen = new Set();
            const values = grid && grid.values ? grid.values : [];
            for (let r = 1; r < values.length; r++) {
                const dim = values[r] && values[r][1];
                if (dim && !seen.has(dim)) { seen.add(dim); dims.push(dim); }
            }

            if (dims.length === 0) {
                container.innerHTML = "<div style='font-size:11px; color:#605e5c;'>Este modelo no tiene dimensiones.</div>";
                return;
            }

            container.innerHTML = "";
            dims.forEach(dim => {
                const cfg = saved[dim] || { required: true, mode: "null" };
                container.appendChild(this.buildDimensionBehaviorRow(dim, cfg));
            });
        } catch (err) {
            console.warn("No se pudieron cargar las dimensiones para 'Comportamiento de dimensiones':", err);
            container.innerHTML = "<div style='font-size:11px; color:#a80000;'>No se pudieron cargar las dimensiones.</div>";
        }
    }
};

// SEGURIDAD (auditoría F-14): todo texto externo (nombres de fichero del
// repositorio, mensajes de error) se escapa antes de meterlo en innerHTML.
function draco_escapeHtml(str) {
    return String(str == null ? "" : str)
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
