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

    loadFromAddresses(rrAddress, rcAddress) {
        this.rr = this._parseRange(rrAddress, { row: 1, col: 1, height: 1, width: 0 });
        this.rc = this._parseRange(rcAddress, { row: 1, col: 1, height: 0, width: 1 });
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

    // Al añadir un campo al eje Filas: crece en anchura y empuja columnas
    onRowFieldAdded() {
        this.rr.width += 1;
        this.moveRC(0, 1);
    },

    // Al quitar un campo del eje Filas
    onRowFieldRemoved() {
        this.rr.width = Math.max(0, this.rr.width - 1);
        this.moveRC(0, -1);
    },

    // Al añadir un campo al eje Columnas: crece en altura y empuja filas
    onColFieldAdded() {
        this.rc.height += 1;
        this.moveRR(1, 0);
    },

    // Al quitar un campo del eje Columnas
    onColFieldRemoved() {
        this.rc.height = Math.max(0, this.rc.height - 1);
        this.moveRR(-1, 0);
    }
};

const TaskPaneApp = {
    draggedElementData: null,

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
        subtotalsOnTop: true,
        overwriteFormats: true,
        autoFitColumns: false
    },

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

    scheduleAutoUpdate() {
        this.setAutoStatus("Cambios pendientes…");
        if (this.autoRefreshTimer) clearTimeout(this.autoRefreshTimer);
        this.autoRefreshTimer = setTimeout(() => this.runAutoSaveAndRefresh(), 700);
    },

    async runAutoSaveAndRefresh() {
        if (this.isAutoRefreshing) {
            this.autoRefreshQueued = true;
            return;
        }
        this.isAutoRefreshing = true;

        try {
            this.setAutoStatus("Guardando…");
            await window.ExcelService.saveEditReportDesign(this.collectDesignPayload());

            if (window.ReportActions && typeof window.ReportActions.actualizar === "function") {
                this.setAutoStatus("Actualizando…");
                await window.ReportActions.actualizar();
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
                this.scheduleAutoUpdate();
            }
        }
    },

    // Guarda el diseño SIN disparar Actualizar() (BigQuery). Se usa al
    // marcar/desmarcar Estático: el propio botón pide explícitamente que
    // eso no dispare un refresco automático.
    async saveDesignOnly() {
        try {
            this.setAutoStatus("Guardando…");
            await window.ExcelService.saveEditReportDesign(this.collectDesignPayload());
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

    async init() {
        if (typeof FilterModal !== "undefined" && FilterModal.init) {
            FilterModal.init();
        }
        this.bindEvents();
        this.registerPendingRibbonActionListener();
        this.loadReportPropertiesFromSettings();
        await this.loadFields();
        await this.loadDesignFromSheet();
        await this.handlePendingRibbonAction();
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

        const btnSave = document.getElementById("btnSaveDesign");
        if (btnSave) btnSave.addEventListener("click", () => this.saveDesign());

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

        // Botón "Opciones de campo" (toggle del panel derecho)
        const btnFieldOptions = document.getElementById("btnFieldOptions");
        if (btnFieldOptions) btnFieldOptions.addEventListener("click", () => this.toggleFieldOptionsPanel());
    },

    /* -------------------------------------------------------------
     * Eje Estático: sombrea la zona y desactiva el drag&drop (tanto
     * soltar campos nuevos como arrastrar/quitar los que ya hay).
     * ----------------------------------------------------------- */
    setZoneLocked(zoneId, locked) {
        const zone = document.querySelector(`.zone-card[data-zone="${zoneId}"]`);
        if (!zone) return;
        zone.classList.toggle("zone-locked", locked);
        zone.querySelectorAll(".dropped-tag").forEach(tag => { tag.draggable = !locked; });
    },

    isZoneLocked(zoneId) {
        const zone = document.querySelector(`.zone-card[data-zone="${zoneId}"]`);
        return !!(zone && zone.classList.contains("zone-locked"));
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
        try {
            const design = await window.ExcelService.loadEditReportDesign();

            this.state.filters = design.filters.map(f => ({
                dimension: f.dimension,
                name: f.name,
                isHierarchy: f.isHierarchy,
                realAttribute: f.realAttribute,
                value: f.value
            }));
            this.state.rows = design.rows.map(r => ({ dimension: r.dimension, name: r.name, isHierarchy: r.isHierarchy }));
            this.state.columns = design.columns.map(c => ({ dimension: c.dimension, name: c.name, isHierarchy: c.isHierarchy }));
            this.state.rowsStatic = design.rowsStatic;
            this.state.colsStatic = design.colsStatic;
            this.state.fieldOptions = design.fieldOptions || {};

            RangeAxis.loadFromAddresses(design.rrAddress, design.rcAddress);

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

            this.state.filters.forEach(f => this.renderTag(filtersContent, "filters", f));
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
     * Guardado del diseño en EDIT_REPORT (botón "Guardar")
     * ----------------------------------------------------------- */
    async saveDesign() {
        const btn = document.getElementById("btnSaveDesign");
        const incomplete = this.state.filters.filter(f => !f.value);

        if (incomplete.length > 0) {
            const proceed = confirm(
                `Hay ${incomplete.length} filtro(s) sin un valor seleccionado (doble clic sobre el filtro para elegirlo). ¿Guardar igualmente?`
            );
            if (!proceed) return;
        }

        try {
            if (btn) { btn.disabled = true; btn.innerText = "Guardando…"; }

            await window.ExcelService.saveEditReportDesign(this.collectDesignPayload());

            if (btn) btn.innerText = "Guardado ✓";
        } catch (err) {
            console.error("Error guardando el diseño en EDIT_REPORT:", err);
            alert("Error al guardar: " + (err.message || err));
            if (btn) btn.innerText = "Guardar";
        } finally {
            if (btn) {
                setTimeout(() => { btn.disabled = false; btn.innerText = "Guardar"; }, 1500);
            }
        }
    },

    /* -------------------------------------------------------------
     * Lista de campos disponibles (izquierda)
     * ----------------------------------------------------------- */
    async loadFields() {
        const container = document.getElementById("availableFieldsContainer");
        if (!container) return;

        container.innerHTML = "<div style='color: #605e5c; padding: 4px;'>Cargando dimensiones...</div>";

        try {
            const result = await window.ExcelService.readDim2Data();

            if (result.error) {
                container.innerHTML = `<div style='color: #a80000; padding: 4px;'>⚠️ ${result.error}</div>`;
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
                header.innerText = dim.dimension.toLowerCase();
                group.appendChild(header);

                dim.hierarchies.forEach(hier => {
                    group.appendChild(this.createFieldElement(dim.dimension, hier, true));
                });

                dim.attributes.forEach(att => {
                    group.appendChild(this.createFieldElement(dim.dimension, att, false));
                });

                container.appendChild(group);
            });
        } catch (err) {
            console.error("Error al cargar dimensiones:", err);
            container.innerHTML = `<div style='color: #a80000; padding: 4px;'>❌ Error: ${err.message || err}</div>`;
        }
    },

    createFieldElement(dim, name, isHierarchy) {
        const div = document.createElement("div");
        div.className = "field-item";
        div.draggable = true;

        const iconClass = isHierarchy ? "field-icon hierarchy-icon" : "field-icon";
        const iconSymbol = isHierarchy ? "🗂️" : "📄";

        div.innerHTML = `
            <span class="${iconClass}">${iconSymbol}</span>
            <span class="field-label">${name.toLowerCase()}</span>
        `;

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
     * Drag & drop entre zonas
     * ----------------------------------------------------------- */
    handleDragOver(e) {
        const zoneId = e.currentTarget.getAttribute("data-zone");
        if (this.isZoneLocked(zoneId)) return; // eje Estático: no admite drop
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        e.currentTarget.classList.add("drag-over");
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

        const { data, sourceTag, sourceZone } = this.draggedElementData;
        const targetContent = dropzoneBox.querySelector(".dropzone-content");

        if (sourceZone === zoneId) {
            // Mismo eje: no hacer nada
            this.draggedElementData = null;
            return;
        }

        // El eje de origen también puede estar bloqueado (no se puede sacar
        // un campo de un eje Estático arrastrándolo a otro sitio).
        if (sourceZone && this.isZoneLocked(sourceZone)) {
            this.draggedElementData = null;
            return;
        }

        if (sourceTag) {
            sourceTag.remove();
            this.removeFromState(sourceZone, data);
        }

        this.addField(targetContent, zoneId, data);

        this.draggedElementData = null;
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
            entry.value = "";
        }

        list.push(entry);

        if (zoneId === "rows") RangeAxis.onRowFieldAdded();
        if (zoneId === "columns") RangeAxis.onColFieldAdded();
        this.refreshRangeLabels();

        this.renderTag(container, zoneId, entry);

        // Un filtro recién soltado aún no tiene valor (se ignora en el WHERE
        // hasta que se elija uno), pero añadir/quitar campos de Filas o
        // Columnas sí cambia el informe de inmediato: autoguardar+actualizar.
        this.scheduleAutoUpdate();
    },

    removeFromState(zoneId, data) {
        const list = this.listForZone(zoneId);
        const idx = list.findIndex(x => x.dimension === data.dim && x.name === data.name);
        if (idx !== -1) list.splice(idx, 1);

        if (zoneId === "rows") RangeAxis.onRowFieldRemoved();
        if (zoneId === "columns") RangeAxis.onColFieldRemoved();
        this.refreshRangeLabels();

        this.scheduleAutoUpdate();
    },

    listForZone(zoneId) {
        if (zoneId === "filters") return this.state.filters;
        if (zoneId === "rows") return this.state.rows;
        if (zoneId === "columns") return this.state.columns;
        return [];
    },

    renderTag(container, zoneId, entry) {
        const tag = document.createElement("div");
        tag.className = "dropped-tag";
        tag.draggable = true;
        tag.dataset.dim = entry.dimension;
        tag.dataset.fieldName = entry.name;
        tag.dataset.isHierarchy = entry.isHierarchy;

        // Filtro sin valor seleccionado todavía: se muestra vacío (no se
        // añade al WHERE de la consulta hasta que el usuario elija un valor
        // con doble clic).
        const titleText = zoneId === "filters"
            ? (entry.value ? `${entry.dimension}.${entry.name}: ${entry.value}` : `${entry.dimension}.${entry.name}: (vacío · doble clic para elegir)`)
            : `${entry.dimension}.${entry.name}`;

        tag.innerHTML = `
            <span class="dropped-tag-title">${titleText}</span>
            <span class="dropped-tag-remove">&times;</span>
        `;

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

        // Solo los filtros abren el modal de selección de valor (doble clic)
        if (zoneId === "filters") {
            tag.addEventListener("dblclick", async () => {
                if (typeof FilterModal === "undefined" || !FilterModal.open) return;

                const result = await FilterModal.open({ dim: entry.dimension, name: entry.name, isHierarchy: entry.isHierarchy });
                if (result) {
                    entry.value = result.value;
                    entry.realAttribute = result.attribute;
                    tag.querySelector(".dropped-tag-title").innerText = `${entry.dimension}.${entry.name}: ${entry.value}`;
                    this.scheduleAutoUpdate();
                }
            });
        }

        tag.querySelector(".dropped-tag-remove").addEventListener("click", (e) => {
            e.stopPropagation();
            tag.remove();
            this.removeFromState(zoneId, { dim: entry.dimension, name: entry.name });
        });

        // Clic (simple) sobre el campo: abre/actualiza el panel "Opciones de
        // campo" con los controles de ESE campo (medida o dimensión).
        tag.classList.add("clickable-for-options");
        tag.addEventListener("click", (e) => {
            if (e.target.closest(".dropped-tag-remove")) return;
            this.selectFieldForOptions(zoneId, entry);
        });

        container.appendChild(tag);
    },

    /* =================================================================
     * PROPIEDADES DEL INFORME (modal): nombre del informe, suprimir
     * ceros en filas/columnas, subtotales arriba, sobrescribir formatos,
     * autoajustar columnas. Se guardan en Office roaming settings, para
     * que también las pueda leer commands.js (ribbon y jsonTo3Matrices).
     * ================================================================= */
    loadReportPropertiesFromSettings() {
        try {
            const raw = Office.context.document.settings.get("draco_reportProperties");
            if (raw) {
                this.reportProperties = Object.assign({}, this.reportProperties, JSON.parse(raw));
            }
        } catch (err) {
            console.warn("No se pudieron leer las propiedades del informe:", err);
        }
    },

    async saveReportPropertiesToSettings() {
        const settings = Office.context.document.settings;
        settings.set("draco_reportProperties", JSON.stringify(this.reportProperties));
        await new Promise((resolve) => settings.saveAsync(resolve));
    },

    openReportPropertiesModal() {
        const modal = document.getElementById("reportPropertiesModal");
        if (!modal) return;

        document.getElementById("propReportName").value = this.reportProperties.reportName || "Report 001";
        document.getElementById("propSuppressZeroRows").checked = !!this.reportProperties.suppressZeroRows;
        document.getElementById("propSuppressZeroCols").checked = !!this.reportProperties.suppressZeroCols;
        document.getElementById("propSubtotalsOnTop").checked = !!this.reportProperties.subtotalsOnTop;
        document.getElementById("propOverwriteFormats").checked = !!this.reportProperties.overwriteFormats;
        document.getElementById("propAutoFitColumns").checked = !!this.reportProperties.autoFitColumns;

        modal.style.display = "flex";
    },

    closeReportPropertiesModal() {
        const modal = document.getElementById("reportPropertiesModal");
        if (modal) modal.style.display = "none";
    },

    async saveReportPropertiesFromModal() {
        this.reportProperties = {
            reportName: (document.getElementById("propReportName").value || "Report 001").trim(),
            suppressZeroRows: document.getElementById("propSuppressZeroRows").checked,
            suppressZeroCols: document.getElementById("propSuppressZeroCols").checked,
            subtotalsOnTop: document.getElementById("propSubtotalsOnTop").checked,
            overwriteFormats: document.getElementById("propOverwriteFormats").checked,
            autoFitColumns: document.getElementById("propAutoFitColumns").checked
        };

        const btn = document.getElementById("btnSaveProperties");
        try {
            if (btn) { btn.disabled = true; btn.innerText = "Guardando…"; }
            await this.saveReportPropertiesToSettings();
            this.closeReportPropertiesModal();
            // Estas propiedades las lee jsonTo3Matrices en cada refresco real;
            // no hace falta lanzar uno aquí, solo quedan guardadas para el próximo.
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
        if (btn) btn.classList.toggle("toggle-active", open);
        if (!open) this.selectedFieldForOptions = null;
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
            ? { numberFormat: "#,##0.00", decimalSeparator: ",", thousandsSeparator: ".", factor: 1, decimals: 2, aggregation: "SUM" }
            : { showTotals: true, sortOrder: "none", expandToLevel: null, visibleLevels: null };
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
        nameEl.innerText = `${entry.dimension}.${entry.name}`;
        body.appendChild(nameEl);

        if (isMeasure) {
            body.appendChild(this.buildMeasureOptionsForm(zoneId, entry, options));
        } else {
            body.appendChild(this.buildDimensionOptionsForm(zoneId, entry, options));
            if (entry.isHierarchy) {
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
            }
        }
    },

    buildDimensionOptionsForm(zoneId, entry, options) {
        const wrap = document.createElement("div");
        wrap.className = "field-options-group";

        wrap.innerHTML = `
            <div class="field-options-section-title">General</div>
            <label class="field-options-checkbox-row">
                <input type="checkbox" id="optShowTotals" ${options.showTotals ? "checked" : ""}>
                Mostrar totales
            </label>
            <label>
                Ordenar
                <select id="optSortOrder">
                    <option value="none" ${options.sortOrder === "none" ? "selected" : ""}>Sin ordenar</option>
                    <option value="asc" ${options.sortOrder === "asc" ? "selected" : ""}>Ascendente</option>
                    <option value="desc" ${options.sortOrder === "desc" ? "selected" : ""}>Descendente</option>
                </select>
            </label>
        `;

        const persist = () => {
            const current = this.getFieldOptions(zoneId, entry);
            current.showTotals = wrap.querySelector("#optShowTotals").checked;
            current.sortOrder = wrap.querySelector("#optSortOrder").value;
            this.setFieldOptions(zoneId, entry, current);
            this.scheduleAutoUpdate();
        };

        wrap.querySelector("#optShowTotals").addEventListener("change", persist);
        wrap.querySelector("#optSortOrder").addEventListener("change", persist);

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

        const expandLabel = document.createElement("label");
        expandLabel.innerHTML = `Expandir hasta nivel`;
        const expandSelect = document.createElement("select");
        expandSelect.id = "optExpandToLevel";
        expandSelect.innerHTML = `<option value="">(todos)</option>` +
            levels.map(l => `<option value="${l.nivel}" ${String(options.expandToLevel) === String(l.nivel) ? "selected" : ""}>Nivel ${l.nivel} — ${l.attribute}</option>`).join("");
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
            row.className = "field-options-checkbox-row";
            row.innerHTML = `<input type="checkbox" data-nivel="${l.nivel}" ${visibleLevels.includes(l.nivel) ? "checked" : ""}> Nivel ${l.nivel} — ${l.attribute}`;
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

    buildMeasureOptionsForm(zoneId, entry, options) {
        const wrap = document.createElement("div");
        wrap.className = "field-options-group";

        wrap.innerHTML = `
            <div class="field-options-section-title">Formato de número</div>
            <label>
                Formato
                <select id="optNumberFormat">
                    <option value="#,##0" ${options.numberFormat === "#,##0" ? "selected" : ""}>1.234 (sin decimales)</option>
                    <option value="#,##0.00" ${options.numberFormat === "#,##0.00" ? "selected" : ""}>1.234,00</option>
                    <option value="0.00%" ${options.numberFormat === "0.00%" ? "selected" : ""}>Porcentaje</option>
                    <option value="#,##0.00 €" ${options.numberFormat === "#,##0.00 €" ? "selected" : ""}>Moneda (€)</option>
                </select>
            </label>
            <label>
                Separador decimal
                <select id="optDecimalSeparator">
                    <option value="," ${options.decimalSeparator === "," ? "selected" : ""}>Coma (,)</option>
                    <option value="." ${options.decimalSeparator === "." ? "selected" : ""}>Punto (.)</option>
                </select>
            </label>
            <label>
                Nº de decimales
                <input type="number" id="optDecimals" min="0" max="10" value="${options.decimals}">
            </label>
            <label>
                Factor de escala
                <select id="optFactor">
                    <option value="1" ${Number(options.factor) === 1 ? "selected" : ""}>Unidades</option>
                    <option value="1000" ${Number(options.factor) === 1000 ? "selected" : ""}>Miles</option>
                    <option value="1000000" ${Number(options.factor) === 1000000 ? "selected" : ""}>Millones</option>
                </select>
            </label>
            <div class="field-options-section-title">Agregación</div>
            <label>
                Función
                <select id="optAggregation">
                    ${["SUM", "AVG", "COUNT", "COUNT_DISTINCT", "MIN", "MAX"].map(a =>
                        `<option value="${a}" ${options.aggregation === a ? "selected" : ""}>${a}</option>`).join("")}
                </select>
            </label>
        `;

        const persist = () => {
            const current = this.getFieldOptions(zoneId, entry);
            current.numberFormat = wrap.querySelector("#optNumberFormat").value;
            current.decimalSeparator = wrap.querySelector("#optDecimalSeparator").value;
            current.decimals = Number(wrap.querySelector("#optDecimals").value) || 0;
            current.factor = Number(wrap.querySelector("#optFactor").value) || 1;
            current.aggregation = wrap.querySelector("#optAggregation").value;
            this.setFieldOptions(zoneId, entry, current);
            this.scheduleAutoUpdate();
        };

        wrap.querySelectorAll("select, input").forEach(el => el.addEventListener("change", persist));

        return wrap;
    }
};
