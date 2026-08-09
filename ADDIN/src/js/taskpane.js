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
        colsStatic: false
    },

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
            await window.ExcelService.saveEditReportDesign({
                filters: this.state.filters,
                rows: this.state.rows,
                columns: this.state.columns,
                rowsStatic: this.state.rowsStatic,
                colsStatic: this.state.colsStatic,
                rrAddress: RangeAxis.addressOf("rr"),
                rcAddress: RangeAxis.addressOf("rc")
            });

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

    async init() {
        if (typeof FilterModal !== "undefined" && FilterModal.init) {
            FilterModal.init();
        }
        this.bindEvents();
        await this.loadFields();
        await this.loadDesignFromSheet();
    },

    bindEvents() {
        const btnRefresh = document.getElementById("btnRefresh");
        if (btnRefresh) btnRefresh.addEventListener("click", () => this.loadFields());

        const fieldSearch = document.getElementById("fieldSearch");
        if (fieldSearch) fieldSearch.addEventListener("input", (e) => this.filterFields(e.target.value));

        const btnSave = document.getElementById("btnSaveDesign");
        if (btnSave) btnSave.addEventListener("click", () => this.saveDesign());

        // Checkboxes Estático / Dinámico (Checkrow / CheckCol del VBA)
        const chkRows = document.getElementById("chkAsymmetricRows");
        const chkCols = document.getElementById("chkAsymmetricCols");

        if (chkRows) {
            chkRows.addEventListener("change", (e) => {
                this.state.rowsStatic = e.target.checked;
                this.updateStaticLabel("rows");
                this.scheduleAutoUpdate();
            });
        }
        if (chkCols) {
            chkCols.addEventListener("change", (e) => {
                this.state.colsStatic = e.target.checked;
                this.updateStaticLabel("cols");
                this.scheduleAutoUpdate();
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

            RangeAxis.loadFromAddresses(design.rrAddress, design.rcAddress);

            // Pintar checkboxes
            const chkRows = document.getElementById("chkAsymmetricRows");
            const chkCols = document.getElementById("chkAsymmetricCols");
            if (chkRows) chkRows.checked = this.state.rowsStatic;
            if (chkCols) chkCols.checked = this.state.colsStatic;
            this.updateStaticLabel("rows");
            this.updateStaticLabel("cols");

            this.refreshRangeLabels();

            // Pintar tags existentes en cada zona
            const filtersContent = document.querySelector('.zone-card[data-zone="filters"] .dropzone-content');
            const rowsContent = document.querySelector('.zone-card[data-zone="rows"] .dropzone-content');
            const colsContent = document.querySelector('.zone-card[data-zone="columns"] .dropzone-content');

            this.state.filters.forEach(f => this.renderTag(filtersContent, "filters", f));
            this.state.rows.forEach(r => this.renderTag(rowsContent, "rows", r));
            this.state.columns.forEach(c => this.renderTag(colsContent, "columns", c));

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

            await window.ExcelService.saveEditReportDesign({
                filters: this.state.filters,
                rows: this.state.rows,
                columns: this.state.columns,
                rowsStatic: this.state.rowsStatic,
                colsStatic: this.state.colsStatic,
                rrAddress: RangeAxis.addressOf("rr"),
                rcAddress: RangeAxis.addressOf("rc")
            });

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
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        e.currentTarget.classList.add("drag-over");
    },

    handleDragLeave(e) {
        e.currentTarget.classList.remove("drag-over");
    },

    handleDrop(e) {
        e.preventDefault();
        const dropzoneBox = e.currentTarget;
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
        const zoneId = dropzoneBox.getAttribute("data-zone");

        if (sourceZone === zoneId) {
            // Mismo eje: no hacer nada
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

        container.appendChild(tag);
    }
};
