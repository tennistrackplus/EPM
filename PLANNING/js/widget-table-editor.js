/**
 * Editor tipo Excel para widgets de tipo TABLA.
 *
 * Estado del widget (se serializa completo en WIDGETS.CONFIG_JSON):
 *   {
 *     rows, cols,                     // dimensiones de la rejilla
 *     cells: { "r_c": CellStyle },    // solo celdas con contenido/formato != por defecto
 *     merges: [{ r, c, rowSpan, colSpan }]   // r,c = celda ancla (arriba-izq.)
 *   }
 *
 * CellStyle: { v, b, i, u, al, ff, fs, col, bg, bt, br, bb, bl }
 *   v=valor, b/i/u=negrita/cursiva/subrayado (1/0), al=alineación,
 *   ff=familia de fuente, fs=tamaño (px), col=color texto, bg=color fondo,
 *   bt/br/bb/bl = borde arriba/derecha/abajo/izquierda (1/0)
 *
 * La selección es siempre un rectángulo {r1,c1,r2,c2} (normalizado). Un
 * clic simple selecciona una celda; con Shift se extiende el rectángulo
 * desde el ancla; arrastrando con el ratón (mousedown + mousemove) también
 * selecciona un rango. Doble clic (o Enter/F2, o teclear directamente)
 * entra en modo edición de contenido de la celda activa.
 */
const WidgetTableEditor = {
    DEFAULT_ROWS: 15,
    DEFAULT_COLS: 8,
    COL_WIDTH: 96,
    ROW_HEIGHT: 26,

    FONTS: ["Arial", "Calibri", "Georgia", "Courier New", "Verdana", "Tahoma"],
    SIZES: [9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36],

    async open(widgetRow) {
        this.widget = {
            id: widgetRow.WIDGET_ID,
            name: widgetRow.WIDGET,
            description: widgetRow.DESCRIPCION || ""
        };

        let configRaw = "{}";
        try {
            const rows = await Provider.runQuery(
                `SELECT CONFIG_JSON FROM ${Provider.qualifyControl("WIDGETS")} WHERE WIDGET_ID = '${Provider.esc(this.widget.id)}'`
            );
            configRaw = (rows[0] && rows[0].CONFIG_JSON) || "{}";
        } catch (err) {
            UI.toast("No se pudo cargar el contenido del widget: " + err.message, "error");
        }

        this.state = this.parseConfig(configRaw);
        this.selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
        this.editingCell = null;
        this.dragging = false;

        this.renderModal();
    },

    blankState() {
        return { rows: this.DEFAULT_ROWS, cols: this.DEFAULT_COLS, cells: {}, merges: [] };
    },

    parseConfig(raw) {
        try {
            const parsed = JSON.parse(raw || "{}");
            return {
                rows: parsed.rows || this.DEFAULT_ROWS,
                cols: parsed.cols || this.DEFAULT_COLS,
                cells: parsed.cells || {},
                merges: parsed.merges || []
            };
        } catch (e) {
            return this.blankState();
        }
    },

    // ------------------------------------------------------------
    // Modal
    // ------------------------------------------------------------
    renderModal() {
        let overlay = document.getElementById("widgetTableModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "widgetTableModal";
            document.body.appendChild(overlay);
        }
        this.overlay = overlay;

        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div class="modal-title-edit-wrap">
                        <h3 id="wteTitle" class="modal-title-editable" contenteditable="true" spellcheck="false" title="Clic para renombrar el widget"></h3>
                        <span class="modal-subtitle">Widget de tipo Tabla</span>
                    </div>
                    <div class="modal-header-right">
                        <button class="modal-close" id="wteClose">&times;</button>
                    </div>
                </div>
                <div class="modal-body modal-body-flush">
                    <div class="wte-toolbar" id="wteToolbar"></div>
                    <div class="wte-grid-wrap" id="wteGridWrap"></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="wteCancel">Cancelar</button>
                    <button class="btn btn-primary" id="wteSave">Guardar widget</button>
                </div>
            </div>`;

        document.getElementById("wteTitle").textContent = this.widget.name;
        document.getElementById("wteClose").addEventListener("click", () => this.close());
        document.getElementById("wteCancel").addEventListener("click", () => this.close());
        document.getElementById("wteSave").addEventListener("click", () => this.save());

        const titleEl = document.getElementById("wteTitle");
        titleEl.addEventListener("keydown", (e) => {
            if (e.key === "Enter") { e.preventDefault(); titleEl.blur(); }
        });
        titleEl.addEventListener("blur", () => {
            const newName = titleEl.textContent.trim();
            if (!newName) { titleEl.textContent = this.widget.name; return; }
            this.widget.name = newName;
        });

        overlay.classList.add("visible");

        this.renderToolbar();
        this.renderGrid();

        document.addEventListener("keydown", this._keydownHandler = (e) => this.onGlobalKeydown(e));
        document.addEventListener("mouseup", this._mouseupHandler = () => { this.dragging = false; });
    },

    close() {
        if (this.overlay) this.overlay.classList.remove("visible");
        document.removeEventListener("keydown", this._keydownHandler);
        document.removeEventListener("mouseup", this._mouseupHandler);
    },

    // ------------------------------------------------------------
    // Toolbar
    // ------------------------------------------------------------
    renderToolbar() {
        const bar = document.getElementById("wteToolbar");
        bar.innerHTML = `
            <select id="wteFont" title="Tipo de letra">
                ${this.FONTS.map(f => `<option value="${f}">${f}</option>`).join("")}
            </select>
            <select id="wteSize" title="Tamaño de letra">
                ${this.SIZES.map(s => `<option value="${s}">${s}</option>`).join("")}
            </select>
            <span class="wte-toolbar-sep"></span>
            <button class="actupd-icon-btn" id="wteBold" title="Negrita"><strong>B</strong></button>
            <button class="actupd-icon-btn" id="wteItalic" title="Cursiva"><em>I</em></button>
            <button class="actupd-icon-btn" id="wteUnderline" title="Subrayado"><span style="text-decoration:underline;">S</span></button>
            <span class="wte-toolbar-sep"></span>
            <label class="wte-color-label" title="Color de texto">A<input type="color" id="wteColor" value="#1a1f2b"></label>
            <label class="wte-color-label" title="Color de relleno">▧<input type="color" id="wteBg" value="#ffffff"></label>
            <span class="wte-toolbar-sep"></span>
            <button class="actupd-icon-btn" id="wteAlignLeft" title="Alinear izquierda">⟸</button>
            <button class="actupd-icon-btn" id="wteAlignCenter" title="Centrar">≡</button>
            <button class="actupd-icon-btn" id="wteAlignRight" title="Alinear derecha">⟹</button>
            <span class="wte-toolbar-sep"></span>
            <button class="btn btn-secondary btn-sm" id="wteBorderOuter" title="Borde exterior de la selección">▢ Exterior</button>
            <button class="btn btn-secondary btn-sm" id="wteBorderAll" title="Todos los bordes de la selección">▦ Todos</button>
            <button class="btn btn-secondary btn-sm" id="wteBorderClear" title="Quitar bordes">✕ Bordes</button>
            <span class="wte-toolbar-sep"></span>
            <button class="btn btn-secondary btn-sm" id="wteMerge" title="Combinar celdas seleccionadas">⛭ Combinar</button>
            <button class="btn btn-secondary btn-sm" id="wteUnmerge" title="Separar celdas">⛝ Separar</button>
            <span class="wte-toolbar-spacer"></span>
            <button class="btn btn-secondary btn-sm" id="wteAddRow" title="Añadir fila">+ Fila</button>
            <button class="btn btn-secondary btn-sm" id="wteDelRow" title="Eliminar fila seleccionada">− Fila</button>
            <button class="btn btn-secondary btn-sm" id="wteAddCol" title="Añadir columna">+ Columna</button>
            <button class="btn btn-secondary btn-sm" id="wteDelCol" title="Eliminar columna seleccionada">− Columna</button>
        `;

        document.getElementById("wteFont").addEventListener("change", (e) => this.applyToSelection(c => c.ff = e.target.value));
        document.getElementById("wteSize").addEventListener("change", (e) => this.applyToSelection(c => c.fs = parseInt(e.target.value, 10)));
        document.getElementById("wteBold").addEventListener("click", () => this.toggleStyle("b"));
        document.getElementById("wteItalic").addEventListener("click", () => this.toggleStyle("i"));
        document.getElementById("wteUnderline").addEventListener("click", () => this.toggleStyle("u"));
        document.getElementById("wteColor").addEventListener("input", (e) => this.applyToSelection(c => c.col = e.target.value));
        document.getElementById("wteBg").addEventListener("input", (e) => this.applyToSelection(c => c.bg = e.target.value));
        document.getElementById("wteAlignLeft").addEventListener("click", () => this.applyToSelection(c => c.al = "left"));
        document.getElementById("wteAlignCenter").addEventListener("click", () => this.applyToSelection(c => c.al = "center"));
        document.getElementById("wteAlignRight").addEventListener("click", () => this.applyToSelection(c => c.al = "right"));
        document.getElementById("wteBorderOuter").addEventListener("click", () => this.applyBorder("outer"));
        document.getElementById("wteBorderAll").addEventListener("click", () => this.applyBorder("all"));
        document.getElementById("wteBorderClear").addEventListener("click", () => this.applyBorder("clear"));
        document.getElementById("wteMerge").addEventListener("click", () => this.mergeSelection());
        document.getElementById("wteUnmerge").addEventListener("click", () => this.unmergeSelection());
        document.getElementById("wteAddRow").addEventListener("click", () => this.insertRow());
        document.getElementById("wteDelRow").addEventListener("click", () => this.deleteRow());
        document.getElementById("wteAddCol").addEventListener("click", () => this.insertCol());
        document.getElementById("wteDelCol").addEventListener("click", () => this.deleteCol());
    },

    syncToolbar() {
        const cell = this.getCell(this.selection.r1, this.selection.c1);
        document.getElementById("wteFont").value = cell.ff || this.FONTS[0];
        document.getElementById("wteSize").value = cell.fs || 12;
        document.getElementById("wteColor").value = cell.col || "#1a1f2b";
        document.getElementById("wteBg").value = (cell.bg && cell.bg !== "transparent") ? cell.bg : "#ffffff";
        document.getElementById("wteBold").classList.toggle("active", !!cell.b);
        document.getElementById("wteItalic").classList.toggle("active", !!cell.i);
        document.getElementById("wteUnderline").classList.toggle("active", !!cell.u);
    },

    // ------------------------------------------------------------
    // Modelo de celdas
    // ------------------------------------------------------------
    cellKey(r, c) { return `${r}_${c}`; },

    getCell(r, c) {
        return this.state.cells[this.cellKey(r, c)] || {};
    },

    ensureCell(r, c) {
        const key = this.cellKey(r, c);
        if (!this.state.cells[key]) this.state.cells[key] = {};
        return this.state.cells[key];
    },

    // Celda ancla que "cubre" (r,c): ella misma, o el ancla de la fusión
    // a la que pertenece.
    anchorFor(r, c) {
        const merge = this.state.merges.find(m =>
            r >= m.r && r < m.r + m.rowSpan && c >= m.c && c < m.c + m.colSpan);
        return merge ? { r: merge.r, c: merge.c } : { r, c };
    },

    coveredMap() {
        const covered = {};
        this.state.merges.forEach(m => {
            for (let r = m.r; r < m.r + m.rowSpan; r++) {
                for (let c = m.c; c < m.c + m.colSpan; c++) {
                    if (r === m.r && c === m.c) continue;
                    covered[this.cellKey(r, c)] = true;
                }
            }
        });
        return covered;
    },

    // ------------------------------------------------------------
    // Selección
    // ------------------------------------------------------------
    normalizedSelection() {
        let { r1, c1, r2, c2 } = this.selection;
        let nr1 = Math.min(r1, r2), nr2 = Math.max(r1, r2);
        let nc1 = Math.min(c1, c2), nc2 = Math.max(c1, c2);

        // Expande la selección para incluir por completo cualquier celda
        // fusionada que quede parcialmente dentro del rectángulo.
        let changed = true;
        while (changed) {
            changed = false;
            this.state.merges.forEach(m => {
                const overlaps = m.r <= nr2 && m.r + m.rowSpan - 1 >= nr1 && m.c <= nc2 && m.c + m.colSpan - 1 >= nc1;
                if (!overlaps) return;
                if (m.r < nr1) { nr1 = m.r; changed = true; }
                if (m.r + m.rowSpan - 1 > nr2) { nr2 = m.r + m.rowSpan - 1; changed = true; }
                if (m.c < nc1) { nc1 = m.c; changed = true; }
                if (m.c + m.colSpan - 1 > nc2) { nc2 = m.c + m.colSpan - 1; changed = true; }
            });
        }
        return { r1: nr1, c1: nc1, r2: nr2, c2: nc2 };
    },

    applyToSelection(fn) {
        const sel = this.normalizedSelection();
        const seen = new Set();
        for (let r = sel.r1; r <= sel.r2; r++) {
            for (let c = sel.c1; c <= sel.c2; c++) {
                const anchor = this.anchorFor(r, c);
                const key = this.cellKey(anchor.r, anchor.c);
                if (seen.has(key)) continue;
                seen.add(key);
                fn(this.ensureCell(anchor.r, anchor.c));
            }
        }
        this.renderGrid();
    },

    toggleStyle(prop) {
        const anchor = this.anchorFor(this.selection.r1, this.selection.c1);
        const current = !!this.getCell(anchor.r, anchor.c)[prop];
        this.applyToSelection(c => c[prop] = current ? 0 : 1);
        this.syncToolbar();
    },

    applyBorder(mode) {
        const sel = this.normalizedSelection();
        const seen = new Set();
        for (let r = sel.r1; r <= sel.r2; r++) {
            for (let c = sel.c1; c <= sel.c2; c++) {
                const anchor = this.anchorFor(r, c);
                const key = this.cellKey(anchor.r, anchor.c);
                if (seen.has(key)) continue;
                seen.add(key);
                const cell = this.ensureCell(anchor.r, anchor.c);
                if (mode === "clear") {
                    cell.bt = cell.br = cell.bb = cell.bl = 0;
                } else if (mode === "all") {
                    cell.bt = cell.br = cell.bb = cell.bl = 1;
                } else if (mode === "outer") {
                    if (r === sel.r1) cell.bt = 1;
                    if (r === sel.r2) cell.bb = 1;
                    if (c === sel.c1) cell.bl = 1;
                    if (c === sel.c2) cell.br = 1;
                }
            }
        }
        this.renderGrid();
    },

    // ------------------------------------------------------------
    // Combinar / separar celdas
    // ------------------------------------------------------------
    mergeSelection() {
        const sel = this.normalizedSelection();
        if (sel.r1 === sel.r2 && sel.c1 === sel.c2) {
            UI.toast("Selecciona más de una celda para combinar.", "error");
            return;
        }
        // Disuelve cualquier fusión existente que se solape con la selección.
        this.state.merges = this.state.merges.filter(m => {
            const overlaps = m.r <= sel.r2 && m.r + m.rowSpan - 1 >= sel.r1 && m.c <= sel.c2 && m.c + m.colSpan - 1 >= sel.c1;
            return !overlaps;
        });
        // Conserva solo el contenido de la celda superior-izquierda; el resto
        // de celdas del rango se vacían (quedan cubiertas por la fusión).
        for (let r = sel.r1; r <= sel.r2; r++) {
            for (let c = sel.c1; c <= sel.c2; c++) {
                if (r === sel.r1 && c === sel.c1) continue;
                delete this.state.cells[this.cellKey(r, c)];
            }
        }
        this.state.merges.push({ r: sel.r1, c: sel.c1, rowSpan: sel.r2 - sel.r1 + 1, colSpan: sel.c2 - sel.c1 + 1 });
        this.selection = { r1: sel.r1, c1: sel.c1, r2: sel.r1, c2: sel.c1 };
        this.renderGrid();
    },

    unmergeSelection() {
        const sel = this.normalizedSelection();
        const before = this.state.merges.length;
        this.state.merges = this.state.merges.filter(m => {
            const overlaps = m.r <= sel.r2 && m.r + m.rowSpan - 1 >= sel.r1 && m.c <= sel.c2 && m.c + m.colSpan - 1 >= sel.c1;
            return !overlaps;
        });
        if (this.state.merges.length === before) {
            UI.toast("La selección no contiene celdas combinadas.", "info");
            return;
        }
        this.renderGrid();
    },

    // ------------------------------------------------------------
    // Insertar / eliminar filas y columnas
    // ------------------------------------------------------------
    insertRow() {
        const idx = Math.min(this.selection.r1, this.selection.r2) + 1; // debajo de la selección
        this.unmergeCrossing("row", idx);

        const newCells = {};
        Object.keys(this.state.cells).forEach(key => {
            const [r, c] = key.split("_").map(Number);
            newCells[this.cellKey(r >= idx ? r + 1 : r, c)] = this.state.cells[key];
        });
        this.state.cells = newCells;

        this.state.merges.forEach(m => {
            if (idx <= m.r) m.r += 1;
            else if (idx > m.r && idx < m.r + m.rowSpan) m.rowSpan += 1;
        });

        this.state.rows += 1;
        this.renderGrid();
    },

    deleteRow() {
        if (this.state.rows <= 1) { UI.toast("El widget necesita al menos una fila.", "error"); return; }
        const idx = Math.min(this.selection.r1, this.selection.r2);
        this.unmergeCrossing("row", idx, true);

        const newCells = {};
        Object.keys(this.state.cells).forEach(key => {
            const [r, c] = key.split("_").map(Number);
            if (r === idx) return;
            newCells[this.cellKey(r > idx ? r - 1 : r, c)] = this.state.cells[key];
        });
        this.state.cells = newCells;

        this.state.merges.forEach(m => { if (idx < m.r) m.r -= 1; });

        this.state.rows -= 1;
        this.selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
        this.renderGrid();
    },

    insertCol() {
        const idx = Math.min(this.selection.c1, this.selection.c2) + 1;
        this.unmergeCrossing("col", idx);

        const newCells = {};
        Object.keys(this.state.cells).forEach(key => {
            const [r, c] = key.split("_").map(Number);
            newCells[this.cellKey(r, c >= idx ? c + 1 : c)] = this.state.cells[key];
        });
        this.state.cells = newCells;

        this.state.merges.forEach(m => {
            if (idx <= m.c) m.c += 1;
            else if (idx > m.c && idx < m.c + m.colSpan) m.colSpan += 1;
        });

        this.state.cols += 1;
        this.renderGrid();
    },

    deleteCol() {
        if (this.state.cols <= 1) { UI.toast("El widget necesita al menos una columna.", "error"); return; }
        const idx = Math.min(this.selection.c1, this.selection.c2);
        this.unmergeCrossing("col", idx, true);

        const newCells = {};
        Object.keys(this.state.cells).forEach(key => {
            const [r, c] = key.split("_").map(Number);
            if (c === idx) return;
            newCells[this.cellKey(r, c > idx ? c - 1 : c)] = this.state.cells[key];
        });
        this.state.cells = newCells;

        this.state.merges.forEach(m => { if (idx < m.c) m.c -= 1; });

        this.state.cols -= 1;
        this.selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
        this.renderGrid();
    },

    // Disuelve (antes de insertar/eliminar) cualquier fusión cuyo rango
    // cruce por el índice de fila/columna afectado, para no dejar el
    // estado de fusiones inconsistente.
    unmergeCrossing(axis, idx, isDelete = false) {
        this.state.merges = this.state.merges.filter(m => {
            if (axis === "row") {
                const crosses = isDelete
                    ? (idx >= m.r && idx < m.r + m.rowSpan && m.rowSpan > 1)
                    : (idx > m.r && idx < m.r + m.rowSpan);
                return !crosses;
            } else {
                const crosses = isDelete
                    ? (idx >= m.c && idx < m.c + m.colSpan && m.colSpan > 1)
                    : (idx > m.c && idx < m.c + m.colSpan);
                return !crosses;
            }
        });
    },

    // ------------------------------------------------------------
    // Render de la rejilla
    // ------------------------------------------------------------
    renderGrid() {
        const wrap = document.getElementById("wteGridWrap");
        const covered = this.coveredMap();
        const sel = this.normalizedSelection();

        let html = `<table class="wte-table"><thead><tr><th class="wte-corner"></th>`;
        for (let c = 0; c < this.state.cols; c++) {
            html += `<th class="wte-colhead" data-colhead="${c}" style="width:${this.COL_WIDTH}px;">${this.colLabel(c)}</th>`;
        }
        html += `</tr></thead><tbody>`;

        for (let r = 0; r < this.state.rows; r++) {
            html += `<tr><th class="wte-rowhead" data-rowhead="${r}" style="height:${this.ROW_HEIGHT}px;">${r + 1}</th>`;
            for (let c = 0; c < this.state.cols; c++) {
                if (covered[this.cellKey(r, c)]) continue;

                const merge = this.state.merges.find(m => m.r === r && m.c === c);
                const rowSpan = merge ? merge.rowSpan : 1;
                const colSpan = merge ? merge.colSpan : 1;
                const cell = this.getCell(r, c);
                const isSelected = r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2;
                const isAnchorSel = r === Math.min(this.selection.r1, this.selection.r2) && c === Math.min(this.selection.c1, this.selection.c2);

                const style = [
                    `font-family:${cell.ff || "inherit"}`,
                    `font-size:${cell.fs || 12}px`,
                    `font-weight:${cell.b ? "700" : "400"}`,
                    `font-style:${cell.i ? "italic" : "normal"}`,
                    `text-decoration:${cell.u ? "underline" : "none"}`,
                    `text-align:${cell.al || "left"}`,
                    `color:${cell.col || "inherit"}`,
                    `background-color:${cell.bg || "#ffffff"}`,
                    `border-top:${cell.bt ? "2px solid #1a1f2b" : "1px solid #E3E6EC"}`,
                    `border-right:${cell.br ? "2px solid #1a1f2b" : "1px solid #E3E6EC"}`,
                    `border-bottom:${cell.bb ? "2px solid #1a1f2b" : "1px solid #E3E6EC"}`,
                    `border-left:${cell.bl ? "2px solid #1a1f2b" : "1px solid #E3E6EC"}`
                ].join(";");

                html += `<td class="wte-cell${isSelected ? " wte-cell-selected" : ""}${isAnchorSel ? " wte-cell-anchor" : ""}"
                    data-r="${r}" data-c="${c}" rowspan="${rowSpan}" colspan="${colSpan}"
                    style="${style}">${UI.escapeHtml(cell.v || "")}</td>`;
            }
            html += `</tr>`;
        }
        html += `</tbody></table>`;
        wrap.innerHTML = html;

        wrap.querySelectorAll(".wte-cell").forEach(td => {
            const r = parseInt(td.dataset.r, 10), c = parseInt(td.dataset.c, 10);
            td.addEventListener("mousedown", (e) => this.onCellMouseDown(e, r, c));
            td.addEventListener("mouseenter", () => this.onCellMouseEnter(r, c));
            td.addEventListener("dblclick", () => this.startEditing(r, c));
        });
        wrap.querySelectorAll("[data-colhead]").forEach(th => {
            th.addEventListener("click", () => {
                const c = parseInt(th.dataset.colhead, 10);
                this.selection = { r1: 0, c1: c, r2: this.state.rows - 1, c2: c };
                this.renderGrid();
                this.syncToolbar();
            });
        });
        wrap.querySelectorAll("[data-rowhead]").forEach(th => {
            th.addEventListener("click", () => {
                const r = parseInt(th.dataset.rowhead, 10);
                this.selection = { r1: r, c1: 0, r2: r, c2: this.state.cols - 1 };
                this.renderGrid();
                this.syncToolbar();
            });
        });

        this.syncToolbar();
    },

    colLabel(c) {
        let label = "";
        c += 1;
        while (c > 0) {
            const rem = (c - 1) % 26;
            label = String.fromCharCode(65 + rem) + label;
            c = Math.floor((c - 1) / 26);
        }
        return label;
    },

    // ------------------------------------------------------------
    // Interacción: selección con clic / arrastre, edición de contenido
    // ------------------------------------------------------------
    onCellMouseDown(e, r, c) {
        if (this.editingCell && (this.editingCell.r !== r || this.editingCell.c !== c)) {
            this.commitEditing();
        }
        if (this.editingCell) return; // ya editando esta misma celda: deja que el clic coloque el cursor

        if (e.shiftKey) {
            this.selection.r2 = r;
            this.selection.c2 = c;
        } else {
            this.selection = { r1: r, c1: c, r2: r, c2: c };
            this.dragging = true;
        }
        this.renderGrid();
    },

    onCellMouseEnter(r, c) {
        if (!this.dragging) return;
        this.selection.r2 = r;
        this.selection.c2 = c;
        this.renderGrid();
    },

    startEditing(r, c) {
        const anchor = this.anchorFor(r, c);
        this.selection = { r1: anchor.r, c1: anchor.c, r2: anchor.r, c2: anchor.c };
        this.editingCell = { r: anchor.r, c: anchor.c };
        this.renderGrid();

        const td = document.querySelector(`.wte-cell[data-r="${anchor.r}"][data-c="${anchor.c}"]`);
        if (!td) return;
        td.contentEditable = "true";
        td.focus();
        document.execCommand("selectAll", false, null);

        td.addEventListener("blur", () => this.commitEditing(), { once: true });
        td.addEventListener("keydown", (e) => {
            if (e.key === "Enter") { e.preventDefault(); td.blur(); }
            if (e.key === "Escape") { e.preventDefault(); this.editingCell = null; this.renderGrid(); }
        });
    },

    commitEditing() {
        if (!this.editingCell) return;
        const { r, c } = this.editingCell;
        const td = document.querySelector(`.wte-cell[data-r="${r}"][data-c="${c}"]`);
        const value = td ? td.innerText.replace(/\n/g, " ").trim() : "";
        if (value) {
            this.ensureCell(r, c).v = value;
        } else if (this.state.cells[this.cellKey(r, c)]) {
            this.state.cells[this.cellKey(r, c)].v = "";
        }
        this.editingCell = null;
        this.renderGrid();
    },

    onGlobalKeydown(e) {
        if (!this.overlay || !this.overlay.classList.contains("visible")) return;
        if (this.editingCell) return; // se gestiona en el propio td

        const active = document.activeElement;
        if (active && active.id === "wteTitle") return;

        if (e.key === "Enter" || e.key === "F2") {
            e.preventDefault();
            this.startEditing(this.selection.r1, this.selection.c1);
            return;
        }
        if (e.key === "Delete" || e.key === "Backspace") {
            e.preventDefault();
            this.applyToSelection(c => c.v = "");
            return;
        }
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            this.startEditing(this.selection.r1, this.selection.c1);
            setTimeout(() => {
                const td = document.querySelector(`.wte-cell.wte-cell-anchor`);
                if (td) td.textContent = e.key;
                if (td) { const range = document.createRange(); range.selectNodeContents(td); range.collapse(false); const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range); }
            }, 0);
        }
    },

    // ------------------------------------------------------------
    // Guardar
    // ------------------------------------------------------------
    async save() {
        if (this.editingCell) this.commitEditing();
        const btn = document.getElementById("wteSave");
        btn.disabled = true;
        btn.textContent = "Guardando...";
        try {
            const configJson = JSON.stringify(this.state).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
            const sql = `UPDATE ${Provider.qualifyControl("WIDGETS")}
                SET ${Widgets.NAME_COL} = '${Provider.esc(this.widget.name)}',
                    CONFIG_JSON = '${configJson}',
                    FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                WHERE ${Widgets.ID_COL} = '${Provider.esc(this.widget.id)}'`;
            await Provider.runQuery(sql);
            UI.toast(`Widget "${this.widget.name}" guardado.`, "success");
            this.close();
            await Widgets.loadList();
        } catch (err) {
            UI.toast("Error al guardar el widget: " + err.message, "error");
        } finally {
            btn.disabled = false;
            btn.textContent = "Guardar widget";
        }
    }
};
