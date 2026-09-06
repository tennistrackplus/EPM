/**
 * Editor tipo Excel para widgets de tipo TABLA.
 *
 * Estado del widget (se serializa completo en WIDGETS.CONFIG_JSON):
 *   {
 *     rows, cols,                     // dimensiones de la rejilla (por defecto 1000 x 200)
 *     cells: { "r_c": CellStyle },    // solo celdas con contenido/formato != por defecto
 *     merges: [{ r, c, rowSpan, colSpan }],  // r,c = celda ancla (arriba-izq.)
 *     colWidths: { "c": px },         // solo columnas cuyo ancho difiere del por defecto
 *     rowHeights: { "r": px }         // solo filas cuya altura difiere de la por defecto
 *   }
 *
 * CellStyle: { v, b, i, u, al, ff, fs, col, bg, bt, br, bb, bl }
 *   v=valor, b/i/u=negrita/cursiva/subrayado (1/0), al=alineación,
 *   ff=familia de fuente, fs=tamaño (px), col=color texto, bg=color fondo,
 *   bt/br/bb/bl = borde arriba/derecha/abajo/izquierda (1/0)
 *
 * La selección es siempre un rectángulo {r1,c1,r2,c2} (r1,c1 = celda ancla
 * donde empezó el clic; r2,c2 = esquina "activa", la que se mueve al
 * arrastrar o con flechas+Shift). Un clic simple selecciona una celda; con
 * Shift se extiende el rectángulo desde el ancla; arrastrando con el ratón
 * también selecciona un rango. Doble clic (o Enter/F2, o teclear
 * directamente) entra en modo edición de contenido de la celda activa.
 *
 * ------------------------------------------------------------------
 * Motor de rejilla (virtualizado)
 * ------------------------------------------------------------------
 * Con 1000 filas x 200 columnas por defecto (200.000 celdas posibles) no se
 * puede pintar un <table> completo: solo se crean nodos DOM para las
 * celdas realmente visibles en el viewport (+ un margen de buffer), tanto
 * para el cuerpo como para las cabeceras de fila/columna. La posición de
 * cada celda se calcula a partir de sumas acumuladas de anchos/altos
 * (this._colOffsets / this._rowOffsets), que se reconstruyen solo cuando
 * cambia la estructura (nº de filas/columnas, o el ancho/alto de alguna),
 * marcado con markDirty(). El scroll es nativo (overflow:auto) sobre
 * #wteBodyScroll; las cabeceras viven en paneles aparte y se desplazan con
 * un transform que sigue al scroll del cuerpo, para quedar siempre fijas.
 *
 * Anchos y altos están ligados al ÍNDICE de fila/columna, igual que las
 * celdas: insertar/eliminar una fila o columna reindexa colWidths/
 * rowHeights exactamente igual que reindexa cells, así que el resto de
 * columnas/filas conservan su tamaño (no se recalculan ni se "corren" los
 * anchos al borrar).
 */
// Estilos del editor incrustados directamente desde JS: así el motor de
// rejilla (posiciones absolutas, virtualización) funciona aunque el HTML
// que incluye este script no tenga bien enlazados (o cacheados) los
// ficheros css/widget-table-editor.css. Sin esta
// red de seguridad, si el <link> no carga, TODAS las celdas caen a
// position:static y se apilan verticalmente, y la virtualización pierde
// el límite de viewport real (el scroller ya no tiene overflow:auto),
// así que se acaban creando los ~200.000 nodos DOM de golpe.
const WTE_EMBEDDED_CSS = `
/* ---------- Editor de widget tipo Tabla (estilo Excel, virtualizado) ---------- */
.wte-toolbar {
    display: flex;
    align-items: center;
    gap: 6px;
    flex-wrap: wrap;
    padding: 8px 10px;
    border-bottom: 1px solid var(--border-subtle);
    background: var(--surface-sunken);
    flex-shrink: 0;
}

.wte-toolbar select {
    height: 30px;
    font-size: var(--fs-sm);
    border: 1px solid var(--border-subtle);
    border-radius: var(--radius-sm);
    padding: 0 6px;
    background: var(--surface-card);
}

#wteFont { width: 110px; }
#wteSize { width: 56px; }

.wte-toolbar-sep {
    width: 1px;
    height: 22px;
    background: var(--border-default);
    margin: 0 2px;
}

.wte-toolbar-spacer { flex: 1; }

.wte-toolbar-hint {
    font-size: var(--fs-sm);
    color: var(--text-muted);
    padding: 0 4px;
}

.wte-color-label {
    position: relative;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 30px;
    height: 30px;
    border: 1px solid var(--border-subtle);
    border-radius: var(--radius-sm);
    cursor: pointer;
    font-size: 13px;
    font-weight: 700;
    color: var(--text-secondary);
}

.wte-color-label input[type="color"] {
    position: absolute;
    inset: 0;
    opacity: 0;
    cursor: pointer;
}

/* ---------- Layout: corner + col headers | row headers + body (scroll) ---------- */
.wte-grid-wrap {
    flex: 1;
    min-height: 0;
    display: flex;
    flex-direction: row;
    overflow: hidden;
    background: var(--surface-card);
}

.wte-grid-main {
    flex: 1;
    min-width: 0;
    min-height: 0;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    user-select: none;
}

.wte-headrow {
    display: flex;
    flex-shrink: 0;
}

.wte-corner {
    flex-shrink: 0;
    width: 50px;
    height: 26px;
    background: var(--surface-sunken);
    border-right: 1px solid var(--border-default);
    border-bottom: 1px solid var(--border-default);
}

.wte-colhead-clip {
    flex: 1;
    min-width: 0;
    height: 26px;
    overflow: hidden;
    position: relative;
    background: var(--surface-sunken);
    border-bottom: 1px solid var(--border-default);
}

.wte-colhead-track {
    position: absolute;
    top: 0;
    left: 0;
    height: 100%;
}

.wte-colhead-cell {
    position: absolute;
    top: 0;
    height: 100%;
    box-sizing: border-box;
    border-right: 1px solid var(--border-default);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: var(--fs-sm);
    font-weight: 600;
    color: var(--text-muted);
    cursor: pointer;
}

.wte-colhead-cell:hover,
.wte-colhead-cell.wte-head-selected {
    background: var(--brand-primary-light);
    color: var(--brand-primary-hover);
}

.wte-resize-col {
    position: absolute;
    top: 0;
    right: -3px;
    width: 6px;
    height: 100%;
    cursor: col-resize;
    z-index: 3;
}

.wte-bodyrow {
    flex: 1;
    min-height: 0;
    display: flex;
}

.wte-rowhead-clip {
    flex-shrink: 0;
    width: 50px;
    overflow: hidden;
    position: relative;
    background: var(--surface-sunken);
    border-right: 1px solid var(--border-default);
}

.wte-rowhead-track {
    position: absolute;
    left: 0;
    top: 0;
    width: 100%;
}

.wte-rowhead-cell {
    position: absolute;
    left: 0;
    width: 100%;
    box-sizing: border-box;
    border-bottom: 1px solid var(--border-default);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: var(--fs-sm);
    font-weight: 600;
    color: var(--text-muted);
    cursor: pointer;
}

.wte-rowhead-cell:hover,
.wte-rowhead-cell.wte-head-selected {
    background: var(--brand-primary-light);
    color: var(--brand-primary-hover);
}

.wte-resize-row {
    position: absolute;
    left: 0;
    bottom: -3px;
    height: 6px;
    width: 100%;
    cursor: row-resize;
    z-index: 3;
}

.wte-body-scroll {
    flex: 1;
    min-height: 0;
    min-width: 0;
    overflow: auto;
    position: relative;
    background: var(--gray-50);
}

.wte-body-canvas {
    position: relative;
}

.wte-cell {
    position: absolute;
    box-sizing: border-box;
    padding: 3px 6px;
    white-space: pre-wrap;
    word-break: break-word;
    overflow: hidden;
    cursor: cell;
    outline: none;
    background: #ffffff;
}

.wte-cell-selected {
    box-shadow: inset 0 0 0 9999px rgba(42, 91, 215, 0.08);
}

.wte-cell-anchor {
    box-shadow: inset 0 0 0 2px var(--brand-primary);
    z-index: 2;
}

.wte-resizing {
    cursor: col-resize !important;
}

.wte-resizing-row {
    cursor: row-resize !important;
}

#wteToggleReport.active,
#wteRecognition.active {
    background: var(--brand-primary-light);
    border-color: var(--brand-primary);
    color: var(--brand-primary-hover);
}

/* El body del modal pasa a ser una fila: columna izquierda (toolbar +
   rejilla) + separador arrastrable + panel de informe, así el panel ocupa
   TODA la altura del modal (antes se quedaba solo a la altura de la
   rejilla, dejando un hueco vacío junto a la toolbar). Se usan dos clases
   (mayor especificidad) para no afectar a .modal-body-flush en otros
   modales de la aplicación. */
.wte-modal-body {
    display: flex !important;
    flex-direction: row !important;
    padding: 0 !important;
    overflow: hidden;
}

.wte-left-column {
    display: flex;
    flex-direction: column;
    flex: 1;
    min-width: 0;
    min-height: 0;
    overflow: hidden;
}

.wte-panel-resizer {
    width: 6px;
    flex-shrink: 0;
    cursor: col-resize;
    background: var(--border-default);
    display: none;
}

.wte-panel-resizer.visible {
    display: block;
}

.wte-panel-resizer:hover,
.wte-panel-resizer.dragging {
    background: var(--brand-primary);
}

/* Panel de informe: el taskpane real del add-in, en un iframe propio.
   Colapsado por defecto (width:0); al abrirse (☰ Informe) se expande.
   El ancho se puede arrastrar con .wte-panel-resizer (ver widget-table-editor.js). */
.wte-report-panel {
    width: 0;
    flex-shrink: 0;
    overflow: hidden;
    background: var(--surface-sunken);
    display: flex;
    flex-direction: column;
}

.wte-report-panel.visible {
    width: 420px;
}

.wte-taskpane-frame {
    flex: 1;
    width: 100%;
    height: 100%;
    border: none;
    background: #fff;
}

.wte-pivot-toggle {
    display: inline-block;
    width: 14px;
    cursor: pointer;
    color: var(--text-muted);
    font-size: 10px;
}

.wte-cell-pivot-head {
    font-weight: 600;
}

`;

const WidgetTableEditor = {
    DEFAULT_ROWS: 1000,
    DEFAULT_COLS: 200,
    DEFAULT_COL_WIDTH: 96,
    DEFAULT_ROW_HEIGHT: 26,
    MIN_COL_WIDTH: 32,
    MIN_ROW_HEIGHT: 18,
    BUFFER_PX: 200,

    FONTS: ["Arial", "Calibri", "Georgia", "Courier New", "Verdana", "Tahoma"],
    SIZES: [9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36],

    async open(widgetRow, project = null) {
        this.ensureStylesInjected();

        this.widget = {
            id: widgetRow.WIDGET_ID,
            name: widgetRow.WIDGET,
            description: widgetRow.DESCRIPCION || ""
        };
        this.project = project || this.project;

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

        // Estado interno del motor de rejilla (se reconstruye en cada open()).
        this._colOffsets = null;
        this._rowOffsets = null;
        this._totalWidth = 0;
        this._totalHeight = 0;
        this._offsetsDirty = true;
        this._rafPending = false;
        this._resizeCleanup = null;

        // Panel de informe (pivot): se crea perezosamente al pulsar "☰ Informe".
        this._reportPanelOpen = false;

        this.renderModal();

        if (typeof WidgetPivot !== "undefined") WidgetPivot.onWidgetOpened();
    },

    // Inyecta (una sola vez) el CSS crítico del motor de rejilla directamente
    // en <head>, para que funcione aunque el <link> a los .css falle o esté
    // cacheado con una versión antigua.
    ensureStylesInjected() {
        if (document.getElementById("wteEmbeddedStyles")) return;
        const style = document.createElement("style");
        style.id = "wteEmbeddedStyles";
        style.textContent = WTE_EMBEDDED_CSS;
        document.head.appendChild(style);
    },

    blankState() {
        return {
            rows: this.DEFAULT_ROWS,
            cols: this.DEFAULT_COLS,
            cells: {},
            merges: [],
            colWidths: {},
            rowHeights: {},
            report: this.defaultReport()
        };
    },

    defaultReport() {
        return {
            cuboId: null,
            rowField: null,   // { dimId, kind:'hierarchy'|'attribute', ref: nombreJerarquia|colId, label }
            colField: null,
            values: [],       // [{ name, column }]
            filters: [],      // [{ dimId, dimLabel, colId, values:[...] }]
            expandedRows: [], // rutas expandidas
            expandedCols: [],
            memberRecognition: false
        };
    },

    parseConfig(raw) {
        try {
            const parsed = JSON.parse(raw || "{}");
            // Object.assign en vez de reconstruir el objeto a mano: así se
            // conserva CUALQUIER clave adicional que el taskpane (host-bridge.js,
            // reportStore.js, etc.) haya ido guardando dentro del estado
            // (taskpaneSettings, hiddenSheets, hiddenSheetMerges, namedRanges...),
            // en vez de descartarla silenciosamente al recargar el widget.
            return Object.assign({}, parsed, {
                rows: parsed.rows || this.DEFAULT_ROWS,
                cols: parsed.cols || this.DEFAULT_COLS,
                cells: parsed.cells || {},
                merges: parsed.merges || [],
                colWidths: parsed.colWidths || {},
                rowHeights: parsed.rowHeights || {},
                report: parsed.report || this.defaultReport()
            });
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
                <div class="modal-body modal-body-flush wte-modal-body">
                    <div class="wte-left-column">
                        <div class="wte-toolbar" id="wteToolbar"></div>
                        <div class="wte-grid-wrap" id="wteGridWrap">
                            <div class="wte-grid-main" id="wteGridMain">
                                <div class="wte-headrow">
                                    <div class="wte-corner" id="wteCorner"></div>
                                    <div class="wte-colhead-clip" id="wteColHeadClip">
                                        <div class="wte-colhead-track" id="wteColHeadTrack"></div>
                                    </div>
                                </div>
                                <div class="wte-bodyrow">
                                    <div class="wte-rowhead-clip" id="wteRowHeadClip">
                                        <div class="wte-rowhead-track" id="wteRowHeadTrack"></div>
                                    </div>
                                    <div class="wte-body-scroll" id="wteBodyScroll">
                                        <div class="wte-body-canvas" id="wteBodyCanvas"></div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="wte-panel-resizer" id="wtePanelResizer"></div>
                    <div class="wte-report-panel" id="wteReportPanel"></div>
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
        this.initGrid();

        document.addEventListener("keydown", this._keydownHandler = (e) => this.onGlobalKeydown(e));
        document.addEventListener("mouseup", this._mouseupHandler = () => { this.dragging = false; });
    },

    close() {
        if (this._resizeCleanup) { this._resizeCleanup(); this._resizeCleanup = null; }
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
            <span class="wte-toolbar-sep"></span>
            <button class="btn btn-secondary btn-sm" id="wteToggleReport" title="Diseñar informe sobre un modelo semántico (cubo)">☰ Informe</button>
            <button class="btn btn-secondary btn-sm" id="wteRecognition" title="Reconocimiento de miembros">🔎 Reconocimiento</button>
            <button class="btn btn-secondary btn-sm" id="wteRefreshReport" title="Actualizar informe">⟳ Actualizar</button>
            <button class="btn btn-secondary btn-sm" id="wteAddFilterTop" title="Añadir filtro">▽ Filtro</button>
            <span class="wte-toolbar-spacer"></span>
            <span class="wte-toolbar-hint" id="wteDimHint"></span>
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
        document.getElementById("wteToggleReport").addEventListener("click", () => this.toggleReportPanel());
        document.getElementById("wteRecognition").addEventListener("click", () => this.callTaskpaneCommand("toggleMemberRecognition"));
        document.getElementById("wteRefreshReport").addEventListener("click", () => this.callTaskpaneCommand("actualizar"));
        document.getElementById("wteAddFilterTop").addEventListener("click", () => this.callTaskpaneCommand("abrirAnadirFiltro"));
        document.getElementById("wteAddRow").addEventListener("click", () => this.insertRow());
        document.getElementById("wteDelRow").addEventListener("click", () => this.deleteRow());
        document.getElementById("wteAddCol").addEventListener("click", () => this.insertCol());
        document.getElementById("wteDelCol").addEventListener("click", () => this.deleteCol());
    },

    syncToolbar() {
        const anchor = this.anchorFor(this.selection.r1, this.selection.c1);
        const cell = this.getCell(anchor.r, anchor.c);
        document.getElementById("wteFont").value = cell.ff || this.FONTS[0];
        document.getElementById("wteSize").value = cell.fs || 12;
        document.getElementById("wteColor").value = cell.col || "#1a1f2b";
        document.getElementById("wteBg").value = (cell.bg && cell.bg !== "transparent") ? cell.bg : "#ffffff";
        document.getElementById("wteBold").classList.toggle("active", !!cell.b);
        document.getElementById("wteItalic").classList.toggle("active", !!cell.i);
        document.getElementById("wteUnderline").classList.toggle("active", !!cell.u);
        const hint = document.getElementById("wteDimHint");
        if (hint) hint.textContent = `${this.state.rows} filas × ${this.state.cols} columnas`;
    },

    // ------------------------------------------------------------
    // Panel de informe: taskpane real del add-in en un iframe — ver
    // widget-taskpane/taskpane.html y js/toggleReportPanel más abajo.
    // ------------------------------------------------------------
    toggleReportPanel() {
        const panel = document.getElementById("wteReportPanel");
        const resizer = document.getElementById("wtePanelResizer");
        this._reportPanelOpen = !this._reportPanelOpen;
        panel.classList.toggle("visible", this._reportPanelOpen);
        resizer.classList.toggle("visible", this._reportPanelOpen);
        document.getElementById("wteToggleReport").classList.toggle("active", this._reportPanelOpen);
        if (this._reportPanelOpen && !panel.querySelector("iframe")) {
            // El taskpane real del add-in (ADDIN/src), copiado tal cual en
            // /widget-taskpane, corriendo en un iframe propio. Solo cambia
            // lo que en el add-in hablaba con Excel real o autenticaba su
            // propia conexión: aquí habla con esta rejilla (host-bridge.js)
            // y con la conexión ya autenticada de esta app (provider-bridge.js).
            const iframe = document.createElement("iframe");
            iframe.className = "wte-taskpane-frame";
            iframe.src = "widget-taskpane/taskpane.html?v=20260906b";
            panel.appendChild(iframe);
            this._taskpaneFrame = iframe;
            if (!resizer._wired) {
                resizer._wired = true;
                resizer.addEventListener("mousedown", (e) => this.startPanelResize(e));
            }
        }
        // El área del cuerpo cambia de ancho: recalcula el viewport visible.
        requestAnimationFrame(() => this.renderGrid());
    },

    // Arrastrar el separador entre la rejilla y el panel del taskpane
    // (pedido explícitamente: poder mover el reparto de espacio entre
    // Excel y el iframe para que quepa todo).
    startPanelResize(e) {
        e.preventDefault();
        const panel = document.getElementById("wteReportPanel");
        const resizer = document.getElementById("wtePanelResizer");
        const startX = e.clientX;
        const startWidth = panel.getBoundingClientRect().width;
        resizer.classList.add("dragging");
        document.body.style.userSelect = "none";

        const onMove = (ev) => {
            // El separador queda a la IZQUIERDA del panel: arrastrar hacia
            // la izquierda debe ENSANCHAR el panel (y viceversa).
            const newWidth = Math.max(280, Math.min(window.innerWidth - 320, startWidth - (ev.clientX - startX)));
            panel.style.width = newWidth + "px";
            requestAnimationFrame(() => this.renderGrid());
        };
        const onUp = () => {
            resizer.classList.remove("dragging");
            document.body.style.userSelect = "";
            document.removeEventListener("mousemove", onMove);
            document.removeEventListener("mouseup", onUp);
        };
        document.addEventListener("mousemove", onMove);
        document.addEventListener("mouseup", onUp);
    },

    // Llama a una función de comando del taskpane (commands.js) dentro del
    // iframe, con el mismo "event" con .completed() que espera un botón de
    // la cinta real de Excel.
    callTaskpaneCommand(fnName) {
        if (!this._reportPanelOpen) this.toggleReportPanel();
        const win = this._taskpaneFrame && this._taskpaneFrame.contentWindow;
        if (!win || typeof win[fnName] !== "function") {
            UI.toast("El panel de informe todavía se está cargando, espera un segundo y vuelve a intentarlo.", "info");
            return;
        }
        win[fnName]({ completed: () => {} });
    },

    // Escritura directa de una celda (valor + estilo/propiedades extra),
    // usada por WidgetPivot para pintar el resultado del informe.
    writeCell(r, c, value, extra = null) {
        const cell = this.ensureCell(r, c);
        cell.v = value;
        if (extra) Object.assign(cell, extra);
    },

    // Vacía por completo un rectángulo de celdas (contenido y estilo),
    // usado antes de repintar un informe para no dejar restos de una
    // ejecución anterior más grande.
    clearRegion(r1, c1, r2, c2) {
        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                delete this.state.cells[this.cellKey(r, c)];
            }
        }
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

    colWidth(c) { return this.state.colWidths[c] || this.DEFAULT_COL_WIDTH; },
    rowHeight(r) { return this.state.rowHeights[r] || this.DEFAULT_ROW_HEIGHT; },

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
        this.state.merges = this.state.merges.filter(m => {
            const overlaps = m.r <= sel.r2 && m.r + m.rowSpan - 1 >= sel.r1 && m.c <= sel.c2 && m.c + m.colSpan - 1 >= sel.c1;
            return !overlaps;
        });
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
    // (colWidths/rowHeights se reindexan igual que cells, para que el
    // resto de columnas/filas conserven su tamaño intacto)
    // ------------------------------------------------------------
    insertRow() {
        const idx = Math.min(this.selection.r1, this.selection.r2) + 1;
        this.unmergeCrossing("row", idx);

        const newCells = {};
        Object.keys(this.state.cells).forEach(key => {
            const [r, c] = key.split("_").map(Number);
            newCells[this.cellKey(r >= idx ? r + 1 : r, c)] = this.state.cells[key];
        });
        this.state.cells = newCells;

        const newHeights = {};
        Object.keys(this.state.rowHeights).forEach(key => {
            const r = Number(key);
            newHeights[r >= idx ? r + 1 : r] = this.state.rowHeights[key];
        });
        this.state.rowHeights = newHeights;

        this.state.merges.forEach(m => {
            if (idx <= m.r) m.r += 1;
            else if (idx > m.r && idx < m.r + m.rowSpan) m.rowSpan += 1;
        });

        this.state.rows += 1;
        this.markDirty();
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

        const newHeights = {};
        Object.keys(this.state.rowHeights).forEach(key => {
            const r = Number(key);
            if (r === idx) return;
            newHeights[r > idx ? r - 1 : r] = this.state.rowHeights[key];
        });
        this.state.rowHeights = newHeights;

        this.state.merges.forEach(m => { if (idx < m.r) m.r -= 1; });

        this.state.rows -= 1;
        this.selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
        this.markDirty();
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

        const newWidths = {};
        Object.keys(this.state.colWidths).forEach(key => {
            const c = Number(key);
            newWidths[c >= idx ? c + 1 : c] = this.state.colWidths[key];
        });
        this.state.colWidths = newWidths;

        this.state.merges.forEach(m => {
            if (idx <= m.c) m.c += 1;
            else if (idx > m.c && idx < m.c + m.colSpan) m.colSpan += 1;
        });

        this.state.cols += 1;
        this.markDirty();
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

        const newWidths = {};
        Object.keys(this.state.colWidths).forEach(key => {
            const c = Number(key);
            if (c === idx) return;
            newWidths[c > idx ? c - 1 : c] = this.state.colWidths[key];
        });
        this.state.colWidths = newWidths;

        this.state.merges.forEach(m => { if (idx < m.c) m.c -= 1; });

        this.state.cols -= 1;
        this.selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
        this.markDirty();
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
    // Motor de rejilla: offsets acumulados, búsqueda binaria y
    // render virtualizado (solo el rango visible + buffer)
    // ------------------------------------------------------------
    markDirty() { this._offsetsDirty = true; },

    buildOffsets() {
        const colOffsets = new Array(this.state.cols + 1);
        let acc = 0;
        for (let c = 0; c < this.state.cols; c++) { colOffsets[c] = acc; acc += this.colWidth(c); }
        colOffsets[this.state.cols] = acc;
        this._colOffsets = colOffsets;
        this._totalWidth = acc;

        const rowOffsets = new Array(this.state.rows + 1);
        acc = 0;
        for (let r = 0; r < this.state.rows; r++) { rowOffsets[r] = acc; acc += this.rowHeight(r); }
        rowOffsets[this.state.rows] = acc;
        this._rowOffsets = rowOffsets;
        this._totalHeight = acc;

        this._offsetsDirty = false;
    },

    ensureOffsets() {
        if (this._offsetsDirty || !this._colOffsets) this.buildOffsets();
    },

    colX(c) { return this._colOffsets[c]; },
    rowY(r) { return this._rowOffsets[r]; },

    // Índice i (0-based, < count) tal que offsets[i] <= pos < offsets[i+1].
    findIndex(offsets, count, pos) {
        if (count <= 0) return 0;
        if (pos <= offsets[0]) return 0;
        if (pos >= offsets[count - 1]) return count - 1;
        let lo = 0, hi = count - 1;
        while (lo < hi) {
            const mid = (lo + hi + 1) >> 1;
            if (offsets[mid] <= pos) lo = mid; else hi = mid - 1;
        }
        return lo;
    },

    initGrid() {
        const scroller = document.getElementById("wteBodyScroll");
        scroller.scrollTop = 0;
        scroller.scrollLeft = 0;
        scroller.addEventListener("scroll", () => this.scheduleRender());

        // Rueda del ratón sobre las cabeceras: se reenvía al scroll del cuerpo.
        ["wteColHeadClip", "wteRowHeadClip", "wteCorner"].forEach(id => {
            const el = document.getElementById(id);
            el.addEventListener("wheel", (e) => {
                e.preventDefault();
                scroller.scrollTop += e.deltaY;
                scroller.scrollLeft += e.deltaX;
            }, { passive: false });
        });

        this.markDirty();
        this.renderGrid();
    },

    scheduleRender() {
        if (this._rafPending) return;
        this._rafPending = true;
        requestAnimationFrame(() => { this._rafPending = false; this.renderViewport(); });
    },

    // Punto de entrada usado por el resto del editor (selección, estilos,
    // insertar/eliminar filas y columnas...). Nombre conservado por
    // compatibilidad con el resto del código.
    renderGrid() { this.renderViewport(); },

    renderViewport() {
        this.ensureOffsets();
        const scroller = document.getElementById("wteBodyScroll");
        if (!scroller) return;
        const scrollLeft = scroller.scrollLeft, scrollTop = scroller.scrollTop;
        const viewW = scroller.clientWidth || 800, viewH = scroller.clientHeight || 400;

        document.getElementById("wteColHeadTrack").style.transform = `translateX(${-scrollLeft}px)`;
        document.getElementById("wteRowHeadTrack").style.transform = `translateY(${-scrollTop}px)`;

        const c1 = this.findIndex(this._colOffsets, this.state.cols, Math.max(0, scrollLeft - this.BUFFER_PX));
        const c2 = this.findIndex(this._colOffsets, this.state.cols, scrollLeft + viewW + this.BUFFER_PX);
        const r1 = this.findIndex(this._rowOffsets, this.state.rows, Math.max(0, scrollTop - this.BUFFER_PX));
        const r2 = this.findIndex(this._rowOffsets, this.state.rows, scrollTop + viewH + this.BUFFER_PX);

        this.renderColHeaders(c1, c2);
        this.renderRowHeaders(r1, r2);
        this.renderCells(r1, r2, c1, c2);
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

    renderColHeaders(c1, c2) {
        const track = document.getElementById("wteColHeadTrack");
        track.style.width = this._totalWidth + "px";
        const sel = this.normalizedSelection();
        let html = "";
        for (let c = c1; c <= c2; c++) {
            const isSel = c >= sel.c1 && c <= sel.c2;
            html += `<div class="wte-colhead-cell${isSel ? " wte-head-selected" : ""}" data-colhead="${c}" style="left:${this.colX(c)}px;width:${this.colWidth(c)}px;">${this.colLabel(c)}<div class="wte-resize-col" data-resizecol="${c}"></div></div>`;
        }
        track.innerHTML = html;
        track.querySelectorAll(".wte-colhead-cell").forEach(el => {
            const c = parseInt(el.dataset.colhead, 10);
            el.addEventListener("mousedown", (e) => {
                if (e.target.closest("[data-resizecol]")) {
                    e.preventDefault();
                    this.startColResize(c, e.clientX);
                    return;
                }
                this.selection = { r1: 0, c1: c, r2: this.state.rows - 1, c2: c };
                this.renderGrid();
            });
        });
    },

    renderRowHeaders(r1, r2) {
        const track = document.getElementById("wteRowHeadTrack");
        track.style.height = this._totalHeight + "px";
        const sel = this.normalizedSelection();
        let html = "";
        for (let r = r1; r <= r2; r++) {
            const isSel = r >= sel.r1 && r <= sel.r2;
            html += `<div class="wte-rowhead-cell${isSel ? " wte-head-selected" : ""}" data-rowhead="${r}" style="top:${this.rowY(r)}px;height:${this.rowHeight(r)}px;">${r + 1}<div class="wte-resize-row" data-resizerow="${r}"></div></div>`;
        }
        track.innerHTML = html;
        track.querySelectorAll(".wte-rowhead-cell").forEach(el => {
            const r = parseInt(el.dataset.rowhead, 10);
            el.addEventListener("mousedown", (e) => {
                if (e.target.closest("[data-resizerow]")) {
                    e.preventDefault();
                    this.startRowResize(r, e.clientY);
                    return;
                }
                this.selection = { r1: r, c1: 0, r2: r, c2: this.state.cols - 1 };
                this.renderGrid();
            });
        });
    },

    startColResize(c, startX) {
        const startWidth = this.colWidth(c);
        document.body.classList.add("wte-resizing");
        const onMove = (e) => {
            this.state.colWidths[c] = Math.max(this.MIN_COL_WIDTH, startWidth + (e.clientX - startX));
            this.markDirty();
            this.renderGrid();
        };
        const onUp = () => {
            document.body.classList.remove("wte-resizing");
            document.removeEventListener("mousemove", onMove);
            document.removeEventListener("mouseup", onUp);
            this._resizeCleanup = null;
        };
        this._resizeCleanup = onUp;
        document.addEventListener("mousemove", onMove);
        document.addEventListener("mouseup", onUp);
    },

    startRowResize(r, startY) {
        const startHeight = this.rowHeight(r);
        document.body.classList.add("wte-resizing-row");
        const onMove = (e) => {
            this.state.rowHeights[r] = Math.max(this.MIN_ROW_HEIGHT, startHeight + (e.clientY - startY));
            this.markDirty();
            this.renderGrid();
        };
        const onUp = () => {
            document.body.classList.remove("wte-resizing-row");
            document.removeEventListener("mousemove", onMove);
            document.removeEventListener("mouseup", onUp);
            this._resizeCleanup = null;
        };
        this._resizeCleanup = onUp;
        document.addEventListener("mousemove", onMove);
        document.addEventListener("mouseup", onUp);
    },

    renderCells(r1, r2, c1, c2) {
        const canvas = document.getElementById("wteBodyCanvas");
        canvas.style.width = this._totalWidth + "px";
        canvas.style.height = this._totalHeight + "px";

        const sel = this.normalizedSelection();
        const renderedAnchors = new Set();
        let html = "";

        // 1) Fusiones que se solapan con el rectángulo visible, aunque su
        // ancla quede fuera del rango (celda combinada parcialmente scrolleada).
        this.state.merges.forEach(m => {
            const overlaps = m.r <= r2 && m.r + m.rowSpan - 1 >= r1 && m.c <= c2 && m.c + m.colSpan - 1 >= c1;
            if (!overlaps) return;
            renderedAnchors.add(this.cellKey(m.r, m.c));
            const w = this._colOffsets[m.c + m.colSpan] - this._colOffsets[m.c];
            const h = this._rowOffsets[m.r + m.rowSpan] - this._rowOffsets[m.r];
            html += this.cellHtml(m.r, m.c, this.colX(m.c), this.rowY(m.r), w, h, sel);
        });

        // 2) Celdas normales del rango visible (sin cubrir, sin ya renderizada).
        const covered = this.coveredMap();
        for (let r = r1; r <= r2; r++) {
            const y = this.rowY(r), h = this.rowHeight(r);
            for (let c = c1; c <= c2; c++) {
                const key = this.cellKey(r, c);
                if (covered[key] || renderedAnchors.has(key)) continue;
                html += this.cellHtml(r, c, this.colX(c), y, this.colWidth(c), h, sel);
            }
        }

        canvas.innerHTML = html;
        canvas.querySelectorAll(".wte-cell").forEach(div => {
            const r = parseInt(div.dataset.r, 10), c = parseInt(div.dataset.c, 10);
            div.addEventListener("mousedown", (e) => this.onCellMouseDown(e, r, c));
            div.addEventListener("mouseenter", () => this.onCellMouseEnter(r, c));
            div.addEventListener("dblclick", () => this.startEditing(r, c));
        });
        canvas.querySelectorAll(".wte-pivot-toggle").forEach(span => {
            span.addEventListener("mousedown", (e) => {
                e.preventDefault();
                e.stopPropagation();
                if (typeof WidgetPivot !== "undefined") {
                    WidgetPivot.toggleMember(parseInt(span.dataset.toggleR, 10), parseInt(span.dataset.toggleC, 10));
                }
            });
        });
    },

    cellHtml(r, c, x, y, w, h, sel) {
        const cell = this.getCell(r, c);
        const isSelected = r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2;
        const isAnchorSel = r === Math.min(this.selection.r1, this.selection.r2) && c === Math.min(this.selection.c1, this.selection.c2);

        const style = [
            `left:${x}px`, `top:${y}px`, `width:${w}px`, `height:${h}px`,
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

        // Celdas de cabecera de un informe (jerarquía): icono ▸/▾ clicable
        // en vez de un listener de hover, tal y como se pidió. El icono es
        // un <span> propio para poder distinguir su clic del de la celda.
        const inner = cell.tg
            ? `<span class="wte-pivot-toggle" data-toggle-r="${r}" data-toggle-c="${c}">${cell.tg.expanded ? "▾" : "▸"}</span>${UI.escapeHtml(cell.v || "")}`
            : UI.escapeHtml(cell.v || "");

        return `<div class="wte-cell${isSelected ? " wte-cell-selected" : ""}${isAnchorSel ? " wte-cell-anchor" : ""}${cell.tg ? " wte-cell-pivot-head" : ""}"
            data-r="${r}" data-c="${c}" style="${style}">${inner}</div>`;
    },

    scrollCellIntoView(r, c) {
        this.ensureOffsets();
        const scroller = document.getElementById("wteBodyScroll");
        const x = this.colX(c), w = this.colWidth(c);
        const y = this.rowY(r), h = this.rowHeight(r);
        if (x < scroller.scrollLeft) scroller.scrollLeft = x;
        else if (x + w > scroller.scrollLeft + scroller.clientWidth) scroller.scrollLeft = x + w - scroller.clientWidth;
        if (y < scroller.scrollTop) scroller.scrollTop = y;
        else if (y + h > scroller.scrollTop + scroller.clientHeight) scroller.scrollTop = y + h - scroller.clientHeight;
        this.renderViewport();
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
        this.fireTaskpaneEvent("onSelectionChanged", r, c);
        this.fireTaskpaneEvent("onSingleClicked", r, c);
    },

    // Dispara, dentro del iframe del taskpane (ver widget-taskpane/js/host-bridge.js),
    // los eventos de hoja de Excel de los que depende el reconocimiento de
    // miembros y el expandir/contraer con clic. No hace nada si el panel
    // "☰ Informe" no está abierto (no hay iframe al que avisar).
    fireTaskpaneEvent(type, r, c) {
        const win = this._taskpaneFrame && this._taskpaneFrame.contentWindow;
        if (win && typeof win.__fireExcelEvent === "function") win.__fireExcelEvent(type, r, c);
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
        this.scrollCellIntoView(anchor.r, anchor.c);

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
        this.fireTaskpaneEvent("onChanged", r, c);
        this.renderGrid();
    },

    onGlobalKeydown(e) {
        if (!this.overlay || !this.overlay.classList.contains("visible")) return;
        if (this.editingCell) return; // se gestiona en el propio div

        const active = document.activeElement;
        if (active && active.id === "wteTitle") return;

        const moveKeys = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
        if (moveKeys[e.key]) {
            e.preventDefault();
            const [dr, dc] = moveKeys[e.key];
            if (e.shiftKey) {
                this.selection.r2 = Math.max(0, Math.min(this.state.rows - 1, this.selection.r2 + dr));
                this.selection.c2 = Math.max(0, Math.min(this.state.cols - 1, this.selection.c2 + dc));
            } else {
                const r = Math.max(0, Math.min(this.state.rows - 1, this.selection.r1 + dr));
                const c = Math.max(0, Math.min(this.state.cols - 1, this.selection.c1 + dc));
                const anchor = this.anchorFor(r, c);
                this.selection = { r1: anchor.r, c1: anchor.c, r2: anchor.r, c2: anchor.c };
            }
            this.scrollCellIntoView(this.selection.r2, this.selection.c2);
            this.fireTaskpaneEvent("onSelectionChanged", this.selection.r2, this.selection.c2);
            return;
        }
        if (e.key === "Tab") {
            e.preventDefault();
            const c = Math.max(0, Math.min(this.state.cols - 1, this.selection.c1 + (e.shiftKey ? -1 : 1)));
            const anchor = this.anchorFor(this.selection.r1, c);
            this.selection = { r1: anchor.r, c1: anchor.c, r2: anchor.r, c2: anchor.c };
            this.scrollCellIntoView(this.selection.r2, this.selection.c2);
            return;
        }
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

window.WidgetTableEditor = WidgetTableEditor;
