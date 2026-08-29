/**
 * ============================================================
 * DRACO PLANNING — ACTUALIZACIÓN DE TABLAS
 * ============================================================
 * Mantenimiento manual de una tabla física del esquema del proyecto
 * (una dimensión, un cubo, o cualquier otra tabla) con:
 *   1) Diseñador (2 bloques):
 *      A) Pantalla de selección: variables sueltas (como en Flujos
 *         de proceso, versión simplificada sin frames) que luego se
 *         pueden usar como "filtro variable" de un campo.
 *      B) Campos de la tabla: orden (arrastrar), descripción, filtro
 *         (ninguno/constante/variable) y validación (ninguna/
 *         constante/valores de dimensión/valores de jerarquía), más
 *         "permite vacío", "mostrar texto" y ayuda de búsqueda
 *         (listbox/buscador tipo SAP/checkbox).
 *   2) Ejecución: pantalla de variables -> grid tipo ALV (buscar,
 *      filtrar, ordenar, añadir/borrar filas, descargar/cargar
 *      fichero) -> Grabar, con un resumen de cambios (altas/
 *      modificaciones/bajas) con detalle.
 *
 * Todo el diseño se guarda como JSON en una única fila de control
 * (DRACO_CONTROL.ACTUALIZACIONES), igual que CUBOS.CAMPOS_JSON.
 *
 * Alcance de esta primera versión (decisiones explícitas):
 *  - El selector de tabla solo lista tablas DENTRO del esquema del
 *    proyecto actual (no se navega todo Snowflake/BigQuery como en
 *    el add-in): es más simple y evita ediciones fuera de proyecto.
 *  - El grid trabaja "todo en memoria": se trae la tabla entera
 *    (con los filtros de cabecera aplicados) de una vez y se
 *    filtra/ordena/pagina en el navegador. Para no congelar el DOM
 *    con tablas de decenas de miles de filas, el HTML solo renderiza
 *    una "ventana" de filas visibles (ver RENDER_CHUNK) aunque el
 *    filtro/orden/edición siempre opera sobre el array completo.
 *  - El log de cambios al grabar es un resumen en pantalla (no se
 *    persiste en base de datos).
 */
const TableUpdates = {
    TABLE: "ACTUALIZACIONES",
    ID_COL: "ACTUALIZACION_ID",
    NAME_COL: "NOMBRE",
    RENDER_CHUNK: 300,

    list: [],
    project: null,
    editing: null,
    editingIsNew: true,
    screen: { title: "", blocks: [] },   // pantalla de variables del diseñador, mismo formato que Flows.editing.screen
    fields: [],
    dragFieldIdx: null,
    dragBlockIdx: null,
    dragFrameVar: null,
    screenCollapsed: false,
    dimensionsCache: [],

    // ---- ejecución (Bloque de "Ejecutar", ver más abajo) ----
    runRecord: null,
    runFields: [],
    runScreen: { title: "", blocks: [] },
    runView: "screen",       // 'screen' | 'table'
    gridState: null,
    selOptState: {},
    _selOptDelegationBound: false,

    /** Parsea VARIABLES_JSON admitiendo el formato antiguo (lista plana de
     * variables) y el nuevo (mismo formato que la pantalla de Flujos de
     * carga: {title, blocks:[{kind:'variable'|'frame'|'text'|'skip'|'line', ...}]}). */
    parseScreen(json) {
        const parsed = this.safeParse(json, null);
        if (!parsed) return { title: "", blocks: [] };
        if (Array.isArray(parsed)) {
            return {
                title: "",
                blocks: parsed.filter(v => v && v.name).map(v => ({
                    id: v.id || Provider.newId(),
                    kind: "variable",
                    variable: { id: v.id || Provider.newId(), name: v.name, label: v.label || v.name, type: v.type || "STRING", selectMode: v.selectMode || "unico" }
                }))
            };
        }
        return { title: parsed.title || "", blocks: Array.isArray(parsed.blocks) ? parsed.blocks : [] };
    },

    /** Lista plana de variables de pantalla (sueltas + las de todos los frames). */
    flatVars(screen) {
        const out = [];
        (screen.blocks || []).forEach(b => {
            if (b.kind === "variable") out.push(b.variable);
            else if (b.kind === "frame") out.push(...(b.variables || []));
        });
        return out;
    },

    // ================================================================
    // LISTADO
    // ================================================================
    async render(container, project) {
        this.container = container;
        this.project = project;

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <h3>Actualización de tablas</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET}</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewActualizacion">
                    + Nueva actualización
                </button>
            </div>
            <div id="actualizacionesListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewActualizacion").addEventListener("click", () => this.openCreateModal());
        await this.loadList();
    },

    async loadList() {
        const wrap = document.getElementById("actualizacionesListWrap");
        try {
            // Autoreparación: si esta tabla de control todavía no existe en el
            // esquema (sesión abierta antes de desplegar este módulo, sin pasar
            // de nuevo por el login que ejecuta DracoSchema.bootstrap), se crea
            // aquí mismo bajo demanda en lugar de fallar.
            try {
                await Provider.runQuery(DracoSchema.ddl(this.TABLE));
            } catch (ddlErr) {
                console.error("No se pudo verificar/crear la tabla de control ACTUALIZACIONES:", ddlErr);
            }

            const sql = `SELECT ${this.ID_COL}, ${this.NAME_COL}, DESCRIPCION, TABLA, VARIABLES_JSON, CAMPOS_JSON
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                         ORDER BY ${this.NAME_COL}`;
            this.list = await Provider.runQuery(sql);

            if (!this.list.length) {
                wrap.innerHTML = `<div class="module-empty">Todavía no hay actualizaciones en este proyecto. Crea la primera con "Nueva actualización".</div>`;
                return;
            }

            wrap.innerHTML = `
                <div class="data-list">
                    <table>
                        <thead><tr><th>Nombre</th><th>Tabla</th><th>Campos</th><th></th></tr></thead>
                        <tbody>
                            ${this.list.map(r => {
                                const fields = this.safeParse(r.CAMPOS_JSON, []);
                                return `
                                <tr>
                                    <td><strong>${UI.escapeHtml(r[this.NAME_COL])}</strong><br><span class="col-type">${UI.escapeHtml(r.DESCRIPCION || "")}</span></td>
                                    <td><span class="table-tag">${UI.escapeHtml(r.TABLA)}</span></td>
                                    <td>${fields.length} campo${fields.length === 1 ? "" : "s"}</td>
                                    <td>
                                        <div class="row-actions">
                                            <button data-run="${r[this.ID_COL]}" title="Ejecutar">▶</button>
                                            <button data-edit="${r[this.ID_COL]}" title="Editar">✎</button>
                                            <button data-del="${r[this.ID_COL]}" class="danger" title="Eliminar">🗑</button>
                                        </div>
                                    </td>
                                </tr>`;
                            }).join("")}
                        </tbody>
                    </table>
                </div>`;

            wrap.querySelectorAll("[data-run]").forEach(btn =>
                btn.addEventListener("click", () => this.startRun(this.list.find(r => r[this.ID_COL] === btn.dataset.run))));
            wrap.querySelectorAll("[data-edit]").forEach(btn =>
                btn.addEventListener("click", () => this.openEditor(this.list.find(r => r[this.ID_COL] === btn.dataset.edit))));
            wrap.querySelectorAll("[data-del]").forEach(btn =>
                btn.addEventListener("click", () => this.remove(btn.dataset.del)));
        } catch (err) {
            wrap.innerHTML = `<div class="module-empty">Error al cargar actualizaciones: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    async remove(id) {
        const rec = this.list.find(r => r[this.ID_COL] === id);
        const ok = await UI.confirm("Eliminar actualización", `¿Seguro que quieres eliminar "${rec ? rec[this.NAME_COL] : ""}"? Esto no borra la tabla física, solo su configuración.`);
        if (!ok) return;
        try {
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
            UI.toast("Actualización eliminada.", "success");
            await this.loadList();
        } catch (err) {
            UI.toast("Error al eliminar: " + err.message, "error");
        }
    },

    safeParse(json, fallback) {
        try { return json ? JSON.parse(json) : fallback; } catch (e) { return fallback; }
    },

    // ================================================================
    // CREAR (nombre + selector de tabla: base de datos / esquema / tabla)
    // ================================================================
    async openCreateModal() {
        let overlay = document.getElementById("actUpdCreateModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "actUpdCreateModal";
            document.body.appendChild(overlay);
        }
        overlay.classList.add("visible");

        overlay.innerHTML = `
            <div class="modal-box">
                <div class="modal-header">
                    <h3>Nueva actualización de tabla</h3>
                    <button class="modal-close" id="actUpdCreateClose">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="form-group">
                        <label>Nombre</label>
                        <input type="text" id="actUpdNewName" placeholder="ej. Mantenimiento presupuesto CECO">
                    </div>
                    <div class="form-group">
                        <label>Descripción (opcional)</label>
                        <input type="text" id="actUpdNewDesc">
                    </div>
                    <div class="form-group">
                        <label>Selector de tabla</label>
                        <div class="table-picker">
                            <div class="table-picker-row"><span>Base de datos</span><strong>${UI.escapeHtml(Provider.databaseLabel())}</strong></div>
                            <div class="table-picker-row"><span>Esquema</span><strong>${UI.escapeHtml(this.project.DATASET)}</strong></div>
                            <div class="table-picker-row">
                                <span>Tabla</span>
                                <select id="actUpdNewTable"><option value="">Cargando tablas...</option></select>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="actUpdCreateCancel">Cancelar</button>
                    <button class="btn btn-primary" id="actUpdCreateContinue">Continuar</button>
                </div>
            </div>`;

        const close = () => overlay.remove();
        document.getElementById("actUpdCreateClose").addEventListener("click", close);
        document.getElementById("actUpdCreateCancel").addEventListener("click", close);

        const tableSelect = document.getElementById("actUpdNewTable");
        try {
            const tables = await Provider.listTablesInContainer(this.project.DATASET);
            tableSelect.innerHTML = tables.length
                ? tables.map(t => `<option value="${UI.escapeHtml(t)}">${UI.escapeHtml(t)}</option>`).join("")
                : `<option value="">No hay tablas en este esquema</option>`;
        } catch (err) {
            tableSelect.innerHTML = `<option value="">Error al listar tablas</option>`;
        }

        document.getElementById("actUpdCreateContinue").addEventListener("click", () => {
            const name = document.getElementById("actUpdNewName").value.trim();
            const table = tableSelect.value;
            if (!name) { UI.toast("Indica un nombre.", "error"); return; }
            if (!table) { UI.toast("Selecciona una tabla.", "error"); return; }

            const draft = {
                [this.ID_COL]: Provider.newId(),
                PROYECTO_ID: this.project.PROYECTO_ID,
                [this.NAME_COL]: name,
                DESCRIPCION: document.getElementById("actUpdNewDesc").value.trim(),
                TABLA: table,
                VARIABLES_JSON: "[]",
                CAMPOS_JSON: "[]"
            };
            close();
            this.editingIsNew = true;
            this.openEditor(draft);
        });
    },

    // ================================================================
    // DISEÑADOR (modal a pantalla completa, 2 bloques)
    // ================================================================
    async openEditor(record) {
        this.editing = record;
        this.editingIsNew = !this.list.some(r => r[this.ID_COL] === record[this.ID_COL]);
        this.screen = this.parseScreen(record.VARIABLES_JSON);
        const savedFields = this.safeParse(record.CAMPOS_JSON, []);

        await this.loadDimensions();

        let physicalCols = [];
        try {
            physicalCols = await Provider.listColumns(this.project.DATASET, record.TABLA);
        } catch (err) {
            UI.toast("No se pudieron leer las columnas de la tabla: " + err.message, "error");
        }

        // Fusiona la config guardada con las columnas físicas actuales:
        // columnas nuevas entran con valores por defecto; columnas que ya
        // no existen en la tabla se descartan.
        const byName = new Map(savedFields.map(f => [f.name, f]));
        this.fields = physicalCols.map((c, idx) => {
            const saved = byName.get(c.name);
            return saved || this.blankFieldConfig(c.name, idx);
        });
        // Respeta el orden guardado si todas las columnas siguen existiendo.
        this.fields.sort((a, b) => (a.order ?? 0) - (b.order ?? 0));

        let overlay = document.getElementById("actUpdEditorModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "actUpdEditorModal";
            document.body.appendChild(overlay);
        }
        overlay.classList.add("visible");
        this.overlay = overlay;

        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div class="modal-title-edit-wrap">
                        <h3 id="actUpdModalTitle" class="modal-title-editable" contenteditable="true" spellcheck="false" title="Clic para renombrar">${UI.escapeHtml(record[this.NAME_COL])}</h3>
                        <span class="modal-subtitle">Tabla: ${UI.escapeHtml(record.TABLA)}</span>
                    </div>
                    <div class="modal-header-right">
                        <button class="modal-close" id="actUpdEditorClose">&times;</button>
                    </div>
                </div>
                <div class="modal-body modal-body-flush">
                    <div id="actUpdScreenPart" class="flow-part flow-part--screen"></div>
                    <div id="actUpdFieldsPart" class="flow-part"></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="actUpdEditorCancel">Cancelar</button>
                    <button class="btn btn-secondary" id="actUpdEditorRun" ${this.editingIsNew ? "disabled" : ""}>▶ Ejecutar</button>
                    <button class="btn btn-primary" id="actUpdEditorSave">Guardar</button>
                </div>
            </div>`;

        document.getElementById("actUpdEditorClose").addEventListener("click", () => overlay.remove());
        document.getElementById("actUpdEditorCancel").addEventListener("click", () => overlay.remove());
        document.getElementById("actUpdEditorSave").addEventListener("click", () => this.save());
        document.getElementById("actUpdEditorRun").addEventListener("click", () => {
            if (!this.editingIsNew) this.startRun(this.editing);
        });

        this.renderScreenBlock();
        this.renderFieldsBlock();
    },

    blankFieldConfig(name, order) {
        return {
            name,
            description: "",
            order,
            filter: { type: "NONE", value: "" },
            validation: {
                type: "NONE",
                constants: [],
                dimensionId: "", dimensionName: "",
                hierarchyName: "", level: 1, node: "",
                allowEmpty: true,
                showText: false,
                searchHelp: "LISTBOX"
            }
        };
    },

    async loadDimensions() {
        const sql = `SELECT DIMENSION_ID, DIMENSION, TABLA, CAMPOS_JSON
                     FROM ${Provider.qualifyControl("DIMENSIONES")}
                     WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                     ORDER BY DIMENSION`;
        this.dimensionsCache = await Provider.runQuery(sql);
    },

    // ---------------- Bloque A: pantalla de entrada de variables ----------------
    // Mismo diseñador que en Flujos de carga (Flows.renderScreenBlock): variables
    // sueltas, frames, textos, espacios y líneas, con alta/edición de cada
    // variable vía el modal compartido UI.openScreenVariableModal (nombre
    // técnico, etiqueta, tipo y modo de selección).
    renderScreenBlock() {
        const part = document.getElementById("actUpdScreenPart");
        const screen = this.screen;

        part.innerHTML = `
            <div class="flow-part-header flow-part-header--screen">
                <button type="button" class="flow-part-toggle" id="actUpdScreenToggle">
                    <span class="flow-group-caret ${this.screenCollapsed ? "is-collapsed" : ""}">▾</span>
                    <span>Pantalla de entrada de variables</span>
                </button>
                <div class="flow-screen-toolbar-mini">
                    <button type="button" class="flow-mini-btn" id="actUpdAddVar" title="Añadir variable">+ Var</button>
                    <button type="button" class="flow-mini-btn" id="actUpdAddFrame" title="Añadir frame">+ Frame</button>
                    <button type="button" class="flow-mini-btn" id="actUpdAddText" title="Añadir texto">+ Texto</button>
                    <button type="button" class="flow-mini-btn" id="actUpdAddSkip" title="Añadir espacio en blanco">+ Espacio</button>
                    <button type="button" class="flow-mini-btn" id="actUpdAddLine" title="Añadir línea separadora">+ Línea</button>
                </div>
            </div>
            <div class="flow-screen-box ${this.screenCollapsed ? "is-collapsed" : ""}" id="actUpdScreenBox">
                <div class="form-group">
                    <label>Título de la pantalla</label>
                    <input type="text" id="actUpdScreenTitle" placeholder="Ej. Filtro de actualización" value="${UI.escapeHtml(screen.title || "")}">
                </div>
                <p class="form-hint">Variables para filtrar qué filas se traen a editar; se pueden usar luego como "filtro variable" de un campo.</p>
                <div class="flow-screen-blocks" id="actUpdScreenBlocks"></div>
            </div>`;

        document.getElementById("actUpdScreenToggle").addEventListener("click", () => {
            this.screenCollapsed = !this.screenCollapsed;
            const box = document.getElementById("actUpdScreenBox");
            box.classList.toggle("is-collapsed", this.screenCollapsed);
            part.querySelector(".flow-group-caret").classList.toggle("is-collapsed", this.screenCollapsed);
        });

        document.getElementById("actUpdScreenTitle").addEventListener("input", (e) => { screen.title = e.target.value; });

        document.getElementById("actUpdAddVar").addEventListener("click", async () => {
            const v = await UI.openScreenVariableModal({});
            if (!v) return;
            screen.blocks.push({ id: Provider.newId(), kind: "variable", variable: { id: Provider.newId(), ...v } });
            this.screenCollapsed = false;
            this.renderScreenBlocksList();
        });
        document.getElementById("actUpdAddFrame").addEventListener("click", () => {
            screen.blocks.push({ id: Provider.newId(), kind: "frame", title: "Nuevo frame", variables: [] });
            this.screenCollapsed = false;
            this.renderScreenBlocksList();
        });
        document.getElementById("actUpdAddText").addEventListener("click", async () => {
            const text = await UI.openScreenTextModal({ current: "" });
            if (text === null) return;
            screen.blocks.push({ id: Provider.newId(), kind: "text", text });
            this.renderScreenBlocksList();
        });
        document.getElementById("actUpdAddSkip").addEventListener("click", () => {
            screen.blocks.push({ id: Provider.newId(), kind: "skip" });
            this.renderScreenBlocksList();
        });
        document.getElementById("actUpdAddLine").addEventListener("click", () => {
            screen.blocks.push({ id: Provider.newId(), kind: "line" });
            this.renderScreenBlocksList();
        });

        this.renderScreenBlocksList();
    },

    renderScreenBlocksList() {
        const wrap = document.getElementById("actUpdScreenBlocks");
        const screen = this.screen;
        if (!wrap) return;

        if (!screen.blocks.length) {
            wrap.innerHTML = `<div class="module-empty module-empty--inline">Sin variables: la tabla se cargará entera al ejecutar (salvo filtros constantes por campo).</div>`;
            return;
        }

        wrap.innerHTML = screen.blocks.map((b, idx) => this.screenBlockHtml(b, idx)).join("");
        this.bindScreenBlocksEvents();
    },

    screenBlockHtml(b, idx) {
        if (b.kind === "variable") {
            return `
                <div class="flow-screen-block flow-screen-block--var" draggable="true" data-block-idx="${idx}">
                    <span class="load-drag-handle">⠿</span>
                    <div class="flow-field-preview-click" data-edit-var="${idx}" title="Clic para configurar la variable">
                        ${this.fieldPreviewHtml(b.variable)}
                    </div>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar">✕</button>
                </div>`;
        }
        if (b.kind === "text") {
            return `
                <div class="flow-screen-block flow-screen-block--text" draggable="true" data-block-idx="${idx}">
                    <span class="load-drag-handle">⠿</span>
                    <div class="flow-screen-text-content" data-edit-text="${idx}" title="Clic para editar">${UI.renderFormattedText(b.text)}</div>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar">✕</button>
                </div>`;
        }
        if (b.kind === "skip") {
            return `
                <div class="flow-screen-block flow-screen-block--skip" draggable="true" data-block-idx="${idx}">
                    <span class="load-drag-handle">⠿</span>
                    <div class="flow-screen-skip-marker">· · · espacio en blanco · · ·</div>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar">✕</button>
                </div>`;
        }
        if (b.kind === "line") {
            return `
                <div class="flow-screen-block flow-screen-block--line" draggable="true" data-block-idx="${idx}">
                    <span class="load-drag-handle">⠿</span>
                    <div class="flow-screen-line-marker"><hr></div>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar">✕</button>
                </div>`;
        }
        // frame
        return `
            <div class="flow-screen-block flow-screen-block--frame" draggable="true" data-block-idx="${idx}">
                <div class="flow-frame-header">
                    <span class="load-drag-handle">⠿</span>
                    <strong data-edit-frame-title="${idx}" title="Clic para renombrar">${UI.escapeHtml(b.title || "Frame")}</strong>
                    <span class="load-fn-toolbar-spacer"></span>
                    <button type="button" class="flow-mini-btn" data-add-frame-var="${idx}">+ Var</button>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar frame">✕</button>
                </div>
                <div class="flow-frame-vars" data-frame-idx="${idx}">
                    ${b.variables.length ? b.variables.map((v, vi) => `
                        <div class="flow-frame-var-row" draggable="true" data-frame-idx="${idx}" data-var-idx="${vi}">
                            <span class="load-drag-handle">⠿</span>
                            <div class="flow-field-preview-click" data-edit-frame-var="${idx}:${vi}" title="Clic para configurar la variable">
                                ${this.fieldPreviewHtml(v)}
                            </div>
                            <button type="button" class="field-remove" data-remove-frame-var="${idx}:${vi}" title="Eliminar">✕</button>
                        </div>`).join("") : `<div class="hierarchy-pool-empty">Sin variables en este frame.</div>`}
                </div>
            </div>`;
    },

    /** Previsualización de una variable de pantalla tal y como se vería de verdad: etiqueta + input, sin recuadro alrededor. */
    fieldPreviewHtml(v) {
        const modeLabels = { rango: "Rango", multiple: "Varios valores", cualquiera: "Select-options" };
        const modeBadge = v.selectMode && modeLabels[v.selectMode] ? `<span class="flow-var-mode-badge">${modeLabels[v.selectMode]}</span>` : "";
        return `
            <div class="flow-field-preview">
                <label>${UI.escapeHtml(v.label || v.name)}${modeBadge}</label>
                <input type="text" disabled placeholder="${UI.escapeHtml(v.name)}">
            </div>`;
    },

    bindScreenBlocksEvents() {
        const wrap = document.getElementById("actUpdScreenBlocks");
        const screen = this.screen;

        wrap.querySelectorAll("[data-remove-block]").forEach(btn => btn.addEventListener("click", () => {
            const idx = parseInt(btn.dataset.removeBlock, 10);
            screen.blocks.splice(idx, 1);
            this.renderScreenBlocksList();
        }));

        wrap.querySelectorAll("[data-add-frame-var]").forEach(btn => btn.addEventListener("click", async () => {
            const idx = parseInt(btn.dataset.addFrameVar, 10);
            const v = await UI.openScreenVariableModal({});
            if (!v) return;
            screen.blocks[idx].variables.push({ id: Provider.newId(), ...v });
            this.renderScreenBlocksList();
        }));

        wrap.querySelectorAll("[data-remove-frame-var]").forEach(btn => btn.addEventListener("click", () => {
            const [bIdx, vIdx] = btn.dataset.removeFrameVar.split(":").map(Number);
            screen.blocks[bIdx].variables.splice(vIdx, 1);
            this.renderScreenBlocksList();
        }));

        wrap.querySelectorAll("[data-edit-frame-title]").forEach(el => el.addEventListener("click", async () => {
            const idx = parseInt(el.dataset.editFrameTitle, 10);
            const val = await UI.openTextPromptModal({ title: "Nombre del frame", label: "Título", value: screen.blocks[idx].title || "" });
            if (val === null) return;
            screen.blocks[idx].title = val || "Frame";
            this.renderScreenBlocksList();
        }));

        wrap.querySelectorAll("[data-edit-text]").forEach(el => el.addEventListener("click", async () => {
            const idx = parseInt(el.dataset.editText, 10);
            const val = await UI.openScreenTextModal({ current: screen.blocks[idx].text || "" });
            if (val === null) return;
            screen.blocks[idx].text = val;
            this.renderScreenBlocksList();
        }));

        wrap.querySelectorAll("[data-edit-var]").forEach(el => el.addEventListener("click", async (e) => {
            e.stopPropagation();
            const idx = parseInt(el.dataset.editVar, 10);
            const current = screen.blocks[idx].variable;
            const v = await UI.openScreenVariableModal({ current });
            if (!v) return;
            screen.blocks[idx].variable = { ...current, ...v };
            this.renderScreenBlocksList();
        }));

        wrap.querySelectorAll("[data-edit-frame-var]").forEach(el => el.addEventListener("click", async (e) => {
            e.stopPropagation();
            const [bIdx, vIdx] = el.dataset.editFrameVar.split(":").map(Number);
            const current = screen.blocks[bIdx].variables[vIdx];
            const v = await UI.openScreenVariableModal({ current });
            if (!v) return;
            screen.blocks[bIdx].variables[vIdx] = { ...current, ...v };
            this.renderScreenBlocksList();
        }));

        // Reordenar bloques de primer nivel arrastrando.
        wrap.querySelectorAll(":scope > .flow-screen-block").forEach(block => {
            block.addEventListener("dragstart", (e) => {
                e.stopPropagation();
                this.dragBlockIdx = parseInt(block.dataset.blockIdx, 10);
                block.classList.add("dragging");
            });
            block.addEventListener("dragend", () => block.classList.remove("dragging"));
            block.addEventListener("dragover", (e) => e.preventDefault());
            block.addEventListener("drop", (e) => {
                e.preventDefault();
                e.stopPropagation();
                const targetIdx = parseInt(block.dataset.blockIdx, 10);
                if (this.dragBlockIdx === null || this.dragBlockIdx === targetIdx) return;
                const [moved] = screen.blocks.splice(this.dragBlockIdx, 1);
                screen.blocks.splice(targetIdx, 0, moved);
                this.dragBlockIdx = null;
                this.renderScreenBlocksList();
            });
        });

        // Reordenar variables dentro de un mismo frame arrastrando.
        wrap.querySelectorAll(".flow-frame-var-row").forEach(row => {
            row.addEventListener("dragstart", (e) => {
                e.stopPropagation();
                this.dragFrameVar = { frameIdx: parseInt(row.dataset.frameIdx, 10), varIdx: parseInt(row.dataset.varIdx, 10) };
                row.classList.add("dragging");
            });
            row.addEventListener("dragend", () => row.classList.remove("dragging"));
            row.addEventListener("dragover", (e) => { e.preventDefault(); e.stopPropagation(); });
            row.addEventListener("drop", (e) => {
                e.preventDefault();
                e.stopPropagation();
                if (!this.dragFrameVar) return;
                const frameIdx = parseInt(row.dataset.frameIdx, 10);
                const targetVarIdx = parseInt(row.dataset.varIdx, 10);
                if (this.dragFrameVar.frameIdx !== frameIdx) { this.dragFrameVar = null; return; }
                const vars = screen.blocks[frameIdx].variables;
                const [moved] = vars.splice(this.dragFrameVar.varIdx, 1);
                vars.splice(targetVarIdx, 0, moved);
                this.dragFrameVar = null;
                this.renderScreenBlocksList();
            });
        });
    },

    // ---------------- Bloque B: campos de la tabla ----------------
    renderFieldsBlock() {
        const part = document.getElementById("actUpdFieldsPart");
        part.innerHTML = `
            <div class="flow-screen-header"><strong>Campos</strong>
                <span class="flow-screen-hint">Arrastra para ordenar · el orden aquí es el orden de columnas al ejecutar</span>
            </div>
            <div class="actupd-fields-header">
                <span></span><span>Campo</span><span>Descripción</span><span>Filtro</span><span>Validación</span>
            </div>
            <div id="actUpdFieldsRows" class="actupd-fields-rows"></div>`;
        this.renderFieldRows();
    },

    renderFieldRows() {
        const rowsEl = document.getElementById("actUpdFieldsRows");
        if (!rowsEl) return;
        if (!this.fields.length) {
            rowsEl.innerHTML = `<div class="hierarchy-levels-empty">Esta tabla no tiene columnas (o no se pudieron leer).</div>`;
            return;
        }

        rowsEl.innerHTML = this.fields.map((f, idx) => {
            const filterActive = f.filter && f.filter.type !== "NONE";
            const validActive = f.validation && f.validation.type !== "NONE";
            return `
            <div class="actupd-field-row" draggable="true" data-idx="${idx}">
                <span class="actupd-drag-handle" title="Arrastrar">⠿</span>
                <span class="actupd-field-name">${UI.escapeHtml(f.name)}</span>
                <input type="text" class="actupd-field-desc" data-idx="${idx}" placeholder="Descripción del campo" value="${UI.escapeHtml(f.description || "")}">
                <button type="button" class="actupd-icon-btn ${filterActive ? "active" : ""}" data-filter-idx="${idx}" title="Filtro">⏷</button>
                <button type="button" class="actupd-icon-btn ${validActive ? "active" : ""}" data-valid-idx="${idx}" title="Validación">✓</button>
            </div>`;
        }).join("");

        rowsEl.querySelectorAll(".actupd-field-desc").forEach(inp => {
            inp.addEventListener("input", (e) => { this.fields[parseInt(e.target.dataset.idx, 10)].description = e.target.value; });
        });
        rowsEl.querySelectorAll("[data-filter-idx]").forEach(btn => {
            btn.addEventListener("click", () => this.openFilterPopover(parseInt(btn.dataset.filterIdx, 10), btn));
        });
        rowsEl.querySelectorAll("[data-valid-idx]").forEach(btn => {
            btn.addEventListener("click", () => this.openValidationModal(parseInt(btn.dataset.validIdx, 10)));
        });

        rowsEl.querySelectorAll(".actupd-field-row").forEach(row => {
            const idx = parseInt(row.dataset.idx, 10);
            row.addEventListener("dragstart", () => { this.dragFieldIdx = idx; row.classList.add("dragging"); });
            row.addEventListener("dragend", () => row.classList.remove("dragging"));
            row.addEventListener("dragover", (e) => e.preventDefault());
            row.addEventListener("drop", (e) => {
                e.preventDefault();
                if (this.dragFieldIdx === null || this.dragFieldIdx === idx) return;
                const [moved] = this.fields.splice(this.dragFieldIdx, 1);
                this.fields.splice(idx, 0, moved);
                this.fields.forEach((f, i) => { f.order = i; });
                this.dragFieldIdx = null;
                this.renderFieldRows();
            });
        });
    },

    // ---- popover de filtro (ninguno / constante / variable) ----
    openFilterPopover(idx, anchorBtn) {
        document.querySelectorAll(".actupd-popover").forEach(p => p.remove());
        const f = this.fields[idx];
        const pop = document.createElement("div");
        pop.className = "actupd-popover";
        pop.innerHTML = `
            <div class="actupd-popover-title">Filtro: ${UI.escapeHtml(f.name)}</div>
            <label><input type="radio" name="filterType" value="NONE" ${f.filter.type === "NONE" ? "checked" : ""}> Ninguno</label>
            <label><input type="radio" name="filterType" value="CONST" ${f.filter.type === "CONST" ? "checked" : ""}> Constante</label>
            <input type="text" id="actUpdFilterConstVal" placeholder="Valor" style="display:${f.filter.type === "CONST" ? "block" : "none"}" value="${f.filter.type === "CONST" ? UI.escapeHtml(f.filter.value) : ""}">
            <label><input type="radio" name="filterType" value="VAR" ${f.filter.type === "VAR" ? "checked" : ""}> Variable</label>
            <select id="actUpdFilterVarSel" style="display:${f.filter.type === "VAR" ? "block" : "none"}">
                <option value="">Selecciona...</option>
                ${this.flatVars(this.screen).map(v => `<option value="${v.id}" ${f.filter.value === v.id ? "selected" : ""}>${UI.escapeHtml(v.label || v.name)}</option>`).join("")}
            </select>
            <div class="actupd-popover-actions">
                <button class="btn btn-primary btn-sm" id="actUpdFilterAccept">Aceptar</button>
            </div>`;
        document.body.appendChild(pop);
        const r = anchorBtn.getBoundingClientRect();
        pop.style.top = `${r.bottom + window.scrollY + 4}px`;
        pop.style.left = `${r.left + window.scrollX}px`;

        pop.querySelectorAll('input[name="filterType"]').forEach(radio => {
            radio.addEventListener("change", () => {
                pop.querySelector("#actUpdFilterConstVal").style.display = radio.value === "CONST" && radio.checked ? "block" : "none";
                pop.querySelector("#actUpdFilterVarSel").style.display = radio.value === "VAR" && radio.checked ? "block" : "none";
            });
        });

        const closeOnOutside = (e) => {
            if (!pop.contains(e.target) && e.target !== anchorBtn) {
                pop.remove();
                document.removeEventListener("click", closeOnOutside, true);
            }
        };
        setTimeout(() => document.addEventListener("click", closeOnOutside, true), 0);

        pop.querySelector("#actUpdFilterAccept").addEventListener("click", () => {
            const type = pop.querySelector('input[name="filterType"]:checked').value;
            f.filter = {
                type,
                value: type === "CONST" ? pop.querySelector("#actUpdFilterConstVal").value : (type === "VAR" ? pop.querySelector("#actUpdFilterVarSel").value : "")
            };
            pop.remove();
            document.removeEventListener("click", closeOnOutside, true);
            this.renderFieldRows();
        });
    },

    // ---- modal de validación (ninguna / constante / dimensión / jerarquía) ----
    openValidationModal(idx) {
        const f = this.fields[idx];
        const v = f.validation;

        let overlay = document.getElementById("actUpdValidModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "actUpdValidModal";
            document.body.appendChild(overlay);
        }
        overlay.classList.add("visible");

        const dimOptions = this.dimensionsCache.map(d => `<option value="${d.DIMENSION_ID}" ${v.dimensionId === d.DIMENSION_ID ? "selected" : ""}>${UI.escapeHtml(d.DIMENSION)}</option>`).join("");

        overlay.innerHTML = `
            <div class="modal-box">
                <div class="modal-header">
                    <h3>Validación: ${UI.escapeHtml(f.name)}</h3>
                    <button class="modal-close" id="actUpdValidClose">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="form-group">
                        <label><input type="radio" name="validType" value="NONE" ${v.type === "NONE" ? "checked" : ""}> Ninguna</label>
                        <label><input type="radio" name="validType" value="CONST" ${v.type === "CONST" ? "checked" : ""}> Constante</label>
                        <label><input type="radio" name="validType" value="DIM" ${v.type === "DIM" ? "checked" : ""}> Valores de dimensión</label>
                        <label><input type="radio" name="validType" value="HIER" ${v.type === "HIER" ? "checked" : ""}> Valores de jerarquía</label>
                    </div>

                    <div id="actUpdValidConst" style="display:${v.type === "CONST" ? "block" : "none"}">
                        <table class="const-table">
                            <thead><tr><th>ID</th><th>Descripción</th><th></th></tr></thead>
                            <tbody id="actUpdConstRows"></tbody>
                        </table>
                        <button class="btn btn-secondary btn-sm" id="actUpdAddConst">+ Añadir valor</button>
                    </div>

                    <div id="actUpdValidDim" style="display:${v.type === "DIM" ? "block" : "none"}">
                        <div class="form-group">
                            <label>Dimensión</label>
                            <select id="actUpdDimSelect"><option value="">Selecciona...</option>${dimOptions}</select>
                        </div>
                    </div>

                    <div id="actUpdValidHier" style="display:${v.type === "HIER" ? "block" : "none"}">
                        <div class="form-group">
                            <label>Dimensión</label>
                            <select id="actUpdHierDimSelect"><option value="">Selecciona...</option>${dimOptions}</select>
                        </div>
                        <div class="form-group">
                            <label>Jerarquía</label>
                            <select id="actUpdHierSelect"><option value="">Selecciona una dimensión primero</option></select>
                        </div>
                        <div class="form-group">
                            <label>Nivel</label>
                            <select id="actUpdHierLevelSelect"><option value="">—</option></select>
                        </div>
                        <div class="form-group">
                            <label>Valor del nodo (se traerán los valores por debajo de este nodo)</label>
                            <input type="text" id="actUpdHierNode" value="${UI.escapeHtml(v.node || "")}">
                        </div>
                    </div>

                    <hr>
                    <div class="form-group">
                        <label><input type="checkbox" id="actUpdAllowEmpty" ${v.allowEmpty ? "checked" : ""}> Permite valor vacío</label>
                    </div>
                    <div class="form-group">
                        <label><input type="checkbox" id="actUpdShowText" ${v.showText ? "checked" : ""}> Mostrar texto descriptivo junto al valor</label>
                    </div>
                    <div class="form-group">
                        <label>Ayuda de búsqueda</label>
                        <select id="actUpdSearchHelp">
                            <option value="LISTBOX" ${v.searchHelp === "LISTBOX" ? "selected" : ""}>Listbox (desplegable)</option>
                            <option value="SEARCH" ${v.searchHelp === "SEARCH" ? "selected" : ""}>Buscador (estilo SAP, con caja de búsqueda)</option>
                            <option value="CHECKBOX" ${v.searchHelp === "CHECKBOX" ? "selected" : ""}>Checkbox (solo si hay exactamente 2 valores posibles)</option>
                        </select>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="actUpdValidCancel">Cancelar</button>
                    <button class="btn btn-primary" id="actUpdValidAccept">Aceptar</button>
                </div>
            </div>`;

        const close = () => overlay.remove();
        document.getElementById("actUpdValidClose").addEventListener("click", close);
        document.getElementById("actUpdValidCancel").addEventListener("click", close);

        overlay.querySelectorAll('input[name="validType"]').forEach(radio => {
            radio.addEventListener("change", () => {
                overlay.querySelector("#actUpdValidConst").style.display = radio.value === "CONST" && radio.checked ? "block" : "none";
                overlay.querySelector("#actUpdValidDim").style.display = radio.value === "DIM" && radio.checked ? "block" : "none";
                overlay.querySelector("#actUpdValidHier").style.display = radio.value === "HIER" && radio.checked ? "block" : "none";
            });
        });

        // -- constante: tabla ID/Descripción editable --
        let constants = (v.constants || []).slice();
        const renderConstRows = () => {
            const tbody = overlay.querySelector("#actUpdConstRows");
            tbody.innerHTML = constants.map((c, i) => `
                <tr>
                    <td><input type="text" class="const-id" data-i="${i}" value="${UI.escapeHtml(c.id)}"></td>
                    <td><input type="text" class="const-desc" data-i="${i}" value="${UI.escapeHtml(c.desc)}"></td>
                    <td><button type="button" class="field-remove" data-rm="${i}">✕</button></td>
                </tr>`).join("");
            tbody.querySelectorAll(".const-id").forEach(inp => inp.addEventListener("input", e => constants[e.target.dataset.i].id = e.target.value));
            tbody.querySelectorAll(".const-desc").forEach(inp => inp.addEventListener("input", e => constants[e.target.dataset.i].desc = e.target.value));
            tbody.querySelectorAll("[data-rm]").forEach(btn => btn.addEventListener("click", () => { constants.splice(parseInt(btn.dataset.rm, 10), 1); renderConstRows(); }));
        };
        renderConstRows();
        overlay.querySelector("#actUpdAddConst").addEventListener("click", () => { constants.push({ id: "", desc: "" }); renderConstRows(); });

        // -- jerarquía: dependencias dimensión -> jerarquía -> nivel --
        const hierDimSel = overlay.querySelector("#actUpdHierDimSelect");
        const hierSel = overlay.querySelector("#actUpdHierSelect");
        const hierLevelSel = overlay.querySelector("#actUpdHierLevelSelect");
        let hierLevelsCache = [];

        const loadHierarchiesFor = async (dimId, preselectHier, preselectLevel) => {
            hierSel.innerHTML = `<option value="">Cargando...</option>`;
            try {
                const rows = await Provider.runQuery(`SELECT JERARQUIA, NIVELES_JSON FROM ${Provider.qualifyControl("JERARQUIAS")} WHERE DIMENSION_ID = '${Provider.esc(dimId)}' ORDER BY JERARQUIA`);
                if (!rows.length) { hierSel.innerHTML = `<option value="">Esta dimensión no tiene jerarquías</option>`; return; }
                hierSel.innerHTML = `<option value="">Selecciona...</option>` + rows.map(r => `<option value="${UI.escapeHtml(r.JERARQUIA)}" ${preselectHier === r.JERARQUIA ? "selected" : ""}>${UI.escapeHtml(r.JERARQUIA)}</option>`).join("");
                hierSel.dataset.rows = JSON.stringify(rows);
                if (preselectHier) fillLevels(preselectHier, preselectLevel);
            } catch (err) {
                hierSel.innerHTML = `<option value="">Error al cargar jerarquías</option>`;
            }
        };

        const fillLevels = (hierName, preselectLevel) => {
            const rows = JSON.parse(hierSel.dataset.rows || "[]");
            const row = rows.find(r => r.JERARQUIA === hierName);
            hierLevelsCache = row ? this.safeParse(row.NIVELES_JSON, []) : [];
            hierLevelSel.innerHTML = hierLevelsCache.map((colId, i) => `<option value="${i + 1}" ${preselectLevel == i + 1 ? "selected" : ""}>Nivel ${i + 1}: ${UI.escapeHtml(colId)}</option>`).join("") || `<option value="">—</option>`;
        };

        if (v.type === "HIER" && v.dimensionId) {
            loadHierarchiesFor(v.dimensionId, v.hierarchyName, v.level);
        }
        hierDimSel.addEventListener("change", () => loadHierarchiesFor(hierDimSel.value));
        hierSel.addEventListener("change", () => fillLevels(hierSel.value));

        document.getElementById("actUpdValidAccept").addEventListener("click", () => {
            const type = overlay.querySelector('input[name="validType"]:checked').value;
            const searchHelp = overlay.querySelector("#actUpdSearchHelp").value;

            f.validation = {
                type,
                constants: type === "CONST" ? constants.filter(c => c.id) : [],
                dimensionId: type === "DIM" ? overlay.querySelector("#actUpdDimSelect").value : (type === "HIER" ? hierDimSel.value : ""),
                dimensionName: type === "DIM"
                    ? (this.dimensionsCache.find(d => d.DIMENSION_ID === overlay.querySelector("#actUpdDimSelect").value) || {}).DIMENSION || ""
                    : (type === "HIER" ? (this.dimensionsCache.find(d => d.DIMENSION_ID === hierDimSel.value) || {}).DIMENSION || "" : ""),
                hierarchyName: type === "HIER" ? hierSel.value : "",
                level: type === "HIER" ? parseInt(hierLevelSel.value || "1", 10) : 1,
                node: type === "HIER" ? overlay.querySelector("#actUpdHierNode").value.trim() : "",
                allowEmpty: overlay.querySelector("#actUpdAllowEmpty").checked,
                showText: overlay.querySelector("#actUpdShowText").checked,
                searchHelp
            };
            close();
            this.renderFieldRows();
        });
    },

    async save() {
        const nameEl = document.getElementById("actUpdModalTitle");
        const name = nameEl ? nameEl.textContent.trim() : this.editing[this.NAME_COL];
        if (!name) { UI.toast("El nombre no puede estar vacío.", "error"); return; }

        this.fields.forEach((f, i) => { f.order = i; });
        const varsJsonPlain = JSON.stringify(this.screen);
        const fieldsJsonPlain = JSON.stringify(this.fields);
        const varsJson = varsJsonPlain.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
        const fieldsJson = fieldsJsonPlain.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
        const id = this.editing[this.ID_COL];

        try {
            if (this.editingIsNew) {
                await Provider.runQuery(`
                    INSERT INTO ${Provider.qualifyControl(this.TABLE)}
                    (${this.ID_COL}, PROYECTO_ID, ${this.NAME_COL}, DESCRIPCION, TABLA, VARIABLES_JSON, CAMPOS_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                    VALUES ('${Provider.esc(id)}', '${Provider.esc(this.project.PROYECTO_ID)}', '${Provider.esc(name)}', '${Provider.esc(this.editing.DESCRIPCION || "")}',
                            '${Provider.esc(this.editing.TABLA)}', '${varsJson}', '${fieldsJson}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`);
                this.editingIsNew = false;
                UI.toast(`Actualización "${name}" creada.`, "success");
            } else {
                await Provider.runQuery(`
                    UPDATE ${Provider.qualifyControl(this.TABLE)}
                    SET ${this.NAME_COL} = '${Provider.esc(name)}', VARIABLES_JSON = '${varsJson}', CAMPOS_JSON = '${fieldsJson}', FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                    WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
                UI.toast(`Actualización "${name}" guardada.`, "success");
            }
            // Refresca el registro en memoria (VARIABLES_JSON/CAMPOS_JSON) con lo
            // que se acaba de grabar, para que "▶ Ejecutar" use los cambios
            // recién guardados sin tener que cerrar y reabrir desde la lista.
            this.editing[this.NAME_COL] = name;
            this.editing.VARIABLES_JSON = varsJsonPlain;
            this.editing.CAMPOS_JSON = fieldsJsonPlain;
            await this.loadList();
            const runBtn = document.getElementById("actUpdEditorRun");
            if (runBtn) runBtn.disabled = false;
        } catch (err) {
            UI.toast("Error al guardar: " + err.message, "error");
        }
    },

    // ================================================================
    // EJECUCIÓN — mismo patrón que flow_run.html (Flujos de carga): una
    // pantalla con dos pestañas, "Pantalla de variables" y, en lugar de
    // "Monitor", "Tabla" (el grid con los datos que trae esa pantalla de
    // variables). Un único modal a pantalla completa; switchRunTab()
    // alterna el contenido de #actUpdRunBody entre ambas vistas.
    // ================================================================
    async startRun(record) {
        if (this.overlay) this.overlay.remove();
        this.runRecord = record;
        this.runScreen = this.parseScreen(record.VARIABLES_JSON);
        this.runFields = this.safeParse(record.CAMPOS_JSON, []).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
        this.gridState = null;
        this.selOptState = {};
        // Al ejecutar directamente desde el listado (sin pasar por "Editar")
        // this.dimensionsCache nunca se ha cargado, así que las validaciones de
        // tipo dimensión/jerarquía no encontraban nada. Se asegura aquí, igual
        // que hace openEditor().
        await this.loadDimensions();
        // Si la pantalla no tiene variables definidas (ni sueltas ni dentro de
        // frames), no tiene sentido mostrar esa pestaña ni pedir al usuario que
        // pulse "Cargar tabla": se muestra únicamente la pestaña Tabla y se
        // carga directamente al abrir. OJO: blocks.length no vale por sí solo,
        // puede haber frames/textos/saltos de línea sin ninguna variable real.
        this.hasScreenVars = this.flatVars(this.runScreen).length > 0;

        let overlay = document.getElementById("actUpdRunModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "actUpdRunModal";
            document.body.appendChild(overlay);
        }
        this.runOverlay = overlay;
        overlay.classList.add("visible");
        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div>
                        <h3>${UI.escapeHtml(this.runScreen.title || record[this.NAME_COL])}</h3>
                        <span class="modal-subtitle">${UI.escapeHtml(record[this.NAME_COL])} · Tabla ${UI.escapeHtml(record.TABLA)}</span>
                    </div>
                    <button class="modal-close" id="actUpdRunClose">&times;</button>
                </div>
                <div class="flow-run-tabs" id="actUpdRunTabs">
                    ${this.hasScreenVars ? `<button type="button" class="flow-run-tab active" id="actUpdRunTabScreen">Pantalla de variables</button>` : ""}
                    <button type="button" class="flow-run-tab ${this.hasScreenVars ? "" : "active"}" id="actUpdRunTabTable">📋 Tabla</button>
                </div>
                <div class="modal-body modal-body-flush" id="actUpdRunBody"></div>
            </div>`;

        document.getElementById("actUpdRunClose").addEventListener("click", () => overlay.remove());
        const screenTabBtn = document.getElementById("actUpdRunTabScreen");
        if (screenTabBtn) screenTabBtn.addEventListener("click", () => this.switchRunTab("screen"));
        document.getElementById("actUpdRunTabTable").addEventListener("click", () => this.switchRunTab("table"));

        if (this.hasScreenVars) {
            this.switchRunTab("screen");
        } else {
            this.runView = "table";
            this.autoLoadRunTable();
        }
    },

    switchRunTab(tab) {
        this.runView = tab;
        const screenTabBtn = document.getElementById("actUpdRunTabScreen");
        if (screenTabBtn) screenTabBtn.classList.toggle("active", tab === "screen");
        document.getElementById("actUpdRunTabTable").classList.toggle("active", tab === "table");
        if (tab === "screen") this.renderRunScreenView();
        else this.renderRunTableView();
    },

    /** Carga la tabla directamente al abrir, sin pasar por la pantalla de
     *  variables (usado cuando la actualización no tiene variables definidas). */
    async autoLoadRunTable() {
        const body = document.getElementById("actUpdRunBody");
        if (body) body.innerHTML = `<div class="module-empty"><span class="spinner"></span> Cargando datos...</div>`;
        try {
            const where = this.buildWhere(this.runFields, {});
            const table = Provider.qualify(this.project.DATASET, this.runRecord.TABLA);
            const rows = await Provider.runQuery(`SELECT * FROM ${table} ${where}`);
            await this.buildGridState(this.runRecord, this.runFields, rows, where);
            this.renderRunTableView();
        } catch (err) {
            UI.toast("Error al cargar la tabla: " + err.message, "error");
            if (body) body.innerHTML = `<div class="module-empty">Error al cargar la tabla.</div>`;
        }
    },

    // ---------------- Pestaña 1: pantalla de variables ----------------
    renderRunScreenView() {
        const body = document.getElementById("actUpdRunBody");
        const blocksHtml = this.runScreen.blocks.length
            ? this.runScreen.blocks.map(b => this.runBlockHtml(b)).join("")
            : `<div class="module-empty module-empty--inline">Esta actualización no tiene variables de pantalla definidas.</div>`;

        body.innerHTML = `
            <div class="flow-run-screen">
                <div class="flow-screen-blocks flow-screen-blocks--run">${blocksHtml}</div>
                <div class="flow-run-actions">
                    <button class="btn btn-primary" id="actUpdRunExecute">▶ Cargar tabla</button>
                    <span class="form-hint" id="actUpdRunHint"></span>
                </div>
            </div>`;

        document.getElementById("actUpdRunExecute").addEventListener("click", () => this.executeRun());
        this.bindSelOptDelegation(body);
    },

    runBlockHtml(b) {
        if (b.kind === "text") {
            return `<div class="flow-screen-block flow-screen-block--text flow-screen-block--static">${UI.renderFormattedText(b.text)}</div>`;
        }
        if (b.kind === "skip") {
            return `<div class="flow-screen-block flow-screen-block--skip flow-screen-block--static"></div>`;
        }
        if (b.kind === "line") {
            return `<div class="flow-screen-block flow-screen-block--line flow-screen-block--static"><hr></div>`;
        }
        if (b.kind === "variable") {
            return `<div class="flow-screen-block flow-screen-block--var flow-screen-block--static">${this.runInputHtml(b.variable)}</div>`;
        }
        // frame
        return `
            <div class="flow-screen-block flow-screen-block--frame flow-screen-block--static">
                <div class="flow-frame-header"><strong>${UI.escapeHtml(b.title || "Frame")}</strong></div>
                <div class="flow-frame-vars">
                    ${(b.variables || []).map(v => `<div class="flow-frame-var-row flow-frame-var-row--static">${this.runInputHtml(v)}</div>`).join("")}
                </div>
            </div>`;
    },

    /** Input real para una variable de pantalla en ejecución, según su tipo. */
    runInputHtml(v) {
        if (v.selectMode && v.selectMode !== "unico") return this.selOptHtml(v);

        const id = `actupdrunvar_${v.id}`;
        const label = `<label for="${id}">${UI.escapeHtml(v.label || v.name)}</label>`;
        if (v.type === "FILE") {
            // Sin orquestador de storage en Actualización de tablas: se recoge
            // el nombre del fichero como valor de texto, sin subirlo.
            return `
                <div class="flow-field-preview flow-field-preview--file">
                    ${label}
                    <label class="file-input-btn" for="${id}">
                        <span class="file-input-btn-icon">📎</span>
                        <span class="file-input-btn-text" data-file-text="${id}">Elegir archivo…</span>
                    </label>
                    <input type="file" class="file-input-native" id="${id}" data-var-name="${UI.escapeHtml(v.name)}" data-var-type="FILE_NAME_ONLY">
                </div>`;
        }
        if (v.type === "BOOLEAN") {
            return `<div class="flow-field-preview flow-field-preview--checkbox"><input type="checkbox" id="${id}" data-var-name="${UI.escapeHtml(v.name)}" data-var-type="BOOLEAN">${label}</div>`;
        }
        const htmlType = { INTEGER: "number", FLOAT: "number", NUMERIC: "number", DATE: "date", DATETIME: "datetime-local", TIMESTAMP: "datetime-local" }[v.type] || "text";
        return `<div class="flow-field-preview">${label}<input type="${htmlType}" id="${id}" data-var-name="${UI.escapeHtml(v.name)}" data-var-type="${UI.escapeHtml(v.type || "STRING")}"></div>`;
    },

    // ---- editor de select-options (rango / varios valores / cualquiera), igual que en flow_run.js ----
    selOptDefaultRow(mode) {
        return { sign: "I", option: mode === "rango" ? "BT" : "EQ", low: "", high: "" };
    },

    selOptNeedsHigh(option) {
        return option === "BT" || option === "NB";
    },

    selOptHtml(v) {
        if (!this.selOptState[v.id]) this.selOptState[v.id] = [this.selOptDefaultRow(v.selectMode)];
        return `<div class="flow-field-preview flow-field-preview--selopt" id="actupd_selopt_wrap_${v.id}">${this.selOptInnerHtml(v)}</div>`;
    },

    selOptInnerHtml(v) {
        const rows = this.selOptState[v.id] || [this.selOptDefaultRow(v.selectMode)];
        return `
            <label>${UI.escapeHtml(v.label || v.name)}</label>
            <div class="selopt-table">
                ${rows.map((r, idx) => this.selOptRowHtml(v, idx, r)).join("")}
            </div>
            <button type="button" class="flow-mini-btn selopt-add-btn" data-selopt-add="${v.id}">+ Valor</button>`;
    },

    SELOPT_MODE_OPTIONS: {
        rango: [
            { value: "BT", label: "Entre" }, { value: "GE", label: "Mayor o igual que" },
            { value: "LE", label: "Menor o igual que" }, { value: "GT", label: "Mayor que" },
            { value: "LT", label: "Menor que" }, { value: "EQ", label: "Igual a" }
        ],
        multiple: [{ value: "EQ", label: "Igual a" }, { value: "NE", label: "Distinto de" }],
        cualquiera: [
            { value: "EQ", label: "Igual a (EQ)" }, { value: "NE", label: "Distinto de (NE)" },
            { value: "GT", label: "Mayor que (GT)" }, { value: "GE", label: "Mayor o igual (GE)" },
            { value: "LT", label: "Menor que (LT)" }, { value: "LE", label: "Menor o igual (LE)" },
            { value: "BT", label: "Entre (BT)" }, { value: "NB", label: "No entre (NB)" },
            { value: "CP", label: "Contiene patrón, admite * (CP)" }, { value: "NP", label: "No contiene patrón (NP)" }
        ]
    },

    selOptRowHtml(v, idx, row) {
        const options = this.SELOPT_MODE_OPTIONS[v.selectMode] || this.SELOPT_MODE_OPTIONS.cualquiera;
        const needsHigh = this.selOptNeedsHigh(row.option);
        return `
            <div class="selopt-row">
                <select class="selopt-sign" data-selopt-field="sign" data-selopt-var="${v.id}" data-selopt-idx="${idx}" title="Incluir / excluir">
                    <option value="I" ${row.sign !== "E" ? "selected" : ""}>Incl.</option>
                    <option value="E" ${row.sign === "E" ? "selected" : ""}>Excl.</option>
                </select>
                <select class="selopt-option" data-selopt-field="option" data-selopt-var="${v.id}" data-selopt-idx="${idx}">
                    ${options.map(o => `<option value="${o.value}" ${row.option === o.value ? "selected" : ""}>${o.label}</option>`).join("")}
                </select>
                <input type="text" class="selopt-low" data-selopt-field="low" data-selopt-var="${v.id}" data-selopt-idx="${idx}"
                       placeholder="Valor" value="${UI.escapeHtml(row.low || "")}">
                <input type="text" class="selopt-high" data-selopt-field="high" data-selopt-var="${v.id}" data-selopt-idx="${idx}"
                       placeholder="y" value="${UI.escapeHtml(row.high || "")}" ${needsHigh ? "" : 'style="display:none"'}>
                <button type="button" class="selopt-remove-btn" data-selopt-remove="${v.id}:${idx}" title="Eliminar valor">✕</button>
            </div>`;
    },

    bindSelOptDelegation(body) {
        if (this._selOptDelegationBound) return;
        this._selOptDelegationBound = true;

        body.addEventListener("click", (e) => {
            const addBtn = e.target.closest("[data-selopt-add]");
            if (addBtn) {
                const varId = addBtn.dataset.seloptAdd;
                const v = this.findRunVariableById(varId);
                if (!v) return;
                this.selOptState[varId] = this.selOptState[varId] || [];
                this.selOptState[varId].push(this.selOptDefaultRow(v.selectMode));
                this.refreshSelOptWrap(v);
                return;
            }
            const removeBtn = e.target.closest("[data-selopt-remove]");
            if (removeBtn) {
                const [varId, idxStr] = removeBtn.dataset.seloptRemove.split(":");
                const idx = parseInt(idxStr, 10);
                const v = this.findRunVariableById(varId);
                if (!v || !this.selOptState[varId]) return;
                this.selOptState[varId].splice(idx, 1);
                if (!this.selOptState[varId].length) this.selOptState[varId].push(this.selOptDefaultRow(v.selectMode));
                this.refreshSelOptWrap(v);
            }
        });

        body.addEventListener("change", (e) => {
            const el = e.target.closest("[data-selopt-field]");
            if (!el) return;
            const varId = el.dataset.seloptVar;
            const idx = parseInt(el.dataset.seloptIdx, 10);
            const field = el.dataset.seloptField;
            const rows = this.selOptState[varId];
            if (!rows || !rows[idx]) return;
            rows[idx][field] = el.value;
            if (field === "option") {
                const v = this.findRunVariableById(varId);
                if (v) this.refreshSelOptWrap(v);
            }
        });
    },

    refreshSelOptWrap(v) {
        const wrap = document.getElementById(`actupd_selopt_wrap_${v.id}`);
        if (!wrap) return;
        wrap.innerHTML = this.selOptInnerHtml(v);
    },

    findRunVariableById(varId) {
        return this.flatVars(this.runScreen).find(v => v.id === varId) || null;
    },

    /** Recoge del DOM {nombre: valor} para las variables de la pantalla de ejecución.
     *  Las variables en modo rango/varios/cualquiera devuelven la tabla de
     *  select-options: [{sign, option, low, high}, ...]. */
    collectRunScreenValues() {
        const values = {};
        document.querySelectorAll("#actUpdRunBody [data-var-name]").forEach(el => {
            const name = el.dataset.varName;
            const type = el.dataset.varType;
            if (type === "FILE_NAME_ONLY") {
                values[name] = (el.files && el.files[0]) ? el.files[0].name : "";
            } else if (type === "BOOLEAN") {
                values[name] = el.checked;
            } else {
                values[name] = el.value;
            }
        });

        this.flatVars(this.runScreen).forEach(v => {
            if (v.selectMode && v.selectMode !== "unico") {
                const rows = (this.selOptState[v.id] || []).filter(r => (r.low || "").toString().trim() !== "");
                values[v.name] = rows.map(r => ({
                    sign: r.sign === "E" ? "E" : "I",
                    option: r.option || "EQ",
                    low: r.low || "",
                    high: this.selOptNeedsHigh(r.option) ? (r.high || "") : ""
                }));
            }
        });

        return values;
    },

    /** Traduce una fila de select-options (sign/option/low/high) a una condición SQL. */
    selOptRowSql(col, row) {
        const low = Provider.esc(row.low), high = Provider.esc(row.high || "");
        switch (row.option) {
            case "EQ": return `${col} = '${low}'`;
            case "NE": return `${col} <> '${low}'`;
            case "GT": return `${col} > '${low}'`;
            case "GE": return `${col} >= '${low}'`;
            case "LT": return `${col} < '${low}'`;
            case "LE": return `${col} <= '${low}'`;
            case "BT": return `${col} BETWEEN '${low}' AND '${high}'`;
            case "NB": return `NOT (${col} BETWEEN '${low}' AND '${high}')`;
            case "CP": return `${col} LIKE '${low.replace(/\*/g, "%")}'`;
            case "NP": return `${col} NOT LIKE '${low.replace(/\*/g, "%")}'`;
            default: return null;
        }
    },

    /** Tabla select-options -> condición SQL: incluidos en OR, excluidos como AND NOT, al estilo select-options de SAP. */
    selOptTableSql(col, rows) {
        const inc = (rows || []).filter(r => r.sign !== "E" && (r.low || "").toString().trim() !== "");
        const exc = (rows || []).filter(r => r.sign === "E" && (r.low || "").toString().trim() !== "");
        const incSql = inc.map(r => this.selOptRowSql(col, r)).filter(Boolean);
        const excSql = exc.map(r => this.selOptRowSql(col, r)).filter(Boolean);
        const parts = [];
        if (incSql.length) parts.push(`(${incSql.join(" OR ")})`);
        excSql.forEach(c => parts.push(`NOT (${c})`));
        return parts.length ? parts.join(" AND ") : null;
    },

    buildWhere(fields, variableValues) {
        const clauses = [];
        fields.forEach(f => {
            if (!f.filter || f.filter.type === "NONE") return;
            if (f.filter.type === "CONST") {
                if (f.filter.value === null || f.filter.value === undefined || f.filter.value === "") return;
                clauses.push(`${f.name} = '${Provider.esc(f.filter.value)}'`);
                return;
            }
            if (f.filter.type === "VAR") {
                const varName = this.varNameById(f.filter.value);
                const val = variableValues[varName];
                if (Array.isArray(val)) {
                    const sql = this.selOptTableSql(f.name, val);
                    if (sql) clauses.push(sql);
                } else if (val !== null && val !== undefined && val !== "") {
                    clauses.push(`${f.name} = '${Provider.esc(val)}'`);
                }
            }
        });
        return clauses.length ? `WHERE ${clauses.join(" AND ")}` : "";
    },

    varNameById(id) {
        const v = this.flatVars(this.runScreen).find(x => x.id === id);
        return v ? v.name : id;
    },

    async executeRun() {
        const btn = document.getElementById("actUpdRunExecute");
        const hint = document.getElementById("actUpdRunHint");
        if (btn) btn.disabled = true;
        if (hint) hint.textContent = "Cargando datos...";

        try {
            const values = this.collectRunScreenValues();
            const where = this.buildWhere(this.runFields, values);
            const table = Provider.qualify(this.project.DATASET, this.runRecord.TABLA);
            const rows = await Provider.runQuery(`SELECT * FROM ${table} ${where}`);
            await this.buildGridState(this.runRecord, this.runFields, rows, where);
            this.switchRunTab("table");
        } catch (err) {
            UI.toast("Error al cargar la tabla: " + err.message, "error");
        } finally {
            if (btn) btn.disabled = false;
            if (hint) hint.textContent = "";
        }
    },

    // ---------------- Pestaña 2: Tabla (antes "Monitor") ----------------
    // Mismo grid tipo "Excel" de siempre (pegar bloques, exportar/importar,
    // filtro por columna, diff altas/mods/bajas al grabar), ahora incrustado
    // en la pestaña "Tabla" en lugar de un monitor de progreso.
    async buildGridState(record, fields, rows, where) {
        const state = {
            record, fields, where,
            columns: fields.map(f => f.name),
            originalRows: rows.map(r => ({ ...r })),
            currentRows: rows.map(r => ({ ...r, __rowId: Provider.newId(), __isNew: false })),
            deletedRowIds: [],
            colFilters: {},
            sortCol: null, sortDir: 1,
            visibleCount: this.RENDER_CHUNK,
            optionsByField: {},
            selectedRowIds: new Set(),
            lastClickedRowId: null
        };
        this.gridState = state;

        // Precarga las opciones de validación (constante/dimensión/jerarquía)
        // una sola vez por campo, para no volver a consultar en cada celda.
        for (const f of fields) {
            if (f.validation && f.validation.type !== "NONE") {
                state.optionsByField[f.name] = await this.resolveValidationOptions(f.validation);
            }
        }
    },

    renderRunTableView() {
        const body = document.getElementById("actUpdRunBody");
        if (!this.gridState) {
            body.innerHTML = this.hasScreenVars
                ? `<div class="module-empty">Todavía no se ha cargado la tabla. Pulsa "▶ Cargar tabla" en la pestaña "Pantalla de variables".</div>`
                : `<div class="module-empty">Todavía no se ha cargado la tabla.</div>`;
            return;
        }
        const state = this.gridState;
        body.innerHTML = `
            <div class="values-toolbar">
                <button class="btn btn-secondary btn-sm" id="actUpdGridAddRow">+ Añadir fila</button>
                <button class="btn btn-secondary btn-sm" id="actUpdGridExportCsv">Exportar CSV</button>
                <button class="btn btn-secondary btn-sm" id="actUpdGridExportXlsx">Exportar Excel</button>
                <button class="btn btn-secondary btn-sm" id="actUpdGridImport">Importar archivo</button>
                <input type="file" id="actUpdGridFileInput" accept=".csv,.xlsx,.xls" style="display:none;">
                <button class="btn btn-danger btn-sm" id="actUpdGridDeleteSel" style="display:none;">🗑 Eliminar seleccionadas</button>
                <span class="values-toolbar-spacer"></span>
                <span class="values-row-count" id="actUpdGridCount"></span>
                <button class="btn btn-primary btn-sm" id="actUpdGridSave">Grabar</button>
            </div>
            <p class="form-hint">Pega bloques de celdas directamente desde Excel (Ctrl+V sobre una celda). Filtra escribiendo bajo el nombre de cada columna. Marca la casilla de varias filas (Mayús+clic para seleccionar un rango) y usa "🗑 Eliminar seleccionadas" para borrarlas de golpe. "Grabar" sustituye estas filas en la tabla; la clave (${UI.escapeHtml(state.columns[0])}) debe ser única.</p>
            <div class="values-grid-wrap values-grid-wrap--modal" id="actUpdGridWrap"><span class="spinner"></span></div>`;

        document.getElementById("actUpdGridAddRow").addEventListener("click", () => {
            const blank = {};
            state.columns.forEach(c => { blank[c] = ""; });
            state.currentRows.push({ ...blank, __rowId: Provider.newId(), __isNew: true });
            this.renderGrid();
        });
        document.getElementById("actUpdGridExportCsv").addEventListener("click", () => this.exportGrid("csv"));
        document.getElementById("actUpdGridExportXlsx").addEventListener("click", () => this.exportGrid("xlsx"));
        document.getElementById("actUpdGridImport").addEventListener("click", () => document.getElementById("actUpdGridFileInput").click());
        document.getElementById("actUpdGridFileInput").addEventListener("change", (e) => this.uploadGrid(e));
        document.getElementById("actUpdGridDeleteSel").addEventListener("click", () => this.deleteSelectedRows());
        document.getElementById("actUpdGridSave").addEventListener("click", () => this.commitGrid());

        this.renderGrid();
    },

    async resolveValidationOptions(validation) {
        try {
            if (validation.type === "CONST") {
                return (validation.constants || []).map(c => ({ id: c.id, desc: c.desc }));
            }
            if (validation.type === "DIM") {
                const dim = this.dimensionsCache.find(d => d.DIMENSION_ID === validation.dimensionId);
                if (!dim) return [];
                const dimFields = this.safeParse(dim.CAMPOS_JSON, []);
                const key = dimFields.find(f => f.__isPrimaryName) || dimFields[0];
                const descAttr = dimFields.find(f => f.isDescription);
                const keyCol = Provider.toIdentifier(key ? key.name : dim.DIMENSION);
                const descCol = descAttr ? Provider.toIdentifier(descAttr.name) : null;
                const table = Provider.qualify(this.project.DATASET, dim.TABLA);
                const sql = descCol
                    ? `SELECT DISTINCT ${keyCol} AS ID, ${descCol} AS DESC_ FROM ${table} ORDER BY ${keyCol}`
                    : `SELECT DISTINCT ${keyCol} AS ID FROM ${table} ORDER BY ${keyCol}`;
                const rows = await Provider.runQuery(sql);
                return rows.map(r => ({ id: r.ID, desc: r.DESC_ || "" }));
            }
            if (validation.type === "HIER") {
                const dim = this.dimensionsCache.find(d => d.DIMENSION_ID === validation.dimensionId);
                if (!dim || !validation.hierarchyName) return [];
                const hierRows = await Provider.runQuery(`SELECT NIVELES_JSON FROM ${Provider.qualifyControl("JERARQUIAS")} WHERE DIMENSION_ID='${Provider.esc(validation.dimensionId)}' AND JERARQUIA='${Provider.esc(validation.hierarchyName)}'`);
                if (!hierRows.length) return [];
                const niveles = this.safeParse(hierRows[0].NIVELES_JSON, []);
                const levelCol = niveles[(validation.level || 1) - 1];
                if (!levelCol) return [];

                const dimFields = this.safeParse(dim.CAMPOS_JSON, []);
                const key = dimFields.find(f => f.__isPrimaryName) || dimFields[0];
                const descAttr = dimFields.find(f => f.isDescription);
                const keyCol = Provider.toIdentifier(key ? key.name : dim.DIMENSION);
                const descCol = descAttr ? Provider.toIdentifier(descAttr.name) : null;
                const table = Provider.qualify(this.project.DATASET, dim.TABLA);
                const sql = descCol
                    ? `SELECT DISTINCT ${keyCol} AS ID, ${descCol} AS DESC_ FROM ${table} WHERE ${levelCol} = '${Provider.esc(validation.node)}' ORDER BY ${keyCol}`
                    : `SELECT DISTINCT ${keyCol} AS ID FROM ${table} WHERE ${levelCol} = '${Provider.esc(validation.node)}' ORDER BY ${keyCol}`;
                const rows = await Provider.runQuery(sql);
                return rows.map(r => ({ id: r.ID, desc: r.DESC_ || "" }));
            }
        } catch (err) {
            console.error("Error resolviendo opciones de validación:", err);
        }
        return [];
    },

    getFilteredSortedRows() {
        const state = this.gridState;
        let rows = state.currentRows;

        const activeFilters = Object.entries(state.colFilters).filter(([, v]) => v);
        if (activeFilters.length) {
            rows = rows.filter(r => activeFilters.every(([col, term]) => String(r[col] ?? "").toLowerCase().includes(term)));
        }
        if (state.sortCol) {
            rows = rows.slice().sort((a, b) => {
                const va = a[state.sortCol] ?? "", vb = b[state.sortCol] ?? "";
                if (va < vb) return -1 * state.sortDir;
                if (va > vb) return 1 * state.sortDir;
                return 0;
            });
        }
        return rows;
    },

    // Borra de golpe todas las filas marcadas con la casilla de selección
    // (equivalente al botón "Eliminar" de un ALV tras seleccionar con
    // casillas / Mayús+clic).
    deleteSelectedRows() {
        const state = this.gridState;
        const n = state.selectedRowIds.size;
        if (!n) return;
        this.syncGridFromDom();
        state.currentRows.forEach(r => {
            if (state.selectedRowIds.has(r.__rowId) && !r.__isNew) state.deletedRowIds.push(r.__rowId);
        });
        state.currentRows = state.currentRows.filter(r => !state.selectedRowIds.has(r.__rowId));
        state.selectedRowIds.clear();
        state.lastClickedRowId = null;
        this.renderGrid();
    },

    renderGrid() {
        const state = this.gridState;
        const wrap = document.getElementById("actUpdGridWrap");

        // Sin filas no tiene sentido pintar las columnas: se muestra un aviso
        // y el usuario decide si añade una fila manualmente (lo que sí pinta
        // ya las columnas) en vez de forzar siempre una fila en blanco.
        if (!state.currentRows.length) {
            wrap.innerHTML = `<div class="module-empty module-empty--inline">No hay datos en esta tabla. Usa "+ Añadir fila" para empezar a cargar registros.</div>`;
            const countEl = document.getElementById("actUpdGridCount");
            if (countEl) countEl.textContent = "0 fila(s)";
            return;
        }

        const filtered = this.getFilteredSortedRows();
        const visible = filtered.slice(0, state.visibleCount);

        // La selección solo debe conservar filas que sigan existiendo (por si
        // se borraron/filtraron entre renders).
        const currentIds = new Set(state.currentRows.map(r => r.__rowId));
        state.selectedRowIds.forEach(id => { if (!currentIds.has(id)) state.selectedRowIds.delete(id); });

        const countEl = document.getElementById("actUpdGridCount");
        if (countEl) countEl.textContent = `${filtered.length} de ${state.currentRows.length} fila(s)`;

        const delBtn = document.getElementById("actUpdGridDeleteSel");
        if (delBtn) {
            if (state.selectedRowIds.size > 0) {
                delBtn.style.display = "";
                delBtn.textContent = `🗑 Eliminar seleccionadas (${state.selectedRowIds.size})`;
            } else {
                delBtn.style.display = "none";
            }
        }

        const allFilteredSelected = filtered.length > 0 && filtered.every(r => state.selectedRowIds.has(r.__rowId));

        // Anchos de columna: fijos para casilla/✕ (no se pueden arrastrar), y
        // ajustables (con un tirador en el borde derecho de la cabecera) para
        // el resto. Se recuerdan en el propio gridState entre renders.
        if (!state.colWidths) state.colWidths = {};
        const DEFAULT_COL_W = 170;
        const DEFAULT_DESC_W = 140;
        const colsHtml = `<col class="actupd-col-fixed">` +
            state.fields.map(f => {
                const w = state.colWidths[f.name] || DEFAULT_COL_W;
                const descCol = f.validation && f.validation.showText
                    ? `<col style="width:${state.colWidths[f.name + "__desc"] || DEFAULT_DESC_W}px">` : "";
                return `<col data-col="${UI.escapeHtml(f.name)}" style="width:${w}px">${descCol}`;
            }).join("") +
            `<col class="actupd-col-fixed">`;

        const headerCells = state.fields.map(f => {
            const arrow = state.sortCol === f.name ? (state.sortDir === 1 ? " ▲" : " ▼") : "";
            return `<th data-sort="${UI.escapeHtml(f.name)}" title="${UI.escapeHtml(f.description || "")}">${UI.escapeHtml(f.description || f.name)}${arrow}<br><span class="col-type">${UI.escapeHtml((f.validation && f.validation.type !== "NONE") ? "validado" : "texto")}</span><span class="actupd-col-resize" data-col="${UI.escapeHtml(f.name)}" title="Arrastra para ajustar el ancho"></span></th>${f.validation && f.validation.showText ? `<th class="actupd-desc-col">Texto<span class="actupd-col-resize" data-col="${UI.escapeHtml(f.name)}__desc" title="Arrastra para ajustar el ancho"></span></th>` : ""}`;
        }).join("");

        const filterCells = state.fields.map(f => `
            <th class="actupd-filter-th"><input type="text" class="actupd-col-filter" data-col="${UI.escapeHtml(f.name)}" placeholder="Filtrar..." value="${UI.escapeHtml(state.colFilters[f.name] || "")}"></th>
            ${f.validation && f.validation.showText ? `<th class="actupd-filter-th actupd-desc-col"></th>` : ""}
        `).join("");

        const rowsHtml = visible.map(row => {
            const cells = state.fields.map(f => this.renderCell(row, f)).join("");
            const checked = state.selectedRowIds.has(row.__rowId);
            return `<tr data-row="${row.__rowId}" class="${checked ? "actupd-row-selected" : ""}"><td class="values-row-select"><input type="checkbox" class="actupd-row-select" data-row="${row.__rowId}" ${checked ? "checked" : ""}></td>${cells}<td class="values-row-remove"><button type="button" data-remove-row="${row.__rowId}" title="Eliminar fila">✕</button></td></tr>`;
        }).join("");

        wrap.innerHTML = `
            <table class="values-grid actupd-values-grid">
                <colgroup>${colsHtml}</colgroup>
                <thead>
                    <tr><th class="values-row-select"><input type="checkbox" id="actUpdSelectAll" ${allFilteredSelected ? "checked" : ""}></th>${headerCells}<th></th></tr>
                    <tr><th class="values-row-select"></th>${filterCells}<th></th></tr>
                </thead>
                <tbody id="actUpdGridBody">${rowsHtml}</tbody>
            </table>
            ${filtered.length > visible.length ? `<button class="btn btn-secondary btn-sm" id="actUpdGridMore">Mostrar más (${filtered.length - visible.length} restantes)</button>` : ""}
        `;

        wrap.querySelectorAll(".actupd-col-resize").forEach(handle => {
            handle.addEventListener("click", (e) => e.stopPropagation());
            handle.addEventListener("mousedown", (e) => {
                e.stopPropagation();
                e.preventDefault();
                const key = handle.dataset.col;
                const isDesc = key.endsWith("__desc");
                const colEl = isDesc
                    // columna "__desc": es el <col> inmediatamente siguiente al
                    // de su campo base (no lleva su propio data-col).
                    ? (() => {
                        const base = wrap.querySelector(`col[data-col="${CSS.escape(key.replace("__desc", ""))}"]`);
                        return base ? base.nextElementSibling : null;
                    })()
                    : wrap.querySelector(`col[data-col="${CSS.escape(key)}"]`);
                if (!colEl) return;
                const startX = e.clientX;
                // OJO: getBoundingClientRect() sobre un <col> no es fiable (no
                // es una caja de renderizado normal); se parte del ancho que
                // nosotros mismos le asignamos, no del que reporte el DOM.
                const startWidth = state.colWidths[key] || (isDesc ? DEFAULT_DESC_W : DEFAULT_COL_W);
                const onMove = (ev) => {
                    colEl.style.width = `${Math.max(60, startWidth + (ev.clientX - startX))}px`;
                };
                const onUp = (ev) => {
                    state.colWidths[key] = Math.max(60, startWidth + (ev.clientX - startX));
                    document.removeEventListener("mousemove", onMove);
                    document.removeEventListener("mouseup", onUp);
                };
                document.addEventListener("mousemove", onMove);
                document.addEventListener("mouseup", onUp);
            });
        });

        const selectAllEl = document.getElementById("actUpdSelectAll");
        if (selectAllEl) {
            selectAllEl.addEventListener("change", () => {
                if (selectAllEl.checked) filtered.forEach(r => state.selectedRowIds.add(r.__rowId));
                else filtered.forEach(r => state.selectedRowIds.delete(r.__rowId));
                this.renderGrid();
            });
        }
        // Casilla por fila: clic normal marca/desmarca solo esa fila; con Mayús
        // pulsado, marca (o desmarca, según el estado de la fila de partida) el
        // rango completo entre la última fila tocada y esta, igual que en un
        // ALV o en Excel.
        wrap.querySelectorAll(".actupd-row-select").forEach(cb => {
            cb.addEventListener("click", (e) => {
                const rowId = cb.dataset.row;
                if (e.shiftKey && state.lastClickedRowId) {
                    const ids = visible.map(r => r.__rowId);
                    const from = ids.indexOf(state.lastClickedRowId);
                    const to = ids.indexOf(rowId);
                    if (from !== -1 && to !== -1) {
                        const [start, end] = from < to ? [from, to] : [to, from];
                        const makeChecked = cb.checked;
                        for (let i = start; i <= end; i++) {
                            if (makeChecked) state.selectedRowIds.add(ids[i]);
                            else state.selectedRowIds.delete(ids[i]);
                        }
                    }
                } else {
                    if (cb.checked) state.selectedRowIds.add(rowId);
                    else state.selectedRowIds.delete(rowId);
                }
                state.lastClickedRowId = rowId;
                this.renderGrid();
            });
        });

        wrap.querySelectorAll("th[data-sort]").forEach(th => {
            th.addEventListener("click", () => {
                const col = th.dataset.sort;
                state.sortDir = state.sortCol === col ? -state.sortDir : 1;
                state.sortCol = col;
                this.renderGrid();
            });
        });
        wrap.querySelectorAll(".actupd-col-filter").forEach(inp => {
            inp.addEventListener("click", (e) => e.stopPropagation());
            inp.addEventListener("input", (e) => {
                state.colFilters[e.target.dataset.col] = e.target.value.toLowerCase();
                state.visibleCount = this.RENDER_CHUNK;
                this.renderGrid();
                const again = wrap.querySelector(`.actupd-col-filter[data-col="${CSS.escape(e.target.dataset.col)}"]`);
                if (again) { again.focus(); again.setSelectionRange(again.value.length, again.value.length); }
            });
        });
        const moreBtn = document.getElementById("actUpdGridMore");
        if (moreBtn) moreBtn.addEventListener("click", () => { state.visibleCount += this.RENDER_CHUNK; this.renderGrid(); });

        wrap.querySelectorAll("[data-remove-row]").forEach(btn => {
            btn.addEventListener("click", () => {
                this.syncGridFromDom();
                const rowId = btn.dataset.removeRow;
                const row = state.currentRows.find(r => r.__rowId === rowId);
                if (row && !row.__isNew) state.deletedRowIds.push(rowId);
                state.currentRows = state.currentRows.filter(r => r.__rowId !== rowId);
                state.selectedRowIds.delete(rowId);
                this.renderGrid();
            });
        });

        wrap.querySelectorAll(".actupd-cell-input").forEach(el => {
            const evt = el.tagName === "SELECT" || el.type === "checkbox" ? "change" : "input";
            el.addEventListener(evt, (e) => this.onCellChange(e, el));
        });
        wrap.querySelectorAll(".actupd-search-trigger").forEach(btn => {
            btn.addEventListener("click", () => this.openSearchHelp(btn.dataset.row, btn.dataset.field));
        });

        const tbody = document.getElementById("actUpdGridBody");
        tbody.addEventListener("paste", (e) => this.handlePaste(e));
    },

    renderCell(row, f) {
        const val = row[f.name] ?? "";
        const v = f.validation;
        let inputHtml;

        if (!v || v.type === "NONE") {
            inputHtml = `<input type="text" class="actupd-cell-input" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}" data-col-idx="${this.gridState.columns.indexOf(f.name)}" value="${UI.escapeHtml(val)}">`;
        } else {
            const options = this.gridState.optionsByField[f.name] || [];
            if (v.searchHelp === "CHECKBOX" && options.length === 2) {
                const checked = String(val) === String(options[1].id);
                inputHtml = `<input type="checkbox" class="actupd-cell-input" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}" data-off="${UI.escapeHtml(options[0].id)}" data-on="${UI.escapeHtml(options[1].id)}" ${checked ? "checked" : ""}>`;
            } else if (v.searchHelp === "SEARCH") {
                inputHtml = `<span class="actupd-search-cell">
                    <input type="text" class="actupd-cell-input" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}" data-col-idx="${this.gridState.columns.indexOf(f.name)}" value="${UI.escapeHtml(val)}">
                    <button type="button" class="actupd-search-trigger" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}" title="Buscar">🔍</button>
                </span>`;
            } else {
                // "Listbox (desplegable)": un <select> nativo real (estándar,
                // accesible, navegable con teclado). No es posible mostrar
                // "clave — texto" en la lista y solo la clave ya cerrado, así
                // que se opta por mostrar solo la clave en ambos casos; la
                // descripción queda disponible como tooltip al pasar el ratón.
                inputHtml = `<select class="actupd-cell-input" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}">
                    <option value="">${v.allowEmpty ? "(vacío)" : "— selecciona —"}</option>
                    ${options.map(o => `<option value="${UI.escapeHtml(o.id)}" ${o.desc ? `title="${UI.escapeHtml(o.desc)}"` : ""} ${String(val) === String(o.id) ? "selected" : ""}>${UI.escapeHtml(o.id)}</option>`).join("")}
                </select>`;
            }
        }

        const cellTd = `<td>${inputHtml}</td>`;
        if (v && v.showText) {
            const options = this.gridState.optionsByField[f.name] || [];
            const opt = options.find(o => String(o.id) === String(val));
            return cellTd + `<td class="actupd-desc-col">${UI.escapeHtml(opt ? opt.desc : "")}</td>`;
        }
        return cellTd;
    },

    /** Vuelca lo que hay en los <input>/<select> del DOM a state.currentRows (por si hay cambios sin evento "change" disparado) */
    syncGridFromDom() {
        const state = this.gridState;
        const tbody = document.getElementById("actUpdGridBody");
        if (!tbody) return;
        tbody.querySelectorAll("tr").forEach(tr => {
            const row = state.currentRows.find(r => r.__rowId === tr.dataset.row);
            if (!row) return;
            tr.querySelectorAll(".actupd-cell-input").forEach(el => {
                if (el.type === "checkbox") row[el.dataset.field] = el.checked ? el.dataset.on : el.dataset.off;
                else row[el.dataset.field] = el.value;
            });
        });
    },

    onCellChange(e, el) {
        const state = this.gridState;
        const rowId = el.dataset.row;
        const field = el.dataset.field;
        const row = state.currentRows.find(r => r.__rowId === rowId);
        if (!row) return;
        if (el.type === "checkbox") {
            row[field] = el.checked ? el.dataset.on : el.dataset.off;
        } else {
            row[field] = el.value;
        }
        // Repinta solo la columna de texto asociada, si aplica, sin re-renderizar todo el grid.
        const f = state.fields.find(x => x.name === field);
        if (f && f.validation && f.validation.showText) {
            const options = state.optionsByField[field] || [];
            const opt = options.find(o => String(o.id) === String(row[field]));
            const td = el.closest("tr").querySelector(".actupd-desc-col");
            if (td) td.textContent = opt ? opt.desc : "";
        }
    },

    /**
     * Pegado de bloques de Excel (mismo comportamiento que Mantenimiento de
     * dimensiones): pegar un bloque con tabulaciones/saltos de línea sobre
     * una celda rellena esa celda y las siguientes en fila/columna,
     * añadiendo filas nuevas si hace falta.
     */
    handlePaste(e) {
        const target = e.target;
        if (!target.matches(".actupd-cell-input") || target.tagName !== "INPUT") return;
        const text = (e.clipboardData || window.clipboardData).getData("text");
        if (!text || (!text.includes("\t") && !text.includes("\n"))) return; // pegado simple de una celda: comportamiento nativo

        e.preventDefault();
        this.syncGridFromDom();

        const state = this.gridState;
        const pastedRows = text.replace(/\r/g, "").split("\n").filter((r, i, arr) => !(i === arr.length - 1 && r === ""));
        const startRowId = target.closest("tr").dataset.row;
        const startRowIdx = state.currentRows.findIndex(r => r.__rowId === startRowId);
        const startColIdx = parseInt(target.dataset.colIdx, 10);

        pastedRows.forEach((rowText, rOffset) => {
            const cells = rowText.split("\t");
            let rowIdx = startRowIdx + rOffset;
            while (rowIdx >= state.currentRows.length) {
                const blank = {};
                state.columns.forEach(c => { blank[c] = ""; });
                state.currentRows.push({ ...blank, __rowId: Provider.newId(), __isNew: true });
            }
            cells.forEach((cellText, cOffset) => {
                const colIdx = startColIdx + cOffset;
                if (colIdx >= state.columns.length) return;
                state.currentRows[rowIdx][state.columns[colIdx]] = cellText;
            });
        });

        this.renderGrid();
    },

    openSearchHelp(rowId, field) {
        const state = this.gridState;
        const options = state.optionsByField[field] || [];
        let overlay = document.getElementById("actUpdSearchHelpModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "actUpdSearchHelpModal";
            document.body.appendChild(overlay);
        }
        overlay.innerHTML = `
            <div class="modal-box">
                <div class="modal-header"><h3>Buscar valor — ${UI.escapeHtml(field)}</h3><button class="modal-close" id="actUpdSHClose">&times;</button></div>
                <div class="modal-body">
                    <input type="text" id="actUpdSHSearch" placeholder="Buscar por ID o descripción...">
                    <div id="actUpdSHResults" class="actupd-sh-results"></div>
                </div>
            </div>`;
        overlay.classList.add("visible");
        const close = () => overlay.remove();
        document.getElementById("actUpdSHClose").addEventListener("click", close);

        const renderResults = (term) => {
            const t = (term || "").toLowerCase();
            const filtered = options.filter(o => !t || String(o.id).toLowerCase().includes(t) || (o.desc || "").toLowerCase().includes(t));
            document.getElementById("actUpdSHResults").innerHTML = filtered.slice(0, 200).map(o =>
                `<div class="actupd-sh-item" data-id="${UI.escapeHtml(o.id)}">${UI.escapeHtml(o.id)}${o.desc ? " — " + UI.escapeHtml(o.desc) : ""}</div>`
            ).join("") || `<div class="hierarchy-levels-empty">Sin resultados.</div>`;
            document.querySelectorAll("#actUpdSHResults .actupd-sh-item").forEach(item => {
                item.addEventListener("click", () => {
                    this.syncGridFromDom();
                    const row = state.currentRows.find(r => r.__rowId === rowId);
                    row[field] = item.dataset.id;
                    close();
                    this.renderGrid();
                });
            });
        };
        renderResults("");
        document.getElementById("actUpdSHSearch").addEventListener("input", (e) => renderResults(e.target.value));
    },

    // ------------------------------------------------------------
    // Exportar / Importar (igual que Mantenimiento de dimensiones)
    // ------------------------------------------------------------
    toAoa() {
        this.syncGridFromDom();
        const state = this.gridState;
        const header = state.columns;
        const body = state.currentRows.map(r => state.columns.map(c => r[c] ?? ""));
        return [header, ...body];
    },

    toCsv(aoa) {
        return aoa.map(row => row.map(cell => {
            const s = String(cell ?? "");
            return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        }).join(",")).join("\n");
    },

    exportGrid(kind) {
        const state = this.gridState;
        if (kind === "csv") {
            const csv = this.toCsv(this.toAoa());
            UI.downloadBlob(`${state.record.TABLA}.csv`, "\uFEFF" + csv, "text/csv;charset=utf-8");
        } else {
            const ws = XLSX.utils.aoa_to_sheet(this.toAoa());
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Datos");
            XLSX.writeFile(wb, `${state.record.TABLA}.xlsx`);
        }
    },

    parseFileToAoa(file) {
        return new Promise((resolve, reject) => {
            const isCsv = /\.csv$/i.test(file.name);
            const reader = new FileReader();
            reader.onerror = () => reject(new Error("No se pudo leer el archivo."));
            reader.onload = () => {
                try {
                    const wb = isCsv
                        ? XLSX.read(reader.result, { type: "string" })
                        : XLSX.read(new Uint8Array(reader.result), { type: "array" });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
                    resolve(aoa);
                } catch (err) {
                    reject(err);
                }
            };
            if (isCsv) reader.readAsText(file, "utf-8");
            else reader.readAsArrayBuffer(file);
        });
    },

    /**
     * Carga desde fichero: SUSTITUYE el contenido del grid, igual que en
     * Mantenimiento de dimensiones. Para conservar la identidad de las
     * filas ya existentes (y así poder distinguir "modificada" de "borrada
     * + añadida" en el resumen de cambios), se empareja cada fila del
     * fichero con una fila original por el valor de su PRIMERA columna
     * (asumida como identificador natural). Si no hay match, se trata
     * como fila nueva; cualquier fila original sin match se marca borrada.
     */
    async uploadGrid(e) {
        const file = e.target.files[0];
        e.target.value = "";
        if (!file) return;
        const state = this.gridState;
        try {
            const aoa = await this.parseFileToAoa(file);
            if (!aoa.length) { UI.toast("El archivo está vacío.", "error"); return; }

            const [headerRow, ...dataRows] = aoa;
            const headerMap = headerRow.map(h => String(h).trim().toUpperCase());
            const parsed = dataRows
                .filter(r => r.some(c => String(c ?? "").trim() !== ""))
                .map(r => {
                    const obj = {};
                    state.columns.forEach(c => {
                        const idx = headerMap.indexOf(c.toUpperCase());
                        obj[c] = idx !== -1 ? String(r[idx] ?? "") : "";
                    });
                    return obj;
                });

            if (!parsed.length) { UI.toast("No se encontraron filas de datos en el archivo.", "error"); return; }

            this.syncGridFromDom();
            const keyCol = state.columns[0];
            const originalByKey = new Map(state.originalRows.map(r => [String(r[keyCol]), r]));
            const currentByKey = new Map(state.currentRows.map(r => [String(r[keyCol]), r]));
            const stillPresentKeys = new Set();

            const newCurrent = parsed.map(pRow => {
                const keyVal = String(pRow[keyCol]);
                const original = originalByKey.get(keyVal);
                const existingCurrent = currentByKey.get(keyVal);
                stillPresentKeys.add(keyVal);
                const rowId = existingCurrent ? existingCurrent.__rowId : Provider.newId();
                const row = { __rowId: rowId, __isNew: existingCurrent ? existingCurrent.__isNew : !original };
                state.columns.forEach(c => { row[c] = pRow[c] !== undefined ? pRow[c] : ""; });
                return row;
            });

            state.originalRows.forEach(r => {
                if (!stillPresentKeys.has(String(r[keyCol]))) state.deletedRowIds.push(String(r[keyCol]));
            });

            state.currentRows = newCurrent;
            state.visibleCount = this.RENDER_CHUNK;
            UI.toast(`Fichero cargado: ${parsed.length} fila(s). Pulsa "Grabar" para confirmarlo en la base de datos.`, "success");
            this.renderGrid();
        } catch (err) {
            UI.toast("Error al leer el fichero: " + err.message, "error");
        }
    },

    computeDiff() {
        const state = this.gridState;
        const keyCol = state.columns[0];
        const originalByKey = new Map(state.originalRows.map(r => [String(r[keyCol]), r]));

        const added = [], modified = [], unchanged = [];
        state.currentRows.forEach(row => {
            if (row.__isNew) { added.push(row); return; }
            const orig = originalByKey.get(String(row[keyCol]));
            if (!orig) { added.push(row); return; }
            const changed = state.columns.some(c => String(orig[c] ?? "") !== String(row[c] ?? ""));
            if (changed) modified.push({ before: orig, after: row }); else unchanged.push(row);
        });

        const currentKeys = new Set(state.currentRows.map(r => String(r[keyCol])));
        const deleted = state.originalRows.filter(r => !currentKeys.has(String(r[keyCol])));

        return { added, modified, deleted };
    },

    /** Marca en rojo (clase .duplicate-row, igual que Mantenimiento de dimensiones) las filas con clave repetida */
    highlightDuplicateRows(dupKeys) {
        const tbody = document.getElementById("actUpdGridBody");
        if (!tbody) return;
        const state = this.gridState;
        const keyCol = state.columns[0];
        tbody.querySelectorAll("tr").forEach(tr => tr.classList.remove("duplicate-row"));
        state.currentRows.forEach(r => {
            if (dupKeys.has(String(r[keyCol]).trim().toUpperCase())) {
                const tr = tbody.querySelector(`tr[data-row="${r.__rowId}"]`);
                if (tr) tr.classList.add("duplicate-row");
            }
        });
    },

    async commitGrid() {
        const state = this.gridState;
        this.syncGridFromDom();

        // Clave única (misma validación que Mantenimiento de dimensiones).
        const keyCol = state.columns[0];
        const seen = new Map();
        const dupKeys = new Set();
        state.currentRows.forEach(r => {
            const k = String(r[keyCol] ?? "").trim().toUpperCase();
            if (!k) return;
            if (seen.has(k)) dupKeys.add(k); else seen.set(k, true);
        });
        if (dupKeys.size) {
            this.renderGrid();
            this.highlightDuplicateRows(dupKeys);
            UI.toast(`No se puede grabar: hay ${dupKeys.size} valor(es) de "${keyCol}" repetidos. Corrige o elimina las filas duplicadas (marcadas en rojo).`, "error");
            return;
        }

        // "Permite vacío" por campo validado.
        const violations = [];
        state.fields.forEach(f => {
            if (f.validation && f.validation.type !== "NONE" && f.validation.allowEmpty === false) {
                state.currentRows.forEach(r => {
                    if (r[f.name] === "" || r[f.name] === null || r[f.name] === undefined) violations.push(f.name);
                });
            }
        });
        if (violations.length) {
            UI.toast(`No se puede grabar: hay valor(es) vacío(s) en campos que no lo permiten (${[...new Set(violations)].join(", ")}).`, "error");
            return;
        }

        const diff = this.computeDiff();
        if (!diff.added.length && !diff.modified.length && !diff.deleted.length) {
            UI.toast("No hay cambios que grabar.", "info");
            return;
        }

        const ok = await UI.confirm("Grabar cambios",
            `Se van a grabar ${diff.added.length} alta(s), ${diff.modified.length} modificación(es) y ${diff.deleted.length} baja(s) en ${state.record.TABLA}. ¿Continuar?`);
        if (!ok) return;

        const btn = document.getElementById("actUpdGridSave");
        if (btn) { btn.disabled = true; btn.textContent = "Guardando..."; }
        try {
            const table = Provider.qualify(this.project.DATASET, state.record.TABLA);
            // Estrategia simple y robusta (no requiere clave garantizada a
            // nivel de motor): se borra el subconjunto filtrado y se
            // reinserta completo. Si no hay filtro (sin variables o tabla de
            // partida vacía), state.where es "" y el motor exige un WHERE
            // explícito para el DELETE — se usa WHERE TRUE para borrar todo.
            await Provider.runQuery(`DELETE FROM ${table} ${state.where || "WHERE TRUE"}`);

            const CHUNK = 500;
            for (let i = 0; i < state.currentRows.length; i += CHUNK) {
                const chunk = state.currentRows.slice(i, i + CHUNK);
                const values = chunk.map(r =>
                    `(${state.columns.map(c => r[c] === "" || r[c] === null || r[c] === undefined ? "NULL" : `'${Provider.esc(r[c])}'`).join(", ")})`
                ).join(", ");
                if (chunk.length) {
                    await Provider.runQuery(`INSERT INTO ${table} (${state.columns.join(", ")}) VALUES ${values}`);
                }
            }

            UI.toast("Cambios grabados correctamente.", "success");
            this.showChangeSummary(diff);
            state.originalRows = state.currentRows.map(r => ({ ...r }));
            state.deletedRowIds = [];
            state.currentRows.forEach(r => { r.__isNew = false; });
        } catch (err) {
            UI.toast("Error al grabar: " + err.message, "error");
        } finally {
            if (btn) { btn.disabled = false; btn.textContent = "Grabar"; }
        }
    },

    showChangeSummary(diff) {
        let overlay = document.getElementById("actUpdSummaryModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "actUpdSummaryModal";
            document.body.appendChild(overlay);
        }
        const state = this.gridState;
        overlay.innerHTML = `
            <div class="modal-box">
                <div class="modal-header"><h3>Resumen de cambios</h3><button class="modal-close" id="actUpdSummaryClose">&times;</button></div>
                <div class="modal-body">
                    <p>✔ ${diff.added.length} fila(s) añadida(s)</p>
                    <p>✎ ${diff.modified.length} fila(s) modificada(s)</p>
                    <p>🗑 ${diff.deleted.length} fila(s) eliminada(s)</p>
                    <button class="btn btn-secondary btn-sm" id="actUpdSummaryDetail">Ver detalle</button>
                    <div id="actUpdSummaryDetailBox" class="actupd-summary-detail" style="display:none;"></div>
                </div>
            </div>`;
        overlay.classList.add("visible");
        document.getElementById("actUpdSummaryClose").addEventListener("click", () => overlay.remove());
        document.getElementById("actUpdSummaryDetail").addEventListener("click", (e) => {
            const box = document.getElementById("actUpdSummaryDetailBox");
            if (box.style.display === "none") {
                const keyCol = state.columns[0];
                box.innerHTML = `
                    ${diff.added.length ? `<h4>Añadidas</h4>` + diff.added.map(r => `<div class="actupd-summary-row">+ ${UI.escapeHtml(String(r[keyCol]))}</div>`).join("") : ""}
                    ${diff.modified.length ? `<h4>Modificadas</h4>` + diff.modified.map(m => `<div class="actupd-summary-row">~ ${UI.escapeHtml(String(m.after[keyCol]))}</div>`).join("") : ""}
                    ${diff.deleted.length ? `<h4>Eliminadas</h4>` + diff.deleted.map(r => `<div class="actupd-summary-row">− ${UI.escapeHtml(String(r[keyCol]))}</div>`).join("") : ""}
                `;
                box.style.display = "block";
                e.target.textContent = "Ocultar detalle";
            } else {
                box.style.display = "none";
                e.target.textContent = "Ver detalle";
            }
        });
    }
};
