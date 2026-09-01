/**
 * ============================================================
 * DRACO PLANNING — FUNCIONES
 * ============================================================
 * Operaciones predefinidas sobre los datos de un cubo del proyecto:
 *   - Copy         copiar datos de un conjunto de características a otro
 *   - Delete       borrar datos planificados
 *   - Move         traspasar/reclasificar valores entre miembros
 *   - Revalue      incrementar/reducir valores por porcentaje o factor
 *   - Distribution distribuir un valor entre varios miembros (pendiente)
 *
 * Mismo patrón que ACTUALIZACIÓN DE TABLAS (js/table-updates.js):
 *   1) Diseñador (modal a pantalla completa):
 *        A) Pantalla de variables (idéntica a Flujos de carga / Actualización
 *           de tablas: variables sueltas, frames, textos, espacios, líneas).
 *        B) Filtros: origen (todas las dimensiones en "All" por defecto,
 *           admite valor(es) concretos o una variable de pantalla), destino
 *           (solo Copy/Move: "Same" por defecto, o valor(es)/variable) y
 *           selector de medidas (todas por defecto). Revalue añade además
 *           el incremento (porcentaje/valor, constante o variable).
 *   2) Ejecución: pantalla de variables -> "Ejecutar" genera el SQL
 *      correspondiente y lo lanza en segundo plano; un Monitor (igual que
 *      el de Flujos de carga) sondea DRACO_CONTROL.FUNCIONES_RUNS cada
 *      pocos segundos mientras la ejecución esté en curso.
 *
 * Todo el diseño se guarda como JSON en una única fila de control
 * (DRACO_CONTROL.FUNCIONES), igual que ACTUALIZACIONES.CAMPOS_JSON.
 */
const Functions = {
    TABLE: "FUNCIONES",
    RUN_TABLE: "FUNCIONES_RUNS",
    ID_COL: "FUNCION_ID",
    NAME_COL: "NOMBRE",

    TYPES: {
        COPY: { label: "Copy", icon: "⧉", desc: "Copiar datos de un conjunto de características a otro" },
        DELETE: { label: "Delete", icon: "🗑", desc: "Borrar datos planificados" },
        MOVE: { label: "Move", icon: "⇄", desc: "Traspasar/reclasificar valores entre miembros" },
        REVALUE: { label: "Revalue", icon: "%", desc: "Incrementar/reducir valores por porcentaje o factor" },
        DISTRIBUTION: { label: "Distribution", icon: "⋔", desc: "Distribuir un valor entre varios miembros" }
    },

    typeChoices() {
        return Object.entries(this.TYPES).map(([key, t]) => ({ key, label: t.label, icon: t.icon, desc: t.desc }));
    },

    list: [],
    project: null,
    cubes: [],
    dimensionsCache: [],
    editing: null,
    editingIsNew: true,
    screen: { title: "", blocks: [] },
    config: {},
    currentCube: null,
    screenCollapsed: false,
    dragBlockIdx: null,
    dragFrameVar: null,

    // ---- ejecución ----
    runRecord: null,
    runCube: null,
    runScreen: { title: "", blocks: [] },
    runConfig: {},
    runView: "screen",       // 'screen' | 'monitor'
    pollHandle: null,
    lastRunId: null,
    selOptState: {},
    _selOptDelegationBound: false,

    safeParse(json, fallback) {
        try { return json ? JSON.parse(json) : fallback; } catch (e) { return fallback; }
    },

    /** Mismo formato de pantalla de variables que Flujos de carga / Actualización de tablas. */
    parseScreen(json) {
        const parsed = this.safeParse(json, null);
        if (!parsed) return { title: "", blocks: [] };
        return { title: parsed.title || "", blocks: Array.isArray(parsed.blocks) ? parsed.blocks : [] };
    },

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
                    <h3>Funciones</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET}</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewFuncion">
                    + Nueva función
                </button>
            </div>
            <div id="funcionesListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewFuncion").addEventListener("click", () => this.openCreateModal());
        await this.loadList();
    },

    async loadCubes() {
        const sql = `SELECT CUBO_ID, CUBOS, DESCRIPCION, TABLA, CAMPOS_JSON
                     FROM ${Provider.qualifyControl("CUBOS")}
                     WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                     ORDER BY CUBOS`;
        const rows = await Provider.runQuery(sql);
        this.cubes = rows.map(r => {
            const spec = this.safeParse(r.CAMPOS_JSON, { dimensions: [], measures: [] });
            return {
                id: r.CUBO_ID,
                name: r.CUBOS,
                table: r.TABLA,
                dims: (spec.dimensions || []).map(d => ({ id: d.id, name: d.name, colId: d.colId, type: d.type })),
                measures: (spec.measures || []).map(m => ({ name: m.name, type: m.type }))
            };
        });
    },

    async loadDimensions() {
        const sql = `SELECT DIMENSION_ID, DIMENSION, DESCRIPCION, TABLA, CAMPOS_JSON
                     FROM ${Provider.qualifyControl("DIMENSIONES")}
                     WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                     ORDER BY DIMENSION`;
        this.dimensionsCache = await Provider.runQuery(sql);
    },

    async loadList() {
        const wrap = document.getElementById("funcionesListWrap");
        try {
            try {
                await Provider.runQuery(DracoSchema.ddl(this.TABLE));
                await Provider.runQuery(DracoSchema.ddl(this.RUN_TABLE));
            } catch (ddlErr) {
                console.error("No se pudo verificar/crear las tablas de control de Funciones:", ddlErr);
            }

            const sql = `SELECT ${this.ID_COL}, ${this.NAME_COL}, DESCRIPCION, TIPO, CUBO_ID, CUBO_NOMBRE, CUBO_TABLA, VARIABLES_JSON, CONFIG_JSON
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                         ORDER BY ${this.NAME_COL}`;
            this.list = await Provider.runQuery(sql);

            if (!this.list.length) {
                wrap.innerHTML = `<div class="module-empty">Todavía no hay funciones en este proyecto. Crea la primera con "Nueva función".</div>`;
                return;
            }

            wrap.innerHTML = `
                <div class="data-list">
                    <table>
                        <thead><tr><th>Nombre</th><th>Tipo</th><th>Cubo</th><th></th></tr></thead>
                        <tbody>
                            ${this.list.map(r => {
                                const typeInfo = this.TYPES[r.TIPO] || { label: r.TIPO, icon: "•" };
                                return `
                                <tr>
                                    <td><strong>${UI.escapeHtml(r[this.NAME_COL])}</strong><br><span class="col-type">${UI.escapeHtml(r.DESCRIPCION || "")}</span></td>
                                    <td><span class="table-tag">${typeInfo.icon} ${UI.escapeHtml(typeInfo.label)}</span></td>
                                    <td><span class="table-tag">${UI.escapeHtml(r.CUBO_NOMBRE || r.CUBO_TABLA)}</span></td>
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
            wrap.innerHTML = `<div class="module-empty">Error al cargar funciones: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    async remove(id) {
        const rec = this.list.find(r => r[this.ID_COL] === id);
        const ok = await UI.confirm("Eliminar función", `¿Seguro que quieres eliminar "${rec ? rec[this.NAME_COL] : ""}"?`);
        if (!ok) return;
        try {
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
            UI.toast("Función eliminada.", "success");
            await this.loadList();
        } catch (err) {
            UI.toast("Error al eliminar: " + err.message, "error");
        }
    },

    // ================================================================
    // CREAR (nombre + selector de cubo + selector de tipo con cajitas)
    // ================================================================
    async openCreateModal() {
        await this.loadCubes();
        if (!this.cubes.length) {
            UI.toast("Este proyecto todavía no tiene cubos creados: crea uno primero en Cubos.", "info");
        }

        let overlay = document.getElementById("funcCreateModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "funcCreateModal";
            document.body.appendChild(overlay);
        }
        overlay.classList.add("visible");

        let selectedType = "COPY";

        overlay.innerHTML = `
            <div class="modal-box modal-wide">
                <div class="modal-header">
                    <h3>Nueva función</h3>
                    <button class="modal-close" id="funcCreateClose">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="form-group">
                        <label>Nombre</label>
                        <input type="text" id="funcNewName" placeholder="ej. Copiar Actual a Presupuesto">
                    </div>
                    <div class="form-group">
                        <label>Descripción (opcional)</label>
                        <input type="text" id="funcNewDesc">
                    </div>
                    <div class="form-group">
                        <label>Cubo</label>
                        <select id="funcNewCubo">
                            <option value="">Selecciona un cubo...</option>
                            ${this.cubes.map(c => `<option value="${c.id}">${UI.escapeHtml(c.name)}</option>`).join("")}
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Tipo de función</label>
                        <div class="origin-type-grid" id="funcNewTypeGrid">
                            ${this.typeChoices().map(t => `
                                <button type="button" class="origin-type-card ${t.key === selectedType ? "active" : ""}" data-type="${t.key}" title="${UI.escapeHtml(t.desc)}">
                                    <span class="origin-type-card-icon">${t.icon}</span>
                                    <span class="origin-type-card-label">${UI.escapeHtml(t.label)}</span>
                                </button>`).join("")}
                        </div>
                        <p class="form-hint" id="funcNewTypeDesc">${UI.escapeHtml(this.TYPES[selectedType].desc)}</p>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="funcCreateCancel">Cancelar</button>
                    <button class="btn btn-primary" id="funcCreateContinue">Continuar</button>
                </div>
            </div>`;

        const close = () => overlay.remove();
        document.getElementById("funcCreateClose").addEventListener("click", close);
        document.getElementById("funcCreateCancel").addEventListener("click", close);

        overlay.querySelectorAll("#funcNewTypeGrid [data-type]").forEach(btn => {
            btn.addEventListener("click", () => {
                selectedType = btn.dataset.type;
                overlay.querySelectorAll("#funcNewTypeGrid [data-type]").forEach(b => b.classList.toggle("active", b === btn));
                document.getElementById("funcNewTypeDesc").textContent = this.TYPES[selectedType].desc;
            });
        });

        document.getElementById("funcCreateContinue").addEventListener("click", () => {
            const name = document.getElementById("funcNewName").value.trim();
            const cuboId = document.getElementById("funcNewCubo").value;
            const cube = this.cubes.find(c => c.id === cuboId);
            if (!name) { UI.toast("Indica un nombre.", "error"); return; }
            if (!cube) { UI.toast("Selecciona un cubo.", "error"); return; }

            const draft = {
                [this.ID_COL]: Provider.newId(),
                PROYECTO_ID: this.project.PROYECTO_ID,
                [this.NAME_COL]: name,
                DESCRIPCION: document.getElementById("funcNewDesc").value.trim(),
                TIPO: selectedType,
                CUBO_ID: cube.id,
                CUBO_NOMBRE: cube.name,
                CUBO_TABLA: cube.table,
                VARIABLES_JSON: "[]",
                CONFIG_JSON: JSON.stringify(this.defaultConfig(selectedType, cube))
            };
            close();
            this.editingIsNew = true;
            this.openEditor(draft);
        });
    },

    defaultConfig(tipo, cube) {
        const originDims = {};
        (cube.dims || []).forEach(d => { originDims[d.colId] = { mode: "ALL", values: [], varId: "" }; });
        const config = {
            origin: { dims: originDims },
            measuresMode: "ALL",
            measures: []
        };
        if (tipo === "COPY" || tipo === "MOVE") {
            const destDims = {};
            (cube.dims || []).forEach(d => { destDims[d.colId] = { mode: "SAME", values: [], varId: "" }; });
            config.dest = { dims: destDims };
        }
        if (tipo === "REVALUE") {
            config.revalue = { kind: "PERCENT", amountMode: "CONST", amountConst: "", amountVarId: "" };
        }
        return config;
    },

    // ================================================================
    // DISEÑADOR (modal a pantalla completa, 2 bloques)
    // ================================================================
    async openEditor(record) {
        this.editing = record;
        this.editingIsNew = !this.list.some(r => r[this.ID_COL] === record[this.ID_COL]);
        this.screen = this.parseScreen(record.VARIABLES_JSON);

        await this.loadDimensions();
        if (!this.cubes.length) await this.loadCubes();
        this.currentCube = this.cubes.find(c => c.id === record.CUBO_ID) ||
            { id: record.CUBO_ID, name: record.CUBO_NOMBRE, table: record.CUBO_TABLA, dims: [], measures: [] };

        this.config = this.safeParse(record.CONFIG_JSON, this.defaultConfig(record.TIPO, this.currentCube));
        // Fusiona con las dimensiones/medidas actuales del cubo (por si se
        // han añadido/quitado desde que se creó esta función).
        this.reconcileConfigWithCube();

        const typeInfo = this.TYPES[record.TIPO] || { label: record.TIPO, icon: "•" };

        let overlay = document.getElementById("funcEditorModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "funcEditorModal";
            document.body.appendChild(overlay);
        }
        overlay.classList.add("visible");
        this.overlay = overlay;

        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div class="modal-title-edit-wrap">
                        <h3 id="funcModalTitle" class="modal-title-editable" contenteditable="true" spellcheck="false" title="Clic para renombrar">${UI.escapeHtml(record[this.NAME_COL])}</h3>
                        <span class="modal-subtitle">${typeInfo.icon} ${UI.escapeHtml(typeInfo.label)} · Cubo: ${UI.escapeHtml(this.currentCube.name)}</span>
                    </div>
                    <div class="modal-header-right">
                        <button class="modal-close" id="funcEditorClose">&times;</button>
                    </div>
                </div>
                <div class="modal-body modal-body-flush">
                    <div id="funcScreenPart" class="flow-part flow-part--screen"></div>
                    <div id="funcConfigPart" class="flow-part"></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="funcEditorCancel">Cancelar</button>
                    <button class="btn btn-secondary" id="funcEditorRun" ${this.editingIsNew ? "disabled" : ""}>▶ Ejecutar</button>
                    <button class="btn btn-primary" id="funcEditorSave">Guardar</button>
                </div>
            </div>`;

        document.getElementById("funcEditorClose").addEventListener("click", () => overlay.remove());
        document.getElementById("funcEditorCancel").addEventListener("click", () => overlay.remove());
        document.getElementById("funcEditorSave").addEventListener("click", () => this.save());
        document.getElementById("funcEditorRun").addEventListener("click", () => {
            if (!this.editingIsNew) this.startRun(this.editing);
        });

        this.renderScreenBlock();
        this.renderConfigBlock();
    },

    /** Si el cubo ganó/perdió dimensiones o medidas desde que se guardó esta
     *  función, añade entradas por defecto para las nuevas y descarta las
     *  de columnas que ya no existen. */
    reconcileConfigWithCube() {
        const dimColIds = (this.currentCube.dims || []).map(d => d.colId);
        const measureNames = (this.currentCube.measures || []).map(m => m.name);

        const reconcileDims = (dims, defaultMode) => {
            const out = {};
            dimColIds.forEach(colId => { out[colId] = (dims && dims[colId]) || { mode: defaultMode, values: [], varId: "" }; });
            return out;
        };

        this.config.origin = this.config.origin || { dims: {} };
        this.config.origin.dims = reconcileDims(this.config.origin.dims, "ALL");
        this.config.measuresMode = this.config.measuresMode || "ALL";
        this.config.measures = (this.config.measures || []).filter(m => measureNames.includes(m));

        if (this.editing.TIPO === "COPY" || this.editing.TIPO === "MOVE") {
            this.config.dest = this.config.dest || { dims: {} };
            this.config.dest.dims = reconcileDims(this.config.dest.dims, "SAME");
        }
        if (this.editing.TIPO === "REVALUE") {
            this.config.revalue = this.config.revalue || { kind: "PERCENT", amountMode: "CONST", amountConst: "", amountVarId: "" };
        }
    },

    // ---------------- Bloque A: pantalla de variables ----------------
    // Mismo diseñador que en Flujos de carga / Actualización de tablas:
    // variables sueltas, frames, textos, espacios y líneas.
    renderScreenBlock() {
        const part = document.getElementById("funcScreenPart");
        const screen = this.screen;

        part.innerHTML = `
            <div class="flow-part-header flow-part-header--screen">
                <button type="button" class="flow-part-toggle" id="funcScreenToggle">
                    <span class="flow-group-caret ${this.screenCollapsed ? "is-collapsed" : ""}">▾</span>
                    <span>Pantalla de entrada de variables</span>
                </button>
                <div class="flow-screen-toolbar-mini">
                    <button type="button" class="flow-mini-btn" id="funcAddVar" title="Añadir variable">+ Var</button>
                    <button type="button" class="flow-mini-btn" id="funcAddFrame" title="Añadir frame">+ Frame</button>
                    <button type="button" class="flow-mini-btn" id="funcAddText" title="Añadir texto">+ Texto</button>
                    <button type="button" class="flow-mini-btn" id="funcAddSkip" title="Añadir espacio en blanco">+ Espacio</button>
                    <button type="button" class="flow-mini-btn" id="funcAddLine" title="Añadir línea separadora">+ Línea</button>
                </div>
            </div>
            <div class="flow-screen-box ${this.screenCollapsed ? "is-collapsed" : ""}" id="funcScreenBox">
                <div class="form-group">
                    <label>Título de la pantalla</label>
                    <input type="text" id="funcScreenTitle" placeholder="Ej. Filtro de la función" value="${UI.escapeHtml(screen.title || "")}">
                </div>
                <p class="form-hint">Variables que se pedirán al ejecutar la función; se pueden usar luego como "Variable" en los filtros de origen/destino o en el incremento de Revalue.</p>
                <div class="flow-screen-blocks" id="funcScreenBlocks"></div>
            </div>`;

        document.getElementById("funcScreenToggle").addEventListener("click", () => {
            this.screenCollapsed = !this.screenCollapsed;
            const box = document.getElementById("funcScreenBox");
            box.classList.toggle("is-collapsed", this.screenCollapsed);
            part.querySelector(".flow-group-caret").classList.toggle("is-collapsed", this.screenCollapsed);
        });

        document.getElementById("funcScreenTitle").addEventListener("input", (e) => { screen.title = e.target.value; });

        document.getElementById("funcAddVar").addEventListener("click", async () => {
            const v = await UI.openScreenVariableModal({ dimensions: this.dimensionsCache });
            if (!v) return;
            screen.blocks.push({ id: Provider.newId(), kind: "variable", variable: { id: Provider.newId(), ...v } });
            this.screenCollapsed = false;
            this.renderScreenBlocksList();
            this.renderConfigBlock();
        });
        document.getElementById("funcAddFrame").addEventListener("click", () => {
            screen.blocks.push({ id: Provider.newId(), kind: "frame", title: "Nuevo frame", variables: [] });
            this.screenCollapsed = false;
            this.renderScreenBlocksList();
        });
        document.getElementById("funcAddText").addEventListener("click", async () => {
            const text = await UI.openScreenTextModal({ current: "" });
            if (text === null) return;
            screen.blocks.push({ id: Provider.newId(), kind: "text", text });
            this.renderScreenBlocksList();
        });
        document.getElementById("funcAddSkip").addEventListener("click", () => {
            screen.blocks.push({ id: Provider.newId(), kind: "skip" });
            this.renderScreenBlocksList();
        });
        document.getElementById("funcAddLine").addEventListener("click", () => {
            screen.blocks.push({ id: Provider.newId(), kind: "line" });
            this.renderScreenBlocksList();
        });

        this.renderScreenBlocksList();
    },

    renderScreenBlocksList() {
        const wrap = document.getElementById("funcScreenBlocks");
        const screen = this.screen;
        if (!wrap) return;

        if (!screen.blocks.length) {
            wrap.innerHTML = `<div class="module-empty module-empty--inline">Sin variables: la función se lanzará solo con los filtros fijos definidos abajo.</div>`;
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
                        </div>`).join("") : `<p class="form-hint">Sin variables en este frame todavía.</p>`}
                </div>
            </div>`;
    },

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
        const wrap = document.getElementById("funcScreenBlocks");
        const screen = this.screen;

        wrap.querySelectorAll("[data-remove-block]").forEach(btn => btn.addEventListener("click", () => {
            const idx = parseInt(btn.dataset.removeBlock, 10);
            screen.blocks.splice(idx, 1);
            this.renderScreenBlocksList();
            this.renderConfigBlock();
        }));

        wrap.querySelectorAll("[data-add-frame-var]").forEach(btn => btn.addEventListener("click", async () => {
            const idx = parseInt(btn.dataset.addFrameVar, 10);
            const v = await UI.openScreenVariableModal({ dimensions: this.dimensionsCache });
            if (!v) return;
            screen.blocks[idx].variables.push({ id: Provider.newId(), ...v });
            this.renderScreenBlocksList();
            this.renderConfigBlock();
        }));

        wrap.querySelectorAll("[data-remove-frame-var]").forEach(btn => btn.addEventListener("click", () => {
            const [bIdx, vIdx] = btn.dataset.removeFrameVar.split(":").map(Number);
            screen.blocks[bIdx].variables.splice(vIdx, 1);
            this.renderScreenBlocksList();
            this.renderConfigBlock();
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
            const v = await UI.openScreenVariableModal({ current, dimensions: this.dimensionsCache });
            if (!v) return;
            screen.blocks[idx].variable = { ...current, ...v };
            this.renderScreenBlocksList();
            this.renderConfigBlock();
        }));

        wrap.querySelectorAll("[data-edit-frame-var]").forEach(el => el.addEventListener("click", async (e) => {
            e.stopPropagation();
            const [bIdx, vIdx] = el.dataset.editFrameVar.split(":").map(Number);
            const current = screen.blocks[bIdx].variables[vIdx];
            const v = await UI.openScreenVariableModal({ current, dimensions: this.dimensionsCache });
            if (!v) return;
            screen.blocks[bIdx].variables[vIdx] = { ...current, ...v };
            this.renderScreenBlocksList();
            this.renderConfigBlock();
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

    // ---------------- Bloque B: filtros (origen/destino/medidas/incremento) ----------------
    renderConfigBlock() {
        const part = document.getElementById("funcConfigPart");
        const tipo = this.editing.TIPO;
        const cube = this.currentCube;

        let html = `
            <div class="flow-screen-header"><strong>Filtros</strong>
                <span class="flow-screen-hint">Por defecto se aplican a todas las combinaciones ("All"/"Same"); cambia a valor(es) concretos o a una variable de la pantalla anterior.</span>
            </div>`;

        if (tipo === "DISTRIBUTION") {
            html += `<div class="module-empty module-empty--inline">La función Distribution estará disponible próximamente.</div>`;
        } else if (!cube.dims.length && !cube.measures.length) {
            html += `<div class="module-empty module-empty--inline">Este cubo no tiene dimensiones ni medidas configuradas.</div>`;
        } else {
            html += this.filterSectionHtml("origin", "Filtro origen", cube, this.config.origin.dims, "ALL");
            if (tipo === "COPY" || tipo === "MOVE") {
                html += this.filterSectionHtml("dest", "Filtro destino", cube, this.config.dest.dims, "SAME");
            }
            html += this.measuresSectionHtml(cube);
            if (tipo === "REVALUE") html += this.revalueSectionHtml();
        }

        part.innerHTML = html;
        this.bindConfigEvents();
    },

    FILTER_MODE_LABELS: {
        origin: [["ALL", "Todos (All)"], ["VALUES", "Valor(es)"], ["VAR", "Variable"]],
        dest: [["SAME", "Igual que origen (Same)"], ["VALUES", "Valor(es)"], ["VAR", "Variable"]]
    },

    filterSectionHtml(kind, title, cube, dimsCfg, defaultMode) {
        if (!cube.dims.length) return "";
        return `
            <div class="func-filter-block">
                <div class="flow-screen-header"><strong>${UI.escapeHtml(title)}</strong></div>
                <div class="func-filter-table">
                    ${cube.dims.map(d => this.filterRowHtml(kind, d, dimsCfg[d.colId] || { mode: defaultMode, values: [], varId: "" })).join("")}
                </div>
            </div>`;
    },

    filterRowHtml(kind, dim, cfg) {
        const vars = this.flatVars(this.screen);
        const modeOptions = this.FILTER_MODE_LABELS[kind];
        let detail = "";
        if (cfg.mode === "VALUES") {
            detail = `
                <span class="func-filter-summary">${cfg.values.length ? cfg.values.length + " valor(es): " + UI.escapeHtml(cfg.values.slice(0, 4).join(", ")) + (cfg.values.length > 4 ? "…" : "") : "Sin valores todavía"}</span>
                <button type="button" class="btn btn-secondary btn-sm func-filter-pick" data-kind="${kind}" data-dim="${dim.colId}">Seleccionar valores</button>`;
        } else if (cfg.mode === "VAR") {
            detail = `
                <select class="func-filter-var" data-kind="${kind}" data-dim="${dim.colId}">
                    <option value="">Selecciona variable...</option>
                    ${vars.map(v => `<option value="${v.id}" ${cfg.varId === v.id ? "selected" : ""}>${UI.escapeHtml(v.label || v.name)}</option>`).join("")}
                </select>`;
        }
        return `
            <div class="func-filter-row" data-kind="${kind}" data-dim="${dim.colId}">
                <span class="func-filter-dim">${UI.escapeHtml(dim.name)}</span>
                <select class="func-filter-mode" data-kind="${kind}" data-dim="${dim.colId}">
                    ${modeOptions.map(([val, label]) => `<option value="${val}" ${cfg.mode === val ? "selected" : ""}>${label}</option>`).join("")}
                </select>
                <div class="func-filter-detail">${detail}</div>
            </div>`;
    },

    measuresSectionHtml(cube) {
        const cfg = this.config;
        return `
            <div class="func-measures-block">
                <div class="flow-screen-header"><strong>Medidas</strong></div>
                <label class="func-measure-all"><input type="checkbox" id="funcMeasuresAll" ${cfg.measuresMode === "ALL" ? "checked" : ""}> Todas las medidas</label>
                <div class="func-measures-grid" id="funcMeasuresGrid" style="display:${cfg.measuresMode === "ALL" ? "none" : "flex"}">
                    ${cube.measures.length ? cube.measures.map(m => `
                        <label class="func-measure-chip"><input type="checkbox" class="func-measure-cb" value="${UI.escapeHtml(m.name)}" ${cfg.measures.includes(m.name) ? "checked" : ""}> ${UI.escapeHtml(m.name)}</label>`).join("")
                        : `<span class="form-hint">Este cubo no tiene medidas.</span>`}
                </div>
            </div>`;
    },

    revalueSectionHtml() {
        const r = this.config.revalue;
        const vars = this.flatVars(this.screen);
        return `
            <div class="func-revalue-block">
                <div class="flow-screen-header"><strong>Incremento</strong></div>
                <div class="func-revalue-row">
                    <select id="funcRevalueKind">
                        <option value="PERCENT" ${r.kind === "PERCENT" ? "selected" : ""}>Porcentaje (%)</option>
                        <option value="VALUE" ${r.kind === "VALUE" ? "selected" : ""}>Valor absoluto</option>
                    </select>
                    <select id="funcRevalueMode">
                        <option value="CONST" ${r.amountMode === "CONST" ? "selected" : ""}>Constante</option>
                        <option value="VAR" ${r.amountMode === "VAR" ? "selected" : ""}>Variable</option>
                    </select>
                    ${r.amountMode === "VAR"
                        ? `<select id="funcRevalueVar"><option value="">Selecciona variable...</option>${vars.map(v => `<option value="${v.id}" ${r.amountVarId === v.id ? "selected" : ""}>${UI.escapeHtml(v.label || v.name)}</option>`).join("")}</select>`
                        : `<input type="number" id="funcRevalueConst" step="any" placeholder="ej. 5 o -10" value="${UI.escapeHtml(r.amountConst || "")}">`}
                </div>
                <p class="form-hint">Con Porcentaje, 10 incrementa un 10% y -10 lo reduce un 10%. Se aplica igual a todas las medidas seleccionadas arriba.</p>
            </div>`;
    },

    bindConfigEvents() {
        const part = document.getElementById("funcConfigPart");
        if (!part) return;

        part.querySelectorAll(".func-filter-mode").forEach(sel => {
            sel.addEventListener("change", () => {
                const { kind, dim } = sel.dataset;
                const dimsCfg = kind === "origin" ? this.config.origin.dims : this.config.dest.dims;
                dimsCfg[dim] = dimsCfg[dim] || { mode: "ALL", values: [], varId: "" };
                dimsCfg[dim].mode = sel.value;
                this.renderConfigBlock();
            });
        });

        part.querySelectorAll(".func-filter-var").forEach(sel => {
            sel.addEventListener("change", () => {
                const { kind, dim } = sel.dataset;
                const dimsCfg = kind === "origin" ? this.config.origin.dims : this.config.dest.dims;
                dimsCfg[dim].varId = sel.value;
            });
        });

        part.querySelectorAll(".func-filter-pick").forEach(btn => {
            btn.addEventListener("click", async () => {
                const { kind, dim: colId } = btn.dataset;
                const dimsCfg = kind === "origin" ? this.config.origin.dims : this.config.dest.dims;
                const dimSpec = this.currentCube.dims.find(d => d.colId === colId);
                if (!dimSpec) return;
                const dimRow = this.dimensionsCache.find(d => d.DIMENSION_ID === dimSpec.id);
                if (!dimRow) { UI.toast("No se encuentra la dimensión de origen de esta columna.", "error"); return; }
                const keyCol = Provider.toIdentifier(dimRow.DIMENSION);
                const result = await UI.openDimensionValuesPickerModal({ project: this.project, dim: dimRow, keyCol, selected: dimsCfg[colId].values || [] });
                if (result === null) return;
                if (kind === "dest" && result.length > 1) {
                    UI.toast("El destino admite un único valor: se usará el primero seleccionado.", "info");
                }
                dimsCfg[colId].values = kind === "dest" ? result.slice(0, 1) : result;
                this.renderConfigBlock();
            });
        });

        const measuresAll = document.getElementById("funcMeasuresAll");
        if (measuresAll) {
            measuresAll.addEventListener("change", () => {
                this.config.measuresMode = measuresAll.checked ? "ALL" : "SOME";
                this.renderConfigBlock();
            });
        }
        part.querySelectorAll(".func-measure-cb").forEach(cb => {
            cb.addEventListener("change", () => {
                const set = new Set(this.config.measures);
                if (cb.checked) set.add(cb.value); else set.delete(cb.value);
                this.config.measures = Array.from(set);
            });
        });

        const revalueKind = document.getElementById("funcRevalueKind");
        if (revalueKind) revalueKind.addEventListener("change", () => { this.config.revalue.kind = revalueKind.value; });
        const revalueMode = document.getElementById("funcRevalueMode");
        if (revalueMode) revalueMode.addEventListener("change", () => { this.config.revalue.amountMode = revalueMode.value; this.renderConfigBlock(); });
        const revalueConst = document.getElementById("funcRevalueConst");
        if (revalueConst) revalueConst.addEventListener("input", () => { this.config.revalue.amountConst = revalueConst.value; });
        const revalueVar = document.getElementById("funcRevalueVar");
        if (revalueVar) revalueVar.addEventListener("change", () => { this.config.revalue.amountVarId = revalueVar.value; });
    },

    // ================================================================
    // GUARDAR
    // ================================================================
    async save() {
        const nameEl = document.getElementById("funcModalTitle");
        const name = nameEl ? nameEl.textContent.trim() : this.editing[this.NAME_COL];
        if (!name) { UI.toast("El nombre no puede estar vacío.", "error"); return; }

        const varsJsonPlain = JSON.stringify(this.screen);
        const configJsonPlain = JSON.stringify(this.config);
        const varsJson = varsJsonPlain.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
        const configJson = configJsonPlain.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
        const id = this.editing[this.ID_COL];

        try {
            if (this.editingIsNew) {
                await Provider.runQuery(`
                    INSERT INTO ${Provider.qualifyControl(this.TABLE)}
                    (${this.ID_COL}, PROYECTO_ID, ${this.NAME_COL}, DESCRIPCION, TIPO, CUBO_ID, CUBO_NOMBRE, CUBO_TABLA, VARIABLES_JSON, CONFIG_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                    VALUES ('${Provider.esc(id)}', '${Provider.esc(this.project.PROYECTO_ID)}', '${Provider.esc(name)}', '${Provider.esc(this.editing.DESCRIPCION || "")}',
                            '${Provider.esc(this.editing.TIPO)}', '${Provider.esc(this.editing.CUBO_ID)}', '${Provider.esc(this.editing.CUBO_NOMBRE || "")}', '${Provider.esc(this.editing.CUBO_TABLA)}',
                            '${varsJson}', '${configJson}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`);
                this.editingIsNew = false;
                UI.toast(`Función "${name}" creada.`, "success");
            } else {
                await Provider.runQuery(`
                    UPDATE ${Provider.qualifyControl(this.TABLE)}
                    SET ${this.NAME_COL} = '${Provider.esc(name)}', VARIABLES_JSON = '${varsJson}', CONFIG_JSON = '${configJson}', FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                    WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
                UI.toast(`Función "${name}" guardada.`, "success");
            }
            this.editing[this.NAME_COL] = name;
            this.editing.VARIABLES_JSON = varsJsonPlain;
            this.editing.CONFIG_JSON = configJsonPlain;
            await this.loadList();
            const runBtn = document.getElementById("funcEditorRun");
            if (runBtn) runBtn.disabled = false;
        } catch (err) {
            UI.toast("Error al guardar: " + err.message, "error");
        }
    },

    // ================================================================
    // EJECUCIÓN — pantalla de variables + Monitor (igual patrón que
    // flow_run.html): un único modal a pantalla completa con 2 pestañas.
    // ================================================================
    async startRun(record) {
        if (this.overlay) this.overlay.remove();
        this.stopPolling();
        this.runRecord = record;
        this.runScreen = this.parseScreen(record.VARIABLES_JSON);
        if (!this.cubes.length) await this.loadCubes();
        this.runCube = this.cubes.find(c => c.id === record.CUBO_ID) ||
            { id: record.CUBO_ID, name: record.CUBO_NOMBRE, table: record.CUBO_TABLA, dims: [], measures: [] };
        this.runConfig = this.safeParse(record.CONFIG_JSON, this.defaultConfig(record.TIPO, this.runCube));
        this.selOptState = {};

        const typeInfo = this.TYPES[record.TIPO] || { label: record.TIPO, icon: "•" };

        let overlay = document.getElementById("funcRunModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "funcRunModal";
            document.body.appendChild(overlay);
        }
        this.runOverlay = overlay;
        overlay.classList.add("visible");
        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div>
                        <h3>${UI.escapeHtml(this.runScreen.title || record[this.NAME_COL])}</h3>
                        <span class="modal-subtitle">${typeInfo.icon} ${UI.escapeHtml(typeInfo.label)} · Cubo ${UI.escapeHtml(this.runCube.name)}</span>
                    </div>
                    <button class="modal-close" id="funcRunClose">&times;</button>
                </div>
                <div class="flow-run-tabs" id="funcRunTabs">
                    <button type="button" class="flow-run-tab active" id="funcRunTabScreen">Pantalla de variables</button>
                    <button type="button" class="flow-run-tab" id="funcRunTabMonitor">Monitor</button>
                </div>
                <div class="modal-body modal-body-flush" id="funcRunBody"></div>
            </div>`;

        document.getElementById("funcRunClose").addEventListener("click", () => { this.stopPolling(); overlay.remove(); });
        document.getElementById("funcRunTabScreen").addEventListener("click", () => this.switchRunTab("screen"));
        document.getElementById("funcRunTabMonitor").addEventListener("click", () => this.switchRunTab("monitor"));

        this.switchRunTab("screen");
    },

    switchRunTab(tab) {
        this.runView = tab;
        document.getElementById("funcRunTabScreen").classList.toggle("active", tab === "screen");
        document.getElementById("funcRunTabMonitor").classList.toggle("active", tab === "monitor");
        if (tab === "screen") this.renderRunScreenView();
        else this.renderMonitorView();
    },

    // ---------------- Pestaña 1: pantalla de variables ----------------
    renderRunScreenView() {
        const body = document.getElementById("funcRunBody");
        const blocksHtml = this.runScreen.blocks.length
            ? this.runScreen.blocks.map(b => this.runBlockHtml(b)).join("")
            : `<div class="module-empty module-empty--inline">Esta función no tiene variables de pantalla definidas.</div>`;

        body.innerHTML = `
            <div class="flow-run-screen">
                <div class="flow-screen-blocks flow-screen-blocks--run">${blocksHtml}</div>
                <div class="flow-run-actions">
                    <button class="btn btn-primary" id="funcRunExecute">▶ Ejecutar función</button>
                    <span class="form-hint" id="funcRunHint"></span>
                </div>
            </div>`;

        document.getElementById("funcRunExecute").addEventListener("click", () => this.executeFunction());
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
        return `
            <div class="flow-screen-block flow-screen-block--frame flow-screen-block--static">
                <div class="flow-frame-header"><strong>${UI.escapeHtml(b.title || "Frame")}</strong></div>
                <div class="flow-frame-vars">
                    ${(b.variables || []).map(v => `<div class="flow-frame-var-row flow-frame-var-row--static">${this.runInputHtml(v)}</div>`).join("")}
                </div>
            </div>`;
    },

    runInputHtml(v) {
        if (v.selectMode && v.selectMode !== "unico") return this.selOptHtml(v);

        const id = `funcrunvar_${v.id}`;
        const label = `<label for="${id}">${UI.escapeHtml(v.label || v.name)}</label>`;
        if (v.type === "BOOLEAN") {
            return `<div class="flow-field-preview flow-field-preview--checkbox"><input type="checkbox" id="${id}" data-var-name="${UI.escapeHtml(v.name)}" data-var-type="BOOLEAN">${label}</div>`;
        }
        const htmlType = { INTEGER: "number", FLOAT: "number", NUMERIC: "number", DATE: "date", DATETIME: "datetime-local", TIMESTAMP: "datetime-local" }[v.type] || "text";
        return `<div class="flow-field-preview">${label}<input type="${htmlType}" id="${id}" data-var-name="${UI.escapeHtml(v.name)}" data-var-type="${UI.escapeHtml(v.type || "STRING")}"></div>`;
    },

    // ---- select-options (rango / varios valores / cualquiera), igual que en flow_run.js / table-updates.js ----
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

    selOptDefaultRow(mode) {
        return { sign: "I", option: mode === "rango" ? "BT" : "EQ", low: "", high: "" };
    },

    selOptNeedsHigh(option) {
        return option === "BT" || option === "NB";
    },

    selOptHtml(v) {
        if (!this.selOptState[v.id]) this.selOptState[v.id] = [this.selOptDefaultRow(v.selectMode)];
        return `<div class="flow-field-preview flow-field-preview--selopt" id="func_selopt_wrap_${v.id}">${this.selOptInnerHtml(v)}</div>`;
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
        const wrap = document.getElementById(`func_selopt_wrap_${v.id}`);
        if (!wrap) return;
        wrap.innerHTML = this.selOptInnerHtml(v);
    },

    findRunVariableById(varId) {
        return this.flatVars(this.runScreen).find(v => v.id === varId) || null;
    },

    varNameById(id) {
        const v = this.flatVars(this.runScreen).find(x => x.id === id);
        return v ? v.name : id;
    },

    /** Recoge del DOM {nombre: valor} para las variables de la pantalla de ejecución. */
    collectRunScreenValues() {
        const values = {};
        document.querySelectorAll("#funcRunBody [data-var-name]").forEach(el => {
            const name = el.dataset.varName;
            const type = el.dataset.varType;
            values[name] = type === "BOOLEAN" ? el.checked : el.value;
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

    /** Valor(es) resueltos de una variable de pantalla ({sign/option/low/high}[] o string). */
    resolveVarValue(varId, values) {
        const name = this.varNameById(varId);
        return values[name];
    },

    /** Condición SQL para una columna según su configuración de filtro
     *  (ALL/SAME -> null: sin condición; VALUES -> IN (...); VAR -> según
     *  el valor resuelto de la variable, admite select-options). */
    filterClauseSql(colId, cfg, values) {
        if (!cfg || cfg.mode === "ALL" || cfg.mode === "SAME") return null;
        if (cfg.mode === "VALUES") {
            const vals = (cfg.values || []).filter(v => v !== null && v !== undefined && v !== "");
            if (!vals.length) return null;
            return `${colId} IN (${vals.map(v => `'${Provider.esc(v)}'`).join(", ")})`;
        }
        if (cfg.mode === "VAR" && cfg.varId) {
            const val = this.resolveVarValue(cfg.varId, values);
            if (Array.isArray(val)) return this.selOptTableSql(colId, val);
            if (val !== null && val !== undefined && val !== "") return `${colId} = '${Provider.esc(val)}'`;
        }
        return null;
    },

    /** Valor único resuelto para un destino en modo VALUES/VAR (nunca SAME/ALL). */
    destValueSql(cfg, values) {
        if (cfg.mode === "VALUES") {
            const v = (cfg.values || [])[0];
            return v !== undefined && v !== null && v !== "" ? `'${Provider.esc(v)}'` : null;
        }
        if (cfg.mode === "VAR" && cfg.varId) {
            const val = this.resolveVarValue(cfg.varId, values);
            const plain = Array.isArray(val) ? (val[0] && val[0].low) : val;
            return plain !== undefined && plain !== null && plain !== "" ? `'${Provider.esc(plain)}'` : null;
        }
        return null;
    },

    buildOriginWhere(cube, config, values) {
        const clauses = [];
        cube.dims.forEach(d => {
            const cfg = config.origin.dims[d.colId];
            const c = this.filterClauseSql(d.colId, cfg, values);
            if (c) clauses.push(c);
        });
        return clauses;
    },

    selectedMeasureNames(cube, config) {
        return config.measuresMode === "ALL" ? cube.measures.map(m => m.name) : (config.measures || []);
    },

    /**
     * Genera el SQL a ejecutar según el tipo de función. Devuelve null si
     * el tipo todavía no tiene generador (Distribution) o si faltan datos
     * imprescindibles (se avisa con un toast desde executeFunction()).
     */
    buildFunctionSql(record, cube, config, values) {
        const table = Provider.qualify(this.project.DATASET, cube.table);
        const originClauses = this.buildOriginWhere(cube, config, values);
        const whereOrigin = originClauses.length ? `WHERE ${originClauses.join(" AND ")}` : "";
        const measures = this.selectedMeasureNames(cube, config);

        if (record.TIPO === "DELETE") {
            if (config.measuresMode === "ALL") {
                return `DELETE FROM ${table} ${whereOrigin}`.trim();
            }
            if (!measures.length) return null;
            const sets = measures.map(m => `${m} = NULL`).join(", ");
            return `UPDATE ${table} SET ${sets} ${whereOrigin}`.trim();
        }

        if (record.TIPO === "REVALUE") {
            if (!measures.length) return null;
            const r = config.revalue || {};
            let amountExpr;
            if (r.amountMode === "CONST") {
                const n = parseFloat(r.amountConst);
                if (isNaN(n)) return null;
                amountExpr = String(n);
            } else if (r.amountMode === "VAR" && r.amountVarId) {
                const raw = this.resolveVarValue(r.amountVarId, values);
                const plain = Array.isArray(raw) ? (raw[0] && raw[0].low) : raw;
                const n = parseFloat(plain);
                if (isNaN(n)) return null;
                amountExpr = String(n);
            } else {
                return null;
            }
            const sets = measures.map(m => r.kind === "PERCENT"
                ? `${m} = ${m} * (1 + (${amountExpr}) / 100)`
                : `${m} = ${m} + (${amountExpr})`).join(", ");
            return `UPDATE ${table} SET ${sets} ${whereOrigin}`.trim();
        }

        if (record.TIPO === "MOVE") {
            const sets = [];
            cube.dims.forEach(d => {
                const destCfg = config.dest.dims[d.colId];
                if (!destCfg || destCfg.mode === "SAME") return;
                const expr = this.destValueSql(destCfg, values);
                if (expr) sets.push(`${d.colId} = ${expr}`);
            });
            if (!sets.length) return null;
            return `UPDATE ${table} SET ${sets.join(", ")} ${whereOrigin}`.trim();
        }

        if (record.TIPO === "COPY") {
            const selectExprs = [];
            const destWhere = [];
            cube.dims.forEach(d => {
                const destCfg = config.dest.dims[d.colId] || { mode: "SAME" };
                if (destCfg.mode === "SAME") {
                    selectExprs.push(d.colId);
                    const originCfg = config.origin.dims[d.colId];
                    const originClause = this.filterClauseSql(d.colId, originCfg, values);
                    if (originClause) destWhere.push(originClause);
                } else {
                    const expr = this.destValueSql(destCfg, values);
                    if (!expr) { selectExprs.push(d.colId); return; }
                    selectExprs.push(expr);
                    destWhere.push(`${d.colId} = ${expr}`);
                }
            });
            const cols = [...cube.dims.map(d => d.colId), ...measures];
            const selects = [...selectExprs, ...measures];
            if (!measures.length) return null;

            const statements = [];
            if (destWhere.length) {
                statements.push(`DELETE FROM ${table} WHERE ${destWhere.join(" AND ")}`);
            }
            statements.push(`INSERT INTO ${table} (${cols.join(", ")})\nSELECT ${selects.join(", ")}\nFROM ${table} ${whereOrigin}`);
            return statements.join(";\n");
        }

        return null;
    },

    // ---------------- Ejecutar + Monitor ----------------
    async executeFunction() {
        const btn = document.getElementById("funcRunExecute");
        const hint = document.getElementById("funcRunHint");
        if (btn) btn.disabled = true;

        try {
            const values = this.collectRunScreenValues();
            const sql = this.buildFunctionSql(this.runRecord, this.runCube, this.runConfig, values);
            if (!sql) {
                UI.toast("No se ha podido generar el SQL: revisa que haya medidas seleccionadas y, si aplica, un incremento válido.", "error");
                return;
            }

            const runId = Provider.newId();
            this.lastRunId = runId;
            const sqlEsc = sql.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
            const varsEsc = JSON.stringify(values).replace(/\\/g, "\\\\").replace(/'/g, "\\'");

            await Provider.runQuery(`
                INSERT INTO ${Provider.qualifyControl(this.RUN_TABLE)}
                (RUN_ID, PROYECTO_ID, FUNCION_ID, ESTADO, SQL_TEXT, VARIABLES_JSON, USUARIO, FECHA_INICIO)
                VALUES ('${Provider.esc(runId)}', '${Provider.esc(this.project.PROYECTO_ID)}', '${Provider.esc(this.runRecord[this.ID_COL])}',
                        'EN_CURSO', '${sqlEsc}', '${varsEsc}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP())`);

            // Se lanza sin esperar: la propia llamada a Provider.runQuery ya es
            // asíncrona (red), así que el SQL sigue en marcha aunque el
            // usuario cambie de pestaña; el Monitor sondea el estado en
            // FUNCIONES_RUNS, igual que el monitor de Flujos de carga.
            this.runInBackground(runId, sql);

            UI.toast(`"${this.runRecord[this.NAME_COL]}" lanzada.`, "success");
            if (hint) hint.textContent = "";
            this.switchRunTab("monitor");
        } catch (err) {
            UI.toast("Error al lanzar la función: " + err.message, "error");
        } finally {
            if (btn) btn.disabled = false;
        }
    },

    async runInBackground(runId, sql) {
        try {
            // Varias sentencias separadas por ';' (caso Copy: DELETE + INSERT).
            const statements = sql.split(/;\s*\n/).map(s => s.trim()).filter(Boolean);
            for (const stmt of statements) {
                await Provider.runQuery(stmt);
            }
            await Provider.runQuery(`
                UPDATE ${Provider.qualifyControl(this.RUN_TABLE)}
                SET ESTADO = 'OK', FECHA_FIN = CURRENT_TIMESTAMP()
                WHERE RUN_ID = '${Provider.esc(runId)}'`);
        } catch (err) {
            const msg = Provider.esc((err.message || "Error desconocido").slice(0, 900)).replace(/'/g, "\\'");
            try {
                await Provider.runQuery(`
                    UPDATE ${Provider.qualifyControl(this.RUN_TABLE)}
                    SET ESTADO = 'ERROR', MENSAJE = '${msg}', FECHA_FIN = CURRENT_TIMESTAMP()
                    WHERE RUN_ID = '${Provider.esc(runId)}'`);
            } catch (e2) {
                console.error("No se pudo registrar el error de la función:", e2);
            }
        } finally {
            if (this.runView === "monitor" && document.getElementById("funcMonitorChainWrap")) this.refreshMonitor();
        }
    },

    async renderMonitorView() {
        const body = document.getElementById("funcRunBody");
        body.innerHTML = `
            <div class="flow-run-monitor">
                <div class="flow-chain-wrap" id="funcMonitorChainWrap"><span class="spinner"></span></div>
                <div class="flow-run-monitor-footer" id="funcMonitorFooter"></div>
            </div>`;
        await this.refreshMonitor();
        this.stopPolling();
        this.pollHandle = setInterval(() => this.refreshMonitor(), 3000);
    },

    stopPolling() {
        if (this.pollHandle) { clearInterval(this.pollHandle); this.pollHandle = null; }
    },

    async refreshMonitor() {
        const wrap = document.getElementById("funcMonitorChainWrap");
        const footer = document.getElementById("funcMonitorFooter");
        if (!wrap || !footer) { this.stopPolling(); return; }

        try {
            const rows = await Provider.runQuery(`
                SELECT RUN_ID, ESTADO, MENSAJE, FECHA_INICIO, FECHA_FIN FROM ${Provider.qualifyControl(this.RUN_TABLE)}
                WHERE FUNCION_ID = '${Provider.esc(this.runRecord[this.ID_COL])}' ORDER BY FECHA_INICIO DESC LIMIT 1`);
            const run = rows[0];

            if (!run) {
                wrap.innerHTML = `<div class="module-empty module-empty--inline">Todavía no se ha ejecutado esta función.</div>`;
                footer.innerHTML = "";
                this.stopPolling();
                return;
            }

            const estado = run.ESTADO;
            const typeInfo = this.TYPES[this.runRecord.TIPO] || { label: this.runRecord.TIPO, icon: "•" };
            wrap.innerHTML = `
                <div class="flow-chain-card flow-run-step flow-run-step--${estado.toLowerCase()}" title="${run.MENSAJE ? UI.escapeHtml(run.MENSAJE) : ""}">
                    <div class="flow-chain-card-name">${typeInfo.icon} ${UI.escapeHtml(typeInfo.label)}</div>
                    <div class="flow-chain-card-meta">${this.stepStatusLabel(estado)}</div>
                </div>`;

            const badge = { EN_CURSO: "table-tag flow-status-scheduled", OK: "table-tag flow-status-ok", ERROR: "table-tag flow-status-error" }[estado] || "table-tag";
            footer.innerHTML = `
                <span class="${badge}">${this.stepStatusLabel(estado)}</span>
                <span class="form-hint">Run ${UI.escapeHtml(run.RUN_ID)}</span>
                ${estado === "ERROR" && run.MENSAJE ? `<div class="form-hint flow-run-error-msg">${UI.escapeHtml(run.MENSAJE)}</div>` : ""}`;

            if (estado !== "EN_CURSO") this.stopPolling();
        } catch (err) {
            footer.innerHTML = `<span class="form-hint">Error consultando el estado: ${UI.escapeHtml(err.message)}</span>`;
            this.stopPolling();
        }
    },

    stepStatusLabel(estado) {
        return { EN_CURSO: "▶ En ejecución", OK: "✔ Completado", ERROR: "✕ Error" }[estado] || estado;
    }
};
