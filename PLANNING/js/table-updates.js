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
    variables: [],
    fields: [],
    dragFieldIdx: null,
    dragVarIdx: null,
    screenCollapsed: false,
    dimensionsCache: [],

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
        this.variables = this.safeParse(record.VARIABLES_JSON, []);
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

    // ---------------- Bloque A: pantalla de selección (variables) ----------------
    renderScreenBlock() {
        const part = document.getElementById("actUpdScreenPart");
        part.innerHTML = `
            <div class="flow-screen-header" id="actUpdScreenToggle">
                <span class="flow-screen-caret">${this.screenCollapsed ? "▶" : "▼"}</span>
                <strong>Pantalla de selección</strong>
                <span class="flow-screen-hint">Variables para filtrar qué filas se traen a editar</span>
            </div>
            <div class="flow-screen-body" id="actUpdScreenBody" style="${this.screenCollapsed ? "display:none;" : ""}">
                <div class="fields-builder-rows" id="actUpdVarsRows"></div>
                <button class="btn btn-secondary btn-sm" id="actUpdAddVar">+ Añadir variable</button>
            </div>`;

        document.getElementById("actUpdScreenToggle").addEventListener("click", () => {
            this.screenCollapsed = !this.screenCollapsed;
            this.renderScreenBlock();
        });
        document.getElementById("actUpdAddVar").addEventListener("click", () => {
            this.variables.push({ id: Provider.newId(), name: "", label: "", type: "STRING" });
            this.renderScreenBlock();
        });

        this.renderVarRows();
    },

    renderVarRows() {
        const rowsEl = document.getElementById("actUpdVarsRows");
        if (!rowsEl) return;
        rowsEl.innerHTML = this.variables.map((v, idx) => `
            <div class="field-row var-row" draggable="true" data-idx="${idx}">
                <input type="text" class="var-name" placeholder="nombre técnico (ej. SOCIEDAD)" value="${UI.escapeHtml(v.name)}">
                <input type="text" class="var-label" placeholder="Etiqueta a mostrar" value="${UI.escapeHtml(v.label)}">
                <select class="var-type">
                    ${["STRING", "INTEGER", "DATE"].map(t => `<option value="${t}" ${t === v.type ? "selected" : ""}>${t}</option>`).join("")}
                </select>
                <button type="button" class="field-remove" title="Eliminar">✕</button>
            </div>`).join("") || `<div class="hierarchy-levels-empty">Sin variables: la tabla se cargará entera al ejecutar (salvo filtros constantes por campo).</div>`;

        rowsEl.querySelectorAll(".var-row").forEach(row => {
            const idx = parseInt(row.dataset.idx, 10);
            row.querySelector(".var-name").addEventListener("input", (e) => { this.variables[idx].name = e.target.value; });
            row.querySelector(".var-label").addEventListener("input", (e) => { this.variables[idx].label = e.target.value; });
            row.querySelector(".var-type").addEventListener("change", (e) => { this.variables[idx].type = e.target.value; });
            row.querySelector(".field-remove").addEventListener("click", () => {
                this.variables.splice(idx, 1);
                this.renderVarRows();
            });
            row.addEventListener("dragstart", () => { this.dragVarIdx = idx; row.classList.add("dragging"); });
            row.addEventListener("dragend", () => row.classList.remove("dragging"));
            row.addEventListener("dragover", (e) => e.preventDefault());
            row.addEventListener("drop", (e) => {
                e.preventDefault();
                if (this.dragVarIdx === null || this.dragVarIdx === idx) return;
                const [moved] = this.variables.splice(this.dragVarIdx, 1);
                this.variables.splice(idx, 0, moved);
                this.dragVarIdx = null;
                this.renderVarRows();
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
                ${this.variables.map(v => `<option value="${v.id}" ${f.filter.value === v.id ? "selected" : ""}>${UI.escapeHtml(v.label || v.name)}</option>`).join("")}
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
        const varsJson = JSON.stringify(this.variables).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
        const fieldsJson = JSON.stringify(this.fields).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
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
            this.editing[this.NAME_COL] = name;
            await this.loadList();
            const runBtn = document.getElementById("actUpdEditorRun");
            if (runBtn) runBtn.disabled = false;
        } catch (err) {
            UI.toast("Error al guardar: " + err.message, "error");
        }
    },

    // ================================================================
    // EJECUCIÓN: pantalla de variables -> grid
    // ================================================================
    async startRun(record) {
        if (this.overlay) this.overlay.remove();
        const variables = this.safeParse(record.VARIABLES_JSON, []);
        const fields = this.safeParse(record.CAMPOS_JSON, []).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));

        if (!variables.length) {
            this.loadAndOpenGrid(record, fields, {});
            return;
        }

        let overlay = document.getElementById("actUpdRunVarsModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "actUpdRunVarsModal";
            document.body.appendChild(overlay);
        }
        overlay.classList.add("visible");
        overlay.innerHTML = `
            <div class="modal-box">
                <div class="modal-header">
                    <h3>${UI.escapeHtml(record[this.NAME_COL])} — Variables</h3>
                    <button class="modal-close" id="actUpdRunVarsClose">&times;</button>
                </div>
                <div class="modal-body">
                    ${variables.map(v => `
                        <div class="form-group">
                            <label>${UI.escapeHtml(v.label || v.name)}</label>
                            <input type="${v.type === "DATE" ? "date" : (v.type === "INTEGER" ? "number" : "text")}" class="actupd-run-var" data-name="${UI.escapeHtml(v.name)}">
                        </div>`).join("")}
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="actUpdRunVarsCancel">Cancelar</button>
                    <button class="btn btn-primary" id="actUpdRunVarsContinue">Continuar</button>
                </div>
            </div>`;
        const close = () => overlay.remove();
        document.getElementById("actUpdRunVarsClose").addEventListener("click", close);
        document.getElementById("actUpdRunVarsCancel").addEventListener("click", close);
        document.getElementById("actUpdRunVarsContinue").addEventListener("click", () => {
            const values = {};
            overlay.querySelectorAll(".actupd-run-var").forEach(inp => { values[inp.dataset.name] = inp.value; });
            close();
            this.loadAndOpenGrid(record, fields, values);
        });
    },

    buildWhere(fields, variableValues) {
        const clauses = [];
        fields.forEach(f => {
            if (!f.filter || f.filter.type === "NONE") return;
            let val = null;
            if (f.filter.type === "CONST") val = f.filter.value;
            else if (f.filter.type === "VAR") val = variableValues[this.varNameById(f.filter.value)];
            if (val === null || val === undefined || val === "") return;
            clauses.push(`${f.name} = '${Provider.esc(val)}'`);
        });
        return clauses.length ? `WHERE ${clauses.join(" AND ")}` : "";
    },

    varNameById(id) {
        // helper resuelto en tiempo de ejecución desde this._runVariables si existe
        const v = (this._runVariables || []).find(x => x.id === id);
        return v ? v.name : id;
    },

    async loadAndOpenGrid(record, fields, variableValues) {
        this._runVariables = this.safeParse(record.VARIABLES_JSON, []);
        const where = this.buildWhere(fields, variableValues);
        const table = Provider.qualify(this.project.DATASET, record.TABLA);

        UI.toast("Cargando datos...", "info");
        try {
            const rows = await Provider.runQuery(`SELECT * FROM ${table} ${where}`);
            await this.openGrid(record, fields, rows, where);
        } catch (err) {
            UI.toast("Error al cargar la tabla: " + err.message, "error");
        }
    },

    // ================================================================
    // GRID (tipo ALV): buscar, filtrar, ordenar, editar, fichero, grabar
    // ================================================================
    async openGrid(record, fields, rows, where) {
        const state = {
            record, fields, where,
            columns: fields.map(f => f.name),
            originalRows: rows.map(r => ({ ...r })),
            currentRows: rows.map(r => ({ ...r, __rowId: Provider.newId(), __isNew: false })),
            deletedRowIds: [],
            search: "",
            sortCol: null, sortDir: 1,
            visibleCount: this.RENDER_CHUNK,
            optionsByField: {}
        };
        this.gridState = state;

        // Precarga las opciones de validación (constante/dimensión/jerarquía)
        // una sola vez por campo, para no volver a consultar en cada celda.
        for (const f of fields) {
            if (f.validation && f.validation.type !== "NONE") {
                state.optionsByField[f.name] = await this.resolveValidationOptions(f.validation);
            }
        }

        let overlay = document.getElementById("actUpdGridModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "actUpdGridModal";
            document.body.appendChild(overlay);
        }
        overlay.classList.add("visible");
        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <h3>${UI.escapeHtml(record[this.NAME_COL])}</h3>
                    <button class="modal-close" id="actUpdGridClose">&times;</button>
                </div>
                <div class="modal-body modal-body-flush">
                    <div class="actupd-grid-toolbar">
                        <input type="text" id="actUpdGridSearch" placeholder="Buscar en todas las columnas...">
                        <button class="btn btn-secondary btn-sm" id="actUpdGridAddRow">+ Añadir fila</button>
                        <button class="btn btn-secondary btn-sm" id="actUpdGridDelRows">Eliminar seleccionadas</button>
                        <button class="btn btn-secondary btn-sm" id="actUpdGridDownload">⇩ Descargar</button>
                        <button class="btn btn-secondary btn-sm" id="actUpdGridUpload">⇧ Cargar fichero</button>
                        <input type="file" id="actUpdGridFileInput" accept=".xlsx,.xls,.csv" style="display:none;">
                        <span class="actupd-grid-count" id="actUpdGridCount"></span>
                        <button class="btn btn-primary btn-sm" id="actUpdGridSave">Grabar</button>
                    </div>
                    <div class="actupd-grid-wrap" id="actUpdGridWrap"><span class="spinner"></span></div>
                </div>
            </div>`;

        document.getElementById("actUpdGridClose").addEventListener("click", () => overlay.remove());
        document.getElementById("actUpdGridSearch").addEventListener("input", (e) => {
            state.search = e.target.value.toLowerCase();
            state.visibleCount = this.RENDER_CHUNK;
            this.renderGrid();
        });
        document.getElementById("actUpdGridAddRow").addEventListener("click", () => {
            const blank = {};
            state.columns.forEach(c => { blank[c] = ""; });
            state.currentRows.push({ ...blank, __rowId: Provider.newId(), __isNew: true });
            this.renderGrid();
        });
        document.getElementById("actUpdGridDelRows").addEventListener("click", () => this.deleteSelectedRows());
        document.getElementById("actUpdGridDownload").addEventListener("click", () => this.downloadGrid());
        document.getElementById("actUpdGridUpload").addEventListener("click", () => document.getElementById("actUpdGridFileInput").click());
        document.getElementById("actUpdGridFileInput").addEventListener("change", (e) => this.uploadGrid(e));
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
        if (state.search) {
            rows = rows.filter(r => state.columns.some(c => String(r[c] ?? "").toLowerCase().includes(state.search)));
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

    renderGrid() {
        const state = this.gridState;
        const wrap = document.getElementById("actUpdGridWrap");
        const filtered = this.getFilteredSortedRows();
        const visible = filtered.slice(0, state.visibleCount);

        const countEl = document.getElementById("actUpdGridCount");
        if (countEl) countEl.textContent = `${filtered.length} de ${state.currentRows.length} filas`;

        const headerCells = state.fields.map(f => {
            const arrow = state.sortCol === f.name ? (state.sortDir === 1 ? " ▲" : " ▼") : "";
            return `<th data-sort="${UI.escapeHtml(f.name)}" title="${UI.escapeHtml(f.description || "")}">${UI.escapeHtml(f.description || f.name)}${arrow}</th>${f.validation && f.validation.showText ? `<th class="actupd-desc-col">Texto</th>` : ""}`;
        }).join("");

        const rowsHtml = visible.map(row => {
            const cells = state.fields.map(f => this.renderCell(row, f)).join("");
            return `<tr data-row="${row.__rowId}"><td class="actupd-sel-col"><input type="checkbox" class="actupd-row-check" data-row="${row.__rowId}"></td>${cells}</tr>`;
        }).join("");

        wrap.innerHTML = `
            <table class="actupd-grid">
                <thead><tr><th class="actupd-sel-col"></th>${headerCells}</tr></thead>
                <tbody>${rowsHtml || `<tr><td colspan="99">Sin filas.</td></tr>`}</tbody>
            </table>
            ${filtered.length > visible.length ? `<button class="btn btn-secondary btn-sm" id="actUpdGridMore">Mostrar más (${filtered.length - visible.length} restantes)</button>` : ""}
        `;

        wrap.querySelectorAll("th[data-sort]").forEach(th => {
            th.addEventListener("click", () => {
                const col = th.dataset.sort;
                state.sortDir = state.sortCol === col ? -state.sortDir : 1;
                state.sortCol = col;
                this.renderGrid();
            });
        });
        const moreBtn = document.getElementById("actUpdGridMore");
        if (moreBtn) moreBtn.addEventListener("click", () => { state.visibleCount += this.RENDER_CHUNK; this.renderGrid(); });

        wrap.querySelectorAll(".actupd-cell-input").forEach(el => {
            const evt = el.tagName === "SELECT" || el.type === "checkbox" ? "change" : "input";
            el.addEventListener(evt, (e) => this.onCellChange(e, el));
        });
        wrap.querySelectorAll(".actupd-search-trigger").forEach(btn => {
            btn.addEventListener("click", () => this.openSearchHelp(btn.dataset.row, btn.dataset.field));
        });
    },

    renderCell(row, f) {
        const val = row[f.name] ?? "";
        const v = f.validation;
        let inputHtml;

        if (!v || v.type === "NONE") {
            inputHtml = `<input type="text" class="actupd-cell-input" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}" value="${UI.escapeHtml(val)}">`;
        } else {
            const options = this.gridState.optionsByField[f.name] || [];
            if (v.searchHelp === "CHECKBOX" && options.length === 2) {
                const checked = String(val) === String(options[1].id);
                inputHtml = `<input type="checkbox" class="actupd-cell-input" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}" data-off="${UI.escapeHtml(options[0].id)}" data-on="${UI.escapeHtml(options[1].id)}" ${checked ? "checked" : ""}>`;
            } else if (v.searchHelp === "SEARCH") {
                inputHtml = `<span class="actupd-search-cell">
                    <input type="text" class="actupd-cell-input" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}" value="${UI.escapeHtml(val)}">
                    <button type="button" class="actupd-search-trigger" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}" title="Buscar">🔍</button>
                </span>`;
            } else {
                inputHtml = `<select class="actupd-cell-input" data-row="${row.__rowId}" data-field="${UI.escapeHtml(f.name)}">
                    <option value="">${v.allowEmpty ? "(vacío)" : "— selecciona —"}</option>
                    ${options.map(o => `<option value="${UI.escapeHtml(o.id)}" ${String(val) === String(o.id) ? "selected" : ""}>${UI.escapeHtml(o.id)}${o.desc ? " — " + UI.escapeHtml(o.desc) : ""}</option>`).join("")}
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
        overlay.classList.add("visible");
        overlay.innerHTML = `
            <div class="modal-box">
                <div class="modal-header"><h3>Buscar valor — ${UI.escapeHtml(field)}</h3><button class="modal-close" id="actUpdSHClose">&times;</button></div>
                <div class="modal-body">
                    <input type="text" id="actUpdSHSearch" placeholder="Buscar por ID o descripción...">
                    <div id="actUpdSHResults" class="actupd-sh-results"></div>
                </div>
            </div>`;
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

    deleteSelectedRows() {
        const state = this.gridState;
        const ids = Array.from(document.querySelectorAll(".actupd-row-check:checked")).map(cb => cb.dataset.row);
        if (!ids.length) { UI.toast("Selecciona al menos una fila.", "info"); return; }
        ids.forEach(id => {
            const row = state.currentRows.find(r => r.__rowId === id);
            if (row && !row.__isNew) state.deletedRowIds.push(id);
        });
        state.currentRows = state.currentRows.filter(r => !ids.includes(r.__rowId));
        this.renderGrid();
    },

    downloadGrid() {
        const state = this.gridState;
        const data = state.currentRows.map(r => {
            const out = {};
            state.columns.forEach(c => { out[c] = r[c]; });
            return out;
        });
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Datos");
        XLSX.writeFile(wb, `${state.record.TABLA}.xlsx`);
    },

    /**
     * Carga desde fichero: SUSTITUYE el contenido del grid. Para conservar
     * la identidad de las filas ya existentes (y así poder distinguir
     * "modificada" de "borrada + añadida" en el resumen de cambios), se
     * empareja cada fila del fichero con una fila original por el valor de
     * su PRIMERA columna (asumida como identificador natural, patrón que
     * siguen las tablas de Draco). Si no hay match, se trata como fila
     * nueva; cualquier fila original sin match en el fichero se marca
     * como eliminada.
     */
    uploadGrid(e) {
        const file = e.target.files[0];
        if (!file) return;
        const state = this.gridState;
        const reader = new FileReader();
        reader.onload = async (ev) => {
            try {
                const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const parsed = XLSX.utils.sheet_to_json(ws, { defval: "" });

                const keyCol = state.columns[0];
                const originalByKey = new Map(state.originalRows.map(r => [String(r[keyCol]), r]));
                const stillPresentKeys = new Set();

                const newCurrent = parsed.map(pRow => {
                    const keyVal = String(pRow[keyCol]);
                    const original = originalByKey.get(keyVal);
                    stillPresentKeys.add(keyVal);
                    const rowId = original ? (state.currentRows.find(r => String(r[keyCol]) === keyVal)?.__rowId || Provider.newId()) : Provider.newId();
                    const row = { __rowId: rowId, __isNew: !original };
                    state.columns.forEach(c => { row[c] = pRow[c] !== undefined ? pRow[c] : ""; });
                    return row;
                });

                state.originalRows.forEach(r => {
                    if (!stillPresentKeys.has(String(r[keyCol]))) {
                        const existing = state.currentRows.find(cr => String(cr[keyCol]) === String(r[keyCol]) && !cr.__isNew);
                        if (existing) state.deletedRowIds.push(existing.__rowId);
                    }
                });

                state.currentRows = newCurrent;
                state.visibleCount = this.RENDER_CHUNK;
                UI.toast(`Fichero cargado: ${parsed.length} filas.`, "success");
                this.renderGrid();
            } catch (err) {
                UI.toast("Error al leer el fichero: " + err.message, "error");
            }
            e.target.value = "";
        };
        reader.readAsArrayBuffer(file);
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

    async commitGrid() {
        const state = this.gridState;

        // Valida "permite vacío" antes de grabar.
        const violations = [];
        state.fields.forEach(f => {
            if (f.validation && f.validation.type !== "NONE" && f.validation.allowEmpty === false) {
                state.currentRows.forEach(r => {
                    if (r[f.name] === "" || r[f.name] === null || r[f.name] === undefined) {
                        violations.push({ campo: f.name, row: r });
                    }
                });
            }
        });
        if (violations.length) {
            UI.toast(`No se puede grabar: ${violations.length} valor(es) vacío(s) en campos que no lo permiten (${[...new Set(violations.map(v => v.campo))].join(", ")}).`, "error");
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

        try {
            const table = Provider.qualify(this.project.DATASET, state.record.TABLA);
            // Estrategia simple y robusta (no requiere clave garantizada):
            // se borra el subconjunto filtrado y se reinserta completo.
            await Provider.runQuery(`DELETE FROM ${table} ${state.where}`);

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
        overlay.classList.add("visible");
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
