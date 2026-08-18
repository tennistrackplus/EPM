/**
 * ============================================================
 * DRACO PLANNING — INTERFACES (MOCKUP)
 * ============================================================
 * Listado de interfaces (cargas de datos) agrupadas por cubo, con
 * alta/edición en dos pasos:
 *   1) Modal pequeño: nombre, origen (tabla/fichero) y cubo destino.
 *   2) Modal casi a pantalla completa (mapeo), que ya no repite esos
 *      3 campos como inputs — se muestran en el título para ganar
 *      espacio vertical.
 * Todo lo relativo al origen, el mapeo y las funciones Python es,
 * por ahora, MOCKUP: se guarda en localStorage, no en las tablas de
 * control reales. Cuando el mockup guste, se decide cómo persistirlo.
 */
const Loads = {
    list: [],
    cubes: [],
    filterCuboId: "",
    editing: null,       // copia de trabajo de la carga que se está editando
    editingIsNew: true,
    dragFieldName: null,
    fileParamsCollapsed: false,

    TYPE_ABBR: { STRING: "Str", INTEGER: "Int", FLOAT: "Flt", NUMERIC: "Num", BOOLEAN: "Bool", DATE: "Date", DATETIME: "DtTm", TIMESTAMP: "Ts" },

    MOCK_TABLES: {
        bigquery: [
            { name: "raw_erp.ventas_linea", fields: [{ name: "id_venta", type: "STRING" }, { name: "fecha", type: "DATE" }, { name: "id_cliente", type: "STRING" }, { name: "id_producto", type: "STRING" }, { name: "cantidad", type: "INTEGER" }, { name: "importe", type: "FLOAT" }] },
            { name: "raw_erp.clientes", fields: [{ name: "id_cliente", type: "STRING" }, { name: "nombre", type: "STRING" }, { name: "pais", type: "STRING" }, { name: "segmento", type: "STRING" }] },
            { name: "raw_rrhh.empleados", fields: [{ name: "id_empleado", type: "STRING" }, { name: "nombre", type: "STRING" }, { name: "departamento", type: "STRING" }, { name: "fecha_alta", type: "DATE" }] }
        ],
        snowflake: [
            { name: "RAW.ERP.VENTAS_LINEA", fields: [{ name: "ID_VENTA", type: "STRING" }, { name: "FECHA", type: "DATE" }, { name: "ID_CLIENTE", type: "STRING" }, { name: "ID_PRODUCTO", type: "STRING" }, { name: "CANTIDAD", type: "INTEGER" }, { name: "IMPORTE", type: "FLOAT" }] },
            { name: "RAW.ERP.PRODUCTOS", fields: [{ name: "ID_PRODUCTO", type: "STRING" }, { name: "DESCRIPCION", type: "STRING" }, { name: "FAMILIA", type: "STRING" }] },
            { name: "RAW.FINANZAS.PRESUPUESTO", fields: [{ name: "EJERCICIO", type: "INTEGER" }, { name: "CENTRO_COSTE", type: "STRING" }, { name: "IMPORTE_PPTO", type: "FLOAT" }] }
        ]
    },

    async render(container, project) {
        this.container = container;
        this.project = project;

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <h3>Interfaces</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET} <span class="mock-badge">MOCKUP</span></p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewLoad">+ Nueva interfaz</button>
            </div>
            <div class="loads-toolbar">
                <label for="loadsCuboFilter">Filtrar por cubo</label>
                <select id="loadsCuboFilter">
                    <option value="">Todos los cubos</option>
                </select>
            </div>
            <div id="loadsListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewLoad").addEventListener("click", () => this.openForm());
        document.getElementById("loadsCuboFilter").addEventListener("change", (e) => {
            this.filterCuboId = e.target.value;
            this.renderList();
        });

        await this.loadCubes();
        this.loadMockList();
        this.renderList();
    },

    async loadCubes() {
        try {
            const sql = `SELECT CUBO_ID, CUBOS, DESCRIPCION, CAMPOS_JSON
                         FROM ${Provider.qualifyControl("CUBOS")}
                         WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                         ORDER BY CUBOS`;
            this.cubes = await Provider.runQuery(sql);
        } catch (err) {
            this.cubes = [];
            UI.toast("Error al cargar los cubos: " + err.message, "error");
        }

        const filterSelect = document.getElementById("loadsCuboFilter");
        if (filterSelect) {
            filterSelect.innerHTML = `<option value="">Todos los cubos</option>` +
                this.cubes.map(c => `<option value="${c.CUBO_ID}">${UI.escapeHtml(c.CUBOS)}</option>`).join("");
            filterSelect.value = this.filterCuboId;
        }
    },

    /** Campos del cubo con etiqueta de grupo corta (D/M) y tipo abreviado */
    cuboFields(cuboId) {
        const cubo = this.cubes.find(c => c.CUBO_ID === cuboId);
        if (!cubo) return [];
        try {
            const spec = JSON.parse(cubo.CAMPOS_JSON || "{}");
            const dims = (spec.dimensions || []).map(d => ({ id: d.colId, name: d.name, type: d.type, group: "D", groupLabel: "Dimensión" }));
            const meas = (spec.measures || []).map(m => ({ id: Provider.toIdentifier(m.name), name: m.name, type: m.type, group: "M", groupLabel: "Medida" }));
            return [...dims, ...meas];
        } catch (e) {
            return [];
        }
    },

    // ------------------------------------------------------------
    // Persistencia mockup (localStorage por proyecto)
    // ------------------------------------------------------------
    storageKey() {
        return `draco_mock_loads_${this.project.PROYECTO_ID}`;
    },

    loadMockList() {
        try {
            this.list = JSON.parse(localStorage.getItem(this.storageKey()) || "[]");
        } catch (e) {
            this.list = [];
        }
    },

    saveMockList() {
        localStorage.setItem(this.storageKey(), JSON.stringify(this.list));
    },

    // ------------------------------------------------------------
    // Listado agrupado por cubo
    // ------------------------------------------------------------
    renderList() {
        const wrap = document.getElementById("loadsListWrap");
        if (!wrap) return;

        if (!this.cubes.length) {
            wrap.innerHTML = `<div class="module-empty">Este proyecto todavía no tiene cubos. Crea al menos un cubo antes de definir interfaces.</div>`;
            return;
        }

        const cubesToShow = this.filterCuboId
            ? this.cubes.filter(c => c.CUBO_ID === this.filterCuboId)
            : this.cubes;

        const groups = cubesToShow.map(cubo => {
            const loads = this.list.filter(l => l.cuboId === cubo.CUBO_ID);
            return { cubo, loads };
        }).filter(g => this.filterCuboId || g.loads.length);

        if (!groups.length) {
            wrap.innerHTML = `<div class="module-empty">Todavía no hay interfaces. Crea la primera con "+ Nueva interfaz".</div>`;
            return;
        }

        wrap.innerHTML = groups.map(({ cubo, loads }) => `
            <div class="load-cube-group">
                <div class="load-cube-group-header">
                    <span class="admin-menu-icon">▣</span>
                    <strong>${UI.escapeHtml(cubo.CUBOS)}</strong>
                    <span class="col-type">${loads.length} interfaz${loads.length === 1 ? "" : "es"}</span>
                </div>
                ${loads.length ? `
                    <div class="data-list">
                        <table>
                            <thead><tr><th>Interfaz</th><th>Origen</th><th>Detalle</th><th>Campos mapeados</th><th></th></tr></thead>
                            <tbody>
                                ${loads.map(l => `
                                    <tr>
                                        <td><strong>${UI.escapeHtml(l.name)}</strong></td>
                                        <td><span class="table-tag">${l.originType === "tabla" ? "Tabla" : "Fichero"}</span></td>
                                        <td>${UI.escapeHtml(this.originSummary(l))}</td>
                                        <td>${this.mappedCount(l)}</td>
                                        <td>
                                            <div class="row-actions">
                                                <button data-edit="${l.id}" title="Editar">✎</button>
                                                <button data-del="${l.id}" class="danger" title="Eliminar">🗑</button>
                                            </div>
                                        </td>
                                    </tr>`).join("")}
                            </tbody>
                        </table>
                    </div>` : `<div class="module-empty module-empty--inline">Sin cargas para este cubo todavía.</div>`}
            </div>`).join("");

        wrap.querySelectorAll("[data-edit]").forEach(btn =>
            btn.addEventListener("click", () => this.openForm(btn.dataset.edit)));
        wrap.querySelectorAll("[data-del]").forEach(btn =>
            btn.addEventListener("click", () => this.remove(btn.dataset.del)));
    },

    originSummary(l) {
        if (l.originType === "tabla") return l.origin.tableName || "(sin tabla seleccionada)";
        const parts = [l.origin.fileType || "?"];
        if (l.origin.fieldSeparator) parts.push(`sep. "${l.origin.fieldSeparator}"`);
        return parts.join(" · ");
    },

    mappedCount(l) {
        const total = this.cuboFields(l.cuboId).length;
        const mapped = Object.values(l.outputMappings || {}).filter(m => m && m.type).length;
        return l.mappingMode === "code" ? "por código" : `${mapped}/${total}`;
    },

    async remove(id) {
        const load = this.list.find(l => l.id === id);
        if (!load) return;
        const ok = await UI.confirm("Eliminar interfaz", `Se eliminará la interfaz <strong>${UI.escapeHtml(load.name)}</strong> (mockup, no afecta a ninguna tabla física).`);
        if (!ok) return;
        this.list = this.list.filter(l => l.id !== id);
        this.saveMockList();
        this.renderList();
        UI.toast(`Interfaz "${load.name}" eliminada.`, "success");
    },

    // ------------------------------------------------------------
    // Datos "en blanco" de una interfaz nueva
    // ------------------------------------------------------------
    blankLoad() {
        return {
            id: Provider.newId(),
            name: "",
            cuboId: this.filterCuboId || (this.cubes[0] ? this.cubes[0].CUBO_ID : ""),
            originType: "tabla",
            origin: {
                connector: Provider.key(),
                tableName: "",
                fileType: "csv",
                hasHeader: true,
                headerLines: 1,
                fieldSeparator: ",",
                decimalSeparator: ".",
                encoding: "UTF-8",
                fields: []
            },
            inputTransformCode: "",
            outputTransformCode: "",
            mappingMode: "visual",
            mappingCode: "",
            outputMappings: {}
        };
    },

    // ------------------------------------------------------------
    // Paso 1: modal pequeño — nombre, origen, cubo destino
    // ------------------------------------------------------------
    openBasicsModal(initial, isNew) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("loadBasicsModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "loadBasicsModal";
                document.body.appendChild(overlay);
            }

            overlay.innerHTML = `
                <div class="modal-box">
                    <div class="modal-header">
                        <h3>${isNew ? "Nueva interfaz" : "Datos básicos de la interfaz"}</h3>
                        <button class="modal-close" id="loadBasicsClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Nombre de la interfaz</label>
                            <input type="text" id="basicsName" placeholder="Ej. Interfaz ventas línea" value="${UI.escapeHtml(initial.name || "")}">
                        </div>
                        <div class="form-group">
                            <label>Origen</label>
                            <div class="segmented" id="basicsOriginType">
                                <button type="button" class="segmented-btn ${initial.originType !== "fichero" ? "active" : ""}" data-origin="tabla">Tabla</button>
                                <button type="button" class="segmented-btn ${initial.originType === "fichero" ? "active" : ""}" data-origin="fichero">Fichero</button>
                            </div>
                        </div>
                        <div class="form-group">
                            <label>Cubo destino</label>
                            <select id="basicsCubo">
                                ${this.cubes.map(c => `<option value="${c.CUBO_ID}" ${c.CUBO_ID === initial.cuboId ? "selected" : ""}>${UI.escapeHtml(c.CUBOS)}</option>`).join("")}
                            </select>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="basicsCancel">Cancelar</button>
                        <button class="btn btn-primary" id="basicsNext">Continuar</button>
                    </div>
                </div>`;

            let originType = initial.originType === "fichero" ? "fichero" : "tabla";
            overlay.querySelectorAll("#basicsOriginType [data-origin]").forEach(btn => {
                btn.addEventListener("click", () => {
                    originType = btn.dataset.origin;
                    overlay.querySelectorAll("#basicsOriginType [data-origin]").forEach(b => b.classList.toggle("active", b === btn));
                });
            });

            overlay.classList.add("visible");
            const nameInput = overlay.querySelector("#basicsName");
            setTimeout(() => { nameInput.focus(); nameInput.select(); }, 50);

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#loadBasicsClose").onclick = () => cleanup(null);
            overlay.querySelector("#basicsCancel").onclick = () => cleanup(null);
            overlay.querySelector("#basicsNext").onclick = () => {
                const name = nameInput.value.trim();
                const cuboId = overlay.querySelector("#basicsCubo").value;
                if (!name) { UI.toast("Indica un nombre para la interfaz.", "error"); return; }
                if (!cuboId) { UI.toast("Selecciona un cubo destino.", "error"); return; }
                cleanup({ name, cuboId, originType });
            };
            nameInput.onkeydown = (e) => { if (e.key === "Enter") overlay.querySelector("#basicsNext").click(); };
        });
    },

    // ------------------------------------------------------------
    // Orquesta los 2 pasos de alta/edición
    // ------------------------------------------------------------
    async openForm(editId = null) {
        if (!this.cubes.length) {
            UI.toast("Crea al menos un cubo antes de definir una interfaz.", "error");
            return;
        }

        const existing = editId ? this.list.find(l => l.id === editId) : null;
        this.editingIsNew = !existing;
        const draft = existing ? JSON.parse(JSON.stringify(existing)) : this.blankLoad();

        const basics = await this.openBasicsModal(draft, this.editingIsNew);
        if (!basics) return;

        const cuboChanged = draft.cuboId !== basics.cuboId;
        const originChanged = draft.originType !== basics.originType;
        Object.assign(draft, basics);
        if (cuboChanged) draft.outputMappings = {};
        if (originChanged) {
            draft.origin.fields = [];
            draft.origin.tableName = "";
            draft.outputMappings = {};
        }

        this.editing = draft;
        this.fileParamsCollapsed = false;
        this.openMainModal();
    },

    /** Reabre el paso 1 desde dentro del modal grande, sin perder el resto del mapeo */
    async editBasicsInline() {
        const before = { cuboId: this.editing.cuboId, originType: this.editing.originType };
        const basics = await this.openBasicsModal(this.editing, false);
        if (!basics) return;

        const cuboChanged = basics.cuboId !== before.cuboId;
        const originChanged = basics.originType !== before.originType;
        Object.assign(this.editing, basics);

        if (cuboChanged) this.editing.outputMappings = {};
        if (originChanged) {
            this.editing.origin.fields = [];
            this.editing.origin.tableName = "";
            this.editing.outputMappings = {};
        }

        this.updateModalHeader();
        if (cuboChanged || originChanged) {
            this.renderOriginPanel();
            this.renderInputFields();
        }
        this.renderOutputFields();
    },

    // ------------------------------------------------------------
    // Paso 2: modal grande (mapeo)
    // ------------------------------------------------------------
    openMainModal() {
        let overlay = document.getElementById("loadFormModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "loadFormModal";
            document.body.appendChild(overlay);
        }
        this.overlay = overlay;

        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div>
                        <h3 id="loadModalTitle"></h3>
                        <span class="modal-subtitle" id="loadModalSubtitle"></span>
                    </div>
                    <div class="modal-header-right">
                        <button class="btn btn-secondary btn-sm" id="btnEditBasics">✎ Nombre / origen / cubo</button>
                        <button class="modal-close" id="loadFormClose">&times;</button>
                    </div>
                </div>
                <div class="modal-body modal-body-flush">
                    <div id="loadOriginPanel" class="load-origin-panel"></div>

                    <div class="load-fn-toolbar">
                        <button class="btn btn-secondary btn-sm" id="btnFnInput">⚙ Función cambiar datos input</button>
                        <button class="btn btn-secondary btn-sm" id="btnFnOutput">⚙ Función cambiar datos output</button>
                        <button class="btn btn-secondary btn-sm" id="btnFnMappingCode">{ } Mapeo por código</button>
                        <span class="load-fn-toolbar-spacer"></span>
                        <span class="form-hint" id="mappingModeHint"></span>
                    </div>

                    <div class="load-mapping-cols" id="loadMappingCols">
                        <div class="load-mapping-col load-mapping-col--input">
                            <div class="load-mapping-col-header">
                                <span>Input</span>
                                <div class="load-mapping-col-actions" id="inputFieldActions"></div>
                            </div>
                            <div class="load-input-fields" id="loadInputFields"></div>
                        </div>
                        <div class="load-mapping-connector" id="loadConnectorWrap">
                            <svg id="loadConnectorSvg"></svg>
                        </div>
                        <div class="load-mapping-col load-mapping-col--output">
                            <div class="load-mapping-col-header">
                                <span>Output (cubo)</span>
                            </div>
                            <div class="load-output-fields" id="loadOutputFields"></div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="loadFormCancel">Cancelar</button>
                    <button class="btn btn-primary" id="loadFormSave">Guardar interfaz</button>
                </div>
            </div>`;

        document.getElementById("loadFormClose").addEventListener("click", () => this.closeForm());
        document.getElementById("loadFormCancel").addEventListener("click", () => this.closeForm());
        document.getElementById("loadFormSave").addEventListener("click", () => this.save());
        document.getElementById("btnEditBasics").addEventListener("click", () => this.editBasicsInline());

        document.getElementById("btnFnInput").addEventListener("click", async () => {
            const code = await UI.openCodeEditorModal({
                title: "Función: cambiar datos input",
                subtitle: this.editing.name,
                code: this.editing.inputTransformCode
            });
            if (code !== null) {
                this.editing.inputTransformCode = code;
                UI.toast("Código de transformación de input guardado (mockup).", "success");
            }
        });

        document.getElementById("btnFnOutput").addEventListener("click", async () => {
            const code = await UI.openCodeEditorModal({
                title: "Función: cambiar datos output",
                subtitle: this.editing.name,
                code: this.editing.outputTransformCode
            });
            if (code !== null) {
                this.editing.outputTransformCode = code;
                UI.toast("Código de transformación de output guardado (mockup).", "success");
            }
        });

        document.getElementById("btnFnMappingCode").addEventListener("click", () => this.toggleMappingCode());

        overlay.classList.add("visible");
        this._onResize = () => this.scheduleConnectorRedraw();
        window.addEventListener("resize", this._onResize);

        this.updateModalHeader();
        this.renderOriginPanel();
        this.renderInputFields();
        this.renderOutputFields();
        this.updateMappingModeUI();
    },

    updateModalHeader() {
        const cubo = this.cubes.find(c => c.CUBO_ID === this.editing.cuboId);
        document.getElementById("loadModalTitle").textContent = this.editing.name;
        document.getElementById("loadModalSubtitle").innerHTML =
            `Cubo: <strong>${UI.escapeHtml(cubo ? cubo.CUBOS : "—")}</strong> · Origen: <strong>${this.editing.originType === "tabla" ? "Tabla" : "Fichero"}</strong>`;
    },

    closeForm() {
        if (this.overlay) this.overlay.classList.remove("visible");
        if (this._onResize) { window.removeEventListener("resize", this._onResize); this._onResize = null; }
        this.editing = null;
    },

    // ------------------------------------------------------------
    // Panel de origen: Tabla o Fichero (con formato de fichero colapsable)
    // ------------------------------------------------------------
    renderOriginPanel() {
        const panel = document.getElementById("loadOriginPanel");
        const o = this.editing.origin;

        if (this.editing.originType === "tabla") {
            panel.innerHTML = `
                <div class="load-origin-row">
                    <div class="form-group">
                        <label>Conector</label>
                        <select id="originConnector">
                            <option value="bigquery" ${o.connector === "bigquery" ? "selected" : ""}>BigQuery</option>
                            <option value="snowflake" ${o.connector === "snowflake" ? "selected" : ""}>Snowflake</option>
                        </select>
                    </div>
                    <div class="form-group" style="flex:1;">
                        <label>Tabla de origen</label>
                        <div class="load-table-picker">
                            <input type="text" id="originTableName" readonly placeholder="Ninguna tabla seleccionada" value="${UI.escapeHtml(o.tableName)}">
                            <button class="btn btn-secondary btn-sm" id="btnPickTable">Buscar tabla…</button>
                        </div>
                    </div>
                </div>`;

            document.getElementById("originConnector").addEventListener("change", (e) => { o.connector = e.target.value; });
            document.getElementById("btnPickTable").addEventListener("click", async () => {
                const tables = this.MOCK_TABLES[o.connector] || [];
                const picked = await UI.openTablePickerModal({ connector: o.connector === "snowflake" ? "Snowflake" : "BigQuery", tables });
                if (!picked) return;
                o.tableName = picked.name;
                o.fields = picked.fields.map(f => ({ ...f, custom: false, filter: null }));
                document.getElementById("originTableName").value = o.tableName;
                this.renderInputFields();
                UI.toast(`Tabla "${picked.name}" seleccionada (mockup).`, "success");
            });
        } else {
            const collapsed = this.fileParamsCollapsed;
            panel.innerHTML = `
                <div class="load-file-params-header">
                    <span>Formato de fichero</span>
                    <button type="button" class="link-btn" id="btnToggleFileParams">${collapsed ? "Mostrar ▾" : "Ocultar ▴"}</button>
                </div>
                <div class="load-origin-row load-origin-row--file ${collapsed ? "is-collapsed" : ""}" id="fileParamsGrid">
                    <div class="form-group">
                        <label>Tipo</label>
                        <select id="fileType">
                            <option value="csv" ${o.fileType === "csv" ? "selected" : ""}>CSV / texto</option>
                            <option value="xlsx" ${o.fileType === "xlsx" ? "selected" : ""}>Excel (.xlsx)</option>
                            <option value="json" ${o.fileType === "json" ? "selected" : ""}>JSON</option>
                            <option value="fixed" ${o.fileType === "fixed" ? "selected" : ""}>Ancho fijo</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Líneas cabecera</label>
                        <input type="number" id="fileHeaderLines" min="0" value="${o.headerLines}">
                    </div>
                    <div class="form-group form-group--check">
                        <label><input type="checkbox" id="fileHasHeader" ${o.hasHeader ? "checked" : ""}> Con cabecera</label>
                    </div>
                    <div class="form-group">
                        <label>Sep. campo</label>
                        <select id="fileFieldSep">
                            <option value="," ${o.fieldSeparator === "," ? "selected" : ""}>Coma (,)</option>
                            <option value=";" ${o.fieldSeparator === ";" ? "selected" : ""}>Punto y coma (;)</option>
                            <option value="\t" ${o.fieldSeparator === "\t" ? "selected" : ""}>Tabulador</option>
                            <option value="|" ${o.fieldSeparator === "|" ? "selected" : ""}>Barra (|)</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Sep. decimal</label>
                        <select id="fileDecSep">
                            <option value="." ${o.decimalSeparator === "." ? "selected" : ""}>Punto (.)</option>
                            <option value="," ${o.decimalSeparator === "," ? "selected" : ""}>Coma (,)</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Codificación</label>
                        <select id="fileEncoding">
                            <option value="UTF-8" ${o.encoding === "UTF-8" ? "selected" : ""}>UTF-8</option>
                            <option value="ISO-8859-1" ${o.encoding === "ISO-8859-1" ? "selected" : ""}>ISO-8859-1</option>
                            <option value="Windows-1252" ${o.encoding === "Windows-1252" ? "selected" : ""}>Windows-1252</option>
                        </select>
                    </div>
                </div>
                <div class="load-origin-row load-origin-row--sample">
                    <button class="btn btn-secondary btn-sm" id="btnSampleFile">📄 Cargar fichero de ejemplo</button>
                    <input type="file" id="sampleFileInput" accept=".csv,.txt,.xlsx,.xls,.json" style="display:none;">
                    <span class="form-hint">Solo se usa para leer la cabecera y proponer campos; no se guarda el fichero.</span>
                </div>`;

            document.getElementById("btnToggleFileParams").addEventListener("click", () => {
                this.fileParamsCollapsed = !this.fileParamsCollapsed;
                this.renderOriginPanel();
            });

            if (!collapsed) {
                const bind = (id, prop, isCheckbox) => document.getElementById(id).addEventListener("change", (e) => {
                    o[prop] = isCheckbox ? e.target.checked : e.target.value;
                });
                bind("fileType", "fileType");
                bind("fileHeaderLines", "headerLines");
                bind("fileHasHeader", "hasHeader", true);
                bind("fileFieldSep", "fieldSeparator");
                bind("fileDecSep", "decimalSeparator");
                bind("fileEncoding", "encoding");
            }

            document.getElementById("btnSampleFile").addEventListener("click", () => document.getElementById("sampleFileInput").click());
            document.getElementById("sampleFileInput").addEventListener("change", (e) => this.handleSampleFile(e));
        }
    },

    async handleSampleFile(e) {
        const file = e.target.files[0];
        e.target.value = "";
        if (!file) return;
        const o = this.editing.origin;

        try {
            let headerNames = [];
            if (/\.(xlsx|xls)$/i.test(file.name)) {
                const buf = await file.arrayBuffer();
                const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
                const headerRowIdx = o.hasHeader ? (o.headerLines || 1) - 1 : -1;
                headerNames = headerRowIdx >= 0 && aoa[headerRowIdx] ? aoa[headerRowIdx].map(String) : (aoa[0] || []).map((_, i) => `campo_${i + 1}`);
            } else if (/\.json$/i.test(file.name)) {
                const text = await file.text();
                const parsed = JSON.parse(text);
                const first = Array.isArray(parsed) ? parsed[0] : parsed;
                headerNames = first ? Object.keys(first) : [];
            } else {
                const text = await file.text();
                const lines = text.replace(/\r/g, "").split("\n").filter(l => l.length);
                const sepLine = o.hasHeader ? lines[(o.headerLines || 1) - 1] : lines[0];
                const sep = o.fieldSeparator === "\t" ? "\t" : o.fieldSeparator;
                headerNames = sepLine ? sepLine.split(sep).map(s => s.trim().replace(/^"|"$/g, "")) : [];
                if (!o.hasHeader) headerNames = headerNames.map((_, i) => `campo_${i + 1}`);
            }

            if (!headerNames.length) {
                UI.toast("No se han podido detectar campos en el fichero de ejemplo.", "error");
                return;
            }

            o.fields = headerNames.filter(Boolean).map(n => ({ name: n, type: "STRING", custom: false, filter: null }));
            this.renderInputFields();
            UI.toast(`${o.fields.length} campo(s) detectado(s) del fichero de ejemplo.`, "success");
        } catch (err) {
            UI.toast("Error al leer el fichero de ejemplo: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Columna de campos de entrada (input) — compacta
    // ------------------------------------------------------------
    renderInputFields() {
        const wrap = document.getElementById("loadInputFields");
        const actions = document.getElementById("inputFieldActions");
        const o = this.editing.origin;
        const isTable = this.editing.originType === "tabla";

        actions.innerHTML = `<button class="btn btn-secondary btn-sm" id="btnAddInputField">+ Añadir${isTable ? " (visual)" : ""}</button>`;
        document.getElementById("btnAddInputField").addEventListener("click", () => {
            o.fields.push({ name: `campo_${o.fields.length + 1}`, type: "STRING", custom: true, filter: null });
            this.renderInputFields();
        });

        if (!o.fields.length) {
            wrap.innerHTML = `<div class="hierarchy-pool-empty">${isTable ? "Selecciona una tabla o añade campos manualmente." : "Añade campos manualmente o carga un fichero de ejemplo."}</div>`;
            this.scheduleConnectorRedraw();
            return;
        }

        wrap.innerHTML = o.fields.map((f, idx) => {
            const hasFilter = !!(f.filter && f.filter.type);
            const filterTitle = hasFilter
                ? `Filtro: ${f.filter.type === "constante" ? `constante (${f.filter.value || "—"})` : "variable"}`
                : "Sin filtro — clic para añadir uno";
            return `
            <div class="load-input-field-row" draggable="true" data-idx="${idx}" data-name="${UI.escapeHtml(f.name)}">
                <span class="load-drag-handle" title="Arrastra a un campo de salida">⠿</span>
                <input type="text" class="load-field-name" data-idx="${idx}" value="${UI.escapeHtml(f.name)}">
                <select class="load-field-type" data-idx="${idx}">
                    ${UI.FIELD_TYPES.map(t => `<option value="${t}" ${t === f.type ? "selected" : ""}>${this.TYPE_ABBR[t] || t}</option>`).join("")}
                </select>
                <button type="button" class="load-filter-btn ${hasFilter ? "has-filter" : ""}" data-filter-idx="${idx}" title="${UI.escapeHtml(filterTitle)}">
                    <svg viewBox="0 0 16 16" width="11" height="11" aria-hidden="true"><path d="M1.5 2.5h13l-4.8 5.6v4.4l-3.4 1.8V8.1z"/></svg>
                </button>
                ${f.custom ? '<span class="load-manual-dot" title="Campo añadido manualmente">●</span>' : ""}
                <button type="button" class="field-remove" data-remove="${idx}" title="Eliminar">✕</button>
            </div>`;
        }).join("");

        wrap.querySelectorAll(".load-field-name").forEach(inp => inp.addEventListener("input", (e) => {
            const idx = parseInt(e.target.dataset.idx, 10);
            const oldName = o.fields[idx].name;
            const newName = e.target.value;
            o.fields[idx].name = newName;
            e.target.closest(".load-input-field-row").dataset.name = newName;
            Object.values(this.editing.outputMappings).forEach(m => {
                if (m && m.type === "campo" && m.value === oldName) m.value = newName;
            });
            this.renderOutputFields();
        }));
        wrap.querySelectorAll(".load-field-type").forEach(sel => sel.addEventListener("change", (e) => {
            o.fields[parseInt(e.target.dataset.idx, 10)].type = e.target.value;
        }));
        wrap.querySelectorAll("[data-remove]").forEach(btn => btn.addEventListener("click", () => {
            const idx = parseInt(btn.dataset.remove, 10);
            const removed = o.fields[idx];
            o.fields.splice(idx, 1);
            Object.entries(this.editing.outputMappings).forEach(([outId, m]) => {
                if (m && m.type === "campo" && m.value === removed.name) delete this.editing.outputMappings[outId];
            });
            this.renderInputFields();
            this.renderOutputFields();
        }));
        wrap.querySelectorAll("[data-filter-idx]").forEach(btn => btn.addEventListener("click", async () => {
            const idx = parseInt(btn.dataset.filterIdx, 10);
            const field = o.fields[idx];
            const result = await UI.openFilterFieldModal({
                title: "Filtro de campo",
                fieldName: field.name,
                current: field.filter
            });
            if (result === null) return;
            field.filter = result === "remove" ? null : result;
            this.renderInputFields();
        }));

        wrap.querySelectorAll(".load-input-field-row").forEach(row => {
            row.addEventListener("dragstart", () => {
                this.dragFieldName = row.dataset.name;
                row.classList.add("dragging");
                document.getElementById("loadMappingCols").classList.add("is-drag-active");
            });
            row.addEventListener("dragend", () => {
                row.classList.remove("dragging");
                const cols = document.getElementById("loadMappingCols");
                if (cols) cols.classList.remove("is-drag-active");
            });
        });

        this.scheduleConnectorRedraw();
    },

    // ------------------------------------------------------------
    // Columna de campos de salida (output) — compacta, drag&drop siempre activo
    // ------------------------------------------------------------
    renderOutputFields() {
        const wrap = document.getElementById("loadOutputFields");
        const fields = this.cuboFields(this.editing.cuboId);

        if (!fields.length) {
            wrap.innerHTML = `<div class="hierarchy-pool-empty">Este cubo no tiene dimensiones ni medidas definidas.</div>`;
            this.scheduleConnectorRedraw();
            return;
        }

        wrap.innerHTML = fields.map(f => {
            const m = this.editing.outputMappings[f.id] || {};
            return `
            <div class="load-output-field-row" data-out="${f.id}">
                <span class="load-output-group-tag load-output-group-tag--${f.group}" title="${UI.escapeHtml(f.groupLabel)}">${f.group}</span>
                <div class="load-output-field-name" title="${UI.escapeHtml(f.name)}">
                    ${UI.escapeHtml(f.name)}
                    <span class="load-output-type-abbr">${this.TYPE_ABBR[f.type] || f.type}</span>
                </div>
                <select class="load-map-type" data-out="${f.id}">
                    <option value="" ${!m.type ? "selected" : ""}>Sin mapear</option>
                    <option value="campo" ${m.type === "campo" ? "selected" : ""}>Campo origen</option>
                    <option value="constante" ${m.type === "constante" ? "selected" : ""}>Constante</option>
                    <option value="variable" ${m.type === "variable" ? "selected" : ""}>Variable</option>
                    <option value="formula" ${m.type === "formula" ? "selected" : ""}>Fórmula</option>
                    <option value="funcion" ${m.type === "funcion" ? "selected" : ""}>Función</option>
                </select>
                <div class="load-map-target" data-out-drop="${f.id}">
                    ${this.mappingSummaryHtml(m)}
                </div>
            </div>`;
        }).join("");

        wrap.querySelectorAll(".load-map-type").forEach(sel => {
            sel.addEventListener("change", (e) => this.onMapTypeChange(e.target.dataset.out, e.target.value));
        });

        // Drag&drop siempre activo en cualquier fila, sea cual sea el tipo de mapeo actual
        wrap.querySelectorAll("[data-out-drop]").forEach(zone => {
            zone.addEventListener("dragover", (e) => e.preventDefault());
            zone.addEventListener("dragenter", () => zone.classList.add("is-drop-hover"));
            zone.addEventListener("dragleave", () => zone.classList.remove("is-drop-hover"));
            zone.addEventListener("drop", (e) => {
                e.preventDefault();
                zone.classList.remove("is-drop-hover");
                if (!this.dragFieldName) return;
                const outId = zone.dataset.outDrop;
                this.editing.outputMappings[outId] = { type: "campo", value: this.dragFieldName };
                this.renderOutputFields();
            });
        });

        this.updateMappingModeUI();
        this.scheduleConnectorRedraw();
    },

    mappingSummaryHtml(m) {
        if (!m || !m.type) return `<span class="load-map-empty">Arrastra un campo o elige un tipo</span>`;
        const labels = {
            campo: `<strong>${UI.escapeHtml(m.value || "—")}</strong>`,
            constante: `Constante: <strong>${UI.escapeHtml(m.value || "—")}</strong>`,
            variable: `Variable: <strong>${UI.escapeHtml(m.value || "(pendiente)")}</strong>`,
            formula: `<code>${UI.escapeHtml((m.value || "").slice(0, 50) || "—")}</code>`,
            funcion: `Función Python asignada`
        };
        return labels[m.type] || "";
    },

    async onMapTypeChange(outId, type) {
        if (!type) {
            delete this.editing.outputMappings[outId];
            this.renderOutputFields();
            return;
        }

        const current = this.editing.outputMappings[outId] || {};

        if (type === "campo") {
            this.editing.outputMappings[outId] = { type: "campo", value: current.value || "" };
            this.renderOutputFields();
            return;
        }

        if (type === "constante") {
            const val = await UI.openTextPromptModal({ title: "Mapeo por constante", label: "Valor constante", value: current.type === "constante" ? current.value : "" });
            if (val === null) { this.renderOutputFields(); return; }
            this.editing.outputMappings[outId] = { type: "constante", value: val };
            this.renderOutputFields();
            return;
        }

        if (type === "variable") {
            UI.toast("La selección de variables se desarrollará más adelante. Por ahora se guarda como pendiente.", "info");
            this.editing.outputMappings[outId] = { type: "variable", value: current.type === "variable" ? current.value : "" };
            this.renderOutputFields();
            return;
        }

        if (type === "formula") {
            const formula = await UI.openFormulaEditorModal({
                title: "Editor de fórmula",
                inputFields: this.editing.origin.fields,
                value: current.type === "formula" ? current.value : ""
            });
            if (formula === null) { this.renderOutputFields(); return; }
            this.editing.outputMappings[outId] = { type: "formula", value: formula };
            this.renderOutputFields();
            return;
        }

        if (type === "funcion") {
            const code = await UI.openCodeEditorModal({
                title: "Función de campo (Python)",
                subtitle: `Campo de salida: ${outId}`,
                code: current.type === "funcion" ? current.value : ""
            });
            if (code === null) { this.renderOutputFields(); return; }
            this.editing.outputMappings[outId] = { type: "funcion", value: code };
            this.renderOutputFields();
        }
    },

    // ------------------------------------------------------------
    // Mapeo por código (sustituye al mapeo campo a campo)
    // ------------------------------------------------------------
    async toggleMappingCode() {
        if (this.editing.mappingMode === "code") {
            const ok = await UI.confirm("Volver a mapeo visual", "Se mantendrá el código guardado, pero el mapeo campo a campo vuelve a estar activo. ¿Continuar?");
            if (!ok) return;
            this.editing.mappingMode = "visual";
            this.updateMappingModeUI();
            return;
        }

        const code = await UI.openCodeEditorModal({
            title: "Mapeo por código",
            subtitle: "Todo el mapeo de esta interfaz se resolverá por código Python, en lugar de campo a campo.",
            code: this.editing.mappingCode || `def mapear(df_input):\n    # Devuelve un DataFrame con las columnas del cubo de destino\n    df_output = df_input.copy()\n    return df_output\n`
        });
        if (code === null) return;
        this.editing.mappingCode = code;
        this.editing.mappingMode = "code";
        this.updateMappingModeUI();
        UI.toast("Mapeo por código guardado. El mapeo campo a campo queda deshabilitado.", "success");
    },

    updateMappingModeUI() {
        const hint = document.getElementById("mappingModeHint");
        const cols = document.getElementById("loadMappingCols");
        if (!hint || !cols) return;

        if (this.editing.mappingMode === "code") {
            hint.textContent = "Mapeo por código activo: todo el mapeo se hace por código.";
            cols.classList.add("is-code-mode");
        } else {
            hint.textContent = "";
            cols.classList.remove("is-code-mode");
        }
    },

    // ------------------------------------------------------------
    // Conector visual entre campos de input y output mapeados por campo
    // ------------------------------------------------------------
    scheduleConnectorRedraw() {
        if (this._connectorRaf) cancelAnimationFrame(this._connectorRaf);
        this._connectorRaf = requestAnimationFrame(() => this.drawConnectors());
    },

    drawConnectors() {
        const svg = document.getElementById("loadConnectorSvg");
        const wrap = document.getElementById("loadConnectorWrap");
        if (!svg || !wrap || !this.editing) return;

        const wrapRect = wrap.getBoundingClientRect();
        svg.setAttribute("width", Math.max(wrapRect.width, 1));
        svg.setAttribute("height", Math.max(wrapRect.height, 1));
        svg.innerHTML = "";

        if (this.editing.mappingMode === "code") return;

        const fields = this.cuboFields(this.editing.cuboId);
        const svgNS = "http://www.w3.org/2000/svg";

        fields.forEach(f => {
            const m = this.editing.outputMappings[f.id];
            if (!m || m.type !== "campo" || !m.value) return;

            const inputRow = document.querySelector(`.load-input-field-row[data-name="${CSS.escape(m.value)}"]`);
            const outputRow = document.querySelector(`.load-output-field-row[data-out="${CSS.escape(f.id)}"]`);
            if (!inputRow || !outputRow) return;

            const ir = inputRow.getBoundingClientRect();
            const or = outputRow.getBoundingClientRect();
            const y1 = ir.top + ir.height / 2 - wrapRect.top;
            const y2 = or.top + or.height / 2 - wrapRect.top;
            const x1 = 0, x2 = wrapRect.width;
            const midX = wrapRect.width / 2;

            const path = document.createElementNS(svgNS, "path");
            path.setAttribute("d", `M ${x1} ${y1} C ${midX} ${y1}, ${midX} ${y2}, ${x2} ${y2}`);
            path.setAttribute("class", "load-connector-line");
            svg.appendChild(path);

            [[x1, y1], [x2, y2]].forEach(([cx, cy]) => {
                const dot = document.createElementNS(svgNS, "circle");
                dot.setAttribute("cx", cx);
                dot.setAttribute("cy", cy);
                dot.setAttribute("r", 3);
                dot.setAttribute("class", "load-connector-dot");
                svg.appendChild(dot);
            });
        });
    },

    // ------------------------------------------------------------
    // Guardado (mockup → localStorage)
    // ------------------------------------------------------------
    save() {
        const idx = this.list.findIndex(l => l.id === this.editing.id);
        if (idx >= 0) this.list[idx] = this.editing;
        else this.list.push(this.editing);

        this.saveMockList();
        const name = this.editing.name;
        this.closeForm();
        this.renderList();
        UI.toast(`Interfaz "${name}" guardada (mockup).`, "success");
    }
};
