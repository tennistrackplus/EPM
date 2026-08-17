/**
 * ============================================================
 * DRACO PLANNING — CARGAS DE DATOS (MOCKUP)
 * ============================================================
 * Listado de cargas de datos agrupadas por cubo, con alta/edición
 * en un modal casi a pantalla completa (mismo patrón que "Valores"
 * de dimensiones). Todo lo relativo al origen (tabla/fichero), el
 * mapeo de campos y las funciones Python es, por ahora, MOCKUP:
 * se guarda en localStorage, no en las tablas de control reales.
 * Cuando el mockup guste, se decide cómo persistirlo en BBDD.
 */
const Loads = {
    list: [],
    cubes: [],
    filterCuboId: "",
    editing: null,   // copia de trabajo de la carga que se está editando en el modal
    dragFieldName: null,

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
                    <h3>Cargas de datos</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET} <span class="mock-badge">MOCKUP</span></p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewLoad">+ Nueva carga de datos</button>
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

    cuboFields(cuboId) {
        const cubo = this.cubes.find(c => c.CUBO_ID === cuboId);
        if (!cubo) return [];
        try {
            const spec = JSON.parse(cubo.CAMPOS_JSON || "{}");
            const dims = (spec.dimensions || []).map(d => ({ id: d.colId, name: d.name, type: d.type, group: "Dimensión" }));
            const meas = (spec.measures || []).map(m => ({ id: Provider.toIdentifier(m.name), name: m.name, type: m.type, group: "Medida" }));
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
            wrap.innerHTML = `<div class="module-empty">Este proyecto todavía no tiene cubos. Crea al menos un cubo antes de definir cargas de datos.</div>`;
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
            wrap.innerHTML = `<div class="module-empty">Todavía no hay cargas de datos. Crea la primera con "+ Nueva carga de datos".</div>`;
            return;
        }

        wrap.innerHTML = groups.map(({ cubo, loads }) => `
            <div class="load-cube-group">
                <div class="load-cube-group-header">
                    <span class="admin-menu-icon">▣</span>
                    <strong>${UI.escapeHtml(cubo.CUBOS)}</strong>
                    <span class="col-type">${loads.length} carga(s)</span>
                </div>
                ${loads.length ? `
                    <div class="data-list">
                        <table>
                            <thead><tr><th>Carga de datos</th><th>Origen</th><th>Detalle</th><th>Campos mapeados</th><th></th></tr></thead>
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
        const ok = await UI.confirm("Eliminar carga de datos", `Se eliminará la carga <strong>${UI.escapeHtml(load.name)}</strong> (mockup, no afecta a ninguna tabla física).`);
        if (!ok) return;
        this.list = this.list.filter(l => l.id !== id);
        this.saveMockList();
        this.renderList();
        UI.toast(`Carga "${load.name}" eliminada.`, "success");
    },

    // ------------------------------------------------------------
    // Alta / edición — modal casi a pantalla completa
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

    openForm(editId = null) {
        const existing = editId ? this.list.find(l => l.id === editId) : null;
        this.editing = existing ? JSON.parse(JSON.stringify(existing)) : this.blankLoad();
        this.editingIsNew = !existing;

        if (!this.cubes.length) {
            UI.toast("Crea al menos un cubo antes de definir una carga de datos.", "error");
            return;
        }

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
                        <h3>${existing ? "Editar carga de datos" : "Nueva carga de datos"}</h3>
                        <span class="modal-subtitle">Mockup — el mapeo y las funciones todavía no se ejecutan ni se guardan en base de datos</span>
                    </div>
                    <button class="modal-close" id="loadFormClose">&times;</button>
                </div>
                <div class="modal-body modal-body-flush">
                    <div class="load-form-top">
                        <div class="form-group">
                            <label>Nombre de la carga de datos</label>
                            <input type="text" id="loadName" placeholder="Ej. Carga ventas línea" value="${UI.escapeHtml(this.editing.name)}">
                        </div>
                        <div class="form-group">
                            <label>Cubo destino</label>
                            <select id="loadCubo">
                                ${this.cubes.map(c => `<option value="${c.CUBO_ID}" ${c.CUBO_ID === this.editing.cuboId ? "selected" : ""}>${UI.escapeHtml(c.CUBOS)}</option>`).join("")}
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Origen</label>
                            <div class="segmented" id="loadOriginType">
                                <button type="button" class="segmented-btn ${this.editing.originType === "tabla" ? "active" : ""}" data-origin="tabla">Tabla</button>
                                <button type="button" class="segmented-btn ${this.editing.originType === "fichero" ? "active" : ""}" data-origin="fichero">Fichero</button>
                            </div>
                        </div>
                    </div>

                    <div id="loadOriginPanel" class="load-origin-panel"></div>

                    <div class="load-fn-toolbar">
                        <button class="btn btn-secondary btn-sm" id="btnFnInput">⚙ Función cambiar datos input</button>
                        <button class="btn btn-secondary btn-sm" id="btnFnOutput">⚙ Función cambiar datos output</button>
                        <button class="btn btn-secondary btn-sm" id="btnFnMappingCode">{ } Mapeo por código</button>
                        <span class="load-fn-toolbar-spacer"></span>
                        <span class="form-hint" id="mappingModeHint"></span>
                    </div>

                    <div class="load-mapping-cols" id="loadMappingCols">
                        <div class="load-mapping-col">
                            <div class="load-mapping-col-header">
                                <span>Campos de entrada (input)</span>
                                <div class="load-mapping-col-actions" id="inputFieldActions"></div>
                            </div>
                            <div class="load-input-fields" id="loadInputFields"></div>
                        </div>
                        <div class="load-mapping-col">
                            <div class="load-mapping-col-header">
                                <span>Campos de salida (output — cubo)</span>
                            </div>
                            <div class="load-output-fields" id="loadOutputFields"></div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="loadFormCancel">Cancelar</button>
                    <button class="btn btn-primary" id="loadFormSave">Guardar carga de datos</button>
                </div>
            </div>`;

        document.getElementById("loadFormClose").addEventListener("click", () => this.closeForm());
        document.getElementById("loadFormCancel").addEventListener("click", () => this.closeForm());
        document.getElementById("loadFormSave").addEventListener("click", () => this.save());

        document.getElementById("loadName").addEventListener("input", (e) => { this.editing.name = e.target.value; });
        document.getElementById("loadCubo").addEventListener("change", (e) => {
            this.editing.cuboId = e.target.value;
            this.editing.outputMappings = {};
            this.renderOutputFields();
        });

        overlay.querySelectorAll("#loadOriginType [data-origin]").forEach(btn => {
            btn.addEventListener("click", () => {
                this.editing.originType = btn.dataset.origin;
                overlay.querySelectorAll("#loadOriginType [data-origin]").forEach(b => b.classList.toggle("active", b === btn));
                this.renderOriginPanel();
                this.renderInputFields();
            });
        });

        document.getElementById("btnFnInput").addEventListener("click", async () => {
            const code = await UI.openCodeEditorModal({
                title: "Función: cambiar datos input",
                subtitle: this.editing.name || "Nueva carga",
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
                subtitle: this.editing.name || "Nueva carga",
                code: this.editing.outputTransformCode
            });
            if (code !== null) {
                this.editing.outputTransformCode = code;
                UI.toast("Código de transformación de output guardado (mockup).", "success");
            }
        });

        document.getElementById("btnFnMappingCode").addEventListener("click", () => this.toggleMappingCode());

        overlay.classList.add("visible");
        this.renderOriginPanel();
        this.renderInputFields();
        this.renderOutputFields();
        this.updateMappingModeUI();
        setTimeout(() => document.getElementById("loadName").focus(), 50);
    },

    closeForm() {
        if (this.overlay) this.overlay.classList.remove("visible");
        this.editing = null;
    },

    // ------------------------------------------------------------
    // Panel de origen: Tabla o Fichero
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

            document.getElementById("originConnector").addEventListener("change", (e) => {
                o.connector = e.target.value;
            });
            document.getElementById("btnPickTable").addEventListener("click", async () => {
                const tables = this.MOCK_TABLES[o.connector] || [];
                const picked = await UI.openTablePickerModal({ connector: o.connector === "snowflake" ? "Snowflake" : "BigQuery", tables });
                if (!picked) return;
                o.tableName = picked.name;
                o.fields = picked.fields.map(f => ({ ...f, custom: false }));
                document.getElementById("originTableName").value = o.tableName;
                this.renderInputFields();
                UI.toast(`Tabla "${picked.name}" seleccionada (mockup).`, "success");
            });
        } else {
            panel.innerHTML = `
                <div class="load-origin-row load-origin-row--file">
                    <div class="form-group">
                        <label>Tipo de fichero</label>
                        <select id="fileType">
                            <option value="csv" ${o.fileType === "csv" ? "selected" : ""}>CSV / texto delimitado</option>
                            <option value="xlsx" ${o.fileType === "xlsx" ? "selected" : ""}>Excel (.xlsx)</option>
                            <option value="json" ${o.fileType === "json" ? "selected" : ""}>JSON</option>
                            <option value="fixed" ${o.fileType === "fixed" ? "selected" : ""}>Ancho fijo</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Líneas de cabecera</label>
                        <input type="number" id="fileHeaderLines" min="0" value="${o.headerLines}">
                    </div>
                    <div class="form-group">
                        <label><input type="checkbox" id="fileHasHeader" ${o.hasHeader ? "checked" : ""}> Primera línea = nombres de campo</label>
                    </div>
                    <div class="form-group">
                        <label>Separador de campo</label>
                        <select id="fileFieldSep">
                            <option value="," ${o.fieldSeparator === "," ? "selected" : ""}>Coma (,)</option>
                            <option value=";" ${o.fieldSeparator === ";" ? "selected" : ""}>Punto y coma (;)</option>
                            <option value="\t" ${o.fieldSeparator === "\t" ? "selected" : ""}>Tabulador</option>
                            <option value="|" ${o.fieldSeparator === "|" ? "selected" : ""}>Barra vertical (|)</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Separador decimal</label>
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
                <div class="load-origin-row">
                    <button class="btn btn-secondary btn-sm" id="btnSampleFile">📄 Cargar fichero de ejemplo (para detectar campos)</button>
                    <input type="file" id="sampleFileInput" accept=".csv,.txt,.xlsx,.xls,.json" style="display:none;">
                    <span class="form-hint">Se usa solo para leer la cabecera y proponer campos; no se guarda el fichero.</span>
                </div>`;

            const bind = (id, prop, isCheckbox) => document.getElementById(id).addEventListener("change", (e) => {
                o[prop] = isCheckbox ? e.target.checked : e.target.value;
            });
            bind("fileType", "fileType");
            bind("fileHeaderLines", "headerLines");
            bind("fileHasHeader", "hasHeader", true);
            bind("fileFieldSep", "fieldSeparator");
            bind("fileDecSep", "decimalSeparator");
            bind("fileEncoding", "encoding");

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

            o.fields = headerNames.filter(Boolean).map(n => ({ name: n, type: "STRING", custom: false }));
            this.renderInputFields();
            UI.toast(`${o.fields.length} campo(s) detectado(s) del fichero de ejemplo.`, "success");
        } catch (err) {
            UI.toast("Error al leer el fichero de ejemplo: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Columna de campos de entrada (input)
    // ------------------------------------------------------------
    renderInputFields() {
        const wrap = document.getElementById("loadInputFields");
        const actions = document.getElementById("inputFieldActions");
        const o = this.editing.origin;
        const isTable = this.editing.originType === "tabla";

        actions.innerHTML = `<button class="btn btn-secondary btn-sm" id="btnAddInputField">+ Añadir campo${isTable ? " (visual)" : ""}</button>`;
        document.getElementById("btnAddInputField").addEventListener("click", () => {
            o.fields.push({ name: `campo_${o.fields.length + 1}`, type: "STRING", custom: true });
            this.renderInputFields();
        });

        if (!o.fields.length) {
            wrap.innerHTML = `<div class="hierarchy-pool-empty">${isTable ? "Selecciona una tabla de origen o añade campos manualmente." : "Añade campos manualmente o carga un fichero de ejemplo."}</div>`;
            return;
        }

        wrap.innerHTML = o.fields.map((f, idx) => `
            <div class="load-input-field-row" draggable="true" data-idx="${idx}">
                <span class="load-drag-handle" title="Arrastra a un campo de salida">⠿</span>
                <input type="text" class="load-field-name" data-idx="${idx}" value="${UI.escapeHtml(f.name)}">
                <select class="load-field-type" data-idx="${idx}">
                    ${UI.FIELD_TYPES.map(t => `<option value="${t}" ${t === f.type ? "selected" : ""}>${t}</option>`).join("")}
                </select>
                ${f.custom ? '<span class="table-tag" title="Campo añadido manualmente">manual</span>' : ""}
                <button type="button" class="field-remove" data-remove="${idx}" title="Eliminar">✕</button>
            </div>`).join("");

        wrap.querySelectorAll(".load-field-name").forEach(inp => inp.addEventListener("input", (e) => {
            o.fields[parseInt(e.target.dataset.idx, 10)].name = e.target.value;
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

        wrap.querySelectorAll(".load-input-field-row").forEach(row => {
            row.addEventListener("dragstart", () => {
                this.dragFieldName = o.fields[parseInt(row.dataset.idx, 10)].name;
                row.classList.add("dragging");
            });
            row.addEventListener("dragend", () => row.classList.remove("dragging"));
        });
    },

    // ------------------------------------------------------------
    // Columna de campos de salida (output — desde el cubo)
    // ------------------------------------------------------------
    renderOutputFields() {
        const wrap = document.getElementById("loadOutputFields");
        const fields = this.cuboFields(this.editing.cuboId);

        if (!fields.length) {
            wrap.innerHTML = `<div class="hierarchy-pool-empty">Este cubo no tiene dimensiones ni medidas definidas.</div>`;
            return;
        }

        wrap.innerHTML = fields.map(f => {
            const m = this.editing.outputMappings[f.id] || {};
            return `
            <div class="load-output-field-row" data-out="${f.id}">
                <div class="load-output-field-name">
                    ${UI.escapeHtml(f.name)}
                    <span class="col-type">${f.group} · ${f.type}</span>
                </div>
                <select class="load-map-type" data-out="${f.id}">
                    <option value="" ${!m.type ? "selected" : ""}>Sin mapear</option>
                    <option value="campo" ${m.type === "campo" ? "selected" : ""}>Campo origen</option>
                    <option value="constante" ${m.type === "constante" ? "selected" : ""}>Constante</option>
                    <option value="variable" ${m.type === "variable" ? "selected" : ""}>Variable</option>
                    <option value="formula" ${m.type === "formula" ? "selected" : ""}>Fórmula</option>
                    <option value="funcion" ${m.type === "funcion" ? "selected" : ""}>Función</option>
                </select>
                <div class="load-map-target ${m.type === "campo" ? "is-dropzone" : ""}" data-drop="${f.id}">
                    ${this.mappingSummaryHtml(m)}
                </div>
            </div>`;
        }).join("");

        wrap.querySelectorAll(".load-map-type").forEach(sel => {
            sel.addEventListener("change", (e) => this.onMapTypeChange(e.target.dataset.out, e.target.value));
        });

        wrap.querySelectorAll("[data-drop]").forEach(zone => {
            zone.addEventListener("dragover", (e) => {
                if (zone.classList.contains("is-dropzone")) e.preventDefault();
            });
            zone.addEventListener("drop", (e) => {
                if (!zone.classList.contains("is-dropzone")) return;
                e.preventDefault();
                if (!this.dragFieldName) return;
                const outId = zone.dataset.drop;
                this.editing.outputMappings[outId] = { type: "campo", value: this.dragFieldName };
                this.renderOutputFields();
            });
        });

        this.updateMappingModeUI();
    },

    mappingSummaryHtml(m) {
        if (!m || !m.type) return `<span class="load-map-empty">Arrastra un campo o elige un tipo de mapeo</span>`;
        const labels = {
            campo: `Campo: <strong>${UI.escapeHtml(m.value || "—")}</strong>`,
            constante: `Constante: <strong>${UI.escapeHtml(m.value || "—")}</strong>`,
            variable: `Variable: <strong>${UI.escapeHtml(m.value || "(pendiente de definir)")}</strong>`,
            formula: `Fórmula: <code>${UI.escapeHtml((m.value || "").slice(0, 60) || "—")}</code>`,
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
            subtitle: "Todo el mapeo de esta carga se resolverá por código Python, en lugar de campo a campo.",
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
    // Guardado (mockup → localStorage)
    // ------------------------------------------------------------
    save() {
        const name = document.getElementById("loadName").value.trim();
        if (!name) {
            UI.toast("Indica un nombre para la carga de datos.", "error");
            return;
        }
        if (!this.editing.cuboId) {
            UI.toast("Selecciona un cubo destino.", "error");
            return;
        }
        if (this.editing.originType === "tabla" && !this.editing.origin.tableName) {
            UI.toast("Selecciona una tabla de origen.", "error");
            return;
        }

        this.editing.name = name;

        const idx = this.list.findIndex(l => l.id === this.editing.id);
        if (idx >= 0) this.list[idx] = this.editing;
        else this.list.push(this.editing);

        this.saveMockList();
        this.closeForm();
        this.renderList();
        UI.toast(`Carga de datos "${name}" guardada (mockup).`, "success");
    }
};
