/**
 * Utilidades de interfaz compartidas por toda la app: toasts,
 * apertura/cierre de modales, confirmaciones y el comportamiento
 * de maximizar/contraer de los bloques principales.
 */
const UI = {
    toast(message, type = "info") {
        const container = document.getElementById("toastContainer");
        if (!container) return;
        const el = document.createElement("div");
        el.className = `toast ${type}`;
        el.textContent = message;
        container.appendChild(el);
        setTimeout(() => {
            el.style.opacity = "0";
            el.style.transition = "opacity 0.25s ease";
            setTimeout(() => el.remove(), 250);
        }, 4200);
    },

    openModal(id) {
        document.getElementById(id).classList.add("visible");
    },

    closeModal(id) {
        document.getElementById(id).classList.remove("visible");
    },

    /** Modal de confirmación genérico. Devuelve una Promise<boolean>. */
    confirm(title, message) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("confirmModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "confirmModal";
                overlay.innerHTML = `
                    <div class="modal-box">
                        <div class="modal-header">
                            <h3 id="confirmModalTitle"></h3>
                            <button class="modal-close" id="confirmModalClose">&times;</button>
                        </div>
                        <div class="modal-body">
                            <p class="confirm-text" id="confirmModalMessage"></p>
                        </div>
                        <div class="modal-footer">
                            <button class="btn btn-secondary" id="confirmModalCancel">Cancelar</button>
                            <button class="btn btn-primary" id="confirmModalOk" style="background-color: var(--color-danger);">Confirmar</button>
                        </div>
                    </div>`;
                document.body.appendChild(overlay);
            }

            document.getElementById("confirmModalTitle").textContent = title;
            document.getElementById("confirmModalMessage").innerHTML = message;
            overlay.classList.add("visible");

            const cleanup = (result) => {
                overlay.classList.remove("visible");
                resolve(result);
            };

            document.getElementById("confirmModalOk").onclick = () => cleanup(true);
            document.getElementById("confirmModalCancel").onclick = () => cleanup(false);
            document.getElementById("confirmModalClose").onclick = () => cleanup(false);
        });
    },

    escapeHtml(str) {
        if (str === null || str === undefined) return "";
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    },

    FIELD_TYPES: ["STRING", "INTEGER", "FLOAT", "NUMERIC", "BOOLEAN", "DATE", "DATETIME", "TIMESTAMP"],

    /** Crea una fila de campo reutilizable para los distintos diseñadores (atributos / medidas) */
    _fieldRow(f = { name: "", type: "STRING", key: false }, { showKey = true } = {}) {
        const row = document.createElement("div");
        row.className = "field-row";
        row.innerHTML = `
            <input type="text" class="f-name" placeholder="ej. codigo_cliente" value="${UI.escapeHtml(f.name)}">
            <select class="f-type">
                ${UI.FIELD_TYPES.map(t => `<option value="${t}" ${t === f.type ? "selected" : ""}>${t}</option>`).join("")}
            </select>
            <span class="field-key">${showKey ? `<input type="checkbox" class="f-key" ${f.key ? "checked" : ""}>` : ""}</span>
            <button type="button" class="field-remove" title="Eliminar">✕</button>`;
        row.querySelector(".field-remove").addEventListener("click", () => row.remove());
        return row;
    },

    _readFieldRows(container) {
        return Array.from(container.querySelectorAll(".field-row")).map(r => ({
            name: r.querySelector(".f-name").value.trim(),
            type: r.querySelector(".f-type").value,
            key: !!r.querySelector(".f-key") && r.querySelector(".f-key").checked
        })).filter(f => f.name);
    },

    /**
     * Modal de Dimensión: nombre + descripción + clave principal (= nombre de
     * la dimensión, automática) + diseñador de atributos. Los atributos
     * también se pueden marcar como clave (para claves compuestas).
     * Devuelve una Promise<{name, description, keyType, attributes}|null>.
     */
    openDimensionFormModal({ title, name = "", description = "", keyType = "STRING", attributes = [], nameEditable = true }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("dimFormModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "dimFormModal";
                overlay.innerHTML = `
                    <div class="modal-box modal-wide">
                        <div class="modal-header">
                            <h3 id="dimFormTitle"></h3>
                            <button class="modal-close" id="dimFormClose">&times;</button>
                        </div>
                        <div class="modal-body">
                            <div class="form-row">
                                <div class="form-group">
                                    <label>Nombre de la dimensión</label>
                                    <input type="text" id="dimFormName" placeholder="Ej. Cuenta">
                                </div>
                            </div>
                            <div class="form-group">
                                <label>Descripción</label>
                                <textarea id="dimFormDesc" rows="2"></textarea>
                            </div>
                            <div class="key-field-box">
                                <div class="key-field-box-info">
                                    <span class="key-field-badge">CLAVE PRINCIPAL</span>
                                    <strong id="dimFormKeyPreview">—</strong>
                                    <span class="key-field-hint">Se crea automáticamente con el nombre de la dimensión.</span>
                                </div>
                                <select id="dimFormKeyType"></select>
                            </div>
                            <div class="form-group">
                                <label>Atributos</label>
                                <div class="fields-builder">
                                    <div class="fields-builder-header">
                                        <span>Nombre del atributo</span><span>Tipo</span><span>Clave</span><span></span>
                                    </div>
                                    <div class="fields-builder-rows" id="dimFormRows"></div>
                                    <div class="fields-builder-footer">
                                        <button class="btn btn-secondary btn-sm" id="dimFormAddField">+ Añadir atributo</button>
                                    </div>
                                </div>
                                <p class="form-hint">Marca "Clave" en un atributo si quieres una clave compuesta (ej. Clase de coste + Sociedad).</p>
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button class="btn btn-secondary" id="dimFormCancel">Cancelar</button>
                            <button class="btn btn-primary" id="dimFormSave">Guardar</button>
                        </div>
                    </div>`;
                document.body.appendChild(overlay);
            }

            const rowsContainer = overlay.querySelector("#dimFormRows");
            const nameInput = overlay.querySelector("#dimFormName");
            const keyPreview = overlay.querySelector("#dimFormKeyPreview");
            const keyTypeSelect = overlay.querySelector("#dimFormKeyType");

            keyTypeSelect.innerHTML = UI.FIELD_TYPES.map(t => `<option value="${t}" ${t === keyType ? "selected" : ""}>${t}</option>`).join("");

            const updateKeyPreview = () => {
                const ident = Provider.toIdentifier(nameInput.value) || "…";
                keyPreview.textContent = ident;
            };
            nameInput.oninput = updateKeyPreview;

            const addRow = (f) => rowsContainer.appendChild(UI._fieldRow(f, { showKey: true }));
            rowsContainer.innerHTML = "";
            attributes.forEach(a => addRow(a));

            overlay.querySelector("#dimFormTitle").textContent = title;
            nameInput.value = name;
            nameInput.disabled = !nameEditable;
            overlay.querySelector("#dimFormDesc").value = description;
            overlay.querySelector("#dimFormAddField").onclick = () => addRow();
            updateKeyPreview();

            overlay.classList.add("visible");
            setTimeout(() => nameInput.focus(), 50);

            const cleanup = (result) => {
                overlay.classList.remove("visible");
                resolve(result);
            };

            overlay.querySelector("#dimFormCancel").onclick = () => cleanup(null);
            overlay.querySelector("#dimFormClose").onclick = () => cleanup(null);
            overlay.querySelector("#dimFormSave").onclick = () => {
                const nameVal = nameInput.value.trim();
                if (!nameVal) {
                    UI.toast("Indica un nombre para la dimensión.", "error");
                    return;
                }
                cleanup({
                    name: nameVal,
                    description: overlay.querySelector("#dimFormDesc").value.trim(),
                    keyType: keyTypeSelect.value,
                    attributes: UI._readFieldRows(rowsContainer)
                });
            };
        });
    },

    /**
     * Modal de Cubo: nombre + descripción + selección de dimensiones ya
     * existentes en el proyecto + diseñador de medidas.
     * Devuelve una Promise<{name, description, dimensionIds, measures}|null>.
     */
    openCubeFormModal({ title, name = "", description = "", dimensionsList = [], selectedDimensionIds = [], measures = [], nameEditable = true }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("cuboFormModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "cuboFormModal";
                overlay.innerHTML = `
                    <div class="modal-box modal-wide">
                        <div class="modal-header">
                            <h3 id="cuboFormTitle"></h3>
                            <button class="modal-close" id="cuboFormClose">&times;</button>
                        </div>
                        <div class="modal-body">
                            <div class="form-row">
                                <div class="form-group">
                                    <label>Nombre del cubo</label>
                                    <input type="text" id="cuboFormName" placeholder="Ej. Ventas">
                                </div>
                            </div>
                            <div class="form-group">
                                <label>Descripción</label>
                                <textarea id="cuboFormDesc" rows="2"></textarea>
                            </div>
                            <div class="form-group">
                                <label>Dimensiones</label>
                                <div class="fields-builder">
                                    <div class="fields-builder-header" style="grid-template-columns: 1fr;">
                                        <span>Dimensiones añadidas al cubo</span>
                                    </div>
                                    <div class="fields-builder-rows">
                                        <div class="dim-picker" id="cuboFormDims"></div>
                                    </div>
                                    <div class="fields-builder-footer">
                                        <button type="button" class="btn btn-secondary btn-sm" id="cuboDimAddBtn">+ Añadir dimensión</button>
                                    </div>
                                </div>
                            </div>
                            <div class="form-group">
                                <label>Medidas</label>
                                <div class="fields-builder">
                                    <div class="fields-builder-header">
                                        <span>Nombre de la medida</span><span>Tipo</span><span></span><span></span>
                                    </div>
                                    <div class="fields-builder-rows" id="cuboFormRows"></div>
                                    <div class="fields-builder-footer">
                                        <button class="btn btn-secondary btn-sm" id="cuboFormAddField">+ Añadir medida</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button class="btn btn-secondary" id="cuboFormCancel">Cancelar</button>
                            <button class="btn btn-primary" id="cuboFormSave">Guardar</button>
                        </div>
                    </div>`;
                document.body.appendChild(overlay);
            }

            const rowsContainer = overlay.querySelector("#cuboFormRows");
            const dimsList = overlay.querySelector("#cuboFormDims");
            const nameInput = overlay.querySelector("#cuboFormName");

            // Estado local: dimensiones ya añadidas al cubo (se va construyendo con "+ Añadir dimensión")
            let addedIds = dimensionsList.length
                ? dimensionsList.filter(d => selectedDimensionIds.includes(d.DIMENSION_ID)).map(d => d.DIMENSION_ID)
                : [];

            const renderDimsList = () => {
                dimsList.innerHTML = addedIds.length
                    ? addedIds.map(id => {
                        const d = dimensionsList.find(x => x.DIMENSION_ID === id);
                        if (!d) return "";
                        return `
                            <div class="dim-picker-item">
                                <span>${UI.escapeHtml(d.DIMENSION)}</span>
                                <span class="table-tag">${UI.escapeHtml(d.TABLA)}</span>
                                <button type="button" class="dim-picker-remove" data-remove-dim="${id}" title="Quitar">✕</button>
                            </div>`;
                    }).join("")
                    : `<p class="form-hint">${dimensionsList.length ? "Todavía no has añadido ninguna dimensión a este cubo." : 'Este proyecto todavía no tiene dimensiones. Créalas primero desde el menú "Dimensiones".'}</p>`;

                dimsList.querySelectorAll("[data-remove-dim]").forEach(btn => {
                    btn.addEventListener("click", () => {
                        addedIds = addedIds.filter(id => id !== btn.dataset.removeDim);
                        renderDimsList();
                    });
                });
            };

            const addBtn = overlay.querySelector("#cuboDimAddBtn");
            addBtn.onclick = async () => {
                const picked = await UI.openDimensionPickerModal({ dimensionsList, excludeIds: addedIds });
                if (!picked) return;
                addedIds.push(picked);
                renderDimsList();
            };
            addBtn.disabled = !dimensionsList.length;

            renderDimsList();

            const addRow = (f) => rowsContainer.appendChild(UI._fieldRow(f, { showKey: false }));
            rowsContainer.innerHTML = "";
            (measures.length ? measures : []).forEach(m => addRow(m));

            overlay.querySelector("#cuboFormTitle").textContent = title;
            nameInput.value = name;
            nameInput.disabled = !nameEditable;
            overlay.querySelector("#cuboFormDesc").value = description;
            overlay.querySelector("#cuboFormAddField").onclick = () => addRow();

            overlay.classList.add("visible");
            setTimeout(() => nameInput.focus(), 50);

            const cleanup = (result) => {
                overlay.classList.remove("visible");
                resolve(result);
            };

            overlay.querySelector("#cuboFormCancel").onclick = () => cleanup(null);
            overlay.querySelector("#cuboFormClose").onclick = () => cleanup(null);
            overlay.querySelector("#cuboFormSave").onclick = () => {
                const nameVal = nameInput.value.trim();
                if (!nameVal) {
                    UI.toast("Indica un nombre para el cubo.", "error");
                    return;
                }
                const measureRows = UI._readFieldRows(rowsContainer);

                if (!addedIds.length && !measureRows.length) {
                    UI.toast("Añade al menos una dimensión o una medida.", "error");
                    return;
                }

                cleanup({
                    name: nameVal,
                    description: overlay.querySelector("#cuboFormDesc").value.trim(),
                    dimensionIds: addedIds,
                    measures: measureRows
                });
            };
        });
    },

    /** Descarga un contenido de texto/binario como archivo en el navegador */
    downloadBlob(filename, content, mime = "text/plain;charset=utf-8") {
        const blob = (content instanceof Blob) ? content : new Blob([content], { type: mime });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 2000);
    },

    /** Modal de elección entre varias opciones (ej. Sustituir todo / Incremental). Devuelve Promise<key|null>. */
    choiceModal(title, message, options) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("choiceModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "choiceModal";
                document.body.appendChild(overlay);
            }
            overlay.innerHTML = `
                <div class="modal-box">
                    <div class="modal-header">
                        <h3>${UI.escapeHtml(title)}</h3>
                        <button class="modal-close" id="choiceModalClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <p class="confirm-text">${message}</p>
                        <div class="choice-modal-options">
                            ${options.map(o => `<button type="button" class="btn ${o.style === "primary" ? "btn-primary" : "btn-secondary"}" data-choice="${o.key}">${UI.escapeHtml(o.label)}</button>`).join("")}
                        </div>
                    </div>
                </div>`;
            overlay.classList.add("visible");

            const cleanup = (result) => {
                overlay.classList.remove("visible");
                resolve(result);
            };
            overlay.querySelector("#choiceModalClose").onclick = () => cleanup(null);
            overlay.querySelectorAll("[data-choice]").forEach(btn => {
                btn.onclick = () => cleanup(btn.dataset.choice);
            });
        });
    },

    /**
     * Popup de selección de una dimensión con buscador + tabla (Dimensión / Descripción).
     * Pensado para proyectos con muchas dimensiones (fácil de encontrar por texto).
     * Devuelve una Promise<dimensionId|null>.
     */
    openDimensionPickerModal({ dimensionsList = [], excludeIds = [] }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("dimPickerModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "dimPickerModal";
                document.body.appendChild(overlay);
            }

            const available = dimensionsList.filter(d => !excludeIds.includes(d.DIMENSION_ID));

            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <h3>Seleccionar dimensión</h3>
                        <button class="modal-close" id="dimPickerClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <input type="text" id="dimPickerSearch" class="dim-picker-search" placeholder="Buscar por nombre o descripción...">
                        <div class="dim-picker-table-wrap">
                            <table class="dim-picker-table">
                                <thead><tr><th>Dimensión</th><th>Descripción</th></tr></thead>
                                <tbody id="dimPickerRows"></tbody>
                            </table>
                        </div>
                    </div>
                </div>`;

            const searchInput = overlay.querySelector("#dimPickerSearch");
            const rowsEl = overlay.querySelector("#dimPickerRows");

            const cleanup = (result) => {
                overlay.classList.remove("visible");
                resolve(result);
            };

            const renderRows = (filterText = "") => {
                const f = filterText.trim().toLowerCase();
                const filtered = !f ? available : available.filter(d =>
                    (d.DIMENSION || "").toLowerCase().includes(f) || (d.DESCRIPCION || "").toLowerCase().includes(f));

                rowsEl.innerHTML = filtered.length
                    ? filtered.map(d => `
                        <tr data-pick="${d.DIMENSION_ID}">
                            <td><strong>${UI.escapeHtml(d.DIMENSION)}</strong></td>
                            <td>${UI.escapeHtml(d.DESCRIPCION || "—")}</td>
                        </tr>`).join("")
                    : `<tr><td colspan="2" class="dim-picker-empty">${available.length ? "Sin resultados para esa búsqueda." : "No quedan más dimensiones por añadir."}</td></tr>`;

                rowsEl.querySelectorAll("[data-pick]").forEach(tr => {
                    tr.addEventListener("click", () => cleanup(tr.dataset.pick));
                });
            };

            searchInput.value = "";
            searchInput.oninput = () => renderRows(searchInput.value);
            renderRows();

            overlay.classList.add("visible");
            setTimeout(() => searchInput.focus(), 50);

            overlay.querySelector("#dimPickerClose").onclick = () => cleanup(null);
        });
    },

    /**
     * Prompt de una sola línea reutilizable (ej. valor de una constante).
     * Devuelve Promise<string|null>.
     */
    openTextPromptModal({ title, label = "Valor", value = "", placeholder = "" }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("textPromptModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "textPromptModal";
                document.body.appendChild(overlay);
            }
            overlay.innerHTML = `
                <div class="modal-box">
                    <div class="modal-header">
                        <h3>${UI.escapeHtml(title)}</h3>
                        <button class="modal-close" id="textPromptClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>${UI.escapeHtml(label)}</label>
                            <input type="text" id="textPromptInput" placeholder="${UI.escapeHtml(placeholder)}" value="${UI.escapeHtml(value)}">
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="textPromptCancel">Cancelar</button>
                        <button class="btn btn-primary" id="textPromptSave">Guardar</button>
                    </div>
                </div>`;
            overlay.classList.add("visible");
            const input = overlay.querySelector("#textPromptInput");
            setTimeout(() => { input.focus(); input.select(); }, 50);

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#textPromptClose").onclick = () => cleanup(null);
            overlay.querySelector("#textPromptCancel").onclick = () => cleanup(null);
            overlay.querySelector("#textPromptSave").onclick = () => cleanup(input.value);
            input.onkeydown = (e) => { if (e.key === "Enter") cleanup(input.value); };
        });
    },

    /**
     * Editor de código Python (mockup, sin ejecución real). Reutilizado para
     * las funciones "cambiar datos input/output", mapeo por código y función
     * de campo. Devuelve Promise<string|null>.
     */
    openCodeEditorModal({ title, subtitle = "", code = "" }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("codeEditorModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "codeEditorModal";
                document.body.appendChild(overlay);
            }
            const defaultCode = code || `def transformar(df):\n    # df: pandas.DataFrame de entrada\n    # ... tu lógica aquí ...\n    return df\n`;
            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <div>
                            <h3>${UI.escapeHtml(title)}</h3>
                            ${subtitle ? `<span class="modal-subtitle">${UI.escapeHtml(subtitle)}</span>` : ""}
                        </div>
                        <button class="modal-close" id="codeEditorClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <p class="form-hint">Editor de Python (mockup) — todavía no se ejecuta, solo se guarda el código junto a la carga de datos.</p>
                        <textarea id="codeEditorArea" class="code-editor-area" spellcheck="false">${UI.escapeHtml(defaultCode)}</textarea>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="codeEditorCancel">Cancelar</button>
                        <button class="btn btn-primary" id="codeEditorSave">Guardar código</button>
                    </div>
                </div>`;
            overlay.classList.add("visible");

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#codeEditorClose").onclick = () => cleanup(null);
            overlay.querySelector("#codeEditorCancel").onclick = () => cleanup(null);
            overlay.querySelector("#codeEditorSave").onclick = () => cleanup(overlay.querySelector("#codeEditorArea").value);
        });
    },

    FORMULA_FUNCTIONS: {
        "Números": ["SUMA(a, b)", "RESTA(a, b)", "MULTIPLICAR(a, b)", "DIVIDIR(a, b)", "REDONDEAR(valor, decimales)", "ABS(valor)"],
        "Texto": ["CONCATENAR(a, b)", "SI(condicion, si_si, si_no)", "MAYUSCULAS(texto)", "MINUSCULAS(texto)", "SUBCADENA(texto, inicio, longitud)", "TRIM(texto)"]
    },

    /**
     * Editor de fórmulas tipo Excel (mockup): funciones numéricas/texto +
     * campos del input insertables. Devuelve Promise<string|null>.
     */
    openFormulaEditorModal({ title, inputFields = [], value = "" }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("formulaEditorModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "formulaEditorModal";
                document.body.appendChild(overlay);
            }
            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <h3>${UI.escapeHtml(title)}</h3>
                        <button class="modal-close" id="formulaClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Fórmula</label>
                            <textarea id="formulaArea" rows="3" placeholder="Ej. CONCATENAR([NOMBRE], &quot; &quot;, [APELLIDO])">${UI.escapeHtml(value)}</textarea>
                        </div>
                        <div class="formula-editor-cols">
                            <div class="formula-editor-col">
                                <div class="hierarchy-col-label">Funciones</div>
                                <div class="formula-fn-groups" id="formulaFnGroups"></div>
                            </div>
                            <div class="formula-editor-col">
                                <div class="hierarchy-col-label">Campos de entrada</div>
                                <div class="formula-field-chips" id="formulaFieldChips">
                                    ${inputFields.length ? inputFields.map(f => `<button type="button" class="hier-chip" data-insert-field="${UI.escapeHtml(f.name)}">${UI.escapeHtml(f.name)}</button>`).join("") : `<div class="hierarchy-pool-empty">No hay campos de entrada todavía.</div>`}
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="formulaCancel">Cancelar</button>
                        <button class="btn btn-primary" id="formulaSave">Guardar fórmula</button>
                    </div>
                </div>`;

            const groupsEl = overlay.querySelector("#formulaFnGroups");
            groupsEl.innerHTML = Object.entries(UI.FORMULA_FUNCTIONS).map(([group, fns]) => `
                <div class="formula-fn-group">
                    <span class="formula-fn-group-title">${group}</span>
                    <div class="formula-fn-list">
                        ${fns.map(fn => `<button type="button" class="hier-chip" data-insert-fn="${UI.escapeHtml(fn)}">${UI.escapeHtml(fn)}</button>`).join("")}
                    </div>
                </div>`).join("");

            const textarea = overlay.querySelector("#formulaArea");
            const insertAtCursor = (text) => {
                const start = textarea.selectionStart ?? textarea.value.length;
                const end = textarea.selectionEnd ?? textarea.value.length;
                textarea.value = textarea.value.slice(0, start) + text + textarea.value.slice(end);
                textarea.focus();
                textarea.selectionStart = textarea.selectionEnd = start + text.length;
            };

            overlay.querySelectorAll("[data-insert-fn]").forEach(btn =>
                btn.addEventListener("click", () => insertAtCursor(btn.dataset.insertFn)));
            overlay.querySelectorAll("[data-insert-field]").forEach(btn =>
                btn.addEventListener("click", () => insertAtCursor(`[${btn.dataset.insertField}]`)));

            overlay.classList.add("visible");
            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#formulaClose").onclick = () => cleanup(null);
            overlay.querySelector("#formulaCancel").onclick = () => cleanup(null);
            overlay.querySelector("#formulaSave").onclick = () => cleanup(textarea.value.trim());
        });
    },

    /**
     * Popup de selección de tabla origen (BigQuery o Snowflake). Mockup:
     * lista de tablas de ejemplo con buscador, igual que el picker de
     * dimensiones. Devuelve Promise<{tableName, fields}|null>.
     */
    openTablePickerModal({ connector, tables = [] }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("tablePickerModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "tablePickerModal";
                document.body.appendChild(overlay);
            }

            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <div>
                            <h3>Seleccionar tabla de origen</h3>
                            <span class="modal-subtitle">Conector: ${UI.escapeHtml(connector)}</span>
                        </div>
                        <button class="modal-close" id="tablePickerClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <input type="text" id="tablePickerSearch" class="dim-picker-search" placeholder="Buscar tabla...">
                        <div class="dim-picker-table-wrap">
                            <table class="dim-picker-table">
                                <thead><tr><th>Tabla</th><th>Columnas</th></tr></thead>
                                <tbody id="tablePickerRows"></tbody>
                            </table>
                        </div>
                        <p class="form-hint">Listado de ejemplo (mockup) — todavía no consulta el catálogo real de ${UI.escapeHtml(connector)}.</p>
                    </div>
                </div>`;

            const searchInput = overlay.querySelector("#tablePickerSearch");
            const rowsEl = overlay.querySelector("#tablePickerRows");
            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };

            const renderRows = (filterText = "") => {
                const f = filterText.trim().toLowerCase();
                const filtered = !f ? tables : tables.filter(t => t.name.toLowerCase().includes(f));
                rowsEl.innerHTML = filtered.length
                    ? filtered.map((t, i) => `
                        <tr data-pick="${i}">
                            <td><strong>${UI.escapeHtml(t.name)}</strong></td>
                            <td>${t.fields.map(f2 => `<span class="table-tag">${UI.escapeHtml(f2.name)}</span>`).join(" ")}</td>
                        </tr>`).join("")
                    : `<tr><td colspan="2" class="dim-picker-empty">Sin resultados.</td></tr>`;
                rowsEl.querySelectorAll("[data-pick]").forEach(tr => {
                    tr.addEventListener("click", () => cleanup(tables[parseInt(tr.dataset.pick, 10)]));
                });
            };

            searchInput.value = "";
            searchInput.oninput = () => renderRows(searchInput.value);
            renderRows();

            overlay.classList.add("visible");
            setTimeout(() => searchInput.focus(), 50);
            overlay.querySelector("#tablePickerClose").onclick = () => cleanup(null);
        });
    },

    /** Activa el comportamiento de maximizar/contraer para todos los .block de la página */
    initBlockControls() {
        const grid = document.getElementById("blocksGrid");
        document.querySelectorAll(".block").forEach(block => {
            const maxBtn = block.querySelector('[data-action="maximize"]');
            const collapseBtn = block.querySelector('[data-action="collapse"]');

            if (maxBtn) {
                maxBtn.addEventListener("click", () => {
                    const isMax = block.classList.toggle("maximized");
                    grid.classList.toggle("has-maximized", isMax);
                    if (isMax) block.classList.remove("collapsed");
                });
            }

            if (collapseBtn) {
                collapseBtn.addEventListener("click", () => {
                    block.classList.toggle("collapsed");
                    collapseBtn.style.transform = block.classList.contains("collapsed") ? "rotate(-90deg)" : "";
                });
            }
        });
    }
};
