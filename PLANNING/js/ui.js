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

    // Tipos válidos para una variable de pantalla de un FLUJO (superconjunto de
    // FIELD_TYPES + FILE, que renderiza un <input type="file"> en flow_run.html
    // y cuyo valor resuelto es la ruta de storage tras subir el fichero).
    SCREEN_VARIABLE_TYPES: ["STRING", "INTEGER", "FLOAT", "NUMERIC", "BOOLEAN", "DATE", "DATETIME", "TIMESTAMP", "FILE"],

    /** Crea una fila de campo reutilizable para los distintos diseñadores (atributos / medidas) */
    /**
     * `showDesc` activa la columna "Descripción" (icono 🏷): marca qué
     * atributo de la dimensión es el atributo descriptivo. Es un estado
     * MUTUAMENTE EXCLUSIVO entre filas del mismo diseñador (como un radio),
     * pero opcional: puede no haber ninguno marcado. Al no ser un
     * <input type="radio"> nativo, se puede desmarcar volviendo a pulsar
     * el que ya está activo.
     */
    _fieldRow(f = { name: "", type: "STRING", key: false, isDescription: false }, { showKey = true, showDesc = false } = {}) {
        const row = document.createElement("div");
        row.className = "field-row";
        row.dataset.isDesc = f.isDescription ? "1" : "0";
        row.innerHTML = `
            <input type="text" class="f-name" placeholder="ej. codigo_cliente" value="${UI.escapeHtml(f.name)}">
            <select class="f-type">
                ${UI.FIELD_TYPES.map(t => `<option value="${t}" ${t === f.type ? "selected" : ""}>${t}</option>`).join("")}
            </select>
            <span class="field-key">${showKey ? `<input type="checkbox" class="f-key" ${f.key ? "checked" : ""}>` : ""}</span>
            ${showDesc ? `<span class="field-desc"><button type="button" class="field-desc-toggle${f.isDescription ? " active" : ""}" title="Marcar como atributo descriptivo">🏷</button></span>` : ""}
            <button type="button" class="field-remove" title="Eliminar">✕</button>`;
        row.querySelector(".field-remove").addEventListener("click", () => row.remove());
        if (showDesc) {
            row.querySelector(".field-desc-toggle").addEventListener("click", (e) => {
                const nowActive = row.dataset.isDesc === "1";
                // Desmarca cualquier otra fila del mismo diseñador (solo puede haber una).
                row.parentElement.querySelectorAll(".field-row").forEach(r => {
                    r.dataset.isDesc = "0";
                    const btn = r.querySelector(".field-desc-toggle");
                    if (btn) btn.classList.remove("active");
                });
                if (!nowActive) {
                    row.dataset.isDesc = "1";
                    e.currentTarget.classList.add("active");
                }
            });
        }
        return row;
    },

    _readFieldRows(container) {
        return Array.from(container.querySelectorAll(".field-row")).map(r => ({
            name: r.querySelector(".f-name").value.trim(),
            type: r.querySelector(".f-type").value,
            key: !!r.querySelector(".f-key") && r.querySelector(".f-key").checked,
            isDescription: r.dataset.isDesc === "1"
        })).filter(f => f.name);
    },

    /**
     * Fila de campo para el diseñador de "Tablas de parametrización":
     * Nombre, Descripción (texto libre), Tipo y Clave. A diferencia de
     * `_fieldRow` (usada en Dimensiones/Cubos), aquí la descripción es un
     * campo de texto por cada fila, no un marcador único tipo radio.
     */
    _paramFieldRow(f = { name: "", description: "", type: "STRING", key: false }) {
        const row = document.createElement("div");
        row.className = "field-row param-field-row";
        row.innerHTML = `
            <input type="text" class="f-name" placeholder="ej. codigo" value="${UI.escapeHtml(f.name)}">
            <input type="text" class="f-description" placeholder="Descripción del campo" value="${UI.escapeHtml(f.description || "")}">
            <select class="f-type">
                ${UI.FIELD_TYPES.map(t => `<option value="${t}" ${t === f.type ? "selected" : ""}>${t}</option>`).join("")}
            </select>
            <span class="field-key"><input type="checkbox" class="f-key" ${f.key ? "checked" : ""}></span>
            <button type="button" class="field-remove" title="Eliminar">✕</button>`;
        row.querySelector(".field-remove").addEventListener("click", () => row.remove());
        return row;
    },

    _readParamFieldRows(container) {
        return Array.from(container.querySelectorAll(".param-field-row")).map(r => ({
            name: r.querySelector(".f-name").value.trim(),
            description: r.querySelector(".f-description").value.trim(),
            type: r.querySelector(".f-type").value,
            key: r.querySelector(".f-key").checked
        })).filter(f => f.name);
    },

    /**
     * Modal de "Tabla de parametrización": nombre + descripción + diseñador
     * de campos (nombre, descripción, tipo, clave). No hay clave automática
     * (a diferencia de Dimensiones): el usuario marca la(s) clave(s) él mismo.
     * Devuelve una Promise<{name, description, fields}|null>.
     */
    openParamTableFormModal({ title, name = "", description = "", fields = [], nameEditable = true }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("paramFormModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "paramFormModal";
                overlay.innerHTML = `
                    <div class="modal-box modal-wide">
                        <div class="modal-header">
                            <h3 id="paramFormTitle"></h3>
                            <button class="modal-close" id="paramFormClose">&times;</button>
                        </div>
                        <div class="modal-body">
                            <div class="form-row">
                                <div class="form-group">
                                    <label>Nombre de la tabla</label>
                                    <input type="text" id="paramFormName" placeholder="Ej. Parametros_calculo">
                                </div>
                            </div>
                            <div class="form-group">
                                <label>Descripción</label>
                                <textarea id="paramFormDesc" rows="2"></textarea>
                            </div>
                            <div class="form-group">
                                <label>Campos</label>
                                <div class="fields-builder fields-builder--params">
                                    <div class="fields-builder-header fields-builder-header--params">
                                        <span>Nombre del campo</span><span>Descripción</span><span>Tipo</span><span class="col-center">Clave</span><span></span>
                                    </div>
                                    <div class="fields-builder-rows" id="paramFormRows"></div>
                                    <div class="fields-builder-footer">
                                        <button class="btn btn-secondary btn-sm" id="paramFormAddField">+ Añadir campo</button>
                                    </div>
                                </div>
                                <p class="form-hint">Marca "Clave" en, al menos, un campo (se admite clave compuesta).</p>
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button class="btn btn-secondary" id="paramFormCancel">Cancelar</button>
                            <button class="btn btn-primary" id="paramFormSave">Guardar</button>
                        </div>
                    </div>`;
                document.body.appendChild(overlay);
            }

            const rowsContainer = overlay.querySelector("#paramFormRows");
            const nameInput = overlay.querySelector("#paramFormName");

            const addRow = (f) => rowsContainer.appendChild(UI._paramFieldRow(f));
            rowsContainer.innerHTML = "";
            (fields.length ? fields : [{ name: "", description: "", type: "STRING", key: true }]).forEach(f => addRow(f));

            overlay.querySelector("#paramFormTitle").textContent = title;
            nameInput.value = name;
            nameInput.disabled = !nameEditable;
            overlay.querySelector("#paramFormDesc").value = description;
            overlay.querySelector("#paramFormAddField").onclick = () => addRow();

            overlay.classList.add("visible");
            setTimeout(() => nameInput.focus(), 50);

            const cleanup = (result) => {
                overlay.classList.remove("visible");
                resolve(result);
            };

            overlay.querySelector("#paramFormCancel").onclick = () => cleanup(null);
            overlay.querySelector("#paramFormClose").onclick = () => cleanup(null);
            overlay.querySelector("#paramFormSave").onclick = () => {
                const nameVal = nameInput.value.trim();
                if (!nameVal) {
                    UI.toast("Indica un nombre para la tabla.", "error");
                    return;
                }
                const fieldsVal = UI._readParamFieldRows(rowsContainer);
                if (!fieldsVal.length) {
                    UI.toast("Añade al menos un campo.", "error");
                    return;
                }
                if (!fieldsVal.some(f => f.key)) {
                    UI.toast("Marca al menos un campo como clave.", "error");
                    return;
                }
                cleanup({
                    name: nameVal,
                    description: overlay.querySelector("#paramFormDesc").value.trim(),
                    fields: fieldsVal
                });
            };
        });
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
                                <div class="fields-builder fields-builder--with-desc">
                                    <div class="fields-builder-header">
                                        <span>Nombre del atributo</span><span>Tipo</span><span class="col-center">Clave</span><span class="col-center">Desc.</span><span></span>
                                    </div>
                                    <div class="fields-builder-rows" id="dimFormRows"></div>
                                    <div class="fields-builder-footer">
                                        <button class="btn btn-secondary btn-sm" id="dimFormAddField">+ Añadir atributo</button>
                                    </div>
                                </div>
                                <p class="form-hint">Marca "Clave" en un atributo si quieres una clave compuesta (ej. Clase de coste + Sociedad). Marca 🏷 en, como mucho, un atributo para indicar cuál es el atributo descriptivo de la dimensión (opcional).</p>
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

            const addRow = (f) => rowsContainer.appendChild(UI._fieldRow(f, { showKey: true, showDesc: true }));
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
     * Popup de filtro para un campo de entrada: permite elegir si el filtro
     * se resuelve por una constante (con valor) o por una variable (por
     * ahora solo se guarda que es de tipo variable, sin catálogo todavía).
     * Devuelve Promise<{type, value}|"remove"|null>.
     */
    openFilterFieldModal({ title = "Filtro de campo", fieldName = "", current = null }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("filterFieldModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "filterFieldModal";
                document.body.appendChild(overlay);
            }

            let type = current && current.type ? current.type : "constante";
            let value = current && current.type === "constante" ? (current.value || "") : "";

            const renderBody = () => {
                const valueBlock = type === "constante"
                    ? `<div class="form-group">
                            <label>Valor constante</label>
                            <input type="text" id="filterFieldValue" placeholder="Ej. 2024" value="${UI.escapeHtml(value)}">
                       </div>`
                    : `<p class="form-hint">El valor de esta variable se asigna en cada flujo, directamente sobre este campo (en la pestaña Flujos → Mapeo de variables).</p>`;

                overlay.innerHTML = `
                    <div class="modal-box">
                        <div class="modal-header">
                            <div>
                                <h3>${UI.escapeHtml(title)}</h3>
                                ${fieldName ? `<span class="modal-subtitle">${UI.escapeHtml(fieldName)}</span>` : ""}
                            </div>
                            <button class="modal-close" id="filterFieldClose">&times;</button>
                        </div>
                        <div class="modal-body">
                            <div class="form-group">
                                <label>Tipo de filtro</label>
                                <div class="segmented" id="filterFieldType">
                                    <button type="button" class="segmented-btn ${type === "constante" ? "active" : ""}" data-ftype="constante">Constante</button>
                                    <button type="button" class="segmented-btn ${type === "variable" ? "active" : ""}" data-ftype="variable">Variable</button>
                                </div>
                            </div>
                            ${valueBlock}
                        </div>
                        <div class="modal-footer">
                            ${current ? `<button class="btn btn-secondary" id="filterFieldRemove">Quitar filtro</button><span class="load-fn-toolbar-spacer"></span>` : ""}
                            <button class="btn btn-secondary" id="filterFieldCancel">Cancelar</button>
                            <button class="btn btn-primary" id="filterFieldSave">Guardar</button>
                        </div>
                    </div>`;

                overlay.querySelectorAll("#filterFieldType [data-ftype]").forEach(btn => {
                    btn.addEventListener("click", () => {
                        if (btn.dataset.ftype === type) return;
                        if (type === "constante") {
                            const valInput = overlay.querySelector("#filterFieldValue");
                            if (valInput) value = valInput.value;
                        }
                        type = btn.dataset.ftype;
                        renderBody();
                    });
                });

                const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
                overlay.querySelector("#filterFieldClose").onclick = () => cleanup(null);
                overlay.querySelector("#filterFieldCancel").onclick = () => cleanup(null);
                const removeBtn = overlay.querySelector("#filterFieldRemove");
                if (removeBtn) removeBtn.onclick = () => cleanup("remove");
                overlay.querySelector("#filterFieldSave").onclick = () => {
                    if (type === "constante") {
                        const valInput = overlay.querySelector("#filterFieldValue");
                        cleanup({ type: "constante", value: valInput ? valInput.value : "" });
                    } else {
                        cleanup({ type: "variable", value: "" });
                    }
                };

                const focusInput = overlay.querySelector("#filterFieldValue");
                if (focusInput) setTimeout(() => { focusInput.focus(); focusInput.select(); }, 50);
            };

            renderBody();
            overlay.classList.add("visible");
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

    // ==========================================================
    // Modales específicos de FLUJOS DE CARGA
    // ==========================================================

    /**
     * Planificador de ejecución para flujos automáticos. Devuelve
     * Promise<schedule|"remove"|null>.
     */
    openScheduleModal({ current = null } = {}) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("scheduleModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "scheduleModal";
                document.body.appendChild(overlay);
            }

            const s = Object.assign({
                startDate: "", startTime: "06:00", repeat: "ninguna",
                weekDays: [], dayOfMonth: 1, intervalValue: 1, intervalUnit: "horas", active: true
            }, current || {});

            const weekDayLabels = ["L", "M", "X", "J", "V", "S", "D"];

            const renderBody = () => {
                overlay.innerHTML = `
                    <div class="modal-box">
                        <div class="modal-header">
                            <h3>Planificar ejecución</h3>
                            <button class="modal-close" id="scheduleClose">&times;</button>
                        </div>
                        <div class="modal-body">
                            <div class="load-origin-row">
                                <div class="form-group">
                                    <label>Fecha de inicio</label>
                                    <input type="date" id="schedDate" value="${UI.escapeHtml(s.startDate)}">
                                </div>
                                <div class="form-group">
                                    <label>Hora de inicio</label>
                                    <input type="time" id="schedTime" value="${UI.escapeHtml(s.startTime)}">
                                </div>
                            </div>
                            <div class="form-group">
                                <label>Repetición</label>
                                <select id="schedRepeat">
                                    <option value="ninguna" ${s.repeat === "ninguna" ? "selected" : ""}>No repetir (una sola vez)</option>
                                    <option value="diaria" ${s.repeat === "diaria" ? "selected" : ""}>Diaria</option>
                                    <option value="semanal" ${s.repeat === "semanal" ? "selected" : ""}>Semanal</option>
                                    <option value="mensual" ${s.repeat === "mensual" ? "selected" : ""}>Mensual</option>
                                    <option value="personalizada" ${s.repeat === "personalizada" ? "selected" : ""}>Personalizada (cada X)</option>
                                </select>
                            </div>
                            ${s.repeat === "semanal" ? `
                                <div class="form-group">
                                    <label>Días de la semana</label>
                                    <div class="weekday-picker" id="schedWeekDays">
                                        ${weekDayLabels.map((d, i) => `<button type="button" class="weekday-btn ${s.weekDays.includes(i) ? "active" : ""}" data-day="${i}">${d}</button>`).join("")}
                                    </div>
                                </div>` : ""}
                            ${s.repeat === "mensual" ? `
                                <div class="form-group">
                                    <label>Día del mes</label>
                                    <input type="number" id="schedDayOfMonth" min="1" max="31" value="${s.dayOfMonth}">
                                </div>` : ""}
                            ${s.repeat === "personalizada" ? `
                                <div class="load-origin-row">
                                    <div class="form-group">
                                        <label>Cada</label>
                                        <input type="number" id="schedIntervalValue" min="1" value="${s.intervalValue}">
                                    </div>
                                    <div class="form-group">
                                        <label>Unidad</label>
                                        <select id="schedIntervalUnit">
                                            <option value="minutos" ${s.intervalUnit === "minutos" ? "selected" : ""}>Minutos</option>
                                            <option value="horas" ${s.intervalUnit === "horas" ? "selected" : ""}>Horas</option>
                                            <option value="dias" ${s.intervalUnit === "dias" ? "selected" : ""}>Días</option>
                                        </select>
                                    </div>
                                </div>` : ""}
                            <div class="form-group form-group--check">
                                <label><input type="checkbox" id="schedActive" ${s.active ? "checked" : ""}> Planificación activa</label>
                            </div>
                        </div>
                        <div class="modal-footer">
                            ${current ? `<button class="btn btn-secondary" id="schedRemove">Quitar planificación</button><span class="load-fn-toolbar-spacer"></span>` : ""}
                            <button class="btn btn-secondary" id="schedCancel">Cancelar</button>
                            <button class="btn btn-primary" id="schedSave">Guardar</button>
                        </div>
                    </div>`;

                const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
                overlay.querySelector("#scheduleClose").onclick = () => cleanup(null);
                overlay.querySelector("#schedCancel").onclick = () => cleanup(null);
                const removeBtn = overlay.querySelector("#schedRemove");
                if (removeBtn) removeBtn.onclick = () => cleanup("remove");

                overlay.querySelector("#schedRepeat").addEventListener("change", (e) => {
                    s.repeat = e.target.value;
                    syncFromInputs();
                    renderBody();
                });

                const weekWrap = overlay.querySelector("#schedWeekDays");
                if (weekWrap) {
                    weekWrap.querySelectorAll("[data-day]").forEach(btn => {
                        btn.addEventListener("click", () => {
                            const day = parseInt(btn.dataset.day, 10);
                            const idx = s.weekDays.indexOf(day);
                            if (idx >= 0) s.weekDays.splice(idx, 1); else s.weekDays.push(day);
                            btn.classList.toggle("active");
                        });
                    });
                }

                const syncFromInputs = () => {
                    s.startDate = overlay.querySelector("#schedDate").value;
                    s.startTime = overlay.querySelector("#schedTime").value;
                    const dom = overlay.querySelector("#schedDayOfMonth");
                    if (dom) s.dayOfMonth = parseInt(dom.value, 10) || 1;
                    const iv = overlay.querySelector("#schedIntervalValue");
                    if (iv) s.intervalValue = parseInt(iv.value, 10) || 1;
                    const iu = overlay.querySelector("#schedIntervalUnit");
                    if (iu) s.intervalUnit = iu.value;
                    s.active = overlay.querySelector("#schedActive").checked;
                };

                overlay.querySelector("#schedSave").onclick = () => {
                    syncFromInputs();
                    if (!s.startDate) { UI.toast("Indica una fecha de inicio.", "error"); return; }
                    cleanup({ ...s });
                };
            };

            renderBody();
            overlay.classList.add("visible");
        });
    },

    /**
     * Selector de interfaz (carga de datos) para añadir a la cadena de un
     * flujo. Devuelve Promise<interfaceObj|null>.
     */
    openInterfacePickerModal({ interfaces = [], cubeNameById = {} }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("interfacePickerModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "interfacePickerModal";
                document.body.appendChild(overlay);
            }

            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <h3>Añadir interfaz a la cadena</h3>
                        <button class="modal-close" id="ifacePickerClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <input type="text" id="ifacePickerSearch" class="dim-picker-search" placeholder="Buscar interfaz...">
                        <div class="dim-picker-table-wrap">
                            <table class="dim-picker-table">
                                <thead><tr><th>Interfaz</th><th>Cubo</th><th>Origen</th></tr></thead>
                                <tbody id="ifacePickerRows"></tbody>
                            </table>
                        </div>
                        <p class="form-hint">La misma interfaz se puede añadir varias veces a la cadena.</p>
                    </div>
                </div>`;

            const searchInput = overlay.querySelector("#ifacePickerSearch");
            const rowsEl = overlay.querySelector("#ifacePickerRows");
            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };

            const renderRows = (filterText = "") => {
                const f = filterText.trim().toLowerCase();
                const filtered = !f ? interfaces : interfaces.filter(i => i.name.toLowerCase().includes(f));
                rowsEl.innerHTML = filtered.length
                    ? filtered.map((i, idx) => `
                        <tr data-pick="${idx}">
                            <td><strong>${UI.escapeHtml(i.name)}</strong></td>
                            <td>${UI.escapeHtml(cubeNameById[i.cuboId] || "—")}</td>
                            <td><span class="table-tag">${i.originType === "tabla" ? "Tabla" : "Fichero"}</span></td>
                        </tr>`).join("")
                    : `<tr><td colspan="3" class="dim-picker-empty">${interfaces.length ? "Sin resultados." : "Todavía no hay interfaces creadas."}</td></tr>`;
                rowsEl.querySelectorAll("[data-pick]").forEach(tr => {
                    tr.addEventListener("click", () => cleanup(filtered[parseInt(tr.dataset.pick, 10)]));
                });
            };

            searchInput.value = "";
            searchInput.oninput = () => renderRows(searchInput.value);
            renderRows();

            overlay.classList.add("visible");
            setTimeout(() => searchInput.focus(), 50);
            overlay.querySelector("#ifacePickerClose").onclick = () => cleanup(null);
        });
    },

    /**
     * Popup para asignar el valor de un "target" del bloque de mapeo de un
     * flujo (variables de fichero / filtro / mapeo de cada paso de la
     * cadena): por constante, o por una variable de la pantalla (si el
     * flujo es manual y hay variables definidas). Devuelve
     * Promise<{type,value}|"remove"|null>.
     */
    openFlowTargetModal({ title = "Asignar valor", targetLabel = "", screenVariables = [], current = null }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("flowTargetModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "flowTargetModal";
                document.body.appendChild(overlay);
            }

            const hasVars = screenVariables.length > 0;
            let type = current && current.type === "variable" && hasVars ? "variable" : "constante";
            let constValue = current && current.type === "constante" ? (current.value || "") : "";
            let varValue = current && current.type === "variable" ? (current.value || "") : (hasVars ? screenVariables[0].name : "");

            const renderBody = () => {
                const valueBlock = type === "constante"
                    ? `<div class="form-group">
                            <label>Valor constante</label>
                            <input type="text" id="flowTargetValue" value="${UI.escapeHtml(constValue)}">
                       </div>`
                    : `<div class="form-group">
                            <label>Variable de pantalla</label>
                            <select id="flowTargetVar">
                                ${screenVariables.map(v => `<option value="${UI.escapeHtml(v.name)}" ${v.name === varValue ? "selected" : ""}>${UI.escapeHtml(v.label || v.name)}</option>`).join("")}
                            </select>
                       </div>`;

                overlay.innerHTML = `
                    <div class="modal-box">
                        <div class="modal-header">
                            <div>
                                <h3>${UI.escapeHtml(title)}</h3>
                                ${targetLabel ? `<span class="modal-subtitle">${UI.escapeHtml(targetLabel)}</span>` : ""}
                            </div>
                            <button class="modal-close" id="flowTargetClose">&times;</button>
                        </div>
                        <div class="modal-body">
                            ${hasVars ? `
                                <div class="form-group">
                                    <label>Origen del valor</label>
                                    <div class="segmented" id="flowTargetType">
                                        <button type="button" class="segmented-btn ${type === "constante" ? "active" : ""}" data-vtype="constante">Constante</button>
                                        <button type="button" class="segmented-btn ${type === "variable" ? "active" : ""}" data-vtype="variable">Variable</button>
                                    </div>
                                </div>` : `<p class="form-hint">Flujo sin pantalla: el valor solo se puede fijar por constante.</p>`}
                            ${valueBlock}
                        </div>
                        <div class="modal-footer">
                            ${current ? `<button class="btn btn-secondary" id="flowTargetRemove">Quitar</button><span class="load-fn-toolbar-spacer"></span>` : ""}
                            <button class="btn btn-secondary" id="flowTargetCancel">Cancelar</button>
                            <button class="btn btn-primary" id="flowTargetSave">Guardar</button>
                        </div>
                    </div>`;

                const typeWrap = overlay.querySelector("#flowTargetType");
                if (typeWrap) {
                    typeWrap.querySelectorAll("[data-vtype]").forEach(btn => {
                        btn.addEventListener("click", () => {
                            if (btn.dataset.vtype === type) return;
                            if (type === "constante") { const i = overlay.querySelector("#flowTargetValue"); if (i) constValue = i.value; }
                            else { const s = overlay.querySelector("#flowTargetVar"); if (s) varValue = s.value; }
                            type = btn.dataset.vtype;
                            renderBody();
                        });
                    });
                }

                const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
                overlay.querySelector("#flowTargetClose").onclick = () => cleanup(null);
                overlay.querySelector("#flowTargetCancel").onclick = () => cleanup(null);
                const removeBtn = overlay.querySelector("#flowTargetRemove");
                if (removeBtn) removeBtn.onclick = () => cleanup("remove");
                overlay.querySelector("#flowTargetSave").onclick = () => {
                    if (type === "constante") {
                        const i = overlay.querySelector("#flowTargetValue");
                        cleanup({ type: "constante", value: i ? i.value : "" });
                    } else {
                        const s = overlay.querySelector("#flowTargetVar");
                        cleanup({ type: "variable", value: s ? s.value : "" });
                    }
                };

                const focusInput = overlay.querySelector("#flowTargetValue");
                if (focusInput) setTimeout(() => { focusInput.focus(); focusInput.select(); }, 50);
            };

            renderBody();
            overlay.classList.add("visible");
        });
    },

    /**
     * Selector de valores concretos de una dimensión (multi-selección),
     * usado por el driver de un paso de Workflow. Consulta la tabla física
     * de la dimensión (columna clave = nombre de la dimensión) y deja
     * marcar/desmarcar filas. Devuelve Promise<string[]|null>.
     */
    openDimensionValuesPickerModal({ project, dim, keyCol, selected = [] }) {
        return new Promise(async (resolve) => {
            let overlay = document.getElementById("dimValuesPickerModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "dimValuesPickerModal";
                document.body.appendChild(overlay);
            }

            const picked = new Set(selected);
            let allValues = [];

            const shell = (bodyHtml) => `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <div>
                            <h3>Valores de ${UI.escapeHtml(dim[Dimensions.NAME_COL])}</h3>
                            <span class="modal-subtitle">Selecciona los valores sobre los que se repartirá este paso</span>
                        </div>
                        <button class="modal-close" id="dimValPickerClose">&times;</button>
                    </div>
                    <div class="modal-body">${bodyHtml}</div>
                </div>`;

            overlay.innerHTML = shell(`<span class="spinner"></span>`);
            overlay.classList.add("visible");
            overlay.querySelector("#dimValPickerClose").onclick = () => { overlay.classList.remove("visible"); resolve(null); };

            try {
                const fullTable = Provider.qualify(project.DATASET, dim.TABLA);
                const rows = await Provider.runQuery(`SELECT DISTINCT ${keyCol} AS V FROM ${fullTable} ORDER BY ${keyCol} LIMIT 5000`);
                allValues = rows.map(r => String(r.V));
            } catch (err) {
                overlay.innerHTML = shell(`<div class="module-empty">Error al leer los valores: ${UI.escapeHtml(err.message)}</div>`);
                return;
            }

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };

            const renderRows = (filterText = "") => {
                const f = filterText.trim().toLowerCase();
                const filtered = !f ? allValues : allValues.filter(v => v.toLowerCase().includes(f));
                const rowsHtml = filtered.length
                    ? filtered.map(v => `
                        <tr data-toggle-val="${UI.escapeHtml(v)}">
                            <td><input type="checkbox" ${picked.has(v) ? "checked" : ""}></td>
                            <td>${UI.escapeHtml(v)}</td>
                        </tr>`).join("")
                    : `<tr><td colspan="2" class="dim-picker-empty">${allValues.length ? "Sin resultados." : "Esta dimensión todavía no tiene valores cargados."}</td></tr>`;

                overlay.innerHTML = shell(`
                    <input type="text" id="dimValPickerSearch" class="dim-picker-search" placeholder="Buscar valor...">
                    <p class="form-hint">${picked.size} valor(es) seleccionado(s). Si no seleccionas ninguno, el paso se repartirá entre <strong>todos</strong> los valores de la dimensión.</p>
                    <div class="dim-picker-table-wrap">
                        <table class="dim-picker-table">
                            <thead><tr><th style="width:32px;"></th><th>${UI.escapeHtml(dim[Dimensions.NAME_COL])}</th></tr></thead>
                            <tbody id="dimValPickerRows">${rowsHtml}</tbody>
                        </table>
                    </div>
                    <div class="modal-footer" style="padding:14px 0 0;">
                        <button class="btn btn-secondary" id="dimValPickerCancel">Cancelar</button>
                        <button class="btn btn-primary" id="dimValPickerSave">Guardar selección</button>
                    </div>`);

                overlay.querySelector("#dimValPickerClose").onclick = () => cleanup(null);
                overlay.querySelector("#dimValPickerCancel").onclick = () => cleanup(null);
                overlay.querySelector("#dimValPickerSave").onclick = () => cleanup(Array.from(picked));
                const search = overlay.querySelector("#dimValPickerSearch");
                search.value = filterText;
                search.oninput = () => renderRows(search.value);
                setTimeout(() => { search.focus(); search.selectionStart = search.selectionEnd = search.value.length; }, 30);

                overlay.querySelectorAll("[data-toggle-val]").forEach(tr => {
                    tr.addEventListener("click", (e) => {
                        const v = tr.dataset.toggleVal;
                        if (picked.has(v)) picked.delete(v); else picked.add(v);
                        renderRows(search.value);
                    });
                });
            };

            renderRows();
        });
    },

    /**
     * Selector de un flujo manual (Flujos de carga con TIPO = manual) para
     * usarlo como tarea de un paso de Workflow. Devuelve Promise<object|null>
     * con el flujo elegido (de la lista `flows` recibida).
     */
    /**
     * Modal "Nueva tarea" (y edición): nombre + descripción + selector de
     * tipo con tarjetas (mismo patrón visual que "Nueva interfaz" en
     * Interfaces: .origin-type-grid / .origin-type-card).
     * `types`: [{key,label,icon}]. `locked`: en edición no se puede cambiar
     * el tipo de una tarea ya creada (se ve la tarjeta activa pero deshabilitada).
     * Devuelve Promise<{name, description, tipo}|null>.
     */
    openTaskFormModal({ title, name = "", description = "", tipo = "", types = [], locked = false }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("wfTaskFormModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wfTaskFormModal";
                document.body.appendChild(overlay);
            }

            let selected = tipo || (types[0] && types[0].key) || "";

            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <h3>${UI.escapeHtml(title)}</h3>
                        <button class="modal-close" id="wfTaskFormClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Nombre de la tarea</label>
                            <input type="text" id="wfTaskFormName" placeholder="Ej. Actualizar tipos de cambio">
                        </div>
                        <div class="form-group">
                            <label>Descripción</label>
                            <textarea id="wfTaskFormDesc" rows="2" placeholder="Opcional"></textarea>
                        </div>
                        <div class="form-group">
                            <label>Tipo de tarea</label>
                            <div class="origin-type-grid origin-type-grid--wf" id="wfTaskFormTypes">
                                ${types.map(t => `
                                    <button type="button" class="origin-type-card ${locked && t.key !== selected ? "is-disabled" : ""} ${t.key === selected ? "active" : ""}"
                                            data-type="${t.key}" ${locked && t.key !== selected ? "disabled" : ""}>
                                        <span class="origin-type-card-icon">${t.icon || "•"}</span>
                                        <span class="origin-type-card-label">${UI.escapeHtml(t.label)}</span>
                                    </button>`).join("")}
                            </div>
                            ${locked ? `<p class="form-hint">El tipo de una tarea no se puede cambiar una vez creada. Elimínala y crea una nueva si necesitas otro tipo.</p>` : ""}
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="wfTaskFormCancel">Cancelar</button>
                        <button class="btn btn-primary" id="wfTaskFormNext">Continuar</button>
                    </div>
                </div>`;

            overlay.querySelectorAll("#wfTaskFormTypes [data-type]:not(:disabled)").forEach(btn => {
                btn.addEventListener("click", () => {
                    selected = btn.dataset.type;
                    overlay.querySelectorAll("#wfTaskFormTypes [data-type]").forEach(b => b.classList.toggle("active", b === btn));
                });
            });

            const nameInput = overlay.querySelector("#wfTaskFormName");
            nameInput.value = name;
            overlay.querySelector("#wfTaskFormDesc").value = description;

            overlay.classList.add("visible");
            setTimeout(() => { nameInput.focus(); nameInput.select(); }, 50);

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#wfTaskFormClose").onclick = () => cleanup(null);
            overlay.querySelector("#wfTaskFormCancel").onclick = () => cleanup(null);
            overlay.querySelector("#wfTaskFormNext").onclick = () => {
                const nameVal = nameInput.value.trim();
                if (!nameVal) { UI.toast("Indica un nombre para la tarea.", "error"); return; }
                if (!selected) { UI.toast("Selecciona un tipo de tarea.", "error"); return; }
                cleanup({ name: nameVal, description: overlay.querySelector("#wfTaskFormDesc").value.trim(), tipo: selected });
            };
        });
    },

    /**
     * Selector de una "Actualización de tablas" existente del proyecto
     * (catálogo de js/table-updates.js), para usarla como referencia de una
     * tarea de Workflow. Devuelve Promise<{id,name}|null>.
     */
    openActualizacionPickerModal({ items = [] } = {}) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("wfActPickerModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wfActPickerModal";
                document.body.appendChild(overlay);
            }
            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <h3>Seleccionar actualización de tablas</h3>
                        <button class="modal-close" id="wfActPickerClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <input type="text" id="wfActPickerSearch" class="dim-picker-search" placeholder="Buscar por nombre...">
                        <div class="dim-picker-table-wrap">
                            <table class="dim-picker-table">
                                <thead><tr><th>Actualización</th><th>Tabla</th></tr></thead>
                                <tbody id="wfActPickerRows"></tbody>
                            </table>
                        </div>
                        ${!items.length ? `<p class="form-hint">Todavía no hay ninguna "Actualización de tablas" definida en este proyecto. Créala primero en Administración → Actualización de tablas.</p>` : ""}
                    </div>
                </div>`;

            const searchInput = overlay.querySelector("#wfActPickerSearch");
            const rowsEl = overlay.querySelector("#wfActPickerRows");
            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };

            const renderRows = (filterText = "") => {
                const f = filterText.trim().toLowerCase();
                const filtered = !f ? items : items.filter(i => (i.name || "").toLowerCase().includes(f));
                rowsEl.innerHTML = filtered.length
                    ? filtered.map((i, idx) => `
                        <tr data-pick="${idx}">
                            <td><strong>${UI.escapeHtml(i.name)}</strong></td>
                            <td><span class="table-tag">${UI.escapeHtml(i.tabla || "—")}</span></td>
                        </tr>`).join("")
                    : `<tr><td colspan="2" class="dim-picker-empty">${items.length ? "Sin resultados." : "No hay actualizaciones de tablas."}</td></tr>`;
                rowsEl.querySelectorAll("[data-pick]").forEach(tr => {
                    tr.addEventListener("click", () => cleanup(filtered[parseInt(tr.dataset.pick, 10)]));
                });
            };

            searchInput.value = "";
            searchInput.oninput = () => renderRows(searchInput.value);
            renderRows();
            overlay.classList.add("visible");
            setTimeout(() => searchInput.focus(), 50);
            overlay.querySelector("#wfActPickerClose").onclick = () => cleanup(null);
        });
    },

    openFlowManualPickerModal({ flows = [] } = {}) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("wfFlowPickerModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wfFlowPickerModal";
                document.body.appendChild(overlay);
            }
            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <h3>Seleccionar flujo manual</h3>
                        <button class="modal-close" id="wfFlowPickerClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <input type="text" id="wfFlowPickerSearch" class="dim-picker-search" placeholder="Buscar flujo...">
                        <div class="dim-picker-table-wrap">
                            <table class="dim-picker-table">
                                <thead><tr><th>Flujo</th></tr></thead>
                                <tbody id="wfFlowPickerRows"></tbody>
                            </table>
                        </div>
                        ${!flows.length ? `<p class="form-hint">Todavía no hay flujos de carga de tipo manual en este proyecto.</p>` : ""}
                    </div>
                </div>`;

            const searchInput = overlay.querySelector("#wfFlowPickerSearch");
            const rowsEl = overlay.querySelector("#wfFlowPickerRows");
            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };

            const renderRows = (filterText = "") => {
                const f = filterText.trim().toLowerCase();
                const filtered = !f ? flows : flows.filter(fl => fl.name.toLowerCase().includes(f));
                rowsEl.innerHTML = filtered.length
                    ? filtered.map((fl, idx) => `<tr data-pick="${idx}"><td><strong>${UI.escapeHtml(fl.name)}</strong></td></tr>`).join("")
                    : `<tr><td class="dim-picker-empty">${flows.length ? "Sin resultados." : "No hay flujos manuales."}</td></tr>`;
                rowsEl.querySelectorAll("[data-pick]").forEach(tr => {
                    tr.addEventListener("click", () => cleanup(filtered[parseInt(tr.dataset.pick, 10)]));
                });
            };

            searchInput.value = "";
            searchInput.oninput = () => renderRows(searchInput.value);
            renderRows();
            overlay.classList.add("visible");
            setTimeout(() => searchInput.focus(), 50);
            overlay.querySelector("#wfFlowPickerClose").onclick = () => cleanup(null);
        });
    },

    /**
     * Alta/edición de una variable de paso de Workflow: nombre técnico,
     * etiqueta y tipo, de valor único (sirve para parametrizar distintas
     * ejecuciones del mismo paso). Devuelve Promise<{name,label,type}|null>.
     */
    openWorkflowVariableModal({ current = null } = {}) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("wfVarModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wfVarModal";
                document.body.appendChild(overlay);
            }
            const v = Object.assign({ name: "", label: "", type: "STRING" }, current || {});
            overlay.innerHTML = `
                <div class="modal-box">
                    <div class="modal-header">
                        <h3>${current ? "Editar variable" : "Nueva variable del paso"}</h3>
                        <button class="modal-close" id="wfVarClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Nombre técnico</label>
                            <input type="text" id="wfVarName" placeholder="ej. mes_cierre" value="${UI.escapeHtml(v.name)}">
                        </div>
                        <div class="form-group">
                            <label>Etiqueta</label>
                            <input type="text" id="wfVarLabel" placeholder="ej. Mes de cierre" value="${UI.escapeHtml(v.label)}">
                        </div>
                        <div class="form-group">
                            <label>Tipo</label>
                            <select id="wfVarType">
                                ${UI.FIELD_TYPES.map(t => `<option value="${t}" ${t === v.type ? "selected" : ""}>${t}</option>`).join("")}
                            </select>
                            <p class="form-hint">Variable de valor único: cada ejecución del workflow puede darle un valor distinto.</p>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="wfVarCancel">Cancelar</button>
                        <button class="btn btn-primary" id="wfVarSave">Guardar</button>
                    </div>
                </div>`;
            overlay.classList.add("visible");
            const nameInput = overlay.querySelector("#wfVarName");
            setTimeout(() => nameInput.focus(), 50);
            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#wfVarClose").onclick = () => cleanup(null);
            overlay.querySelector("#wfVarCancel").onclick = () => cleanup(null);
            overlay.querySelector("#wfVarSave").onclick = () => {
                const name = overlay.querySelector("#wfVarName").value.trim();
                const label = overlay.querySelector("#wfVarLabel").value.trim();
                const type = overlay.querySelector("#wfVarType").value;
                if (!name) { UI.toast("Indica un nombre técnico para la variable.", "error"); return; }
                cleanup({ name, label: label || name, type });
            };
        });
    },

    /**
     * Asigna el valor de una variable que expone una tarea de un paso de
     * Workflow: por constante, o mapeada a una variable del propio paso;
     * con un botón para ocultarla en la pantalla de ejecución (útil cuando
     * ya se ha fijado un valor y no tiene sentido volver a pedirlo).
     * Devuelve Promise<{type,value,hidden}|"remove"|null>.
     */
    openWorkflowValueModal({ title = "Asignar valor", targetLabel = "", stepVariables = [], current = null }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("wfValueModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wfValueModal";
                document.body.appendChild(overlay);
            }

            const hasVars = stepVariables.length > 0;
            let type = current && current.type === "variable" && hasVars ? "variable" : "constante";
            let constValue = current && current.type === "constante" ? (current.value || "") : "";
            let varValue = current && current.type === "variable" ? (current.value || "") : (hasVars ? stepVariables[0].name : "");
            let hidden = !!(current && current.hidden);

            const renderBody = () => {
                const valueBlock = type === "constante"
                    ? `<div class="form-group">
                            <label>Valor constante</label>
                            <input type="text" id="wfValueValue" value="${UI.escapeHtml(constValue)}">
                       </div>`
                    : `<div class="form-group">
                            <label>Variable del paso</label>
                            <select id="wfValueVar">
                                ${stepVariables.map(v => `<option value="${UI.escapeHtml(v.name)}" ${v.name === varValue ? "selected" : ""}>${UI.escapeHtml(v.label || v.name)}</option>`).join("")}
                            </select>
                       </div>`;

                overlay.innerHTML = `
                    <div class="modal-box">
                        <div class="modal-header">
                            <div>
                                <h3>${UI.escapeHtml(title)}</h3>
                                ${targetLabel ? `<span class="modal-subtitle">${UI.escapeHtml(targetLabel)}</span>` : ""}
                            </div>
                            <button class="modal-close" id="wfValueClose">&times;</button>
                        </div>
                        <div class="modal-body">
                            ${hasVars ? `
                                <div class="form-group">
                                    <label>Origen del valor</label>
                                    <div class="segmented" id="wfValueType">
                                        <button type="button" class="segmented-btn ${type === "constante" ? "active" : ""}" data-vtype="constante">Constante</button>
                                        <button type="button" class="segmented-btn ${type === "variable" ? "active" : ""}" data-vtype="variable">Variable del paso</button>
                                    </div>
                                </div>` : `<p class="form-hint">Este paso todavía no tiene variables — el valor solo se puede fijar por constante.</p>`}
                            ${valueBlock}
                            <div class="form-group">
                                <label><input type="checkbox" id="wfValueHidden" ${hidden ? "checked" : ""}> Ocultar en la pantalla de ejecución</label>
                                <p class="form-hint">Si ya has fijado el valor aquí, no hace falta volver a pedirlo al ejecutar el paso.</p>
                            </div>
                        </div>
                        <div class="modal-footer">
                            ${current ? `<button class="btn btn-secondary" id="wfValueRemove">Quitar</button><span class="load-fn-toolbar-spacer"></span>` : ""}
                            <button class="btn btn-secondary" id="wfValueCancel">Cancelar</button>
                            <button class="btn btn-primary" id="wfValueSave">Guardar</button>
                        </div>
                    </div>`;

                const typeWrap = overlay.querySelector("#wfValueType");
                if (typeWrap) {
                    typeWrap.querySelectorAll("[data-vtype]").forEach(btn => {
                        btn.addEventListener("click", () => {
                            if (btn.dataset.vtype === type) return;
                            if (type === "constante") { const i = overlay.querySelector("#wfValueValue"); if (i) constValue = i.value; }
                            else { const s = overlay.querySelector("#wfValueVar"); if (s) varValue = s.value; }
                            const h = overlay.querySelector("#wfValueHidden"); if (h) hidden = h.checked;
                            type = btn.dataset.vtype;
                            renderBody();
                        });
                    });
                }

                const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
                overlay.querySelector("#wfValueClose").onclick = () => cleanup(null);
                overlay.querySelector("#wfValueCancel").onclick = () => cleanup(null);
                const removeBtn = overlay.querySelector("#wfValueRemove");
                if (removeBtn) removeBtn.onclick = () => cleanup("remove");
                overlay.querySelector("#wfValueSave").onclick = () => {
                    const h = overlay.querySelector("#wfValueHidden").checked;
                    if (type === "constante") {
                        const i = overlay.querySelector("#wfValueValue");
                        cleanup({ type: "constante", value: i ? i.value : "", hidden: h });
                    } else {
                        const s = overlay.querySelector("#wfValueVar");
                        cleanup({ type: "variable", value: s ? s.value : "", hidden: h });
                    }
                };

                const focusInput = overlay.querySelector("#wfValueValue");
                if (focusInput) setTimeout(() => { focusInput.focus(); focusInput.select(); }, 50);
            };

            renderBody();
            overlay.classList.add("visible");
        });
    },

    /**
     * Alta/edición de una variable de pantalla (nombre técnico, etiqueta,
     * tipo y validación). Devuelve Promise<{name,label,type,selectMode,validation}|null>.
     *
     * `dimensions` (opcional): lista de dimensiones del proyecto ([{DIMENSION_ID,DIMENSION,...}])
     * para poblar el selector de "Valores de dimensión" / "Valores de jerarquía"
     * de la validación. Si no se pasa, esos selectores aparecen vacíos.
     */
    openScreenVariableModal({ current = null, dimensions = [] } = {}) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("screenVarModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "screenVarModal";
                document.body.appendChild(overlay);
            }

            const v = Object.assign({ name: "", label: "", type: "STRING", selectMode: "unico" }, current || {});
            const validation = Object.assign({
                type: "NONE", constants: [], dimensionId: "", dimensionName: "",
                hierarchyName: "", level: 1, node: "", allowEmpty: false, showText: true, searchHelp: "LISTBOX"
            }, v.validation || {});

            const dimOptions = dimensions.map(d => `<option value="${d.DIMENSION_ID}" ${validation.dimensionId === d.DIMENSION_ID ? "selected" : ""}>${UI.escapeHtml(d.DIMENSION)}</option>`).join("");

            overlay.innerHTML = `
                <div class="modal-box">
                    <div class="modal-header">
                        <h3>${current ? "Editar variable" : "Nueva variable"}</h3>
                        <button class="modal-close" id="svClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="segmented" id="svTabs">
                            <button type="button" class="segmented-btn active" data-tab="propiedades">Propiedades</button>
                            <button type="button" class="segmented-btn" data-tab="validacion">Validación</button>
                        </div>

                        <div id="svTabProps">
                            <div class="form-group">
                                <label>Nombre técnico</label>
                                <input type="text" id="svName" placeholder="ej. fecha_desde" value="${UI.escapeHtml(v.name)}">
                            </div>
                            <div class="form-group">
                                <label>Etiqueta en pantalla</label>
                                <input type="text" id="svLabel" placeholder="ej. Fecha desde" value="${UI.escapeHtml(v.label)}">
                            </div>
                            <div class="form-group">
                                <label>Tipo</label>
                                <select id="svType">
                                    ${UI.SCREEN_VARIABLE_TYPES.map(t => `<option value="${t}" ${t === v.type ? "selected" : ""}>${t}</option>`).join("")}
                                </select>
                                <p class="form-hint">FILE: en la pantalla de ejecución del flujo se pedirá un fichero; su valor resuelto será la ruta de storage tras subirlo.</p>
                            </div>
                            <div class="form-group">
                                <label>Modo de selección</label>
                                <select id="svSelectMode">
                                    <option value="unico" ${v.selectMode === "unico" ? "selected" : ""}>Valor único</option>
                                    <option value="rango" ${v.selectMode === "rango" ? "selected" : ""}>Rango (desde – hasta)</option>
                                    <option value="multiple" ${v.selectMode === "multiple" ? "selected" : ""}>Varios valores</option>
                                    <option value="cualquiera" ${v.selectMode === "cualquiera" ? "selected" : ""}>Cualquiera (select-options)</option>
                                </select>
                                <p class="form-hint">Si eliges rango, varios valores o cualquiera, en la pantalla de ejecución aparecerá una tabla de valores (incluir/excluir + operador) al estilo select-options de SAP; el código Python recibirá esa tabla (JSON) en lugar de un valor único.</p>
                            </div>
                        </div>

                        <div id="svTabValid" style="display:none">
                            <div class="form-group">
                                <label>Validación</label>
                                <label><input type="radio" name="svValidType" value="NONE" ${validation.type === "NONE" ? "checked" : ""}> Ninguna</label>
                                <label><input type="radio" name="svValidType" value="CONST" ${validation.type === "CONST" ? "checked" : ""}> Constante</label>
                                <label><input type="radio" name="svValidType" value="DIM" ${validation.type === "DIM" ? "checked" : ""}> Valores de dimensión</label>
                                <label><input type="radio" name="svValidType" value="HIER" ${validation.type === "HIER" ? "checked" : ""}> Valores de jerarquía</label>
                            </div>

                            <div id="svValidConst" style="display:${validation.type === "CONST" ? "block" : "none"}">
                                <table class="const-table">
                                    <thead><tr><th>ID</th><th>Descripción</th><th></th></tr></thead>
                                    <tbody id="svConstRows"></tbody>
                                </table>
                                <button class="btn btn-secondary btn-sm" id="svAddConst">+ Añadir valor</button>
                            </div>

                            <div id="svValidDim" style="display:${validation.type === "DIM" ? "block" : "none"}">
                                <div class="form-group">
                                    <label>Dimensión</label>
                                    <select id="svDimSelect"><option value="">Selecciona...</option>${dimOptions}</select>
                                </div>
                            </div>

                            <div id="svValidHier" style="display:${validation.type === "HIER" ? "block" : "none"}">
                                <div class="form-group">
                                    <label>Dimensión</label>
                                    <select id="svHierDimSelect"><option value="">Selecciona...</option>${dimOptions}</select>
                                </div>
                                <div class="form-group">
                                    <label>Jerarquía</label>
                                    <select id="svHierSelect"><option value="">Selecciona una dimensión primero</option></select>
                                </div>
                                <div class="form-group">
                                    <label>Nivel</label>
                                    <select id="svHierLevelSelect"><option value="">—</option></select>
                                </div>
                                <div class="form-group">
                                    <label>Valor del nodo (se traerán los valores por debajo de este nodo)</label>
                                    <input type="text" id="svHierNode" value="${UI.escapeHtml(validation.node || "")}">
                                </div>
                            </div>

                            <div id="svValidCommon" style="display:${validation.type === "NONE" ? "none" : "block"}">
                                <div class="form-group">
                                    <label><input type="checkbox" id="svAllowEmpty" ${validation.allowEmpty ? "checked" : ""}> Permite valor vacío</label>
                                </div>
                                <div class="form-group">
                                    <label><input type="checkbox" id="svShowText" ${validation.showText ? "checked" : ""}> Mostrar texto descriptivo junto al valor</label>
                                </div>
                                <div class="form-group">
                                    <label>Ayuda de búsqueda</label>
                                    <select id="svSearchHelp">
                                        <option value="LISTBOX" ${validation.searchHelp === "LISTBOX" ? "selected" : ""}>Listbox (desplegable)</option>
                                        <option value="SEARCH" ${validation.searchHelp === "SEARCH" ? "selected" : ""}>Buscador (estilo SAP, con caja de búsqueda)</option>
                                        <option value="CHECKBOX" ${validation.searchHelp === "CHECKBOX" ? "selected" : ""}>Checkbox (solo si hay exactamente 2 valores posibles)</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="svCancel">Cancelar</button>
                        <button class="btn btn-primary" id="svSave">Guardar</button>
                    </div>
                </div>`;

            overlay.classList.add("visible");
            const nameInput = overlay.querySelector("#svName");
            setTimeout(() => { nameInput.focus(); }, 50);

            // -- pestañas: Propiedades | Validación --
            const tabProps = overlay.querySelector("#svTabProps");
            const tabValid = overlay.querySelector("#svTabValid");
            overlay.querySelectorAll("#svTabs [data-tab]").forEach(btn => {
                btn.addEventListener("click", () => {
                    overlay.querySelectorAll("#svTabs [data-tab]").forEach(b => b.classList.remove("active"));
                    btn.classList.add("active");
                    const isProps = btn.dataset.tab === "propiedades";
                    tabProps.style.display = isProps ? "flex" : "none";
                    tabValid.style.display = isProps ? "none" : "flex";
                });
            });

            // -- validación: alternar bloques según el tipo elegido --
            overlay.querySelectorAll('input[name="svValidType"]').forEach(radio => {
                radio.addEventListener("change", () => {
                    overlay.querySelector("#svValidConst").style.display = radio.value === "CONST" && radio.checked ? "block" : "none";
                    overlay.querySelector("#svValidDim").style.display = radio.value === "DIM" && radio.checked ? "block" : "none";
                    overlay.querySelector("#svValidHier").style.display = radio.value === "HIER" && radio.checked ? "block" : "none";
                    overlay.querySelector("#svValidCommon").style.display = radio.value === "NONE" && radio.checked ? "none" : "block";
                });
            });

            // -- constante: tabla ID/Descripción editable --
            let constants = (validation.constants || []).slice();
            const renderConstRows = () => {
                const tbody = overlay.querySelector("#svConstRows");
                tbody.innerHTML = constants.map((c, i) => `
                    <tr>
                        <td><input type="text" class="sv-const-id" data-i="${i}" value="${UI.escapeHtml(c.id)}"></td>
                        <td><input type="text" class="sv-const-desc" data-i="${i}" value="${UI.escapeHtml(c.desc)}"></td>
                        <td><button type="button" class="field-remove" data-rm="${i}">✕</button></td>
                    </tr>`).join("");
                tbody.querySelectorAll(".sv-const-id").forEach(inp => inp.addEventListener("input", e => constants[e.target.dataset.i].id = e.target.value));
                tbody.querySelectorAll(".sv-const-desc").forEach(inp => inp.addEventListener("input", e => constants[e.target.dataset.i].desc = e.target.value));
                tbody.querySelectorAll("[data-rm]").forEach(btn => btn.addEventListener("click", () => { constants.splice(parseInt(btn.dataset.rm, 10), 1); renderConstRows(); }));
            };
            renderConstRows();
            overlay.querySelector("#svAddConst").addEventListener("click", () => { constants.push({ id: "", desc: "" }); renderConstRows(); });

            // -- jerarquía: dependencias dimensión -> jerarquía -> nivel --
            const hierDimSel = overlay.querySelector("#svHierDimSelect");
            const hierSel = overlay.querySelector("#svHierSelect");
            const hierLevelSel = overlay.querySelector("#svHierLevelSelect");
            let hierLevelsCache = [];

            const safeParseJson = (json, fallback) => {
                try { const p = JSON.parse(json); return p == null ? fallback : p; } catch (e) { return fallback; }
            };

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
                const rows = safeParseJson(hierSel.dataset.rows || "[]", []);
                const row = rows.find(r => r.JERARQUIA === hierName);
                hierLevelsCache = row ? safeParseJson(row.NIVELES_JSON, []) : [];
                hierLevelSel.innerHTML = hierLevelsCache.map((colId, i) => `<option value="${i + 1}" ${preselectLevel == i + 1 ? "selected" : ""}>Nivel ${i + 1}: ${UI.escapeHtml(colId)}</option>`).join("") || `<option value="">—</option>`;
            };

            if (validation.type === "HIER" && validation.dimensionId) {
                loadHierarchiesFor(validation.dimensionId, validation.hierarchyName, validation.level);
            }
            hierDimSel.addEventListener("change", () => loadHierarchiesFor(hierDimSel.value));
            hierSel.addEventListener("change", () => fillLevels(hierSel.value));

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#svClose").onclick = () => cleanup(null);
            overlay.querySelector("#svCancel").onclick = () => cleanup(null);
            overlay.querySelector("#svSave").onclick = () => {
                const name = overlay.querySelector("#svName").value.trim();
                const label = overlay.querySelector("#svLabel").value.trim();
                const type = overlay.querySelector("#svType").value;
                const selectMode = overlay.querySelector("#svSelectMode").value;
                if (!name) { UI.toast("Indica un nombre técnico para la variable.", "error"); return; }

                const validType = overlay.querySelector('input[name="svValidType"]:checked').value;
                const newValidation = {
                    type: validType,
                    constants: validType === "CONST" ? constants.filter(c => c.id) : [],
                    dimensionId: validType === "DIM" ? overlay.querySelector("#svDimSelect").value : (validType === "HIER" ? hierDimSel.value : ""),
                    dimensionName: validType === "DIM"
                        ? (dimensions.find(d => d.DIMENSION_ID === overlay.querySelector("#svDimSelect").value) || {}).DIMENSION || ""
                        : (validType === "HIER" ? (dimensions.find(d => d.DIMENSION_ID === hierDimSel.value) || {}).DIMENSION || "" : ""),
                    hierarchyName: validType === "HIER" ? hierSel.value : "",
                    level: validType === "HIER" ? parseInt(hierLevelSel.value || "1", 10) : 1,
                    node: validType === "HIER" ? overlay.querySelector("#svHierNode").value.trim() : "",
                    allowEmpty: overlay.querySelector("#svAllowEmpty").checked,
                    showText: overlay.querySelector("#svShowText").checked,
                    searchHelp: overlay.querySelector("#svSearchHelp").value
                };

                cleanup({ name, label: label || name, type, selectMode, validation: newValidation });
            };
        });
    },

    /**
     * Editor de texto explicativo con formato mínimo (negrita/cursiva vía
     * marcadores tipo markdown). Devuelve Promise<string|null>.
     */
    openScreenTextModal({ current = "" } = {}) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("screenTextModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "screenTextModal";
                document.body.appendChild(overlay);
            }

            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <h3>Texto explicativo</h3>
                        <button class="modal-close" id="stClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="load-fn-toolbar" style="border:none; padding:0 0 8px;">
                            <button type="button" class="btn btn-secondary btn-sm" id="stBold"><strong>N</strong></button>
                            <button type="button" class="btn btn-secondary btn-sm" id="stItalic"><em>K</em></button>
                            <span class="form-hint">Selecciona texto y pulsa N / K para aplicar negrita o cursiva.</span>
                        </div>
                        <textarea id="stArea" rows="6" style="width:100%; font-family:inherit; font-size:var(--fs-base); padding:8px 10px; border:1px solid var(--border-default); border-radius:var(--radius-sm);">${UI.escapeHtml(current)}</textarea>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="stCancel">Cancelar</button>
                        <button class="btn btn-primary" id="stSave">Guardar</button>
                    </div>
                </div>`;

            overlay.classList.add("visible");
            const area = overlay.querySelector("#stArea");
            setTimeout(() => area.focus(), 50);

            const wrapSelection = (marker) => {
                const start = area.selectionStart, end = area.selectionEnd;
                const text = area.value;
                const selected = text.slice(start, end) || "texto";
                area.value = text.slice(0, start) + marker + selected + marker + text.slice(end);
                area.focus();
                area.selectionStart = start + marker.length;
                area.selectionEnd = start + marker.length + selected.length;
            };
            overlay.querySelector("#stBold").onclick = () => wrapSelection("**");
            overlay.querySelector("#stItalic").onclick = () => wrapSelection("_");

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#stClose").onclick = () => cleanup(null);
            overlay.querySelector("#stCancel").onclick = () => cleanup(null);
            overlay.querySelector("#stSave").onclick = () => cleanup(area.value);
        });
    },

    /** Convierte marcadores **negrita** / _cursiva_ a HTML simple, escapando el resto. */
    renderFormattedText(text) {
        let html = UI.escapeHtml(text || "");
        html = html.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
        html = html.replace(/_(.+?)_/g, "<em>$1</em>");
        return html.replace(/\n/g, "<br>");
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
