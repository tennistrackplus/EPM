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
                                <div class="dim-picker" id="cuboFormDims"></div>
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
            const dimsContainer = overlay.querySelector("#cuboFormDims");
            const nameInput = overlay.querySelector("#cuboFormName");

            dimsContainer.innerHTML = dimensionsList.length
                ? dimensionsList.map(d => `
                    <label class="dim-picker-item">
                        <input type="checkbox" value="${d.DIMENSION_ID}" ${selectedDimensionIds.includes(d.DIMENSION_ID) ? "checked" : ""}>
                        <span>${UI.escapeHtml(d.DIMENSION)}</span>
                        <span class="table-tag">${UI.escapeHtml(d.TABLA)}</span>
                    </label>`).join("")
                : `<p class="form-hint">Todavía no hay dimensiones en este proyecto. Crea alguna primero desde el menú "Dimensiones".</p>`;

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
                const dimensionIds = Array.from(dimsContainer.querySelectorAll("input[type=checkbox]:checked")).map(c => c.value);
                const measureRows = UI._readFieldRows(rowsContainer);

                if (!dimensionIds.length && !measureRows.length) {
                    UI.toast("Selecciona al menos una dimensión o añade una medida.", "error");
                    return;
                }

                cleanup({
                    name: nameVal,
                    description: overlay.querySelector("#cuboFormDesc").value.trim(),
                    dimensionIds,
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
