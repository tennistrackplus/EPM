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

    FIELD_TYPES: ["STRING", "INTEGER", "FLOAT64", "NUMERIC", "BOOLEAN", "DATE", "DATETIME", "TIMESTAMP"],

    /**
     * Modal genérico de "entidad con campos", usado tanto por Dimensiones como
     * por Cubos: nombre + descripción + diseñador de columnas.
     * Devuelve una Promise<{name, description, fields}|null>.
     */
    openEntityFormModal({ title, nameLabel = "Nombre", namePlaceholder = "", name = "", description = "", fields = [], nameEditable = true }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("entityFormModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "entityFormModal";
                overlay.innerHTML = `
                    <div class="modal-box modal-wide">
                        <div class="modal-header">
                            <h3 id="entityFormTitle"></h3>
                            <button class="modal-close" id="entityFormClose">&times;</button>
                        </div>
                        <div class="modal-body">
                            <div class="form-row">
                                <div class="form-group">
                                    <label id="entityFormNameLabel"></label>
                                    <input type="text" id="entityFormName">
                                </div>
                            </div>
                            <div class="form-group">
                                <label>Descripción</label>
                                <textarea id="entityFormDesc" rows="2"></textarea>
                            </div>
                            <div class="form-group">
                                <label>Campos</label>
                                <div class="fields-builder">
                                    <div class="fields-builder-header">
                                        <span>Nombre del campo</span><span>Tipo</span><span>Clave</span><span></span>
                                    </div>
                                    <div class="fields-builder-rows" id="entityFormRows"></div>
                                    <div class="fields-builder-footer">
                                        <button class="btn btn-secondary btn-sm" id="entityFormAddField">+ Añadir campo</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button class="btn btn-secondary" id="entityFormCancel">Cancelar</button>
                            <button class="btn btn-primary" id="entityFormSave">Guardar</button>
                        </div>
                    </div>`;
                document.body.appendChild(overlay);
            }

            const rowsContainer = overlay.querySelector("#entityFormRows");

            const addRow = (f = { name: "", type: "STRING", key: false }) => {
                const row = document.createElement("div");
                row.className = "field-row";
                row.innerHTML = `
                    <input type="text" class="f-name" placeholder="ej. codigo_cliente" value="${UI.escapeHtml(f.name)}">
                    <select class="f-type">
                        ${UI.FIELD_TYPES.map(t => `<option value="${t}" ${t === f.type ? "selected" : ""}>${t}</option>`).join("")}
                    </select>
                    <span class="field-key"><input type="checkbox" class="f-key" ${f.key ? "checked" : ""}></span>
                    <button type="button" class="field-remove" title="Eliminar campo">✕</button>`;
                row.querySelector(".field-remove").addEventListener("click", () => row.remove());
                rowsContainer.appendChild(row);
            };

            rowsContainer.innerHTML = "";
            (fields.length ? fields : [{ name: "", type: "STRING", key: true }]).forEach(f => addRow(f));

            overlay.querySelector("#entityFormTitle").textContent = title;
            overlay.querySelector("#entityFormNameLabel").textContent = nameLabel;
            const nameInput = overlay.querySelector("#entityFormName");
            nameInput.value = name;
            nameInput.placeholder = namePlaceholder;
            nameInput.disabled = !nameEditable;
            overlay.querySelector("#entityFormDesc").value = description;
            overlay.querySelector("#entityFormAddField").onclick = () => addRow();

            overlay.classList.add("visible");
            setTimeout(() => nameInput.focus(), 50);

            const cleanup = (result) => {
                overlay.classList.remove("visible");
                resolve(result);
            };

            overlay.querySelector("#entityFormCancel").onclick = () => cleanup(null);
            overlay.querySelector("#entityFormClose").onclick = () => cleanup(null);
            overlay.querySelector("#entityFormSave").onclick = () => {
                const nameVal = nameInput.value.trim();
                if (!nameVal) {
                    UI.toast("Indica un nombre.", "error");
                    return;
                }
                const fieldRows = Array.from(rowsContainer.querySelectorAll(".field-row")).map(r => ({
                    name: r.querySelector(".f-name").value.trim(),
                    type: r.querySelector(".f-type").value,
                    key: r.querySelector(".f-key").checked
                })).filter(f => f.name);

                if (!fieldRows.length) {
                    UI.toast("Añade al menos un campo.", "error");
                    return;
                }

                cleanup({
                    name: nameVal,
                    description: overlay.querySelector("#entityFormDesc").value.trim(),
                    fields: fieldRows
                });
            };
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
