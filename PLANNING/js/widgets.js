/**
 * Módulo de Widgets. Un widget es una pieza visual reutilizable de tres
 * tipos: FILTRO, TABLA (editor tipo Excel) o GRAFICO. En esta primera
 * entrega el módulo cubre solo el alta/edición de datos básicos (nombre,
 * tipo, descripción) y el borrado — igual que el resto de módulos del
 * menú de Administración. El editor específico de cada tipo (celdas y
 * formato de la tabla, definición del informe sobre el modelo semántico,
 * configuración del gráfico, y la exportación a widget HTML embebible)
 * se añadirá en próximas iteraciones sobre CONFIG_JSON.
 */
const Widgets = {
    TABLE: "WIDGETS",
    NAME_COL: "WIDGET",
    ID_COL: "WIDGET_ID",
    list: [],

    TYPES: [
        { value: "FILTRO",  label: "Filtro",   icon: "⌕" },
        { value: "TABLA",   label: "Tabla",    icon: "▦" },
        { value: "GRAFICO", label: "Gráfico",  icon: "▲" }
    ],

    async render(container, project) {
        this.container = container;
        this.project = project;

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <h3>Widgets</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET}</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewWidget">
                    + Nuevo widget
                </button>
            </div>
            <div id="widgetsListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewWidget").addEventListener("click", () => this.openForm());
        await this.loadList();
    },

    typeInfo(tipo) {
        return this.TYPES.find(t => t.value === tipo) || { label: tipo, icon: "•" };
    },

    async loadList() {
        const wrap = document.getElementById("widgetsListWrap");
        try {
            const sql = `SELECT ${this.ID_COL}, ${this.NAME_COL}, TIPO, DESCRIPCION, CUBO_ID
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                         ORDER BY ${this.NAME_COL}`;
            this.list = await Provider.runQuery(sql);

            if (!this.list.length) {
                wrap.innerHTML = `<div class="module-empty">Todavía no hay widgets en este proyecto. Crea el primero con "Nuevo widget".</div>`;
                return;
            }

            wrap.innerHTML = `
                <div class="data-list">
                    <table>
                        <thead><tr><th>Widget</th><th>Tipo</th><th>Descripción</th><th></th></tr></thead>
                        <tbody>
                            ${this.list.map(w => {
                                const info = this.typeInfo(w.TIPO);
                                return `
                                <tr>
                                    <td><strong>${UI.escapeHtml(w[this.NAME_COL])}</strong></td>
                                    <td><span class="widget-type-badge widget-type-${w.TIPO.toLowerCase()}">${info.icon} ${info.label}</span></td>
                                    <td><span class="col-type">${UI.escapeHtml(w.DESCRIPCION || "—")}</span></td>
                                    <td>
                                        <div class="row-actions">
                                            <button data-edit="${w[this.ID_COL]}" title="Editar">✎</button>
                                            <button data-del="${w[this.ID_COL]}" class="danger" title="Eliminar">🗑</button>
                                        </div>
                                    </td>
                                </tr>`;
                            }).join("")}
                        </tbody>
                    </table>
                </div>`;

            wrap.querySelectorAll("[data-edit]").forEach(btn =>
                btn.addEventListener("click", () => this.openForm(btn.dataset.edit)));
            wrap.querySelectorAll("[data-del]").forEach(btn =>
                btn.addEventListener("click", () => this.remove(btn.dataset.del)));
        } catch (err) {
            wrap.innerHTML = `<div class="module-empty">Error al cargar widgets: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    // ------------------------------------------------------------
    // Popup "Nuevo widget" / datos básicos: nombre + tipo (Filtro /
    // Tabla / Gráfico) + descripción. Mismo patrón que Flows.openBasicsModal.
    // ------------------------------------------------------------
    openBasicsModal(initial, isNew) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("widgetBasicsModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "widgetBasicsModal";
                document.body.appendChild(overlay);
            }

            overlay.innerHTML = `
                <div class="modal-box">
                    <div class="modal-header">
                        <h3>${isNew ? "Nuevo widget" : "Editar widget"}</h3>
                        <button class="modal-close" id="widgetBasicsClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="flow-type-toggle-wrap">
                            <div class="flow-type-toggle" id="widgetBasicsType">
                                ${this.TYPES.map(t => `
                                    <button type="button" class="flow-type-toggle-btn ${initial.type === t.value ? "active" : ""}" data-wtype="${t.value}">${t.icon} ${t.label}</button>
                                `).join("")}
                            </div>
                        </div>
                        <div class="form-group">
                            <label>Nombre del widget</label>
                            <input type="text" id="widgetBasicsName" placeholder="Ej. Ventas por región" value="${UI.escapeHtml(initial.name || "")}">
                        </div>
                        <div class="form-group">
                            <label>Descripción</label>
                            <textarea id="widgetBasicsDesc" rows="2" placeholder="Describe brevemente qué muestra este widget...">${UI.escapeHtml(initial.description || "")}</textarea>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="widgetBasicsCancel">Cancelar</button>
                        <button class="btn btn-primary" id="widgetBasicsNext">${isNew ? "Continuar" : "Guardar"}</button>
                    </div>
                </div>`;

            let type = initial.type || "TABLA";
            overlay.querySelectorAll("#widgetBasicsType [data-wtype]").forEach(btn => {
                btn.addEventListener("click", () => {
                    type = btn.dataset.wtype;
                    overlay.querySelectorAll("#widgetBasicsType [data-wtype]").forEach(b => b.classList.toggle("active", b === btn));
                });
            });

            overlay.classList.add("visible");
            const nameInput = overlay.querySelector("#widgetBasicsName");
            setTimeout(() => { nameInput.focus(); nameInput.select(); }, 50);

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#widgetBasicsClose").onclick = () => cleanup(null);
            overlay.querySelector("#widgetBasicsCancel").onclick = () => cleanup(null);
            overlay.querySelector("#widgetBasicsNext").onclick = () => {
                const name = nameInput.value.trim();
                const description = overlay.querySelector("#widgetBasicsDesc").value.trim();
                if (!name) { UI.toast("Indica un nombre para el widget.", "error"); return; }
                cleanup({ name, type, description });
            };
            nameInput.onkeydown = (e) => { if (e.key === "Enter") overlay.querySelector("#widgetBasicsNext").click(); };
        });
    },

    async openForm(editId = null) {
        const editing = editId ? this.list.find(w => w[this.ID_COL] === editId) : null;
        const initial = editing
            ? { name: editing[this.NAME_COL], type: editing.TIPO, description: editing.DESCRIPCION || "" }
            : { name: "", type: "TABLA", description: "" };

        const result = await this.openBasicsModal(initial, !editing);
        if (!result) return;

        const savedId = await this.save(editing, result);
        if (!savedId) return;

        if (result.type === "TABLA" && typeof WidgetTableEditor !== "undefined") {
            const row = this.list.find(w => w[this.ID_COL] === savedId) ||
                { [this.ID_COL]: savedId, [this.NAME_COL]: result.name, DESCRIPCION: result.description };
            await WidgetTableEditor.open(row);
            return;
        }

        // El editor específico de Filtro / Gráfico llega en la siguiente
        // iteración. De momento avisamos de que el widget ya existe.
        const info = this.typeInfo(result.type);
        UI.toast(`El editor de contenido para widgets de tipo "${info.label}" se añadirá en el siguiente paso.`, "info");
    },

    // Devuelve el WIDGET_ID guardado (nuevo o existente), o null si falló.
    async save(editing, { name, type, description }) {
        const widgetId = editing ? editing[this.ID_COL] : Provider.newId();
        try {
            if (editing) {
                const sql = `UPDATE ${Provider.qualifyControl(this.TABLE)}
                    SET ${this.NAME_COL} = '${Provider.esc(name)}',
                        TIPO = '${Provider.esc(type)}',
                        DESCRIPCION = '${Provider.esc(description)}',
                        FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                    WHERE ${this.ID_COL} = '${Provider.esc(widgetId)}'`;
                await Provider.runQuery(sql);
                UI.toast(`Widget "${name}" actualizado.`, "success");
            } else {
                const sql = `INSERT INTO ${Provider.qualifyControl(this.TABLE)}
                    (${this.ID_COL}, PROYECTO_ID, ${this.NAME_COL}, TIPO, DESCRIPCION, CONFIG_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                    VALUES ('${Provider.esc(widgetId)}', '${Provider.esc(this.project.PROYECTO_ID)}', '${Provider.esc(name)}',
                            '${Provider.esc(type)}', '${Provider.esc(description)}', '{}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`;
                await Provider.runQuery(sql);
                UI.toast(`Widget "${name}" creado.`, "success");
            }
            await this.loadList();
            return widgetId;
        } catch (err) {
            UI.toast("Error al guardar el widget: " + err.message, "error");
            return null;
        }
    },

    async remove(id) {
        const widget = this.list.find(w => w[this.ID_COL] === id);
        if (!widget) return;

        const ok = await UI.confirm(
            "Eliminar widget",
            `Se eliminará el widget <strong>${UI.escapeHtml(widget[this.NAME_COL])}</strong>. Esta acción no se puede deshacer.`
        );
        if (!ok) return;

        try {
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
            UI.toast(`Widget "${widget[this.NAME_COL]}" eliminado.`, "success");
            await this.loadList();
        } catch (err) {
            UI.toast("Error al eliminar el widget: " + err.message, "error");
        }
    }
};
