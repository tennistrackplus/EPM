/**
 * Módulo de Dimensiones. La clave principal de la tabla física es
 * siempre el propio nombre de la dimensión (ej. dimensión CUENTA ->
 * columna CUENTA, clave). El resto de campos son "atributos", que
 * también se pueden marcar como clave para soportar claves compuestas
 * (ej. Clase de coste + Sociedad).
 */
const Dimensions = {
    TABLE: "DIMENSIONES",
    NAME_COL: "DIMENSION",
    ID_COL: "DIMENSION_ID",
    list: [],

    async render(container, project) {
        this.container = container;
        this.project = project;

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <h3>Dimensiones</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET}</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewDimension">
                    + Nueva dimensión
                </button>
            </div>
            <div id="dimensionsListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewDimension").addEventListener("click", () => this.openForm());
        await this.loadList();
    },

    async loadList() {
        const wrap = document.getElementById("dimensionsListWrap");
        try {
            const sql = `SELECT ${this.ID_COL}, ${this.NAME_COL}, DESCRIPCION, TABLA, CAMPOS_JSON
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                         ORDER BY ${this.NAME_COL}`;
            this.list = await Provider.runQuery(sql);

            if (!this.list.length) {
                wrap.innerHTML = `<div class="module-empty">Todavía no hay dimensiones en este proyecto. Crea la primera con "Nueva dimensión".</div>`;
                return;
            }

            wrap.innerHTML = `
                <div class="data-list">
                    <table>
                        <thead><tr><th>Dimensión</th><th>Descripción</th><th>Tabla</th><th></th></tr></thead>
                        <tbody>
                            ${this.list.map(d => `
                                <tr>
                                    <td><strong>${UI.escapeHtml(d[this.NAME_COL])}</strong></td>
                                    <td>${UI.escapeHtml(d.DESCRIPCION || "—")}</td>
                                    <td><span class="table-tag">${UI.escapeHtml(d.TABLA)}</span></td>
                                    <td>
                                        <div class="row-actions">
                                            <button data-values="${d[this.ID_COL]}" title="Actualizar valores">▤</button>
                                            <button data-hier="${d[this.ID_COL]}" title="Jerarquías">⛓</button>
                                            <button data-edit="${d[this.ID_COL]}" title="Editar">✎</button>
                                            <button data-del="${d[this.ID_COL]}" class="danger" title="Eliminar">🗑</button>
                                        </div>
                                    </td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>`;

            wrap.querySelectorAll("[data-edit]").forEach(btn =>
                btn.addEventListener("click", () => this.openForm(btn.dataset.edit)));
            wrap.querySelectorAll("[data-del]").forEach(btn =>
                btn.addEventListener("click", () => this.remove(btn.dataset.del)));
            wrap.querySelectorAll("[data-values]").forEach(btn =>
                btn.addEventListener("click", () => this.openDataEditor(btn.dataset.values)));
            wrap.querySelectorAll("[data-hier]").forEach(btn =>
                btn.addEventListener("click", () => this.openHierarchyManager(btn.dataset.hier)));
        } catch (err) {
            wrap.innerHTML = `<div class="module-empty">Error al cargar dimensiones: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    parseFields(dim) {
        try { return JSON.parse(dim.CAMPOS_JSON || "[]"); } catch (e) { return []; }
    },

    async openForm(editId = null) {
        const editing = editId ? this.list.find(d => d[this.ID_COL] === editId) : null;
        let allFields = editing ? this.parseFields(editing) : [];
        const keyField = allFields.find(f => f.__isPrimaryName) || allFields[0] || { type: "STRING" };
        const attributes = allFields.filter(f => f !== keyField);

        const result = await UI.openDimensionFormModal({
            title: editing ? `Editar dimensión: ${editing[this.NAME_COL]}` : "Nueva dimensión",
            name: editing ? editing[this.NAME_COL] : "",
            description: editing ? (editing.DESCRIPCION || "") : "",
            keyType: keyField.type || "STRING",
            attributes,
            nameEditable: !editing
        });

        if (!result) return;
        await this.save(editing, result);
    },

    async save(editing, { name, description, keyType, attributes }) {
        const ident = Provider.toIdentifier(name);
        if (!ident) {
            UI.toast("El nombre de la dimensión no es válido.", "error");
            return;
        }
        const tableName = `${DracoConfig.prefix}${ident}`;
        const fullTable = Provider.qualify(this.project.DATASET, tableName);

        // La clave principal siempre es el propio nombre de la dimensión.
        const keyField = { name: ident, type: keyType, key: true, __isPrimaryName: true };
        const allFields = [keyField, ...attributes];

        const colDefs = allFields.map(f => {
            const colIdent = Provider.toIdentifier(f.name);
            return `${colIdent} ${Provider.mapFieldType(f.type)}`;
        }).join(", ");

        try {
            await Provider.runQuery(`CREATE OR REPLACE TABLE ${fullTable} (${colDefs})`);

            const camposJson = JSON.stringify(allFields).replace(/\\/g, "\\\\").replace(/'/g, "\\'");

            if (editing) {
                const sql = `UPDATE ${Provider.qualifyControl(this.TABLE)}
                    SET DESCRIPCION = '${Provider.esc(description)}',
                        CAMPOS_JSON = '${camposJson}',
                        FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                    WHERE ${this.ID_COL} = '${Provider.esc(editing[this.ID_COL])}'`;
                await Provider.runQuery(sql);
                UI.toast(`Dimensión "${name}" actualizada.`, "success");
            } else {
                const id = Provider.newId();
                const sql = `INSERT INTO ${Provider.qualifyControl(this.TABLE)}
                    (${this.ID_COL}, PROYECTO_ID, ${this.NAME_COL}, DESCRIPCION, TABLA, CAMPOS_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                    VALUES ('${Provider.esc(id)}', '${Provider.esc(this.project.PROYECTO_ID)}', '${Provider.esc(name)}', '${Provider.esc(description)}',
                            '${Provider.esc(tableName)}', '${camposJson}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`;
                await Provider.runQuery(sql);
                UI.toast(`Dimensión "${name}" creada.`, "success");
            }

            await this.loadList();
            Draco.renderProgress();
        } catch (err) {
            UI.toast("Error al guardar la dimensión: " + err.message, "error");
        }
    },

    async remove(id) {
        const dim = this.list.find(d => d[this.ID_COL] === id);
        if (!dim) return;

        const ok = await UI.confirm(
            "Eliminar dimensión",
            `Se eliminará la dimensión <strong>${UI.escapeHtml(dim[this.NAME_COL])}</strong> y su tabla <strong>${dim.TABLA}</strong> con todos sus datos.`
        );
        if (!ok) return;

        try {
            await Provider.runQuery(`DROP TABLE IF EXISTS ${Provider.qualify(this.project.DATASET, dim.TABLA)}`);
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
            const delHier = `DELETE FROM ${Provider.qualifyControl("JERARQUIAS")} WHERE DIMENSION_ID = '${Provider.esc(id)}'`;
            await Provider.runQuery(delHier);
            UI.toast(`Dimensión "${dim[this.NAME_COL]}" eliminada.`, "success");
            await this.loadList();
            Draco.renderProgress();
        } catch (err) {
            UI.toast("Error al eliminar la dimensión: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Navegación a sub-pantallas (Valores / Jerarquías)
    // ------------------------------------------------------------
    openDataEditor(id) {
        const dim = this.list.find(d => d[this.ID_COL] === id);
        if (!dim) return;
        DimensionData.render(this.container, this.project, dim, () => this.render(this.container, this.project));
    },

    openHierarchyManager(id) {
        const dim = this.list.find(d => d[this.ID_COL] === id);
        if (!dim) return;
        Hierarchies.render(this.container, this.project, dim, () => this.render(this.container, this.project));
    }
};
