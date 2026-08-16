/**
 * Módulo de Cubos. Misma mecánica que Dimensiones, a través de Provider
 * (funciona igual sobre BigQuery o Snowflake).
 */
const Cubes = {
    TABLE: "CUBOS",
    NAME_COL: "CUBOS",
    ID_COL: "CUBO_ID",
    list: [],

    async render(container, project) {
        this.container = container;
        this.project = project;

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <h3>Cubos</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET}</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewCubo">
                    + Nuevo cubo
                </button>
            </div>
            <div id="cubesListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewCubo").addEventListener("click", () => this.openForm());
        await this.loadList();
    },

    async loadList() {
        const wrap = document.getElementById("cubesListWrap");
        try {
            const sql = `SELECT ${this.ID_COL}, ${this.NAME_COL}, DESCRIPCION, TABLA, CAMPOS_JSON
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                         ORDER BY ${this.NAME_COL}`;
            this.list = await Provider.runQuery(sql);

            if (!this.list.length) {
                wrap.innerHTML = `<div class="module-empty">Todavía no hay cubos en este proyecto. Crea el primero con "Nuevo cubo".</div>`;
                return;
            }

            wrap.innerHTML = `
                <div class="data-list">
                    <table>
                        <thead><tr><th>Cubo</th><th>Descripción</th><th>Tabla</th><th></th></tr></thead>
                        <tbody>
                            ${this.list.map(c => `
                                <tr>
                                    <td><strong>${UI.escapeHtml(c[this.NAME_COL])}</strong></td>
                                    <td>${UI.escapeHtml(c.DESCRIPCION || "—")}</td>
                                    <td><span class="table-tag">${UI.escapeHtml(c.TABLA)}</span></td>
                                    <td>
                                        <div class="row-actions">
                                            <button data-edit="${c[this.ID_COL]}" title="Editar">✎</button>
                                            <button data-del="${c[this.ID_COL]}" class="danger" title="Eliminar">🗑</button>
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
        } catch (err) {
            wrap.innerHTML = `<div class="module-empty">Error al cargar cubos: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    async openForm(editId = null) {
        const editing = editId ? this.list.find(c => c[this.ID_COL] === editId) : null;
        let fields = [];
        if (editing && editing.CAMPOS_JSON) {
            try { fields = JSON.parse(editing.CAMPOS_JSON); } catch (e) { fields = []; }
        }

        const result = await UI.openEntityFormModal({
            title: editing ? `Editar cubo: ${editing[this.NAME_COL]}` : "Nuevo cubo",
            nameLabel: "Nombre del cubo",
            namePlaceholder: "Ej. Ventas",
            name: editing ? editing[this.NAME_COL] : "",
            description: editing ? (editing.DESCRIPCION || "") : "",
            fields,
            nameEditable: !editing
        });

        if (!result) return;
        await this.save(editing, result);
    },

    async save(editing, { name, description, fields }) {
        const ident = Provider.toIdentifier(name);
        if (!ident) {
            UI.toast("El nombre del cubo no es válido.", "error");
            return;
        }
        const tableName = `${DracoConfig.prefix}${ident}`;
        const fullTable = Provider.qualify(this.project.DATASET, tableName);

        const colDefs = fields.map(f => {
            const colIdent = Provider.toIdentifier(f.name);
            return `${colIdent} ${Provider.mapFieldType(f.type)}`;
        }).join(", ");

        try {
            await Provider.runQuery(`CREATE OR REPLACE TABLE ${fullTable} (${colDefs})`);

            const camposJson = JSON.stringify(fields).replace(/\\/g, "\\\\").replace(/'/g, "\\'");

            if (editing) {
                const sql = `UPDATE ${Provider.qualifyControl(this.TABLE)}
                    SET DESCRIPCION = '${Provider.esc(description)}',
                        CAMPOS_JSON = '${camposJson}',
                        FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                    WHERE ${this.ID_COL} = '${Provider.esc(editing[this.ID_COL])}'`;
                await Provider.runQuery(sql);
                UI.toast(`Cubo "${name}" actualizado.`, "success");
            } else {
                const id = Provider.newId();
                const sql = `INSERT INTO ${Provider.qualifyControl(this.TABLE)}
                    (${this.ID_COL}, PROYECTO_ID, ${this.NAME_COL}, DESCRIPCION, TABLA, CAMPOS_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                    VALUES ('${Provider.esc(id)}', '${Provider.esc(this.project.PROYECTO_ID)}', '${Provider.esc(name)}', '${Provider.esc(description)}',
                            '${Provider.esc(tableName)}', '${camposJson}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`;
                await Provider.runQuery(sql);
                UI.toast(`Cubo "${name}" creado.`, "success");
            }

            await this.loadList();
            Draco.renderProgress();
        } catch (err) {
            UI.toast("Error al guardar el cubo: " + err.message, "error");
        }
    },

    async remove(id) {
        const cubo = this.list.find(c => c[this.ID_COL] === id);
        if (!cubo) return;

        const ok = await UI.confirm(
            "Eliminar cubo",
            `Se eliminará el cubo <strong>${UI.escapeHtml(cubo[this.NAME_COL])}</strong> y su tabla <strong>${cubo.TABLA}</strong> con todos sus datos.`
        );
        if (!ok) return;

        try {
            await Provider.runQuery(`DROP TABLE IF EXISTS ${Provider.qualify(this.project.DATASET, cubo.TABLA)}`);
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
            UI.toast(`Cubo "${cubo[this.NAME_COL]}" eliminado.`, "success");
            await this.loadList();
            Draco.renderProgress();
        } catch (err) {
            UI.toast("Error al eliminar el cubo: " + err.message, "error");
        }
    }
};
