/**
 * ============================================================
 * DRACO PLANNING — TABLAS DE PARAMETRIZACIÓN
 * ============================================================
 * Módulo sencillo para crear tablas físicas de parametrización:
 * a diferencia de las Dimensiones, aquí NO hay clave automática ni
 * jerarquías ni editor de valores — solo el diseño de la tabla
 * (nombre, descripción y campos con nombre/descripción/tipo/clave)
 * y la creación/sincronización de la tabla física correspondiente.
 */
const Parametrizacion = {
    TABLE: "PARAMETRIZACIONES",
    NAME_COL: "NOMBRE",
    ID_COL: "PARAMETRIZACION_ID",
    list: [],

    async render(container, project) {
        this.container = container;
        this.project = project;

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <h3>Tablas de parametrización</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET}</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewParamTable">
                    + Nueva tabla
                </button>
            </div>
            <p class="form-hint">Tablas físicas auxiliares para parámetros de cálculo, catálogos, etc. Aquí solo se define su estructura (campos, tipos y clave); la carga o edición de sus valores se hace desde tu herramienta habitual de consultas SQL.</p>
            <div id="paramTablesListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewParamTable").addEventListener("click", () => this.openForm());
        await this.loadList();
    },

    async loadList() {
        const wrap = document.getElementById("paramTablesListWrap");
        try {
            const sql = `SELECT ${this.ID_COL}, ${this.NAME_COL}, DESCRIPCION, TABLA, CAMPOS_JSON
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                         ORDER BY ${this.NAME_COL}`;
            this.list = await Provider.runQuery(sql);

            if (!this.list.length) {
                wrap.innerHTML = `<div class="module-empty">Todavía no hay tablas de parametrización en este proyecto. Crea la primera con "Nueva tabla".</div>`;
                return;
            }

            wrap.innerHTML = `
                <div class="data-list">
                    <table>
                        <thead><tr><th>Tabla</th><th>Descripción</th><th>Tabla física</th><th>Campos</th><th></th></tr></thead>
                        <tbody>
                            ${this.list.map(p => `
                                <tr>
                                    <td><strong>${UI.escapeHtml(p[this.NAME_COL])}</strong></td>
                                    <td>${UI.escapeHtml(p.DESCRIPCION || "—")}</td>
                                    <td><span class="table-tag">${UI.escapeHtml(p.TABLA)}</span></td>
                                    <td>${this.parseFields(p).length}</td>
                                    <td>
                                        <div class="row-actions">
                                            <button data-edit="${p[this.ID_COL]}" title="Editar">✎</button>
                                            <button data-del="${p[this.ID_COL]}" class="danger" title="Eliminar">🗑</button>
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
            wrap.innerHTML = `<div class="module-empty">Error al cargar las tablas de parametrización: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    parseFields(p) {
        try { return JSON.parse(p.CAMPOS_JSON || "[]"); } catch (e) { return []; }
    },

    async openForm(editId = null) {
        const editing = editId ? this.list.find(p => p[this.ID_COL] === editId) : null;
        const fields = editing ? this.parseFields(editing) : [];

        const result = await UI.openParamTableFormModal({
            title: editing ? `Editar tabla de parametrización: ${editing[this.NAME_COL]}` : "Nueva tabla de parametrización",
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
            UI.toast("El nombre de la tabla no es válido.", "error");
            return;
        }
        const tableName = `${DracoConfig.prefix}PARAM_${ident}`;

        // Nombres físicos (identificador) + tipo, para sincronizar columnas
        // sin recrear la tabla (ver Provider.syncTableColumns).
        const physicalFields = fields.map(f => ({ name: Provider.toIdentifier(f.name), type: f.type }));

        try {
            // No se usa CREATE OR REPLACE TABLE: eso borraría los datos ya
            // cargados cada vez que se edita la estructura. Se sincronizan
            // las columnas físicas (se añaden las nuevas, se quitan las
            // eliminadas) preservando los datos del resto.
            await Provider.syncTableColumns(this.project.DATASET, tableName, physicalFields);

            const camposJson = JSON.stringify(fields).replace(/\\/g, "\\\\").replace(/'/g, "\\'");

            if (editing) {
                const sql = `UPDATE ${Provider.qualifyControl(this.TABLE)}
                    SET DESCRIPCION = '${Provider.esc(description)}',
                        CAMPOS_JSON = '${camposJson}',
                        FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                    WHERE ${this.ID_COL} = '${Provider.esc(editing[this.ID_COL])}'`;
                await Provider.runQuery(sql);
                UI.toast(`Tabla de parametrización "${name}" actualizada.`, "success");
            } else {
                const id = Provider.newId();
                const sql = `INSERT INTO ${Provider.qualifyControl(this.TABLE)}
                    (${this.ID_COL}, PROYECTO_ID, ${this.NAME_COL}, DESCRIPCION, TABLA, CAMPOS_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                    VALUES ('${Provider.esc(id)}', '${Provider.esc(this.project.PROYECTO_ID)}', '${Provider.esc(name)}', '${Provider.esc(description)}',
                            '${Provider.esc(tableName)}', '${camposJson}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`;
                await Provider.runQuery(sql);
                UI.toast(`Tabla de parametrización "${name}" creada.`, "success");
            }

            await this.loadList();
        } catch (err) {
            UI.toast("Error al guardar la tabla de parametrización: " + err.message, "error");
        }
    },

    async remove(id) {
        const p = this.list.find(x => x[this.ID_COL] === id);
        if (!p) return;

        const ok = await UI.confirm(
            "Eliminar tabla de parametrización",
            `Se eliminará la tabla <strong>${UI.escapeHtml(p[this.NAME_COL])}</strong> y su tabla física <strong>${p.TABLA}</strong> con todos sus datos.`
        );
        if (!ok) return;

        try {
            await Provider.runQuery(`DROP TABLE IF EXISTS ${Provider.qualify(this.project.DATASET, p.TABLA)}`);
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
            UI.toast(`Tabla de parametrización "${p[this.NAME_COL]}" eliminada.`, "success");
            await this.loadList();
        } catch (err) {
            UI.toast("Error al eliminar la tabla de parametrización: " + err.message, "error");
        }
    }
};
