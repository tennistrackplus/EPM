/**
 * Módulo de Cubos. Un cubo se compone de:
 *  - Dimensiones: seleccionadas entre las ya creadas en el proyecto.
 *    Cada una aporta una columna FK (misma clave e igual tipo que la
 *    clave principal de esa dimensión).
 *  - Medidas: campos numéricos/otros definidos libremente, como antes.
 */
const Cubes = {
    TABLE: "CUBOS",
    NAME_COL: "CUBOS",
    ID_COL: "CUBO_ID",
    list: [],
    dimensionsCache: [],

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
        await this.loadDimensions();
        await this.loadList();
    },

    async loadDimensions() {
        const sql = `SELECT DIMENSION_ID, DIMENSION, DESCRIPCION, TABLA, CAMPOS_JSON
                     FROM ${Provider.qualifyControl("DIMENSIONES")}
                     WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                     ORDER BY DIMENSION`;
        this.dimensionsCache = await Provider.runQuery(sql);
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
                        <thead><tr><th>Cubo</th><th>Dimensiones</th><th>Medidas</th><th>Tabla</th><th></th></tr></thead>
                        <tbody>
                            ${this.list.map(c => {
                                const spec = this.parseSpec(c);
                                return `
                                <tr>
                                    <td><strong>${UI.escapeHtml(c[this.NAME_COL])}</strong><br><span class="col-type">${UI.escapeHtml(c.DESCRIPCION || "")}</span></td>
                                    <td>${spec.dimensions.map(d => `<span class="table-tag">${UI.escapeHtml(d.name)}</span>`).join(" ") || "—"}</td>
                                    <td>${spec.measures.map(m => `<span class="table-tag">${UI.escapeHtml(m.name)}</span>`).join(" ") || "—"}</td>
                                    <td><span class="table-tag">${UI.escapeHtml(c.TABLA)}</span></td>
                                    <td>
                                        <div class="row-actions">
                                            <button data-edit="${c[this.ID_COL]}" title="Editar">✎</button>
                                            <button data-del="${c[this.ID_COL]}" class="danger" title="Eliminar">🗑</button>
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
            wrap.innerHTML = `<div class="module-empty">Error al cargar cubos: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    parseSpec(cubo) {
        try {
            const spec = JSON.parse(cubo.CAMPOS_JSON || "{}");
            return { dimensions: spec.dimensions || [], measures: spec.measures || [] };
        } catch (e) {
            return { dimensions: [], measures: [] };
        }
    },

    async openForm(editId = null) {
        const editing = editId ? this.list.find(c => c[this.ID_COL] === editId) : null;
        const spec = editing ? this.parseSpec(editing) : { dimensions: [], measures: [] };

        if (!this.dimensionsCache.length) {
            UI.toast("Aviso: este proyecto todavía no tiene dimensiones creadas.", "info");
        }

        const result = await UI.openCubeFormModal({
            title: editing ? `Editar cubo: ${editing[this.NAME_COL]}` : "Nuevo cubo",
            name: editing ? editing[this.NAME_COL] : "",
            description: editing ? (editing.DESCRIPCION || "") : "",
            dimensionsList: this.dimensionsCache,
            selectedDimensionIds: spec.dimensions.map(d => d.id),
            measures: spec.measures,
            nameEditable: !editing
        });

        if (!result) return;
        await this.save(editing, result);
    },

    async save(editing, { name, description, dimensionIds, measures }) {
        const ident = Provider.toIdentifier(name);
        if (!ident) {
            UI.toast("El nombre del cubo no es válido.", "error");
            return;
        }
        const tableName = `${DracoConfig.prefix}${ident}`;

        const selectedDims = dimensionIds.map(id => this.dimensionsCache.find(d => d.DIMENSION_ID === id)).filter(Boolean);
        const dimSpecs = selectedDims.map(d => {
            const fields = Dimensions.parseFields(d);
            const keyField = fields[0] || { type: "STRING" };
            return {
                id: d.DIMENSION_ID,
                name: d.DIMENSION,
                colId: Provider.toIdentifier(d.DIMENSION),
                type: keyField.type || "STRING"
            };
        });

        // Nombres físicos (identificador) + tipo, para sincronizar columnas
        // sin recrear la tabla (ver Provider.syncTableColumns).
        const physicalFields = [
            ...dimSpecs.map(d => ({ name: d.colId, type: d.type })),
            ...measures.map(f => ({ name: Provider.toIdentifier(f.name), type: f.type }))
        ];

        if (!physicalFields.length) {
            UI.toast("El cubo necesita al menos una dimensión o una medida.", "error");
            return;
        }

        try {
            // No se usa CREATE OR REPLACE TABLE: eso borraría todos los
            // datos ya cargados en el cubo cada vez que se edita. Se
            // sincronizan las columnas físicas (se añaden las nuevas, se
            // quitan las eliminadas) preservando los datos del resto.
            await Provider.syncTableColumns(this.project.DATASET, tableName, physicalFields);

            const spec = { dimensions: dimSpecs, measures };
            const camposJson = JSON.stringify(spec).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
            const cuboId = editing ? editing[this.ID_COL] : Provider.newId();

            if (editing) {
                const sql = `UPDATE ${Provider.qualifyControl(this.TABLE)}
                    SET DESCRIPCION = '${Provider.esc(description)}',
                        CAMPOS_JSON = '${camposJson}',
                        FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                    WHERE ${this.ID_COL} = '${Provider.esc(cuboId)}'`;
                await Provider.runQuery(sql);
                UI.toast(`Cubo "${name}" actualizado.`, "success");
            } else {
                const sql = `INSERT INTO ${Provider.qualifyControl(this.TABLE)}
                    (${this.ID_COL}, PROYECTO_ID, ${this.NAME_COL}, DESCRIPCION, TABLA, CAMPOS_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                    VALUES ('${Provider.esc(cuboId)}', '${Provider.esc(this.project.PROYECTO_ID)}', '${Provider.esc(name)}', '${Provider.esc(description)}',
                            '${Provider.esc(tableName)}', '${camposJson}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`;
                await Provider.runQuery(sql);
                UI.toast(`Cubo "${name}" creado.`, "success");
            }

            await this.loadList();
            Draco.renderProgress();

            // Genera y guarda el modelo semántico (YAML) de este cubo. No
            // bloquea ni revierte el guardado del cubo si falla: solo avisa.
            if (typeof SemanticModel !== "undefined") {
                SemanticModel.generateAndSave(
                    this.project,
                    { id: cuboId, name, description, table: tableName },
                    spec
                );
            }
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
