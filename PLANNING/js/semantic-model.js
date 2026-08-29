/**
 * ============================================================
 * DRACO PLANNING — MODELO SEMÁNTICO (YAML)
 * ============================================================
 * Al guardar un Cubo (crear o editar), genera un "modelo semántico"
 * en YAML —tabla de hechos, dimensiones con sus atributos, y
 * jerarquías— y lo guarda en el mismo sitio donde Draco guarda ya
 * cualquier otro fichero: el stage de Snowflake (@DRACO_LANDING),
 * reutilizando exactamente el mecanismo de js/storage.js (trocear +
 * INSERT + SP_FINALIZE_FILE_UPLOAD) que ya usa flow_run.html.
 *
 * El objetivo es que un proceso externo (el add-in de Excel, carpeta
 * addin/src) pueda leer luego ese YAML desde el stage y rellenar sus
 * pestañas model_dimension / model_Fact / model_hier sin tener que
 * volver a consultar DRACO_CONTROL. ESTE MÓDULO SOLO GENERA Y GUARDA
 * el fichero — no toca nada del add-in.
 *
 * Estructura completa documentada en README.md, sección
 * "Modelo semántico (YAML)".
 *
 * Si el proveedor activo no es Snowflake, o la subida al stage falla
 * por lo que sea (falta de conectividad, permisos...), el modelo se
 * genera igualmente y se ofrece como descarga en el navegador, para
 * no bloquear nunca el guardado del cubo.
 */
const SemanticModel = {
    // Carpeta dentro del stage donde se guardan los modelos.
    FOLDER: "semantic_models",
    // Versión del propio formato del YAML (para que el add-in pueda
    // saber si sabe leerlo).
    FORMAT_VERSION: 1,

    // ------------------------------------------------------------
    // 1. Construcción del objeto de modelo (antes de serializar)
    // ------------------------------------------------------------
    /**
     * @param project     fila de PROYECTOS
     * @param cuboInfo    { id, name, description, table } del cubo recién guardado
     * @param spec        { dimensions: [{id,name,colId,type}], measures: [{name,type}] } (el mismo spec que Cubes.save ya calcula)
     * @param dimensionRows filas completas de DIMENSIONES del proyecto (Cubes.dimensionsCache)
     * @param hierarchyRows filas de JERARQUIAS de las dimensiones usadas en este cubo
     */
    build(project, cuboInfo, spec, dimensionRows, hierarchyRows) {
        const dimById = new Map(dimensionRows.map(d => [d.DIMENSION_ID, d]));

        const dimensions = spec.dimensions.map(dSpec => {
            const dimRow = dimById.get(dSpec.id);
            const fields = dimRow ? Dimensions.parseFields(dimRow) : [];
            const keyField = fields.find(f => f.__isPrimaryName) || fields[0] || { name: dSpec.name, type: dSpec.type };
            const descAttr = fields.find(f => f.isDescription);

            const attributes = fields.map(f => ({
                name: f.name,
                column: Provider.toIdentifier(f.name),
                type: f.type || "STRING",
                is_key: !!(f.key || f.__isPrimaryName),
                is_description: !!f.isDescription
            }));

            const hierarchies = hierarchyRows
                .filter(h => h.DIMENSION_ID === dSpec.id)
                .map(h => {
                    let niveles = [];
                    try { niveles = JSON.parse(h.NIVELES_JSON || "[]"); } catch (e) { niveles = []; }
                    return {
                        name: h.JERARQUIA,
                        levels: niveles.map((colId, idx) => {
                            const f = fields.find(x => Provider.toIdentifier(x.name) === colId);
                            return { level: idx + 1, attribute: f ? f.name : colId, column: colId };
                        })
                    };
                });

            return {
                name: dimRow ? dimRow.DIMENSION : dSpec.name,
                dimension_id: dSpec.id,
                table: dimRow ? dimRow.TABLA : null,
                description: (dimRow && dimRow.DESCRIPCION) || null,
                key_attribute: keyField.name,
                key_column: Provider.toIdentifier(keyField.name),
                // Atributo marcado con el tick 🏷 en el diseñador de la dimensión.
                // Puede no haber ninguno -> null.
                description_attribute: descAttr ? descAttr.name : null,
                description_column: descAttr ? Provider.toIdentifier(descAttr.name) : null,
                attributes,
                hierarchies
            };
        });

        const foreign_keys = spec.dimensions.map(dSpec => ({
            dimension: dSpec.name,
            column: dSpec.colId,
            references_dimension: dSpec.name,
            references_column: dSpec.colId
        }));

        return {
            format_version: this.FORMAT_VERSION,
            model: {
                name: cuboInfo.name,
                cube_id: cuboInfo.id || null,
                project: project.PROYECTO,
                project_id: project.PROYECTO_ID,
                engine: Provider.key(),
                database: Provider.key() === "snowflake" ? SF.getDatabase() : (typeof BQ !== "undefined" ? BQ.getGcpProject() : null),
                schema: project.DATASET,
                generated_at: new Date().toISOString()
            },
            fact: {
                name: cuboInfo.name,
                table: cuboInfo.table,
                description: cuboInfo.description || null,
                measures: (spec.measures || []).map(m => ({
                    name: m.name,
                    column: Provider.toIdentifier(m.name),
                    type: m.type || "STRING"
                })),
                foreign_keys
            },
            dimensions
        };
    },

    // ------------------------------------------------------------
    // 2. Serializador YAML minimalista (sin dependencias externas)
    // ------------------------------------------------------------
    _scalar(v) {
        if (v === null || v === undefined) return "null";
        if (typeof v === "boolean" || typeof v === "number") return String(v);
        const s = String(v);
        const needsQuote = s === "" ||
            /^\s|\s$/.test(s) ||
            /[:#]/.test(s) ||
            /^[-?:,\[\]{}#&*!|>'"%@`]/.test(s) ||
            /^(true|false|null|yes|no|~)$/i.test(s) ||
            /^-?\d+(\.\d+)?$/.test(s);
        if (!needsQuote) return s;
        return `"${s.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\n/g, "\\n")}"`;
    },

    _lines(value, indent) {
        const pad = "  ".repeat(indent);
        const lines = [];
        if (Array.isArray(value)) {
            if (!value.length) { lines.push(`${pad}[]`); return lines; }
            value.forEach(item => {
                if (item && typeof item === "object") {
                    const sub = this._lines(item, indent + 1);
                    if (sub.length) {
                        const stripped = sub[0].slice((indent + 1) * 2);
                        lines.push(`${pad}- ${stripped}`);
                        lines.push(...sub.slice(1));
                    } else {
                        lines.push(`${pad}- {}`);
                    }
                } else {
                    lines.push(`${pad}- ${this._scalar(item)}`);
                }
            });
            return lines;
        }
        if (value && typeof value === "object") {
            const entries = Object.entries(value).filter(([, v]) => v !== undefined);
            if (!entries.length) { lines.push(`${pad}{}`); return lines; }
            entries.forEach(([k, v]) => {
                if (Array.isArray(v)) {
                    if (!v.length) lines.push(`${pad}${k}: []`);
                    else { lines.push(`${pad}${k}:`); lines.push(...this._lines(v, indent + 1)); }
                } else if (v && typeof v === "object") {
                    const subEntries = Object.entries(v).filter(([, vv]) => vv !== undefined);
                    if (!subEntries.length) lines.push(`${pad}${k}: {}`);
                    else { lines.push(`${pad}${k}:`); lines.push(...this._lines(v, indent + 1)); }
                } else {
                    lines.push(`${pad}${k}: ${this._scalar(v)}`);
                }
            });
            return lines;
        }
        lines.push(`${pad}${this._scalar(value)}`);
        return lines;
    },

    toYaml(value) {
        return this._lines(value, 0).join("\n") + "\n";
    },

    // ------------------------------------------------------------
    // 3. Guardado: genera el YAML de un cubo y lo sube al stage
    // ------------------------------------------------------------
    buildPath(project, cuboIdent) {
        const projIdent = Provider.toIdentifier(project.PROYECTO) || project.PROYECTO_ID;
        return `${this.FOLDER}/${projIdent}/${cuboIdent}.yaml`;
    },

    /** Añade (si no existen aún) las columnas de control del modelo a CUBOS. Idempotente. */
    async ensureColumns() {
        const t = Provider.qualifyControl("CUBOS");
        await Provider.runQuery(`ALTER TABLE ${t} ADD COLUMN IF NOT EXISTS MODELO_YAML_PATH STRING`);
        await Provider.runQuery(`ALTER TABLE ${t} ADD COLUMN IF NOT EXISTS MODELO_YAML_FECHA TIMESTAMP`);
    },

    /**
     * Genera el modelo semántico del cubo recién guardado y lo persiste.
     * Nunca lanza: si algo falla, avisa por toast pero no revierte el
     * guardado del cubo (que ya ha ocurrido cuando se llama a esto).
     */
    async generateAndSave(project, cuboInfo, spec) {
        try {
            const dimIds = (spec.dimensions || []).map(d => d.id).filter(Boolean);
            let hierarchyRows = [];
            if (dimIds.length) {
                const inList = dimIds.map(id => `'${Provider.esc(id)}'`).join(", ");
                hierarchyRows = await Provider.runQuery(`
                    SELECT DIMENSION_ID, JERARQUIA, NIVELES_JSON
                    FROM ${Provider.qualifyControl("JERARQUIAS")}
                    WHERE DIMENSION_ID IN (${inList})`);
            }

            const model = this.build(project, cuboInfo, spec, Cubes.dimensionsCache, hierarchyRows);
            const yamlText = this.toYaml(model);

            const ident = Provider.toIdentifier(cuboInfo.name) || cuboInfo.id;
            const path = this.buildPath(project, ident);
            const fileName = path.split("/").pop();

            let uploaded = false;
            if (Provider.key() === "snowflake" && typeof Storage !== "undefined") {
                try {
                    const file = new File([yamlText], fileName, { type: "text/yaml" });
                    await Storage.upload(file, path);
                    uploaded = true;
                } catch (upErr) {
                    console.error("No se pudo subir el modelo semántico al stage de Snowflake:", upErr);
                }
            }

            try {
                await this.ensureColumns();
                await Provider.runQuery(`
                    UPDATE ${Provider.qualifyControl("CUBOS")}
                    SET MODELO_YAML_PATH = '${Provider.esc(uploaded ? path : "")}',
                        MODELO_YAML_FECHA = CURRENT_TIMESTAMP()
                    WHERE ${Cubes.ID_COL} = '${Provider.esc(cuboInfo.id)}'`);
            } catch (regErr) {
                console.error("No se pudo registrar la ruta del modelo semántico en CUBOS:", regErr);
            }

            if (uploaded) {
                const stage = (typeof DracoConfig !== "undefined" && DracoConfig.snowflakeUploadStage) || "@DRACO_LANDING";
                UI.toast(`Modelo semántico guardado en ${stage}/${path}`, "success");
            } else {
                UI.toast("Modelo semántico generado. No se pudo subir automáticamente al stage: se descarga el YAML.", "info");
                UI.downloadBlob(fileName, yamlText, "text/yaml;charset=utf-8");
            }

            // Si el proveedor activo es BigQuery, además del YAML se genera el
            // .lkml equivalente (mismo formato que ya exporta el add-in de
            // Excel, ver js/lkml-export.js) y se hace commit directo en el
            // repositorio de GitHub configurado en DracoConfig.semanticModelGithub.
            // Nunca bloquea el guardado del cubo: si falla, solo avisa.
            if (Provider.key() === "bigquery") {
                await this.generateAndPushLkml(model, path);
            }
        } catch (err) {
            console.error("Error generando el modelo semántico:", err);
            UI.toast("Aviso: el cubo se guardó, pero falló la generación del modelo semántico YAML.", "error");
        }
    },

    // ------------------------------------------------------------
    // 4. BigQuery: además del YAML, genera el .lkml y lo sube a GitHub
    // ------------------------------------------------------------
    /**
     * @param model    objeto de modelo ya construido por this.build()
     * @param yamlPath ruta (dentro de FOLDER) del .yaml ya generado, para
     *                 derivar la ruta hermana del .lkml (mismo nombre)
     */
    async generateAndPushLkml(model, yamlPath) {
        try {
            if (typeof LkmlExport === "undefined") {
                console.error("js/lkml-export.js no está cargado: no se puede generar el .lkml.");
                return;
            }
            if (typeof GithubRepo === "undefined") {
                console.error("js/github-repo.js no está cargado: no se puede subir el .lkml a GitHub.");
                return;
            }

            const lkmlText = LkmlExport.buildContent(model);
            const lkmlPath = yamlPath.replace(/\.yaml$/i, ".lkml");

            // Justo antes de llamar a GitHub: el token no sale de
            // DracoConfig.semanticModelGithub.token (config.js), sino del
            // secreto "github-pat-draco" en Google Secret Manager.
            const cfg = (typeof DracoConfig !== "undefined" && DracoConfig.semanticModelGithub) || {};
            const token = await BQ.getGithubPatFromSecretManager();
            const repoConfig = { url: cfg.url, branch: cfg.branch, token };

            await GithubRepo.putFile(
                lkmlPath,
                lkmlText,
                `Actualiza ${lkmlPath} desde Draco Planning (cubo "${model.model.name}")`,
                repoConfig
            );

            UI.toast(`Modelo LookML guardado en GitHub: ${lkmlPath}`, "success");
        } catch (err) {
            console.error("No se pudo generar/subir el modelo LookML a GitHub:", err);
            UI.toast("Aviso: el modelo semántico YAML se guardó, pero falló la generación/subida del .lkml a GitHub: " + err.message, "error");
        }
    }
};
