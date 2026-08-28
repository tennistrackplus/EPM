/**
 * ============================================================
 * DRACO PLANNING — EXPORTADOR LOOKML (.lkml)
 * ============================================================
 * Genera el contenido LookML de un cubo a partir del MISMO objeto de
 * modelo semántico que ya construye SemanticModel.build() (ver
 * js/semantic-model.js y la sección "Modelo semántico (YAML)" del
 * README): { format_version, model, fact, dimensions }.
 *
 * Es el equivalente en Planning de ADDIN/src/js/lkmlExport.js: genera
 * exactamente el mismo tipo de fichero (view de hechos + measures, un
 * view por dimensión con sus atributos y jerarquías documentadas como
 * comentario, y un explore con los joins), para que el resultado sea
 * indistinguible de un .lkml exportado a mano desde el add-in de Excel
 * (ver ejemplo real: cualquier .lkml ya generado por el diseñador de
 * modelos semánticos, p.ej. DEMO1.lkml).
 *
 * SOLO se usa cuando el proveedor activo es BigQuery (ver
 * SemanticModel.generateAndSave en js/semantic-model.js): es una
 * función PURA, no toca Provider ni hace peticiones — de eso se
 * encarga js/github-repo.js.
 */
const LkmlExport = {

    sanitize(name) {
        return String(name || "")
            .toLowerCase()
            .replace(/[^a-z0-9_]+/g, "_")
            .replace(/^_+|_+$/g, "");
    },

    qualifiedTable(database, schema, table) {
        return [database, schema, table].filter(Boolean).join(".");
    },

    // Las medidas del modelo Planning solo llevan un tipo de dato físico
    // (STRING/INTEGER/FLOAT/NUMERIC/BOOLEAN/DATE/DATETIME/TIMESTAMP), no
    // un tipo de agregación LookML: se infiere sum para tipos numéricos
    // y count para el resto (equivalente al "sum" por defecto que ya
    // aplicaba ADDIN/src/js/lkmlExport.js cuando no había agregación).
    measureAggregation(type) {
        const numeric = ["INTEGER", "FLOAT", "FLOAT64", "NUMERIC", "BIGNUMERIC"];
        return numeric.includes(String(type || "").toUpperCase()) ? "sum" : "count";
    },

    /**
     * @param model  objeto devuelto por SemanticModel.build() (o leído de
     *               vuelta del propio YAML): { model, fact, dimensions }
     */
    buildContent(model) {

        const meta = (model && model.model) || {};
        const fact = (model && model.fact) || {};
        const dimensions = (model && model.dimensions) || [];

        const modelId = this.sanitize(meta.name) || "modelo";
        const factViewName = `${modelId}_fact`;
        const factTable = this.qualifiedTable(meta.database, meta.schema, fact.table);

        const lines = [];

        lines.push(`# Modelo semántico: ${meta.name || modelId}`);
        lines.push("# Generado automáticamente al guardar el cubo en Draco Planning.");
        lines.push("");

        /* ---------------- VIEW: tabla de hechos + measures ---------------- */

        lines.push(`view: ${factViewName} {`);
        if (factTable) lines.push(`  sql_table_name: \`${factTable}\` ;;`);
        lines.push("");

        (fact.measures || []).forEach(m => {
            const measureName = this.sanitize(m.name);
            if (!measureName) return;

            lines.push(`  measure: ${measureName} {`);
            lines.push(`    type: ${this.measureAggregation(m.type)}`);
            lines.push(`    sql: \${TABLE}.${m.column || this.sanitize(m.name)} ;;`);
            lines.push("  }");
            lines.push("");
        });

        lines.push("}");
        lines.push("");

        /* ---------------- VIEW por cada dimensión (con sus atributos) ---------------- */

        dimensions.forEach(dim => {

            const dimId = this.sanitize(dim.name);
            if (!dimId) return;

            const dimTable = this.qualifiedTable(meta.database, meta.schema, dim.table);
            const attributes = dim.attributes || [];

            lines.push(`view: ${dimId} {`);
            if (dimTable) lines.push(`  sql_table_name: \`${dimTable}\` ;;`);
            lines.push("");

            if (attributes.length > 0) {
                attributes.forEach(a => {
                    const attrId = this.sanitize(a.column || a.name);
                    if (!attrId) return;

                    lines.push(`  dimension: ${attrId} {`);
                    if (a.is_key) lines.push("    primary_key: yes");
                    if (a.type) lines.push(`    # tipo de origen: ${a.type}`);
                    lines.push(`    sql: \${TABLE}.${a.column || attrId} ;;`);
                    lines.push("  }");
                    lines.push("");
                });
            } else {
                // Dimensión sin atributos todavía: al menos su propia clave.
                lines.push(`  dimension: ${dimId} {`);
                lines.push("    primary_key: yes");
                lines.push(`    sql: \${TABLE}.${dim.key_column || dimId} ;;`);
                lines.push("  }");
                lines.push("");
            }

            // Jerarquías: se documentan como comentario con el orden de niveles
            // (LookML no tiene un bloque nativo de jerarquía genérica).
            (dim.hierarchies || []).forEach(hier => {
                if (!hier.levels || hier.levels.length === 0) return;
                lines.push(`  # Jerarquía "${hier.name}":`);
                hier.levels.forEach(lvl => {
                    lines.push(`  #   ${lvl.level}. ${lvl.attribute}`);
                });
                lines.push("");
            });

            lines.push("}");
            lines.push("");
        });

        /* ---------------- EXPLORE: relaciones fact-dimensión ---------------- */

        lines.push(`explore: ${modelId} {`);
        lines.push(`  view_name: ${factViewName}`);
        lines.push("");

        (fact.foreign_keys || []).forEach(fk => {

            const dimId = this.sanitize(fk.dimension || fk.references_dimension);
            if (!dimId) return;

            const fkColumn = fk.column;
            const refColumn = fk.references_column || fkColumn;

            lines.push(`  join: ${dimId} {`);
            lines.push(`    sql_on: \${${factViewName}.${fkColumn}} = \${${dimId}.${refColumn}} ;;`);
            lines.push("    relationship: many_to_one");
            lines.push("  }");
            lines.push("");
        });

        lines.push("}");

        return lines.join("\n");
    }
};
