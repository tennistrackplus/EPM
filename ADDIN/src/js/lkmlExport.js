/**
 * LkmlExport
 * ------------------------------------------------------------------------
 * Genera el contenido LookML (.lkml) de un modelo semántico a partir de su
 * objeto { fact, fields } (el mismo formato que guarda SemanticModelStore,
 * ver js/semanticModelStore.js):
 *
 *   model.fact   = { project, dataset, table }
 *   model.fields = [{
 *       enabled, type: "DIMENSION" | "MEASURE", name,
 *       // DIMENSION:
 *       relProject, relDataset, relTable,      // tabla de la dimensión (o la propia fact si está vacío)
 *       attributes: [{ name, alias, dataType, isKey, enabled }],
 *       hierarchies: [{ name, levels: [{ attribute }] }],
 *       // MEASURE:
 *       aggregation, format
 *   }]
 *
 * Es una función PURA (no toca Excel ni Office.context.document): la usa
 * tanto el taskpane (con el modelo leído de SemanticModelStore) como el
 * diálogo independiente "Guardar modelo semántico" (saveSemanticModel.html),
 * que recibe el modelo ya "aplanado" por querystring porque un diálogo de
 * Office no tiene acceso al modelo de objetos de Excel.
 *
 * Genera:
 *   - Un "view" con la tabla de hechos y sus measures.
 *   - Un "view" por cada dimensión, con sus atributos como dimension: y la
 *     clave marcada con primary_key: yes.
 *   - Las jerarquías de cada dimensión, documentadas como comentario (LookML
 *     no tiene un bloque nativo de jerarquía genérica).
 *   - Un "explore" con el join (relación) de cada dimensión contra la
 *     tabla de hechos.
 */
(function () {

    function sanitize(name) {
        return String(name || "")
            .toLowerCase()
            .replace(/[^a-z0-9_]+/g, "_")
            .replace(/^_+|_+$/g, "");
    }

    function qualifiedTable(project, dataset, table) {
        return [project, dataset, table].filter(Boolean).join(".");
    }

    function buildContent(modelName, model) {

        const fact = (model && model.fact) || {};
        const fields = (model && model.fields) || [];

        const modelId = sanitize(modelName) || "modelo";
        const factViewName = `${modelId}_fact`;
        const factTable = qualifiedTable(fact.project, fact.dataset, fact.table);

        const dimensions = fields.filter(f => f.enabled && f.type === "DIMENSION");
        const measures = fields.filter(f => f.enabled && f.type === "MEASURE");

        const lines = [];

        lines.push(`# Modelo semántico: ${modelName}`);
        lines.push("# Generado automáticamente desde el diseñador de modelos semánticos.");
        lines.push("");

        /* ---------------- VIEW: tabla de hechos + measures ---------------- */

        lines.push(`view: ${factViewName} {`);
        if (factTable) lines.push(`  sql_table_name: \`${factTable}\` ;;`);
        lines.push("");

        measures.forEach(m => {
            const measureName = sanitize(m.name);
            if (!measureName) return;

            lines.push(`  measure: ${measureName} {`);
            lines.push(`    type: ${sanitize(m.aggregation) || "sum"}`);
            lines.push(`    sql: \${TABLE}.${m.name} ;;`);
            if (m.format) lines.push(`    value_format_name: ${sanitize(m.format)}`);
            lines.push("  }");
            lines.push("");
        });

        lines.push("}");
        lines.push("");

        /* ---------------- VIEW por cada dimensión (con sus atributos) ---------------- */

        dimensions.forEach(dim => {

            const dimId = sanitize(dim.name);
            if (!dimId) return;

            const dimTable = qualifiedTable(
                dim.relProject || fact.project,
                dim.relDataset || fact.dataset,
                dim.relTable || fact.table
            );

            const attributes = (dim.attributes || []).filter(a => a.enabled);

            lines.push(`view: ${dimId} {`);
            if (dimTable) lines.push(`  sql_table_name: \`${dimTable}\` ;;`);
            lines.push("");

            if (attributes.length > 0) {
                attributes.forEach(a => {
                    const attrId = sanitize(a.alias || a.name);
                    if (!attrId) return;

                    lines.push(`  dimension: ${attrId} {`);
                    if (a.isKey) lines.push("    primary_key: yes");
                    if (a.dataType) lines.push(`    # tipo de origen: ${a.dataType}`);
                    lines.push(`    sql: \${TABLE}.${a.name} ;;`);
                    lines.push("  }");
                    lines.push("");
                });
            } else {
                // Dimensión sin atributos definidos todavía: al menos su propia clave.
                lines.push(`  dimension: ${dimId} {`);
                lines.push("    primary_key: yes");
                lines.push(`    sql: \${TABLE}.${dim.name} ;;`);
                lines.push("  }");
                lines.push("");
            }

            // Jerarquías: se documentan como comentario con el orden de niveles.
            (dim.hierarchies || []).forEach(hier => {
                if (!hier.levels || hier.levels.length === 0) return;
                lines.push(`  # Jerarquía "${hier.name}":`);
                hier.levels.forEach((lvl, i) => {
                    lines.push(`  #   ${i + 1}. ${lvl.attribute}`);
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

        dimensions.forEach(dim => {

            const dimId = sanitize(dim.name);
            if (!dimId) return;

            const attributes = (dim.attributes || []).filter(a => a.enabled);
            const keyAttr = attributes.find(a => a.isKey);
            const keyId = sanitize(keyAttr ? (keyAttr.alias || keyAttr.name) : dim.name);

            lines.push(`  join: ${dimId} {`);
            lines.push(`    sql_on: \${${factViewName}.${dim.name}} = \${${dimId}.${keyId}} ;;`);
            lines.push("    relationship: many_to_one");
            lines.push("  }");
            lines.push("");
        });

        lines.push("}");

        return lines.join("\n");
    }

    // Descarga "content" como fichero de texto llamado "fileName" en el
    // equipo del usuario (no requiere backend: Blob + enlace temporal).
    function downloadTextFile(fileName, content) {
        const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
        const url = URL.createObjectURL(blob);

        const link = document.createElement("a");
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        URL.revokeObjectURL(url);
    }

    window.LkmlExport = { buildContent, downloadTextFile };

})();
