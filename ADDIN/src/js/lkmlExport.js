/**
 * LkmlExport
 * ------------------------------------------------------------------------
 * Genera el contenido LookML (.lkml) de un modelo semántico a partir de su
 * objeto { fact, fields } (el mismo formato que guarda SemanticModelStore).
 *
 * Se ha extraído de semantic_model.js para poder reutilizarlo también desde
 * el diálogo independiente de "Guardar modelo semántico" (saveSemanticModel.html),
 * que NO tiene acceso a Office.context.document.settings (y por tanto no
 * puede usar SemanticModelStore directamente): recibe los modelos ya
 * "aplanados" por querystring y solo necesita esta función pura para
 * construir el LookML.
 *
 * TODO (paso posterior, igual que en el origen): esto es una plantilla
 * mínima con la tabla de hechos y el listado de campos. La generación
 * completa (relaciones, jerarquías...) se ampliará más adelante.
 */
(function () {

    function buildContent(modelName, model) {

        const fact = (model && model.fact) || { project: "", dataset: "", table: "" };
        const fields = (model && model.fields) || [];

        const viewName = modelName.toLowerCase().replace(/[^a-z0-9_]+/g, "_");
        const sqlTableName = [fact.project, fact.dataset, fact.table].filter(Boolean).join(".");

        const lines = [];
        lines.push(`view: ${viewName} {`);
        if (sqlTableName) {
            lines.push(`  sql_table_name: \`${sqlTableName}\` ;;`);
        }
        lines.push("");

        fields.forEach(field => {
            const fieldName = (field.alias || field.name || "").toLowerCase().replace(/[^a-z0-9_]+/g, "_");
            if (!fieldName) return;

            if (field.type === "measure") {
                lines.push(`  measure: ${fieldName} {`);
                lines.push(`    type: ${(field.aggregation || "sum").toLowerCase()}`);
                lines.push(`    sql: \${TABLE}.${field.name} ;;`);
                lines.push("  }");
            } else {
                lines.push(`  dimension: ${fieldName} {`);
                lines.push(`    sql: \${TABLE}.${field.name} ;;`);
                lines.push("  }");
            }
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
