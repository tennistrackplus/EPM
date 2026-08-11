/**
 * Vista previa de datos (diálogo abierto con Office.context.ui.displayDialogAsync
 * desde semantic_model.js). Es una página independiente: no tiene acceso al
 * modelo de objetos de Excel, solo necesita el token de BigQuery (compartido
 * vía localStorage, mismo origen que el resto del add-in) y los parámetros
 * project/dataset/table recibidos por querystring.
 */

function getAuthTokenForPreview() {
    const token = localStorage.getItem("bigquery_access_token");
    const expires = localStorage.getItem("bigquery_token_expires");
    if (!token || !expires || Date.now() >= parseInt(expires, 10)) {
        return null;
    }
    return token;
}

function showPreviewError(msg) {
    const status = document.getElementById("previewStatus");
    status.textContent = msg;
    status.classList.add("preview-error");
    status.style.display = "block";
    document.getElementById("previewTableWrapper").style.display = "none";
}

function formatCellValue(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object") return JSON.stringify(v);
    return String(v);
}

function renderPreviewTable(rows, fields) {
    const table = document.getElementById("previewTable");
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    const headerRow = document.createElement("tr");
    fields.forEach(f => {
        const th = document.createElement("th");
        th.textContent = f.name;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    rows.forEach(r => {
        const tr = document.createElement("tr");
        const cells = r.f || [];
        cells.forEach(cell => {
            const td = document.createElement("td");
            td.textContent = formatCellValue(cell.v);
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    document.getElementById("previewStatus").style.display = "none";
    document.getElementById("previewTableWrapper").style.display = "block";
}

async function runDataPreview() {
    const params = new URLSearchParams(window.location.search);
    const project = params.get("project");
    const dataset = params.get("dataset");
    const table = params.get("table");

    document.getElementById("previewTitle").textContent = table
        ? `Vista previa: ${table}`
        : "Vista previa de datos";

    document.getElementById("previewSubtitle").textContent =
        (project && dataset && table)
            ? `${project}.${dataset}.${table} · primeras 500 filas`
            : "";

    if (!project || !dataset || !table) {
        showPreviewError("No se ha indicado ninguna tabla para previsualizar.");
        return;
    }

    const token = getAuthTokenForPreview();
    if (!token) {
        showPreviewError("Sesión de BigQuery no válida o expirada. Inicia sesión de nuevo desde el panel principal y vuelve a intentarlo.");
        return;
    }

    const escapedTable = "`" + project + "." + dataset + "." + table + "`";
    const sql = "SELECT * FROM " + escapedTable + " LIMIT 500";

    try {
        const response = await fetch(
            `https://bigquery.googleapis.com/bigquery/v2/projects/${encodeURIComponent(project)}/queries`,
            {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ query: sql, useLegacySql: false })
            }
        );

        const data = await response.json();

        if (data.error) {
            showPreviewError("Error de BigQuery: " + data.error.message);
            return;
        }

        const fields = (data.schema && data.schema.fields) || [];
        const rows = data.rows || [];

        if (fields.length === 0) {
            showPreviewError("No se ha podido leer el esquema de la tabla.");
            return;
        }

        if (rows.length === 0) {
            showPreviewError("La tabla no tiene filas (o la consulta no ha devuelto datos).");
            return;
        }

        renderPreviewTable(rows, fields);

    } catch (err) {
        console.error("Error al consultar BigQuery para la vista previa:", err);
        showPreviewError("Error al consultar BigQuery: " + (err.message || err));
    }
}

document.addEventListener("DOMContentLoaded", runDataPreview);
