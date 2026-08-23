/**
 * Vista previa de datos (diálogo abierto con Office.context.ui.displayDialogAsync
 * desde semantic_model.js). Es una página independiente: no tiene acceso al
 * modelo de objetos de Excel, solo necesita la sesión del proveedor activo
 * (compartida vía localStorage, mismo origen que el resto del add-in) y los
 * parámetros project/dataset/table recibidos por querystring.
 */

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
        fields.forEach(f => {
            const td = document.createElement("td");
            td.textContent = formatCellValue(r[f.name]);
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

    if (!Provider.isConnected()) {
        showPreviewError(`Sesión de ${Provider.label()} no válida o expirada. Inicia sesión de nuevo desde el panel principal y vuelve a intentarlo.`);
        return;
    }

    const sql = "SELECT * FROM " + Provider.qualify(project, dataset, table) + " LIMIT 500";

    try {
        const { fields, rows } = await Provider.runQuery(sql, project, dataset);

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
        console.error(`Error al consultar ${Provider.label()} para la vista previa:`, err);
        showPreviewError(`Error al consultar ${Provider.label()}: ` + (err.message || err));
    }
}

document.addEventListener("DOMContentLoaded", runDataPreview);
