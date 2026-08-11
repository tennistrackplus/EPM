/**
 * Vista previa de jerarquía (diálogo abierto con Office.context.ui.displayDialogAsync
 * desde el editor de jerarquías de semantic_model.js). Página independiente,
 * sin acceso al modelo de objetos de Excel: recibe la especificación de la
 * jerarquía (tabla + niveles) por querystring, la consulta contra BigQuery
 * con SELECT DISTINCT ... ORDER BY (igual que buildHierarchySQL en
 * excelService.js) y pinta el resultado como árbol anidado por nivel.
 */

function getAuthTokenForHierPreview() {
    const token = localStorage.getItem("bigquery_access_token");
    const expires = localStorage.getItem("bigquery_token_expires");
    if (!token || !expires || Date.now() >= parseInt(expires, 10)) {
        return null;
    }
    return token;
}

function showHierPreviewError(msg) {
    const status = document.getElementById("previewStatus");
    status.textContent = msg;
    status.classList.add("preview-error");
    status.style.display = "block";
    document.getElementById("previewTreeWrapper").style.display = "none";
}

/**
 * Agrupa las filas planas (ya ordenadas por BigQuery según los niveles) en
 * un árbol: cada nodo agrupa los valores repetidos de un nivel bajo su
 * padre del nivel anterior.
 */
function buildHierTree(rows, levelCount) {
    const root = { children: new Map() };

    rows.forEach(row => {
        let node = root;
        const cells = row.f || [];
        for (let i = 0; i < levelCount; i++) {
            const raw = cells[i] ? cells[i].v : "";
            const value = (raw === null || raw === undefined || String(raw).trim() === "")
                ? "(vacío)"
                : String(raw);

            if (!node.children.has(value)) {
                node.children.set(value, { children: new Map() });
            }
            node = node.children.get(value);
        }
    });

    return root;
}

function renderHierTree(root, levels, container) {
    container.innerHTML = "";

    function renderLevel(node, depth, parentUl) {
        node.children.forEach((childNode, value) => {
            const li = document.createElement("li");
            li.className = "hier-preview-node";

            const label = document.createElement("div");
            label.className = "hier-preview-label";
            label.innerHTML =
                `<span class="hier-preview-level-tag">${levels[depth] ? levels[depth].label : ""}</span>` +
                `<span class="hier-preview-value"></span>`;
            label.querySelector(".hier-preview-value").textContent = value;
            li.appendChild(label);

            if (childNode.children.size > 0 && depth + 1 < levels.length) {
                const ul = document.createElement("ul");
                ul.className = "hier-preview-children";
                renderLevel(childNode, depth + 1, ul);
                li.appendChild(ul);
            }

            parentUl.appendChild(li);
        });
    }

    const rootUl = document.createElement("ul");
    rootUl.className = "hier-preview-root";
    renderLevel(root, 0, rootUl);
    container.appendChild(rootUl);
}

async function runHierarchyPreview() {
    const params = new URLSearchParams(window.location.search);
    const specRaw = params.get("spec");

    if (!specRaw) {
        showHierPreviewError("No se ha indicado ninguna jerarquía para previsualizar.");
        return;
    }

    let spec;
    try {
        spec = JSON.parse(specRaw);
    } catch (e) {
        showHierPreviewError("No se ha podido leer la definición de la jerarquía.");
        return;
    }

    const project = spec.project;
    const dataset = spec.dataset;
    const table = spec.table;
    const hierarchyName = spec.hierarchyName;
    const levels = spec.levels || [];

    document.getElementById("previewTitle").textContent = hierarchyName
        ? `Jerarquía: ${hierarchyName}`
        : "Vista previa de jerarquía";

    document.getElementById("previewSubtitle").textContent = levels.map(l => l.label).join(" › ");

    if (!project || !dataset || !table || levels.length === 0) {
        showHierPreviewError("La jerarquía no tiene niveles válidos para previsualizar.");
        return;
    }

    const token = getAuthTokenForHierPreview();
    if (!token) {
        showHierPreviewError("Sesión de BigQuery no válida o expirada. Inicia sesión de nuevo desde el panel principal y vuelve a intentarlo.");
        return;
    }

    const fieldList = levels.map(l => l.field).join(", ");
    const escapedTable = "`" + project + "." + dataset + "." + table + "`";
    const sql = "SELECT DISTINCT " + fieldList + " FROM " + escapedTable +
        " ORDER BY " + fieldList + " LIMIT 500";

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
            showHierPreviewError("Error de BigQuery: " + data.error.message);
            return;
        }

        const rows = data.rows || [];

        if (rows.length === 0) {
            showHierPreviewError("No hay datos para mostrar esta jerarquía.");
            return;
        }

        const tree = buildHierTree(rows, levels.length);

        document.getElementById("previewStatus").style.display = "none";
        const wrapper = document.getElementById("previewTreeWrapper");
        wrapper.style.display = "block";
        renderHierTree(tree, levels, wrapper);

    } catch (err) {
        console.error("Error al consultar BigQuery para la vista previa de jerarquía:", err);
        showHierPreviewError("Error al consultar BigQuery: " + (err.message || err));
    }
}

document.addEventListener("DOMContentLoaded", runHierarchyPreview);
