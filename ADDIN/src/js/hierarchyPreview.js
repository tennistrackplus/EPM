/**
 * Vista previa de jerarquía (diálogo abierto con Office.context.ui.displayDialogAsync
 * desde el editor de jerarquías de semantic_model.js). Página independiente,
 * sin acceso al modelo de objetos de Excel: recibe la especificación de la
 * jerarquía (tabla + niveles) por querystring, la consulta contra el
 * proveedor activo con SELECT DISTINCT ... ORDER BY (igual que
 * buildHierarchySQL en excelService.js) y pinta el resultado como árbol
 * anidado por nivel.
 */

function showHierPreviewError(msg) {
    const status = document.getElementById("previewStatus");
    status.textContent = msg;
    status.classList.add("preview-error");
    status.style.display = "block";
    document.getElementById("previewTreeWrapper").style.display = "none";
}

/**
 * Agrupa las filas planas (ya ordenadas por el proveedor según los niveles) en
 * un árbol: cada nodo agrupa los valores repetidos de un nivel bajo su
 * padre del nivel anterior.
 */
function buildHierTree(rows, levels) {
    const root = { children: new Map() };

    rows.forEach(row => {
        let node = root;
        for (let i = 0; i < levels.length; i++) {
            const raw = row[levels[i].field];
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

    if (!Provider.isConnected()) {
        showHierPreviewError(`Sesión de ${Provider.label()} no válida o expirada. Inicia sesión de nuevo desde el panel principal y vuelve a intentarlo.`);
        return;
    }

    const fieldList = levels.map(l => l.field).join(", ");
    const sql = "SELECT DISTINCT " + fieldList + " FROM " + Provider.qualify(project, dataset, table) +
        " ORDER BY " + fieldList + " LIMIT 500";

    try {
        const { rows } = await Provider.runQuery(sql, project, dataset);

        if (rows.length === 0) {
            showHierPreviewError("No hay datos para mostrar esta jerarquía.");
            return;
        }

        const tree = buildHierTree(rows, levels);

        document.getElementById("previewStatus").style.display = "none";
        const wrapper = document.getElementById("previewTreeWrapper");
        wrapper.style.display = "block";
        renderHierTree(tree, levels, wrapper);

    } catch (err) {
        console.error(`Error al consultar ${Provider.label()} para la vista previa de jerarquía:`, err);
        showHierPreviewError(`Error al consultar ${Provider.label()}: ` + (err.message || err));
    }
}

document.addEventListener("DOMContentLoaded", runHierarchyPreview);
