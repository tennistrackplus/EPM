/**
 * Lógica del Modelo Semántico BigQuery EPM para Office Add-in
 */

let fieldsState = [];
let currentConfigFieldIndex = null;
let currentTreeTarget = "FACT"; // "FACT" o "DIM"

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initEvents();
    }
});

function initEvents() {
    document.getElementById("sapSearchBtn").addEventListener("click", () => openTreeModal("FACT"));
    document.getElementById("factFullConcat").addEventListener("click", () => openTreeModal("FACT"));

    document.getElementById("dimSapSearchBtn").addEventListener("click", () => openTreeModal("DIM"));
    document.getElementById("dimRelFullConcat").addEventListener("click", () => openTreeModal("DIM"));

    document.getElementById("btnCloseTreeModal").addEventListener("click", closeTreeModal);

    document.getElementById("btnSaveMeasureModal").addEventListener("click", saveMeasureModal);
    document.getElementById("btnCloseMeasureModal").addEventListener("click", () => {
        document.getElementById("measureModal").style.display = "none";
    });

    document.getElementById("btnSaveDimModal").addEventListener("click", saveDimModal);
    document.getElementById("btnCloseDimModal").addEventListener("click", () => {
        document.getElementById("dimensionModal").style.display = "none";
    });

    document.getElementById("btnGenerateModel").addEventListener("click", generateSemanticModelInExcel);
}

function getAuthToken() {
    const token = localStorage.getItem("bigquery_access_token");
    const expires = localStorage.getItem("bigquery_token_expires");
    if (!token || !expires || Date.now() >= parseInt(expires)) {
        alert("Sesión no válida o expirada. Por favor, inicia sesión de nuevo.");
        return null;
    }
    return token;
}

async function openTreeModal(target = "FACT") {
    currentTreeTarget = target;
    document.getElementById("treeModal").style.display = "block";
    const container = document.getElementById("treeContainer");
    container.innerHTML = "Cargando proyectos...";

    let autoProject = null;
    let autoDataset = null;

    if (target === "DIM") {
        autoProject = document.getElementById("dimRelProject").value || document.getElementById("factProject").value;
        autoDataset = document.getElementById("dimRelDataset").value || document.getElementById("factDataset").value;
    }

    await loadProjectsTree(container, autoProject, autoDataset);
}

function closeTreeModal() {
    document.getElementById("treeModal").style.display = "none";
}

async function loadProjectsTree(container, autoProject = null, autoDataset = null) {
    const token = getAuthToken();
    if (!token) return;

    try {
        const response = await fetch("https://bigquery.googleapis.com/bigquery/v2/projects", {
            headers: { "Authorization": `Bearer ${token}` }
        });
        const data = await response.json();
        container.innerHTML = "";

        if (data.projects && data.projects.length > 0) {
            const ul = document.createElement("ul");
            ul.className = "tree-list";
            for (const p of data.projects) {
                const projectId = p.id || p.projectReference.projectId;
                const li = document.createElement("li");
                li.className = "tree-item";
                
                const header = document.createElement("div");
                header.className = "tree-header";
                header.innerHTML = `<span class="tree-toggle">▶</span> 📁 <strong>${projectId}</strong>`;
                
                const childrenDiv = document.createElement("div");
                childrenDiv.className = "tree-children";
                
                let loaded = false;
                header.addEventListener("click", async () => {
                    const toggleSpan = header.querySelector(".tree-toggle");
                    if (childrenDiv.classList.contains("open")) {
                        childrenDiv.classList.remove("open");
                        toggleSpan.textContent = "▶";
                    } else {
                        childrenDiv.classList.add("open");
                        toggleSpan.textContent = "▼";
                        if (!loaded) {
                            childrenDiv.innerHTML = "<div style='margin-left:20px; font-size:11px; color:#666;'>Cargando datasets...</div>";
                            await loadDatasetsTree(projectId, childrenDiv, autoDataset);
                            loaded = true;
                        }
                    }
                });

                li.appendChild(header);
                li.appendChild(childrenDiv);
                ul.appendChild(li);

                if (autoProject && projectId === autoProject) {
                    childrenDiv.classList.add("open");
                    header.querySelector(".tree-toggle").textContent = "▼";
                    childrenDiv.innerHTML = "<div style='margin-left:20px; font-size:11px; color:#666;'>Cargando datasets...</div>";
                    loaded = true;
                    loadDatasetsTree(projectId, childrenDiv, autoDataset);
                }
            }
            container.appendChild(ul);
        } else {
            container.innerHTML = "No se encontraron proyectos.";
        }
    } catch (err) {
        console.error("Error al cargar árbol de proyectos:", err);
        container.innerHTML = "Error al cargar proyectos.";
    }
}

async function loadDatasetsTree(projectId, container, autoDataset = null) {
    const token = getAuthToken();
    if (!token) return;

    try {
        const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets`, {
            headers: { "Authorization": `Bearer ${token}` }
        });
        const data = await response.json();
        container.innerHTML = "";

        if (data.datasets && data.datasets.length > 0) {
            const ul = document.createElement("ul");
            ul.className = "tree-list";
            for (const ds of data.datasets) {
                const datasetId = ds.datasetReference.datasetId;
                const li = document.createElement("li");
                li.className = "tree-item";
                
                const header = document.createElement("div");
                header.className = "tree-header";
                header.innerHTML = `<span class="tree-toggle">▶</span> 📊 ${datasetId}`;
                
                const childrenDiv = document.createElement("div");
                childrenDiv.className = "tree-children";
                
                let loaded = false;
                header.addEventListener("click", async (e) => {
                    e.stopPropagation();
                    const toggleSpan = header.querySelector(".tree-toggle");
                    if (childrenDiv.classList.contains("open")) {
                        childrenDiv.classList.remove("open");
                        toggleSpan.textContent = "▶";
                    } else {
                        childrenDiv.classList.add("open");
                        toggleSpan.textContent = "▼";
                        if (!loaded) {
                            childrenDiv.innerHTML = "<div style='margin-left:20px; font-size:11px; color:#666;'>Cargando tablas...</div>";
                            await loadTablesTree(projectId, datasetId, childrenDiv);
                            loaded = true;
                        }
                    }
                });

                li.appendChild(header);
                li.appendChild(childrenDiv);
                ul.appendChild(li);

                if (autoDataset && datasetId === autoDataset) {
                    childrenDiv.classList.add("open");
                    header.querySelector(".tree-toggle").textContent = "▼";
                    childrenDiv.innerHTML = "<div style='margin-left:20px; font-size:11px; color:#666;'>Cargando tablas...</div>";
                    loaded = true;
                    loadTablesTree(projectId, datasetId, childrenDiv);
                }
            }
            container.appendChild(ul);
        } else {
            container.innerHTML = "<div style='margin-left:20px; font-size:11px; color:#666;'>No hay datasets en este proyecto.</div>";
        }
    } catch (err) {
        console.error("Error al cargar datasets:", err);
        container.innerHTML = "<div style='margin-left:20px; font-size:11px; color:#666;'>Error al cargar datasets.</div>";
    }
}

async function loadTablesTree(projectId, datasetId, container) {
    const token = getAuthToken();
    if (!token) return;

    try {
        const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables`, {
            headers: { "Authorization": `Bearer ${token}` }
        });
        const data = await response.json();
        container.innerHTML = "";

        if (data.tables && data.tables.length > 0) {
            const ul = document.createElement("ul");
            ul.className = "tree-list";
            data.tables.forEach(tbl => {
                const tableId = tbl.tableReference.tableId;
                const li = document.createElement("li");
                li.className = "tree-item";
                
                const itemDiv = document.createElement("div");
                itemDiv.className = "tree-header table-item";
                itemDiv.innerHTML = `📋 ${tableId}`;
                
                itemDiv.addEventListener("click", (e) => {
                    e.stopPropagation();
                    selectFactTable(projectId, datasetId, tableId);
                });

                li.appendChild(itemDiv);
                ul.appendChild(li);
            });
            container.appendChild(ul);
        } else {
            container.innerHTML = "<div style='margin-left:20px; font-size:11px; color:#666;'>No hay tablas en este dataset.</div>";
        }
    } catch (err) {
        console.error("Error al cargar tablas:", err);
        container.innerHTML = "<div style='margin-left:20px; font-size:11px; color:#666;'>Error al cargar tablas.</div>";
    }
}

function selectFactTable(projectId, datasetId, tableId) {
    if (currentTreeTarget === "FACT") {
        document.getElementById("factProject").value = projectId;
        document.getElementById("factDataset").value = datasetId;
        document.getElementById("factTable").value = tableId;
        document.getElementById("factFullConcat").value = `${projectId}.${datasetId}.${tableId}`;
        closeTreeModal();
        fetchFactFields();
    } else if (currentTreeTarget === "DIM") {
        document.getElementById("dimRelProject").value = projectId;
        document.getElementById("dimRelDataset").value = datasetId;
        document.getElementById("dimRelTable").value = tableId;
        document.getElementById("dimRelFullConcat").value = `${projectId}.${datasetId}.${tableId}`;
        closeTreeModal();
        fetchDimensionAttributes(true);
    }
}

async function fetchFactFields() {
    const token = getAuthToken();
    if (!token) return;

    const projectId = document.getElementById("factProject").value.trim();
    const datasetId = document.getElementById("factDataset").value;
    const tableId = document.getElementById("factTable").value;

    if (!projectId || !datasetId || !tableId) {
        alert("Por favor selecciona una Tabla de Hechos.");
        return;
    }

    try {
        const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables/${tableId}`, {
            headers: { "Authorization": `Bearer ${token}` }
        });
        const data = await response.json();

        if (data.schema && data.schema.fields) {
            fieldsState = data.schema.fields.map(f => {
                const isNumeric = ["INTEGER", "FLOAT", "NUMERIC", "BIGNUMERIC"].includes(f.type);
                return {
                    name: f.name,
                    alias: f.name,
                    dataType: f.type,
                    type: isNumeric ? "MEASURE" : "DIMENSION",
                    enabled: true,
                    // Config Medida
                    aggregation: "SUM",
                    format: "Auto",
                    // Config Dimensión (Relación)
                    relProject: "",
                    relDataset: "",
                    relTable: "",
                    attributes: []
                };
            });

            renderFieldsTable();
            document.getElementById("fieldsCard").style.display = "block";
        }
    } catch (err) {
        console.error("Error al obtener esquema:", err);
    }
}

function renderFieldsTable() {
    const tbody = document.getElementById("fieldsTbody");
    tbody.innerHTML = "";

    fieldsState.forEach((field, idx) => {
        const tr = document.createElement("tr");

        tr.innerHTML = `
            <td><strong>${field.name}</strong></td>
            <td><input type="text" value="${field.alias}" onchange="updateFieldAlias(${idx}, this.value)" /></td>
            <td>
                <select onchange="updateFieldType(${idx}, this.value)">
                    <option value="DIMENSION" ${field.type === "DIMENSION" ? "selected" : ""}>DIMENSION</option>
                    <option value="MEASURE" ${field.type === "MEASURE" ? "selected" : ""}>MEASURE</option>
                </select>
            </td>
            <td class="checkbox-cell">
                <input type="checkbox" ${field.enabled ? "checked" : ""} onchange="updateFieldEnabled(${idx}, this.checked)" />
            </td>
            <td>
                <button class="btn btn-sm" onclick="openConfigModal(${idx})">Modificar</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function updateFieldAlias(index, val) { fieldsState[index].alias = val; }
function updateFieldType(index, val) { fieldsState[index].type = val; }
function updateFieldEnabled(index, val) { fieldsState[index].enabled = val; }

function openConfigModal(index) {
    currentConfigFieldIndex = index;
    const field = fieldsState[index];

    if (!field.enabled) {
        alert("Habilita el campo para poder modificar sus propiedades.");
        return;
    }

    if (field.type === "MEASURE") {
        document.getElementById("modalMeasureFieldName").textContent = field.name;
        document.getElementById("modalMeasureAlias").value = field.alias;
        document.getElementById("modalMeasureAgg").value = field.aggregation;
        document.getElementById("modalMeasureFormat").value = field.format;
        document.getElementById("measureModal").style.display = "block";
    } else {
        document.getElementById("modalDimFieldName").textContent = field.name;
        document.getElementById("modalDimAlias").value = field.alias;
        
        const proj = field.relProject || document.getElementById("factProject").value;
        const ds = field.relDataset || document.getElementById("factDataset").value;
        const tbl = field.relTable || "";

        document.getElementById("dimRelProject").value = proj;
        document.getElementById("dimRelDataset").value = ds;
        document.getElementById("dimRelTable").value = tbl;

        if (proj && ds && tbl) {
            document.getElementById("dimRelFullConcat").value = `${proj}.${ds}.${tbl}`;
            fetchDimensionAttributes(false);
        } else {
            document.getElementById("dimRelFullConcat").value = "";
            document.getElementById("attributesContainer").style.display = "none";
        }

        document.getElementById("dimensionModal").style.display = "block";
    }
}

function saveMeasureModal() {
    if (currentConfigFieldIndex !== null) {
        fieldsState[currentConfigFieldIndex].alias = document.getElementById("modalMeasureAlias").value;
        fieldsState[currentConfigFieldIndex].aggregation = document.getElementById("modalMeasureAgg").value;
        fieldsState[currentConfigFieldIndex].format = document.getElementById("modalMeasureFormat").value;
        renderFieldsTable();
    }
    document.getElementById("measureModal").style.display = "none";
}

async function fetchDimensionAttributes(forceRefetch = false) {
    const token = getAuthToken();
    if (!token) return;

    const projectId = document.getElementById("dimRelProject").value.trim();
    const datasetId = document.getElementById("dimRelDataset").value;
    const tableId = document.getElementById("dimRelTable").value;

    if (!projectId || !datasetId || !tableId) {
        return;
    }

    const field = fieldsState[currentConfigFieldIndex];

    if (!forceRefetch && field.attributes && field.attributes.length > 0 && field.relTable === tableId && field.relProject === projectId && field.relDataset === datasetId) {
        renderAttributesTable(field.attributes);
        document.getElementById("attributesContainer").style.display = "block";
        return;
    }

    try {
        const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables/${tableId}`, {
            headers: { "Authorization": `Bearer ${token}` }
        });
        const data = await response.json();

        if (data.schema && data.schema.fields) {
            field.relProject = projectId;
            field.relDataset = datasetId;
            field.relTable = tableId;

            const existingAttrsMap = new Map((field.attributes || []).map(a => [a.name, a]));

            field.attributes = data.schema.fields.map((attr, idx) => {
                const existing = existingAttrsMap.get(attr.name);
                if (existing && !forceRefetch) {
                    return existing;
                }
                return {
                    name: attr.name,
                    alias: attr.name,
                    dataType: attr.type,
                    isKey: idx === 0,
                    enabled: true,
                    hier1: "",
                    hier2: ""
                };
            });

            renderAttributesTable(field.attributes);
            document.getElementById("attributesContainer").style.display = "block";
        }
    } catch (err) {
        console.error("Error al obtener atributos de la dimensión:", err);
    }
}

function renderAttributesTable(attributes) {
    const tbody = document.getElementById("dimAttributesTbody");
    tbody.innerHTML = "";

    attributes.forEach((attr, idx) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td><strong>${attr.name}</strong></td>
            <td><input type="text" value="${attr.alias}" onchange="updateAttrAlias(${idx}, this.value)" /></td>
            <td class="checkbox-cell">
                <input type="radio" name="dimKeyGroup" ${attr.isKey ? "checked" : ""} onchange="updateAttrKey(${idx})" />
            </td>
            <td class="checkbox-cell">
                <input type="checkbox" ${attr.enabled ? "checked" : ""} onchange="updateAttrEnabled(${idx}, this.checked)" />
            </td>
            <td><input type="text" style="width:40px;" value="${attr.hier1 || ''}" onchange="updateAttrHier1(${idx}, this.value)" /></td>
            <td><input type="text" style="width:40px;" value="${attr.hier2 || ''}" onchange="updateAttrHier2(${idx}, this.value)" /></td>
        `;
        tbody.appendChild(tr);
    });
}

function updateAttrAlias(idx, val) { fieldsState[currentConfigFieldIndex].attributes[idx].alias = val; }
function updateAttrKey(selectedIdx) {
    fieldsState[currentConfigFieldIndex].attributes.forEach((attr, idx) => {
        attr.isKey = (idx === selectedIdx);
    });
}
function updateAttrEnabled(idx, val) { fieldsState[currentConfigFieldIndex].attributes[idx].enabled = val; }
function updateAttrHier1(idx, val) { fieldsState[currentConfigFieldIndex].attributes[idx].hier1 = val; }
function updateAttrHier2(idx, val) { fieldsState[currentConfigFieldIndex].attributes[idx].hier2 = val; }

function saveDimModal() {
    if (currentConfigFieldIndex !== null) {
        fieldsState[currentConfigFieldIndex].alias = document.getElementById("modalDimAlias").value;
        renderFieldsTable();
    }
    document.getElementById("dimensionModal").style.display = "none";
}

/**
 * Vuelca el modelo semántico completo a Excel rellenando las 5 pestañas objetivo
 */
async function generateSemanticModelInExcel() {
    const factProject = document.getElementById("factProject").value.trim();
    const factDataset = document.getElementById("factDataset").value;
    const factTable = document.getElementById("factTable").value;

    const enabledFields = fieldsState.filter(f => f.enabled);

    // 1. MODEL_RELATIONSHIP
    const relData = [["FACT_PROJECT", "FACT_DATASET", "FACT_TABLE", "FACT_COLUMN", "DIM_PROJECT", "DIM_DATASET", "DIM_TABLE", "DIM_KEY_COLUMN"]];
    enabledFields.filter(f => f.type === "DIMENSION" && f.relTable).forEach(f => {
        const keyAttr = f.attributes.find(a => a.isKey);
        relData.push([
            factProject, factDataset, factTable, f.name,
            f.relProject, f.relDataset, f.relTable, keyAttr ? keyAttr.name : ""
        ]);
    });

    // 2. MODEL_DIMENSION
    const dimData = [["DIMENSION", "ALIAS", "PROJECT", "DATASET", "TABLE"]];
    enabledFields.filter(f => f.type === "DIMENSION").forEach(f => {
        dimData.push([
            f.name, f.alias,
            f.relProject || factProject,
            f.relDataset || factDataset,
            f.relTable || factTable
        ]);
    });

    // 3. MODEL_MEASURES
    const meaData = [["MEASURE", "ALIAS", "AGGREGATION", "FORMAT"]];
    enabledFields.filter(f => f.type === "MEASURE").forEach(f => {
        meaData.push([f.name, f.alias, f.aggregation, f.format]);
    });

    // 4. MODEL_ATRIBUTES
    const attrData = [["DIMENSION", "ATTRIBUTE_COLUMN", "ALIAS", "TYPE", "IS_KEY", "ENABLED"]];
    enabledFields.filter(f => f.type === "DIMENSION").forEach(f => {
        if (f.attributes && f.attributes.length > 0) {
            f.attributes.forEach(a => {
                attrData.push([f.name, a.name, a.alias, a.dataType, a.isKey ? "YES" : "NO", a.enabled ? "YES" : "NO"]);
            });
        }
    });

    // 5. MODEL_HIER
    const hierData = [["DIMENSION", "HIERARCHY_LEVEL", "ATTRIBUTE_COLUMN"]];
    enabledFields.filter(f => f.type === "DIMENSION").forEach(f => {
        if (f.attributes) {
            f.attributes.forEach(a => {
                if (a.hier1) hierData.push([f.name, `LEVEL_${a.hier1}`, a.name]);
                if (a.hier2) hierData.push([f.name, `LEVEL_${a.hier2}`, a.name]);
            });
        }
    });

    try {
        await Excel.run(async (context) => {
            const sheetsMap = [
                { name: "MODEL_RELATIONSHIP", data: relData },
                { name: "MODEL_DIMENSION", data: dimData },
                { name: "MODEL_MEASURES", data: meaData },
                { name: "MODEL_ATRIBUTES", data: attrData },
                { name: "MODEL_HIER", data: hierData }
            ];

            const sheets = context.workbook.worksheets;

            for (const item of sheetsMap) {
                let sheet = sheets.getItemOrNullObject(item.name);
                await context.sync();

                if (sheet.isNullObject) {
                    sheet = sheets.add(item.name);
                } else {
                    sheet.getUsedRangeOrNullObject().clear();
                }

                if (item.data.length > 0) {
                    const rows = item.data.length;
                    const cols = item.data[0].length;
                    const range = sheet.getRangeByIndexes(0, 0, rows, cols);
                    range.values = item.data;
                }
            }

            await context.sync();
            alert("¡Modelo Semántico generado con éxito en Excel!");
        });
    } catch (err) {
        console.error("Error al escribir el modelo semántico en Excel:", err);
        alert("Error al escribir en Excel: " + err.message);
    }
}