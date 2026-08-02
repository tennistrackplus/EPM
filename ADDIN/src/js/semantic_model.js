/**
 * Lógica del Modelo Semántico BigQuery EPM para Office Add-in
 */

let fieldsState = [];
let currentConfigFieldIndex = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initEvents();
        loadProjects("factProject");
        loadDatasets("factProject", "factDataset");
    }
});

function initEvents() {
    document.getElementById("factProject").addEventListener("change", () => {
        loadDatasets("factProject", "factDataset");
    });

    document.getElementById("factDataset").addEventListener("change", () => {
        loadTables("factProject", "factDataset", "factTable");
    });

    document.getElementById("btnLoadFields").addEventListener("click", fetchFactFields);

    document.getElementById("dimRelProject").addEventListener("change", () => {
        loadDatasets("dimRelProject", "dimRelDataset");
    });

    document.getElementById("dimRelDataset").addEventListener("change", () => {
        loadTables("dimRelProject", "dimRelDataset", "dimRelTable");
    });

    document.getElementById("btnLoadDimAttributes").addEventListener("click", fetchDimensionAttributes);

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

async function loadProjects(projectSelectId) {
    const token = getAuthToken();
    if (!token) return;

    const projectSelect = document.getElementById(projectSelectId);
    if (!projectSelect) return;

    projectSelect.innerHTML = '<option value="">Cargando...</option>';

    try {
        const response = await fetch("https://bigquery.googleapis.com/bigquery/v2/projects", {
            headers: { "Authorization": `Bearer ${token}` }
        });
        const data = await response.json();

        projectSelect.innerHTML = '<option value="">-- Seleccionar Proyecto --</option>';
        if (data.projects) {
            data.projects.forEach(p => {
                const opt = document.createElement("option");
                opt.value = p.id || p.projectReference.projectId;
                opt.textContent = p.id || p.projectReference.projectId;
                projectSelect.appendChild(opt);
            });
        }
    } catch (err) {
        console.error("Error al cargar proyectos:", err);
    }
}

async function loadDatasets(projectIdInputId, datasetSelectId) {
    const token = getAuthToken();
    if (!token) return;

    const projectId = document.getElementById(projectIdInputId).value.trim();
    if (!projectId) return;

    const datasetSelect = document.getElementById(datasetSelectId);
    datasetSelect.innerHTML = '<option value="">Cargando...</option>';

    try {
        const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets`, {
            headers: { "Authorization": `Bearer ${token}` }
        });
        const data = await response.json();
        
        datasetSelect.innerHTML = '<option value="">-- Seleccionar Dataset --</option>';
        if (data.datasets) {
            data.datasets.forEach(ds => {
                const opt = document.createElement("option");
                opt.value = ds.datasetReference.datasetId;
                opt.textContent = ds.datasetReference.datasetId;
                datasetSelect.appendChild(opt);
            });
        }
    } catch (err) {
        console.error("Error al cargar datasets:", err);
    }
}

async function loadTables(projectIdInputId, datasetSelectId, tableSelectId) {
    const token = getAuthToken();
    if (!token) return;

    const projectId = document.getElementById(projectIdInputId).value.trim();
    const datasetId = document.getElementById(datasetSelectId).value;
    const tableSelect = document.getElementById(tableSelectId);

    if (!projectId || !datasetId) return;

    tableSelect.innerHTML = '<option value="">Cargando...</option>';

    try {
        const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables`, {
            headers: { "Authorization": `Bearer ${token}` }
        });
        const data = await response.json();
        
        tableSelect.innerHTML = '<option value="">-- Seleccionar Tabla --</option>';
        if (data.tables) {
            data.tables.forEach(tbl => {
                const opt = document.createElement("option");
                opt.value = tbl.tableReference.tableId;
                opt.textContent = tbl.tableReference.tableId;
                tableSelect.appendChild(opt);
            });
        }
    } catch (err) {
        console.error("Error al cargar tablas:", err);
    }
}

async function fetchFactFields() {
    const token = getAuthToken();
    if (!token) return;

    const projectId = document.getElementById("factProject").value.trim();
    const datasetId = document.getElementById("factDataset").value;
    const tableId = document.getElementById("factTable").value;

    if (!projectId || !datasetId || !tableId) {
        alert("Por favor selecciona Proyecto, Dataset y Tabla de Hechos.");
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

async function openConfigModal(index) {
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
        
        await loadProjects("dimRelProject");
        document.getElementById("dimRelProject").value = field.relProject || document.getElementById("factProject").value;
        
        loadDatasets("dimRelProject", "dimRelDataset");
        document.getElementById("attributesContainer").style.display = "none";
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

async function fetchDimensionAttributes() {
    const token = getAuthToken();
    if (!token) return;

    const projectId = document.getElementById("dimRelProject").value.trim();
    const datasetId = document.getElementById("dimRelDataset").value;
    const tableId = document.getElementById("dimRelTable").value;

    if (!projectId || !datasetId || !tableId) {
        alert("Por favor selecciona Proyecto, Dataset y Tabla de la Dimensión.");
        return;
    }

    try {
        const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables/${tableId}`, {
            headers: { "Authorization": `Bearer ${token}` }
        });
        const data = await response.json();

        if (data.schema && data.schema.fields) {
            const field = fieldsState[currentConfigFieldIndex];
            field.relProject = projectId;
            field.relDataset = datasetId;
            field.relTable = tableId;

            field.attributes = data.schema.fields.map((attr, idx) => {
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
            <td><input type="text" style="width:40px;" value="${attr.hier1}" onchange="updateAttrHier1(${idx}, this.value)" /></td>
            <td><input type="text" style="width:40px;" value="${attr.hier2}" onchange="updateAttrHier2(${idx}, this.value)" /></td>
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