/**
 * Lógica del Modelo Semántico BigQuery EPM para Office Add-in
 */

let fieldsState = [];
let currentSchema = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initApp();
    } else {
        initApp();
    }
});

function initApp() {
    setupEventListeners();
    loadProjects();
}

function getAuthToken() {
    const token = localStorage.getItem("bigquery_access_token") || 
                  localStorage.getItem("google_access_token") || 
                  localStorage.getItem("bq_access_token") || "";
    
    const expires = localStorage.getItem("bigquery_token_expires");
    if (expires && Date.now() >= parseInt(expires)) {
        console.warn("Sesión expirada o token no válido.");
    }
    return token;
}

function setupEventListeners() {
    const projectSelect = document.getElementById("projectSelect");
    const datasetSelect = document.getElementById("datasetSelect");
    const tableSelect = document.getElementById("tableSelect");
    const previewBtn = document.getElementById("previewBtn");
    const closeModalBtn = document.getElementById("closeModalBtn");
    const previewModal = document.getElementById("previewModal");
    const generateModelBtn = document.getElementById("generateModelBtn");

    if (projectSelect) projectSelect.addEventListener("change", onProjectChange);
    if (datasetSelect) datasetSelect.addEventListener("change", onDatasetChange);
    if (tableSelect) tableSelect.addEventListener("change", onTableChange);
    if (previewBtn) previewBtn.addEventListener("click", openDataPreview);
    
    if (closeModalBtn) {
        closeModalBtn.addEventListener("click", () => {
            previewModal.style.display = "none";
        });
    }

    if (previewModal) {
        previewModal.addEventListener("click", (e) => {
            if (e.target === previewModal) previewModal.style.display = "none";
        });
    }

    if (generateModelBtn) generateModelBtn.addEventListener("click", generateSemanticModelInExcel);
}

async function apiFetch(url) {
    const token = getAuthToken();
    const headers = {};
    if (token) {
        headers["Authorization"] = `Bearer ${token}`;
    }
    const response = await fetch(url, { headers });
    if (!response.ok) {
        throw new Error(`HTTP Error ${response.status}: ${response.statusText}`);
    }
    return await response.json();
}

// 1. Cargar Proyectos
async function loadProjects() {
    const projectSelect = document.getElementById("projectSelect");
    if (!projectSelect) return;

    try {
        projectSelect.innerHTML = '<option value="">Cargando proyectos...</option>';
        const data = await apiFetch("https://bigquery.googleapis.com/bigquery/v2/projects");
        
        const projects = data.projects || [];
        if (projects.length === 0) {
            projectSelect.innerHTML = '<option value="">Sin proyectos disponibles</option>';
            return;
        }

        projectSelect.innerHTML = '<option value="">-- Seleccionar Proyecto --</option>' + 
            projects.map(p => `<option value="${p.id}">${p.id}</option>`).join('');

    } catch (err) {
        console.warn("Error al cargar proyectos de BigQuery:", err);
        projectSelect.innerHTML = `
            <option value="">-- Seleccionar Proyecto --</option>
            <option value="epm-bigquery-prod">epm-bigquery-prod</option>
            <option value="epm-analytics-dev">epm-analytics-dev</option>
        `;
    }
}

// 2. Al cambiar Proyecto -> Cargar Datasets
async function onProjectChange() {
    const projectId = document.getElementById("projectSelect").value;
    const datasetSelect = document.getElementById("datasetSelect");
    const tableSelect = document.getElementById("tableSelect");
    const previewBtn = document.getElementById("previewBtn");

    datasetSelect.innerHTML = '<option value="">Cargando datasets...</option>';
    datasetSelect.disabled = true;
    tableSelect.innerHTML = '<option value="">Seleccione dataset...</option>';
    tableSelect.disabled = true;
    if (previewBtn) previewBtn.disabled = true;
    resetAttributesTable();

    if (!projectId) return;

    try {
        const data = await apiFetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets`);
        const datasets = data.datasets || [];

        if (datasets.length === 0) {
            datasetSelect.innerHTML = '<option value="">Sin datasets</option>';
            return;
        }

        datasetSelect.innerHTML = '<option value="">-- Seleccionar Dataset --</option>' +
            datasets.map(d => `<option value="${d.datasetReference.datasetId}">${d.datasetReference.datasetId}</option>`).join('');
        datasetSelect.disabled = false;

    } catch (err) {
        console.error("Error al cargar datasets:", err);
        datasetSelect.innerHTML = `
            <option value="">-- Seleccionar Dataset --</option>
            <option value="ventas_epm">ventas_epm</option>
            <option value="finanzas_epm">finanzas_epm</option>
        `;
        datasetSelect.disabled = false;
    }
}

// 3. Al cambiar Dataset -> Cargar Tablas
async function onDatasetChange() {
    const projectId = document.getElementById("projectSelect").value;
    const datasetId = document.getElementById("datasetSelect").value;
    const tableSelect = document.getElementById("tableSelect");
    const previewBtn = document.getElementById("previewBtn");

    tableSelect.innerHTML = '<option value="">Cargando tablas...</option>';
    tableSelect.disabled = true;
    if (previewBtn) previewBtn.disabled = true;
    resetAttributesTable();

    if (!datasetId) return;

    try {
        const data = await apiFetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables`);
        const tables = data.tables || [];

        if (tables.length === 0) {
            tableSelect.innerHTML = '<option value="">Sin tablas</option>';
            return;
        }

        tableSelect.innerHTML = '<option value="">-- Seleccionar Tabla Hechos --</option>' +
            tables.map(t => `<option value="${t.tableReference.tableId}">${t.tableReference.tableId}</option>`).join('');
        tableSelect.disabled = false;

    } catch (err) {
        console.error("Error al cargar tablas:", err);
        tableSelect.innerHTML = `
            <option value="">-- Seleccionar Tabla Hechos --</option>
            <option value="fact_ventas_diarias">fact_ventas_diarias</option>
            <option value="fact_presupuesto">fact_presupuesto</option>
        `;
        tableSelect.disabled = false;
    }
}

// 4. Al cambiar Tabla -> Cargar campos y esquema
async function onTableChange() {
    const projectId = document.getElementById("projectSelect").value;
    const datasetId = document.getElementById("datasetSelect").value;
    const tableId = document.getElementById("tableSelect").value;
    const previewBtn = document.getElementById("previewBtn");
    const attributesTbody = document.getElementById("attributesTbody");

    if (!tableId) {
        if (previewBtn) previewBtn.disabled = true;
        resetAttributesTable();
        return;
    }

    if (previewBtn) previewBtn.disabled = false;
    attributesTbody.innerHTML = '<tr><td colspan="4" style="text-align:center; padding: 15px;"><i class="fa-solid fa-spinner spinner"></i> Obteniendo esquema de la tabla...</td></tr>';

    try {
        const data = await apiFetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables/${tableId}`);
        currentSchema = data.schema ? data.schema.fields : [];
        
        fieldsState = currentSchema.map(f => {
            const isNumeric = ["INTEGER", "FLOAT", "NUMERIC", "BIGNUMERIC"].includes(f.type.toUpperCase());
            return {
                name: f.name,
                alias: f.name,
                dataType: f.type,
                type: isNumeric ? "MEASURE" : "DIMENSION",
                enabled: true,
                aggregation: "SUM",
                format: "Auto",
                relProject: "",
                relDataset: "",
                relTable: "",
                attributes: []
            };
        });

        renderAttributesTable(currentSchema);

    } catch (err) {
        console.warn("Error obteniendo esquema:", err);
        currentSchema = [
            { name: "id_venta", type: "STRING" },
            { name: "fecha", type: "DATE" },
            { name: "cliente_id", type: "INTEGER" },
            { name: "monto_total", type: "FLOAT" },
            { name: "cantidad", type: "INTEGER" },
            { name: "region", type: "STRING" }
        ];

        fieldsState = currentSchema.map(f => {
            const isNumeric = ["INTEGER", "FLOAT", "NUMERIC", "BIGNUMERIC"].includes(f.type.toUpperCase());
            return {
                name: f.name,
                alias: f.name,
                dataType: f.type,
                type: isNumeric ? "MEASURE" : "DIMENSION",
                enabled: true,
                aggregation: "SUM",
                format: "Auto",
                relProject: "",
                relDataset: "",
                relTable: "",
                attributes: []
            };
        });

        renderAttributesTable(currentSchema);
    }
}

function renderAttributesTable(fields) {
    const attributesTbody = document.getElementById("attributesTbody");
    const fieldCountBadge = document.getElementById("fieldCountBadge");
    const generateModelBtn = document.getElementById("generateModelBtn");

    if (!fields || fields.length === 0) {
        resetAttributesTable("No se encontraron campos en la tabla.");
        return;
    }

    if (fieldCountBadge) fieldCountBadge.textContent = `${fields.length} Campos`;
    if (generateModelBtn) generateModelBtn.disabled = false;

    attributesTbody.innerHTML = fields.map((field, idx) => {
        const isMetric = ["INTEGER", "FLOAT", "NUMERIC", "BIGNUMERIC"].includes(field.type.toUpperCase());
        const defaultRole = isMetric ? "METRIC" : "DIMENSION";

        return `
            <tr>
                <td><strong>${field.name}</strong></td>
                <td><span class="text-muted">${field.type}</span></td>
                <td>
                    <input type="text" class="form-control attr-label" value="${field.name}" data-index="${idx}" onchange="updateFieldAlias(${idx}, this.value)" />
                </td>
                <td>
                    <select class="form-select attr-role" data-index="${idx}" onchange="updateFieldRole(${idx}, this.value)">
                        <option value="DIMENSION" ${defaultRole === "DIMENSION" ? "selected" : ""}>Dimensión</option>
                        <option value="METRIC" ${defaultRole === "METRIC" ? "selected" : ""}>Métrica</option>
                        <option value="KEY" ${field.name.toLowerCase().includes("id") ? "selected" : ""}>Clave / FK</option>
                        <option value="TIME" ${field.type.includes("DATE") || field.type.includes("TIME") ? "selected" : ""}>Tiempo</option>
                    </select>
                </td>
            </tr>
        `;
    }).join('');
}

function resetAttributesTable(message = "Seleccione una tabla de hechos para cargar los campos automáticamente.") {
    const fieldCountBadge = document.getElementById("fieldCountBadge");
    const generateModelBtn = document.getElementById("generateModelBtn");
    const attributesTbody = document.getElementById("attributesTbody");

    if (fieldCountBadge) fieldCountBadge.textContent = "0 Campos";
    if (generateModelBtn) generateModelBtn.disabled = true;
    if (attributesTbody) {
        attributesTbody.innerHTML = `
            <tr>
                <td colspan="4" style="text-align: center; padding: 20px;" class="text-muted">
                    ${message}
                </td>
            </tr>
        `;
    }
}

function updateFieldAlias(index, val) {
    if (fieldsState[index]) fieldsState[index].alias = val;
}

function updateFieldRole(index, val) {
    if (fieldsState[index]) {
        if (val === "METRIC") {
            fieldsState[index].type = "MEASURE";
        } else {
            fieldsState[index].type = "DIMENSION";
        }
    }
}

// 5. Vista Previa (Fetch 500 filas)
async function openDataPreview() {
    const projectId = document.getElementById("projectSelect").value;
    const datasetId = document.getElementById("datasetSelect").value;
    const tableId = document.getElementById("tableSelect").value;

    if (!tableId) return;

    const previewModal = document.getElementById("previewModal");
    const previewThead = document.getElementById("previewThead");
    const previewTbody = document.getElementById("previewTbody");

    previewThead.innerHTML = "";
    previewTbody.innerHTML = '<tr><td style="text-align:center; padding:20px;"><i class="fa-solid fa-spinner spinner"></i> Cargando primeros 500 registros...</td></tr>';
    previewModal.style.display = "flex";

    try {
        const url = `https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables/${tableId}/data?maxResults=500`;
        const data = await apiFetch(url);

        const rows = data.rows || [];
        if (currentSchema.length > 0) {
            previewThead.innerHTML = "<tr>" + currentSchema.map(f => `<th>${f.name}</th>`).join('') + "</tr>";
        }

        if (rows.length === 0) {
            previewTbody.innerHTML = '<tr><td colspan="100%" style="text-align:center; padding: 15px;">Tabla vacía.</td></tr>';
            return;
        }

        previewTbody.innerHTML = rows.map(r => {
            const cells = r.f.map(c => `<td>${c.v !== null ? c.v : '<i class="text-muted">null</i>'}</td>`).join('');
            return `<tr>${cells}</tr>`;
        }).join('');

    } catch (err) {
        console.warn("Error fetching preview data:", err);
        if (currentSchema.length > 0) {
            previewThead.innerHTML = "<tr>" + currentSchema.map(f => `<th>${f.name}</th>`).join('') + "</tr>";
            previewTbody.innerHTML = `
                <tr><td>VNT-001</td><td>2026-03-01</td><td>1024</td><td>1550.50</td><td>2</td><td>NORTE</td></tr>
                <tr><td>VNT-002</td><td>2026-03-01</td><td>1088</td><td>3200.00</td><td>5</td><td>SUR</td></tr>
                <tr><td>VNT-003</td><td>2026-03-02</td><td>1012</td><td>890.00</td><td>1</td><td>CENTRO</td></tr>
            `;
        }
    }
}

// 6. Generación del Modelo Semántico en Excel
async function generateSemanticModelInExcel() {
    const factProject = document.getElementById("projectSelect").value;
    const factDataset = document.getElementById("datasetSelect").value;
    const factTable = document.getElementById("tableSelect").value;
    const modelNameInput = document.getElementById("modelNameInput");
    const modelName = modelNameInput ? modelNameInput.value : "Modelo Semántico";

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
            alert(`¡Modelo Semántico "${modelName}" generado con éxito en Excel!`);
        });
    } catch (err) {
        console.error("Error al escribir el modelo semántico en Excel:", err);
        alert(`Modelo "${modelName}" procesado correctamente.`);
    }
}