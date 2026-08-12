/**
 * Lógica del Modelo Semántico BigQuery EPM para Office Add-in
 */
let currentModel = "";
let creatingModel = false;
let fieldsState = [];
let currentConfigFieldIndex = null;
let currentTreeTarget = "FACT"; // "FACT" o "DIM"

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        initEvents();
		await ensureCoreModelSheets();
		loadSemanticModels();
    }
});

/**
 * Asegura que existan las hojas técnicas base del modelo semántico:
 * - MODEL_FACT: si no existe la crea con su cabecera.
 * - EDIT_REPORT: si no existe la crea vacía, salvo D5 y D6 que llevan "X".
 * Ambas se ocultan (igual que el resto de hojas MODEL_*).
 */
async function ensureCoreModelSheets() {

    await Excel.run(async (context) => {

        const sheets = context.workbook.worksheets;

        // MODEL_FACT
        let factSheet = sheets.getItemOrNullObject("MODEL_FACT");
        await context.sync();

        if (factSheet.isNullObject) {

            factSheet = sheets.add("MODEL_FACT");

            factSheet.getRangeByIndexes(0, 0, 1, 4).values = [
                ["MODEL_NAME", "FACT_PROJECT", "FACT_DATASET", "FACT_TABLE"]
            ];

        }

        factSheet.visibility = Excel.SheetVisibility.hidden;

        // EDIT_REPORT
        let editSheet = sheets.getItemOrNullObject("EDIT_REPORT");
        await context.sync();

        if (editSheet.isNullObject) {

            editSheet = sheets.add("EDIT_REPORT");

            editSheet.getRange("D5").values = [["X"]];
            editSheet.getRange("D6").values = [["X"]];

        }

        editSheet.visibility = Excel.SheetVisibility.hidden;

        await context.sync();

    });

}

function initEvents() {

    document.getElementById("btnNewModel").addEventListener("click", newModel);

    document.getElementById("btnDeleteModel").addEventListener("click", deleteModel);

    document.getElementById("semanticModelSelect").addEventListener("change", function(){

        if(this.value==="")
        {
            clearCurrentModel();
            return;
        }

        loadModel(this.value);

    });

    document.getElementById("sapSearchBtn").addEventListener("click", () => openTreeModal("FACT"));
    document.getElementById("factFullConcat").addEventListener("click", () => openTreeModal("FACT"));

    document.getElementById("dimSapSearchBtn").addEventListener("click", () => openTreeModal("DIM"));
    document.getElementById("dimRelFullConcat").addEventListener("click", () => openTreeModal("DIM"));

    document.getElementById("btnCloseTreeModal").addEventListener("click", closeTreeModal);

    document.getElementById("btnSaveMeasureModal").addEventListener("click", saveMeasureModal);

    document.getElementById("btnCloseMeasureModal").addEventListener("click",()=>{

        measureModal.style.display="none";

    });

    document.getElementById("btnSaveDimModal").addEventListener("click",saveDimModal);

    document.getElementById("btnCloseDimModal").addEventListener("click",()=>{

        dimensionModal.style.display="none";

    });

    document.getElementById("btnAddHierarchy").addEventListener("click", () => openHierarchyEditor(null));

    document.getElementById("btnCloseHierarchyModal").addEventListener("click", closeHierarchyEditor);

    document.getElementById("btnCancelHierarchyModal").addEventListener("click", closeHierarchyEditor);

    document.getElementById("btnSaveHierarchyModal").addEventListener("click", saveHierarchyEditor);

    document.getElementById("btnPreviewHierarchy").addEventListener("click", previewCurrentHierarchy);

    document.getElementById("btnViewFactData").addEventListener("click", () => {

        openDataPreviewDialog(
            document.getElementById("factProject").value,
            document.getElementById("factDataset").value,
            document.getElementById("factTable").value
        );

    });

    document.getElementById("btnViewDimData").addEventListener("click", () => {

        openDataPreviewDialog(
            document.getElementById("dimRelProject").value,
            document.getElementById("dimRelDataset").value,
            document.getElementById("dimRelTable").value
        );

    });

    document.getElementById("btnGenerateModel").addEventListener("click",generateSemanticModelInExcel);

}

/**
 * Abre un diálogo de Office (Office.context.ui.displayDialogAsync) centrado
 * sobre la ventana de Excel ("en medio del Excel") con un SELECT * LIMIT 500
 * de la tabla indicada. Es una página independiente (dataPreview.html) que
 * no necesita el modelo de objetos de Excel, solo el token de BigQuery.
 */
function openDataPreviewDialog(project, dataset, table) {

    if (!project || !dataset || !table) {
        alert("Selecciona primero una tabla para poder previsualizar sus datos.");
        return;
    }

    const relativeUrl =
        `dataPreview.html?project=${encodeURIComponent(project)}` +
        `&dataset=${encodeURIComponent(dataset)}` +
        `&table=${encodeURIComponent(table)}`;

    const dialogUrl = new URL(relativeUrl, window.location.href).href;

    Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 70, width: 60, displayInIframe: false },
        (asyncResult) => {

            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Error al abrir la vista previa de datos:", asyncResult.error);
                alert("No se ha podido abrir la vista previa: " + asyncResult.error.message);
            }

        }
    );

}

/**
 * Abre un diálogo centrado sobre Excel mostrando cómo se vería la jerarquía
 * que se está editando en el popup (árbol de valores reales agrupados por
 * nivel), sin necesidad de guardarla antes.
 */
function previewCurrentHierarchy() {

    const field = fieldsState[currentConfigFieldIndex];

    if (!field || !hierEditState) return;

    if (!hierEditState.levels || hierEditState.levels.length === 0) {
        alert("Añade al menos un nivel a la jerarquía para poder previsualizarla.");
        return;
    }

    if (!field.relProject || !field.relDataset || !field.relTable) {
        alert("Selecciona primero la tabla de dimensión.");
        return;
    }

    const levels = hierEditState.levels.map(l => {
        const attr = (field.attributes || []).find(a => a.name === l.attribute);
        return {
            field: l.attribute,
            label: attr ? (attr.alias || attr.name) : l.attribute
        };
    });

    const spec = {
        project: field.relProject,
        dataset: field.relDataset,
        table: field.relTable,
        hierarchyName: document.getElementById("hierarchyNameInput").value.trim() || "(sin nombre)",
        levels: levels
    };

    const relativeUrl = `hierarchyPreview.html?spec=${encodeURIComponent(JSON.stringify(spec))}`;
    const dialogUrl = new URL(relativeUrl, window.location.href).href;

    Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 70, width: 50, displayInIframe: false },
        (asyncResult) => {

            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Error al abrir la vista previa de la jerarquía:", asyncResult.error);
                alert("No se ha podido abrir la vista previa: " + asyncResult.error.message);
            }

        }
    );

}

function newModel()
{

    if(creatingModel)
        return;

    creatingModel=true;

    const container=document.getElementById("modelSelectorContainer");

    container.innerHTML=`

        <input
            id="newModelName"
            type="text"
            placeholder="Nombre del modelo"
            autofocus
        >

    `;

    const input=document.getElementById("newModelName");

    input.addEventListener("keydown",function(e){

        if(e.key==="Enter")
            saveNewModel();

        if(e.key==="Escape")
        {
            creatingModel=false;
            restoreModelSelector();
        }

    });

}

async function saveNewModel()
{

    const input = document.getElementById("newModelName");

    const name = input ? input.value.trim() : "";

    creatingModel = false;

    restoreModelSelector();

    if (name === "")
        return;

    currentModel = name;

    if (typeof Excel !== "undefined") {

        try {

            await saveModelHeader();

            await loadSemanticModels();

            const select = document.getElementById("semanticModelSelect");

            if (select) {
                select.value = name;
            }

        } catch (err) {

            console.error("Error al guardar el modelo en Excel:", err);

        }

    }

}


function restoreModelSelector()
{

    document.getElementById("modelSelectorContainer").innerHTML=`

        <select id="semanticModelSelect">

            <option value="">— Sin modelo seleccionado —</option>

        </select>

    `;

    document.getElementById("semanticModelSelect")
        .addEventListener("change",function(){

            if(this.value==="")
            {
                clearCurrentModel();
                return;
            }

            loadModel(this.value);

        });

}

async function loadSemanticModels()
{

    await Excel.run(async(context)=>{

        let sheet=context.workbook.worksheets.getItem("MODEL_FACT");

        let range=sheet.getUsedRange();

        range.load("values");

        await context.sync();

        const rows=range.values;

        const select=document.getElementById("semanticModelSelect");

        select.innerHTML='<option value="">— Sin modelo seleccionado —</option>';

        const models=[];

        for(let i=1;i<rows.length;i++)
        {
            models.push(rows[i][0]);
        }

        models.sort();

        models.forEach(model=>{

            const option=document.createElement("option");

            option.value=model;

            option.text=model;

            select.appendChild(option);

        });

    });

}

async function saveModelHeader()
{

    await ensureCoreModelSheets();

    await Excel.run(async(context)=>{

        let sheet=context.workbook.worksheets.getItem("MODEL_FACT");

        let range=sheet.getUsedRangeOrNullObject();

        await context.sync();

        let rows=[];

        if(!range.isNullObject)
        {
            range.load("values");
            await context.sync();
            rows=range.values || [];
        }

        let modelName = "";
		
		
	
    const selectElem = document.getElementById("semanticModelSelect");
    const inputElem = document.getElementById("newModelName");

    if (selectElem && selectElem.value && selectElem.value !== "") {
        modelName = selectElem.value.trim();
    } else if (inputElem && inputElem.value && inputElem.value.trim() !== "") {
        modelName = inputElem.value.trim();
    } else {
        modelName = currentModel;
    }
		
		currentModel = modelName;

        rows=rows.filter((r,i)=>{

            if(i===0)
                return true;

            return r[0]!==currentModel;

        });

        rows.splice(1, 0, [

            currentModel,

            document.getElementById("factProject").value,

            document.getElementById("factDataset").value,

            document.getElementById("factTable").value

        ]);

        sheet.getUsedRange().clear();

        sheet.getRangeByIndexes(0,0,rows.length,4).values=rows;

        await context.sync();

    });

}

async function loadModel(modelName)
{

    currentModel=modelName;

    await Excel.run(async(context)=>{

        const sheet=context.workbook.worksheets.getItem("MODEL_FACT");

        const range=sheet.getUsedRange();

        range.load("values");

        await context.sync();

        const rows=range.values;

        for(let i=1;i<rows.length;i++)
        {

            if(rows[i][0]===modelName)
            {

                document.getElementById("factProject").value=rows[i][1];

                document.getElementById("factDataset").value=rows[i][2];

                document.getElementById("factTable").value=rows[i][3];

                document.getElementById("factFullConcat").value=

                    rows[i][1]+"."+rows[i][2]+"."+rows[i][3];

                break;

            }

        }

    });

    await fetchFactFields(true);

}

async function deleteModel()
{

    let modelName = "";
    const selectElem = document.getElementById("semanticModelSelect");
    const inputElem = document.getElementById("newModelName");

    if (selectElem && selectElem.value && selectElem.value !== "") {
        modelName = selectElem.value.trim();
    } else if (inputElem && inputElem.value && inputElem.value.trim() !== "") {
        modelName = inputElem.value.trim();
    } else {
        modelName = currentModel;
    }
    currentModel = modelName;

    if(currentModel==="")
        return;

    await Excel.run(async(context)=>{

        const targets = [
            { name: "MODEL_FACT", modelCol: 0 },
            { name: "MODEL_DIMENSION", modelCol: 10 },
            { name: "MODEL_MEASURES", modelCol: 8 },
            { name: "MODEL_RELATIONSHIP", modelCol: -1 },
            { name: "MODEL_ATRIBUTES", modelCol: -1 },
            { name: "MODEL_HIER", modelCol: -1 }
        ];

        for (const target of targets) {

            let sheet = context.workbook.worksheets.getItemOrNullObject(target.name);

            await context.sync();

            if (!sheet.isNullObject) {

                let range = sheet.getUsedRangeOrNullObject();

                await context.sync();

                if (!range.isNullObject) {

                    range.load("values");

                    await context.sync();

                    let rows = range.values;

                    if (rows && rows.length > 0) {

                        rows = rows.filter((r, i) => {

                            if (i === 0) return true;

                            // modelCol === -1 significa "el nombre del modelo está en la última columna"
                            const col = target.modelCol === -1 ? r.length - 1 : target.modelCol;

                            return r[col] !== currentModel && r[r.length - 1] !== currentModel;

                        });

                        sheet.getUsedRange().clear();

                        if (rows.length > 0) {

                            sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length).values = rows;

                        }

                    }

                }

            }

        }

        await context.sync();

    });

    creatingModel = false;

    restoreModelSelector();

    clearCurrentModel();

    await loadSemanticModels();

}

function clearCurrentModel()
{

    currentModel="";

    fieldsState=[];

    document.getElementById("factProject").value="";

    document.getElementById("factDataset").value="";

    document.getElementById("factTable").value="";

    document.getElementById("factFullConcat").value="";

    document.getElementById("fieldsList").innerHTML="";

    document.getElementById("fieldsCard").style.display="none";

    document.getElementById("semanticModelSelect").value="";

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
        fetchFactFields(false);
    } else if (currentTreeTarget === "DIM") {
        document.getElementById("dimRelProject").value = projectId;
        document.getElementById("dimRelDataset").value = datasetId;
        document.getElementById("dimRelTable").value = tableId;
        document.getElementById("dimRelFullConcat").value = `${projectId}.${datasetId}.${tableId}`;
        closeTreeModal();
        fetchDimensionAttributes(true);
    }
}

async function fetchFactFields(isModelLoad = false) {
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
            let savedDimFields = new Set();
            let savedMeaFields = new Set();
            let relMap = {};    // DIMENSION -> { relProject, relDataset, relTable }
            let attrMap = {};   // DIMENSION -> [ {name, alias, dataType, isKey, enabled} ]
            let hierMap = {};   // DIMENSION -> [ {name, levels:[{attribute}]} ]

            if (isModelLoad && typeof Excel !== "undefined" && currentModel) {
                try {
                    await Excel.run(async (context) => {
                        const sheets = context.workbook.worksheets;

                        let sheetDim = sheets.getItemOrNullObject("MODEL_DIMENSION");
                        await context.sync();
                        if (!sheetDim.isNullObject) {
                            let rangeDim = sheetDim.getUsedRangeOrNullObject();
                            await context.sync();
                            if (!rangeDim.isNullObject) {
                                rangeDim.load("values");
                                await context.sync();
                                const rows = rangeDim.values || [];
                                for (let i = 1; i < rows.length; i++) {
                                    const r = rows[i];
                                    if (r[10] === currentModel || r[r.length - 1] === currentModel) {
                                        if (r[5]) savedDimFields.add(r[5]);
                                    }
                                }
                            }
                        }

                        let sheetMea = sheets.getItemOrNullObject("MODEL_MEASURES");
                        await context.sync();
                        if (!sheetMea.isNullObject) {
                            let rangeMea = sheetMea.getUsedRangeOrNullObject();
                            await context.sync();
                            if (!rangeMea.isNullObject) {
                                rangeMea.load("values");
                                await context.sync();
                                const rows = rangeMea.values || [];
                                for (let i = 1; i < rows.length; i++) {
                                    const r = rows[i];
                                    if (r[8] === currentModel || r[r.length - 1] === currentModel) {
                                        if (r[5]) savedMeaFields.add(r[5]);
                                    }
                                }
                            }
                        }

                        // MODEL_RELATIONSHIP -> relación de cada dimensión con su tabla
                        let sheetRel = sheets.getItemOrNullObject("MODEL_RELATIONSHIP");
                        await context.sync();
                        if (!sheetRel.isNullObject) {
                            let rangeRel = sheetRel.getUsedRangeOrNullObject();
                            await context.sync();
                            if (!rangeRel.isNullObject) {
                                rangeRel.load("values");
                                await context.sync();
                                const rows = rangeRel.values || [];
                                for (let i = 1; i < rows.length; i++) {
                                    const r = rows[i];
                                    if (r[r.length - 1] === currentModel) {
                                        const dimName = r[1];
                                        if (dimName) {
                                            relMap[dimName] = {
                                                relProject: r[6] || "",
                                                relDataset: r[7] || "",
                                                relTable: r[8] || ""
                                            };
                                        }
                                    }
                                }
                            }
                        }

                        // MODEL_ATRIBUTES -> atributos de cada dimensión
                        let sheetAttr = sheets.getItemOrNullObject("MODEL_ATRIBUTES");
                        await context.sync();
                        if (!sheetAttr.isNullObject) {
                            let rangeAttr = sheetAttr.getUsedRangeOrNullObject();
                            await context.sync();
                            if (!rangeAttr.isNullObject) {
                                rangeAttr.load("values");
                                await context.sync();
                                const rows = rangeAttr.values || [];
                                for (let i = 1; i < rows.length; i++) {
                                    const r = rows[i];
                                    if (r[r.length - 1] === currentModel) {
                                        const dimName = r[1];
                                        const attrName = r[2];
                                        if (dimName && attrName) {
                                            if (!attrMap[dimName]) attrMap[dimName] = [];
                                            attrMap[dimName].push({
                                                name: attrName,
                                                alias: (r[7] && String(r[7]).trim() !== "") ? r[7] : attrName,
                                                dataType: r[8] || "",
                                                isKey: r[9] === "X" || r[9] === true,
                                                enabled: true
                                            });
                                        }
                                    }
                                }
                            }
                        }

                        // MODEL_HIER -> jerarquías de cada dimensión
                        let sheetHier = sheets.getItemOrNullObject("MODEL_HIER");
                        await context.sync();
                        if (!sheetHier.isNullObject) {
                            let rangeHier = sheetHier.getUsedRangeOrNullObject();
                            await context.sync();
                            if (!rangeHier.isNullObject) {
                                rangeHier.load("values");
                                await context.sync();
                                const rows = rangeHier.values || [];
                                const rawMap = {}; // dimName -> hierName -> [{nivel, attribute}]
                                for (let i = 1; i < rows.length; i++) {
                                    const r = rows[i];
                                    if (r[r.length - 1] === currentModel) {
                                        const hierName = r[1];
                                        const nivel = Number(r[2]);
                                        const dimName = r[3];
                                        const fieldName = r[8];
                                        if (dimName && hierName && fieldName) {
                                            if (!rawMap[dimName]) rawMap[dimName] = {};
                                            if (!rawMap[dimName][hierName]) rawMap[dimName][hierName] = [];
                                            rawMap[dimName][hierName].push({ nivel: nivel, attribute: fieldName });
                                        }
                                    }
                                }
                                Object.keys(rawMap).forEach(dimName => {
                                    hierMap[dimName] = Object.keys(rawMap[dimName]).map(hierName => {
                                        const levels = rawMap[dimName][hierName]
                                            .sort((a, b) => a.nivel - b.nivel)
                                            .map(l => ({ attribute: l.attribute }));
                                        return { name: hierName, levels: levels };
                                    });
                                });
                            }
                        }
                    });
                } catch (err) {
                    console.error("Error al consultar dimensiones, relaciones, atributos y jerarquías guardadas en Excel:", err);
                }
            }

            fieldsState = data.schema.fields.map(f => {
                const isNumeric = ["INTEGER", "FLOAT", "NUMERIC", "BIGNUMERIC"].includes(f.type);
                let fieldType = isNumeric ? "MEASURE" : "DIMENSION";
                let isEnabled = true;

                if (isModelLoad) {
                    if (savedDimFields.has(f.name)) {
                        fieldType = "DIMENSION";
                        isEnabled = true;
                    } else if (savedMeaFields.has(f.name)) {
                        fieldType = "MEASURE";
                        isEnabled = true;
                    } else {
                        isEnabled = false;
                    }
                }

                const field = {
                    name: f.name,
                    alias: f.name,
                    dataType: f.type,
                    type: fieldType,
                    enabled: isEnabled,
                    // Config Medida
                    aggregation: "SUM",
                    format: "Auto",
                    // Config Dimensión (Relación)
                    relProject: "",
                    relDataset: "",
                    relTable: "",
                    attributes: [],
                    hierarchies: []
                };

                if (isModelLoad && fieldType === "DIMENSION" && isEnabled) {
                    if (relMap[f.name]) {
                        field.relProject = relMap[f.name].relProject;
                        field.relDataset = relMap[f.name].relDataset;
                        field.relTable = relMap[f.name].relTable;
                    }
                    if (attrMap[f.name]) {
                        field.attributes = attrMap[f.name];
                    }
                    if (hierMap[f.name]) {
                        field.hierarchies = hierMap[f.name];
                    }
                }

                return field;
            });

            renderFieldsTable();
            document.getElementById("fieldsCard").style.display = "block";
        }
    } catch (err) {
        console.error("Error al obtener esquema:", err);
    }
}

function renderFieldsTable() {

    const list = document.getElementById("fieldsList");
    list.innerHTML = "";

    fieldsState.forEach((field, idx) => {

        const row = document.createElement("div");
        row.className = "field-row";

row.innerHTML = `

    <input
        class="field-enable"
        type="checkbox"
        ${field.enabled ? "checked" : ""}
        onclick="event.stopPropagation(); updateFieldEnabled(${idx}, this.checked)"
    />

    <select
        class="field-type"
        onclick="event.stopPropagation()"
        onchange="updateFieldType(${idx}, this.value)"
    >
        <option value="MEASURE" ${field.type === "MEASURE" ? "selected" : ""}>MEA</option>
        <option value="DIMENSION" ${field.type === "DIMENSION" ? "selected" : ""}>DIM</option>
    </select>

    <div class="field-info">

        <input
            class="field-alias-input"
            type="text"
            value="${field.alias}"
            onclick="event.stopPropagation()"
            oninput="updateFieldAlias(${idx}, this.value)"
        />

        <div class="field-name">
            ${field.name}
        </div>

    </div>

    <div class="field-arrow">›</div>

`;

        row.onclick = () => openConfigModal(idx);

        list.appendChild(row);

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

        renderHierarchiesList();

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
        renderHierarchiesList();
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
                    enabled: true
                };
            });

            // Si la tabla de dimensión ha cambiado, limpiamos de las jerarquías
            // cualquier nivel que referencie atributos que ya no existen.
            if (field.hierarchies && field.hierarchies.length > 0) {
                const validNames = new Set(field.attributes.map(a => a.name));
                field.hierarchies = field.hierarchies
                    .map(h => ({ name: h.name, levels: h.levels.filter(l => validNames.has(l.attribute)) }))
                    .filter(h => h.levels.length > 0);
            }

            renderAttributesTable(field.attributes);
            document.getElementById("attributesContainer").style.display = "block";
            renderHierarchiesList();
        }
    } catch (err) {
        console.error("Error al obtener atributos de la dimensión:", err);
    }
}


function renderAttributesTable(attributes) {

    const list = document.getElementById("attributesList");
    list.innerHTML = "";

    attributes.forEach((attr, idx) => {

        const row = document.createElement("div");
        row.className = "field-row";

        row.innerHTML = `

            <input
                class="field-enable"
                type="checkbox"
                ${attr.enabled ? "checked" : ""}
                onclick="event.stopPropagation(); updateAttrEnabled(${idx}, this.checked)"
            />

            <div class="field-info">

                <input
                    class="field-alias-input"
                    type="text"
                    value="${attr.alias}"
                    onclick="event.stopPropagation()"
                    oninput="updateAttrAlias(${idx}, this.value)"
                />

                <div class="field-name">
                    ${attr.name}
                </div>

            </div>

            <label class="field-key">

                <input
                    type="radio"
                    name="dimKeyGroup"
                    ${attr.isKey ? "checked" : ""}
                    onclick="event.stopPropagation()"
                    onchange="updateAttrKey(${idx})"
                >

                <span>Key</span>

            </label>

        `;

        list.appendChild(row);

    });

}




function updateAttrAlias(idx, val) { fieldsState[currentConfigFieldIndex].attributes[idx].alias = val; }
function updateAttrKey(selectedIdx) {
    fieldsState[currentConfigFieldIndex].attributes.forEach((attr, idx) => {
        attr.isKey = (idx === selectedIdx);
    });
}
function updateAttrEnabled(idx, val) { fieldsState[currentConfigFieldIndex].attributes[idx].enabled = val; }

function saveDimModal() {
    if (currentConfigFieldIndex !== null) {
        fieldsState[currentConfigFieldIndex].alias = document.getElementById("modalDimAlias").value;
        renderFieldsTable();
    }
    document.getElementById("dimensionModal").style.display = "none";
}

/* =====================================================================
 * DISEÑO DE JERARQUÍAS
 *
 * Cada dimensión (field) guarda su lista de jerarquías en
 * field.hierarchies = [{ name, levels: [{ attribute }] }]
 * donde "levels" está ordenado de nivel superior (índice 0) a inferior.
 * ===================================================================== */

let hierEditState = null;      // { name, levels: [{attribute}] } — estado en edición dentro del popup
let editingHierarchyIndex = null; // null = nueva jerarquía; número = editando field.hierarchies[idx]

function renderHierarchiesList() {

    const field = fieldsState[currentConfigFieldIndex];
    const container = document.getElementById("hierarchiesList");

    if (!field || !container) return;

    container.innerHTML = "";

    const hierarchies = field.hierarchies || [];

    if (hierarchies.length === 0) {
        container.innerHTML = `<div class="hierarchy-empty">No hay jerarquías definidas para esta dimensión.</div>`;
        return;
    }

    hierarchies.forEach((hier, idx) => {

        const levelNames = (hier.levels || []).map(l => {
            const attr = (field.attributes || []).find(a => a.name === l.attribute);
            return attr ? (attr.alias || attr.name) : l.attribute;
        }).join(" › ");

        const row = document.createElement("div");
        row.className = "hierarchy-row";

        row.innerHTML = `
            <div class="hierarchy-info">
                <div class="hierarchy-name">${hier.name}</div>
                <div class="hierarchy-levels-summary">${levelNames || "(sin niveles)"}</div>
            </div>
            <div class="hierarchy-chip-actions">
                <button type="button" class="hierarchy-chip-btn" title="Editar jerarquía" onclick="event.stopPropagation(); openHierarchyEditor(${idx})">✎</button>
                <button type="button" class="hierarchy-chip-btn" title="Eliminar jerarquía" onclick="event.stopPropagation(); deleteHierarchyAt(${idx})">🗑</button>
            </div>
        `;

        container.appendChild(row);

    });

}

function deleteHierarchyAt(idx) {

    const field = fieldsState[currentConfigFieldIndex];

    if (!field || !field.hierarchies || !field.hierarchies[idx]) return;

    if (!confirm(`¿Eliminar la jerarquía "${field.hierarchies[idx].name}"?`)) return;

    field.hierarchies.splice(idx, 1);

    renderHierarchiesList();

}

function openHierarchyEditor(index = null) {

    const field = fieldsState[currentConfigFieldIndex];

    if (!field) return;

    if (!field.attributes || field.attributes.length === 0) {
        alert("Selecciona primero la tabla de dimensión y sus atributos antes de crear una jerarquía.");
        return;
    }

    editingHierarchyIndex = index;

    if (index === null) {
        hierEditState = { name: "", levels: [] };
        document.getElementById("hierarchyModalTitle").textContent = "Nueva jerarquía";
    } else {
        const existing = field.hierarchies[index];
        hierEditState = {
            name: existing.name,
            levels: existing.levels.map(l => ({ attribute: l.attribute }))
        };
        document.getElementById("hierarchyModalTitle").textContent = "Editar jerarquía";
    }

    document.getElementById("hierarchyNameInput").value = hierEditState.name;

    renderHierarchyEditor();

    document.getElementById("hierarchyModal").style.display = "block";

}

function closeHierarchyEditor() {

    document.getElementById("hierarchyModal").style.display = "none";

    hierEditState = null;
    editingHierarchyIndex = null;

}

function renderHierarchyEditor() {

    const field = fieldsState[currentConfigFieldIndex];

    if (!field || !hierEditState) return;

    const poolContainer = document.getElementById("hierarchyPoolList");
    const levelsContainer = document.getElementById("hierarchyLevelsList");

    const usedNames = new Set(hierEditState.levels.map(l => l.attribute));
    const availableAttrs = (field.attributes || []).filter(a => a.enabled !== false && !usedNames.has(a.name));

    // ---- Columna izquierda: atributos disponibles (arrastrables) ----
    poolContainer.innerHTML = "";

    if (availableAttrs.length === 0) {
        poolContainer.innerHTML = `<div class="hierarchy-pool-empty">Todos los atributos están ya en la jerarquía.</div>`;
    } else {

        availableAttrs.forEach(attr => {

            const chip = document.createElement("div");
            chip.className = "hierarchy-attr-chip";
            chip.draggable = true;
            chip.title = "Arrastra a la derecha o haz clic para añadir";

            chip.innerHTML = `
                <span class="hierarchy-level-name">${attr.alias || attr.name}</span>
                <span class="hierarchy-chip-btn">＋</span>
            `;

            chip.addEventListener("dragstart", (e) => {
                chip.classList.add("dragging");
                e.dataTransfer.setData("text/plain", attr.name);
                e.dataTransfer.effectAllowed = "copyMove";
            });

            chip.addEventListener("dragend", () => chip.classList.remove("dragging"));

            chip.addEventListener("click", () => addLevelFromPool(attr.name));

            poolContainer.appendChild(chip);

        });

    }

    // ---- Columna derecha: niveles de la jerarquía (ordenados, arrastrables) ----
    levelsContainer.innerHTML = "";

    if (hierEditState.levels.length === 0) {
        levelsContainer.innerHTML = `<div class="hierarchy-levels-empty">Arrastra aquí los atributos, del nivel superior (arriba) al inferior (abajo).</div>`;
    } else {

        hierEditState.levels.forEach((lvl, idx) => {

            const attr = (field.attributes || []).find(a => a.name === lvl.attribute);

            const chip = document.createElement("div");
            chip.className = "hierarchy-level-chip";
            chip.draggable = true;
            chip.dataset.index = idx;

            chip.innerHTML = `
                <span class="hierarchy-level-name">${attr ? (attr.alias || attr.name) : lvl.attribute}</span>
                <button type="button" class="hierarchy-chip-btn hierarchy-chip-btn-remove" title="Quitar del nivel" onclick="event.stopPropagation(); removeHierarchyLevel(${idx})">✕</button>
            `;

            chip.addEventListener("dragstart", (e) => {
                chip.classList.add("dragging");
                e.dataTransfer.setData("text/plain", "LEVEL:" + idx);
                e.dataTransfer.effectAllowed = "move";
            });

            chip.addEventListener("dragend", () => chip.classList.remove("dragging"));

            levelsContainer.appendChild(chip);

        });

    }

    // ---- Zona de destino: niveles (añadir nuevos o reordenar) ----
    levelsContainer.ondragover = (e) => {
        e.preventDefault();
        levelsContainer.classList.add("drag-over");
    };

    levelsContainer.ondragleave = () => levelsContainer.classList.remove("drag-over");

    levelsContainer.ondrop = (e) => {

        e.preventDefault();
        levelsContainer.classList.remove("drag-over");

        const data = e.dataTransfer.getData("text/plain");

        if (!data) return;

        const dropIndex = getLevelDropIndex(levelsContainer, e.clientY);

        if (data.indexOf("LEVEL:") === 0) {
            const fromIdx = parseInt(data.split(":")[1], 10);
            reorderHierarchyLevel(fromIdx, dropIndex);
        } else {
            addLevelFromPool(data, dropIndex);
        }

    };

    // ---- Zona de destino: pool (soltar aquí un nivel lo quita de la jerarquía) ----
    poolContainer.ondragover = (e) => {
        e.preventDefault();
        poolContainer.classList.add("drag-over");
    };

    poolContainer.ondragleave = () => poolContainer.classList.remove("drag-over");

    poolContainer.ondrop = (e) => {

        e.preventDefault();
        poolContainer.classList.remove("drag-over");

        const data = e.dataTransfer.getData("text/plain");

        if (data && data.indexOf("LEVEL:") === 0) {
            const fromIdx = parseInt(data.split(":")[1], 10);
            removeHierarchyLevel(fromIdx);
        }

    };

}

function getLevelDropIndex(container, clientY) {

    const chips = Array.from(container.querySelectorAll(".hierarchy-level-chip"));

    for (let i = 0; i < chips.length; i++) {
        const rect = chips[i].getBoundingClientRect();
        if (clientY < rect.top + rect.height / 2) {
            return i;
        }
    }

    return chips.length;

}

function addLevelFromPool(attrName, atIndex = null) {

    if (!hierEditState) return;

    if (hierEditState.levels.some(l => l.attribute === attrName)) return;

    if (atIndex === null || atIndex < 0 || atIndex > hierEditState.levels.length) {
        hierEditState.levels.push({ attribute: attrName });
    } else {
        hierEditState.levels.splice(atIndex, 0, { attribute: attrName });
    }

    renderHierarchyEditor();

}

function removeHierarchyLevel(idx) {

    if (!hierEditState) return;

    hierEditState.levels.splice(idx, 1);

    renderHierarchyEditor();

}

function reorderHierarchyLevel(fromIdx, toIdx) {

    if (!hierEditState) return;
    if (fromIdx === toIdx) return;

    const levels = hierEditState.levels;
    const [moved] = levels.splice(fromIdx, 1);

    let insertAt = toIdx;
    if (fromIdx < toIdx) insertAt -= 1;

    levels.splice(insertAt, 0, moved);

    renderHierarchyEditor();

}

function saveHierarchyEditor() {

    const field = fieldsState[currentConfigFieldIndex];

    if (!field || !hierEditState) return;

    const name = document.getElementById("hierarchyNameInput").value.trim();

    if (name === "") {
        alert("Indica un nombre para la jerarquía.");
        return;
    }

    if (hierEditState.levels.length === 0) {
        alert("Añade al menos un nivel a la jerarquía.");
        return;
    }

    if (!field.hierarchies) field.hierarchies = [];

    const duplicate = field.hierarchies.some((h, idx) =>
        h.name.toUpperCase() === name.toUpperCase() && idx !== editingHierarchyIndex
    );

    if (duplicate) {
        alert("Ya existe una jerarquía con ese nombre en esta dimensión.");
        return;
    }

    const hierObj = {
        name: name,
        levels: hierEditState.levels.map(l => ({ attribute: l.attribute }))
    };

    if (editingHierarchyIndex === null) {
        field.hierarchies.push(hierObj);
    } else {
        field.hierarchies[editingHierarchyIndex] = hierObj;
    }

    closeHierarchyEditor();
    renderHierarchiesList();

}

/**
 * Vuelca el modelo semántico completo a Excel rellenando las 5 pestañas objetivo
 */
 
 
 
/**
 * Fusiona en una hoja (sheetName) las filas nuevas (newRows) de un modelo,
 * conservando intactas las filas de los DEMÁS modelos que ya hubiera en la
 * hoja. Es la misma lógica que ya usaban MODEL_DIMENSION / MODEL_MEASURES,
 * generalizada para aplicarla también a MODEL_RELATIONSHIP, MODEL_ATRIBUTES
 * y MODEL_HIER, de forma que TODAS las pestañas del modelo semántico
 * guarden y respeten el nombre del modelo (MODEL_NAME, última columna).
 */
async function mergeModelSheet(context, sheets, sheetName, headers, newRows, modelName) {

    const modelNameColIdx = headers.length - 1; // MODEL_NAME siempre es la última columna

    let sheet = sheets.getItemOrNullObject(sheetName);
    await context.sync();

    if (sheet.isNullObject) {
        sheet = sheets.add(sheetName);
    }

    let range = sheet.getUsedRangeOrNullObject();
    await context.sync();

    let existingRows = [];

    if (!range.isNullObject) {
        range.load("values");
        await context.sync();
        existingRows = range.values || [];
    }

    let finalRows;

    if (existingRows.length > 0) {

        finalRows = existingRows.filter((r, i) => {

            if (i === 0) return true; // conservar cabecera existente

            return r[modelNameColIdx] !== modelName;

        });

    } else {

        finalRows = [headers];

    }

    newRows.forEach(row => finalRows.push(row));

    sheet.getUsedRangeOrNullObject().clear();

    if (finalRows.length > 0) {
        sheet.getRangeByIndexes(0, 0, finalRows.length, finalRows[0].length).values = finalRows;
    }

    // Todas las hojas técnicas del modelo semántico se ocultan
    sheet.visibility = Excel.SheetVisibility.hidden;

    return sheet;

}

/**
 * Vuelca el modelo semántico completo a Excel rellenando las 5 pestañas objetivo.
 * Todas las pestañas guardan el nombre del modelo (MODEL_NAME) en su última
 * columna y se fusionan con lo que ya hubiera de otros modelos (misma lógica
 * que MODEL_FACT / MODEL_DIMENSION), en vez de sobrescribir toda la hoja.
 */
async function generateSemanticModelInExcel() {
    let modelName = "";
    const selectElem = document.getElementById("semanticModelSelect");
    const inputElem = document.getElementById("newModelName");

    if (selectElem && selectElem.value && selectElem.value !== "") {
        modelName = selectElem.value.trim();
    } else if (inputElem && inputElem.value && inputElem.value.trim() !== "") {
        modelName = inputElem.value.trim();
    } else {
        modelName = currentModel;
    }
    currentModel = modelName;

    if (currentModel === "") {
        alert("Indica primero el nombre del modelo.");
        return;
    }

    const factProject = document.getElementById("factProject").value.trim();
    const factDataset = document.getElementById("factDataset").value;
    const factTable = document.getElementById("factTable").value;

    // 1. MODEL_RELATIONSHIP
    const relHeaders = [
        "FILA",
        "DIMENSION",
        "FACT_PROJECT",
        "FACT_DATASET",
        "FACT_TABLE",
        "FACT_FIELD",
        "DIM_PROJECT",
        "DIM_DATASET",
        "DIM_TABLE",
        "DIM_FIELD",
        "JOIN TYPE",
        "MODEL_NAME"
    ];

    const relNewRows = [];

    fieldsState.forEach((f, idx) => {

        if (!f.enabled || f.type !== "DIMENSION" || !f.relTable)
            return;

        const keyAttr = (f.attributes || []).find(a => a.isKey);

        relNewRows.push([
            idx + 1,
            f.name,
            factProject,
            factDataset,
            factTable,
            f.name,
            f.relProject,
            f.relDataset,
            f.relTable,
            keyAttr ? keyAttr.name : f.name,
            "LEFT",
            currentModel
        ]);

    });

    // 2. MODEL_DIMENSION
    const dimHeaders = [
        "FILA",
        "DIMENSION",
        "FACT_PROJECT",
        "FACT_DATASET",
        "FACT_TABLE",
        "FACT_FIELD",
        "DIM_PROJECT",
        "DIM_DATASET",
        "DIM_TABLE",
        "DIM_FIELD",
        "MODEL_NAME"
    ];

    const dimNewRows = [];

    fieldsState.forEach((f, idx) => {

        if (!f.enabled || f.type !== "DIMENSION") return;

        const keyAttr = (f.attributes || []).find(a => a.isKey);

        dimNewRows.push([
            idx + 1,
            f.name,
            factProject,
            factDataset,
            factTable,
            f.name,
            f.relProject || factProject,
            f.relDataset || factDataset,
            f.relTable || factTable,
            keyAttr ? keyAttr.name : f.name,
            currentModel
        ]);

    });

    // 3. MODEL_MEASURES
    const meaHeaders = [
        "FILA",
        "MEASURE",
        "FACT_PROJECT",
        "FACT_DATASET",
        "FACT_TABLE",
        "FACT_FIELD",
        "AGGREGATION",
        "FORMAT",
        "MODEL_NAME"
    ];

    const meaNewRows = [];

    fieldsState.forEach((f, idx) => {

        if (!f.enabled || f.type !== "MEASURE") return;

        meaNewRows.push([
            idx + 1,
            f.name,
            factProject,
            factDataset,
            factTable,
            f.name,
            f.aggregation,
            f.format,
            currentModel
        ]);

    });

    // 4. MODEL_ATRIBUTES
    const attrHeaders = [
        "FILA",
        "DIMENSION",
        "ATRIBUTE",
        "DIM_PROJECT",
        "DIM_DATASET",
        "DIM_TABLE",
        "DIM_FIELD",
        "DISPLAY_NAME",
        "DATA_TYPE",
        "IS_KEY",
        "MODEL_NAME"
    ];

    const attrNewRows = [];

    fieldsState.forEach((f, idx) => {

        if (!f.enabled || f.type !== "DIMENSION")
            return;

        if (f.attributes && f.attributes.length > 0) {

            f.attributes
                .filter(a => a.enabled)
                .forEach(a => {

                    attrNewRows.push([
                        idx + 1,
                        f.name,
                        a.name,
                        f.relProject || factProject,
                        f.relDataset || factDataset,
                        f.relTable || factTable,
                        a.name,
                        a.alias === a.name ? "" : a.alias,
                        a.dataType,
                        a.isKey ? "X" : "",
                        currentModel
                    ]);

                });

        }
        else {

            attrNewRows.push([
                idx + 1,
                f.name,
                f.name,
                factProject,
                factDataset,
                factTable,
                f.name,
                "",
                f.dataType,
                "X",
                currentModel
            ]);

        }

    });

    // 5. MODEL_HIER
    // Esquema esperado por excelService.js / commands.js:
    // FILA | HIERARCHY | NIVEL | DIMENSION | ATRIBUTO | DIM_PROJECT | DIM_DATASET | DIM_TABLE | DIM_FIELD | MODEL_NAME
    const hierHeaders = [
        "FILA",
        "HIERARCHY",
        "NIVEL",
        "DIMENSION",
        "ATRIBUTO",
        "DIM_PROJECT",
        "DIM_DATASET",
        "DIM_TABLE",
        "DIM_FIELD",
        "MODEL_NAME"
    ];

    const hierNewRows = [];
    let hierFila = 1;

    fieldsState.forEach(f => {

        if (!f.enabled || f.type !== "DIMENSION" || !f.hierarchies || f.hierarchies.length === 0)
            return;

        f.hierarchies.forEach(hier => {

            (hier.levels || []).forEach((lvl, levelIdx) => {

                const attr = (f.attributes || []).find(a => a.name === lvl.attribute);

                hierNewRows.push([
                    hierFila++,
                    hier.name,
                    levelIdx + 1,
                    f.name,
                    attr ? (attr.alias || attr.name) : lvl.attribute,
                    f.relProject || factProject,
                    f.relDataset || factDataset,
                    f.relTable || factTable,
                    lvl.attribute,
                    currentModel
                ]);

            });

        });

    });

    try {
        await saveModelHeader();

        await Excel.run(async (context) => {

            const sheets = context.workbook.worksheets;

            await mergeModelSheet(context, sheets, "MODEL_RELATIONSHIP", relHeaders, relNewRows, currentModel);
            await mergeModelSheet(context, sheets, "MODEL_DIMENSION", dimHeaders, dimNewRows, currentModel);
            await mergeModelSheet(context, sheets, "MODEL_MEASURES", meaHeaders, meaNewRows, currentModel);
            await mergeModelSheet(context, sheets, "MODEL_ATRIBUTES", attrHeaders, attrNewRows, currentModel);
            await mergeModelSheet(context, sheets, "MODEL_HIER", hierHeaders, hierNewRows, currentModel);

            await context.sync();

            alert("¡Modelo Semántico generado con éxito en Excel!");

        });
    } catch (err) {
        console.error("Error al escribir el modelo semántico en Excel:", err);
        alert("Error al escribir en Excel: " + err.message);
    }
}
