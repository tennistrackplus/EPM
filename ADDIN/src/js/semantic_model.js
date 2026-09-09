/**
 * Lógica del Modelo Semántico BigQuery EPM para Office Add-in
 */
let currentModel = "";
let creatingModel = false;
let fieldsState = [];
let currentConfigFieldIndex = null;
let currentTreeTarget = "FACT"; // "FACT" o "DIM"

/**
 * Muestra un aviso visual (toast) en la parte superior de la taskpane,
 * en lugar de un alert() nativo del navegador. type: "success" | "error".
 */
function showToast(message, type = "success", duration = 3500) {

    const container = document.getElementById("appToastContainer");

    if (!container) {
        // Fallback por si el contenedor no existe todavía en el DOM.
        alert(message);
        return;
    }

    const toast = document.createElement("div");

    toast.className = `app-toast ${type === "error" ? "error" : "success"}`;
    toast.textContent = (type === "error" ? "⚠ " : "✔ ") + message;

    container.appendChild(toast);

    // Forzar reflow para que la transición de entrada se aplique.
    requestAnimationFrame(() => toast.classList.add("visible"));

    setTimeout(() => {
        toast.classList.remove("visible");
        setTimeout(() => toast.remove(), 250);
    }, duration);

}

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        initEvents();
        setFactCardVisible(false); // sin modelo seleccionado todavía
		await ensureCoreModelSheets();
		loadSemanticModels();
    }
});

/**
 * Muestra u oculta la caja "Tabla de hechos": solo tiene sentido si hay un
 * modelo semántico seleccionado (o recién creado). Con "— Sin modelo
 * seleccionado —" no debe aparecer.
 */
function setFactCardVisible(visible) {

    const card = document.getElementById("factTableCard");

    if (card) {
        card.style.display = visible ? "" : "none";
    }

}

/**
 * Limpia los campos de la tabla de hechos y la lista de campos en el DOM
 * (usado tanto al deseleccionar el modelo como al crear uno nuevo vacío).
 */
function resetFactAndFieldsUI() {

    document.getElementById("factProject").value = "";
    document.getElementById("factDataset").value = "";
    document.getElementById("factTable").value = "";
    document.getElementById("factFullConcat").value = "";

    fieldsState = [];

    document.getElementById("fieldsList").innerHTML = "";
    document.getElementById("fieldsCard").style.display = "none";

}

/**
 * Asegura que exista la hoja técnica EDIT_REPORT (estado del diseño del
 * informe: filtros/filas/columnas). Se crea vacía si no existe, salvo D5 y
 * D6 que llevan "X", y se oculta.
 *
 * El modelo semántico en sí (antes MODEL_FACT/MODEL_DIMENSION/...) ya NO
 * se guarda en hojas: vive en SemanticModelStore (JSON dentro de las
 * propiedades del documento, sin pestañas ni siquiera ocultas).
 */
async function ensureCoreModelSheets() {

    await Excel.run(async (context) => {

        const sheets = context.workbook.worksheets;

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

    // "Abrir"/"Guardar" ya no muestran un popup dentro del taskpane: abren
    // el mismo diálogo independiente (Office Dialog API) que usan los
    // botones del ribbon "Abrir modelo semántico" / "Guardar modelo
    // semántico" del manifiesto (ver commands.js), así que el
    // comportamiento es idéntico se entre por donde se entre.
    document.getElementById("btnOpenLkmlModel").addEventListener("click", openLkmlDialogFromTaskpane);
    document.getElementById("btnSaveLkmlModel").addEventListener("click", openSaveLkmlDialogFromTaskpane);

}

/**
 * Abre el diálogo independiente "Abrir modelo semántico (.lkml)"
 * (openSemanticModel.html) con Office.context.ui.displayDialogAsync: la
 * misma página que abre directamente el botón del ribbon "Abrir modelo
 * semántico", para que el comportamiento no dependa de este taskpane.
 */
function openLkmlDialogFromTaskpane() {

    window.LkmlOpenBridge.openOpenLkmlDialog((modelName) => {

        // El diálogo ya ha leído el .lkml, generado el modelo semántico y
        // LkmlOpenBridge lo ha guardado en SemanticModelStore: aquí solo
        // hace falta refrescar este taskpane para que se vea seleccionado.
        creatingModel = false;
        restoreModelSelector();
        loadSemanticModels();

        const select = document.getElementById("semanticModelSelect");

        if (select) {
            select.value = modelName;
        }

        loadModel(modelName);

        showToast(`Modelo semántico "${modelName}" importado desde LookML.`, "success");

    });

}

/**
 * Abre el diálogo independiente "Guardar modelo semántico (.lkml)"
 * (saveSemanticModel.html) con Office.context.ui.displayDialogAsync: la
 * misma página que abre directamente el botón del ribbon "Guardar modelo
 * semántico". El diálogo no tiene acceso a SemanticModelStore (no
 * comparte el modelo de objetos del documento), así que aquí, que sí lo
 * tenemos, le pasamos los modelos ya "aplanados" en JSON por querystring.
 */
/**
 * Abre el diálogo independiente "Guardar modelo semántico (.lkml)"
 * (saveSemanticModel.html) con Office.context.ui.displayDialogAsync: la
 * misma página que abre directamente el botón del ribbon "Guardar modelo
 * semántico". La escritura real en EDIT_REPORT!G1 la hace LkmlSaveBridge
 * (ver js/lkmlSaveBridge.js), que sí tiene acceso a Excel.run desde este
 * taskpane, ya que el propio diálogo no tiene acceso al modelo de
 * objetos del documento.
 */
function openSaveLkmlDialogFromTaskpane() {

    const models = window.SemanticModelStore.getAllModels();
    const active = currentModel || window.SemanticModelStore.getActiveModelName();

    window.LkmlSaveBridge.openSaveLkmlDialog(models, active);

}

/**
 * Abre un diálogo de Office (Office.context.ui.displayDialogAsync) centrado
 * sobre la ventana de Excel ("en medio del Excel") con un SELECT * LIMIT 500
 * de la tabla indicada. Es una página independiente (dataPreview.html) que
 * no necesita el modelo de objetos de Excel, solo el token de BigQuery.
 */
function openDataPreviewDialog(project, dataset, table) {

    if (!project || !dataset || !table) {
        showToast("Selecciona primero una tabla para poder previsualizar sus datos.", "error");
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
                showToast("No se ha podido abrir la vista previa: " + asyncResult.error.message, "error");
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
        showToast("Añade al menos un nivel a la jerarquía para poder previsualizarla.", "error");
        return;
    }

    if (!field.relProject || !field.relDataset || !field.relTable) {
        showToast("Selecciona primero la tabla de dimensión.", "error");
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
                showToast("No se ha podido abrir la vista previa: " + asyncResult.error.message, "error");
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

    try {

        await window.SemanticModelStore.saveModel(name, {
            fact: { project: "", dataset: "", table: "" },
            fields: []
        });

        resetFactAndFieldsUI();
        setFactCardVisible(true);

        loadSemanticModels();

        const select = document.getElementById("semanticModelSelect");

        if (select) {
            select.value = name;
        }

    } catch (err) {

        console.error("Error al guardar el modelo:", err);
        showToast("Error al crear el modelo: " + err.message, "error");

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

function loadSemanticModels()
{

    const models = window.SemanticModelStore.listModelNames();

    let select=document.getElementById("semanticModelSelect");

    // Si el selector no existe todavía (p.ej. porque se está mostrando el
    // input de "nuevo modelo" del modelSelectorContainer), lo restauramos
    // primero: evita el "Cannot set properties of null (setting innerHTML)"
    // si esta función se llama en ese momento.
    if (!select) {
        creatingModel = false;
        restoreModelSelector();
        select = document.getElementById("semanticModelSelect");
    }

    if (!select) return;

    select.innerHTML='<option value="">— Sin modelo seleccionado —</option>';

    models.forEach(model=>{

        const option=document.createElement("option");

        option.value=model;

        option.text=model;

        select.appendChild(option);

    });

}

/**
 * Determina el nombre del modelo actual a partir del selector / input de
 * "nuevo modelo" del panel, igual que hacía antes de escribir en MODEL_FACT.
 */
function resolveCurrentModelName()
{

    const selectElem = document.getElementById("semanticModelSelect");
    const inputElem = document.getElementById("newModelName");

    if (selectElem && selectElem.value && selectElem.value !== "") {
        return selectElem.value.trim();
    }
    if (inputElem && inputElem.value && inputElem.value.trim() !== "") {
        return inputElem.value.trim();
    }
    return currentModel;

}

async function loadModel(modelName)
{

    currentModel=modelName;

    setFactCardVisible(true);

    await window.SemanticModelStore.setActiveModelName(modelName);

    const model = window.SemanticModelStore.getModel(modelName);

    if (model && model.fact) {

        document.getElementById("factProject").value=model.fact.project || "";

        document.getElementById("factDataset").value=model.fact.dataset || "";

        document.getElementById("factTable").value=model.fact.table || "";

        document.getElementById("factFullConcat").value=

            (model.fact.project||"")+"."+(model.fact.dataset||"")+"."+(model.fact.table||"");

    }

    await fetchFactFields(true);

}

async function deleteModel()
{

    const modelName = resolveCurrentModelName();

    currentModel = modelName;

    if(currentModel==="")
        return;

    await window.SemanticModelStore.deleteModel(currentModel);

    creatingModel = false;

    restoreModelSelector();

    clearCurrentModel();

    loadSemanticModels();

}

function clearCurrentModel()
{

    currentModel="";

    resetFactAndFieldsUI();

    setFactCardVisible(false);

    const select = document.getElementById("semanticModelSelect");

    if (select) {
        select.value="";
    }

}

function getAuthToken() {
    if (!Provider.isConnected()) {
        showToast(`Sesión de ${Provider.label()} no válida o expirada. Por favor, inicia sesión de nuevo.`, "error");
        return null;
    }
    return true; // solo se usa como comprobación booleana; Provider gestiona el token internamente
}

async function openTreeModal(target = "FACT") {
    currentTreeTarget = target;
    document.getElementById("treeModal").style.display = "block";
    const container = document.getElementById("treeContainer");
    container.innerHTML = `Cargando ${Provider.level1Label()}s...`;

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
        const items = await Provider.listLevel1();
        container.innerHTML = "";

        if (items.length > 0) {
            const ul = document.createElement("ul");
            ul.className = "tree-list";
            for (const p of items) {
                const projectId = p.id;
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
                            childrenDiv.innerHTML = `<div style='margin-left:20px; font-size:11px; color:#666;'>Cargando ${Provider.level2Label()}s...</div>`;
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
                    childrenDiv.innerHTML = `<div style='margin-left:20px; font-size:11px; color:#666;'>Cargando ${Provider.level2Label()}s...</div>`;
                    loaded = true;
                    loadDatasetsTree(projectId, childrenDiv, autoDataset);
                }
            }
            container.appendChild(ul);
        } else {
            container.innerHTML = `No se encontraron elementos de ${Provider.level1Label()}.`;
        }
    } catch (err) {
        console.error("Error al cargar árbol de proyectos:", err);
        container.innerHTML = `Error al cargar ${Provider.level1Label()}s.`;
    }
}

async function loadDatasetsTree(projectId, container, autoDataset = null) {
    const token = getAuthToken();
    if (!token) return;

    try {
        const items = await Provider.listLevel2(projectId);
        container.innerHTML = "";

        if (items.length > 0) {
            const ul = document.createElement("ul");
            ul.className = "tree-list";
            for (const ds of items) {
                const datasetId = ds.id;
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
            container.innerHTML = `<div style='margin-left:20px; font-size:11px; color:#666;'>No hay ${Provider.level2Label()}s en este ${Provider.level1Label()}.</div>`;
        }
    } catch (err) {
        console.error("Error al cargar datasets:", err);
        container.innerHTML = `<div style='margin-left:20px; font-size:11px; color:#666;'>Error al cargar ${Provider.level2Label()}s.</div>`;
    }
}

async function loadTablesTree(projectId, datasetId, container) {
    const token = getAuthToken();
    if (!token) return;

    try {
        const items = await Provider.listTables(projectId, datasetId);
        container.innerHTML = "";

        if (items.length > 0) {
            const ul = document.createElement("ul");
            ul.className = "tree-list";
            items.forEach(tbl => {
                const tableId = tbl.id;
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
        showToast("Por favor selecciona una Tabla de Hechos.", "error");
        return;
    }

    try {
        const fields = await Provider.getTableFields(projectId, datasetId, tableId);

        if (fields && fields.length > 0) {

            // Config ya guardada de este modelo (SemanticModelStore), por
            // nombre de campo, para fusionarla con el esquema recién leído
            // del proveedor (por si hay columnas nuevas o eliminadas).
            let savedByName = {};

            if (isModelLoad && currentModel) {
                const savedModel = window.SemanticModelStore.getModel(currentModel);
                if (savedModel && savedModel.fields) {
                    savedModel.fields.forEach(f => { savedByName[f.name] = f; });
                }
            }

            fieldsState = fields.map(f => {
                const isNumeric = ["INTEGER", "FLOAT", "NUMERIC", "BIGNUMERIC"].includes(f.type);
                const saved = savedByName[f.name];

                let fieldType = isNumeric ? "MEASURE" : "DIMENSION";
                let isEnabled = true;

                if (isModelLoad) {
                    if (saved) {
                        fieldType = saved.type || fieldType;
                        isEnabled = !!saved.enabled;
                    } else {
                        isEnabled = false;
                    }
                }

                return {
                    name: f.name,
                    alias: f.name,
                    dataType: f.type,
                    type: fieldType,
                    enabled: isEnabled,
                    // Config Medida
                    aggregation: (saved && saved.aggregation) || "SUM",
                    format: (saved && saved.format) || "Auto",
                    // Config Dimensión (Relación)
                    relProject: (saved && saved.relProject) || "",
                    relDataset: (saved && saved.relDataset) || "",
                    relTable: (saved && saved.relTable) || "",
                    attributes: (saved && saved.attributes) || [],
                    hierarchies: (saved && saved.hierarchies) || []
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
        showToast("Habilita el campo para poder modificar sus propiedades.", "error");
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
        const fields = await Provider.getTableFields(projectId, datasetId, tableId);

        if (fields && fields.length > 0) {
            field.relProject = projectId;
            field.relDataset = datasetId;
            field.relTable = tableId;

            const existingAttrsMap = new Map((field.attributes || []).map(a => [a.name, a]));

            field.attributes = fields.map((attr, idx) => {
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
        showToast("Selecciona primero la tabla de dimensión y sus atributos antes de crear una jerarquía.", "error");
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
        showToast("Indica un nombre para la jerarquía.", "error");
        return;
    }

    if (hierEditState.levels.length === 0) {
        showToast("Añade al menos un nivel a la jerarquía.", "error");
        return;
    }

    if (!field.hierarchies) field.hierarchies = [];

    const duplicate = field.hierarchies.some((h, idx) =>
        h.name.toUpperCase() === name.toUpperCase() && idx !== editingHierarchyIndex
    );

    if (duplicate) {
        showToast("Ya existe una jerarquía con ese nombre en esta dimensión.", "error");
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

    showToast(`Jerarquía "${name}" guardada.`, "success");

}

/**
 * Guarda el modelo semántico completo en SemanticModelStore (JSON dentro de
 * las propiedades del documento, sin pestañas ni siquiera ocultas). Ya no
 * hace falta reconstruir 5 tablas distintas (MODEL_RELATIONSHIP,
 * MODEL_DIMENSION, MODEL_MEASURES, MODEL_ATRIBUTES, MODEL_HIER): fieldsState
 * ya contiene toda esa información (relación, atributos y jerarquías por
 * dimensión), así que se guarda tal cual junto con la tabla de hechos.
 * Las filas equivalentes a esas 5 tablas se siguen pudiendo generar al
 * vuelo cuando algo las necesita (ver SemanticModelStore.getModelGrid()).
 */
async function generateSemanticModelInExcel() {

    const modelName = resolveCurrentModelName();

    // Si se pulsa "Guardar" mientras todavía se está escribiendo el nombre
    // del nuevo modelo (el "+"), el selector #semanticModelSelect no existe
    // en el DOM en ese momento (está sustituido por el input #newModelName).
    // Hay que resolver el nombre PRIMERO (arriba) y restaurar el selector
    // AQUÍ, antes de que nada más intente tocar #semanticModelSelect —si no,
    // loadSemanticModels() de más abajo revienta con
    // "Cannot set properties of null (setting 'innerHTML')".
    if (creatingModel) {
        creatingModel = false;
        restoreModelSelector();
    }

    currentModel = modelName;

    if (currentModel === "") {
        showToast("Indica primero el nombre del modelo.", "error");
        return;
    }

    setFactCardVisible(true);

    const fact = {
        project: document.getElementById("factProject").value.trim(),
        dataset: document.getElementById("factDataset").value,
        table: document.getElementById("factTable").value
    };

    try {

        await window.SemanticModelStore.saveModel(currentModel, {
            fact: fact,
            fields: fieldsState
        });

        await window.SemanticModelStore.setActiveModelName(currentModel);

        loadSemanticModels();

        const select = document.getElementById("semanticModelSelect");

        if (select) {
            select.value = currentModel;
        }

        showToast("¡Modelo Semántico generado con éxito!", "success");

    } catch (err) {
        console.error("Error al guardar el modelo semántico:", err);
        showToast("Error al guardar el modelo: " + err.message, "error");
    }
}

/* ============================================================
 * Los antiguos popups "Abrir modelo semántico (.lkml)" y "Guardar
 * modelo semántico (.lkml)" (con sus pestañas Servidor/Local para
 * GitHub/GitLab o fichero local) ya NO viven aquí como un modal
 * superpuesto dentro del taskpane. Esa lógica se ha movido a dos
 * páginas de diálogo independientes (Office.context.ui.displayDialogAsync),
 * que se abren igual desde los botones del ribbon del manifiesto
 * ("Abrir modelo semántico" / "Guardar modelo semántico", ver
 * commands.js) que desde los botones "📂 Abrir" / "📤 LookML" de este
 * taskpane (ver openLkmlDialogFromTaskpane() / openSaveLkmlDialogFromTaskpane()
 * más arriba):
 *   - js/openSemanticModelDialog.js  (openSemanticModel.html)
 *   - js/saveSemanticModelDialog.js (saveSemanticModel.html)
 *   - js/lkmlExport.js               (generación del contenido LookML,
 *                                      compartida por el diálogo de guardado)
 *   - js/lkmlImport.js               (parseo del contenido LookML a modelo
 *                                      semántico, usado por el diálogo de
 *                                      apertura)
 *   - js/lkmlSaveBridge.js           (escribe el LookML guardado en
 *                                      EDIT_REPORT!G1, ya que el diálogo de
 *                                      guardado no tiene acceso a Excel)
 *   - js/lkmlOpenBridge.js           (guarda en SemanticModelStore el
 *                                      modelo importado, ya que el diálogo
 *                                      de apertura no tiene acceso a
 *                                      Office.context.document.settings)
 * ============================================================ */
