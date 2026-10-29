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

            // Recién creada: se decide (después) si las EPM_* son del XLAM o del add-in
            setTimeout(() => { if (window.DracoFn) window.DracoFn.ensure(); }, 0);

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
    // Medidas calculadas / filtradas del modelo
    document.getElementById("btnNewCalcMeasure").addEventListener("click", () => createVirtualMeasure("CALC"));
    document.getElementById("btnNewFilteredMeasure").addEventListener("click", () => createVirtualMeasure("FILTERED"));
    document.getElementById("btnAddMeasureCond").addEventListener("click", () => addMeasureCondRow({}));
    document.getElementById("btnAddCalcAttr").addEventListener("click", () => openCalcAttrDialog(fieldsState[currentConfigFieldIndex], null));
    document.getElementById("modalMeasureFormula").addEventListener("input", validateMeasureFormulaUI);

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
                return;
            }
            // El diálogo pide aquí el token de la sesión (BQ.requestSessionFromOpener).
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                if (window.BQ) BQ.handleDialogMessage(dialog, arg.message);
            });

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
        let select = l.attribute;
        if (attr && attr.calc && attr.expression && window.DracoAttrFormula) {
            // Atributo calculado: su expresión, con el nombre del atributo como alias
            try { select = window.DracoAttrFormula.toSql(attr.expression, c => c, Provider.key ? Provider.key() : "bigquery") + " AS " + attr.name; }
            catch (e) { select = l.attribute; }
        }
        return {
            field: l.attribute,
            select,
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
                return;
            }
            // El diálogo pide aquí el token de la sesión (BQ.requestSessionFromOpener).
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                if (window.BQ) BQ.handleDialogMessage(dialog, arg.message);
            });

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
                // SEGURIDAD: projectId viene de la API del proveedor
                // (BigQuery); aunque suele tener un charset
                // restringido, se construye con textContent por defensa
                // en profundidad en vez de innerHTML.
                const toggleSpan1 = document.createElement("span");
                toggleSpan1.className = "tree-toggle";
                toggleSpan1.textContent = "\u25b6"; // ▶
                const strongEl = document.createElement("strong");
                strongEl.textContent = projectId;
                header.appendChild(toggleSpan1);
                header.appendChild(document.createTextNode(" \ud83d\udcc1 "));
                header.appendChild(strongEl);
                
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
                // SEGURIDAD: datasetId viene de la API del proveedor -> textContent.
                const toggleSpan2 = document.createElement("span");
                toggleSpan2.className = "tree-toggle";
                toggleSpan2.textContent = "\u25b6"; // ▶
                header.appendChild(toggleSpan2);
                header.appendChild(document.createTextNode(" \ud83d\udcca " + datasetId));
                
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
                // SEGURIDAD: tableId viene de la API del proveedor -> textContent.
                itemDiv.textContent = "\ud83d\udccb " + tableId;
                
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
        const factP = document.getElementById("factProject").value.trim();
        const factD = document.getElementById("factDataset").value;
        const factT = document.getElementById("factTable").value;
        if (projectId === factP && datasetId === factD && tableId === factT) {
            // Elegir la tabla de hechos como "tabla de dimensión" generaba un
            // JOIN de la tabla consigo misma: se usa el modo sin JOIN.
            closeTreeModal();
            showToast("Es la tabla de hechos: para crear una dimensión con sus columnas, selecciónalas en la lista y usa clic derecho ▸ Agrupar.", "error");
            return;
        }
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
                    hierarchies: (saved && saved.hierarchies) || [],
                    // Atributo de texto (descripción) de la clave, "" = ninguno.
                    // El modo "ID / ID y texto" se elige por campo en el informe.
                    textAttribute: (saved && saved.textAttribute) || "",
                    // Miembros calculados del modelo por atributo (ver openModelCalcMembers)
                    calcMembers: (saved && saved.calcMembers) || {}
                };
            });

            // Grupos guardados (dimensiones de la tabla de hechos cuyo nombre
            // no es una columna) y medidas calculadas / filtradas: se añaden
            // tras las columnas.
            if (isModelLoad) {
                const colNames = new Set(fieldsState.map(f => f.name));
                Object.keys(savedByName).forEach(n => {
                    const v = savedByName[n];
                    if (!colNames.has(n) && isVirtualMeasure(v)) fieldsState.push(Object.assign({}, v));
                });
                Object.keys(savedByName).forEach(n => {
                    const g = savedByName[n];
                    if (colNames.has(n) || g.type !== "DIMENSION" || g.relTable) return;
                    const attrs = (g.attributes || []).filter(a => colNames.has(a.name));
                    if (!attrs.length) return;
                    fieldsState.push(Object.assign({}, g, { isGroup: true, relProject: "", relDataset: "", relTable: "", attributes: attrs }));
                });
            }

            // Columnas agrupadas: no son campos sueltos aunque el modelo
            // guardado (o importado de un .lkml) no lo dijera.
            groupedOwnersMap().forEach((owner, colName) => {
                const col = fieldsState.find(f => f.name === colName);
                if (col) col.enabled = false;
            });

            smSelection.clear();
            renderFieldsTable();
            document.getElementById("fieldsCard").style.display = "block";
            // Modelo recién abierto = sin cambios; una tabla nueva = cambios sin guardar
            if (isModelLoad) smMarkSaved(); else smSavedSnapshot = null;
        }
    } catch (err) {
        console.error("Error al obtener esquema:", err);
    }
}

function renderFieldsTable() {

    const list = document.getElementById("fieldsList");
    list.innerHTML = "";
    closeSmContextMenu();

    // Orden de pintado: primero los grupos (dimensiones de la tabla de
    // hechos) con sus columnas sangradas debajo; luego el resto de campos.
    // fieldsState no cambia de orden.
    const owners = groupedOwnersMap();
    const order = [];
    fieldsState.forEach((field, idx) => {
        if (!isFactGroup(field) || owners.has(field.name)) return;
        order.push(idx);
        fieldsState.forEach((child, cIdx) => { if (owners.get(child.name) === field.name) order.push(cIdx); });
    });
    fieldsState.forEach((field, idx) => {
        if (isFactGroup(field) || owners.has(field.name)) return;
        order.push(idx);
    });
    smRenderOrder = order;
    const present = new Set(fieldsState.map(f => f.name));
    smSelection.forEach(n => { if (!present.has(n)) smSelection.delete(n); });

    order.forEach(idx => {
        const field = fieldsState[idx];
        const owner = owners.get(field.name) || "";
        const group = isFactGroup(field);

        const row = document.createElement("div");
        row.className = "field-row";
        row.dataset.name = field.name;
        if (owner) {
            row.classList.add("field-grouped");
            row.title = `Campo de la dimensión ${owner}. Clic derecho ▸ Quitar del grupo.`;
        }
        if (group) row.classList.add("field-group-row");
        if (smSelection.has(field.name)) row.classList.add("field-selected");

        // SEGURIDAD: field.alias es texto libre editado por el usuario (y
        // field.name puede venir de un modelo importado desde un repo Git
        // compartido) -> se construyen los nodos con DOM API en vez de
        // interpolar en un template literal HTML.
        const enableInput = document.createElement("input");
        enableInput.className = "field-enable";
        enableInput.type = "checkbox";
        enableInput.checked = !!field.enabled;
        enableInput.addEventListener("click", (e) => e.stopPropagation());
        enableInput.addEventListener("change", () => updateFieldEnabled(idx, enableInput.checked));

        const typeSelect = document.createElement("select");
        typeSelect.className = "field-type";
        typeSelect.addEventListener("click", (e) => e.stopPropagation());
        typeSelect.addEventListener("change", () => updateFieldType(idx, typeSelect.value));
        [["MEASURE", "MEA"], ["DIMENSION", "DIM"]].forEach(([value, text]) => {
            const opt = document.createElement("option");
            opt.value = value;
            opt.textContent = text;
            opt.selected = field.type === value;
            typeSelect.appendChild(opt);
        });

        const fieldInfo = document.createElement("div");
        fieldInfo.className = "field-info";

        const aliasInput = document.createElement("input");
        aliasInput.className = "field-alias-input";
        aliasInput.type = "text";
        aliasInput.value = field.alias;
        aliasInput.addEventListener("click", (e) => e.stopPropagation());
        aliasInput.addEventListener("input", () => updateFieldAlias(idx, aliasInput.value));

        const fieldNameDiv = document.createElement("div");
        fieldNameDiv.className = "field-name";
        const virtualMeasure = isVirtualMeasure(field);
        fieldNameDiv.textContent = group
            ? `${field.name} · ${(field.attributes || []).filter(a => a.name !== field.name).length || (field.attributes || []).length} campos`
            : virtualMeasure ? `${field.name} · ${field.measureKind === "CALC" ? "calculada" : "filtrada"}` : field.name;
        if (virtualMeasure) row.classList.add("field-virtual");

        fieldInfo.appendChild(aliasInput);
        fieldInfo.appendChild(fieldNameDiv);

        if (owner || group || virtualMeasure) typeSelect.disabled = true;
        if (owner) enableInput.disabled = true;

        row.appendChild(enableInput);
        row.appendChild(typeSelect);
        row.appendChild(fieldInfo);

        // Flecha de propiedades: en los campos sueltos y en los grupos (no en
        // los campos agrupados, que se configuran desde su grupo).
        if (!owner) {
            const arrowDiv = document.createElement("div");
            arrowDiv.className = "field-arrow";
            arrowDiv.textContent = "\u203a"; // ›
            arrowDiv.title = "Propiedades";
            arrowDiv.addEventListener("click", (e) => { e.stopPropagation(); openConfigModal(idx); });
            row.appendChild(arrowDiv);
        }

        // Clic: selecciona (Ctrl = añadir/quitar, Mayús = rango). Doble clic:
        // propiedades. Clic derecho: agrupar / desagrupar.
        row.addEventListener("click", (e) => onFieldRowClick(e, idx));
        row.addEventListener("dblclick", (e) => {
            if (e.target.closest("input,select")) return;
            if (owner) { const g = fieldsState.findIndex(f => f.name === owner); if (g >= 0) openConfigModal(g); }
            else openConfigModal(idx);
        });
        row.addEventListener("contextmenu", (e) => openSmContextMenu(e, idx));

        list.appendChild(row);
    });
}

function updateFieldAlias(index, val) { fieldsState[index].alias = val; }
function updateFieldType(index, val) {
    const field = fieldsState[index];
    if (val !== "DIMENSION" && isFactGroup(field)) ungroup(field);
    field.type = val;
    if (val !== "DIMENSION" && (field.attributes || []).some(a => a.name !== field.name)) {
        field.attributes = [];
        field.hierarchies = [];
        field.textAttribute = "";
    }
    renderFieldsTable();
}
function updateFieldEnabled(index, val) { fieldsState[index].enabled = val; }

function openConfigModal(index) {
    currentConfigFieldIndex = index;
    const field = fieldsState[index];

    if (!field.enabled) {
        showToast("Habilita el campo para poder modificar sus propiedades.", "error");
        return;
    }

    if (field.type === "MEASURE") {
        document.getElementById("modalMeasureFieldName").textContent = field.name + (isVirtualMeasure(field)
            ? (field.measureKind === "CALC" ? "  (calculada)" : "  (filtrada)") : "");
        document.getElementById("modalMeasureAlias").value = field.alias;
        document.getElementById("modalMeasureAgg").value = field.aggregation;
        document.getElementById("modalMeasureFormat").value = field.format;
        openVirtualMeasureSections(field);
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

        // Grupo de columnas de la tabla de hechos: sin selector de tabla; sus
        // atributos son las columnas agrupadas. El resto, como siempre.
        const group = isFactGroup(field);
        document.getElementById("dimTableSelector").style.display = group ? "none" : "";
        document.getElementById("dimGroupHint").style.display = group ? "" : "none";

        if (group) {
            document.getElementById("dimRelFullConcat").value = "";
            renderAttributesTable(field.attributes);
            document.getElementById("attributesContainer").style.display = "block";
        } else if (proj && ds && tbl) {
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
    if (currentConfigFieldIndex !== null && !saveVirtualMeasureSections(fieldsState[currentConfigFieldIndex])) return; // errores a la vista
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

            const calcAttrs = (field.attributes || []).filter(a => a.calc);
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
            // Atributos calculados: se conservan si sus columnas siguen en la tabla
            const colSet = new Set(field.attributes.map(a => a.name.toUpperCase()));
            calcAttrs.forEach(a => {
                const cols = window.DracoAttrFormula ? window.DracoAttrFormula.columns(a.expression) : [];
                if (cols.every(c => colSet.has(String(c).toUpperCase())) && !colSet.has(a.name.toUpperCase())) field.attributes.push(a);
            });

            // Si la tabla de dimensión ha cambiado, limpiamos de las jerarquías
            // cualquier nivel que referencie atributos que ya no existen.
            if (field.hierarchies && field.hierarchies.length > 0) {
                const validNames = new Set(field.attributes.map(a => a.name));
                field.hierarchies = field.hierarchies
                    .map(h => ({
                        name: h.name,
                        levels: h.levels
                            .filter(l => validNames.has(l.attribute))
                            .map(l => ({ attribute: l.attribute, textAttribute: validNames.has(l.textAttribute) ? l.textAttribute : "" }))
                    }))
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

    const ownerField = fieldsState[currentConfigFieldIndex];
    if (ownerField) validateDimTextAttribute(ownerField);

    const list = document.getElementById("attributesList");
    list.innerHTML = "";

    attributes.forEach((attr, idx) => {

        const row = document.createElement("div");
        row.className = "field-row";

        // SEGURIDAD: attr.alias/attr.name, igual que en renderFieldsTable,
        // no son de confianza (texto libre / modelo importado) -> DOM API
        // en vez de template literal HTML.
        const enableInput = document.createElement("input");
        enableInput.className = "field-enable";
        enableInput.type = "checkbox";
        enableInput.checked = !!attr.enabled;
        enableInput.addEventListener("click", (e) => e.stopPropagation());
        enableInput.addEventListener("change", () => updateAttrEnabled(idx, enableInput.checked));

        const fieldInfo = document.createElement("div");
        fieldInfo.className = "field-info";

        const aliasInput = document.createElement("input");
        aliasInput.className = "field-alias-input";
        aliasInput.type = "text";
        aliasInput.value = attr.alias;
        aliasInput.addEventListener("click", (e) => e.stopPropagation());
        aliasInput.addEventListener("input", () => updateAttrAlias(idx, aliasInput.value));

        const fieldNameDiv = document.createElement("div");
        fieldNameDiv.className = "field-name";
        if (attr.calc) {
            // Atributo calculado: «ƒx NOMBRE = fórmula»; clic para editarlo
            row.classList.add("calc-attr-row");
            const b = document.createElement("span");
            b.className = "calc-attr-badge";
            b.textContent = "\u0192x";
            fieldNameDiv.appendChild(b);
            fieldNameDiv.appendChild(document.createTextNode(attr.name + " = " + (attr.expression || "")));
            fieldNameDiv.title = "Clic para editar la fórmula";
            fieldNameDiv.addEventListener("click", (e) => { e.stopPropagation(); openCalcAttrDialog(fieldsState[currentConfigFieldIndex], attr); });
        } else {
            fieldNameDiv.textContent = attr.name;
        }

        fieldInfo.appendChild(aliasInput);
        fieldInfo.appendChild(fieldNameDiv);

        const keyLabel = document.createElement("label");
        keyLabel.className = "field-key";
        const keyRadio = document.createElement("input");
        keyRadio.type = "radio";
        keyRadio.name = "dimKeyGroup";
        keyRadio.checked = !!attr.isKey;
        keyRadio.disabled = !!attr.calc; // un atributo calculado no puede ser la clave
        keyRadio.addEventListener("click", (e) => e.stopPropagation());
        keyRadio.addEventListener("change", () => updateAttrKey(idx));
        const keySpan = document.createElement("span");
        keySpan.textContent = "Key";
        keyLabel.appendChild(keyRadio);
        keyLabel.appendChild(keySpan);

        const field = fieldsState[currentConfigFieldIndex];
        const textLabel = document.createElement("label");
        textLabel.className = "field-key";
        textLabel.title = "Atributo de texto (descripción) de la clave: en el informe se puede mostrar junto al ID";
        const textCb = document.createElement("input");
        textCb.type = "checkbox";
        textCb.checked = !!field && field.textAttribute === attr.name;
        textCb.disabled = !(textCb.checked || canBeTextAttribute(attr));
        textCb.addEventListener("click", (e) => e.stopPropagation());
        textCb.addEventListener("change", () => updateAttrText(idx, textCb.checked));
        const textSpan = document.createElement("span");
        textSpan.textContent = "Texto";
        textLabel.appendChild(textCb);
        textLabel.appendChild(textSpan);

        row.appendChild(enableInput);
        row.appendChild(fieldInfo);
        row.appendChild(keyLabel);
        row.appendChild(textLabel);

        // Miembros calculados DEL MODELO para este atributo (p. ej. OIBDA =
        // [Ingresos] - [Gastos] en CUENTA.NIVEL1): salen en todos los informes.
        if (!attr.calc && window.DracoCalcEditor && field) {
            const n = ((field.calcMembers || {})[String(attr.name).toUpperCase()] || []).length;
            const fxBtn = document.createElement("button");
            fxBtn.type = "button";
            fxBtn.className = "attr-remove-btn attr-fx-btn" + (n ? " has-calcs" : "");
            fxBtn.textContent = n ? "\u0192x " + n : "\u0192x";
            fxBtn.title = n ? `${n} miembro(s) calculado(s) en el modelo` : "Miembros calculados del modelo (p. ej. OIBDA = [Ingresos] - [Gastos])";
            fxBtn.addEventListener("click", (e) => { e.stopPropagation(); openModelCalcMembers(field, attr.name); });
            row.appendChild(fxBtn);
        }

        // Dimensión de la tabla de hechos: se puede sacar una columna del
        // grupo (vuelve a la lista de campos como campo suelto). La columna
        // propia de la dimensión no se puede quitar.
        if (attr.calc) {
            const rm = document.createElement("button");
            rm.type = "button";
            rm.className = "attr-remove-btn";
            rm.title = "Eliminar el atributo calculado";
            rm.textContent = "\u2715";
            rm.addEventListener("click", (e) => {
                e.stopPropagation();
                dropAttributeFromField(field, attr.name);
                renderAttributesTable(field.attributes);
                renderHierarchiesList();
            });
            row.appendChild(rm);
        } else if (isFactGroup(field) && attr.name !== field.name) {
            const rm = document.createElement("button");
            rm.type = "button";
            rm.className = "attr-remove-btn";
            rm.title = "Quitar de la dimensión";
            rm.textContent = "\u2715";
            rm.addEventListener("click", (e) => { e.stopPropagation(); removeFactColumnFromDimension(field, attr.name); });
            row.appendChild(rm);
        }

        list.appendChild(row);

    });

}

/* ---------------------------------------------------------------------
 * Atributo de TEXTO (descripción) de la clave de la dimensión: se marca con
 * el checkbox "Texto" junto a "Key" (uno o ninguno). Solo atributos
 * habilitados, de tipo texto y que no sean la clave. Si el informe pide
 * "ID y texto" para la clave, se muestra este atributo en una 2ª columna.
 * ------------------------------------------------------------------- */
function isTextAttributeType(attr) {
    return /STRING|VARCHAR|CHAR|TEXT/i.test(String((attr && attr.dataType) || ""));
}

function canBeTextAttribute(attr) {
    return !!attr && attr.enabled && !attr.isKey && isTextAttributeType(attr);
}

function textAttributeCandidates(field, excludeNames) {
    const exclude = new Set((excludeNames || []).map(n => String(n)));
    return (field.attributes || []).filter(a => canBeTextAttribute(a) && !exclude.has(a.name));
}

// Quita el texto marcado si ya no es válido (deshabilitado, ahora es la clave...)
function validateDimTextAttribute(field) {
    if (field.textAttribute && !(field.attributes || []).some(a => a.name === field.textAttribute && canBeTextAttribute(a))) {
        field.textAttribute = "";
    }
}

function updateAttrAlias(idx, val) { fieldsState[currentConfigFieldIndex].attributes[idx].alias = val; }
function updateAttrKey(selectedIdx) {
    const field = fieldsState[currentConfigFieldIndex];
    field.attributes.forEach((attr, idx) => {
        attr.isKey = (idx === selectedIdx);
    });
    validateDimTextAttribute(field);
    renderAttributesTable(field.attributes);
}
function updateAttrEnabled(idx, val) {
    const field = fieldsState[currentConfigFieldIndex];
    field.attributes[idx].enabled = val;
    validateDimTextAttribute(field);
    renderAttributesTable(field.attributes);
}
function updateAttrText(idx, checked) {
    const field = fieldsState[currentConfigFieldIndex];
    field.textAttribute = checked ? field.attributes[idx].name : "";
    renderAttributesTable(field.attributes);
}

/* =====================================================================
 * GRUPOS: DIMENSIONES SOBRE LA TABLA DE HECHOS (sin JOIN)
 *
 * En la lista de campos se seleccionan varias columnas (Ctrl/Mayús + clic)
 * y con clic derecho ▸ «Agrupar» se crea una dimensión con el nombre que
 * se pida. Esa dimensión es la que tiene la flecha de propiedades: dentro
 * se ven sus columnas (atributos) y se crean las jerarquías, igual que en
 * una dimensión con tabla, pero sin selector de tabla.
 *
 * Un grupo es un campo más de fieldsState (isGroup, sin relTable) cuyo
 * nombre NO es una columna: el SQL usa sus atributos como f.<columna>, sin
 * JOIN (ver getTableAlias en commands.js). Las columnas agrupadas dejan de
 * ser campos sueltos (enabled = false) y se pintan sangradas bajo el grupo.
 * Que una columna está agrupada se deduce de los atributos del grupo.
 * También se tratan como grupo las dimensiones de la tabla de hechos con
 * atributos creadas con la versión anterior (columna + columnas).
 * ===================================================================== */

let smSelection = new Set();   // nombres de campo seleccionados en la lista
let smAnchor = null;           // último clic (para Mayús + clic)
let smRenderOrder = [];        // índices de fieldsState en el orden pintado

function isFactDimension(field) {
    return !!field && field.type === "DIMENSION" && !field.relTable;
}

function isFactGroup(field) {
    return isFactDimension(field) && (!!field.isGroup || (field.attributes || []).some(a => a.name !== field.name));
}

/** Map nombreColumna -> nombreGrupo de las columnas agrupadas. */
function groupedOwnersMap() {
    const names = new Set(fieldsState.map(f => f.name));
    const map = new Map();
    fieldsState.forEach(g => {
        if (!isFactGroup(g)) return;
        (g.attributes || []).forEach(a => {
            if (a.name !== g.name && names.has(a.name) && !map.has(a.name)) map.set(a.name, g.name);
        });
    });
    return map;
}

/** Columnas que se pueden meter en un grupo: sueltas, sin tabla de dimensión propia. */
function isGroupableColumn(field, owners) {
    return !!field && !isFactGroup(field) && !owners.has(field.name) && !field.relTable;
}

/** Quita un atributo de una dimensión y lo limpia de sus jerarquías y del texto. */
function dropAttributeFromField(field, attrName) {
    field.attributes = (field.attributes || []).filter(a => a.name !== attrName);
    if (field.textAttribute === attrName) field.textAttribute = "";
    field.hierarchies = (field.hierarchies || [])
        .map(h => ({
            name: h.name,
            levels: (h.levels || []).filter(l => l.attribute !== attrName)
                .map(l => (l.textAttribute === attrName ? { attribute: l.attribute } : l))
        }))
        .filter(h => h.levels.length > 0);
    if (field.attributes.length && !field.attributes.some(a => a.isKey)) field.attributes[0].isKey = true;
}

function attrFromColumn(col, isKey) {
    return { name: col.name, alias: col.alias || col.name, dataType: col.dataType, isKey: !!isKey, enabled: true };
}

function createGroup(name, columns) {
    const insertAt = Math.min.apply(null, columns.map(c => fieldsState.indexOf(c)));
    const group = {
        name, alias: name, dataType: "STRING", type: "DIMENSION", enabled: true, isGroup: true,
        aggregation: "SUM", format: "Auto",
        relProject: "", relDataset: "", relTable: "",
        attributes: columns.map((c, i) => attrFromColumn(c, i === 0)),
        hierarchies: [], textAttribute: ""
    };
    columns.forEach(c => { c.enabled = false; c.type = "DIMENSION"; });
    fieldsState.splice(insertAt, 0, group);
    smSelection = new Set([name]);
    smAnchor = name;
    renderFieldsTable();
    showToast(`Dimensión ${name} creada con ${columns.length} campo(s). Ábrela (›) para crear jerarquías.`, "success");
}

function addColumnsToGroup(group, columns) {
    columns.forEach(c => {
        if ((group.attributes || []).some(a => a.name === c.name)) return;
        group.attributes.push(attrFromColumn(c, !(group.attributes || []).length));
        c.enabled = false;
        c.type = "DIMENSION";
    });
    renderFieldsTable();
    showToast(`${columns.length} campo(s) añadido(s) a ${group.name}.`, "success");
}

function removeFactColumnFromDimension(group, columnName) {
    if (!group || columnName === group.name) return;
    dropAttributeFromField(group, columnName);
    const col = fieldsState.find(f => f.name === columnName);
    if (col) col.enabled = true; // vuelve a ser un campo suelto
    if (group.isGroup && !(group.attributes || []).length) {
        fieldsState.splice(fieldsState.indexOf(group), 1);
        closeDimModalIfOpen();
    } else if (document.getElementById("dimensionModal").style.display === "block"
        && fieldsState[currentConfigFieldIndex] === group) {
        renderAttributesTable(group.attributes);
        renderHierarchiesList();
    }
    renderFieldsTable();
}

function ungroup(group) {
    (group.attributes || []).forEach(a => {
        if (a.name === group.name) return;
        const col = fieldsState.find(f => f.name === a.name);
        if (col) col.enabled = true;
    });
    if (group.isGroup) {
        fieldsState.splice(fieldsState.indexOf(group), 1);
    } else {
        // dimensión antigua (columna + columnas): se queda la columna sola
        group.attributes = [];
        group.hierarchies = [];
        group.textAttribute = "";
    }
    smSelection.clear();
    renderFieldsTable();
}

function closeDimModalIfOpen() {
    const m = document.getElementById("dimensionModal");
    if (m) m.style.display = "none";
}

/* ---------- selección en la lista ---------- */

function onFieldRowClick(e, idx) {
    const field = fieldsState[idx];
    const name = field.name;
    if (e.shiftKey && smAnchor) {
        const names = smRenderOrder.map(i => fieldsState[i].name);
        const a = names.indexOf(smAnchor), b = names.indexOf(name);
        if (a >= 0 && b >= 0) {
            if (!(e.ctrlKey || e.metaKey)) smSelection.clear();
            names.slice(Math.min(a, b), Math.max(a, b) + 1).forEach(n => smSelection.add(n));
        }
        paintSelection();
        return;
    }
    if (e.ctrlKey || e.metaKey) {
        if (smSelection.has(name)) smSelection.delete(name); else smSelection.add(name);
        smAnchor = name;
        paintSelection();
        return;
    }
    smSelection = new Set([name]);
    smAnchor = name;
    paintSelection();
}

function paintSelection() {
    document.querySelectorAll("#fieldsList .field-row").forEach(r => {
        r.classList.toggle("field-selected", smSelection.has(r.dataset.name));
    });
}

/* ---------- menú contextual ---------- */

function closeSmContextMenu() {
    const m = document.getElementById("smCtxMenu");
    if (m) m.remove();
}

function openSmContextMenu(e, idx) {
    e.preventDefault();
    closeSmContextMenu();
    const field = fieldsState[idx];
    if (!smSelection.has(field.name)) { smSelection = new Set([field.name]); smAnchor = field.name; paintSelection(); }

    const owners = groupedOwnersMap();
    const selected = fieldsState.filter(f => smSelection.has(f.name));
    const groups = selected.filter(f => isFactGroup(f));
    const columns = selected.filter(f => isGroupableColumn(f, owners));
    const members = selected.filter(f => owners.has(f.name));

    const items = [];
    if (columns.length && !groups.length && !members.length) {
        items.push(["Agrupar en una dimensión…", () => askGroupName(columns)]);
    }
    if (groups.length === 1 && columns.length) {
        items.push([`Añadir al grupo ${groups[0].name}`, () => addColumnsToGroup(groups[0], columns)]);
    }
    if (members.length) {
        items.push(["Quitar del grupo", () => members.forEach(m => {
            const g = fieldsState.find(f => f.name === owners.get(m.name));
            if (g) removeFactColumnFromDimension(g, m.name);
        })]);
    }
    if (groups.length === 1 && !columns.length) {
        items.push(["Propiedades…", () => openConfigModal(fieldsState.indexOf(groups[0]))]);
        items.push([`Desagrupar ${groups[0].name}`, () => ungroup(groups[0])]);
    }
    const vms = selected.filter(f => isVirtualMeasure(f));
    if (vms.length === 1 && selected.length === 1) {
        items.push(["Propiedades…", () => openConfigModal(fieldsState.indexOf(vms[0]))]);
        items.push([`Eliminar ${vms[0].name}`, () => deleteVirtualMeasure(vms[0])]);
    }
    if (!items.length) {
        items.push(["Selecciona campos sueltos (Ctrl o Mayús + clic) para agrupar", null]);
    }
    items.push(["Nueva medida calculada…", () => createVirtualMeasure("CALC")]);
    items.push(["Nueva medida filtrada…", () => createVirtualMeasure("FILTERED")]);

    const menu = document.createElement("div");
    menu.id = "smCtxMenu";
    menu.className = "sm-ctx-menu";
    items.forEach(([label, fn]) => {
        const b = document.createElement("button");
        b.type = "button";
        b.textContent = label;
        if (!fn) { b.disabled = true; b.style.color = "#94a3b8"; b.style.cursor = "default"; }
        else b.addEventListener("click", () => { closeSmContextMenu(); fn(); });
        menu.appendChild(b);
    });
    document.body.appendChild(menu);
    const w = menu.offsetWidth, h = menu.offsetHeight;
    menu.style.left = Math.max(4, Math.min(e.clientX, window.innerWidth - w - 4)) + "px";
    menu.style.top = Math.max(4, Math.min(e.clientY, window.innerHeight - h - 4)) + "px";
    setTimeout(() => {
        document.addEventListener("mousedown", function off(ev) {
            if (!menu.contains(ev.target)) { closeSmContextMenu(); document.removeEventListener("mousedown", off); }
        });
    }, 0);
}

/* ---------- diálogos propios (Office no admite prompt/confirm) ---------- */

function smDialog({ title, text, input, buttons }) {
    return new Promise(resolve => {
        const back = document.createElement("div");
        back.className = "sm-dialog-backdrop";
        const box = document.createElement("div");
        box.className = "sm-dialog";
        box.setAttribute("role", "dialog");
        const h = document.createElement("h3");
        h.textContent = title;
        box.appendChild(h);
        if (text) { const p = document.createElement("div"); p.textContent = text; box.appendChild(p); }
        let inp = null, err = null;
        if (input) {
            inp = document.createElement("input");
            inp.type = "text";
            inp.value = input.value || "";
            inp.placeholder = input.placeholder || "";
            box.appendChild(inp);
            err = document.createElement("div");
            err.className = "sm-dialog-err";
            box.appendChild(err);
        }
        const actions = document.createElement("div");
        actions.className = "sm-dialog-actions";
        const done = (v) => { back.remove(); resolve(v); };
        buttons.forEach(bd => {
            const b = document.createElement("button");
            b.type = "button";
            b.className = bd.primary ? "btn-primary btn-sm" : "btn-secondary btn-sm";
            b.textContent = bd.label;
            b.addEventListener("click", () => {
                if (bd.value === "ok" && input) {
                    const msg = input.validate ? input.validate(inp.value.trim()) : "";
                    if (msg) { err.textContent = msg; inp.focus(); return; }
                    done(inp.value.trim());
                    return;
                }
                done(bd.value);
            });
            actions.appendChild(b);
        });
        box.appendChild(actions);
        back.appendChild(box);
        back.addEventListener("keydown", (e) => {
            if (e.key === "Escape") done(null);
            if (e.key === "Enter" && inp) { const ok = [...actions.children].find(x => x.className.indexOf("btn-primary") === 0); if (ok) ok.click(); }
        });
        document.body.appendChild(back);
        (inp || actions.lastChild).focus();
        if (inp) inp.select();
    });
}

async function askGroupName(columns) {
    const suggestion = String(columns[0].name).replace(/(_N\d+|_\d+|_TXT|_DESC)$/i, "");
    const taken = new Set(fieldsState.map(f => f.name.toUpperCase()));
    const name = await smDialog({
        title: "Agrupar en una dimensión",
        text: `${columns.length} campo(s): ${columns.map(c => c.name).join(", ")}`,
        input: {
            value: taken.has(suggestion.toUpperCase()) ? suggestion + "_DIM" : suggestion,
            placeholder: "Nombre de la dimensión",
            validate: (v) => {
                if (!v) return "Pon un nombre.";
                if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(v)) return "Solo letras, números y _ (sin espacios ni acentos), sin empezar por número.";
                if (taken.has(v.toUpperCase())) return "Ya hay un campo con ese nombre.";
                return "";
            }
        },
        buttons: [{ label: "Cancelar", value: null }, { label: "Agrupar", value: "ok", primary: true }]
    });
    if (name) createGroup(name, columns);
}

/* =====================================================================
 * MEDIDAS CALCULADAS Y FILTRADAS DEL MODELO
 *   - Calculada: fórmula sobre otras medidas ([Ingresos] - [Gastos],
 *     DIVIDIR([Margen]; [Ventas])...). Se calcula después de sumar.
 *   - Filtrada: una medida base solo donde se cumplen condiciones sobre
 *     columnas de la tabla de hechos (Importe con ESCENARIO = REAL).
 * Son campos de fieldsState que no son columnas (measureKind CALC /
 * FILTERED); el SQL lo genera commands.js (dracoMeasureAggSQL).
 * ===================================================================== */

function isVirtualMeasure(f) {
    return !!f && f.type === "MEASURE" && (f.measureKind === "CALC" || f.measureKind === "FILTERED");
}

function baseMeasures() {
    return fieldsState.filter(f => f.type === "MEASURE" && !isVirtualMeasure(f));
}

async function createVirtualMeasure(kind) {
    const taken = new Set(fieldsState.map(f => f.name.toUpperCase()));
    const name = await smDialog({
        title: kind === "CALC" ? "Nueva medida calculada" : "Nueva medida filtrada",
        text: kind === "CALC" ? "Fórmula sobre otras medidas, p. ej. OIBDA = [Ingresos] - [Gastos]." : "Una medida solo donde se cumplen unas condiciones, p. ej. Importe con Escenario = REAL.",
        input: {
            value: kind === "CALC" ? "MEDIDA_CALC" : "MEDIDA_FILTRADA",
            placeholder: "Nombre de la medida",
            validate: (v) => {
                if (!v) return "Pon un nombre.";
                if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(v)) return "Solo letras, números y _ (sin espacios ni acentos), sin empezar por número.";
                if (taken.has(v.toUpperCase())) return "Ya hay un campo con ese nombre.";
                return "";
            }
        },
        buttons: [{ label: "Cancelar", value: null }, { label: "Crear", value: "ok", primary: true }]
    });
    if (!name) return;
    const base = baseMeasures()[0];
    fieldsState.push({
        name, alias: name, type: "MEASURE", enabled: true, dataType: "NUMERIC",
        measureKind: kind, formula: "", base: kind === "FILTERED" && base ? base.name : "",
        conditions: [], aggregation: (base && base.aggregation) || "SUM", format: (base && base.format) || "Auto",
        relProject: "", relDataset: "", relTable: "", attributes: [], hierarchies: [], textAttribute: ""
    });
    renderFieldsTable();
    openConfigModal(fieldsState.length - 1);
}

async function deleteVirtualMeasure(field) {
    const users = fieldsState.filter(f => isVirtualMeasure(f) && f !== field
        && ((f.measureKind === "CALC" && window.DracoMeasureFormula.refs(f.formula).some(r => r.toUpperCase() === field.name.toUpperCase()))
            || (f.measureKind === "FILTERED" && String(f.base).toUpperCase() === field.name.toUpperCase())));
    const ok = await smDialog({
        title: `Eliminar ${field.name}`,
        text: users.length ? `La usan: ${users.map(u => u.name).join(", ")}. Dejarán de funcionar hasta que las corrijas.` : "Se quita del modelo al guardarlo.",
        buttons: [{ label: "Cancelar", value: null }, { label: "Eliminar", value: "del", primary: true }]
    });
    if (ok !== "del") return;
    fieldsState.splice(fieldsState.indexOf(field), 1);
    smSelection.clear();
    renderFieldsTable();
}

function openVirtualMeasureSections(field) {
    const kind = isVirtualMeasure(field) ? field.measureKind : "";
    document.getElementById("measureFilteredSection").style.display = kind === "FILTERED" ? "" : "none";
    document.getElementById("measureCalcSection").style.display = kind === "CALC" ? "" : "none";
    // Una calculada no se agrega: cada medida de la fórmula lleva la suya
    document.getElementById("modalMeasureAggGroup").style.display = kind === "CALC" ? "none" : "";

    if (kind === "FILTERED") {
        const sel = document.getElementById("modalMeasureBase");
        sel.innerHTML = "";
        baseMeasures().forEach(m => {
            const o = document.createElement("option");
            o.value = m.name; o.textContent = m.alias && m.alias !== m.name ? `${m.alias} (${m.name})` : m.name;
            sel.appendChild(o);
        });
        sel.value = field.base || (baseMeasures()[0] || {}).name || "";
        const box = document.getElementById("modalMeasureConds");
        box.innerHTML = "";
        (field.conditions && field.conditions.length ? field.conditions : [{}]).forEach(addMeasureCondRow);
    }
    if (kind === "CALC") {
        document.getElementById("modalMeasureFormula").value = field.formula || "";
        const chips = document.getElementById("modalMeasureRefs");
        chips.innerHTML = "";
        fieldsState.filter(f => f.type === "MEASURE" && f.enabled && f !== field).forEach(m => {
            const b = document.createElement("button");
            b.type = "button";
            b.textContent = m.name;
            b.title = "Insertar [" + m.name + "]";
            b.addEventListener("click", () => insertIntoFormula("[" + m.name + "]"));
            chips.appendChild(b);
        });
        validateMeasureFormulaUI();
    }
}

function insertIntoFormula(text) {
    const ta = document.getElementById("modalMeasureFormula");
    const a = ta.selectionStart != null ? ta.selectionStart : ta.value.length;
    const b = ta.selectionEnd != null ? ta.selectionEnd : ta.value.length;
    ta.value = ta.value.slice(0, a) + text + ta.value.slice(b);
    ta.focus();
    ta.setSelectionRange(a + text.length, a + text.length);
    validateMeasureFormulaUI();
}

/** Columnas de la tabla de hechos para las condiciones (no virtuales). */
function factColumnsForConditions() {
    return fieldsState.filter(f => !isVirtualMeasure(f) && !f.isGroup);
}

function addMeasureCondRow(c) {
    const box = document.getElementById("modalMeasureConds");
    const row = document.createElement("div");
    row.className = "measure-cond-row";
    const col = document.createElement("select");
    col.className = "mc-col";
    factColumnsForConditions().forEach(f => {
        const o = document.createElement("option");
        o.value = f.name; o.textContent = f.name;
        col.appendChild(o);
    });
    if (c.column) col.value = c.column;
    const op = document.createElement("select");
    op.className = "mc-op";
    [["IN", "es"], ["NOT_IN", "no es"]].forEach(([v, t]) => {
        const o = document.createElement("option"); o.value = v; o.textContent = t; op.appendChild(o);
    });
    op.value = c.op === "NOT_IN" ? "NOT_IN" : "IN";
    const vals = document.createElement("input");
    vals.type = "text";
    vals.className = "mc-vals";
    vals.placeholder = "REAL, BUDGET";
    vals.value = (c.values || []).join(", ");
    const rm = document.createElement("button");
    rm.type = "button";
    rm.className = "attr-remove-btn";
    rm.textContent = "\u2715";
    rm.title = "Quitar condición";
    rm.addEventListener("click", () => row.remove());
    row.append(col, op, vals, rm);
    box.appendChild(row);
}

/** Comprueba la fórmula (sintaxis, medidas existentes, sin ciclos). Devuelve el error o "". */
function checkMeasureFormula(field, formula) {
    const byName = new Map(fieldsState.filter(f => f.type === "MEASURE" && f.enabled).map(f => [f.name.toUpperCase(), f]));
    const visit = (f, stack) => {
        const key = f.name.toUpperCase();
        if (stack.indexOf(key) !== -1) throw new Error("La fórmula se usa a sí misma: " + stack.concat(key).join(" → "));
        const isSelf = key === String(field.name).toUpperCase();
        if (!isSelf && f.measureKind !== "CALC") return;
        // La medida que se está editando usa la fórmula del cuadro, no la guardada
        window.DracoMeasureFormula.toSql(isSelf ? formula : f.formula, (ref) => {
            const m = byName.get(String(ref).toUpperCase());
            if (!m) throw new Error(`No hay ninguna medida «${ref}» (o está deshabilitada).`);
            visit(m, stack.concat(key));
            return "1";
        });
    };
    try { visit(field, []); return ""; }
    catch (e) { return e.message || String(e); }
}

function validateMeasureFormulaUI() {
    const field = fieldsState[currentConfigFieldIndex];
    const msg = document.getElementById("modalMeasureFormulaMsg");
    if (!field || !msg) return "";
    const formula = document.getElementById("modalMeasureFormula").value.trim();
    const err = formula ? checkMeasureFormula(field, formula) : "Escribe la fórmula.";
    msg.className = "measure-formula-msg " + (err ? "err" : "ok");
    msg.textContent = err || "Fórmula correcta.";
    return err;
}

/** Guarda lo propio de una medida calculada/filtrada. false = hay errores. */
function saveVirtualMeasureSections(field) {
    if (!isVirtualMeasure(field)) return true;
    if (field.measureKind === "CALC") {
        const err = validateMeasureFormulaUI();
        if (err) { showToast(err, "error"); return false; }
        field.formula = document.getElementById("modalMeasureFormula").value.trim();
        return true;
    }
    const base = document.getElementById("modalMeasureBase").value;
    const conds = [...document.querySelectorAll("#modalMeasureConds .measure-cond-row")].map(r => {
        const column = r.querySelector(".mc-col").value;
        const colField = fieldsState.find(f => f.name === column);
        return {
            column,
            dataType: colField ? colField.dataType : "STRING",
            op: r.querySelector(".mc-op").value,
            values: r.querySelector(".mc-vals").value.split(",").map(v => v.trim()).filter(v => v !== "")
        };
    }).filter(c => c.column && c.values.length);
    if (!base) { showToast("Elige la medida base.", "error"); return false; }
    if (!conds.length) { showToast("Añade al menos una condición con valores.", "error"); return false; }
    field.base = base;
    field.conditions = conds;
    return true;
}

/* =====================================================================
 * MIEMBROS CALCULADOS DEL MODELO
 * Misma ventana «Cálculos» que en el informe (calcDialog.html, vía
 * DracoCalcEditor.openFor), pero se guardan en la dimensión del modelo:
 * field.calcMembers[ATRIBUTO] = [cálculos]. El motor los junta con los del
 * informe (ver attachDracoCalcs en commands.js); se usan los de tipo «por
 * miembro» y, si el informe tiene uno con el mismo nombre, manda el del informe.
 * ===================================================================== */
function openModelCalcMembers(field, attrName) {
    if (!window.DracoCalcEditor) return;
    const U = String(attrName).toUpperCase();
    window.DracoCalcEditor.openFor({
        dimension: field.name,
        attribute: attrName,
        attributes: [attrName],
        getCalcs: () => ((field.calcMembers || {})[U] || []),
        saveCalcs: (a, list) => {
            field.calcMembers = field.calcMembers || {};
            const k = String(a).toUpperCase();
            // Solo «por miembro»: los «por posición» dependen de cómo se vea cada informe
            const clean = (list || []).filter(c => c && c.mode !== "position");
            if ((list || []).length !== clean.length) showToast("En el modelo solo se guardan los cálculos «por miembro».", "error");
            if (clean.length) field.calcMembers[k] = clean; else delete field.calcMembers[k];
            renderAttributesTable(field.attributes);
        },
        loadMembers: (a) => loadModelAttrMembers(field, a)
    });
}

/** Valores distintos de un atributo, consultando su tabla (para la ventana de cálculos). */
async function loadModelAttrMembers(field, attrName) {
    try {
        const project = field.relProject || document.getElementById("factProject").value;
        const dataset = field.relDataset || document.getElementById("factDataset").value;
        const table = field.relTable || document.getElementById("factTable").value;
        const attr = (field.attributes || []).find(a => a.name === attrName);
        let expr = attrName;
        if (attr && attr.calc && window.DracoAttrFormula) expr = window.DracoAttrFormula.toSql(attr.expression, c => c, Provider.key ? Provider.key() : "bigquery");
        if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(attrName)) throw new Error("Nombre de atributo no válido.");
        const sql = Provider.limit("SELECT DISTINCT " + expr + " AS V FROM " + Provider.qualify(project, dataset, table) + " ORDER BY 1", 5000);
        const { rows } = await Provider.runQuery(sql, project, dataset);
        const members = Array.from(new Set((rows || []).map(r => r.V !== undefined ? r.V : r.v).filter(v => v !== null && v !== undefined).map(String)));
        return { members, visible: null, error: null };
    } catch (e) {
        return { members: [], visible: null, error: (e && e.message) || String(e) };
    }
}

/* =====================================================================
 * ATRIBUTOS CALCULADOS (columnas nuevas de una dimensión calculadas con
 * otras columnas de SU tabla; ver DracoAttrFormula en calcFormula.js)
 * ===================================================================== */
function openCalcAttrDialog(field, attr) {
    if (!field) return;
    const cols = (field.attributes || []).filter(a => !a.calc).map(a => a.name);
    const back = document.createElement("div");
    back.className = "sm-dialog-backdrop";
    const box = document.createElement("div");
    box.className = "sm-dialog";
    box.style.maxWidth = "420px";
    const h = document.createElement("h3");
    h.textContent = attr ? "Editar atributo calculado" : "Nuevo atributo calculado";
    const name = document.createElement("input");
    name.type = "text"; name.placeholder = "Nombre (p. ej. SOC_CECO)"; name.value = attr ? attr.name : "";
    if (attr) name.disabled = true; // el nombre se usa en jerarquías e informes
    const fx = document.createElement("textarea");
    fx.rows = 3; fx.spellcheck = false; fx.placeholder = 'CONCATENAR([SOC]; "-"; [CECO])';
    fx.value = attr ? attr.expression || "" : "";
    const chips = document.createElement("div");
    chips.className = "measure-ref-chips";
    cols.forEach(c => {
        const b = document.createElement("button");
        b.type = "button"; b.textContent = c;
        b.addEventListener("click", () => {
            const a = fx.selectionStart != null ? fx.selectionStart : fx.value.length;
            const z = fx.selectionEnd != null ? fx.selectionEnd : fx.value.length;
            fx.value = fx.value.slice(0, a) + "[" + c + "]" + fx.value.slice(z);
            fx.focus(); validate();
        });
        chips.appendChild(b);
    });
    const hint = document.createElement("div");
    hint.className = "display-hint";
    hint.textContent = "Columnas entre corchetes y textos entre comillas. Texto: CONCATENAR(a; b; …) o a & b, IZQUIERDA(t; n), DERECHA(t; n), "
        + "EXTRAE(t; inicio; n), MAYUS, MINUS, ESPACIOS, SUSTITUIR(t; buscar; poner), LARGO. Número: REDONDEAR, TRUNCAR, ABS. "
        + "Fecha: AÑO, MES, TRIMESTRE, DIA, INICIO_MES, INICIO_AÑO, TEXTO_FECHA(f; \"AAAA-MM\"). SI(condición; a; b). "
        + "Grupos: AGRUPAR([CUENTA]; \"60*,61*\"; \"Personal\"; \"62*\"; \"Servicios\"; \"Otros\") y TRAMOS([IMPORTE]; 1000; 5000).";
    const msg = document.createElement("div");
    msg.className = "measure-formula-msg";
    const actions = document.createElement("div");
    actions.className = "sm-dialog-actions";
    const cancel = document.createElement("button");
    cancel.type = "button"; cancel.className = "btn-secondary btn-sm"; cancel.textContent = "Cancelar";
    const ok = document.createElement("button");
    ok.type = "button"; ok.className = "btn-primary btn-sm"; ok.textContent = attr ? "Guardar" : "Crear";
    actions.append(cancel, ok);
    box.append(h, name, fx, chips, hint, msg, actions);
    back.appendChild(box);
    document.body.appendChild(back);
    (attr ? fx : name).focus();

    const taken = new Set((field.attributes || []).filter(a => a !== attr).map(a => a.name.toUpperCase()));
    function validate() {
        let err = "";
        const n = name.value.trim(), e = fx.value.trim();
        if (!n) err = "Pon un nombre.";
        else if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(n)) err = "Nombre: solo letras, números y _ (sin espacios ni acentos).";
        else if (taken.has(n.toUpperCase())) err = "Ya hay un atributo con ese nombre.";
        else if (!e) err = "Escribe la fórmula.";
        else {
            try {
                window.DracoAttrFormula.toSql(e, c => c, "bigquery");
                const unknown = window.DracoAttrFormula.columns(e).filter(c => !cols.some(x => x.toUpperCase() === c.toUpperCase()));
                if (unknown.length) err = "Columnas que no están en esta tabla: " + unknown.join(", ");
            } catch (ex) { err = ex.message; }
        }
        msg.className = "measure-formula-msg " + (err ? "err" : "ok");
        msg.textContent = err || "Fórmula correcta.";
        ok.disabled = !!err;
        return !err;
    }
    name.addEventListener("input", validate);
    fx.addEventListener("input", validate);
    cancel.addEventListener("click", () => back.remove());
    back.addEventListener("keydown", (e) => { if (e.key === "Escape") back.remove(); });
    ok.addEventListener("click", () => {
        if (!validate()) return;
        const expression = fx.value.trim();
        const dataType = window.DracoAttrFormula.resultType(expression);
        if (attr) {
            attr.expression = expression;
            attr.dataType = dataType;
        } else {
            field.attributes = field.attributes || [];
            field.attributes.push({ name: name.value.trim(), alias: name.value.trim(), dataType, isKey: false, enabled: true, calc: true, expression });
        }
        back.remove();
        renderAttributesTable(field.attributes);
        renderHierarchiesList();
    });
    validate();
}

/* ---------- cambios sin guardar (aviso al ir a «Editar informes») ---------- */

let smSavedSnapshot = null;

function smSnapshot() {
    try {
        return JSON.stringify({
            model: currentModel || "",
            p: (document.getElementById("factProject") || {}).value || "",
            d: (document.getElementById("factDataset") || {}).value || "",
            t: (document.getElementById("factTable") || {}).value || "",
            f: fieldsState
        });
    } catch (e) { return ""; }
}

function smMarkSaved() { smSavedSnapshot = smSnapshot(); }

function smIsDirty() {
    if (!fieldsState || !fieldsState.length) return false;
    return smSavedSnapshot === null || smSnapshot() !== smSavedSnapshot;
}

window.SemanticModelEditor = {
    isDirty: smIsDirty,
    /** Aviso al salir con cambios sin guardar: resuelve true si se puede salir. */
    confirmLeave: async function () {
        const r = await smDialog({
            title: "Cambios sin guardar",
            text: "Has modificado el modelo semántico y no lo has guardado. Si sales ahora, los informes seguirán usando la versión guardada.",
            buttons: [
                { label: "Seguir editando", value: "stay", primary: true },
                { label: "Salir sin guardar", value: "leave" }
            ]
        });
        return r === "leave";
    }
};

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

/** Selector del atributo de texto en cada nivel de jerarquía (oculto de momento). */
const SHOW_LEVEL_TEXT_SELECTOR = false;

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

        // SEGURIDAD: hier.name y levelNames (alias/nombre de atributos) son
        // texto libre / potencialmente importado -> DOM API en vez de
        // template literal HTML.
        const infoDiv = document.createElement("div");
        infoDiv.className = "hierarchy-info";
        const nameDiv = document.createElement("div");
        nameDiv.className = "hierarchy-name";
        nameDiv.textContent = hier.name;
        const summaryDiv = document.createElement("div");
        summaryDiv.className = "hierarchy-levels-summary";
        summaryDiv.textContent = levelNames || "(sin niveles)";
        infoDiv.appendChild(nameDiv);
        infoDiv.appendChild(summaryDiv);

        const actionsDiv = document.createElement("div");
        actionsDiv.className = "hierarchy-chip-actions";
        const editBtn = document.createElement("button");
        editBtn.type = "button";
        editBtn.className = "hierarchy-chip-btn";
        editBtn.title = "Editar jerarquía";
        editBtn.textContent = "\u270e"; // ✎
        editBtn.addEventListener("click", (e) => { e.stopPropagation(); openHierarchyEditor(idx); });
        const delBtn = document.createElement("button");
        delBtn.type = "button";
        delBtn.className = "hierarchy-chip-btn";
        delBtn.title = "Eliminar jerarquía";
        delBtn.textContent = "\ud83d\uddd1"; // 🗑
        delBtn.addEventListener("click", (e) => { e.stopPropagation(); deleteHierarchyAt(idx); });
        actionsDiv.appendChild(editBtn);
        actionsDiv.appendChild(delBtn);

        row.appendChild(infoDiv);
        row.appendChild(actionsDiv);

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
            levels: existing.levels.map(l => ({ attribute: l.attribute, textAttribute: l.textAttribute || "" }))
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

            // SEGURIDAD: attr.alias/attr.name no son de confianza -> DOM API.
            const nameSpan = document.createElement("span");
            nameSpan.className = "hierarchy-level-name";
            nameSpan.textContent = attr.alias || attr.name;
            const addSpan = document.createElement("span");
            addSpan.className = "hierarchy-chip-btn";
            addSpan.textContent = "\uff0b"; // ＋
            chip.appendChild(nameSpan);
            chip.appendChild(addSpan);

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

            // SEGURIDAD: attr.alias/attr.name/lvl.attribute no son de
            // confianza -> DOM API en vez de innerHTML.
            const nameSpan2 = document.createElement("span");
            nameSpan2.className = "hierarchy-level-name";
            nameSpan2.textContent = attr ? (attr.alias || attr.name) : lvl.attribute;
            const removeBtn = document.createElement("button");
            removeBtn.type = "button";
            removeBtn.className = "hierarchy-chip-btn hierarchy-chip-btn-remove";
            removeBtn.title = "Quitar del nivel";
            removeBtn.textContent = "\u2715"; // ✕
            removeBtn.addEventListener("click", (e) => { e.stopPropagation(); removeHierarchyLevel(idx); });
            chip.appendChild(nameSpan2);

            // "ID y texto": atributo de texto de este nivel (no puede ser otro
            // nivel de la jerarquía ni el texto de otro nivel).
            // OCULTO de momento (SHOW_LEVEL_TEXT_SELECTOR): se pensará cómo
            // presentarlo. Los textos ya guardados en los niveles se conservan.
            if (SHOW_LEVEL_TEXT_SELECTOR) {
                const taken = hierEditState.levels
                    .filter((l, j) => j !== idx).map(l => l.textAttribute).filter(Boolean)
                    .concat(hierEditState.levels.map(l => l.attribute));
                const textSel = document.createElement("select");
                textSel.className = "hierarchy-level-text-select";
                textSel.title = "Atributo de texto de este nivel";
                textSel.draggable = false;
                const noneOpt = document.createElement("option");
                noneOpt.value = "";
                noneOpt.textContent = "(sin texto)";
                textSel.appendChild(noneOpt);
                textAttributeCandidates(field, taken).forEach(a => {
                    const o = document.createElement("option");
                    o.value = a.name;
                    o.textContent = a.alias || a.name;
                    textSel.appendChild(o);
                });
                textSel.value = lvl.textAttribute || "";
                ["click", "mousedown"].forEach(ev => textSel.addEventListener(ev, (e) => e.stopPropagation()));
                textSel.addEventListener("change", () => { lvl.textAttribute = textSel.value; renderHierarchyEditor(); });
                chip.appendChild(textSel);
            }

            chip.appendChild(removeBtn);

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
        hierEditState.levels.push({ attribute: attrName, textAttribute: "" });
    } else {
        hierEditState.levels.splice(atIndex, 0, { attribute: attrName, textAttribute: "" });
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
        levels: hierEditState.levels.map(l => (l.textAttribute
            ? { attribute: l.attribute, textAttribute: l.textAttribute }
            : { attribute: l.attribute }))
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

        // La marca de Draco Planning se conserva solo si la tabla de hechos
        // sigue siendo la misma que generó Planning.
        const previous = window.SemanticModelStore.getModel(currentModel);
        const sameFact = !!(previous && previous.fact
            && previous.fact.project === fact.project
            && previous.fact.dataset === fact.dataset
            && previous.fact.table === fact.table);
        const toSave = { fact: fact, fields: fieldsState };
        if (previous && previous.dracoPlanning === true && sameFact) toSave.dracoPlanning = true;
        await window.SemanticModelStore.saveModel(currentModel, toSave);
        smMarkSaved();

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
