/**
 * Diálogo independiente "Guardar modelo semántico (.lkml)".
 * ------------------------------------------------------------------------
 * Se abre con Office.context.ui.displayDialogAsync directamente desde el
 * botón del ribbon "Guardar modelo semántico" (ver js/commands.js) o desde
 * el botón "📤 LookML" del taskpane (ver semantic_model.js), en ambos casos
 * a través de js/lkmlSaveBridge.js: es la MISMA página independiente, no un
 * popup superpuesto dentro del taskpane.
 *
 * En ambas pestañas, "Guardar":
 *   1) Genera el LookML del modelo elegido (tabla de hechos, dimensiones,
 *      atributos, jerarquías, measures y relaciones) con LkmlExport.
 *   2) Guarda de verdad el fichero:
 *        - Local:    lo descarga al equipo del usuario.
 *        - Servidor: hace commit del fichero en el repositorio Git
 *                     (GitHub/GitLab) de la conexión elegida, vía GitRepo.
 *   3) Envía ese mismo LookML a quien abrió el diálogo con
 *      Office.context.ui.messageParent(), porque este diálogo NO tiene
 *      acceso al modelo de objetos de Excel. Es js/lkmlSaveBridge.js, en el
 *      lado que sí tiene ese acceso, quien lo escribe en EDIT_REPORT!G1 y
 *      cierra el diálogo.
 *
 * Los modelos semánticos tampoco están disponibles aquí directamente (no
 * hay Office.context.document.settings en un diálogo): quien abre esta
 * página se los pasa ya "aplanados" en JSON por querystring
 * (?models=...&active=...).
 */

let modelsData = {};

function showToast(message, type = "success", duration = 4000) {

    const container = document.getElementById("appToastContainer");

    if (!container) {
        alert(message);
        return;
    }

    const toast = document.createElement("div");
    toast.className = `app-toast ${type === "error" ? "error" : "success"}`;
    toast.textContent = (type === "error" ? "⚠ " : "✔ ") + message;

    container.appendChild(toast);

    requestAnimationFrame(() => toast.classList.add("visible"));

    setTimeout(() => {
        toast.classList.remove("visible");
        setTimeout(() => toast.remove(), 250);
    }, duration);

}

function closeDialog() {
    try {
        if (Office.context.ui && typeof Office.context.ui.closeContainer === "function") {
            Office.context.ui.closeContainer();
            return;
        }
    } catch (error) {
        console.error("Error al cerrar el diálogo con closeContainer:", error);
    }
    // Alternativa por si closeContainer no está disponible en este host:
    // el diálogo también es una ventana normal, así que window.close()
    // debería cerrarla igualmente.
    try {
        window.close();
    } catch (error2) {
        console.error("Error al cerrar la ventana del diálogo:", error2);
    }
}

function parseQueryParams() {
    const params = new URLSearchParams(window.location.search);

    try {
        modelsData = JSON.parse(decodeURIComponent(params.get("models") || "{}")) || {};
    } catch (e) {
        console.error("No se han podido leer los modelos semánticos recibidos:", e);
        modelsData = {};
    }

    return { active: params.get("active") || "" };
}

function getModel(modelName) {
    return modelsData[modelName] || null;
}

function ensureLkmlExtension(fileName) {
    fileName = (fileName || "").trim();
    if (!fileName) return fileName;
    return fileName.toLowerCase().endsWith(".lkml") ? fileName : `${fileName}.lkml`;
}

function joinRepoPath(folder, fileName) {
    const cleanFolder = (folder || "").replace(/^\/+|\/+$/g, "");
    return cleanFolder ? `${cleanFolder}/${fileName}` : fileName;
}

function populateModelSelect(activeModel) {
    const select = document.getElementById("lkmlSaveModelSelect");
    if (!select) return;

    const models = Object.keys(modelsData).sort();

    select.innerHTML = '<option value="">— Sin modelo seleccionado —</option>';

    models.forEach(model => {
        const option = document.createElement("option");
        option.value = model;
        option.text = model;
        select.appendChild(option);
    });

    if (activeModel && models.includes(activeModel)) {
        select.value = activeModel;
    }
}

function setActiveTab(tabId) {
    document.getElementById("lkmlSaveTabServer").classList.toggle("active", tabId === "server");
    document.getElementById("lkmlSaveTabLocal").classList.toggle("active", tabId === "local");
    document.getElementById("lkmlSavePanelServer").classList.toggle("active", tabId === "server");
    document.getElementById("lkmlSavePanelLocal").classList.toggle("active", tabId === "local");
    updateConfirmButtonState();
}

function getActiveTab() {
    return document.getElementById("lkmlSaveTabLocal").classList.contains("active") ? "local" : "server";
}

// Rellena el selector de conexiones con las que tienen repositorio de
// modelos semánticos configurado (GitHub/GitLab).
function populateConnectionSelect() {
    const select = document.getElementById("lkmlSaveServerConnection");
    if (!select) return;

    const connections = (Connections && typeof Connections.list === "function")
        ? Connections.list() : [];
    const withRepo = connections.filter(c => {
        const repo = c.config && c.config.semanticRepo;
        return repo && (repo.type === "github" || repo.type === "gitlab") && repo.url;
    });

    select.innerHTML = "";
    if (withRepo.length === 0) {
        select.innerHTML = "<option value=\"\">— Ninguna conexión tiene repositorio configurado —</option>";
        select.disabled = true;
        return;
    }

    select.disabled = false;
    select.innerHTML = "<option value=\"\">— Selecciona una conexión —</option>";
    withRepo.forEach(conn => {
        const opt = document.createElement("option");
        opt.value = conn.id;
        opt.textContent = `${conn.name} (${conn.config.semanticRepo.type === "github" ? "GitHub" : "GitLab"})`;
        select.appendChild(opt);
    });

    const activeId = Connections.getActiveId ? Connections.getActiveId() : null;
    if (activeId && withRepo.some(c => c.id === activeId)) select.value = activeId;
}

// Lista el contenido de la carpeta actual del repositorio (navegación de
// carpetas). Al hacer clic en un .lkml existente se propone como nombre de
// fichero (para sobrescribirlo al guardar).
async function updateServerList() {
    const list = document.getElementById("lkmlSaveServerList");
    if (!list) return;

    const connectionId = document.getElementById("lkmlSaveServerConnection").value;
    const path = document.getElementById("lkmlSaveServerPath").value.trim();

    if (!connectionId) {
        list.classList.add("is-empty");
        list.innerHTML = "<span class=\"lkml-empty-hint\">Selecciona una conexión para explorar las carpetas del repositorio.</span>";
        return;
    }

    const conn = Connections.getById(connectionId);
    const repoConfig = conn && conn.config && conn.config.semanticRepo;
    if (!repoConfig) return;

    list.classList.add("is-empty");
    list.innerHTML = "<span class=\"lkml-empty-hint\">Cargando…</span>";

    try {
        const items = await GitRepo.listContents(repoConfig, path);

        if (items.length === 0) {
            list.classList.add("is-empty");
            list.innerHTML = "<span class=\"lkml-empty-hint\">Esta carpeta no tiene subcarpetas ni ficheros .lkml.</span>";
            return;
        }

        list.classList.remove("is-empty");
        list.innerHTML = "";
        items.forEach(item => {
            const row = document.createElement("div");
            row.className = "lkml-server-item";
            row.innerHTML = `<span>${item.type === "dir" ? "📁" : "📄"}</span><span>${item.name}</span>`;

            if (item.type === "dir") {
                row.addEventListener("click", () => {
                    document.getElementById("lkmlSaveServerPath").value = item.path;
                    updateServerList();
                });
            } else {
                row.addEventListener("click", () => {
                    list.querySelectorAll(".lkml-server-item.selected").forEach(el => el.classList.remove("selected"));
                    row.classList.add("selected");
                    document.getElementById("lkmlSaveServerFileName").value = item.name;
                    updateConfirmButtonState();
                });
            }
            list.appendChild(row);
        });
    } catch (err) {
        console.error("Error al listar el repositorio:", err);
        list.classList.add("is-empty");
        list.innerHTML = `<span class="lkml-empty-hint">${err.message || "Error al listar el repositorio."}</span>`;
    }
}

// Propone un nombre de fichero (<modelo>.lkml) en ambas pestañas al elegir
// un modelo, si el usuario todavía no ha escrito uno propio.
function prefillFileNames() {
    const modelName = document.getElementById("lkmlSaveModelSelect").value;
    if (!modelName) return;

    const suggested = `${modelName}.lkml`;

    const localInput = document.getElementById("lkmlSaveLocalFileName");
    if (localInput && localInput.value.trim() === "") localInput.value = suggested;

    const serverInput = document.getElementById("lkmlSaveServerFileName");
    if (serverInput && serverInput.value.trim() === "") serverInput.value = suggested;
}

function updateConfirmButtonState() {
    const btn = document.getElementById("btnConfirmSaveLkml");
    if (!btn) return;

    const modelName = document.getElementById("lkmlSaveModelSelect").value;
    if (!modelName) {
        btn.disabled = true;
        return;
    }

    if (getActiveTab() === "local") {
        const fileName = document.getElementById("lkmlSaveLocalFileName").value.trim();
        btn.disabled = fileName === "";
        return;
    }

    const connectionId = document.getElementById("lkmlSaveServerConnection").value;
    const fileName = document.getElementById("lkmlSaveServerFileName").value.trim();
    btn.disabled = !(connectionId && fileName !== "");
}

function setSaving(isSaving, label) {
    const btn = document.getElementById("btnConfirmSaveLkml");
    const cancelBtn = document.getElementById("btnCancelSaveLkml");
    if (!btn) return;

    btn.textContent = isSaving ? (label || "Guardando…") : "Guardar";
    if (cancelBtn) cancelBtn.disabled = isSaving;
    if (isSaving) {
        btn.disabled = true;
    } else {
        updateConfirmButtonState();
    }
}

// Envía el LookML ya guardado (local o en el repositorio) a quien abrió el
// diálogo, para que lo escriba en EDIT_REPORT!G1 (este diálogo no tiene
// acceso a Excel). El diálogo se queda a la espera: lo cierra el padre al
// terminar, o nos avisa del error si falla (ver el handler más abajo).
function sendContentToParentForG1(modelName, content) {
    setSaving(true, "Escribiendo en Excel…");
    Office.context.ui.messageParent(JSON.stringify({ modelName, content }));
}

function initEvents() {

    document.getElementById("btnCancelSaveLkml").addEventListener("click", closeDialog);

    document.getElementById("lkmlSaveTabServer").addEventListener("click", () => setActiveTab("server"));
    document.getElementById("lkmlSaveTabLocal").addEventListener("click", () => setActiveTab("local"));

    document.getElementById("lkmlSaveModelSelect").addEventListener("change", () => {
        prefillFileNames();
        updateConfirmButtonState();
    });

    document.getElementById("lkmlSaveServerConnection").addEventListener("change", () => {
        document.getElementById("lkmlSaveServerPath").value = "";
        updateServerList();
        updateConfirmButtonState();
    });

    document.getElementById("lkmlSaveServerUpBtn").addEventListener("click", () => {
        const pathInput = document.getElementById("lkmlSaveServerPath");
        const parts = pathInput.value.split("/").filter(Boolean);
        parts.pop();
        pathInput.value = parts.join("/");
        updateServerList();
    });

    document.getElementById("lkmlSaveServerPath").addEventListener("keydown", (e) => {
        if (e.key !== "Enter") return;
        updateServerList();
    });

    document.getElementById("lkmlSaveServerFileName").addEventListener("input", updateConfirmButtonState);
    document.getElementById("lkmlSaveLocalFileName").addEventListener("input", updateConfirmButtonState);

    document.getElementById("btnConfirmSaveLkml").addEventListener("click", async () => {

        const modelName = document.getElementById("lkmlSaveModelSelect").value;
        if (!modelName) return;

        let content;
        try {
            content = LkmlExport.buildContent(modelName, getModel(modelName));
        } catch (err) {
            console.error("Error al generar el LookML:", err);
            showToast("Error al generar el LookML: " + err.message, "error");
            return;
        }

        if (getActiveTab() === "local") {

            const fileName = ensureLkmlExtension(document.getElementById("lkmlSaveLocalFileName").value);
            if (!fileName) return;

            try {
                LkmlExport.downloadTextFile(fileName, content);
            } catch (err) {
                console.error("Error al descargar el fichero .lkml:", err);
                showToast("Error al descargar el fichero: " + err.message, "error");
                return;
            }

            sendContentToParentForG1(modelName, content);
            return;
        }

        // Pestaña Servidor: commit real del fichero en el repositorio.
        const connectionId = document.getElementById("lkmlSaveServerConnection").value;
        const folder = document.getElementById("lkmlSaveServerPath").value.trim();
        const fileName = ensureLkmlExtension(document.getElementById("lkmlSaveServerFileName").value);
        if (!connectionId || !fileName) return;

        const conn = Connections.getById(connectionId);
        const repoConfig = conn && conn.config && conn.config.semanticRepo;
        if (!repoConfig) {
            showToast("La conexión elegida no tiene un repositorio configurado.", "error");
            return;
        }

        const fullPath = joinRepoPath(folder, fileName);

        setSaving(true, "Guardando en el repositorio…");
        try {
            await GitRepo.putFile(repoConfig, fullPath, content, `Actualiza ${fullPath} desde el editor de modelos semánticos`);
        } catch (err) {
            console.error("Error al guardar en el repositorio:", err);
            setSaving(false);
            showToast("Error al guardar en el repositorio: " + err.message, "error");
            return;
        }

        sendContentToParentForG1(modelName, content);

    });

    // Si quien nos abrió no ha podido escribir en EDIT_REPORT!G1 (p.ej. la
    // hoja no existe o el libro está protegido), nos lo dice aquí para que
    // el usuario lo sepa y pueda reintentar en vez de quedarse sin saber
    // qué ha pasado. El fichero (local o en el repositorio) ya se ha
    // guardado en ese punto, así que solo se avisa del fallo al escribir
    // en Excel.
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {

        let payload = null;
        try {
            payload = JSON.parse(arg.message);
        } catch (e) {
            payload = null;
        }

        setSaving(false);
        showToast(
            (payload && payload.error)
                ? `El fichero se ha guardado, pero no se ha podido escribir en Excel: ${payload.error}`
                : "El fichero se ha guardado, pero no se ha podido escribir en Excel.",
            "error"
        );

    });

}

Office.onReady(() => {
    const { active } = parseQueryParams();

    initEvents();
    populateModelSelect(active);
    setActiveTab("server");
    populateConnectionSelect();
    updateServerList();
    prefillFileNames();
    updateConfirmButtonState();
});
