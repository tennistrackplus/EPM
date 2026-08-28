/**
 * Diálogo independiente "Guardar modelo semántico (.lkml)".
 * ------------------------------------------------------------------------
 * Se abre con Office.context.ui.displayDialogAsync directamente desde el
 * botón del ribbon "Guardar modelo semántico" (ver commands.js) o desde el
 * botón "📤 LookML" del taskpane (ver semantic_model.js): en ambos casos es
 * la MISMA página independiente, no un popup superpuesto dentro del
 * taskpane.
 *
 * Esta página NO tiene acceso a Office.context.document.settings (los
 * diálogos de Office no comparten el modelo de objetos del documento con
 * el taskpane, igual que dataPreview.html / hierarchyPreview.html), así
 * que no puede leer SemanticModelStore directamente. Por eso, quien abre
 * el diálogo (commands.js o semantic_model.js, que sí tienen acceso)
 * le pasa los modelos ya "aplanados" en JSON por querystring
 * (?models=...&active=...). A partir de ahí esta página es autónoma.
 */

let modelsData = {};
let lkmlSaveSelection = null; // { source: "local"|"server", connectionId?, path? }

function showToast(message, type = "success", duration = 3500) {

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
        Office.context.ui.closeContainer();
    } catch (error) {
        console.error("Error al cerrar el diálogo:", error);
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

function populateLkmlSaveModelSelect(activeModel) {
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

function setLkmlSaveActiveTab(tabId) {
    document.getElementById("lkmlSaveTabServer").classList.toggle("active", tabId === "server");
    document.getElementById("lkmlSaveTabLocal").classList.toggle("active", tabId === "local");
    document.getElementById("lkmlSavePanelServer").classList.toggle("active", tabId === "server");
    document.getElementById("lkmlSavePanelLocal").classList.toggle("active", tabId === "local");
    updateSaveLkmlConfirmButtonState();
}

// Rellena el selector de conexiones con las que tienen repositorio de
// modelos semánticos configurado (mismo criterio que en "Abrir").
function populateLkmlSaveConnectionSelect() {
    const select = document.getElementById("lkmlSaveServerConnection");
    if (!select) return;

    const connections = (window.Connections && typeof window.Connections.list === "function")
        ? window.Connections.list() : [];
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

    const activeId = window.Connections.getActiveId ? window.Connections.getActiveId() : null;
    if (activeId && withRepo.some(c => c.id === activeId)) select.value = activeId;
}

// Lista el contenido de la carpeta actual del repositorio (navegación de
// carpetas, igual que en "Abrir"). Aquí es solo para elegir el destino: al
// hacer clic en un .lkml existente se propone como nombre de fichero (para
// sobrescribirlo), pero no se lee su contenido.
async function updateLkmlSaveServerList() {
    const list = document.getElementById("lkmlSaveServerList");
    if (!list) return;

    const connectionId = document.getElementById("lkmlSaveServerConnection").value;
    const path = document.getElementById("lkmlSaveServerPath").value.trim();

    if (!connectionId) {
        list.classList.add("is-empty");
        list.innerHTML = "<span class=\"lkml-empty-hint\">Selecciona una conexión para explorar las carpetas del repositorio.</span>";
        return;
    }

    const conn = window.Connections.getById(connectionId);
    const repoConfig = conn && conn.config && conn.config.semanticRepo;
    if (!repoConfig) return;

    list.classList.add("is-empty");
    list.innerHTML = "<span class=\"lkml-empty-hint\">Cargando…</span>";

    try {
        const items = await window.GitRepo.listContents(repoConfig, path);

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
                    updateLkmlSaveServerList();
                });
            } else {
                row.addEventListener("click", () => {
                    list.querySelectorAll(".lkml-server-item.selected").forEach(el => el.classList.remove("selected"));
                    row.classList.add("selected");
                    document.getElementById("lkmlSaveServerFileName").value = item.name;
                    updateSaveLkmlConfirmButtonState();
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
function prefillLkmlSaveFileNames() {
    const modelName = document.getElementById("lkmlSaveModelSelect").value;
    if (!modelName) return;

    const suggested = `${modelName}.lkml`;

    const localInput = document.getElementById("lkmlSaveLocalFileName");
    if (localInput && localInput.value.trim() === "") localInput.value = suggested;

    const serverInput = document.getElementById("lkmlSaveServerFileName");
    if (serverInput && serverInput.value.trim() === "") serverInput.value = suggested;
}

function updateSaveLkmlConfirmButtonState() {
    const btn = document.getElementById("btnConfirmSaveLkml");
    if (!btn) return;

    const modelName = document.getElementById("lkmlSaveModelSelect").value;
    const activeTab = document.getElementById("lkmlSaveTabLocal").classList.contains("active") ? "local" : "server";

    if (!modelName) {
        btn.disabled = true;
        return;
    }

    if (activeTab === "local") {
        const fileName = document.getElementById("lkmlSaveLocalFileName").value.trim();
        btn.disabled = fileName === "";
        return;
    }

    const connectionId = document.getElementById("lkmlSaveServerConnection").value;
    const fileName = document.getElementById("lkmlSaveServerFileName").value.trim();
    btn.disabled = !(connectionId && fileName !== "");
}

function initEvents() {

    document.getElementById("btnCancelSaveLkml").addEventListener("click", closeDialog);

    document.getElementById("lkmlSaveTabServer").addEventListener("click", () => setLkmlSaveActiveTab("server"));
    document.getElementById("lkmlSaveTabLocal").addEventListener("click", () => setLkmlSaveActiveTab("local"));

    document.getElementById("lkmlSaveModelSelect").addEventListener("change", () => {
        prefillLkmlSaveFileNames();
        updateSaveLkmlConfirmButtonState();
    });

    document.getElementById("lkmlSaveServerConnection").addEventListener("change", () => {
        document.getElementById("lkmlSaveServerPath").value = "";
        updateLkmlSaveServerList();
        updateSaveLkmlConfirmButtonState();
    });

    document.getElementById("lkmlSaveServerUpBtn").addEventListener("click", () => {
        const pathInput = document.getElementById("lkmlSaveServerPath");
        const parts = pathInput.value.split("/").filter(Boolean);
        parts.pop();
        pathInput.value = parts.join("/");
        updateLkmlSaveServerList();
    });

    document.getElementById("lkmlSaveServerPath").addEventListener("keydown", (e) => {
        if (e.key !== "Enter") return;
        updateLkmlSaveServerList();
    });

    document.getElementById("lkmlSaveServerFileName").addEventListener("input", updateSaveLkmlConfirmButtonState);
    document.getElementById("lkmlSaveLocalFileName").addEventListener("input", updateSaveLkmlConfirmButtonState);

    // "Guardar": en Local descarga el .lkml al equipo; en Servidor, por
    // ahora, solo confirma el destino elegido (el commit/push real al
    // repositorio se conectará en un paso posterior).
    document.getElementById("btnConfirmSaveLkml").addEventListener("click", () => {

        const modelName = document.getElementById("lkmlSaveModelSelect").value;
        if (!modelName) return;

        const activeTab = document.getElementById("lkmlSaveTabLocal").classList.contains("active") ? "local" : "server";

        if (activeTab === "local") {

            let fileName = document.getElementById("lkmlSaveLocalFileName").value.trim();
            if (!fileName) return;
            if (!fileName.toLowerCase().endsWith(".lkml")) fileName += ".lkml";

            try {
                const content = LkmlExport.buildContent(modelName, getModel(modelName));
                LkmlExport.downloadTextFile(fileName, content);
                showToast(`Fichero "${fileName}" descargado.`, "success");
                setTimeout(closeDialog, 900);
            } catch (err) {
                console.error("Error al generar el LookML:", err);
                showToast("Error al generar el LookML: " + err.message, "error");
            }

            return;
        }

        // Pestaña Servidor: todavía no hay commit/push real al repositorio.
        const connectionId = document.getElementById("lkmlSaveServerConnection").value;
        const path = document.getElementById("lkmlSaveServerPath").value.trim();
        const fileName = document.getElementById("lkmlSaveServerFileName").value.trim();
        if (!connectionId || !fileName) return;

        lkmlSaveSelection = { source: "server", connectionId, path, fileName };
        showToast("Destino seleccionado. El guardado en el repositorio del servidor estará disponible en una fase posterior.", "success");
        setTimeout(closeDialog, 900);

    });

}

Office.onReady(() => {
    const { active } = parseQueryParams();

    initEvents();
    populateLkmlSaveModelSelect(active);
    setLkmlSaveActiveTab("server");
    populateLkmlSaveConnectionSelect();
    updateLkmlSaveServerList();
    prefillLkmlSaveFileNames();
    updateSaveLkmlConfirmButtonState();
});
