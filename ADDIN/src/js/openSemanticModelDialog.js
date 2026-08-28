/**
 * Diálogo independiente "Abrir modelo semántico (.lkml)".
 * ------------------------------------------------------------------------
 * Se abre con Office.context.ui.displayDialogAsync directamente desde el
 * botón del ribbon "Abrir modelo semántico" (ver commands.js) o desde el
 * botón "📂 Abrir" del taskpane (ver semantic_model.js): en ambos casos es
 * la MISMA página independiente, no un popup superpuesto dentro del
 * taskpane. No tiene acceso al modelo de objetos de Excel (igual que
 * dataPreview.html / hierarchyPreview.html): solo usa las conexiones
 * guardadas (localStorage, compartido con el resto del add-in) y el
 * repositorio Git elegido (GitRepo, vía fetch).
 *
 * A propósito, igual que el modal original del que procede, esta pantalla
 * NO procesa todavía el fichero elegido: solo deja "seleccionado" un
 * origen (local o de servidor). La importación real del LookML al modelo
 * semántico se conectará en un paso posterior.
 */

let lkmlSelection = null; // { source: "local"|"server", file?: File, connectionId?, path?, name? }

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

function setLkmlActiveTab(tabId) {
    document.getElementById("lkmlTabServer").classList.toggle("active", tabId === "server");
    document.getElementById("lkmlTabLocal").classList.toggle("active", tabId === "local");
    document.getElementById("lkmlPanelServer").classList.toggle("active", tabId === "server");
    document.getElementById("lkmlPanelLocal").classList.toggle("active", tabId === "local");
    updateLkmlConfirmButtonState();
}

// Rellena el selector de conexiones: solo las que tienen un repositorio de
// modelos semánticos (GitHub/GitLab) configurado.
function populateLkmlConnectionSelect() {
    const select = document.getElementById("lkmlServerConnection");
    if (!select) return;

    const connections = (typeof Connections !== "undefined" && typeof Connections.list === "function")
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

// Lista el contenido de la carpeta actual del repositorio de la conexión
// elegida (vía GitRepo: solo metadatos — nombre/ruta/tipo — nunca el
// contenido de ningún fichero). Las carpetas navegan; los .lkml se
// pueden seleccionar.
async function updateLkmlServerList() {
    const list = document.getElementById("lkmlServerList");
    if (!list) return;

    const connectionId = document.getElementById("lkmlServerConnection").value;
    const path = document.getElementById("lkmlServerPath").value.trim();

    if (!connectionId) {
        list.classList.add("is-empty");
        list.innerHTML = "<span class=\"lkml-empty-hint\">Selecciona una conexión para explorar los ficheros .lkml del repositorio.</span>";
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
                    document.getElementById("lkmlServerPath").value = item.path;
                    lkmlSelection = null;
                    updateLkmlConfirmButtonState();
                    updateLkmlServerList();
                });
            } else {
                row.addEventListener("click", () => {
                    list.querySelectorAll(".lkml-server-item.selected").forEach(el => el.classList.remove("selected"));
                    row.classList.add("selected");
                    lkmlSelection = { source: "server", connectionId, path: item.path, name: item.name };
                    updateLkmlConfirmButtonState();
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

function updateLkmlConfirmButtonState() {
    const btn = document.getElementById("btnConfirmOpenLkml");
    if (!btn) return;
    btn.disabled = !lkmlSelection;
}

function handleLkmlLocalFile(file) {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".lkml")) {
        showToast("Selecciona un fichero con extensión .lkml", "error");
        return;
    }
    lkmlSelection = { source: "local", file };

    document.getElementById("lkmlSelectedFileName").textContent = file.name;
    document.getElementById("lkmlSelectedFile").style.display = "flex";
    updateLkmlConfirmButtonState();
}

function initEvents() {

    document.getElementById("btnCancelOpenLkml").addEventListener("click", closeDialog);

    document.getElementById("lkmlTabServer").addEventListener("click", () => setLkmlActiveTab("server"));
    document.getElementById("lkmlTabLocal").addEventListener("click", () => setLkmlActiveTab("local"));

    document.getElementById("lkmlServerConnection").addEventListener("change", () => {
        lkmlSelection = null;
        document.getElementById("lkmlServerPath").value = "";
        updateLkmlServerList();
        updateLkmlConfirmButtonState();
    });

    document.getElementById("lkmlServerUpBtn").addEventListener("click", () => {
        const pathInput = document.getElementById("lkmlServerPath");
        const parts = pathInput.value.split("/").filter(Boolean);
        parts.pop();
        pathInput.value = parts.join("/");
        lkmlSelection = null;
        updateLkmlConfirmButtonState();
        updateLkmlServerList();
    });

    document.getElementById("lkmlServerPath").addEventListener("keydown", (e) => {
        if (e.key !== "Enter") return;
        lkmlSelection = null;
        updateLkmlConfirmButtonState();
        updateLkmlServerList();
    });

    // Selector de fichero local: clic en la dropzone abre el <input type="file">
    const dropzone = document.getElementById("lkmlDropzone");
    const fileInput = document.getElementById("lkmlFileInput");

    dropzone.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", (e) => handleLkmlLocalFile(e.target.files && e.target.files[0]));

    dropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropzone.classList.add("dragover");
    });
    dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
    dropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropzone.classList.remove("dragover");
        const file = e.dataTransfer.files && e.dataTransfer.files[0];
        handleLkmlLocalFile(file);
    });

    // "Seleccionar": por ahora no hace nada con el fichero, solo confirma
    // el origen elegido y cierra el diálogo (la importación real del LookML
    // se conectará en un paso posterior).
    document.getElementById("btnConfirmOpenLkml").addEventListener("click", () => {
        if (!lkmlSelection) return;
        const label = lkmlSelection.source === "local"
            ? lkmlSelection.file.name
            : lkmlSelection.name;
        showToast(`Fichero seleccionado: ${label}`, "success");
        setTimeout(closeDialog, 900);
    });

}

Office.onReady(() => {
    initEvents();
    setLkmlActiveTab("server");
    populateLkmlConnectionSelect();
    updateLkmlServerList();
    updateLkmlConfirmButtonState();
});
