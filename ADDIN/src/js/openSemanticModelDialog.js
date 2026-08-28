/**
 * Diálogo independiente "Abrir modelo semántico (.lkml)".
 * ------------------------------------------------------------------------
 * Se abre con Office.context.ui.displayDialogAsync directamente desde el
 * botón del ribbon "Abrir modelo semántico" (ver commands.js) o desde el
 * botón "📂 Abrir" del taskpane (ver semantic_model.js), en ambos casos a
 * través de js/lkmlOpenBridge.js: es la MISMA página independiente, no un
 * popup superpuesto dentro del taskpane. No tiene acceso al modelo de
 * objetos de Excel ni a Office.context.document.settings (igual que
 * dataPreview.html / hierarchyPreview.html): solo usa las conexiones
 * guardadas (localStorage, compartido con el resto del add-in) y el
 * repositorio Git elegido (GitRepo, vía fetch).
 *
 * Al pulsar "Seleccionar": pide el nombre del modelo semántico (obligatorio),
 * lee el fichero .lkml elegido (local o del repositorio, vía
 * GitRepo.getFileContent), lo convierte en un modelo semántico con
 * LkmlImport.parseContent (ver js/lkmlImport.js) y lo envía a quien abrió
 * el diálogo con Office.context.ui.messageParent(), porque este diálogo NO
 * tiene acceso a SemanticModelStore. Es js/lkmlOpenBridge.js, en el lado
 * que sí tiene ese acceso, quien lo guarda de verdad y cierra el diálogo.
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
                    prefillModelName(item.name);
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

    const nameInput = document.getElementById("lkmlOpenModelName");
    const modelName = nameInput ? nameInput.value.trim() : "";

    btn.disabled = !lkmlSelection || modelName === "";
}

// Sugiere el nombre del modelo a partir del nombre del fichero (<nombre>.lkml)
// si el usuario todavía no ha escrito uno propio.
function prefillModelName(fileName) {
    const nameInput = document.getElementById("lkmlOpenModelName");
    if (!nameInput || nameInput.value.trim() !== "") return;

    const suggested = (fileName || "").replace(/\.lkml$/i, "");
    if (suggested) nameInput.value = suggested;
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
    prefillModelName(file.name);
    updateLkmlConfirmButtonState();
}

function readFileAsText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error || new Error("No se ha podido leer el fichero."));
        reader.readAsText(file);
    });
}

function setImporting(isImporting, label) {
    const btn = document.getElementById("btnConfirmOpenLkml");
    const cancelBtn = document.getElementById("btnCancelOpenLkml");
    if (!btn) return;

    btn.textContent = isImporting ? (label || "Importando…") : "Seleccionar";
    if (cancelBtn) cancelBtn.disabled = isImporting;
    if (isImporting) {
        btn.disabled = true;
    } else {
        updateLkmlConfirmButtonState();
    }
}

function initEvents() {

    document.getElementById("btnCancelOpenLkml").addEventListener("click", closeDialog);

    document.getElementById("lkmlOpenModelName").addEventListener("input", updateLkmlConfirmButtonState);

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

    // "Seleccionar": lee el fichero elegido (local o del repositorio),
    // genera el modelo semántico a partir de su LookML (LkmlImport) y lo
    // envía a quien abrió el diálogo (LkmlOpenBridge, ver
    // js/lkmlOpenBridge.js), que es quien tiene acceso a SemanticModelStore
    // para guardarlo de verdad con el nombre indicado.
    document.getElementById("btnConfirmOpenLkml").addEventListener("click", async () => {

        if (!lkmlSelection) return;

        const modelName = document.getElementById("lkmlOpenModelName").value.trim();

        if (!modelName) {
            showToast("Indica el nombre con el que se guardará el modelo semántico.", "error");
            return;
        }

        setImporting(true, "Leyendo fichero…");

        try {

            let content;

            if (lkmlSelection.source === "local") {
                content = await readFileAsText(lkmlSelection.file);
            } else {
                const conn = Connections.getById(lkmlSelection.connectionId);
                const repoConfig = conn && conn.config && conn.config.semanticRepo;
                if (!repoConfig) throw new Error("La conexión elegida no tiene un repositorio configurado.");
                content = await GitRepo.getFileContent(repoConfig, lkmlSelection.path);
            }

            setImporting(true, "Generando modelo…");

            const model = LkmlImport.parseContent(content);

            setImporting(true, "Guardando…");

            Office.context.ui.messageParent(JSON.stringify({ modelName, model }));

        } catch (err) {
            console.error("Error al leer o generar el modelo semántico:", err);
            setImporting(false);
            showToast("Error al importar el fichero: " + err.message, "error");
        }

    });

    // Si quien nos abrió (LkmlOpenBridge) no ha podido guardar el modelo en
    // SemanticModelStore, nos lo dice aquí para que el usuario lo sepa y
    // pueda reintentar en vez de quedarse sin saber qué ha pasado.
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {

        let payload = null;
        try {
            payload = JSON.parse(arg.message);
        } catch (e) {
            payload = null;
        }

        setImporting(false);
        showToast(
            (payload && payload.error)
                ? `No se ha podido guardar el modelo semántico: ${payload.error}`
                : "No se ha podido guardar el modelo semántico.",
            "error"
        );

    });

}

Office.onReady(() => {
    initEvents();
    setLkmlActiveTab("server");
    populateLkmlConnectionSelect();
    updateLkmlServerList();
    updateLkmlConfirmButtonState();
});
