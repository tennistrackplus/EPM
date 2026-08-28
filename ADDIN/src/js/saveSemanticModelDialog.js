/**
 * Diálogo independiente "Guardar modelo semántico (.lkml)".
 * ------------------------------------------------------------------------
 * Se abre con Office.context.ui.displayDialogAsync directamente desde el
 * botón del ribbon "Guardar modelo semántico" (ver js/commands.js) o desde
 * el botón "📤 LookML" del taskpane (ver semantic_model.js), en ambos casos
 * a través de js/lkmlSaveBridge.js: es la MISMA página independiente, no un
 * popup superpuesto dentro del taskpane.
 *
 * Al pulsar "Guardar" NO se descarga ningún fichero ni se escribe en
 * GitHub/GitLab: se genera el LookML del modelo elegido (tabla de hechos,
 * dimensiones, atributos, jerarquías, measures y relaciones) y se envía a
 * quien abrió el diálogo con Office.context.ui.messageParent(), porque este
 * diálogo NO tiene acceso al modelo de objetos de Excel. Es
 * js/lkmlSaveBridge.js, en el lado que sí tiene ese acceso, quien escribe
 * el resultado en la celda EDIT_REPORT!G1 y cierra el diálogo.
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

function updateConfirmButtonState() {
    const btn = document.getElementById("btnConfirmSaveLkml");
    if (!btn) return;
    btn.disabled = document.getElementById("lkmlSaveModelSelect").value === "";
}

function setSaving(isSaving) {
    const btn = document.getElementById("btnConfirmSaveLkml");
    const cancelBtn = document.getElementById("btnCancelSaveLkml");
    if (!btn) return;

    btn.textContent = isSaving ? "Guardando…" : "Guardar";
    btn.disabled = isSaving || document.getElementById("lkmlSaveModelSelect").value === "";
    if (cancelBtn) cancelBtn.disabled = isSaving;
}

function initEvents() {

    document.getElementById("btnCancelSaveLkml").addEventListener("click", closeDialog);

    document.getElementById("lkmlSaveModelSelect").addEventListener("change", updateConfirmButtonState);

    // "Guardar": genera el LookML del modelo elegido y se lo envía a quien
    // abrió el diálogo (messageParent), que es quien tiene acceso a Excel
    // para escribirlo en EDIT_REPORT!G1. El diálogo se queda a la espera:
    // lo cierra el padre cuando termina, o nos avisa del error si falla.
    document.getElementById("btnConfirmSaveLkml").addEventListener("click", () => {

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

        setSaving(true);
        Office.context.ui.messageParent(JSON.stringify({ modelName, content }));

    });

    // Si quien nos abrió no ha podido escribir en EDIT_REPORT!G1 (p.ej. la
    // hoja no existe o el libro está protegido), nos lo dice aquí para que
    // el usuario lo sepa y pueda reintentar en vez de quedarse sin saber
    // qué ha pasado.
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {

        let payload = null;
        try {
            payload = JSON.parse(arg.message);
        } catch (e) {
            payload = null;
        }

        setSaving(false);
        showToast(
            (payload && payload.error) ? `No se ha podido guardar: ${payload.error}` : "No se ha podido guardar.",
            "error"
        );

    });

}

Office.onReady(() => {
    const { active } = parseQueryParams();

    initEvents();
    populateModelSelect(active);
    updateConfirmButtonState();
});
