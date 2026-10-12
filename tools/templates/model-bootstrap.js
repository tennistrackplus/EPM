/**
 * model-bootstrap.js — sustituye a la vista de conexión y al selector de modelos del addin.
 * (Generado por tools/sync-addin.py desde tools/templates/.)
 *
 * El selector "Modelo semántico" del diseñador va asociado 1:1 a los cubos del proyecto, y esa lista y la
 * descarga de cada modelo (.lkml) las hace el ANFITRIÓN (frameElement.__dracoHost):
 *     host.cubes                -> [{ name }]
 *     host.resolveModel(name)   -> { fact, fields } | null     (formato de SemanticModelStore.saveModel)
 *
 *   1) Se registra cada cubo como modelo "vacío" (sin red) para que el selector los liste; si ya había un
 *      modelo con datos (se reabre una tabla guardada), se respeta.
 *   2) El modelo del informe activo se resuelve ya; el de cualquier otro cubo, cuando el usuario lo elige
 *      (se intercepta TaskPaneApp.onModelSelectorChange sin tocar taskpane.js).
 *   3) Se avisa a Office.onReady (host-bridge) y a quien espere el fin del arranque (__wteInitDonePromise).
 */
(function () {
    // TaskPaneApp es un `const` de taskpane.js (no cuelga de window): se expone para que el diseñador pueda llamarlo.
    if (typeof TaskPaneApp !== "undefined") window.TaskPaneApp = TaskPaneApp;
    const pending = {};

    function host() { return window.frameElement && window.frameElement.__dracoHost; }

    async function resolveModel(name) {
        const entry = pending[name];
        if (!entry || entry.resolved) return;
        entry.resolved = true;
        try {
            const model = await host().resolveModel(name);
            if (model && model.fields && model.fields.length) await window.SemanticModelStore.saveModel(name, model);
            else { entry.resolved = false; console.warn(`model-bootstrap: el cubo "${name}" no tiene modelo semántico.`); }   // se reintenta al volver a elegirlo
        } catch (err) {
            entry.resolved = false;
            console.error(`model-bootstrap: error al leer el modelo de "${name}":`, err);
        }
    }

    function patchModelSelectorChange() {
        if (typeof TaskPaneApp === "undefined" || TaskPaneApp.__dracoPatched) return false;
        const original = TaskPaneApp.onModelSelectorChange.bind(TaskPaneApp);
        TaskPaneApp.onModelSelectorChange = async function (modelName) {
            if (modelName) await resolveModel(modelName);
            return original(modelName);
        };
        TaskPaneApp.__dracoPatched = true;
        return true;
    }

    function patchInitDone() {
        if (typeof TaskPaneApp === "undefined" || TaskPaneApp.__dracoInitPatched) return false;
        const originalInit = TaskPaneApp.init.bind(TaskPaneApp);
        let done;
        window.__wteInitDonePromise = new Promise((resolve) => { done = resolve; });
        TaskPaneApp.init = async function (...args) { const r = await originalInit(...args); done(); return r; };
        TaskPaneApp.__dracoInitPatched = true;
        return true;
    }

    async function bootstrapInner() {
        const h = host();
        if (!h || !Array.isArray(h.cubes)) { console.error("model-bootstrap: el anfitrión no ofrece cubos."); return; }
        const existing = (typeof window.SemanticModelStore.getAllModels === "function")
            ? await window.SemanticModelStore.getAllModels() : null;
        for (const cube of h.cubes) {
            pending[cube.name] = { resolved: false };
            const already = existing && existing[cube.name] && existing[cube.name].fields && existing[cube.name].fields.length;
            if (already) pending[cube.name].resolved = true;               // ya viene con datos (tabla guardada)
            else await window.SemanticModelStore.saveModel(cube.name, { fact: {}, fields: [] });
        }
        patchModelSelectorChange();
        patchInitDone();
        const active = window.ReportStore && window.ReportStore.getActiveReport ? window.ReportStore.getActiveReport() : null;
        const activeName = active ? active.semanticModelName : "";
        if (activeName && pending[activeName]) {
            await resolveModel(activeName);
            await window.SemanticModelStore.setActiveModelName(activeName);
        }
        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.populateModelSelectorMain) await TaskPaneApp.populateModelSelectorMain();
    }

    async function bootstrap() {
        try { await bootstrapInner(); }
        catch (err) { console.error("model-bootstrap: error inesperado:", err); }
        finally { patchModelSelectorChange(); patchInitDone(); if (window.__wteSignalModelsReady) window.__wteSignalModelsReady(); }
    }

    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", bootstrap); else bootstrap();
})();
