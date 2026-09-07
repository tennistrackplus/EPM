/**
 * PlanningChangesStore
 * ------------------------------------------------------------------------
 * Cambios pendientes de un "informe de planificación" (propiedad del
 * informe "planningReport"): cada vez que el usuario edita a mano una
 * celda del rango de Valores (Draco_<id>_Values) de un informe con esa
 * propiedad activada, se guarda aquí:
 *   - el formato que tenía la celda ANTES de marcarla (para poder
 *     devolvérselo si "Sobrescribir formatos" está desactivado, ver
 *     restoreOriginalFormat en commands.js),
 *   - el valor nuevo que escribió el usuario.
 *
 * Con esto, el botón de ribbon "Planificación > Guardar" puede recorrer
 * solo las celdas realmente cambiadas para construir el INSERT, sin tener
 * que releer toda la tabla pintada.
 *
 * Se persiste en los roaming settings del documento
 * (Office.context.document.settings), el mismo mecanismo que ya usan
 * ReportStore/SemanticModelStore/FilterRangeStore, bajo una sola clave:
 *   { [reportId]: { "<hoja>!<celda>": { sheetName, address, oldFormat, newValue } } }
 *
 * Este fichero se carga tanto en taskpane.html como en commands.html
 * (igual que los otros stores), para que el seguimiento de cambios
 * funcione aunque el taskpane no esté abierto en ese momento.
 */

(function () {

    const PCS_KEY = "draco_planningChanges"; // JSON: { [reportId]: { [sheet!addr]: entry } }

    function _readStoreRaw() {
        try {
            const raw = Office.context.document.settings.get(PCS_KEY);
            if (!raw) return {};
            const parsed = JSON.parse(raw);
            return (parsed && typeof parsed === "object") ? parsed : {};
        } catch (e) {
            console.error("PlanningChangesStore: JSON de draco_planningChanges corrupto, se reinicia.", e);
            return {};
        }
    }

    function _saveAsync() {
        return new Promise((resolve, reject) => {
            Office.context.document.settings.saveAsync((asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    reject(asyncResult.error);
                } else {
                    resolve();
                }
            });
        });
    }

    async function _writeStoreRaw(storeObj) {
        Office.context.document.settings.set(PCS_KEY, JSON.stringify(storeObj));
        await _saveAsync();
    }

    /**
     * Registra (o actualiza) el cambio de una celda. `data` se fusiona con
     * lo que ya hubiera para esa misma celda — así, si el usuario edita la
     * MISMA celda varias veces antes de guardar, el oldFormat guardado
     * sigue siendo el ORIGINAL (el de la primera vez), no el que tuviera
     * ya marcado en cian.
     */
    async function set(reportId, sheetName, address, data) {
        const store = _readStoreRaw();
        const key = `${sheetName}!${address}`;
        if (!store[reportId]) store[reportId] = {};
        store[reportId][key] = Object.assign({ sheetName, address }, store[reportId][key], data);
        await _writeStoreRaw(store);
        return store[reportId][key];
    }

    function getAllForReport(reportId) {
        const store = _readStoreRaw();
        return store[reportId] || {};
    }

    function hasAny(reportId) {
        return Object.keys(getAllForReport(reportId)).length > 0;
    }

    async function clearForReport(reportId) {
        const store = _readStoreRaw();
        if (store[reportId]) {
            delete store[reportId];
            await _writeStoreRaw(store);
        }
    }

    async function removeOne(reportId, sheetName, address) {
        const store = _readStoreRaw();
        const key = `${sheetName}!${address}`;
        if (store[reportId] && store[reportId][key]) {
            delete store[reportId][key];
            if (Object.keys(store[reportId]).length === 0) delete store[reportId];
            await _writeStoreRaw(store);
        }
    }

    window.PlanningChangesStore = { set, getAllForReport, hasAny, clearForReport, removeOne };

})();
