/**
 * FilterRangeStore
 * ------------------------------------------------------------------------
 * Metadatos de los rangos con nombre "Draco_Filter_…" que crea el botón
 * "Añadir filtro" del taskpane sobre la celda seleccionada en Excel.
 *
 * Un rango con nombre de Excel solo puede apuntar a UNA celda; aquí se
 * guarda, para cada nombre de rango, a qué dimensión/campo (y si es una
 * jerarquía) corresponde, a qué informe está atado (o null si es "Todos
 * los informes") y el ÚLTIMO filtro aplicado (para poder reabrir el
 * selector con la selección anterior ya marcada, y para poder repintar la
 * celda sin volver a preguntar).
 *
 * Se persiste en los "roaming settings" del documento
 * (Office.context.document.settings), el mismo mecanismo que ya usan
 * ReportStore y SemanticModelStore, bajo UNA sola clave JSON:
 *   { [rangeName]: { dim, name, isHierarchy, reportId, filter } }
 *
 * reportId es un número (id del informe) o null si el filtro se creó para
 * "Todos los informes" (sufijo "all" en el nombre del rango).
 *
 * Este fichero se carga tanto en taskpane.html como en commands.html
 * (igual que reportStore.js/semanticModelStore.js), para que los
 * metadatos estén disponibles también desde el runtime oculto de
 * comandos (necesario para poder abrir el selector de filtro por clic
 * aunque el taskpane esté cerrado).
 */

(function () {

    const FRS_KEY = "draco_filterRanges"; // JSON: { [rangeName]: meta }

    function _readStoreRaw() {
        try {
            const raw = Office.context.document.settings.get(FRS_KEY);
            if (!raw) return {};
            const parsed = JSON.parse(raw);
            return (parsed && typeof parsed === "object") ? parsed : {};
        } catch (e) {
            console.error("FilterRangeStore: JSON de draco_filterRanges corrupto, se reinicia.", e);
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
        Office.context.document.settings.set(FRS_KEY, JSON.stringify(storeObj));
        await _saveAsync();
    }

    /**
     * Crea o actualiza los metadatos de un rango de filtro. `meta` se
     * fusiona con lo que ya hubiera (así se puede llamar solo con
     * { filter: ... } para actualizar el último filtro aplicado sin
     * repetir dim/name/reportId).
     */
    async function set(rangeName, meta) {
        const store = _readStoreRaw();
        store[rangeName] = Object.assign({}, store[rangeName], meta);
        await _writeStoreRaw(store);
        return store[rangeName];
    }

    function get(rangeName) {
        const store = _readStoreRaw();
        return store[rangeName] || null;
    }

    async function remove(rangeName) {
        const store = _readStoreRaw();
        if (store[rangeName]) {
            delete store[rangeName];
            await _writeStoreRaw(store);
        }
    }

    function listNames() {
        return Object.keys(_readStoreRaw());
    }

    function getAll() {
        return _readStoreRaw();
    }

    window.FilterRangeStore = { set, get, remove, listNames, getAll };

})();
