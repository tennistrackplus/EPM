/**
 * SemanticModelStore
 * ------------------------------------------------------------------------
 * Sustituye a las pestañas MODEL_FACT / MODEL_DIMENSION / MODEL_MEASURES /
 * MODEL_RELATIONSHIP / MODEL_ATRIBUTES / MODEL_HIER.
 *
 * Todos los modelos semánticos del libro se guardan en UNA sola clave JSON
 * dentro de Office.context.document.settings (el mismo mecanismo de
 * "roaming settings" que ya usa el add-in para draco_reportProperties,
 * draco_pendingAction, etc.): es un almacén que viaja dentro del propio
 * .xlsx pero que NO es una hoja, ni siquiera oculta, así que no aparece
 * nunca en "Mostrar hoja".
 *
 * Forma de cada modelo guardado:
 *   {
 *     fact: { project, dataset, table },
 *     fields: [ ...mismo objeto que fieldsState en semantic_model.js... ]
 *   }
 *
 * Con un solo objeto por modelo (identificado por su nombre) ya queda listo
 * para poder tener varios modelos semánticos y, más adelante, un selector
 * que decida cuál usar al construir consultas. Por ahora, si no hay ningún
 * modelo marcado como "activo", se usa el primero (orden alfabético) — ver
 * getActiveModelName().
 *
 * Para no reescribir toda la lógica de excelService.js / commands.js (que
 * lee las hojas MODEL_* con coordenadas fila/columna fijas), getModelGrid()
 * genera al vuelo un "grid" {values, startRow:0, startCol:0} con EXACTAMENTE
 * el mismo formato de cabecera/columnas que tenían esas hojas, construido a
 * partir del JSON. El resto del código (cellValue, lastRowInColumnValues,
 * existsAttribute, buildAttributeSQL, buildHierarchySQL...) sigue
 * funcionando sin cambios.
 */

(function () {

    const SMS_KEY = "epm_semanticModels";          // JSON: { [modelName]: modelObj }
    const SMS_ACTIVE_KEY = "epm_activeSemanticModel"; // nombre del modelo activo

    /* ---------------------------------------------------------------
     * Persistencia bruta (Office roaming settings)
     * ------------------------------------------------------------- */

    function _readStoreRaw() {
        try {
            const raw = Office.context.document.settings.get(SMS_KEY);
            if (!raw) return {};
            const parsed = JSON.parse(raw);
            return (parsed && typeof parsed === "object") ? parsed : {};
        } catch (e) {
            console.error("SemanticModelStore: JSON de epm_semanticModels corrupto, se reinicia.", e);
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
        Office.context.document.settings.set(SMS_KEY, JSON.stringify(storeObj));
        await _saveAsync();
    }

    /* ---------------------------------------------------------------
     * API pública: modelo a modelo
     * ------------------------------------------------------------- */

    function getAllModels() {
        return _readStoreRaw();
    }

    function listModelNames() {
        return Object.keys(_readStoreRaw()).sort();
    }

    function getModel(modelName) {
        const store = _readStoreRaw();
        return store[modelName] || null;
    }

    async function saveModel(modelName, modelObj) {
        const store = _readStoreRaw();
        store[modelName] = modelObj;
        await _writeStoreRaw(store);

        // Si todavía no hay modelo activo, este pasa a serlo (para que el
        // editor de informes tenga algo que leer desde el primer momento).
        const active = Office.context.document.settings.get(SMS_ACTIVE_KEY);
        if (!active) {
            await setActiveModelName(modelName);
        }
    }

    async function deleteModel(modelName) {
        const store = _readStoreRaw();
        delete store[modelName];
        await _writeStoreRaw(store);

        if (Office.context.document.settings.get(SMS_ACTIVE_KEY) === modelName) {
            const remaining = Object.keys(store).sort();
            await setActiveModelName(remaining[0] || "");
        }
    }

    /**
     * Modelo "activo": el que usa el editor de informes mientras no exista
     * un selector visible. Si no hay ninguno marcado explícitamente (o el
     * marcado ya no existe), se usa el primero por orden alfabético.
     */
    function getActiveModelName() {
        const names = listModelNames();
        if (names.length === 0) return "";

        const active = Office.context.document.settings.get(SMS_ACTIVE_KEY);
        if (active && names.includes(active)) return active;

        return names[0]; // por ahora: el primero
    }

    async function setActiveModelName(modelName) {
        Office.context.document.settings.set(SMS_ACTIVE_KEY, modelName || "");
        await _saveAsync();
    }

    /* ---------------------------------------------------------------
     * Constructores de filas (mismo layout que las antiguas hojas MODEL_*)
     * ------------------------------------------------------------- */

    function buildFactRow(modelName, fact) {
        fact = fact || {};
        return [modelName, fact.project || "", fact.dataset || "", fact.table || ""];
    }

    function buildRelationshipRows(modelName, fact, fields) {
        fact = fact || {};
        const rows = [];
        (fields || []).forEach((f, idx) => {
            if (!f.enabled || f.type !== "DIMENSION" || !f.relTable) return;
            const keyAttr = (f.attributes || []).find(a => a.isKey);
            rows.push([
                idx + 1, f.name,
                fact.project, fact.dataset, fact.table, f.name,
                f.relProject, f.relDataset, f.relTable,
                keyAttr ? keyAttr.name : f.name,
                "LEFT", modelName
            ]);
        });
        return rows;
    }

    function buildDimensionRows(modelName, fact, fields) {
        fact = fact || {};
        const rows = [];
        (fields || []).forEach((f, idx) => {
            if (!f.enabled || f.type !== "DIMENSION") return;
            const keyAttr = (f.attributes || []).find(a => a.isKey);
            rows.push([
                idx + 1, f.name,
                fact.project, fact.dataset, fact.table, f.name,
                f.relProject || fact.project, f.relDataset || fact.dataset, f.relTable || fact.table,
                keyAttr ? keyAttr.name : f.name,
                modelName
            ]);
        });
        return rows;
    }

    function buildMeasureRows(modelName, fact, fields) {
        fact = fact || {};
        const rows = [];
        (fields || []).forEach((f, idx) => {
            if (!f.enabled || f.type !== "MEASURE") return;
            rows.push([
                idx + 1, f.name,
                fact.project, fact.dataset, fact.table, f.name,
                f.aggregation, f.format, modelName
            ]);
        });
        return rows;
    }

    function buildAttributeRows(modelName, fact, fields) {
        fact = fact || {};
        const rows = [];
        (fields || []).forEach((f, idx) => {
            if (!f.enabled || f.type !== "DIMENSION") return;

            if (f.attributes && f.attributes.length > 0) {
                f.attributes.filter(a => a.enabled).forEach(a => {
                    rows.push([
                        idx + 1, f.name, a.name,
                        f.relProject || fact.project, f.relDataset || fact.dataset, f.relTable || fact.table,
                        a.name, a.alias === a.name ? "" : a.alias, a.dataType,
                        a.isKey ? "X" : "", modelName
                    ]);
                });
            } else {
                rows.push([
                    idx + 1, f.name, f.name,
                    fact.project, fact.dataset, fact.table,
                    f.name, "", f.dataType, "X", modelName
                ]);
            }
        });
        return rows;
    }

    function buildHierarchyRows(modelName, fact, fields) {
        fact = fact || {};
        const rows = [];
        let fila = 1;
        (fields || []).forEach(f => {
            if (!f.enabled || f.type !== "DIMENSION" || !f.hierarchies || f.hierarchies.length === 0) return;
            f.hierarchies.forEach(hier => {
                (hier.levels || []).forEach((lvl, levelIdx) => {
                    const attr = (f.attributes || []).find(a => a.name === lvl.attribute);
                    rows.push([
                        fila++, hier.name, levelIdx + 1, f.name,
                        attr ? (attr.alias || attr.name) : lvl.attribute,
                        f.relProject || fact.project, f.relDataset || fact.dataset, f.relTable || fact.table,
                        lvl.attribute, modelName
                    ]);
                });
            });
        });
        return rows;
    }

    const HEADERS = {
        MODEL_FACT: ["MODEL_NAME", "FACT_PROJECT", "FACT_DATASET", "FACT_TABLE"],
        MODEL_RELATIONSHIP: ["FILA", "DIMENSION", "FACT_PROJECT", "FACT_DATASET", "FACT_TABLE", "FACT_FIELD", "DIM_PROJECT", "DIM_DATASET", "DIM_TABLE", "DIM_FIELD", "JOIN TYPE", "MODEL_NAME"],
        MODEL_DIMENSION: ["FILA", "DIMENSION", "FACT_PROJECT", "FACT_DATASET", "FACT_TABLE", "FACT_FIELD", "DIM_PROJECT", "DIM_DATASET", "DIM_TABLE", "DIM_FIELD", "MODEL_NAME"],
        MODEL_MEASURES: ["FILA", "MEASURE", "FACT_PROJECT", "FACT_DATASET", "FACT_TABLE", "FACT_FIELD", "AGGREGATION", "FORMAT", "MODEL_NAME"],
        MODEL_ATRIBUTES: ["FILA", "DIMENSION", "ATRIBUTE", "DIM_PROJECT", "DIM_DATASET", "DIM_TABLE", "DIM_FIELD", "DISPLAY_NAME", "DATA_TYPE", "IS_KEY", "MODEL_NAME"],
        MODEL_HIER: ["FILA", "HIERARCHY", "NIVEL", "DIMENSION", "ATRIBUTO", "DIM_PROJECT", "DIM_DATASET", "DIM_TABLE", "DIM_FIELD", "MODEL_NAME"]
    };

    const ROW_BUILDERS = {
        MODEL_RELATIONSHIP: buildRelationshipRows,
        MODEL_DIMENSION: buildDimensionRows,
        MODEL_MEASURES: buildMeasureRows,
        MODEL_ATRIBUTES: buildAttributeRows,
        MODEL_HIER: buildHierarchyRows
    };

    /**
     * Genera un grid {values, startRow:0, startCol:0} igual al que devolvía
     * getValuesGrid(context, sheetName) sobre la antigua hoja MODEL_*.
     *
     * sheetName: "MODEL_FACT" | "MODEL_RELATIONSHIP" | "MODEL_DIMENSION" |
     *            "MODEL_MEASURES" | "MODEL_ATRIBUTES" | "MODEL_HIER"
     * modelNameOverride: opcional, para pedir un modelo concreto en vez del
     *                    activo (hueco ya preparado para el futuro selector).
     */
    async function getModelGrid(sheetName, modelNameOverride) {
        const header = HEADERS[sheetName];

        if (sheetName === "MODEL_FACT") {
            const store = _readStoreRaw();
            const rows = Object.keys(store).sort().map(name => buildFactRow(name, store[name].fact));
            return { values: [header, ...rows], startRow: 0, startCol: 0 };
        }

        const modelName = modelNameOverride || getActiveModelName();
        if (!modelName) {
            return { values: [header], startRow: 0, startCol: 0 };
        }

        const model = getModel(modelName);
        if (!model) {
            return { values: [header], startRow: 0, startCol: 0 };
        }

        const builder = ROW_BUILDERS[sheetName];
        const rows = builder ? builder(modelName, model.fact, model.fields) : [];
        return { values: [header, ...rows], startRow: 0, startCol: 0 };
    }

    window.SemanticModelStore = {
        getAllModels,
        listModelNames,
        getModel,
        saveModel,
        deleteModel,
        getActiveModelName,
        setActiveModelName,
        getModelGrid,
        // expuestos por si hace falta reconstruir filas a mano en otro sitio
        buildRelationshipRows,
        buildDimensionRows,
        buildMeasureRows,
        buildAttributeRows,
        buildHierarchyRows
    };

})();
