/* ==========================================================================
 * ExcelService — traducción de las macros de lectura/escritura del
 * diseñador de informes (frmReportDesigner2 + módulo de filtros).
 * ========================================================================== */

const SVC_CRLF = "\r\n";

/* ---------------------------------------------------------------------
 * Utilidades de rango ("grids"), igual que en commands.js: se carga el
 * used range de una hoja una sola vez y se accede por coordenada absoluta
 * (fila,columna) como hacía ws.Cells(R,C) en VBA.
 * ------------------------------------------------------------------- */

async function getValuesGrid(context, sheetName) {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = sheet.getUsedRangeOrNullObject();
    used.load(["values", "rowIndex", "columnIndex", "isNullObject"]);
    await context.sync();

    if (used.isNullObject) {
        return { values: [], startRow: 0, startCol: 0 };
    }

    return { values: used.values, startRow: used.rowIndex, startCol: used.columnIndex };
}

function cellValue(grid, row1, col1) {
    const r = row1 - 1 - grid.startRow;
    const c = col1 - 1 - grid.startCol;
    if (r < 0 || c < 0 || r >= grid.values.length) return "";
    const rowArr = grid.values[r];
    if (!rowArr || c >= rowArr.length) return "";
    const v = rowArr[c];
    return v === null || v === undefined ? "" : v;
}

function lastRowInColumnValues(grid, col1) {
    const c = col1 - 1 - grid.startCol;
    for (let r = grid.values.length - 1; r >= 0; r--) {
        const rowArr = grid.values[r];
        const v = rowArr && c >= 0 && c < rowArr.length ? rowArr[c] : "";
        if (v !== null && v !== undefined && String(v) !== "") {
            return r + 1 + grid.startRow;
        }
    }
    return grid.startRow;
}

// Convierte {row,col} (base 1) en dirección tipo "B16"
function numToColLetters(n) {
    let s = "";
    while (n > 0) {
        const rem = (n - 1) % 26;
        s = String.fromCharCode(65 + rem) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

function addressFromRC(row, col) {
    return numToColLetters(col) + row;
}

// Convierte "B16" (o "$B$16") en {row, col} (base 1)
function parseAddress(addr) {
    addr = String(addr).trim();
    if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
    if (addr.indexOf(":") !== -1) addr = addr.split(":")[0];
    addr = addr.replace(/\$/g, "");
    const m = addr.match(/^([A-Za-z]+)(\d+)$/);
    if (!m) return null;
    const colLetters = m[1].toUpperCase();
    const row = parseInt(m[2], 10);
    let col = 0;
    for (let i = 0; i < colLetters.length; i++) {
        col = col * 26 + (colLetters.charCodeAt(i) - 64);
    }
    return { row, col };
}

/* ---------------------------------------------------------------------
 * BuildFilterValuesSQL / ExistsAttribute / BuildAttributeSQL /
 * ExistsHierarchy / BuildHierarchySQL
 * ------------------------------------------------------------------- */

function existsAttribute(atribGrid, dimension, attributeName) {
    const lastRow = lastRowInColumnValues(atribGrid, 1);
    for (let R = 2; R <= lastRow; R++) {
        if (String(cellValue(atribGrid, R, 2)).toUpperCase() === String(dimension).toUpperCase()
            && String(cellValue(atribGrid, R, 3)).toUpperCase() === String(attributeName).toUpperCase()) {
            return true;
        }
    }
    return false;
}

function buildAttributeSQL(atribGrid, dimension, attributeName) {
    const lastRow = lastRowInColumnValues(atribGrid, 1);
    let project = "", dataset = "", table = "", field = "";

    for (let R = 2; R <= lastRow; R++) {
        if (String(cellValue(atribGrid, R, 2)).toUpperCase() === String(dimension).toUpperCase()
            && String(cellValue(atribGrid, R, 3)).toUpperCase() === String(attributeName).toUpperCase()) {
            project = cellValue(atribGrid, R, 4);
            dataset = cellValue(atribGrid, R, 5);
            table = cellValue(atribGrid, R, 6);
            field = cellValue(atribGrid, R, 7);
            break;
        }
    }

    return "SELECT DISTINCT" + SVC_CRLF
        + "    " + field + SVC_CRLF
        + "FROM " + Provider.qualify(project, dataset, table) + SVC_CRLF
        + "ORDER BY " + field;
}

function existsHierarchy(hierGrid, dimension, hierarchy) {
    const lastRow = lastRowInColumnValues(hierGrid, 1);
    for (let R = 2; R <= lastRow; R++) {
        if (String(cellValue(hierGrid, R, 4)).toUpperCase() === String(dimension).toUpperCase()
            && String(cellValue(hierGrid, R, 2)).toUpperCase() === String(hierarchy).toUpperCase()) {
            return true;
        }
    }
    return false;
}

function buildHierarchySQL(hierGrid, dimension, hierarchy) {
    const lastRow = lastRowInColumnValues(hierGrid, 1);
    const fields = [];
    let project = "", dataset = "", table = "";

    for (let R = 2; R <= lastRow; R++) {
        if (String(cellValue(hierGrid, R, 4)).toUpperCase() === String(dimension).toUpperCase()
            && String(cellValue(hierGrid, R, 2)).toUpperCase() === String(hierarchy).toUpperCase()) {
            fields.push(cellValue(hierGrid, R, 9));  // I: DIM_FIELD
            project = cellValue(hierGrid, R, 6);      // F: DIM_PROJECT
            dataset = cellValue(hierGrid, R, 7);      // G: DIM_DATASET
            table = cellValue(hierGrid, R, 8);        // H: DIM_TABLE
        }
    }

    let sql = "SELECT DISTINCT" + SVC_CRLF;
    sql += fields.map(f => "    " + f).join("," + SVC_CRLF);
    sql += SVC_CRLF + "FROM " + Provider.qualify(project, dataset, table) + SVC_CRLF;
    sql += "ORDER BY" + SVC_CRLF;
    sql += fields.map(f => "    " + f).join("," + SVC_CRLF);

    return sql;
}

/* ---------------------------------------------------------------------
 * ExecuteSQL (idéntico al usado en commands.js)
 * ------------------------------------------------------------------- */

/**
 * Convierte filas normalizadas (array de objetos { COLUMNA: valor }) al
 * mismo formato de texto que devuelve BigQuery y que sabe leer
 * parseMemberJsonTree()/filterModal.js: buscan el literal `"v":` una vez
 * por celda, en el mismo orden columna a columna del SELECT. Copia local
 * (misma lógica que en commands.js) para no depender del orden de carga
 * entre los dos ficheros.
 */
function svcRowsToPseudoBqJson(rows) {
    // A diferencia de snowflakeRowsToPseudoBqJson() (usada para el informe
    // principal, que solo escanea "v": posicionalmente), loadJsonTree()
    // del diálogo de filtro SÍ necesita el nombre de cada columna
    // (regex /"name":\s*"..."/) para etiquetar cada valor — por eso aquí
    // hace falta además el bloque "schema" con los nombres reales.
    let schema = '{"fields":[]}';
    if (rows.length > 0) {
        const fields = Object.keys(rows[0]).map(name => `{"name": "${name.replace(/"/g, '\\"')}"}`);
        schema = `{"fields":[${fields.join(",")}]}`;
    }

    let out = `{"schema":${schema},"rows":[`;
    rows.forEach((row, i) => {
        if (i > 0) out += ",";
        out += '{"f":[';
        const values = Object.values(row);
        values.forEach((val, j) => {
            if (j > 0) out += ",";
            if (val === null || val === undefined) {
                out += '{"v": null}';
            } else {
                const text = String(val).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
                out += '{"v": "' + text + '"}';
            }
        });
        out += "]}";
    });
    out += "]}";
    return out;
}

async function executeSQLBigQuery(sql) {
    // ÚNICO cambio respecto al add-in original en todo este fichero: en vez
    // de autenticar y llamar directamente a BigQuery/Snowflake por su
    // cuenta (con sus propios tokens en localStorage), se reutiliza la
    // conexión YA autenticada de la aplicación anfitriona (PLANNING),
    // accesible como window.parent.Provider. El resultado se envuelve en
    // el mismo pseudo-JSON de BigQuery que ya sabían leer parseMemberJsonTree
    // y el resto de commands.js/taskpane.js, así que nada más cambia.
    const rows = await window.parent.Provider.runQuery(sql);
    return svcRowsToPseudoBqJson(rows);
}

/* ---------------------------------------------------------------------
 * ExcelService
 * ------------------------------------------------------------------- */

const ExcelService = {

    /**
     * LoadDimensions + LoadAttributes: dimensiones (antes MODEL_ATRIBUTES),
     * atributos + nombres de jerarquía (antes MODEL_HIER, solo NIVEL=1),
     * y la pseudo-dimensión "MEASURE" con las medidas (antes MODEL_MEASURES).
     * Los datos ya no viven en hojas: se generan al vuelo desde el modelo
     * semántico activo guardado por SemanticModelStore (por ahora, el
     * primero que exista; ver SemanticModelStore.getActiveModelName).
     */
    async readDim2Data() {
        try {
            {
                const atribGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
                const hierGrid = await window.SemanticModelStore.getModelGrid("MODEL_HIER");
                const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");

                const order = [];
                const dimensionsMap = {};

                const lastAttrRow = lastRowInColumnValues(atribGrid, 2);
                for (let R = 2; R <= lastAttrRow; R++) {
                    const dimName = String(cellValue(atribGrid, R, 2)).trim();
                    const attrName = String(cellValue(atribGrid, R, 3)).trim();
                    if (!dimName || !attrName) continue;

                    if (!dimensionsMap[dimName]) {
                        dimensionsMap[dimName] = { dimension: dimName, hierarchies: [], attributes: [] };
                        order.push(dimName);
                    }
                    if (!dimensionsMap[dimName].attributes.includes(attrName)) {
                        dimensionsMap[dimName].attributes.push(attrName);
                    }
                }

                // Nombres de jerarquía (MODEL_HIER, columna B), solo el nivel 1 marca el nombre único
                const lastHierRow = lastRowInColumnValues(hierGrid, 4);
                for (let R = 2; R <= lastHierRow; R++) {
                    const dimName = String(cellValue(hierGrid, R, 4)).trim(); // D: DIMENSION
                    const hierName = String(cellValue(hierGrid, R, 2)).trim(); // B: HIERARCHY
                    const nivel = Number(cellValue(hierGrid, R, 3));           // C: NIVEL
                    if (!dimName || !hierName || nivel !== 1) continue;

                    if (!dimensionsMap[dimName]) {
                        dimensionsMap[dimName] = { dimension: dimName, hierarchies: [], attributes: [] };
                        order.push(dimName);
                    }
                    if (!dimensionsMap[dimName].hierarchies.includes(hierName)) {
                        dimensionsMap[dimName].hierarchies.push(hierName);
                    }
                }

                // Pseudo-dimensión MEASURE (MODEL_MEASURES, columna B)
                const measureNames = [];
                const lastMeasRow = lastRowInColumnValues(measuresGrid, 2);
                for (let R = 2; R <= lastMeasRow; R++) {
                    const name = String(cellValue(measuresGrid, R, 2)).trim();
                    if (name) measureNames.push(name);
                }

                const dimensionsList = order.map(d => dimensionsMap[d]);
                dimensionsList.push({ dimension: "MEASURE", hierarchies: [], attributes: measureNames });

                return { data: dimensionsList };
            }
        } catch (error) {
            console.error("Error leyendo el modelo semántico (atributos/jerarquías/medidas):", error);
            return { error: "Error al leer los datos del modelo." };
        }
    },

    /**
     * BuildFilterValuesSQL: decide si "name" es un atributo suelto o una
     * jerarquía y genera el SQL correspondiente. Devuelve "" si no existe.
     */
    async buildFilterValuesSQL(dimension, name) {
        const atribGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
        const hierGrid = await window.SemanticModelStore.getModelGrid("MODEL_HIER");

        if (existsAttribute(atribGrid, dimension, name)) {
            return buildAttributeSQL(atribGrid, dimension, name);
        }
        if (existsHierarchy(hierGrid, dimension, name)) {
            return buildHierarchySQL(hierGrid, dimension, name);
        }
        return "";
    },

    /**
     * Devuelve los niveles (nivel + nombre de atributo) de una jerarquía
     * concreta, leídos de MODEL_HIER, para poder listarlos en el panel
     * "Opciones de campo" (expandir hasta nivel / mostrar niveles).
     */
    async getHierarchyLevels(dimension, hierarchy) {
        const hierGrid = await window.SemanticModelStore.getModelGrid("MODEL_HIER");
        const lastRow = lastRowInColumnValues(hierGrid, 4);
        const levels = [];

        for (let R = 2; R <= lastRow; R++) {
            if (String(cellValue(hierGrid, R, 4)).toUpperCase() === String(dimension).toUpperCase()
                && String(cellValue(hierGrid, R, 2)).toUpperCase() === String(hierarchy).toUpperCase()) {
                levels.push({
                    nivel: Number(cellValue(hierGrid, R, 3)),
                    attribute: String(cellValue(hierGrid, R, 5)).trim()
                });
            }
        }

        levels.sort((a, b) => a.nivel - b.nivel);
        return levels;
    },

    async executeSQL(sql) {
        return await executeSQLBigQuery(sql);
    },

    /**
     * LoadFilters + LoadRows + LoadCols + LoadFixed + LoadRanges:
     * lee el estado actual del diseño del informe indicado (o el activo si
     * no se pasa reportId) desde ReportStore (JSON). El único dato que
     * sigue viniendo de la hoja física es el ancla E1 (celda seleccionada
     * al pulsar "Editar informe" la primera vez).
     */
    async loadEditReportDesign(reportId) {
        const id = reportId || (window.ReportStore ? window.ReportStore.getActiveReportId() : null);
        const report = id && window.ReportStore ? window.ReportStore.getReport(id) : null;
        const design = report ? (report.design || {}) : {};

        // ---- Celda de anclaje para la PRIMERA vez que se monta el informe
        // (EDIT_REPORT!E1, sigue siendo física: la escribe captureActiveEditContext
        // en taskpane.js). Si está vacía o no es una dirección válida, se usa
        // A1. Solo se aplica cuando todavía no hay rrAddress/rcAddress
        // guardados (ver RangeAxis.loadFromAddresses en taskpane.js).
        let anchorAddress = "";
        try {
            anchorAddress = await Excel.run(async (context) => {
                const cell = context.workbook.worksheets.getItem("EDIT_REPORT").getRange("E1");
                cell.load("values");
                await context.sync();
                return String((cell.values && cell.values[0] && cell.values[0][0]) || "").trim();
            });
        } catch (e) {
            // La hoja EDIT_REPORT todavía no existe: sin ancla, se usará A1.
        }

        return {
            filters: design.filters || [],
            rows: design.rows || [],
            columns: design.columns || [],
            rowsStatic: !!design.rowsStatic,
            colsStatic: !!design.colsStatic,
            rrAddress: design.rrAddress || "",
            rcAddress: design.rcAddress || "",
            anchorAddress,
            fieldOptions: design.fieldOptions || {}
        };
    },

    /**
     * SaveFilters + SaveRows + SaveCols + SaveFixed + SaveRanges:
     * guarda el diseño del informe indicado (o el activo) en ReportStore
     * (JSON), expandiendo jerarquías a través de MODEL_HIER igual que hacía
     * antes frmReportDesigner2/esta misma función al escribir en la hoja.
     * YA NO escribe nada en EDIT_REPORT: esa hoja se reserva para el
     * último SQL/JSON (X1/Y1), la hoja/celda activas (D1/E1) y el picker
     * de doble clic del XLAM.
     */
    async saveEditReportDesign(reportId, state) {
        const id = reportId || (window.ReportStore ? window.ReportStore.getActiveReportId() : null);
        if (!id || !window.ReportStore) return;
        await window.ReportStore.saveDesign(id, state);
    },

    /**
     * Guarda las "Propiedades del informe" (modal del taskpane) en
     * ReportStore, para el informe indicado (o el activo). Sustituye a la
     * antigua escritura en EDIT_REPORT!D2:D6 + clave global
     * "draco_reportProperties": ahora cada informe tiene las suyas.
     */
    async saveReportPropertiesToSheetCells(props, reportId) {
        const id = reportId || (window.ReportStore ? window.ReportStore.getActiveReportId() : null);
        if (!id || !window.ReportStore) return;
        await window.ReportStore.saveReportProperties(id, props);
    },

    /**
     * Lista de modelos semánticos disponibles (nombres únicos, ordenados).
     * Usado para rellenar el listbox "Modelo semántico" de Propiedades del
     * informe y, en el futuro, para el selector de modelo semántico del
     * editor de informes.
     */
    async getSemanticModels() {
        return window.SemanticModelStore.listModelNames();
    }
};

window.ExcelService = ExcelService;
window.ReportDesignerUtils = { addressFromRC, parseAddress, numToColLetters };
