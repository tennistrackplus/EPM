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
    let out = '{"rows":[';
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
    if (Provider.key() === "snowflake") {
        const rows = await SF.runQuery(sql);
        return svcRowsToPseudoBqJson(rows);
    }

    // BigQuery (comportamiento original, sin cambios)
    const token = localStorage.getItem("bigquery_access_token");
    const expires = localStorage.getItem("bigquery_token_expires");

    if (!token || !expires || Date.now() >= parseInt(expires)) {
        throw new Error("No hay una sesión activa de BigQuery. Inicia sesión en el panel primero.");
    }

    const projectId = "bigqueryexcelconnector";
    const url = "https://bigquery.googleapis.com/bigquery/v2/projects/" + projectId + "/queries";

    const body = JSON.stringify({ query: sql, useLegacySql: false });

    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json"
        },
        body: body
    });

    return await response.text();
}

/* ---------------------------------------------------------------------
 * ExcelService
 * ------------------------------------------------------------------- */

const ExcelService = {

    /**
     * LoadDimensions + LoadAttributes: dimensiones (MODEL_ATRIBUTES),
     * atributos + nombres de jerarquía (MODEL_HIER, solo NIVEL=1),
     * y la pseudo-dimensión "MEASURE" con las medidas de MODEL_MEASURES.
     */
    async readDim2Data() {
        try {
            return await Excel.run(async (context) => {
                const atribGrid = await getValuesGrid(context, "MODEL_ATRIBUTES");
                const hierGrid = await getValuesGrid(context, "MODEL_HIER");
                const measuresGrid = await getValuesGrid(context, "MODEL_MEASURES");

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
            });
        } catch (error) {
            console.error("Error leyendo MODEL_ATRIBUTES/MODEL_HIER/MODEL_MEASURES:", error);
            return { error: "Error al leer los datos del modelo." };
        }
    },

    /**
     * BuildFilterValuesSQL: decide si "name" es un atributo suelto o una
     * jerarquía y genera el SQL correspondiente. Devuelve "" si no existe.
     */
    async buildFilterValuesSQL(dimension, name) {
        return await Excel.run(async (context) => {
            const atribGrid = await getValuesGrid(context, "MODEL_ATRIBUTES");
            const hierGrid = await getValuesGrid(context, "MODEL_HIER");

            if (existsAttribute(atribGrid, dimension, name)) {
                return buildAttributeSQL(atribGrid, dimension, name);
            }
            if (existsHierarchy(hierGrid, dimension, name)) {
                return buildHierarchySQL(hierGrid, dimension, name);
            }
            return "";
        });
    },

    /**
     * Devuelve los niveles (nivel + nombre de atributo) de una jerarquía
     * concreta, leídos de MODEL_HIER, para poder listarlos en el panel
     * "Opciones de campo" (expandir hasta nivel / mostrar niveles).
     */
    async getHierarchyLevels(dimension, hierarchy) {
        return await Excel.run(async (context) => {
            const hierGrid = await getValuesGrid(context, "MODEL_HIER");
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
        });
    },

    async executeSQL(sql) {
        return await executeSQLBigQuery(sql);
    },

    /**
     * LoadFilters + LoadRows + LoadCols + LoadFixed + LoadRanges:
     * lee el estado actual del diseño desde EDIT_REPORT.
     */
    async loadEditReportDesign() {
        return await Excel.run(async (context) => {
            const grid = await getValuesGrid(context, "EDIT_REPORT");

            // ---- Filtros: C=Dimension, D=Atributo real, E=Filtro(valor), F=Jerarquia ----
            const filters = [];
            {
                const lastRow = lastRowInColumnValues(grid, 3);
                for (let R = 15; R <= lastRow; R++) {
                    const dimension = String(cellValue(grid, R, 3)).trim();
                    if (dimension === "") continue;

                    const realAttribute = String(cellValue(grid, R, 4)).trim();
                    const value = String(cellValue(grid, R, 5)).trim();
                    const hierarchy = String(cellValue(grid, R, 6)).trim();

                    filters.push({
                        dimension: dimension,
                        name: hierarchy !== "" ? hierarchy : realAttribute,
                        isHierarchy: hierarchy !== "",
                        realAttribute: realAttribute,
                        value: value
                    });
                }
            }

            // ---- Filas: H=Dimension, I=Atributo, J=Nivel, K=Jerarquia (solo Nivel=1) ----
            const rows = [];
            {
                const lastRow = lastRowInColumnValues(grid, 8);
                for (let R = 15; R <= lastRow; R++) {
                    const dimension = String(cellValue(grid, R, 8)).trim();
                    if (dimension === "") continue;
                    if (Number(cellValue(grid, R, 10)) !== 1) continue; // solo el primer nivel

                    const attribute = String(cellValue(grid, R, 9)).trim();
                    const hierarchy = String(cellValue(grid, R, 11)).trim();

                    rows.push({
                        dimension: dimension,
                        name: hierarchy !== "" ? hierarchy : attribute,
                        isHierarchy: hierarchy !== ""
                    });
                }
            }

            // ---- Columnas: N=Dimension, O=Atributo, P=Nivel, Q=Jerarquia (solo Nivel=1) ----
            const columns = [];
            {
                const lastRow = lastRowInColumnValues(grid, 14);
                for (let R = 15; R <= lastRow; R++) {
                    const dimension = String(cellValue(grid, R, 14)).trim();
                    if (dimension === "") continue;
                    if (Number(cellValue(grid, R, 16)) !== 1) continue; // solo el primer nivel

                    const attribute = String(cellValue(grid, R, 15)).trim();
                    const hierarchy = String(cellValue(grid, R, 17)).trim();

                    columns.push({
                        dimension: dimension,
                        name: hierarchy !== "" ? hierarchy : attribute,
                        isHierarchy: hierarchy !== ""
                    });
                }
            }

            // ---- Fijo (Estático/Dinámico): H12 filas, N12 columnas ----
            const rowsStatic = String(cellValue(grid, 12, 8)).trim().toUpperCase() === "X";
            const colsStatic = String(cellValue(grid, 12, 14)).trim().toUpperCase() === "X";

            // ---- Rangos: H10 filas, N10 columnas ----
            const rrAddress = String(cellValue(grid, 10, 8)).trim();
            const rcAddress = String(cellValue(grid, 10, 14)).trim();

            // ---- Celda de anclaje para la PRIMERA vez que se monta el
            // informe (EDIT_REPORT!E1). Si está vacía o no es una dirección
            // válida, se usa A1 (comportamiento anterior). Solo se aplica
            // cuando todavía no hay rrAddress/rcAddress guardados (ver
            // RangeAxis.loadFromAddresses en taskpane.js).
            const anchorAddress = String(cellValue(grid, 1, 5)).trim(); // E1

            // ---- Opciones por campo (mostrar totales, orden, expandir hasta
            // nivel, niveles visibles, formato de medida...): un único JSON
            // guardado en S1, clave "zona|dimension|nombre" -> objeto opciones.
            let fieldOptions = {};
            try {
                const raw = String(cellValue(grid, 1, 19)).trim(); // S1
                if (raw) fieldOptions = JSON.parse(raw);
            } catch (e) {
                console.warn("No se pudieron leer las opciones de campo (S1):", e);
            }

            return {
                filters, rows, columns,
                rowsStatic, colsStatic,
                rrAddress, rcAddress, anchorAddress,
                fieldOptions
            };
        });
    },

    /**
     * SaveFilters + SaveRows + SaveCols + SaveFixed + SaveRanges:
     * vuelca el estado del diseño a EDIT_REPORT, expandiendo jerarquías
     * a través de MODEL_HIER igual que hacía frmReportDesigner2.
     */
    async saveEditReportDesign(state) {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("EDIT_REPORT");
            const hierGrid = await getValuesGrid(context, "MODEL_HIER");

            // Limpiar bloques igual que hacía el VBA (ClearContents desde fila 15 hasta el final)
            sheet.getRange("C15:F5000").clear(Excel.ClearApplyTo.contents);
            // H..K = dim/attr/nivel/jerarquía; L = subtotales; M = orden (UP/DOWN)
            sheet.getRange("H15:M5000").clear(Excel.ClearApplyTo.contents);
            // N..Q = dim/attr/nivel/jerarquía; R = subtotales; S = orden (UP/DOWN)
            // (S1, usada para el JSON de opciones por campo, no se toca: el
            // borrado empieza en la fila 15)
            sheet.getRange("N15:S5000").clear(Excel.ClearApplyTo.contents);

            // Devuelve ["X"|"", "UP"|"DOWN"|""] a partir de las opciones de
            // campo (guardadas por el taskpane con clave "<zone>|DIM|NOMBRE").
            // Solo tiene sentido para atributos NO jerarquía (ver
            // renderFieldOptionsBody/buildDimensionOptionsForm): un campo de
            // jerarquía no tiene ni "Subtotales" ni "Ordenar" en el panel,
            // así que sus filas expandidas se dejan sin valor en L/M o R/S.
            // De momento estos valores NO se usan al construir el SQL.
            const subtotalAndOrder = (zoneId, dimension, name) => {
                const opts = (state.fieldOptions || {})[`${zoneId}|${dimension}|${name}`] || {};
                const subtotal = opts.showTotals ? "X" : "";
                const order = opts.sortOrder === "asc" ? "UP" : (opts.sortOrder === "desc" ? "DOWN" : "");
                return [subtotal, order];
            };

            // Niveles marcados en "Opciones de campo" para una jerarquía
            // (clave "<zone>|DIM|NOMBRE"). null = todos los niveles (valor
            // por defecto: al hacer drag&drop se rellenan todos los
            // niveles); un array = solo esos niveles se escriben en
            // EDIT_REPORT (los desmarcados se quitan).
            const visibleLevelsFor = (zoneId, dimension, name) => {
                const opts = (state.fieldOptions || {})[`${zoneId}|${dimension}|${name}`] || {};
                return Array.isArray(opts.visibleLevels) ? opts.visibleLevels : null;
            };

            // ---- SaveFilters: C,D,E,F ----
            let filaFilters = 15;
            for (const f of state.filters) {
                const row = sheet.getRangeByIndexes(filaFilters - 1, 2, 1, 4); // C..F (col index 2..5)
                row.values = [[
                    f.dimension,
                    f.realAttribute || "",
                    f.value || "",
                    f.isHierarchy ? f.name : ""
                ]];
                filaFilters++;
            }

            // ---- SaveRows: H,I,J,K (expandiendo jerarquía si aplica) ----
            let filaRows = 15;
            for (const r of state.rows) {
                if (r.isHierarchy && existsHierarchy(hierGrid, r.dimension, r.name)) {
                    const visibleLevels = visibleLevelsFor("rows", r.dimension, r.name);
                    const lastRow = lastRowInColumnValues(hierGrid, 4);
                    for (let R = 2; R <= lastRow; R++) {
                        if (String(cellValue(hierGrid, R, 4)).toUpperCase() === r.dimension.toUpperCase()
                            && String(cellValue(hierGrid, R, 2)).toUpperCase() === r.name.toUpperCase()) {
                            const nivel = Number(cellValue(hierGrid, R, 3));
                            if (visibleLevels && !visibleLevels.includes(nivel)) continue; // nivel desmarcado: se quita de EDIT_REPORT
                            const row = sheet.getRangeByIndexes(filaRows - 1, 7, 1, 4); // H..K
                            row.values = [[
                                cellValue(hierGrid, R, 4),  // DIMENSION
                                cellValue(hierGrid, R, 5),  // ATRIBUTO
                                cellValue(hierGrid, R, 3),  // NIVEL
                                cellValue(hierGrid, R, 2)   // JERARQUIA
                            ]];
                            filaRows++;
                        }
                    }
                } else {
                    const row = sheet.getRangeByIndexes(filaRows - 1, 7, 1, 4);
                    row.values = [[r.dimension, r.name, 1, ""]];
                    const [subtotal, order] = subtotalAndOrder("rows", r.dimension, r.name);
                    sheet.getRangeByIndexes(filaRows - 1, 11, 1, 2).values = [[subtotal, order]]; // L,M
                    filaRows++;
                }
            }

            // ---- SaveCols: N,O,P,Q (expandiendo jerarquía si aplica) ----
            let filaCols = 15;
            for (const c of state.columns) {
                if (c.isHierarchy && existsHierarchy(hierGrid, c.dimension, c.name)) {
                    const visibleLevels = visibleLevelsFor("columns", c.dimension, c.name);
                    const lastRow = lastRowInColumnValues(hierGrid, 4);
                    for (let R = 2; R <= lastRow; R++) {
                        if (String(cellValue(hierGrid, R, 4)).toUpperCase() === c.dimension.toUpperCase()
                            && String(cellValue(hierGrid, R, 2)).toUpperCase() === c.name.toUpperCase()) {
                            const nivel = Number(cellValue(hierGrid, R, 3));
                            if (visibleLevels && !visibleLevels.includes(nivel)) continue; // nivel desmarcado: se quita de EDIT_REPORT
                            const row = sheet.getRangeByIndexes(filaCols - 1, 13, 1, 4); // N..Q
                            row.values = [[
                                cellValue(hierGrid, R, 4),
                                cellValue(hierGrid, R, 5),
                                cellValue(hierGrid, R, 3),
                                cellValue(hierGrid, R, 2)
                            ]];
                            filaCols++;
                        }
                    }
                } else {
                    const row = sheet.getRangeByIndexes(filaCols - 1, 13, 1, 4);
                    row.values = [[c.dimension, c.name, 1, ""]];
                    const [subtotal, order] = subtotalAndOrder("columns", c.dimension, c.name);
                    sheet.getRangeByIndexes(filaCols - 1, 17, 1, 2).values = [[subtotal, order]]; // R,S
                    filaCols++;
                }
            }

            // ---- SaveFixed: H12 / N12 ----
            sheet.getRange("H12").values = [[state.rowsStatic ? "X" : ""]];
            sheet.getRange("N12").values = [[state.colsStatic ? "X" : ""]];

            // ---- SaveRanges: H10 / N10 ----
            sheet.getRange("H10").values = [[state.rrAddress || ""]];
            sheet.getRange("N10").values = [[state.rcAddress || ""]];

            // ---- Opciones por campo: JSON en S1 ----
            sheet.getRange("S1").values = [[JSON.stringify(state.fieldOptions || {})]];

            await context.sync();
        });
    },

    /**
     * Vuelca las "Propiedades del informe" (modal del taskpane) también
     * como celdas individuales en EDIT_REPORT!D2:D6 (además de guardarlas
     * en Office roaming settings, que es de donde las sigue leyendo
     * jsonTo3Matrices/getDracoReportProperties). Cada checkbox se escribe
     * como "X" (marcado) o "" (desmarcado), igual que el resto de checks
     * de la hoja (H12/N12 = Estático).
     *   D2: Suprimir ceros en filas
     *   D3: Suprimir ceros en columnas
     *   D4: Mostrar subtotales arriba
     *   D5: Sobrescribir formatos al actualizar
     *   D6: Autoajustar ancho de columnas al actualizar
     */
    async saveReportPropertiesToSheetCells(props) {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("EDIT_REPORT");
            const x = (v) => (v ? "X" : "");
            sheet.getRange("D2").values = [[x(props.suppressZeroRows)]];
            sheet.getRange("D3").values = [[x(props.suppressZeroCols)]];
            sheet.getRange("D4").values = [[x(props.subtotalsOnTop)]];
            sheet.getRange("D5").values = [[x(props.overwriteFormats)]];
            sheet.getRange("D6").values = [[x(props.autoFitColumns)]];
            await context.sync();
        });
    },

    /**
     * Lista de modelos semánticos disponibles (columna A de MODEL_FACT,
     * sin cabecera, únicos y ordenados). Usado SOLO para rellenar el
     * listbox "Modelo semántico" de Propiedades del informe: es estético,
     * de momento no cambia nada al seleccionar uno distinto.
     */
    async getSemanticModels() {
        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject("MODEL_FACT");
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) return [];

            const range = sheet.getUsedRangeOrNullObject();
            range.load(["values", "isNullObject"]);
            await context.sync();
            if (range.isNullObject) return [];

            const rows = range.values;
            const seen = new Set();
            for (let i = 1; i < rows.length; i++) { // fila 0 = cabecera
                const v = String(rows[i][0] || "").trim();
                if (v) seen.add(v);
            }
            return Array.from(seen).sort();
        });
    }
};

window.ExcelService = ExcelService;
window.ReportDesignerUtils = { addressFromRC, parseAddress, numToColLetters };
