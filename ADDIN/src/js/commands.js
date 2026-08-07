/**
 * Lógica para comandos ejecutados en segundo plano por Excel
 */
Office.onReady(() => {
    // Handshake completado para comandos
});

/**
 * Función que maneja el botón 'Ocultar panel' (HidePaneButton) definido en el manifiesto
 * @param {Office.AddinCommands.Event} event
 */
function hidePane(event) {
    try {
        if (Office.context && Office.context.ui) {
            Office.context.ui.closeContainer();
        }
    } catch (error) {
        console.error("Error al cerrar el contenedor:", error);
    }

    // OBLIGATORIO: Informar a Excel que la función finalizó para no bloquear el runtime
    if (event) {
        event.completed();
    }
}

/**
 * Ejecuta la consulta SQL en BigQuery y vuelca los resultados en la hoja activa comenzando en A1 de forma optimizada
 * @param {Office.AddinCommands.Event} event
 */
async function writeHolaInA1(event) {
    try {
        const token = localStorage.getItem("bigquery_access_token");
        const expires = localStorage.getItem("bigquery_token_expires");

        // Comprobación de token de autenticación
        if (!token || !expires || Date.now() >= parseInt(expires)) {
            console.error("No hay una sesión activa de BigQuery. Inicia sesión en el panel primero.");
            return;
        }

        const projectId = "bigqueryexcelconnector";
        const sqlQuery = "select * from `ANALYTICS.DIM_CECO`";

        // Petición a la API de BigQuery con deshabilitación de Legacy SQL
        const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/queries`, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                query: sqlQuery,
                useLegacySql: false
            })
        });

        const data = await response.json();

        if (data.error) {
            console.error("Error devuelto por la API de BigQuery:", data.error.message);
            return;
        }

        if (!data.schema || !data.schema.fields) {
            console.warn("La consulta no devolvió estructuras de datos válidas.");
            return;
        }

        const fields = data.schema.fields;
        const fieldsCount = fields.length;
        const rawRows = data.rows || [];
        const rowsCount = rawRows.length;

        // 1. Extraer los nombres de las columnas (Cabecera)
        const headers = new Array(fieldsCount);
        for (let i = 0; i < fieldsCount; i++) {
            headers[i] = fields[i].name;
        }

        // 2. Conversión optimizada de datos a matriz 2D
        const gridData = new Array(rowsCount + 1);
        gridData[0] = headers;

        for (let i = 0; i < rowsCount; i++) {
            const rowCells = rawRows[i].f;
            const rowArray = new Array(fieldsCount);
            for (let j = 0; j < fieldsCount; j++) {
                const val = rowCells[j].v;
                rowArray[j] = val !== null && val !== undefined ? val : "";
            }
            gridData[i + 1] = rowArray;
        }

        // 3. Escribir los resultados en Excel optimizando el rendimiento visual
        await Excel.run(async (context) => {
            // Suspender el redibujado de la pantalla en Excel durante la inserción
            context.workbook.application.suspendScreenUpdatingUntilNextSync();

            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const totalRows = gridData.length;

            // Definir el rango total e inyectar la matriz completa de una sola vez
            const range = sheet.getRangeByIndexes(0, 0, totalRows, fieldsCount);
            range.values = gridData;

            await context.sync();
        });

    } catch (error) {
        console.error("Error al ejecutar la consulta o pintar los datos en Excel:", error);
    } finally {
        // OBLIGATORIO: Informar a Excel que la función finalizó
        if (event) {
            event.completed();
        }
    }
}

/* ==========================================================================
 * ============  TRADUCCIÓN LITERAL DE LAS MACROS VBA (BOTÓN ACTUALIZAR)  ==
 * ==========================================================================
 * Todo lo que sigue es una traducción 1:1 (misma lógica, mismos nombres,
 * mismo orden de pasos) del módulo VBA que generaba el informe:
 *   LoadReportDefinition -> BuildSQL_Fixed -> ExecuteSQL -> JSON_PaintValues
 * No se ha optimizado ni cambiado ningún criterio de negocio, solo se ha
 * adaptado a JavaScript/Office.js porque la lectura de celdas en Excel
 * Online/Desktop vía Office.js es asíncrona y por lotes (no existe un
 * equivalente 1 a 1 de "Cells(R,C).Value" sin hacer una llamada de red por
 * celda). Por eso se cargan rangos completos una vez ("grids") y luego se
 * accede a ellos con las mismas coordenadas absolutas de fila/columna que
 * usaba el VBA (ws.Cells(R, C)).
 *
 * AJUSTES ACORDADOS CON EL USUARIO (respecto al VBA original):
 *   - ProjectId: hardcodeado a "bigqueryexcelconnector" (la hoja
 *     "M.S. - Definición" no existe en este complemento).
 *   - Token: se lee de localStorage("bigquery_access_token"), que es el
 *     mecanismo que ya usa este Add-in (la hoja "CONFIG" no existe).
 *   - Se añade (nuevo requerimiento) la escritura del SQL generado en la
 *     celda A1 de la hoja "CSV_RESULT" antes de ejecutarlo contra BigQuery.
 *
 * FUNCIÓN NO DEFINIDA EN EL VBA ORIGINAL: "GetAttributeType"
 *   Se llama desde SQLValue/BuildWhere pero no estaba definida en ningún
 *   módulo. Se ha implementado leyendo la columna "DATA_TYPE" (columna I)
 *   de la hoja MODEL_ATRIBUTES, cruzando por DIMENSION (col B) y ATRIBUTE
 *   (col C), que es exactamente el esquema que ya genera semantic_model.js
 *   al construir esa hoja. Si no encuentra el atributo, devuelve "" (se
 *   comporta como no-numérico, igual que en el flujo original).
 * ========================================================================== */

const CRLF = "\r\n";

/**
 * Estado global del informe (equivalente a las variables Public del módulo VBA)
 */
const ReportState = {
    FilterCount: 0,
    RowCount: 0,
    ColumnCount: 0,
    AttrRowCount: 0,
    AttrColumnCount: 0,
    MeasureCount: 0,
    Filters: [],
    Rows: [],
    Columns: [],
    Measures: []
};

/* ---------------------------------------------------------------------
 * Utilidades de lectura de rangos ("grids") para poder acceder a las
 * celdas por coordenada absoluta (fila,columna) igual que ws.Cells(R,C)
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

async function getFormulaGrid(context, sheetName) {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = sheet.getUsedRangeOrNullObject();
    used.load(["formulas", "rowIndex", "columnIndex", "isNullObject"]);
    await context.sync();

    if (used.isNullObject) {
        return { formulas: [], startRow: 0, startCol: 0 };
    }

    return { formulas: used.formulas, startRow: used.rowIndex, startCol: used.columnIndex };
}

// Equivalente a ws.Cells(row1, col1).Value (row1/col1 en base 1, como en VBA)
function cellValue(grid, row1, col1) {
    const r = row1 - 1 - grid.startRow;
    const c = col1 - 1 - grid.startCol;
    if (r < 0 || c < 0 || r >= grid.values.length) return "";
    const rowArr = grid.values[r];
    if (!rowArr || c >= rowArr.length) return "";
    const v = rowArr[c];
    return v === null || v === undefined ? "" : v;
}

// Equivalente a ws.Cells(row1, col1).Formula
function cellFormula(grid, row1, col1) {
    const r = row1 - 1 - grid.startRow;
    const c = col1 - 1 - grid.startCol;
    if (r < 0 || c < 0 || r >= grid.formulas.length) return "";
    const rowArr = grid.formulas[r];
    if (!rowArr || c >= rowArr.length) return "";
    const v = rowArr[c];
    return v === null || v === undefined ? "" : v;
}

// Equivalente a ws.Cells(row1, col1).HasFormula
function cellHasFormula(grid, row1, col1) {
    const f = cellFormula(grid, row1, col1);
    return typeof f === "string" && f.charAt(0) === "=";
}

// Equivalente a ws.Cells(ws.Rows.Count, col1).End(xlUp).Row sobre un grid de VALORES
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

// Equivalente a ws.Cells(ws.Rows.Count, col1).End(xlUp).Row sobre un grid de FÓRMULAS
function lastRowInColumnFormulas(grid, col1) {
    const c = col1 - 1 - grid.startCol;
    for (let r = grid.formulas.length - 1; r >= 0; r--) {
        const rowArr = grid.formulas[r];
        const v = rowArr && c >= 0 && c < rowArr.length ? rowArr[c] : "";
        if (v !== null && v !== undefined && String(v) !== "") {
            return r + 1 + grid.startRow;
        }
    }
    return grid.startRow;
}

// Equivalente a ws.Cells(row1, ws.Columns.Count).End(xlToLeft).Column sobre un grid de FÓRMULAS
function lastColInRowFormulas(grid, row1) {
    const r = row1 - 1 - grid.startRow;
    if (r < 0 || r >= grid.formulas.length) return grid.startCol;
    const rowArr = grid.formulas[r] || [];
    for (let c = rowArr.length - 1; c >= 0; c--) {
        const v = rowArr[c];
        if (v !== null && v !== undefined && String(v) !== "") {
            return c + 1 + grid.startCol;
        }
    }
    return grid.startCol;
}

// Convierte una dirección de celda tipo "$B$16" o "CSV_RESULT!$B$16" en {row, col} (base 1)
function parseAddress(addr) {
    addr = String(addr).trim();
    if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
    addr = addr.replace(/\$/g, "");
    const m = addr.match(/^([A-Za-z]+)(\d+)$/);
    if (!m) throw new Error("Dirección de celda no válida: " + addr);
    const colLetters = m[1].toUpperCase();
    const row = parseInt(m[2], 10);
    let col = 0;
    for (let i = 0; i < colLetters.length; i++) {
        col = col * 26 + (colLetters.charCodeAt(i) - 64);
    }
    return { row, col };
}

/* ---------------------------------------------------------------------
 * LoadReportDefinition / LoadFilters / LoadRows / LoadColumns
 * ------------------------------------------------------------------- */

function loadFilters(editReportGrid) {
    ReportState.FilterCount = 0;
    ReportState.Filters = [];

    let R = 15;

    while (String(cellValue(editReportGrid, R, 3)).trim() !== "") {
        ReportState.FilterCount++;
        ReportState.Filters.push({
            Dimension: String(cellValue(editReportGrid, R, 3)).trim(),
            AttributeName: String(cellValue(editReportGrid, R, 4)).trim(),
            Value: String(cellValue(editReportGrid, R, 5)).trim()
        });
        R++;
    }
}

function loadRows(editReportGrid) {
    ReportState.RowCount = 0;
    ReportState.AttrRowCount = 0;
    ReportState.MeasureCount = 0;
    ReportState.Rows = [];
    ReportState.Measures = [];

    let R = 15;

    while (String(cellValue(editReportGrid, R, 14)).trim() !== "") {
        const dim = String(cellValue(editReportGrid, R, 14)).trim();

        if (dim.toUpperCase() === "MEASURE") {
            ReportState.MeasureCount++;
            ReportState.Measures.push({
                Name: String(cellValue(editReportGrid, R, 15)).trim()
            });
        } else {
            ReportState.RowCount++;

            if (Number(cellValue(editReportGrid, R, 16)) === 1) {
                ReportState.AttrRowCount++;
            }

            const hierRaw = String(cellValue(editReportGrid, R, 16)).trim();

            ReportState.Rows.push({
                Position: Number(cellValue(editReportGrid, R, 13)),
                Dimension: dim,
                AttributeName: String(cellValue(editReportGrid, R, 15)).trim(),
                Hierarchy: hierRaw === "" ? 0 : Number(hierRaw)
            });
        }

        R++;
    }
}

function loadColumns(editReportGrid) {
    ReportState.ColumnCount = 0;
    ReportState.AttrColumnCount = 0;
    ReportState.Columns = [];

    let R = 15;

    while (String(cellValue(editReportGrid, R, 8)).trim() !== "") {
        const dim = String(cellValue(editReportGrid, R, 8)).trim();

        if (dim.toUpperCase() === "MEASURE") {
            ReportState.MeasureCount++;
            ReportState.Measures.push({
                Name: String(cellValue(editReportGrid, R, 9)).trim()
            });
        } else {
            if (Number(cellValue(editReportGrid, R, 10)) === 1) {
                ReportState.AttrColumnCount++;
            }

            ReportState.ColumnCount++;

            const hierRaw = String(cellValue(editReportGrid, R, 10)).trim();

            ReportState.Columns.push({
                Position: Number(cellValue(editReportGrid, R, 7)),
                Dimension: dim,
                AttributeName: String(cellValue(editReportGrid, R, 9)).trim(),
                Hierarchy: hierRaw === "" ? 0 : Number(hierRaw)
            });
        }

        R++;
    }
}

function loadReportDefinition(editReportGrid) {
    loadFilters(editReportGrid);
    loadRows(editReportGrid);
    loadColumns(editReportGrid);
}

/* ---------------------------------------------------------------------
 * GetTableAlias / BuscarFilaDimension / BuscarMedida / DimensionIsUsed
 * ------------------------------------------------------------------- */

function buscarFilaDimension(relGrid, dimension) {
    const lastRow = lastRowInColumnValues(relGrid, 1);

    for (let R = 2; R <= lastRow; R++) {
        if (String(cellValue(relGrid, R, 2)).trim().toUpperCase() === String(dimension).trim().toUpperCase()) {
            return Number(cellValue(relGrid, R, 1));
        }
    }

    return 0;
}

function getTableAlias(relGrid, dimension) {
    const fila = buscarFilaDimension(relGrid, dimension);
    return fila === 0 ? "f" : "d" + fila;
}

function buscarMedida(measuresGrid, measureName) {
    const lastRow = lastRowInColumnValues(measuresGrid, 1);

    for (let R = 2; R <= lastRow; R++) {
        if (String(cellValue(measuresGrid, R, 2)).trim().toUpperCase() === String(measureName).trim().toUpperCase()) {
            return R;
        }
    }

    return 0;
}

function dimensionIsUsed(dimension) {
    const upper = String(dimension).toUpperCase();

    for (let i = 0; i < ReportState.Rows.length; i++) {
        if (ReportState.Rows[i].Dimension.toUpperCase() === upper) return true;
    }
    for (let i = 0; i < ReportState.Columns.length; i++) {
        if (ReportState.Columns[i].Dimension.toUpperCase() === upper) return true;
    }
    for (let i = 0; i < ReportState.Filters.length; i++) {
        if (ReportState.Filters[i].Dimension.toUpperCase() === upper) return true;
    }

    return false;
}

/* ---------------------------------------------------------------------
 * GetAttributeType (no definida en el VBA original, ver nota arriba)
 * ------------------------------------------------------------------- */

function getAttributeType(atributesGrid, dimension, attribute) {
    const lastRow = lastRowInColumnValues(atributesGrid, 1);

    for (let R = 2; R <= lastRow; R++) {
        const dim = String(cellValue(atributesGrid, R, 2)).trim();
        const attr = String(cellValue(atributesGrid, R, 3)).trim();

        if (dim.toUpperCase() === String(dimension).trim().toUpperCase() &&
            attr.toUpperCase() === String(attribute).trim().toUpperCase()) {
            return String(cellValue(atributesGrid, R, 9)).trim().toUpperCase();
        }
    }

    return "STRING";
}

function sqlValue(atributesGrid, dimension, atributo, valor) {
    if (getAttributeType(atributesGrid, dimension, atributo) === "INTEGER") {
        return valor;
    }
    return "'" + String(valor).replace(/'/g, "''") + "'";
}

/* ---------------------------------------------------------------------
 * GetFormulaArgumentValue
 * ------------------------------------------------------------------- */

async function getFormulaArgumentValue(context, sheetName, arg) {
    arg = String(arg).trim();

    // Texto literal
    if (arg.charAt(0) === '"' && arg.charAt(arg.length - 1) === '"') {
        return arg.substring(1, arg.length - 1);
    }

    try {
        let targetSheet = sheetName;
        let address = arg;

        if (arg.indexOf("!") !== -1) {
            const parts = arg.split("!");
            targetSheet = parts[0].replace(/'/g, "");
            address = parts[1];
        }

        const sheet = context.workbook.worksheets.getItem(targetSheet);
        const range = sheet.getRange(address);
        range.load("values");
        await context.sync();

        return String(range.values[0][0]);
    } catch (err) {
        return arg.replace(/"/g, "");
    }
}

/* ---------------------------------------------------------------------
 * ReadRowDefinitions / ReadColumnDefinitions
 * ------------------------------------------------------------------- */

// Replica Replace(F, buscado, "") de VBA (sustituye TODAS las apariciones)
function replaceAll(text, search) {
    return text.split(search).join("");
}

async function readRowDefinitions(context, editReportGrid, csvGrid) {
    const items = [];

    const RRows = parseAddress(cellValue(editReportGrid, 10, 8)); // EDIT_REPORT!H10

    const lastRow = lastRowInColumnFormulas(csvGrid, RRows.col);
    const lastCol = RRows.col + ReportState.AttrColumnCount - 1;

    for (let R = RRows.row; R <= lastRow; R++) {
        for (let Col = RRows.col; Col <= lastCol; Col++) {
            if (cellHasFormula(csvGrid, R, Col)) {
                let F = String(cellFormula(csvGrid, R, Col));

                F = replaceAll(F, "=@");
                F = replaceAll(F, "=EPM_VALUE(");
                F = replaceAll(F, ")");

                const V = F.indexOf(";") !== -1 ? F.split(";") : F.split(",");

                items.push({
                    R: R,
                    Dimension: await getFormulaArgumentValue(context, "CSV_RESULT", V[0]),
                    AttributeName: await getFormulaArgumentValue(context, "CSV_RESULT", V[1]),
                    Value: await getFormulaArgumentValue(context, "CSV_RESULT", V[2]),
                    Display: await getFormulaArgumentValue(context, "CSV_RESULT", V[3])
                });
            } else {
                break; // Exit For (solo la iteración de columnas)
            }
        }
    }

    return items;
}

async function readColumnDefinitions(context, editReportGrid, csvGrid) {
    const items = [];

    const RCols = parseAddress(cellValue(editReportGrid, 10, 14)); // EDIT_REPORT!N10

    const lastCol = lastColInRowFormulas(csvGrid, RCols.row);
    const lastRow = RCols.row + ReportState.AttrRowCount - 1;

    for (let Col = RCols.col; Col <= lastCol; Col++) {
        for (let R = RCols.row; R <= lastRow; R++) {
            if (cellHasFormula(csvGrid, R, Col)) {
                let F = String(cellFormula(csvGrid, R, Col));

                F = replaceAll(F, "=@");
                F = replaceAll(F, "=EPM_VALUE(");
                F = replaceAll(F, ")");

                const V = F.indexOf(";") !== -1 ? F.split(";") : F.split(",");

                items.push({
                    R: Col,
                    Dimension: await getFormulaArgumentValue(context, "CSV_RESULT", V[0]),
                    AttributeName: await getFormulaArgumentValue(context, "CSV_RESULT", V[1]),
                    Value: await getFormulaArgumentValue(context, "CSV_RESULT", V[2]),
                    Display: await getFormulaArgumentValue(context, "CSV_RESULT", V[3])
                });
            } else {
                break; // Exit For (solo la iteración de filas)
            }
        }
    }

    return items;
}

/* ---------------------------------------------------------------------
 * BuildSelectBase / BuildFrom / BuildJoins / BuildBaseWhere /
 * BuildGroupByBase / BuildBaseRow / BuildColumns / BuildSQL_Fixed
 * ------------------------------------------------------------------- */

function buildSelectBase(relGrid, rowsDefs, colDefs) {
    const dict = new Map();

    for (const v of rowsDefs) {
        const key = (v.Dimension + "|" + v.AttributeName).toUpperCase();
        if (!dict.has(key)) {
            dict.set(key, getTableAlias(relGrid, v.Dimension) + "." + v.AttributeName);
        }
    }

    for (const v of colDefs) {
        const key = (v.Dimension + "|" + v.AttributeName).toUpperCase();
        if (!dict.has(key)) {
            dict.set(key, getTableAlias(relGrid, v.Dimension) + "." + v.AttributeName);
        }
    }

    let sql = "SELECT" + CRLF;
    for (const val of dict.values()) {
        sql += "    " + val + "," + CRLF;
    }
    sql += "    SUM(f.IMPORTE) AS IMPORTE";

    return sql;
}

function buildFrom(measuresGrid) {
    const R = buscarMedida(measuresGrid, ReportState.Measures[0].Name);

    return "FROM `" +
        cellValue(measuresGrid, R, 3) + "." +
        cellValue(measuresGrid, R, 4) + "." +
        cellValue(measuresGrid, R, 5) +
        "` f";
}

function buildJoins(relGrid) {
    let sql = "";

    const lastRow = lastRowInColumnValues(relGrid, 1);

    for (let R = 2; R <= lastRow; R++) {
        const dimname = cellValue(relGrid, R, 2);

        if (dimensionIsUsed(dimname)) {
            sql += CRLF
                + cellValue(relGrid, R, 11)
                + " JOIN `"
                + cellValue(relGrid, R, 7) + "." + cellValue(relGrid, R, 8) + "." + cellValue(relGrid, R, 9)
                + "` d" + cellValue(relGrid, R, 1) + CRLF
                + "    ON f."
                + cellValue(relGrid, R, 6)
                + " = d"
                + cellValue(relGrid, R, 1)
                + "."
                + cellValue(relGrid, R, 10)
                + CRLF;
        }
    }

    return sql;
}

function buildBaseWhere(atributesGrid, relGrid, rowsDefs, colDefs) {
    const dict = new Map();

    for (const v of colDefs) {
        const campo = getTableAlias(relGrid, v.Dimension) + "." + v.AttributeName;
        const key = (campo + "|" + v.Value).toUpperCase();
        if (!dict.has(key)) {
            dict.set(key, campo + "=" + sqlValue(atributesGrid, v.Dimension, v.AttributeName, v.Value));
        }
    }

    for (const v of rowsDefs) {
        const campo = getTableAlias(relGrid, v.Dimension) + "." + v.AttributeName;
        const key = (campo + "|" + v.Value).toUpperCase();
        if (!dict.has(key)) {
            dict.set(key, campo + "=" + sqlValue(atributesGrid, v.Dimension, v.AttributeName, v.Value));
        }
    }

    let sql = "WHERE" + CRLF + "(" + CRLF;
    let first = true;

    for (const val of dict.values()) {
        if (first) {
            sql += "      " + val;
            first = false;
        } else {
            sql += CRLF + "   OR " + val;
        }
    }

    sql += CRLF + ")";

    return sql;
}

function buildBaseRow(atributesGrid, rowsDefs) {
    const dict = new Map();

    for (const v of rowsDefs) {
        const key = String(v.R);
        const cond = "    " + v.AttributeName + "=" + sqlValue(atributesGrid, v.Dimension, v.AttributeName, v.Value);

        if (!dict.has(key)) {
            dict.set(key, cond);
        } else {
            dict.set(key, dict.get(key) + CRLF + "AND " + v.AttributeName + "=" + sqlValue(atributesGrid, v.Dimension, v.AttributeName, v.Value));
        }
    }

    let sql = "BASE_ROW AS (" + CRLF + CRLF;
    let first = true;

    for (const [rowId, cond] of dict.entries()) {
        if (!first) {
            sql += CRLF + "UNION ALL" + CRLF + CRLF;
        }

        sql += "SELECT" + CRLF
            + "    " + rowId + " AS ROW_ID," + CRLF
            + "    *" + CRLF
            + "FROM BASE" + CRLF
            + "WHERE" + CRLF
            + cond;

        first = false;
    }

    sql += CRLF + CRLF + ")";

    return sql;
}

function buildColumnsSQL(atributesGrid, colDefs) {
    const dict = new Map();

    for (const v of colDefs) {
        const key = String(v.R);
        const cond = "    " + v.AttributeName + "=" + sqlValue(atributesGrid, v.Dimension, v.AttributeName, v.Value);

        if (!dict.has(key)) {
            dict.set(key, cond);
        } else {
            dict.set(key, dict.get(key) + CRLF + "AND " + v.AttributeName + "=" + sqlValue(atributesGrid, v.Dimension, v.AttributeName, v.Value));
        }
    }

    let sql = "";
    let first = true;

    for (const [colId, cond] of dict.entries()) {
        if (!first) {
            sql += CRLF + "UNION ALL" + CRLF + CRLF;
        }

        sql += "SELECT" + CRLF
            + "    ROW_ID," + CRLF
            + "    " + colId + " AS COLUMN_ID," + CRLF
            + "    SUM(IMPORTE) AS IMPORTE" + CRLF
            + "FROM BASE_ROW" + CRLF
            + "WHERE" + CRLF
            + cond + CRLF
            + "GROUP BY ROW_ID";

        first = false;
    }

    sql += CRLF + CRLF
        + "ORDER BY" + CRLF
        + "    ROW_ID," + CRLF
        + "    COLUMN_ID";

    return sql;
}

function buildGroupByBase(relGrid, rowsDefs, colDefs) {
    const dict = new Map();

    for (const v of rowsDefs) {
        const key = (v.Dimension + "|" + v.AttributeName).toUpperCase();
        if (!dict.has(key)) {
            dict.set(key, getTableAlias(relGrid, v.Dimension) + "." + v.AttributeName);
        }
    }

    for (const v of colDefs) {
        const key = (v.Dimension + "|" + v.AttributeName).toUpperCase();
        if (!dict.has(key)) {
            dict.set(key, getTableAlias(relGrid, v.Dimension) + "." + v.AttributeName);
        }
    }

    let sql = "GROUP BY" + CRLF;
    for (const val of dict.values()) {
        sql += "    " + val + "," + CRLF;
    }

    sql = sql.slice(0, -3); // Left(SQL, Len(SQL) - 3)

    return sql;
}

async function buildSQLFixed(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid) {
    const rowsDefs = await readRowDefinitions(context, editReportGrid, csvGrid);
    const colDefs = await readColumnDefinitions(context, editReportGrid, csvGrid);

    let sql = "";

    sql += "WITH BASE AS (" + CRLF + CRLF;
    sql += buildSelectBase(relGrid, rowsDefs, colDefs) + CRLF + CRLF;
    sql += buildFrom(measuresGrid) + CRLF + CRLF;
    sql += buildJoins(relGrid) + CRLF + CRLF;
    sql += buildBaseWhere(atributesGrid, relGrid, rowsDefs, colDefs) + CRLF + CRLF;
    sql += buildGroupByBase(relGrid, rowsDefs, colDefs) + CRLF + CRLF;
    sql += ")," + CRLF + CRLF;
    sql += buildBaseRow(atributesGrid, rowsDefs) + CRLF + CRLF;
    sql += buildColumnsSQL(atributesGrid, colDefs);

    return sql;
}

/* ---------------------------------------------------------------------
 * ExecuteSQL / EscapeJSON
 * ------------------------------------------------------------------- */

function escapeJSON(text) {
    text = text.split("\\").join("\\\\");
    text = text.split('"').join('\\"');
    text = text.split("\r\n").join("\\n");
    text = text.split("\r").join("\\n");
    text = text.split("\n").join("\\n");
    return text;
}

async function executeSQL(sql) {
    const token = localStorage.getItem("bigquery_access_token");
    const expires = localStorage.getItem("bigquery_token_expires");

    if (!token || !expires || Date.now() >= parseInt(expires)) {
        throw new Error("No hay una sesión activa de BigQuery. Inicia sesión en el panel primero.");
    }

    const projectId = "bigqueryexcelconnector";

    const url = "https://bigquery.googleapis.com/bigquery/v2/projects/" + projectId + "/queries";

    const body = '{"query":"' + escapeJSON(sql) + '","useLegacySql":false}';

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
 * JSON_PaintValues
 * ------------------------------------------------------------------- */

function parseJsonValueTriples(json) {
    const triples = [];

    let pos = 0;
    let campo = 0;
    let rowId = 0;
    let colId = 0;

    while (true) {
        const idx = json.indexOf('"v":', pos);
        if (idx === -1) break;

        let ini = idx + 4;

        while (json.charAt(ini) === ":" || json.charAt(ini) === " ") {
            ini++;
        }

        let fin;
        let texto;

        if (json.charAt(ini) === '"') {
            ini++;
            fin = json.indexOf('"', ini);
            texto = json.substring(ini, fin);
        } else {
            fin = json.indexOf("}", ini);
            texto = json.substring(ini, fin);
        }

        campo++;

        if (campo === 1) {
            rowId = parseInt(texto, 10);
        } else if (campo === 2) {
            colId = parseInt(texto, 10);
        } else if (campo === 3) {
            triples.push({ row: rowId, col: colId, text: texto });
            campo = 0;
        }

        pos = fin + 1;
    }

    return triples;
}

// Igual que asignar Texto a FormulaR1C1 en VBA: si es numérico, Excel lo interpreta como número
function coerceCellLiteral(text) {
    if (text.trim() !== "" && !isNaN(text)) {
        return Number(text);
    }
    return text;
}

async function jsonPaintValues(context, json) {
    const triples = parseJsonValueTriples(json);
    const sheet = context.workbook.worksheets.getItem("CSV_RESULT");

    for (const t of triples) {
        const range = sheet.getRangeByIndexes(t.row - 1, t.col - 1, 1, 1);
        range.values = [[coerceCellLiteral(t.text)]];
    }

    await context.sync();
}

/* ---------------------------------------------------------------------
 * Actualizar_informe_fixed -> actualizarInformeFixed (botón "Actualizar")
 * ------------------------------------------------------------------- */

/**
 * Traducción literal de Actualizar_informe_fixed():
 *   LoadReportDefinition
 *   SQL = BuildSQL_Fixed
 *   [NUEVO] escribir SQL en CSV_RESULT!A1
 *   Json = ExecuteSQL(SQL)
 *   JSON_PaintValues(Json)
 * @param {Office.AddinCommands.Event} event
 */
async function actualizarInformeFixed(event) {
    try {
        let sql;

        // 1) LoadReportDefinition + BuildSQL_Fixed + escritura de A1
        await Excel.run(async (context) => {
            const editReportGrid = await getValuesGrid(context, "EDIT_REPORT");
            const relGrid = await getValuesGrid(context, "MODEL_RELATIONSHIP");
            const measuresGrid = await getValuesGrid(context, "MODEL_MEASURES");
            const atributesGrid = await getValuesGrid(context, "MODEL_ATRIBUTES");
            const csvGrid = await getFormulaGrid(context, "CSV_RESULT");

            loadReportDefinition(editReportGrid);

            sql = await buildSQLFixed(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid);

            const csvSheet = context.workbook.worksheets.getItem("CSV_RESULT");
            csvSheet.getRange("A1").values = [[sql]];

            await context.sync();
        });

        // 2) ExecuteSQL contra BigQuery
        const json = await executeSQL(sql);

        // 3) JSON_PaintValues
        await Excel.run(async (context) => {
            await jsonPaintValues(context, json);
        });

    } catch (error) {
        console.error("Error al actualizar el informe:", error);
    } finally {
        // OBLIGATORIO: Informar a Excel que la función finalizó
        if (event) {
            event.completed();
        }
    }
}

// Asociar el nombre del comando del manifiesto con la función JavaScript
Office.actions.associate("hidePane", hidePane);
Office.actions.associate("writeHolaInA1", writeHolaInA1);
Office.actions.associate("actualizarInformeFixed", actualizarInformeFixed);
