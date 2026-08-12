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
// Si es un rango (p.ej. "A9:B9"), toma la primera celda, igual que .Row/.Column de un Range en VBA
function parseAddress(addr) {
    addr = String(addr).trim();
    if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
    if (addr.indexOf(":") !== -1) addr = addr.split(":")[0];
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

// NOTA: en el VBA original esta función podía "Evaluate" una referencia de
// celda si el argumento no venía entre comillas. En la práctica, el código
// que pinta las fórmulas EPM_VALUE(dimension, atributo, valor, display)
// siempre escribe los 4 argumentos como literales entre comillas, así que
// esa rama nunca se ejecuta con datos reales. Se ha simplificado a una
// función síncrona (sin llamadas a Excel) que solo quita las comillas,
// evitando además que un argumento inesperado sin comillas provoque un
// context.sync() fallido a mitad de la construcción del SQL, lo que podía
// dejar el resto del lote en un estado inestable.


/* ---------------------------------------------------------------------
 * ReadRowDefinitions / ReadColumnDefinitions
 * ------------------------------------------------------------------- */

// Replica Replace(F, buscado, "") de VBA (sustituye TODAS las apariciones)
function replaceAll(text, search) {
    return text.split(search).join("");
}



async function getFormulaArgumentValue(context, arg) {
    arg = String(arg).trim();

    // Si empieza por comillas, es texto literal
    if (arg.charAt(0) === '"') {
        return arg.substring(1, arg.length - 1);
    }

    // Si no tiene comillas, es una referencia a una celda
    const range = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(arg);

    range.load("values");

    await context.sync();

    return range.values[0][0];
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

                const V = F.indexOf(";") !== -1
                    ? F.split(";")
                    : F.split(",");

                items.push({
                    R: R,
                    Dimension: await getFormulaArgumentValue(context, V[0]),
                    AttributeName: await getFormulaArgumentValue(context, V[1]),
                    Value: await getFormulaArgumentValue(context, V[2]),
                    Display: await getFormulaArgumentValue(context, V[3])
                });

            } else {
                break;
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

                const V = F.indexOf(";") !== -1
                    ? F.split(";")
                    : F.split(",");

                items.push({
                    R: Col,
                    Dimension: await getFormulaArgumentValue(context, V[0]),
                    AttributeName: await getFormulaArgumentValue(context, V[1]),
                    Value: await getFormulaArgumentValue(context, V[2]),
                    Display: await getFormulaArgumentValue(context, V[3])
                });

            } else {
                break;
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

function indentBlock(text, prefix) {
    return text.split(CRLF).map(line => prefix + line).join(CRLF);
}

/**
 * Construye, para las definiciones de un eje (filas o columnas), la
 * expresión ARRAY(SELECT <idFieldName> FROM UNNEST([STRUCT(...), ...]) WHERE COND)
 * que sustituye al patrón "un SELECT/UNION ALL por cada ROW_ID/COLUMN_ID".
 * Agrupa igual que hacían BuildBaseRow/BuildColumns (por v.R), concatenando
 * las condiciones de cada grupo con AND, y genera un único STRUCT por
 * ROW_ID/COLUMN_ID en vez de un SELECT completo contra BASE/BASE_ROW.
 */
function buildIdArray(atributesGrid, defs, idFieldName) {
    const dict = new Map();

    for (const v of defs) {
        const key = String(v.R);
        const condPart = v.AttributeName + "=" + sqlValue(atributesGrid, v.Dimension, v.AttributeName, v.Value);

        if (!dict.has(key)) {
            dict.set(key, [condPart]);
        } else {
            dict.get(key).push(condPart);
        }
    }

    const structs = [];
    for (const [id, conds] of dict.entries()) {
        structs.push("STRUCT(" + id + " AS " + idFieldName + ", (" + conds.join(" AND ") + ") AS COND)");
    }

    let sql = "ARRAY(" + CRLF;
    sql += "    SELECT " + idFieldName + " FROM UNNEST([" + CRLF;
    sql += structs.map(s => "        " + s).join("," + CRLF) + CRLF;
    sql += "    ])" + CRLF;
    sql += "    WHERE COND" + CRLF;
    sql += ")";

    return sql;
}

async function buildSQLFixed(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid) {
    const rowsDefs = await readRowDefinitions(context, editReportGrid, csvGrid);
    const colDefs = await readColumnDefinitions(context, editReportGrid, csvGrid);

    // Diagnóstico: abre las herramientas de desarrollador (F12) del panel de
    // tareas / comandos para ver exactamente qué Dimension/Attribute/Value
    // se ha leído de las fórmulas EPM_VALUE de CSV_RESULT.
    console.log("readRowDefinitions ->", rowsDefs);
    console.log("readColumnDefinitions ->", colDefs);

    const rowIdsArray = buildIdArray(atributesGrid, rowsDefs, "ROW_ID");
    const columnIdsArray = buildIdArray(atributesGrid, colDefs, "COLUMN_ID");

    let sql = "";

    // ---- CTE BASE: se escanea la tabla de hechos UNA sola vez ----
    sql += "WITH BASE AS (" + CRLF + CRLF;
    sql += buildSelectBase(relGrid, rowsDefs, colDefs) + CRLF + CRLF;
    sql += buildFrom(measuresGrid) + CRLF + CRLF;
    sql += buildJoins(relGrid) + CRLF + CRLF;
    sql += buildBaseWhere(atributesGrid, relGrid, rowsDefs, colDefs) + CRLF + CRLF;
    sql += buildGroupByBase(relGrid, rowsDefs, colDefs) + CRLF;
    sql += ")," + CRLF + CRLF;

    // ---- CTE TAGGED: etiqueta cada fila de BASE con los ROW_ID/COLUMN_ID
    //      que cumple, sin volver a escanear BASE por cada uno ----
    sql += "TAGGED AS (" + CRLF + CRLF;
    sql += "    SELECT" + CRLF;
    sql += "        IMPORTE," + CRLF + CRLF;
    sql += indentBlock(rowIdsArray, "        ") + " AS ROW_IDS," + CRLF + CRLF;
    sql += indentBlock(columnIdsArray, "        ") + " AS COLUMN_IDS" + CRLF + CRLF;
    sql += "    FROM BASE" + CRLF;
    sql += ")" + CRLF + CRLF;

    // ---- SELECT final: cruce ROW_ID x COLUMN_ID vía UNNEST + agregación ----
    sql += "SELECT" + CRLF;
    sql += "    ROW_ID," + CRLF;
    sql += "    COLUMN_ID," + CRLF;
    sql += "    SUM(IMPORTE) AS IMPORTE" + CRLF;
    sql += "FROM TAGGED," + CRLF;
    sql += "UNNEST(ROW_IDS) AS ROW_ID," + CRLF;
    sql += "UNNEST(COLUMN_IDS) AS COLUMN_ID" + CRLF;
    sql += "GROUP BY ROW_ID, COLUMN_ID" + CRLF;
    sql += "ORDER BY ROW_ID, COLUMN_ID";

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
 * Núcleo de Actualizar_informe_fixed(), sin manejo de `event` (para poder
 * reutilizarlo tanto desde el botón del ribbon como desde el dispatcher
 * Actualizar()).
 */
async function actualizarInformeFixedCore() {
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
}

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
        await actualizarInformeFixedCore();
    } catch (error) {
        console.error("Error al actualizar el informe (fixed):", error);
    } finally {
        if (event) {
            event.completed();
        }
    }
}

/* ==========================================================================
 * ===============  TRADUCCIÓN DEL FLUJO "Actualizar_informe"  ============
 * ==========================================================================
 * Este es el flujo "Dinámico" (con GROUPING SETS), usado cuando al menos
 * uno de los dos ejes NO está marcado como Estático en EDIT_REPORT (H12/N12).
 * A diferencia del flujo _Fixed, aquí SÍ se soportan jerarquías desplegables
 * en el informe (totales/subtotales vía GROUPING SETS).
 *
 * FUNCIÓN NO DEFINIDA EN NINGÚN MÓDULO VBA QUE ME HAYAS PASADO: "GetMeasureSQL"
 *   La llaman BuildSelect/BuildSelectEx pero no apareció su código en ningún
 *   sitio. La he implementado leyendo MODEL_MEASURES (columnas AGGREGATION=G
 *   y FACT_FIELD=F) y generando "AGREGACION(f.CAMPO) AS NombreMedida" — por
 *   ejemplo "SUM(f.IMPORTE) AS IMPORTE". Si tu implementación real hace algo
 *   distinto (otro alias, otra tabla, formato...), dímelo y la ajusto.
 * ========================================================================== */

/**
 * GetMeasureSQL — ver nota de la cabecera de esta sección sobre la
 * implementación asumida.
 */
function getMeasureSQL(measuresGrid, measureName) {
    const R = buscarMedida(measuresGrid, measureName);
    if (R === 0) return "";
    const aggregation = String(cellValue(measuresGrid, R, 7)).trim(); // G: AGGREGATION
    const field = String(cellValue(measuresGrid, R, 6)).trim();       // F: FACT_FIELD
    return aggregation + "(f." + field + ") AS " + measureName;
}

/**
 * BuildSelect: columnas del eje Columnas, luego eje Filas, luego medidas.
 */
function buildSelect(measuresGrid, relGrid) {
    let sql = "SELECT" + CRLF;
    let first = true;

    for (const c of ReportState.Columns) {
        if (!first) sql += "," + CRLF;
        sql += "    " + getTableAlias(relGrid, c.Dimension) + "." + c.AttributeName;
        first = false;
    }

    for (const r of ReportState.Rows) {
        if (!first) sql += "," + CRLF;
        sql += "    " + getTableAlias(relGrid, r.Dimension) + "." + r.AttributeName;
        first = false;
    }

    for (const m of ReportState.Measures) {
        if (!first) sql += "," + CRLF;
        sql += "    " + getMeasureSQL(measuresGrid, m.Name);
        first = false;
    }

    return sql;
}

/**
 * BuildWhere: a diferencia de SQLValue (flujo Fixed), aquí se consideran
 * sin comillas varios tipos numéricos de BigQuery, no solo "INTEGER".
 */
const UNQUOTED_TYPES = ["INTEGER", "INT64", "FLOAT", "NUMERIC", "BIGNUMERIC"];

function buildWhere(atributesGrid, relGrid) {
    // Filtros "vacíos" (sin valor seleccionado en el taskpane) no deben
    // añadirse al WHERE: se ignoran por completo, como si no existieran.
    const activeFilters = ReportState.Filters.filter(f => String(f.Value).trim() !== "");

    if (activeFilters.length === 0) return "";

    let sql = "WHERE" + CRLF;

    for (let i = 0; i < activeFilters.length; i++) {
        const f = activeFilters[i];
        const tipo = getAttributeType(atributesGrid, f.Dimension, f.AttributeName);

        sql += getTableAlias(relGrid, f.Dimension) + "." + f.AttributeName + " = ";

        if (UNQUOTED_TYPES.includes(tipo)) {
            sql += f.Value;
        } else {
            sql += "'" + String(f.Value).replace(/'/g, "''") + "'";
        }

        if (i < activeFilters.length - 1) {
            sql += CRLF + "AND ";
        }
    }

    return sql;
}

/**
 * GenerateHierarchyCombinations: traducción del "odómetro" en VBA. Dado un
 * array de longitudes máximas [h1, h2, ...], genera todas las combinaciones
 * desde [h1,h2,...] hasta [1,1,...] en orden decreciente (nunca llega a 0).
 */
function generateHierarchyCombinations(hierarchies) {
    const result = [];
    const current = hierarchies.slice();

    while (true) {
        result.push(current.slice());

        let i = current.length - 1;
        let continueLoop = false;

        while (true) {
            current[i] -= 1;

            if (current[i] > 0) {
                break;
            } else {
                current[i] = hierarchies[i];
                i -= 1;

                if (i < 0) {
                    continueLoop = true;
                    break;
                }
            }
        }

        if (continueLoop) break;
    }

    return result;
}

/**
 * Buildconfigsets: construye la cláusula GROUP BY GROUPING SETS (...)
 * combinando niveles de jerarquía de Filas y Columnas.
 */
function buildConfigSets(relGrid) {
    // ---- DIMENSIONS: Filas primero, luego Columnas ----
    const dimensions = [null]; // índice 1-based, dimensions[0] sin usar

    for (const r of ReportState.Rows) {
        dimensions.push(getTableAlias(relGrid, r.Dimension) + "." + r.AttributeName);
    }
    for (const c of ReportState.Columns) {
        dimensions.push(getTableAlias(relGrid, c.Dimension) + "." + c.AttributeName);
    }

    // ---- HIERARCHIES: longitud de cada tramo de "misma dimensión consecutiva" ----
    const hierarchies = [];

    let lastDim = "";
    let contador = 0;
    for (const r of ReportState.Rows) {
        if (r.Dimension !== lastDim) {
            if (lastDim !== "") hierarchies.push(contador);
            lastDim = r.Dimension;
            contador = 1;
        } else {
            contador++;
        }
    }
    if (ReportState.Rows.length > 0) hierarchies.push(contador);

    lastDim = "";
    contador = 0;
    for (const c of ReportState.Columns) {
        if (c.Dimension !== lastDim) {
            if (lastDim !== "") hierarchies.push(contador);
            lastDim = c.Dimension;
            contador = 1;
        } else {
            contador++;
        }
    }
    if (ReportState.Columns.length > 0) hierarchies.push(contador);

    // ---- HIERARCHIES ACUM ----
    const hierarchiesAcum = [];
    let acum = 0;
    for (const h of hierarchies) {
        acum += h;
        hierarchiesAcum.push(acum);
    }

    const combinations = generateHierarchyCombinations(hierarchies);

    let texto = "";
    let first = true;

    for (const item of combinations) {
        if (first) {
            texto += "GROUP BY GROUPING SETS ( ( ";
            first = false;
        } else {
            texto += ",( ";
        }

        for (let i = 0; i < item.length; i++) {
            for (let j = 1; j <= item[i]; j++) {
                if (item[i] > 0) {
                    if (i !== 0 || j !== 1) {
                        texto += ", ";
                    }
                    if (i === 0) {
                        texto += dimensions[j];
                    } else {
                        texto += dimensions[j + hierarchiesAcum[i - 1]];
                    }
                } else {
                    texto += "[TOTAL]";
                }
            }
        }

        texto += " )";
    }

    texto += " )";

    return texto;
}

/**
 * BuildFinalSelect: traducción literal, incluyendo el cruce ROW_ID/COLUMN_ID
 * "invertido" tal cual está en tu VBA (ROW_ID se calcula a partir del eje
 * Columnas y COLUMN_ID a partir del eje Filas) — se deja igual, según lo
 * acordado.
 */
function buildFinalSelect() {
    let sql = "SELECT" + CRLF + CRLF;

    sql += "    DENSE_RANK() OVER (" + CRLF + "        ORDER BY";
    let first = true;
    for (const c of ReportState.Columns) {
        if (c.Hierarchy > 0) {
            sql += first ? (CRLF + "            " + c.AttributeName) : ("," + CRLF + "            " + c.AttributeName);
            first = false;
        }
    }
    sql += CRLF + "    ) AS ROW_ID," + CRLF + CRLF;

    sql += "    DENSE_RANK() OVER (" + CRLF + "        ORDER BY";
    first = true;
    for (const r of ReportState.Rows) {
        if (r.Hierarchy > 0) {
            sql += first ? (CRLF + "            " + r.AttributeName) : ("," + CRLF + "            " + r.AttributeName);
            first = false;
        }
    }
    sql += CRLF + "    ) AS COLUMN_ID," + CRLF + CRLF;

    first = true;
    for (const c of ReportState.Columns) {
        sql += first ? ("    " + c.AttributeName) : ("," + CRLF + "    " + c.AttributeName);
        first = false;
    }
    for (const r of ReportState.Rows) {
        sql += "," + CRLF + "    " + r.AttributeName;
    }

    sql += "," + CRLF + "    IMPORTE" + CRLF + CRLF + "FROM REPORT_DATA";

    return sql;
}

/**
 * BuildFinalWhere: excluye de cada eje las filas de detalle que no
 * correspondan al grouping-set de totales de esa jerarquía (IS NULL).
 */
function buildFinalWhere() {
    let sql = "";
    let currentDim = "";
    let first = true;

    // ---- COLUMNAS ----
    for (const c of ReportState.Columns) {
        if (c.Hierarchy > 0) {
            if (currentDim !== c.Dimension.toUpperCase()) {
                if (currentDim !== "") {
                    sql += CRLF + CRLF + ")" + CRLF + CRLF + "AND NOT (" + CRLF;
                } else {
                    sql += "WHERE NOT (" + CRLF;
                }
                currentDim = c.Dimension.toUpperCase();
                first = true;
            }

            sql += first
                ? ("    " + c.AttributeName + " IS NULL")
                : (CRLF + "    AND " + c.AttributeName + " IS NULL");
            first = false;
        }
    }

    sql += CRLF + ")";

    // ---- FILAS ----
    currentDim = "";

    for (const r of ReportState.Rows) {
        if (r.Hierarchy > 0) {
            if (currentDim !== r.Dimension.toUpperCase()) {
                if (currentDim !== "") {
                    sql += CRLF + CRLF + ")" + CRLF + CRLF + "AND NOT (" + CRLF;
                } else {
                    sql += CRLF + CRLF + "AND NOT (" + CRLF;
                }
                currentDim = r.Dimension.toUpperCase();
                first = true;
            }

            sql += first
                ? ("    " + r.AttributeName + " IS NULL")
                : (CRLF + "    AND " + r.AttributeName + " IS NULL");
            first = false;
        }
    }

    if (currentDim !== "") {
        sql += CRLF + ")";
    }

    return sql;
}

/**
 * BuildSQL: ensambla el flujo "Dinámico" completo.
 */
function buildSQL(relGrid, measuresGrid, atributesGrid) {
    let sql = "";

    sql += "WITH REPORT_DATA AS (" + CRLF + CRLF;
    sql += buildSelect(measuresGrid, relGrid) + CRLF + CRLF;
    sql += buildFrom(measuresGrid) + CRLF + CRLF;
    sql += buildJoins(relGrid) + CRLF + CRLF;
    sql += buildWhere(atributesGrid, relGrid) + CRLF + CRLF;
    sql += buildConfigSets(relGrid) + CRLF + CRLF;
    sql += ")" + CRLF + CRLF;
    sql += buildFinalSelect() + CRLF + CRLF;
    sql += buildFinalWhere() + CRLF + CRLF;
    sql += "ORDER BY" + CRLF + "    ROW_ID," + CRLF + "    COLUMN_ID";

    return sql;
}

/* ---------------------------------------------------------------------
 * Expandir/Contraer jerarquías (Filas/Columnas) sin ocultar filas de
 * Excel: se filtra y renumera el propio resultado antes de pintarlo, y
 * se mantiene en memoria (dura mientras el libro está abierto; no se
 * guarda en la hoja) qué nodos están contraídos para cada eje, mientras
 * ese eje no cambie de campos.
 * ------------------------------------------------------------------- */

const DracoCollapseState = {
    rows: { signature: null, collapsed: new Set() },
    cols: { signature: null, collapsed: new Set() }
};

let DracoLastJson = null;           // último JSON de BigQuery, para repintar sin re-consultar
let DracoHandlerRegistered = false; // evita registrar el listener de clic más de una vez
let DracoSuppressChangeEvents = false; // true mientras jsonTo3Matrices pinta celdas (evita que el reconocimiento de miembros reaccione a nuestras propias escrituras)
let DracoIndicatorMap = new Map();  // "row_col" -> { axis, nodeKey } de los indicadores +/- pintados

// Firma del eje = lista de "DIMENSION.ATRIBUTO:NIVEL" de sus campos, en
// orden. Si cambia (se añade/quita/reordena un campo en ESE eje), se
// entiende que la jerarquía cambió y se resetea su estado de contraído.
function computeDracoAxisSignature(editReportGrid, dimCol, attrCol, hierCol, count) {
    const parts = [];
    for (let i = 1; i <= count; i++) {
        const R = i + 14;
        const dim = String(cellValue(editReportGrid, R, dimCol)).trim().toUpperCase();
        const attr = String(cellValue(editReportGrid, R, attrCol)).trim().toUpperCase();
        const nivel = String(cellValue(editReportGrid, R, hierCol)).trim();
        parts.push(dim + "." + attr + ":" + nivel);
    }
    return parts.join("|");
}

function resetDracoCollapseIfAxisChanged(axis, signature) {
    const st = DracoCollapseState[axis];
    if (st.signature !== signature) {
        st.signature = signature;
        st.collapsed = new Set();
    }
}

/**
 * Filtra (oculta hijos de nodos contraídos) y renumera 1..N el diccionario
 * de un eje (rowDict o colDict), a partir del set de nodos contraídos.
 *
 * "Profundidad" de una fila = longitud del prefijo contiguo de valores no
 * nulos empezando en el campo 1 (si aparece un NULL, el resto de campos
 * posteriores —aunque tengan valor, p.ej. otro campo plano detrás de la
 * jerarquía— no cuenta: son datos de otra jerarquía/campo, no de esta).
 */
function filterAndCompactDracoAxis(dict, count, collapsedSet, flags, options) {
    const opts = options || {};
    const subtotalsOnTop = !!opts.subtotalsOnTop;

    const items = [];
    for (const V of dict.values()) {
        let deepest = 0;
        for (let i = 1; i <= count; i++) {
            if (String(V[i]).toLowerCase().indexOf("null") === 0) break;
            deepest = i;
        }
        items.push({ oldId: Number(V[0]), V, deepest });
    }
    items.sort((a, b) => a.oldId - b.oldId);

    function prefixKey(item, p) {
        const parts = [];
        for (let i = 1; i <= p; i++) parts.push(String(item.V[i]));
        return p + "\u00A7" + parts.join("\u241F");
    }

    // "Mostrar subtotales arriba" (propiedades del informe): por defecto se
    // respeta el orden natural del resultado (subtotal justo antes que su
    // detalle, que es como suele venir el GROUPING SETS); si el usuario NO
    // quiere los subtotales arriba, se reordena para que cada nodo-resumen
    // pase a ir DESPUÉS de todos sus descendientes.
    if (!subtotalsOnTop) {
        items.sort((a, b) => {
            const len = Math.min(a.deepest, b.deepest);
            let samePrefix = true;
            for (let i = 1; i <= len; i++) {
                if (String(a.V[i]) !== String(b.V[i])) { samePrefix = false; break; }
            }
            if (samePrefix && a.deepest !== b.deepest) {
                // Uno es ancestro (subtotal) del otro: el ancestro va DESPUÉS.
                return a.deepest < b.deepest ? 1 : -1;
            }
            return a.oldId - b.oldId; // sin relación de parentesco directa: orden natural
        });
    }

    // Excluidas: alguna de sus filas ancestro (p < profundidad propia) está contraída.
    const kept = items.filter(item => {
        for (let p = 1; p < item.deepest; p++) {
            if (collapsedSet.has(prefixKey(item, p))) return false;
        }
        return true;
    });

    // ¿Tiene hijos? — existe OTRA fila del resultado (contraída o no) que
    // comparte el mismo prefijo y llega más profundo.
    function hasChildren(item) {
        if (item.deepest <= 0) return false;
        const key = prefixKey(item, item.deepest);
        return items.some(other => other.deepest > item.deepest && prefixKey(other, item.deepest) === key);
    }

    // Columna/fila física (iAux) en la que cae el nivel `deepest` de este
    // nodo — es la MISMA celda donde ya se pinta su etiqueta.
    function fieldAtDeepest(item) {
        let field = 0;
        for (let k = 1; k <= item.deepest; k++) {
            if (flags[k] === 1) field++;
        }
        return field;
    }

    const dictOut = new Map();
    const idMap = new Map();     // oldId -> newId (para remapear FACT)
    const indicators = [];       // { newId, field, nodeKey, collapsed }

    kept.forEach((item, idx) => {
        const newId = idx + 1;
        idMap.set(item.oldId, newId);
        const arr = item.V.slice();
        arr[0] = newId;
        dictOut.set(String(newId), arr);

        if (item.deepest > 0 && hasChildren(item)) {
            const nodeKey = prefixKey(item, item.deepest);
            indicators.push({ newId, field: fieldAtDeepest(item), nodeKey, collapsed: collapsedSet.has(nodeKey) });
        }
    });

    return { dict: dictOut, idMap, indicators };
}

/**
 * Devuelve las propiedades del informe (nombre, suprimir ceros, subtotales
 * arriba, sobrescribir formatos, autoajustar columnas), guardadas por el
 * modal "Propiedades del informe" del taskpane en Office roaming settings.
 * Accesible desde cualquier contexto (taskpane o commands.html/ribbon).
 */
function getDracoReportProperties() {
    const defaults = {
        reportName: "Report 001",
        suppressZeroRows: false,
        suppressZeroCols: false,
        subtotalsOnTop: true,
        overwriteFormats: true,
        autoFitColumns: false
    };
    try {
        const raw = Office.context.document.settings.get("draco_reportProperties");
        if (!raw) return defaults;
        const parsed = JSON.parse(raw);
        return Object.assign({}, defaults, parsed);
    } catch (e) {
        console.warn("No se pudieron leer las propiedades del informe, se usan valores por defecto:", e);
        return defaults;
    }
}

// Construye el literal de fórmula EPM_VALUE("DIM","ATRIBUTO","VALOR","DISPLAY")
// usado para "congelar" como texto editable las celdas de un eje marcado
// como Estático (mismo formato que readRowDefinitions/readColumnDefinitions
// ya saben leer para el flujo Fijo).
function buildEpmValueFormula(dim, attr, text) {
    const esc = (s) => String(s === null || s === undefined ? "" : s).replace(/"/g, '""');
    return '=EPM_VALUE("' + esc(dim) + '","' + esc(attr) + '","' + esc(text) + '","' + esc(text) + '")';
}

/* ---------------------------------------------------------------------
 * Conversión INMEDIATA (sin esperar al próximo refresco/BigQuery) de las
 * celdas ya pintadas de Draco_001_Rows / Draco_001_Cols entre texto plano
 * y fórmula EPM_VALUE, al marcar/desmarcar un eje como Estático desde el
 * taskpane. Reutiliza la misma correspondencia {dim, attr} por nivel que
 * usa jsonTo3Matrices (columnas H/I para Filas, N/O para Columnas de
 * EDIT_REPORT), pero solo toca el rango con nombre ya existente: si
 * todavía no se ha pintado ninguna tabla no hace nada (el próximo
 * refresco la pintará ya en el modo correcto).
 * @param {"rows"|"columns"} axis
 * @param {boolean} makeStatic true = texto -> EPM_VALUE; false = EPM_VALUE -> texto
 */
// {dim, attr} por posición (nivel) de un eje, igual que fieldsFilas/fieldsColumnas
// en jsonTo3Matrices. Requiere que ReportState ya esté cargado (loadReportDefinition).
function buildAxisFieldsTable(editReportGrid, axis) {
    const totalDimFilas = ReportState.ColumnCount; // nº de niveles del eje Filas
    const totalDimCols = ReportState.RowCount;      // nº de niveles del eje Columnas
    const fields = [];
    if (axis === "rows") {
        for (let i = 1; i <= totalDimFilas; i++) {
            fields[i] = { dim: cellValue(editReportGrid, i + 14, 8), attr: cellValue(editReportGrid, i + 14, 9) };
        }
    } else {
        for (let i = 1; i <= totalDimCols; i++) {
            fields[i] = { dim: cellValue(editReportGrid, i + 14, 14), attr: cellValue(editReportGrid, i + 14, 15) };
        }
    }
    return fields;
}

async function convertAxisStaticFormulas(axis, makeStatic) {
    const rangeName = axis === "rows" ? "Draco_001_Rows" : "Draco_001_Cols";
    DracoSuppressChangeEvents = true;
    try {

    await Excel.run(async (context) => {
        const editReportGrid = await getValuesGrid(context, "EDIT_REPORT");
        loadReportDefinition(editReportGrid);
        const fields = buildAxisFieldsTable(editReportGrid, axis);

        const namedRange = context.workbook.names.getItemOrNullObject(rangeName);
        namedRange.load("isNullObject");
        await context.sync();
        if (namedRange.isNullObject) {
            // Todavía no hay tabla pintada: nada que convertir ahora mismo.
            return;
        }

        const range = namedRange.getRange();
        range.load(["values", "formulas", "rowCount", "columnCount"]);
        await context.sync();

        const GLYPH_PREFIX = /^[▸▾]\s+/; // indicador +/- fusionado (solo aplica en eje Dinámico)
        const EPM_RE = /^=\s*EPM_VALUE\s*\(/i;
        const EPM_VALOR_RE = /EPM_VALUE\s*\(\s*"(?:[^"]|"")*"\s*,\s*"(?:[^"]|"")*"\s*,\s*"((?:[^"]|"")*)"/i;

        const newFormulas = [];
        let changed = false;

        for (let r = 0; r < range.rowCount; r++) {
            const rowOut = [];
            for (let c = 0; c < range.columnCount; c++) {
                const currentValue = range.values[r][c];
                const currentFormula = range.formulas[r][c];

                if (currentValue === "" || currentValue === null || currentValue === undefined) {
                    rowOut.push(currentFormula);
                    continue;
                }

                // Nivel del campo dentro del eje: en Filas crece por COLUMNA,
                // en Columnas crece por FILA (misma orientación que al pintar).
                const level = (axis === "rows" ? c : r) + 1;
                const field = fields[level];
                const isEpmFormula = typeof currentFormula === "string" && EPM_RE.test(currentFormula);

                if (makeStatic) {
                    if (isEpmFormula || !field) {
                        rowOut.push(currentFormula);
                        continue;
                    }
                    const text = String(currentValue).replace(GLYPH_PREFIX, "");
                    rowOut.push(buildEpmValueFormula(field.dim, field.attr, text));
                    changed = true;
                } else {
                    if (!isEpmFormula) {
                        rowOut.push(currentFormula);
                        continue;
                    }
                    const match = currentFormula.match(EPM_VALOR_RE);
                    const text = match ? match[1].replace(/""/g, '"') : String(currentValue);
                    rowOut.push(text);
                    changed = true;
                }
            }
            newFormulas.push(rowOut);
        }

        if (changed) {
            range.values = newFormulas;
            await context.sync();
        }
    });

    } finally {
        DracoSuppressChangeEvents = false;
    }
}

// Convierte un número de columna (1-based) en letras de columna Excel ("A", "AB"...)
function dracoColToLetters(col) {
    let s = "";
    while (col > 0) {
        const rem = (col - 1) % 26;
        s = String.fromCharCode(65 + rem) + s;
        col = Math.floor((col - 1) / 26);
    }
    return s;
}

/**
 * Limpia todos los comentarios (notas) previos de la hoja y añade uno por
 * cada celda indicadora registrada en DracoIndicatorMap, a modo de tooltip
 * que se ve al pasar el ratón por encima ("esto se puede clicar").
 */
async function refreshDracoIndicatorTooltips(context, sheet) {
    try {
        sheet.comments.load("items");
        await context.sync();
        if (sheet.comments.items.length > 0) {
            sheet.comments.items.forEach(c => c.delete());
            await context.sync();
        }
    } catch (e) {
        console.warn("No se pudieron limpiar los tooltips Draco previos:", e);
        return;
    }

    if (DracoIndicatorMap.size === 0) return;

    try {
        for (const key of DracoIndicatorMap.keys()) {
            const [rowStr, colStr] = key.split("_");
            const addr = dracoColToLetters(Number(colStr)) + rowStr;
            context.workbook.comments.add("'CSV_RESULT'!" + addr, "Clic para expandir / contraer");
        }
        await context.sync();
    } catch (e) {
        console.warn("No se pudieron crear los tooltips Draco:", e);
    }
}

/**
 * Localiza, dentro de Draco_001_Rows/Draco_001_Cols, a qué {dim, attr}
 * corresponde una celda concreta de CSV_RESULT (equivalente a
 * GetDimAttrRows/GetDimAttrCols del VBA, pero usando los rangos con
 * nombre en lugar de recalcular límites a mano).
 * Devuelve null si la celda no pertenece a ninguno de los dos ejes.
 */
async function locateDracoAxisField(context, addr) {
    const sheet = context.workbook.worksheets.getItem("CSV_RESULT");
    const cell = sheet.getRange(addr);
    cell.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

    const rowsNamed = context.workbook.names.getItemOrNullObject("Draco_001_Rows");
    const colsNamed = context.workbook.names.getItemOrNullObject("Draco_001_Cols");
    rowsNamed.load("isNullObject");
    colsNamed.load("isNullObject");
    await context.sync();

    if (cell.rowCount !== 1 || cell.columnCount !== 1) return null;

    let axis = null;
    let level = null;

    if (!rowsNamed.isNullObject) {
        const r = rowsNamed.getRange();
        r.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
        await context.sync();
        const withinRows = cell.rowIndex >= r.rowIndex && cell.rowIndex < r.rowIndex + r.rowCount
            && cell.columnIndex >= r.columnIndex && cell.columnIndex < r.columnIndex + r.columnCount;
        if (withinRows) {
            axis = "rows";
            level = (cell.columnIndex - r.columnIndex) + 1;
        }
    }

    if (!axis && !colsNamed.isNullObject) {
        const c = colsNamed.getRange();
        c.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
        await context.sync();
        const withinCols = cell.rowIndex >= c.rowIndex && cell.rowIndex < c.rowIndex + c.rowCount
            && cell.columnIndex >= c.columnIndex && cell.columnIndex < c.columnIndex + c.columnCount;
        if (withinCols) {
            axis = "columns";
            level = (cell.rowIndex - c.rowIndex) + 1;
        }
    }

    if (!axis) return null;

    const editReportGrid = await getValuesGrid(context, "EDIT_REPORT");
    loadReportDefinition(editReportGrid);
    const fields = buildAxisFieldsTable(editReportGrid, axis);
    const field = fields[level];
    if (!field || !field.dim) return null;

    return { axis, level, dim: field.dim, attr: field.attr };
}

/**
 * "Reconocimiento de miembros" (traducción de Workbook_SheetChange +
 * Helpvalue del VBA): con el pulsador activado, si el usuario teclea un
 * valor en una celda de Draco_001_Rows/Draco_001_Cols que todavía no es
 * una fórmula EPM_VALUE, se abre el buscador de miembros (FilterModal,
 * la misma ventana que usan los filtros) precargado con lo escrito, y al
 * elegir un valor se sustituye la celda por la fórmula EPM_VALUE
 * correspondiente. Si se cancela, se deja el texto tal cual (igual que
 * en VBA).
 *
 * LIMITACIÓN DE LA PLATAFORMA: Office.js no expone un evento de doble
 * clic sobre una celda (a diferencia de Workbook_SheetBeforeDoubleClick
 * en VBA); Excel.Worksheet solo ofrece onChanged (tras escribir+Enter) y
 * onSelectionChanged (al cambiar de celda seleccionada). Por eso este
 * "reconocimiento" se dispara al escribir un valor (onChanged, fiable al
 * 100%) y, como aproximación al doble clic, también al hacer clic sobre
 * una celda de esos rangos que esté VACÍA (onSelectionChanged) — abre el
 * buscador directamente sin necesidad de escribir nada antes.
 */
async function handleDracoMemberRecognitionChanged(eventArgs) {
    try {
        if (DracoSuppressChangeEvents) {
            console.log("[Draco] onChanged ignorado: escritura programática en curso (refresco/pintado).");
            return;
        }
        if (!Office.context.document.settings.get("draco_memberRecognition")) return;

        let addr = (eventArgs && eventArgs.address) || "";
        if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
        if (!addr) return;
        console.log("[Draco] onChanged: reconocimiento de miembros activo, evaluando celda", addr);

        let located = null;
        let currentText = "";

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("CSV_RESULT");
            const cell = sheet.getRange(addr);
            cell.load(["rowCount", "columnCount", "values", "formulas"]);
            await context.sync();
            if (cell.rowCount !== 1 || cell.columnCount !== 1) return;

            const value = cell.values[0][0];
            const formula = cell.formulas[0][0];
            if (value === "" || value === null || value === undefined) return;
            // Ya es EPM_VALUE (por ejemplo, porque nosotros mismos la acabamos de
            // escribir): salir para no reabrir el buscador en bucle.
            if (typeof formula === "string" && /^=\s*EPM_VALUE\s*\(/i.test(formula)) return;

            currentText = String(value);
            located = await locateDracoAxisField(context, addr);
        });

        console.log("[Draco] onChanged: resultado de locateDracoAxisField ->", located);
        if (!located) return;
        await openMemberRecognitionPicker(addr, located, currentText);
    } catch (e) {
        console.error("Error en el reconocimiento de miembros:", e);
    }
}

/**
 * Aproximación al doble clic (ver comentario anterior): clic sobre una
 * celda VACÍA de Draco_001_Rows/Draco_001_Cols con el reconocimiento
 * activado abre directamente el buscador de miembros.
 */
async function handleDracoMemberRecognitionSelection(eventArgs) {
    try {
        if (DracoSuppressChangeEvents) return;
        if (!Office.context.document.settings.get("draco_memberRecognition")) return;

        let addr = (eventArgs && eventArgs.address) || "";
        if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
        if (!addr) return;
        console.log("[Draco] onSelectionChanged: reconocimiento de miembros activo, evaluando celda", addr);

        let located = null;

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("CSV_RESULT");
            const cell = sheet.getRange(addr);
            cell.load(["rowCount", "columnCount", "values"]);
            await context.sync();
            if (cell.rowCount !== 1 || cell.columnCount !== 1) return;

            const value = cell.values[0][0];
            if (value !== "" && value !== null && value !== undefined) return; // solo celdas vacías

            located = await locateDracoAxisField(context, addr);
        });

        console.log("[Draco] onSelectionChanged: resultado de locateDracoAxisField ->", located);
        if (!located) return;
        await openMemberRecognitionPicker(addr, located, "");
    } catch (e) {
        console.error("Error en el reconocimiento de miembros (clic):", e);
    }
}

/**
 * Parser mínimo del JSON de BigQuery a una lista plana {text, attribute,
 * value} — copia local de loadJsonTree (filterModal.js) para no depender
 * del DOM del taskpane: el diálogo del picker se ejecuta en su PROPIA
 * ventana/contexto y no puede acceder a nada del taskpane directamente.
 */
function parseMemberJsonTree(json) {
    const fieldMatches = [...json.matchAll(/"name":\s*"([^"]+)"/g)];
    const campos = fieldMatches.map((m) => m[1]);
    if (campos.length === 0) return [];

    const ultimos = campos.map(() => "");
    const valores = campos.map(() => "");
    const valueMatches = [...json.matchAll(/"v":\s*"([^"]*)"/g)];

    const items = [];
    let nivel = 0;

    for (const m of valueMatches) {
        valores[nivel] = m[1];
        nivel++;
        if (nivel > campos.length - 1) {
            for (let i = 0; i < campos.length; i++) {
                if (valores[i] !== ultimos[i]) {
                    items.push({ text: " ".repeat(i * 4) + valores[i], attribute: campos[i], value: valores[i] });
                }
            }
            for (let i = 0; i < campos.length; i++) ultimos[i] = valores[i];
            nivel = 0;
        }
    }
    return items;
}

/**
 * Abre el buscador de miembros como una ventana de DIÁLOGO de Office
 * (Office.context.ui.displayDialogAsync): aparece centrada sobre la
 * ventana de Excel, en SU PROPIA ventana, sin necesidad de mostrar el
 * taskpane en absoluto (a diferencia de FilterModal, que vive dentro del
 * DOM de taskpane.html).
 *
 * Un diálogo no tiene acceso al modelo de objetos de Excel, así que:
 *   1) Aquí (con Excel.run disponible) se resuelve primero la lista de
 *      valores (buildFilterValuesSQL + executeSQL, igual que FilterModal).
 *   2) Se abre el diálogo (memberPicker.html) y, en cuanto avisa que está
 *      listo, se le envían los items por mensaje (dialog.messageChild).
 *   3) Cuando el usuario elige uno (o cancela), el diálogo manda un
 *      mensaje de vuelta (DialogMessageReceived) y aquí se escribe la
 *      fórmula EPM_VALUE y se cierra el diálogo.
 *
 * NOTA: Office.js no permite "anclar" un diálogo a la posición de una
 * celda concreta (no hay API para eso); displayDialogAsync solo permite
 * centrarlo sobre la ventana de Excel con un tamaño (%) dado, que es lo
 * más parecido a "en medio del Excel" que soporta la plataforma.
 */
async function openMemberRecognitionPicker(addr, located, initialSearch) {
    console.log("[Draco] openMemberRecognitionPicker: iniciando para", addr, located);

    let items = [];
    try {
        const sql = await window.ExcelService.buildFilterValuesSQL(located.dim, located.attr);
        console.log("[Draco] SQL de valores construida:", sql);
        if (!sql) {
            console.warn("[Draco] No se ha podido construir el SQL (dim/attr no reconocidos):", located);
            return;
        }
        const json = await window.ExcelService.executeSQL(sql);
        items = parseMemberJsonTree(json);
        console.log("[Draco] Nº de items recibidos para el picker:", items.length);
    } catch (err) {
        console.error("[Draco] Error obteniendo los valores del picker:", err);
        return;
    }

    const dialogUrl = new URL("memberPicker.html", window.location.href).href;
    console.log("[Draco] Abriendo diálogo del picker en:", dialogUrl);

    await new Promise((resolve) => {
        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 55, width: 28, displayInIframe: false },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error(
                        "[Draco] displayDialogAsync ha fallado:",
                        asyncResult.error && asyncResult.error.code,
                        asyncResult.error && asyncResult.error.message
                    );
                    resolve();
                    return;
                }

                console.log("[Draco] Diálogo abierto correctamente.");
                const dialog = asyncResult.value;
                let settled = false;

                const closeDialog = () => {
                    try { dialog.close(); } catch (e) { /* ya cerrado */ }
                };

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                    console.log("[Draco] Mensaje recibido del diálogo:", arg.message);
                    let payload;
                    try {
                        payload = JSON.parse(arg.message);
                    } catch (err) {
                        console.error("[Draco] Mensaje del diálogo no es JSON válido:", err);
                        return;
                    }

                    if (payload.type === "ready") {
                        console.log("[Draco] Diálogo listo: enviando items…");
                        dialog.messageChild(JSON.stringify({
                            items,
                            initialSearch,
                            title: `${located.dim} / ${located.attr}`
                        }));
                        return;
                    }

                    if (payload.type === "select") {
                        console.log("[Draco] Valor elegido:", payload.value, "/", payload.attribute);
                        settled = true;
                        closeDialog();
                        try {
                            await Excel.run(async (context) => {
                                const sheet = context.workbook.worksheets.getItem("CSV_RESULT");
                                const cell = sheet.getRange(addr);
                                cell.values = [[buildEpmValueFormula(located.dim, payload.attribute, payload.value)]];
                                await context.sync();
                            });
                            console.log("[Draco] Fórmula EPM_VALUE escrita correctamente en", addr);
                        } catch (err) {
                            console.error("[Draco] Error escribiendo la fórmula EPM_VALUE:", err);
                        }
                        resolve();
                        return;
                    }

                    if (payload.type === "cancel") {
                        console.log("[Draco] Selección cancelada por el usuario.");
                        settled = true;
                        closeDialog();
                        resolve();
                    }
                });

                dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                    // 12006 = el usuario cerró el diálogo con la X.
                    console.warn("[Draco] DialogEventReceived:", arg.error);
                    if (!settled) resolve();
                });
            }
        );
    });
}


async function handleEditReportMemberPickerRequest(eventArgs) {
    try {
        if (!eventArgs || !eventArgs.address) return;

        if (!eventArgs.address.toUpperCase().includes("A5")) return;

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("EDIT_REPORT");

            const values = sheet.getRange("A1:A5");
            values.load("values");

            await context.sync();

            const dimension = String(values.values[0][0] || "").trim();
            const attribute = String(values.values[1][0] || "").trim();
            const searchValue = String(values.values[2][0] || "").trim();
            const addr = String(values.values[3][0] || "").trim();
            const request = String(values.values[4][0] || "").trim().toUpperCase();

            if (request !== "X") return;

            // Consumimos la petición
            sheet.getRange("A5").clear(Excel.ClearApplyTo.contents);
            await context.sync();

            if (!dimension || !attribute || !addr) return;

            await openMemberRecognitionPicker(
                addr,
                {
                    dim: dimension,
                    attr: attribute
                },
                searchValue
            );
        });

    } catch (error) {
        console.error(
            "[Draco] Error procesando petición Member Picker:",
            error
        );
    }
}



async function registerDracoSelectionHandler(context, sheet) {
    if (DracoHandlerRegistered) return;

    sheet.onSelectionChanged.add(handleDracoSelectionChanged);
    sheet.onSelectionChanged.add(handleDracoMemberRecognitionSelection);
    sheet.onChanged.add(handleDracoMemberRecognitionChanged);

    // NUEVO: petición de apertura del Member Picker desde EDIT_REPORT
    const editReport = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
    editReport.load("isNullObject");
    await context.sync();

    if (!editReport.isNullObject) {
        editReport.onChanged.add(handleEditReportMemberPickerRequest);
    }

    await context.sync();
    DracoHandlerRegistered = true;
}

/**
 * Registro TEMPRANO de los listeners de CSV_RESULT (indicadores +/- y
 * Reconocimiento de miembros), independiente de que se haya pulsado
 * "Actualizar informe" alguna vez. Se llama desde TaskPaneApp.init(), que
 * se ejecuta en cuanto arranca el Shared Runtime (con CUALQUIER clic del
 * ribbon, no solo al abrir el taskpane) — así el reconocimiento de
 * miembros funciona aunque el panel nunca se haya abierto en la sesión.
 * Si la hoja CSV_RESULT todavía no existe (primer uso, sin refresco
 * previo) no hace nada; el refresco la registrará igualmente cuando
 * pinte la primera tabla.
 */
async function ensureDracoHandlersRegistered() {
    if (DracoHandlerRegistered) return;
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject("CSV_RESULT");
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) {
                console.log("[Draco] ensureDracoHandlersRegistered: CSV_RESULT no existe todavía (sin refresco previo).");
                return;
            }
            await registerDracoSelectionHandler(context, sheet);
            console.log("[Draco] Listeners de CSV_RESULT registrados de forma temprana (sin necesidad de refrescar).");
        });
    } catch (e) {
        console.warn("[Draco] No se pudieron registrar los listeners de forma temprana:", e);
    }
}

/**
 * Handler del clic (selección de una sola celda) sobre un indicador +/-:
 * alterna el estado contraído/expandido de ese nodo y repinta reutilizando
 * el último JSON de BigQuery (sin volver a consultar).
 */
async function handleDracoSelectionChanged(eventArgs) {
    if (DracoIndicatorMap.size === 0) return;

    try {
        let addr = (eventArgs && eventArgs.address) || "";
        if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
        if (!addr) return;

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("CSV_RESULT");
            const range = sheet.getRange(addr);
            range.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
            await context.sync();

            if (range.rowCount !== 1 || range.columnCount !== 1) return;

            const key = (range.rowIndex + 1) + "_" + (range.columnIndex + 1);
            const meta = DracoIndicatorMap.get(key);
            if (!meta) return;

            const st = DracoCollapseState[meta.axis];
            if (st.collapsed.has(meta.nodeKey)) {
                st.collapsed.delete(meta.nodeKey);
            } else {
                st.collapsed.add(meta.nodeKey);
            }

            if (DracoLastJson) {
                await jsonTo3Matrices(context, DracoLastJson);
            }
        });
    } catch (e) {
        console.error("Error al contraer/expandir jerarquía Draco:", e);
    }
}

/**
 * JSON_To_3_Matrices: traducción literal, incluyendo el mismo cruce
 * "invertido" H/N que usa el resto del módulo (total_dim_filas=ColumnCount
 * leído de columnas H/I/J; total_dim_cols=RowCount leído de columnas N/O/P).
 */
async function jsonTo3Matrices(context, json) {
    DracoSuppressChangeEvents = true; // evita que el reconocimiento de miembros reaccione a este pintado
    try {
        return await jsonTo3MatricesCore(context, json);
    } finally {
        DracoSuppressChangeEvents = false;
    }
}

async function jsonTo3MatricesCore(context, json) {
    DracoLastJson = json; // cache para poder repintar en un toggle +/- sin re-consultar BigQuery

    const totalCampos = 2 + ReportState.RowCount + ReportState.ColumnCount + ReportState.MeasureCount;

    // ---- Extraer todos los "v" ----
    const valores = [];
    let pos = 0;

    while (true) {
        const idx = json.indexOf('"v":', pos);
        if (idx === -1) break;

        let ini = idx + 5;
        while (json.charAt(ini) === " ") ini++;

        let fin, texto;
        if (json.charAt(ini) === '"') {
            ini++;
            fin = json.indexOf('"', ini);
            texto = json.substring(ini, fin);
        } else {
            fin = json.indexOf("}", ini);
            texto = json.substring(ini, fin);
        }

        valores.push(texto);
        pos = fin + 1;
    }

    const filas = Math.floor(valores.length / totalCampos);

    // ---- FACT ----
    const fact = [];
    for (let i = 0; i < filas; i++) {
        const base = i * totalCampos;
        const row = [];
        for (let j = 0; j < totalCampos; j++) row.push(valores[base + j]);
        fact.push(row); // fact[i][j] con j 0-based (VBA era 1-based)
    }

    const rowDict = new Map();
    const colDict = new Map();

    for (let i = 0; i < filas; i++) {
        const f = fact[i]; // f[0]=ROW_ID, f[1]=COLUMN_ID, f[2..]=dims, f[last]=IMPORTE

        // ROW KEY
        let rowKey = f[0];
        for (let j = 1; j <= ReportState.ColumnCount; j++) rowKey += "|" + f[1 + j];
        if (!rowDict.has(rowKey)) {
            const arr = [f[0]]; // arr[0]=ROW_ID
            for (let j = 1; j <= ReportState.ColumnCount; j++) arr.push(f[1 + j]);
            rowDict.set(rowKey, arr);
        }

        // COLUMN KEY
        let colKey = f[1];
        for (let j = 1; j <= ReportState.RowCount; j++) colKey += "|" + f[1 + ReportState.ColumnCount + j];
        if (!colDict.has(colKey)) {
            const arr = [f[1]]; // arr[0]=COLUMN_ID
            for (let j = 1; j <= ReportState.RowCount; j++) arr.push(f[1 + ReportState.ColumnCount + j]);
            colDict.set(colKey, arr);
        }
    }

    // ---- PINTAR ----
    const editReportGrid = await getValuesGrid(context, "EDIT_REPORT");

    const RRows = parseAddress(cellValue(editReportGrid, 10, 8));  // EDIT_REPORT!H10
    const RCols = parseAddress(cellValue(editReportGrid, 10, 14)); // EDIT_REPORT!N10

    const rowsOffRow = RRows.row - 1;
    const rowsOffCol = RRows.col - 1;
    const colsOffRow = RCols.row - 1;
    const colsOffCol = RCols.col - 1;

    const totalDimFilas = ReportState.ColumnCount;
    const totalDimCols = ReportState.RowCount;

    // Estático/Dinámico por eje (H12 filas, N12 columnas) y propiedades
    // generales del informe (nombre, suprimir ceros, subtotales arriba,
    // sobrescribir formatos...), guardadas desde el taskpane.
    const rowsStatic = String(cellValue(editReportGrid, 12, 8)).trim().toUpperCase() === "X";
    const colsStatic = String(cellValue(editReportGrid, 12, 14)).trim().toUpperCase() === "X";
    const reportProps = getDracoReportProperties();

    // Precalcular una sola vez el flag (NIVEL) de cada posición del eje,
    // en vez de releerlo de EDIT_REPORT en cada iteración de cada ROW_ID/COLUMN_ID.
    const flagsFilas = [];
    const fieldsFilas = []; // {dim, attr} por posición — solo hace falta si el eje es Estático (fórmulas EPM_VALUE)
    for (let i = 1; i <= totalDimFilas; i++) {
        flagsFilas[i] = Number(cellValue(editReportGrid, i + 14, 10)); // J
        fieldsFilas[i] = { dim: cellValue(editReportGrid, i + 14, 8), attr: cellValue(editReportGrid, i + 14, 9) }; // H, I
    }

    const flagsColumnas = [];
    const fieldsColumnas = [];
    for (let i = 1; i <= totalDimCols; i++) {
        flagsColumnas[i] = Number(cellValue(editReportGrid, i + 14, 16)); // P
        fieldsColumnas[i] = { dim: cellValue(editReportGrid, i + 14, 14), attr: cellValue(editReportGrid, i + 14, 15) }; // N, O
    }

    // ---- Expandir/Contraer: si el eje (misma lista de campos) no ha
    // cambiado desde el último refresco, se respeta qué nodos estaban
    // contraídos; si ha cambiado, se resetea (todo expandido) para ESE eje.
    const rowsSignature = computeDracoAxisSignature(editReportGrid, 8, 9, 10, totalDimFilas);
    const colsSignature = computeDracoAxisSignature(editReportGrid, 14, 15, 16, totalDimCols);
    resetDracoCollapseIfAxisChanged("rows", rowsSignature);
    resetDracoCollapseIfAxisChanged("cols", colsSignature);

    // "Suprimir ceros en filas/columnas" (propiedades del informe): se quita
    // del diccionario, ANTES de compactar, cualquier ROW_ID/COLUMN_ID cuyo
    // valor sea siempre cero (o vacío) en todo el FACT.
    if (reportProps.suppressZeroRows || reportProps.suppressZeroCols) {
        const rowsWithValue = new Set();
        const colsWithValue = new Set();
        for (let i = 0; i < filas; i++) {
            const f = fact[i];
            const n = Number(f[totalCampos - 1]);
            if (!isNaN(n) && n !== 0) {
                rowsWithValue.add(String(f[0]));
                colsWithValue.add(String(f[1]));
            }
        }
        if (reportProps.suppressZeroRows) {
            for (const k of Array.from(rowDict.keys())) {
                if (!rowsWithValue.has(String(rowDict.get(k)[0]))) rowDict.delete(k);
            }
        }
        if (reportProps.suppressZeroCols) {
            for (const k of Array.from(colDict.keys())) {
                if (!colsWithValue.has(String(colDict.get(k)[0]))) colDict.delete(k);
            }
        }
    }

    const rowsFilter = filterAndCompactDracoAxis(rowDict, totalDimFilas, DracoCollapseState.rows.collapsed, flagsFilas, { subtotalsOnTop: reportProps.subtotalsOnTop });
    const colsFilter = filterAndCompactDracoAxis(colDict, totalDimCols, DracoCollapseState.cols.collapsed, flagsColumnas, { subtotalsOnTop: reportProps.subtotalsOnTop });

    console.log("jsonTo3Matrices diagnóstico:", {
        totalCampos, filas,
        RowCount: ReportState.RowCount, ColumnCount: ReportState.ColumnCount, MeasureCount: ReportState.MeasureCount,
        H10: cellValue(editReportGrid, 10, 8), N10: cellValue(editReportGrid, 10, 14),
        RRows, RCols, rowsOffRow, rowsOffCol, colsOffRow, colsOffCol,
        totalDimFilas, totalDimCols,
        rowsStatic, colsStatic, reportProps,
        rowDictSize: rowDict.size, colDictSize: colDict.size,
        filasVisibles: rowsFilter.dict.size, columnasVisibles: colsFilter.dict.size
    });

    const sheet = context.workbook.worksheets.getItem("CSV_RESULT");

    // Antes de refrescar: borrar formato Y contenido de los rangos con
    // nombre Draco_001_Rows / Draco_001_Cols / Draco_001_Values de la
    // ejecución anterior (si existen), para no arrastrar colores/bordes
    // de una tabla previa más grande o con otra forma.
    await clearDracoNamedRanges(context);

    // Limpiar todo lo pintado en ejecuciones anteriores (incluidas posibles
    // fórmulas EPM_VALUE residuales de una tabla previa más grande, y su
    // formato: relleno, fuente, bordes..., y también la franja de
    // indicadores +/- de la ejecución anterior), pero preservando la fila 1
    // (A1=SQL, B1=JSON) que se escribe aparte.
    DracoIndicatorMap = new Map();
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load(["isNullObject", "rowIndex", "rowCount"]);
    await context.sync();

    if (!usedRange.isNullObject) {
        const firstDataRow = Math.max(usedRange.rowIndex, 1); // índice 0-based: fila 2 en adelante
        const lastRow = usedRange.rowIndex + usedRange.rowCount; // exclusivo
        if (lastRow > firstDataRow) {
            sheet.getRangeByIndexes(firstDataRow, 0, lastRow - firstDataRow, 16384)
                .clear(Excel.ClearApplyTo.all);
        }
    }
    await context.sync();

    /* -------------------------------------------------------------
     * 1) FACT — construir el bloque completo en memoria y escribirlo
     *    de una sola vez con un único range.values = [...]. Las filas
     *    cuyo ROW_ID/COLUMN_ID original haya quedado oculto por una
     *    jerarquía contraída se descartan, y los IDs restantes se
     *    remapean a la numeración compactada 1..N.
     * ----------------------------------------------------------- */
    const factCells = new Map(); // "row_col" -> {row, col, value}
    for (let i = 0; i < filas; i++) {
        const f = fact[i];
        const newRowId = rowsFilter.idMap.get(Number(f[0]));
        const newColId = colsFilter.idMap.get(Number(f[1]));
        if (newRowId === undefined || newColId === undefined) continue; // oculto por contraído
        const row = newRowId + rowsOffRow;
        const col = newColId + colsOffCol;
        const value = coerceCellLiteral(String(f[totalCampos - 1]));
        factCells.set(row + "_" + col, { row, col, value });
    }
    await writeCellBlock(context, sheet, factCells);

    /* -------------------------------------------------------------
     * 2) FILAS — mismo cálculo que antes (última entrada no-nula por
     *    ROW_ID "gana", igual que el pintado secuencial original),
     *    pero acumulado en un Map en vez de escribir celda a celda,
     *    y usando el diccionario ya filtrado/renumerado.
     * ----------------------------------------------------------- */
    const filasCells = new Map(); // "row_col" -> {row, col, value, indent, field}
    let maxRowId = 0;
    for (const V of rowsFilter.dict.values()) {
        if (Number(V[0]) > maxRowId) maxRowId = Number(V[0]);

        let iAux = 0;
        for (let i = 1; i <= totalDimFilas; i++) {
            const flag = flagsFilas[i];
            if (flag === 1) iAux++;

            if (String(V[i]).toLowerCase().indexOf("null") !== 0) {
                const row = Number(V[0]) + rowsOffRow;
                const col = iAux + rowsOffCol;
                const indent = flag === 1 ? 0 : Math.max(0, flag - 1);
                const text = coerceCellLiteral(V[i]);
                // Eje Estático: se escribe como fórmula EPM_VALUE editable
                // (mismo formato que lee el flujo Fijo), no como texto plano.
                const cellVal = rowsStatic
                    ? buildEpmValueFormula(fieldsFilas[i].dim, fieldsFilas[i].attr, text)
                    : text;
                // Sobrescribe si ya había una entrada (mismo comportamiento que
                // el bucle secuencial original: el último nivel no-nulo gana).
                filasCells.set(row + "_" + col, { row, col, value: cellVal, indent, field: iAux });
            }
        }
    }

    // Fusionar el indicador +/- en la misma celda del nivel (nodos con
    // hijos), en vez de una columna aparte: "− España" / "+ España".
    // No aplica si el eje es Estático (esas celdas ya son fórmulas EPM_VALUE
    // editables a mano; anteponer un glifo de texto las rompería).
    if (!rowsStatic) {
        const byLogicalKey = new Map(rowsFilter.indicators.map(it => [it.newId + "_" + it.field, it]));
        for (const [physKey, cell] of filasCells) {
            const newId = cell.row - rowsOffRow;
            const ind = byLogicalKey.get(newId + "_" + cell.field);
            if (ind) {
                const glyph = ind.collapsed ? "▸" : "▾";
                cell.value = glyph + " " + cell.value;
                DracoIndicatorMap.set(physKey, { axis: "rows", nodeKey: ind.nodeKey });
            }
        }
    }

    await writeCellBlock(context, sheet, filasCells);
    if (reportProps.overwriteFormats) {
        await writeIndentAndColorRuns(context, sheet, filasCells, "col", rowsOffCol);
    }

    /* -------------------------------------------------------------
     * 3) COLUMNAS — análogo a FILAS
     * ----------------------------------------------------------- */
    const columnasCells = new Map();
    let maxColId = 0;
    for (const V of colsFilter.dict.values()) {
        if (Number(V[0]) > maxColId) maxColId = Number(V[0]);

        let iAux = 0;
        for (let i = 1; i <= totalDimCols; i++) {
            const flag = flagsColumnas[i];
            if (flag === 1) iAux++;

            if (String(V[i]).toLowerCase().indexOf("null") !== 0) {
                const row = iAux + colsOffRow;
                const col = Number(V[0]) + colsOffCol;
                const indent = flag === 1 ? 0 : Math.max(0, flag - 1);
                const text = coerceCellLiteral(V[i]);
                const cellVal = colsStatic
                    ? buildEpmValueFormula(fieldsColumnas[i].dim, fieldsColumnas[i].attr, text)
                    : text;
                columnasCells.set(row + "_" + col, { row, col, value: cellVal, indent, field: iAux });
            }
        }
    }

    // Igual que en FILAS: el glifo +/- se fusiona en la propia celda
    // (salvo eje Estático, ver comentario arriba).
    if (!colsStatic) {
        const byLogicalKey = new Map(colsFilter.indicators.map(it => [it.newId + "_" + it.field, it]));
        for (const [physKey, cell] of columnasCells) {
            const newId = cell.col - colsOffCol;
            const ind = byLogicalKey.get(newId + "_" + cell.field);
            if (ind) {
                const glyph = ind.collapsed ? "▸" : "▾";
                cell.value = glyph + " " + cell.value;
                DracoIndicatorMap.set(physKey, { axis: "cols", nodeKey: ind.nodeKey });
            }
        }
    }

    await writeCellBlock(context, sheet, columnasCells);
    if (reportProps.overwriteFormats) {
        await writeIndentAndColorRuns(context, sheet, columnasCells, "row", colsOffRow);
    }

    /* -------------------------------------------------------------
     * 4) RANGOS CON NOMBRE Draco_001_Rows / Draco_001_Cols / Draco_001_Values
     *    + formato general (Segoe UI 9, número en Values, bordes finos),
     *    salvo que "Sobrescribir formatos" esté desactivado en las
     *    propiedades del informe (entonces solo se (re)definen los rangos).
     * ----------------------------------------------------------- */
    await applyDracoNamedRanges(context, sheet, {
        RRows, RCols, totalDimFilas, totalDimCols, maxRowId, maxColId,
        applyVisualFormat: reportProps.overwriteFormats
    });

    /* -------------------------------------------------------------
     * 5) Tooltip al pasar el ratón ("Clic para expandir / contraer")
     *    sobre cada celda que lleva el indicador +/- fusionado.
     * ----------------------------------------------------------- */
    await refreshDracoIndicatorTooltips(context, sheet);

    // 6) Registrar (una sola vez por sesión) los listeners de clic/edición:
    //    resuelven los indicadores +/- pintados arriba y el "Reconocimiento
    //    de miembros" sobre las celdas de Draco_001_Rows/Draco_001_Cols.
    await registerDracoSelectionHandler(context, sheet);

    // 7) Autoajustar ancho de columnas (propiedades del informe), si procede.
    if (reportProps.autoFitColumns) {
        try {
            sheet.getUsedRange().format.autofitColumns();
            await context.sync();
        } catch (e) {
            console.warn("No se pudo autoajustar el ancho de columnas:", e);
        }
    }

    console.log("jsonTo3Matrices: pintado OK ->", {
        factCeldas: factCells.size, filasCeldas: filasCells.size, columnasCeldas: columnasCells.size,
        maxRowId, maxColId,
        indicadoresFilas: rowsFilter.indicators.length, indicadoresColumnas: colsFilter.indicators.length
    });
}

/* ---------------------------------------------------------------------
 * Formato "Draco_001_*": paleta de color por nivel de jerarquía, rangos
 * con nombre, tipografía y bordes.
 * ------------------------------------------------------------------- */

// RGB(226,232,240) / RGB(241,245,249) / RGB(248,250,252) / blanco (a partir
// de aquí se repite para niveles 4, 5, 6...). Texto RGB(13,23,42) en todos.
const DRACO_LEVEL_PALETTE = [
    { fill: "#E2E8F0", font: "#0D172A" }, // nivel 1
    { fill: "#F1F5F9", font: "#0D172A" }, // nivel 2
    { fill: "#F8FAFC", font: "#0D172A" }, // nivel 3
    { fill: "#FFFFFF", font: "#0D172A" }  // nivel 4 en adelante
];

const DRACO_BORDER_COLOR = "#0D172A"; // RGB(13,23,42)
const DRACO_FONT_NAME = "Segoe UI";
const DRACO_FONT_SIZE = 9;

function dracoColorForLevel(fieldBase1, indent) {
    const idx = Math.min(Math.max(fieldBase1 - 1, 0) + Math.max(indent, 0), DRACO_LEVEL_PALETTE.length - 1);
    return DRACO_LEVEL_PALETTE[idx];
}

/**
 * Aplica indentLevel + color de relleno/texto + fuente a un bloque de
 * celdas ya escrito (filasCells o columnasCells), agrupando en tramos
 * contiguos para minimizar llamadas a la API.
 *
 * axis === "col": el "nivel" viene dado por la COLUMNA (bloque de FILAS,
 *   la jerarquía crece hacia la derecha); se agrupa por columna y se
 *   recorren tramos contiguos de fila con el mismo indent.
 * axis === "row": el "nivel" viene dado por la FILA (bloque de COLUMNAS,
 *   la jerarquía crece hacia abajo); se agrupa por fila y se recorren
 *   tramos contiguos de columna con el mismo indent.
 */
async function writeIndentAndColorRuns(context, sheet, cellsMap, axis, fieldOffset) {
    if (cellsMap.size === 0) return;

    const groupKeyName = axis === "col" ? "col" : "row";
    const runKeyName = axis === "col" ? "row" : "col";

    const groups = new Map();
    for (const c of cellsMap.values()) {
        const g = c[groupKeyName];
        if (!groups.has(g)) groups.set(g, []);
        groups.get(g).push(c);
    }

    for (const [g, list] of groups) {
        list.sort((a, b) => a[runKeyName] - b[runKeyName]);
        const fieldBase1 = g - fieldOffset; // 1-based: posición del campo en el eje

        let runStart = 0;
        for (let k = 1; k <= list.length; k++) {
            const endOfRun = k === list.length
                || list[k][runKeyName] !== list[k - 1][runKeyName] + 1
                || list[k].indent !== list[k - 1].indent;

            if (endOfRun) {
                const first = list[runStart];
                const last = list[k - 1];
                const style = dracoColorForLevel(fieldBase1, first.indent);

                let range;
                if (axis === "col") {
                    const numRows = last.row - first.row + 1;
                    range = sheet.getRangeByIndexes(first.row - 1, g - 1, numRows, 1);
                } else {
                    const numCols = last.col - first.col + 1;
                    range = sheet.getRangeByIndexes(g - 1, first.col - 1, 1, numCols);
                }

                range.format.indentLevel = first.indent;
                range.format.font.name = DRACO_FONT_NAME;
                range.format.font.size = DRACO_FONT_SIZE;
                range.format.fill.color = style.fill;
                range.format.font.color = style.font;

                runStart = k;
            }
        }
    }

    await context.sync();
}

/**
 * Borra (formato + contenido) los rangos de la ejecución anterior a los
 * que apuntaban los nombres Draco_001_Rows/Cols/Values, si existen.
 */
async function clearDracoNamedRanges(context) {
    const names = ["Draco_001_Rows", "Draco_001_Cols", "Draco_001_Values"];
    const items = names.map(n => context.workbook.names.getItemOrNullObject(n));
    items.forEach(it => it.load("isNullObject"));
    await context.sync();

    let anyToClear = false;
    for (const it of items) {
        if (!it.isNullObject) {
            it.getRange().clear(Excel.ClearApplyTo.all);
            anyToClear = true;
        }
    }
    if (anyToClear) {
        await context.sync();
    }
}

/**
 * Crea/actualiza los 3 rangos con nombre a partir de la última tabla
 * pintada, y aplica: fuente Segoe UI 9 en los tres, número 2 decimales +
 * separador de miles + centrado en Draco_001_Values, y un borde fino
 * RGB(13,23,42) alrededor de cada uno de los 3 rangos.
 */
async function applyDracoNamedRanges(context, sheet, dims) {
    const { RRows, RCols, totalDimFilas, totalDimCols, maxRowId, maxColId, applyVisualFormat } = dims;
    const doFormat = applyVisualFormat !== false; // por defecto, sí formatear (comportamiento previo)

    if (maxRowId <= 0 || maxColId <= 0 || totalDimFilas <= 0 || totalDimCols <= 0) {
        // No hay datos suficientes para definir una tabla: no se crean rangos.
        return;
    }

    const rowsRange = sheet.getRangeByIndexes(RRows.row - 1, RRows.col - 1, maxRowId, totalDimFilas);
    const colsRange = sheet.getRangeByIndexes(RCols.row - 1, RCols.col - 1, totalDimCols, maxColId);
    const valuesRange = sheet.getRangeByIndexes(RRows.row - 1, RCols.col - 1, maxRowId, maxColId);

    // "Sobrescribir formatos" (propiedades del informe) desactivado: se
    // conservan el color/fuente/bordes que el usuario haya tocado a mano,
    // y solo se (re)definen los 3 rangos con nombre para que apunten al
    // tamaño actual de la tabla.
    if (doFormat) {
        // ---- Fuente Segoe UI 9 en los tres rangos ----
        for (const r of [rowsRange, colsRange, valuesRange]) {
            r.format.font.name = DRACO_FONT_NAME;
            r.format.font.size = DRACO_FONT_SIZE;
        }

        // ---- Valores: 2 decimales + separador de miles, centrado ----
        valuesRange.numberFormat = [["#,##0.00"]];
        valuesRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

        // ---- Bordes: primero se elimina cualquier borde existente del rango
        //      (externo + interior, restos de refrescos anteriores con otra
        //      forma/tamaño de tabla) y SOLO DESPUÉS se pinta el borde exterior
        //      fino nuevo. Hacerlo en el mismo lote sin sync intermedio hace
        //      que Excel no aplique bien el cambio, así que se separan en dos
        //      pasadas con su propio context.sync().
        const ALL_BORDER_EDGES = [
            Excel.BorderIndex.edgeTop, Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft, Excel.BorderIndex.edgeRight,
            Excel.BorderIndex.insideHorizontal, Excel.BorderIndex.insideVertical
        ];
        const OUTER_BORDER_EDGES = [
            Excel.BorderIndex.edgeTop, Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft, Excel.BorderIndex.edgeRight
        ];

        // Pasada 1: quitar el borde del rango por completo.
        for (const r of [rowsRange, colsRange, valuesRange]) {
            for (const edge of ALL_BORDER_EDGES) {
                r.format.borders.getItem(edge).style = Excel.BorderLineStyle.none;
            }
        }
        await context.sync();

        // Pasada 2: pintar el borde exterior fino, color RGB(13,23,42).
        for (const r of [rowsRange, colsRange, valuesRange]) {
            for (const edge of OUTER_BORDER_EDGES) {
                const border = r.format.borders.getItem(edge);
                border.style = Excel.BorderLineStyle.continuous;
                border.weight = Excel.BorderWeight.thin;
                border.color = DRACO_BORDER_COLOR;
            }
        }
        await context.sync();
    }

    // ---- (Re)definir los nombres apuntando a los rangos recién pintados ----
    const defs = [
        { name: "Draco_001_Rows", range: rowsRange },
        { name: "Draco_001_Cols", range: colsRange },
        { name: "Draco_001_Values", range: valuesRange }
    ];

    for (const d of defs) {
        const existing = context.workbook.names.getItemOrNullObject(d.name);
        existing.load("isNullObject");
        await context.sync();
        if (!existing.isNullObject) {
            existing.delete();
            await context.sync();
        }
        context.workbook.names.add(d.name, d.range);
    }

    await context.sync();
}

/**
 * Escribe un conjunto de celdas {row,col,value} en el MÍNIMO número de
 * llamadas a la API de Excel posible: agrupa por bounding box, LEE el
 * contenido actual de ese rectángulo (para no pisar nada fuera de las
 * celdas concretas que tocan — mismo comportamiento que escribir celda a
 * celda sin ClearContents previo) y hace una única escritura de vuelta.
 */
async function writeCellBlock(context, sheet, cellsMap) {
    if (cellsMap.size === 0) return;

    let minRow = Infinity, maxRow = -Infinity, minCol = Infinity, maxCol = -Infinity;
    for (const c of cellsMap.values()) {
        if (c.row < minRow) minRow = c.row;
        if (c.row > maxRow) maxRow = c.row;
        if (c.col < minCol) minCol = c.col;
        if (c.col > maxCol) maxCol = c.col;
    }

    const numRows = maxRow - minRow + 1;
    const numCols = maxCol - minCol + 1;

    // Bounding box demasiado disperso (pocas celdas reales en un área enorme):
    // más rápido ir celda a celda que leer/escribir un rectángulo gigante.
    if (numRows * numCols > 50000 && numRows * numCols > cellsMap.size * 50) {
        for (const c of cellsMap.values()) {
            sheet.getRangeByIndexes(c.row - 1, c.col - 1, 1, 1).values = [[c.value]];
        }
        await context.sync();
        return;
    }

    const range = sheet.getRangeByIndexes(minRow - 1, minCol - 1, numRows, numCols);
    range.load("values");
    await context.sync();

    const grid = range.values.map(r => r.slice());
    for (const c of cellsMap.values()) {
        grid[c.row - minRow][c.col - minCol] = c.value;
    }

    range.values = grid;
    await context.sync();
}

/**
 * Aplica el indentLevel agrupando en tramos contiguos (misma columna, filas
 * consecutivas, mismo indentLevel) para minimizar llamadas a la API,
 * respetando exactamente los valores por celda calculados antes.
 */
async function writeIndentRuns(context, sheet, cellsMap) {
    if (cellsMap.size === 0) return;

    const byCol = new Map();
    for (const c of cellsMap.values()) {
        if (!byCol.has(c.col)) byCol.set(c.col, []);
        byCol.get(c.col).push(c);
    }

    for (const [col, list] of byCol) {
        list.sort((a, b) => a.row - b.row);

        let runStart = 0;
        for (let k = 1; k <= list.length; k++) {
            const endOfRun = k === list.length
                || list[k].row !== list[k - 1].row + 1
                || list[k].indent !== list[k - 1].indent;

            if (endOfRun) {
                const first = list[runStart];
                const last = list[k - 1];
                const numRows = last.row - first.row + 1;
                sheet.getRangeByIndexes(first.row - 1, col - 1, numRows, 1).format.indentLevel = first.indent;
                runStart = k;
            }
        }
    }

    await context.sync();
}

/**
 * Núcleo de Actualizar_informe(), sin manejo de `event`.
 */
async function actualizarInformeCore() {
    let sql;

    await Excel.run(async (context) => {
        const editReportGrid = await getValuesGrid(context, "EDIT_REPORT");
        const relGrid = await getValuesGrid(context, "MODEL_RELATIONSHIP");
        const measuresGrid = await getValuesGrid(context, "MODEL_MEASURES");
        const atributesGrid = await getValuesGrid(context, "MODEL_ATRIBUTES");

        loadReportDefinition(editReportGrid);

        sql = buildSQL(relGrid, measuresGrid, atributesGrid);

        console.log("BuildSQL ->", sql);

        const csvSheet = context.workbook.worksheets.getItem("CSV_RESULT");
        csvSheet.getRange("A1").values = [[sql]];

        await context.sync();
    });

    const json = await executeSQL(sql);

    console.log("JSON de BigQuery ->", json);

    await Excel.run(async (context) => {
        const csvSheet = context.workbook.worksheets.getItem("CSV_RESULT");
        const EXCEL_CELL_CHAR_LIMIT = 32000; // límite real de Excel: 32767
        const jsonForCell = json.length > EXCEL_CELL_CHAR_LIMIT
            ? json.substring(0, EXCEL_CELL_CHAR_LIMIT) + " ...(truncado, JSON completo en la consola F12)"
            : json;
        csvSheet.getRange("B1").values = [[jsonForCell]];
        await context.sync();
    });

    await Excel.run(async (context) => {
        await jsonTo3Matrices(context, json);
    });
}

/**
 * Traducción de Actualizar_informe(): LoadReportDefinition, BuildSQL,
 * ExecuteSQL, JSON_To_3_Matrices.
 * @param {Office.AddinCommands.Event} event
 */
async function actualizarInforme(event) {
    try {
        await actualizarInformeCore();
    } catch (error) {
        console.error("Error al actualizar el informe (dinámico):", error);
    } finally {
        if (event) {
            event.completed();
        }
    }
}

/**
 * Traducción de Actualizar(): dispatcher según EDIT_REPORT!H12/N12.
 * @param {Office.AddinCommands.Event} event
 */
async function actualizar(event) {
    try {
        const isFixed = await Excel.run(async (context) => {
            const grid = await getValuesGrid(context, "EDIT_REPORT");
            const h12 = String(cellValue(grid, 12, 8)).trim().toUpperCase();
            const n12 = String(cellValue(grid, 12, 14)).trim().toUpperCase();
            return h12 === "X" && n12 === "X";
        });

        if (isFixed) {
            await actualizarInformeFixedCore();
        } else {
            await actualizarInformeCore();
        }
    } catch (error) {
        console.error("Error al actualizar el informe:", error);
    } finally {
        if (event) {
            event.completed();
        }
    }
}

// Exponer las funciones de actualización para poder llamarlas también desde
// el taskpane (p.ej. tras guardar el diseño), no solo desde el ribbon.
window.ReportActions = {
    actualizar, actualizarInforme, actualizarInformeFixed,
    toggleRefreshPaused, toggleMemberRecognition, openReportProperties, openFieldOptions,
    convertAxisStaticFormulas, ensureDracoHandlersRegistered
};



// Asociar el nombre del comando del manifiesto con la función JavaScript

/**
 * Función placeholder para botones del ribbon aún no implementados.
 * No hace nada salvo completar el evento, para que el botón no falle.
 */
function comingSoon(event) {
    console.log("Esta función todavía no está implementada.");
    if (event) {
        event.completed();
    }
}

/* ---------------------------------------------------------------------
 * Botones del ribbon "tipo pulsador": Pausar refresco / Reconocimiento
 * de miembros. Por ahora SOLO guardan su estado (Office roaming
 * settings, visible desde cualquier contexto: ribbon y taskpane); no hay
 * ninguna lógica de negocio todavía detrás de ellos.
 * ------------------------------------------------------------------- */
function toggleDracoSetting(key) {
    const settings = Office.context.document.settings;
    const current = !!settings.get(key);
    settings.set(key, !current);
    return new Promise((resolve) => settings.saveAsync(() => resolve(!current)));
}

async function requestRibbonLabelUpdate(controlId, label) {
    try {
        if (Office.ribbon && Office.ribbon.requestUpdate) {
            await Office.ribbon.requestUpdate({
                tabs: [{
                    id: "DracoBITab",
                    groups: [{ id: "GroupOpcionesInforme", controls: [{ id: controlId, label }] }]
                }]
            });
        }
    } catch (e) {
        console.warn("No se pudo actualizar la etiqueta del ribbon (" + controlId + "):", e);
    }
}

async function toggleRefreshPaused(event) {
    try {
        const nowOn = await toggleDracoSetting("draco_refreshPaused");
        console.log("Pausar refresco:", nowOn ? "activado" : "desactivado");
        await requestRibbonLabelUpdate("BtnPausarRefresco", nowOn ? "Refresco: Pausado" : "Pausar refresco");
    } catch (error) {
        console.error("Error al alternar 'Pausar refresco':", error);
    } finally {
        if (event) event.completed();
    }
}

async function toggleMemberRecognition(event) {
    try {
        const nowOn = await toggleDracoSetting("draco_memberRecognition");
        console.log("Reconocimiento de miembros:", nowOn ? "activado" : "desactivado");
        await requestRibbonLabelUpdate("BtnReconocimientoMiembros", nowOn ? "Miembros: activado" : "Reconoc. de miembros");
    } catch (error) {
        console.error("Error al alternar 'Reconocimiento de miembros':", error);
    } finally {
        if (event) event.completed();
    }
}

/**
 * Botones del ribbon "Propiedades del informe" y "Opciones de campo":
 * abren/traen al frente el taskpane y le dejan marcada una acción
 * pendiente en Office roaming settings; taskpane.js la recoge al arrancar
 * (o al detectar el cambio de settings) y abre el modal/panel correspondiente.
 */
/**
 * Botones del ribbon "Propiedades del informe" y "Opciones de campo".
 *
 * Desde que el complemento usa Shared Runtime (ver <Runtimes> en el
 * manifiesto, resid="Taskpane.Url"), estas funciones se ejecutan en el
 * MISMO contexto JS que taskpane.js (que además sigue vivo aunque el
 * panel esté oculto, por "lifetime=long"). Eso permite:
 *   1) Llamar directamente a los métodos de TaskPaneApp (sin pasar por
 *      Office roaming settings ni esperar a que taskpane.js relea nada).
 *   2) Usar Office.addin.showAsTaskpane(), que EXIGE Shared Runtime; sin
 *      él lanza "RichApi.Error: La API solo se aplica al complemento que
 *      usa Shared Runtime" (el error que se veía antes de este cambio).
 *
 * Se conserva el mecanismo antiguo (Office roaming settings +
 * SettingsChanged en taskpane.js) como red de seguridad, por si en algún
 * host el runtime compartido tardase en levantar TaskPaneApp.
 */
async function openReportProperties(event) {
    try {
        console.log("[Draco] openReportProperties: shared runtime ¿activo? ->", typeof TaskPaneApp !== "undefined");
        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.openReportPropertiesModal) {
            TaskPaneApp.openReportPropertiesModal();
            console.log("[Draco] openReportPropertiesModal() ejecutado directamente (shared runtime OK).");
        } else {
            console.warn("[Draco] TaskPaneApp NO existe en este contexto: usando fallback de settings (¿manifest sin Runtimes recargado?).");
            const settings = Office.context.document.settings;
            settings.set("draco_pendingAction", "properties");
            await new Promise((resolve) => settings.saveAsync(resolve));
        }
        if (Office.addin && Office.addin.showAsTaskpane) {
            await Office.addin.showAsTaskpane();
            console.log("[Draco] showAsTaskpane() resuelto sin error.");
        } else {
            console.warn("[Draco] Office.addin.showAsTaskpane no está disponible en este runtime.");
        }
    } catch (error) {
        console.error("Error al abrir Propiedades del informe:", error);
    } finally {
        if (event) event.completed();
    }
}

async function openFieldOptions(event) {
    try {
        console.log("[Draco] openFieldOptions: shared runtime ¿activo? ->", typeof TaskPaneApp !== "undefined");
        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.setFieldOptionsPanelOpen) {
            TaskPaneApp.setFieldOptionsPanelOpen(true);
            console.log("[Draco] setFieldOptionsPanelOpen(true) ejecutado directamente (shared runtime OK).");
        } else {
            console.warn("[Draco] TaskPaneApp NO existe en este contexto: usando fallback de settings (¿manifest sin Runtimes recargado?).");
            const settings = Office.context.document.settings;
            settings.set("draco_pendingAction", "fieldOptions");
            await new Promise((resolve) => settings.saveAsync(resolve));
        }
        if (Office.addin && Office.addin.showAsTaskpane) {
            await Office.addin.showAsTaskpane();
            console.log("[Draco] showAsTaskpane() resuelto sin error.");
        } else {
            console.warn("[Draco] Office.addin.showAsTaskpane no está disponible en este runtime.");
        }
    } catch (error) {
        console.error("Error al abrir Opciones de campo:", error);
    } finally {
        if (event) event.completed();
    }
}

// Nota: este fichero se carga tanto en el runtime de comandos (commands.html)
// como, ahora, dentro del propio taskpane (taskpane.html), para poder
// disparar Actualizar()/ActualizarInforme() (y por tanto jsonTo3Matrices)
// automáticamente al guardar cambios en el diseñador. Office.actions.associate
// solo tiene efecto real en el runtime de comandos; se protege con try/catch
// por si el host no expone esa API fuera de ese contexto.
try {
    Office.actions.associate("hidePane", hidePane);
    Office.actions.associate("writeHolaInA1", writeHolaInA1);
    Office.actions.associate("actualizarInformeFixed", actualizarInformeFixed);
    Office.actions.associate("actualizarInforme", actualizarInforme);
    Office.actions.associate("actualizar", actualizar);
    Office.actions.associate("comingSoon", comingSoon);
    Office.actions.associate("toggleRefreshPaused", toggleRefreshPaused);
    Office.actions.associate("toggleMemberRecognition", toggleMemberRecognition);
    Office.actions.associate("openReportProperties", openReportProperties);
    Office.actions.associate("openFieldOptions", openFieldOptions);
} catch (e) {
    console.warn("Office.actions.associate no disponible en este contexto:", e);
}
