/**
 * Lógica para comandos ejecutados en segundo plano por Excel
 */
Office.onReady(() => {
    // Handshake completado para comandos
});

// ==========================================
// VARIABLES GLOBALES DE SESIÓN
// ==========================================
var RowCount = 0;
var ColumnCount = 0;
var MeasureCount = 0;
var AttrColumnCount = 1;
var AttrRowCount = 1;

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

    if (event) {
        event.completed();
    }
}

/**
 * Función principal activada por el botón 'Actualizar' (RefreshReportButton)
 * Traduce la ejecución VBA completa de generación e inserción
 * @param {Office.AddinCommands.Event} event
 */
async function actualizarInformeFixed(event) {
    try {
        var SQL = "";
        var Json = "";

        await LoadReportDefinition();

        SQL = await BuildSQL_Fixed();

        console.log(SQL);

        Json = await ExecuteSQL(SQL);
        console.log(Json);

        await JSON_PaintValues(Json);

    } catch (error) {
        console.error("Error en la actualización del informe:", error);
    } finally {
        if (event) {
            event.completed();
        }
    }
}

// ==========================================
// CONSTRUCCIÓN DE SQL FIXED
// ==========================================
async function BuildSQL_Fixed() {
    var SQL = "";

    SQL = "";

    SQL = SQL + "WITH BASE AS (" + "\r\n\r\n";

    SQL = SQL + await BuildSelectBase() + "\r\n\r\n";

    SQL = SQL + await BuildFrom() + "\r\n\r\n";

    SQL = SQL + await BuildJoins() + "\r\n\r\n";

    SQL = SQL + await BuildBaseWhere(await ReadRowDefinitions(), await ReadColumnDefinitions()) + "\r\n\r\n";

    SQL = SQL + await BuildGroupByBase() + "\r\n\r\n";

    SQL = SQL + ")," + "\r\n\r\n";

    SQL = SQL + await BuildBaseRow(await ReadRowDefinitions()) + "\r\n\r\n";

    SQL = SQL + await BuildColumns(await ReadColumnDefinitions());

    return SQL;
}

async function BuildSelectBase() {
    var Rows = await ReadRowDefinitions();
    var Cols = await ReadColumnDefinitions();
    var Dict = {};
    var SQL = "";

    for (var i = 0; i < Rows.length; i++) {
        var V = Rows[i];
        var key = String(V[1] + "|" + V[2]).toUpperCase();
        if (!Dict[key]) {
            Dict[key] = (await GetTableAlias(String(V[1]))) + "." + String(V[2]);
        }
    }

    for (var j = 0; j < Cols.length; j++) {
        var V = Cols[j];
        var key = String(V[1] + "|" + V[2]).toUpperCase();
        if (!Dict[key]) {
            Dict[key] = (await GetTableAlias(String(V[1]))) + "." + String(V[2]);
        }
    }

    SQL = "SELECT" + "\r\n";

    for (var key in Dict) {
        var S = Dict[key];
        SQL = SQL + "    " + S + "," + "\r\n";
    }

    SQL = SQL + "    SUM(f.IMPORTE) AS IMPORTE";

    return SQL;
}

async function BuildFrom() {
    var SQL = "";

    SQL = "FROM `FACT_TABLE` f";

    return SQL;
}

async function BuildJoins() {
    return await Excel.run(async (context) => {
        var SQL = "";
        var ws = context.workbook.worksheets.getItem("MODEL_RELATIONSHIP");
        var usedRange = ws.getUsedRange();
        usedRange.load(["values", "rowCount"]);
        await context.sync();

        var values = usedRange.values;
        var lastRow = values.length;

        for (var R = 1; R < lastRow; R++) {
            var id = values[R][0];
            var dimName = values[R][1];
            var joinCondition = values[R][2];

            if (dimName) {
                SQL = SQL + "LEFT JOIN `" + dimName + "` d" + id + "\r\n";
                SQL = SQL + "   ON " + joinCondition + "\r\n";
            }
        }

        return SQL;
    });
}

async function BuildGroupByBase() {
    var Rows = await ReadRowDefinitions();
    var Cols = await ReadColumnDefinitions();
    var Dict = {};
    var SQL = "";

    for (var i = 0; i < Rows.length; i++) {
        var V = Rows[i];
        var key = String(V[1] + "|" + V[2]).toUpperCase();
        if (!Dict[key]) {
            Dict[key] = (await GetTableAlias(String(V[1]))) + "." + String(V[2]);
        }
    }

    for (var j = 0; j < Cols.length; j++) {
        var V = Cols[j];
        var key = String(V[1] + "|" + V[2]).toUpperCase();
        if (!Dict[key]) {
            Dict[key] = (await GetTableAlias(String(V[1]))) + "." + String(V[2]);
        }
    }

    SQL = "GROUP BY" + "\r\n";

    for (var key in Dict) {
        var S = Dict[key];
        SQL = SQL + "    " + S + "," + "\r\n";
    }

    SQL = SQL.substring(0, SQL.length - 3);

    return SQL;
}

async function BuildBaseWhere(Rows, Columns) {
    var SQL = "";
    var Dict = {};

    for (var i = 0; i < Columns.length; i++) {
        var V = Columns[i];
        var Campo = (await GetTableAlias(String(V[1]))) + "." + String(V[2]);
        var key = (Campo + "|" + String(V[3])).toUpperCase();
        if (!Dict[key]) {
            Dict[key] = Campo + "=" + (await SQLValue(String(V[1]), String(V[2]), String(V[3])));
        }
    }

    for (var j = 0; j < Rows.length; j++) {
        var V = Rows[j];
        var Campo = (await GetTableAlias(String(V[1]))) + "." + String(V[2]);
        var key = (Campo + "|" + String(V[3])).toUpperCase();
        if (!Dict[key]) {
            Dict[key] = Campo + "=" + (await SQLValue(String(V[1]), String(V[2]), String(V[3])));
        }
    }

    SQL = "WHERE" + "\r\n" + "(" + "\r\n";

    var First = true;

    for (var key in Dict) {
        var V = Dict[key];
        if (First) {
            SQL = SQL + "      " + V;
            First = false;
        } else {
            SQL = SQL + "\r\n" + "   OR " + V;
        }
    }

    SQL = SQL + "\r\n" + ")";

    return SQL;
}

async function BuildBaseRow(Rows) {
    var SQL = "";
    var Dict = {};

    for (var i = 0; i < Rows.length; i++) {
        var V = Rows[i];
        var rowIdKey = String(V[0]);
        var cond = "    " + V[2] + "=" + (await SQLValue(String(V[1]), String(V[2]), String(V[3])));
        if (!Dict[rowIdKey]) {
            Dict[rowIdKey] = cond;
        } else {
            Dict[rowIdKey] = Dict[rowIdKey] + "\r\n" + "AND " + V[2] + "=" + (await SQLValue(String(V[1]), String(V[2]), String(V[3])));
        }
    }

    SQL = "BASE_ROW AS (" + "\r\n\r\n";

    var First = true;

    for (var RowID in Dict) {
        if (!First) {
            SQL = SQL + "\r\n" + "UNION ALL" + "\r\n\r\n";
        }

        SQL = SQL +
            "SELECT" + "\r\n" +
            "    " + RowID + " AS ROW_ID," + "\r\n" +
            "    *" + "\r\n" +
            "FROM BASE" + "\r\n" +
            "WHERE" + "\r\n" +
            Dict[RowID];

        First = false;
    }

    SQL = SQL + "\r\n\r\n" + ")";

    return SQL;
}

async function BuildColumns(Columns) {
    var SQL = "";
    var Dict = {};

    for (var i = 0; i < Columns.length; i++) {
        var V = Columns[i];
        var colIdKey = String(V[0]);
        var cond = "    " + V[2] + "=" + (await SQLValue(String(V[1]), String(V[2]), String(V[3])));
        if (!Dict[colIdKey]) {
            Dict[colIdKey] = cond;
        } else {
            Dict[colIdKey] = Dict[colIdKey] + "\r\n" + "AND " + V[2] + "=" + (await SQLValue(String(V[1]), String(V[2]), String(V[3])));
        }
    }

    var First = true;

    for (var ColID in Dict) {
        if (!First) {
            SQL = SQL + "\r\n" + "UNION ALL" + "\r\n\r\n";
        }

        SQL = SQL +
            "SELECT" + "\r\n" +
            "    ROW_ID," + "\r\n" +
            "    " + ColID + " AS COLUMN_ID," + "\r\n" +
            "    SUM(IMPORTE) AS IMPORTE" + "\r\n" +
            "FROM BASE_ROW" + "\r\n" +
            "WHERE" + "\r\n" +
            Dict[ColID] + "\r\n" +
            "GROUP BY ROW_ID";

        First = false;
    }

    SQL = SQL + "\r\n\r\n" +
        "ORDER BY" + "\r\n" +
        "    ROW_ID," + "\r\n" +
        "    COLUMN_ID";

    return SQL;
}

// ==========================================
// EJECUCIÓN Y ESCAPE JSON
// ==========================================
async function ExecuteSQL(SQL) {
    return await Excel.run(async (context) => {
        var sheetDef = context.workbook.worksheets.getItem("M.S. - Definición");
        var sheetConfig = context.workbook.worksheets.getItem("CONFIG");

        var rangeProjectId = sheetDef.getRange("C11");
        var rangeToken = sheetConfig.getRange("B3");

        rangeProjectId.load("values");
        rangeToken.load("values");

        await context.sync();

        var ProjectId = rangeProjectId.values[0][0];
        var Token = rangeToken.values[0][0];

        if (!Token) {
            Token = localStorage.getItem("bigquery_access_token");
        }

        var url = "https://bigquery.googleapis.com/bigquery/v2/projects/" + ProjectId + "/queries";

        var body = `{"query":"${EscapeJSON(SQL)}","useLegacySql":false}`;

        var response = await fetch(url, {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + Token,
                "Content-Type": "application/json"
            },
            body: body
        });

        return await response.text();
    });
}

function EscapeJSON(Text) {
    Text = Text.replace(/\\/g, "\\\\");
    Text = Text.replace(/"/g, '\\"');
    Text = Text.replace(/\r\n/g, "\\n");
    Text = Text.replace(/\r/g, "\\n");
    Text = Text.replace(/\n/g, "\\n");
    return Text;
}

// ==========================================
// PINTADO DE VALORES
// ==========================================
async function JSON_PaintValues(Json) {
    await Excel.run(async (context) => {
        var ws = context.workbook.worksheets.getItem("CSV_RESULT");

        var pos = 0;
        var Campo = 0;
        var RowID = 0;
        var ColID = 0;
        var Texto = "";

        while (true) {
            pos = Json.indexOf('"v":', pos);
            if (pos === -1) break;

            var Ini = pos + 4;

            while (Json.charAt(Ini) === ":" || Json.charAt(Ini) === " ") {
                Ini++;
            }

            var Fin = 0;
            if (Json.charAt(Ini) === '"') {
                Ini++;
                Fin = Json.indexOf('"', Ini);
            } else {
                Fin = Json.indexOf('}', Ini);
            }

            Texto = Json.substring(Ini, Fin);
            Campo++;

            switch (Campo) {
                case 1:
                    RowID = parseInt(Texto, 10);
                    break;
                case 2:
                    ColID = parseInt(Texto, 10);
                    break;
                case 3:
                    var cell = ws.getCell(RowID - 1, ColID - 1);
                    cell.values = [[Texto]];
                    Campo = 0;
                    break;
            }

            pos = Fin + 1;
        }

        await context.sync();
    });
}

// ==========================================
// AUXILIARES Y LECTURA
// ==========================================
async function LoadReportDefinition() {
    await Excel.run(async (context) => {
        var sheetConfig = context.workbook.worksheets.getItem("CONFIG");

        var rangeAttrCol = sheetConfig.getRange("B5");
        var rangeAttrRow = sheetConfig.getRange("B6");

        rangeAttrCol.load("values");
        rangeAttrRow.load("values");

        await context.sync();

        if (rangeAttrCol.values[0][0]) {
            AttrColumnCount = Number(rangeAttrCol.values[0][0]);
        }
        if (rangeAttrRow.values[0][0]) {
            AttrRowCount = Number(rangeAttrRow.values[0][0]);
        }
    });
}

async function GetTableAlias(Dimension) {
    var fila = await BuscarFilaDimension(Dimension);

    if (fila === 0) {
        return "f";
    } else {
        return "d" + fila;
    }
}

async function BuscarFilaDimension(Dimension) {
    return await Excel.run(async (context) => {
        var ws = context.workbook.worksheets.getItem("MODEL_RELATIONSHIP");
        var usedRange = ws.getUsedRange();
        usedRange.load(["values", "rowCount"]);
        await context.sync();

        var values = usedRange.values;
        var lastRow = values.length;

        for (var R = 1; R < lastRow; R++) {
            if (String(values[R][1]).trim().toUpperCase() === Dimension.trim().toUpperCase()) {
                return Number(values[R][0]);
            }
        }

        return 0;
    });
}

async function ReadRowDefinitions() {
    return await Excel.run(async (context) => {
        var C = [];
        var wsEdit = context.workbook.worksheets.getItem("EDIT_REPORT");
        var wsCsv = context.workbook.worksheets.getItem("CSV_RESULT");

        var rRowsRef = wsEdit.getRange("H10");
        rRowsRef.load("values");
        await context.sync();

        var RRows = wsCsv.getRange(rRowsRef.values[0][0]);
        RRows.load(["rowIndex", "columnIndex"]);
        await context.sync();

        var usedRange = wsCsv.getUsedRange();
        usedRange.load(["rowCount", "formulas"]);
        await context.sync();

        var startRow = RRows.rowIndex;
        var startCol = RRows.columnIndex;
        var lastRow = usedRange.rowCount;

        var lastCol = startCol + AttrColumnCount - 1;

        for (var R = startRow; R < lastRow; R++) {
            for (var Col = startCol; Col <= lastCol; Col++) {
                var F = String(usedRange.formulas[R][Col]);

                if (F && F.startsWith("=")) {
                    F = F.replace("=@", "");
                    F = F.replace("=EPM_VALUE(", "");
                    if (F.endsWith(")")) F = F.substring(0, F.length - 1);

                    var V = F.includes(";") ? F.split(";") : F.split(",");

                    var Item = [
                        R + 1,
                        await GetFormulaArgumentValue(wsCsv, V[0]),
                        await GetFormulaArgumentValue(wsCsv, V[1]),
                        await GetFormulaArgumentValue(wsCsv, V[2]),
                        await GetFormulaArgumentValue(wsCsv, V[3])
                    ];

                    C.push(Item);
                } else {
                    break;
                }
            }
        }

        return C;
    });
}

async function ReadColumnDefinitions() {
    return await Excel.run(async (context) => {
        var C = [];
        var wsEdit = context.workbook.worksheets.getItem("EDIT_REPORT");
        var wsCsv = context.workbook.worksheets.getItem("CSV_RESULT");

        var rColsRef = wsEdit.getRange("N10");
        rColsRef.load("values");
        await context.sync();

        var RCols = wsCsv.getRange(rColsRef.values[0][0]);
        RCols.load(["rowIndex", "columnIndex"]);
        await context.sync();

        var usedRange = wsCsv.getUsedRange();
        usedRange.load(["columnCount", "formulas"]);
        await context.sync();

        var startRow = RCols.rowIndex;
        var startCol = RCols.columnIndex;
        var lastCol = usedRange.columnCount;

        var lastRow = startRow + AttrRowCount - 1;

        for (var Col = startCol; Col < lastCol; Col++) {
            for (var R = startRow; R <= lastRow; R++) {
                var F = String(usedRange.formulas[R][Col]);

                if (F && F.startsWith("=")) {
                    F = F.replace("=@", "");
                    F = F.replace("=EPM_VALUE(", "");
                    if (F.endsWith(")")) F = F.substring(0, F.length - 1);

                    var V = F.includes(";") ? F.split(";") : F.split(",");

                    var Item = [
                        Col + 1,
                        await GetFormulaArgumentValue(wsCsv, V[0]),
                        await GetFormulaArgumentValue(wsCsv, V[1]),
                        await GetFormulaArgumentValue(wsCsv, V[2]),
                        await GetFormulaArgumentValue(wsCsv, V[3])
                    ];

                    C.push(Item);
                } else {
                    break;
                }
            }
        }

        return C;
    });
}

async function GetFormulaArgumentValue(ws, Arg) {
    Arg = Arg.trim();

    if (Arg.startsWith('"') && Arg.endsWith('"')) {
        return Arg.substring(1, Arg.length - 1);
    }

    return Arg.replace(/"/g, "");
}

async function GetAttributeType(Dimension, Atributo) {
    return await Excel.run(async (context) => {
        var ws = context.workbook.worksheets.getItem("MODEL_ATTRIBUTES");
        var usedRange = ws.getUsedRange();
        usedRange.load(["values", "rowCount"]);
        await context.sync();

        var values = usedRange.values;
        var lastRow = values.length;

        for (var R = 1; R < lastRow; R++) {
            if (
                String(values[R][0]).trim().toUpperCase() === Dimension.trim().toUpperCase() &&
                String(values[R][1]).trim().toUpperCase() === Atributo.trim().toUpperCase()
            ) {
                return String(values[R][2]).trim().toUpperCase();
            }
        }

        return "STRING";
    });
}

async function GetHierarchyMaxLevel(DimName, IsRows) {
    return await Excel.run(async (context) => {
        var ws = context.workbook.worksheets.getItem("MODEL_HIERARCHIES");
        var usedRange = ws.getUsedRange();
        usedRange.load(["values", "rowCount"]);
        await context.sync();

        var values = usedRange.values;
        var lastRow = values.length;

        for (var R = 1; R < lastRow; R++) {
            if (String(values[R][0]).trim().toUpperCase() === DimName.trim().toUpperCase()) {
                return Number(values[R][1]);
            }
        }

        return 1;
    });
}

async function SQLValue(Dimension, Atributo, valor) {
    var attrType = await GetAttributeType(Dimension, Atributo);
    if (attrType === "INTEGER") {
        return valor;
    } else {
        return "'" + valor.replace(/'/g, "''") + "'";
    }
}

/**
 * Función legacy de ejemplo
 * @param {Office.AddinCommands.Event} event
 */
async function writeHolaInA1(event) {
    return actualizarInformeFixed(event);
}

// Asociar las acciones del manifiesto con las funciones JavaScript
Office.actions.associate("hidePane", hidePane);
Office.actions.associate("writeHolaInA1", writeHolaInA1);
Office.actions.associate("actualizarInformeFixed", actualizarInformeFixed);