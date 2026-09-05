/**
 * Lógica para comandos ejecutados en segundo plano por Excel
 */
Office.onReady(() => {
    // Handshake completado para comandos

    // Clic en celdas de Draco_XXX_Rows (cualquier hoja) -> escribe el
    // resultado (zona + jerarquía SI/NO) en A51.
    // Este mismo fichero (commands.js) se carga tanto en el runtime de
    // comandos del ribbon (commands.html) como en el taskpane
    // (taskpane.html lo incluye antes que taskpane.js), así que este
    // Office.onReady se dispara con lo que ocurra ANTES: que el usuario
    // abra el panel de tareas o que pulse cualquier botón del ribbon
    // (p.ej. "Actualizar"). No depende de ningún flujo del taskpane.
    ensureDracoRowsClickLoggerRegistered().catch(e => {
        console.warn("[Draco] No se pudo registrar el listener de Draco_*_Rows desde Office.onReady:", e);
    });
});

/**
 * Botón de ribbon "Abrir modelo semántico" (ModeloAbrirButton).
 * Abre directamente el diálogo independiente de importación LookML
 * (Office.context.ui.displayDialogAsync), sin depender de que el taskpane
 * del modelo semántico esté abierto ni de ningún popup dentro de él.
 * @param {Office.AddinCommands.Event} event
 */
function abrirModeloSemantico(event) {
    try {
        // LkmlOpenBridge (js/lkmlOpenBridge.js) abre openSemanticModel.html
        // y, cuando el usuario elige un fichero y confirma, guarda el
        // modelo importado en SemanticModelStore desde este mismo runtime
        // de comandos (que sí tiene acceso a Office.context.document.settings).
        window.LkmlOpenBridge.openOpenLkmlDialog();
    } catch (error) {
        console.error("Error al abrir el diálogo de apertura de modelo semántico:", error);
    } finally {
        if (event) event.completed();
    }
}

/**
 * Botón de ribbon "Guardar modelo semántico" (ModeloGuardarButton).
 * Abre directamente el diálogo independiente de exportación a LookML; la
 * escritura real en EDIT_REPORT!G1 la hace LkmlSaveBridge (ver
 * js/lkmlSaveBridge.js), que sí tiene acceso a Excel.run desde este mismo
 * runtime de comandos.
 * @param {Office.AddinCommands.Event} event
 */
function guardarModeloSemantico(event) {
    try {
        const models = window.SemanticModelStore.getAllModels();
        const active = window.SemanticModelStore.getActiveModelName();
        window.LkmlSaveBridge.openSaveLkmlDialog(models, active);
    } catch (error) {
        console.error("Error al abrir el diálogo de guardado de modelo semántico:", error);
    } finally {
        if (event) event.completed();
    }
}

/**
 * Botón de ribbon "Abrir bucket" (AbrirBucketButton).
 * Abre un diálogo que lista los .xlsx/.xlsm del bucket de Google Cloud
 * Storage configurado en la conexión BigQuery (bucketBrowser.html) y
 * permite descargar el elegido; el navegador se encarga de abrirlo o
 * guardarlo (no reemplaza el libro activo).
 * @param {Office.AddinCommands.Event} event
 */

/**
 * Botón de ribbon "Abrir bucket" (AbrirBucketButton).
 * Versión limpia (sin diagnóstico) — pega esto sustituyendo tanto la
 * versión original como la de depuración que te pasé antes.
 */
function abrirDesdeBucket(event) {
    try {
        const url = new URL("bucketBrowser.html", window.location.href);
        const sessionParams = window.BQ ? BQ.getSessionQueryParams() : "";
        if (sessionParams) url.search = sessionParams;
        Office.context.ui.displayDialogAsync(url.href, { height: 55, width: 40, displayInIframe: false }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("No se pudo abrir el diálogo del bucket:", asyncResult.error.message);
                return;
            }
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => dialog.close());
        });
    } catch (error) {
        console.error("Error al abrir el explorador del bucket:", error);
    } finally {
        if (event) event.completed();
    }
}



/**
 * Botón de ribbon "Guardar en bucket" (GuardarBucketButton).
 * Abre primero el diálogo de selección de proyecto/bucket (saveBucket.html,
 * ver js/gcsSaveBridge.js) y, con esa elección, sube el .xlsx activo (tal
 * cual está guardado) al bucket elegido. Como es un botón de ribbon sin
 * taskpane propio, el resultado se muestra en un pequeño diálogo
 * (uploadStatus.html) en vez de un alert() bloqueante.
 * @param {Office.AddinCommands.Event} event
 */
function guardarExcelEnBucket(event) {
    function showResult(ok, msg) {
        try {
            const url = new URL("uploadStatus.html", window.location.href);
            url.searchParams.set("ok", ok ? "1" : "0");
            url.searchParams.set("msg", msg);
            Office.context.ui.displayDialogAsync(url.href, { height: 25, width: 30, displayInIframe: false }, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("No se pudo abrir el diálogo de estado de la subida:", asyncResult.error.message);
                    return;
                }
                const dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => dialog.close());
                dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {});
            });
        } catch (dialogError) {
            console.error("No se pudo mostrar el resultado de la subida a Cloud Storage:", dialogError);
        }
    }

    let eventCompleted = false;
    function completeEventOnce() {
        if (eventCompleted) return;
        eventCompleted = true;
        if (event) event.completed();
    }

    try {
        if (!window.GCS || !window.GcsSaveBridge) {
            throw new Error("No se encontró el módulo de exportación a Cloud Storage (js/gcsExport.js o js/gcsSaveBridge.js).");
        }
        window.GcsSaveBridge.openSaveBucketDialog(async ({ bucket, objectName }) => {
            try {
                const result = await window.GCS.saveActiveWorkbookToBucketNamed(bucket, objectName);
                showResult(true, `Archivo subido correctamente a gs://${result.bucket}/${result.name}`);
            } catch (error) {
                console.error("Error al subir el Excel a Cloud Storage:", error);
                showResult(false, (error && error.message) ? error.message : String(error));
            } finally {
                completeEventOnce();
            }
        });
        // El diálogo de selección se muestra de forma asíncrona; si el
        // usuario cancela o tarda demasiado, no hay callback y hay que
        // completar el evento del ribbon igualmente para no dejarlo "colgado".
        setTimeout(completeEventOnce, 120000);
    } catch (error) {
        console.error("Error al abrir el selector de bucket:", error);
        showResult(false, (error && error.message) ? error.message : String(error));
        completeEventOnce();
    }
}

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
 * Nombre de la hoja de resultados ("CSV_RESULT" por defecto): desde que
 * "Editar report" guarda en EDIT_REPORT!D1 el nombre de la pestaña sobre
 * la que se pulsó, cada informe puede pintarse en una hoja de resultados
 * distinta. Todo el código que antes usaba el literal "CSV_RESULT" debe
 * usar en su lugar el valor de EDIT_REPORT!D1 (con "CSV_RESULT" como
 * valor por defecto si D1 está vacío, por compatibilidad).
 * ------------------------------------------------------------------- */
const DEFAULT_RESULT_SHEET_NAME = "CSV_RESULT";

function activeReportIdOrNull() {
    return window.ReportStore ? window.ReportStore.getActiveReportId() : null;
}

// Nº de informe con ceros a la izquierda (mismo formato que _pad3 en
// reportStore.js), para nombrar tanto la hoja de resultados por defecto
// como los rangos con nombre Draco_<id>_Rows/Cols/Values de ESE informe.
function pad3(n) {
    return String(n).padStart(3, "0");
}

/**
 * Nombres de los 3 rangos con nombre de un informe concreto. Antes eran
 * literales fijos ("Draco_001_Rows/Cols/Values") compartidos por TODOS los
 * informes del libro: refrescar el informe 2 borraba y reescribía el mismo
 * nombre que ya apuntaba a las celdas del informe 1 (lo que a su vez
 * BORRABA lo pintado del informe 1). Con el id del informe en el nombre,
 * cada uno tiene su propio juego de rangos.
 */
function dracoRangeNames(reportId) {
    const suffix = reportId ? pad3(reportId) : "001"; // "001" = compatibilidad si no hay ReportStore/informe activo
    return {
        rows: `Draco_${suffix}_Rows`,
        cols: `Draco_${suffix}_Cols`,
        values: `Draco_${suffix}_Values`
    };
}

// Variante SÍNCRONA: a partir de un grid de EDIT_REPORT ya cargado en
// memoria. reportId es opcional (por defecto, el informe activo); se
// consulta primero la hoja guardada PARA ESE INFORME en ReportStore, y solo
// si no la tiene (informe creado antes de que esto se guardara por
// informe) se recurre al valor físico compartido de EDIT_REPORT!D1, como
// antes.
function resultSheetNameFromGrid(editReportGrid, reportId) {
    const id = reportId !== undefined ? reportId : activeReportIdOrNull();
    if (window.ReportStore && id) {
        const stored = window.ReportStore.getReportResultSheetName(id);
        if (stored) return stored;
    }
    const v = String(cellValue(editReportGrid, 1, 4)).trim(); // EDIT_REPORT!D1 (compatibilidad)
    return v || DEFAULT_RESULT_SHEET_NAME;
}

// Variante ASÍNCRONA: cuando todavía no hay un grid de EDIT_REPORT cargado
// en el punto donde hace falta el nombre de la hoja (lee D1 directamente
// como fallback). Mismo criterio que resultSheetNameFromGrid.
async function getDracoResultSheetName(context, reportId) {
    const id = reportId !== undefined ? reportId : activeReportIdOrNull();
    if (window.ReportStore && id) {
        const stored = window.ReportStore.getReportResultSheetName(id);
        if (stored) return stored;
    }
    try {
        const cell = context.workbook.worksheets.getItem("EDIT_REPORT").getRange("D1");
        cell.load("values");
        await context.sync();
        const v = String((cell.values && cell.values[0] && cell.values[0][0]) || "").trim();
        return v || DEFAULT_RESULT_SHEET_NAME;
    } catch (e) {
        return DEFAULT_RESULT_SHEET_NAME;
    }
}

/**
 * Recorre los informes guardados en ReportStore y devuelve el id de aquel
 * cuya hoja de resultados (design.resultSheetName) coincide con
 * sheetName. Es el camino inverso a getReportResultSheetName/
 * resultSheetNameFromGrid, necesario porque la petición de
 * expandir/contraer desde EDIT_REPORT!T1 llega con el NOMBRE de la
 * pestaña, no con el id de informe.
 *
 * Si ningún informe guardado tiene esa hoja asignada (p.ej. informe
 * antiguo, creado antes de guardar resultSheetName por informe, o
 * simplemente no hay ReportStore disponible) se usa como último recurso
 * el informe activo, igual que hace el resto del código con
 * EDIT_REPORT!D1 (ver resultSheetNameFromGrid).
 */
function reportIdForResultSheet(sheetName) {
    if (window.ReportStore && sheetName) {
        const store = window.ReportStore.getAllReports() || {};
        for (const idStr of Object.keys(store)) {
            const report = store[idStr];
            const stored = report && report.design ? report.design.resultSheetName : "";
            if (stored && stored === sheetName) return Number(idStr);
        }
    }
    return activeReportIdOrNull();
}

/**
 * Crea la hoja de resultados de un informe si todavía no existe (antes se
 * asumía que "CSV_RESULT" ya venía en la plantilla del libro; con una hoja
 * por informe, cada informe nuevo necesita la suya la primera vez que se
 * refresca).
 */
async function ensureDracoResultSheetExists(context, sheetName) {
    const existing = context.workbook.worksheets.getItemOrNullObject(sheetName);
    existing.load("isNullObject");
    await context.sync();
    if (existing.isNullObject) {
        context.workbook.worksheets.add(sheetName);
        await context.sync();
    }
}

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

/* ---------------------------------------------------------------------
 * getEditReportGrid(): sustituye a getValuesGrid(context,"EDIT_REPORT")
 * en todas las lecturas del DISEÑO del informe (filtros/filas/columnas/
 * Estático/rangos H10-N10/H12-N12). Combina dos fuentes:
 *
 *   - El diseño del informe ACTIVO (ReportStore.getReportGrid), que ahora
 *     vive en JSON (Office roaming settings) y no en la hoja.
 *   - La hoja física EDIT_REPORT, para todo lo que sigue viviendo ahí:
 *     D1/E1 (hoja/celda activas), X1/Y1 (último SQL/JSON), B1
 *     (reconocimiento de miembros), A5 y el resto del picker de doble
 *     clic con el XLAM.
 *
 * Las coordenadas de diseño (ver isDesignOwnedCell) las gana SIEMPRE el
 * JSON, aunque estén vacías (para que un filtro/fila/columna borrado en
 * el taskpane desaparezca también aquí); el resto lo gana la hoja.
 * ------------------------------------------------------------------- */
function isDesignOwnedCell(row1, col1) {
    if (row1 === 10 && (col1 === 8 || col1 === 14)) return true;  // H10 / N10 (rangos)
    if (row1 === 12 && (col1 === 8 || col1 === 14)) return true;  // H12 / N12 (Estático)
    if (row1 >= 15) {
        if (col1 >= 3 && col1 <= 6) return true;   // C:F  filtros
        if (col1 >= 8 && col1 <= 13) return true;  // H:M  filas
        if (col1 >= 14 && col1 <= 19) return true; // N:S  columnas
    }
    return false;
}

function mergeEditReportGrid(physicalGrid, designGrid) {
    const maxRow = Math.max(
        physicalGrid.values.length + physicalGrid.startRow,
        designGrid.values.length + designGrid.startRow
    );
    const maxCol = 26; // hasta Z: cubre sobradamente D1/E1 (col 4/5), B1 (col 2), X1/Y1 (col 24/25)
    const values = [];
    for (let r = 0; r < maxRow; r++) {
        const row1 = r + 1;
        const rowArr = [];
        for (let c = 0; c < maxCol; c++) {
            const col1 = c + 1;
            rowArr.push(isDesignOwnedCell(row1, col1)
                ? cellValue(designGrid, row1, col1)
                : cellValue(physicalGrid, row1, col1));
        }
        values.push(rowArr);
    }
    return { values, startRow: 0, startCol: 0 };
}

// reportId es opcional: por defecto, el informe activo (mismo
// comportamiento que antes). Se puede pasar explícito para operar sobre UN
// informe concreto sin depender de cuál esté activo en el taskpane en ese
// momento (necesario para "Refrescar todos", que recorre varios informes
// uno a uno).
async function getEditReportGrid(context, reportIdOverride) {
    const physicalGrid = await getValuesGrid(context, "EDIT_REPORT");
    const reportId = reportIdOverride !== undefined ? reportIdOverride : activeReportIdOrNull();
    const designGrid = (window.ReportStore && reportId)
        ? window.ReportStore.getReportGrid(reportId)
        : { values: [], startRow: 0, startCol: 0 };
    return mergeEditReportGrid(physicalGrid, designGrid);
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

// Igual que parseAddress, pero para un rango completo ("T1", "T1:V1"...):
// devuelve el rectángulo {r1,c1,r2,c2} (base 1, normalizado) que ocupa.
// Se usa para saber si una escritura/pegado en EDIT_REPORT TOCA alguna de
// las celdas de control T1/U1/V1 aunque venga como un rango más amplio.
function parseAddressRange(addr) {
    addr = String(addr).replace(/\$/g, "");
    const parts = addr.split(":");
    const first = parseAddress(parts[0]);
    const second = parts.length > 1 ? parseAddress(parts[1]) : first;
    return {
        r1: Math.min(first.row, second.row),
        r2: Math.max(first.row, second.row),
        c1: Math.min(first.col, second.col),
        c2: Math.max(first.col, second.col)
    };
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

/**
 * Añade a ReportState.Filters los filtros "bloqueados" creados con el
 * botón "Añadir filtro" (rangos con nombre Draco_Filter_…, ver
 * FilterRangeStore): los atados directamente a este informe (meta.reportId
 * === reportId), más los creados para "Todos los informes" (meta.reportId
 * null). Solo participan los que YA tienen un filtro elegido (meta.filter
 * distinto de null); el rango recién creado, sin doble clic todavía, no
 * añade ninguna condición. buildWhere() no distingue estos de los filtros
 * normales de la zona "Filtros": se combinan todos con AND, igual que
 * antes.
 */
function appendLockedFilterRangesToState(reportId) {
    if (!window.FilterRangeStore) return;
    const all = window.FilterRangeStore.getAll();

    for (const rangeName of Object.keys(all)) {
        const meta = all[rangeName];
        if (!meta || !meta.filter) continue;

        const isAll = meta.reportId === null || meta.reportId === undefined;
        const appliesToThisReport = isAll
            || (reportId !== null && reportId !== undefined && Number(meta.reportId) === Number(reportId));
        if (!appliesToThisReport) continue;

        ReportState.Filters.push({
            Dimension: meta.dim,
            AttributeName: meta.isHierarchy ? "" : meta.name,
            Value: JSON.stringify(meta.filter),
            Locked: true,
            RangeName: rangeName
        });
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
                Hierarchy: hierRaw === "" ? 0 : Number(hierRaw),
                // Columna R de EDIT_REPORT: "X" = generar subtotal para este
                // campo (ver buildConfigSetsWithSubtotals). No se usa en
                // absoluto si NINGÚN campo del informe tiene esta marca.
                Subtotal: String(cellValue(editReportGrid, R, 18)).trim().toUpperCase() === "X"
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
                Hierarchy: hierRaw === "" ? 0 : Number(hierRaw),
                // Columna L de EDIT_REPORT: "X" = generar subtotal para este
                // campo (ver buildConfigSetsWithSubtotals).
                Subtotal: String(cellValue(editReportGrid, R, 12)).trim().toUpperCase() === "X"
            });
        }

        R++;
    }
}

function loadReportDefinition(editReportGrid, reportId) {
    loadFilters(editReportGrid);
    appendLockedFilterRangesToState(reportId !== undefined ? reportId : activeReportIdOrNull());
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

    return "FROM " +
        Provider.qualify(
            cellValue(measuresGrid, R, 3),
            cellValue(measuresGrid, R, 4),
            cellValue(measuresGrid, R, 5)
        ) + " f";
}

function buildJoins(relGrid) {
    let sql = "";

    const lastRow = lastRowInColumnValues(relGrid, 1);

    for (let R = 2; R <= lastRow; R++) {
        const dimname = cellValue(relGrid, R, 2);

        if (dimensionIsUsed(dimname)) {
            sql += CRLF
                + cellValue(relGrid, R, 11)
                + " JOIN "
                + Provider.qualify(cellValue(relGrid, R, 7), cellValue(relGrid, R, 8), cellValue(relGrid, R, 9))
                + " d" + cellValue(relGrid, R, 1) + CRLF
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

    if (Provider.key() === "snowflake") {
        // Snowflake no tiene UNNEST([STRUCT(...)]) ni ARRAY(SELECT...WHERE...);
        // el equivalente es construir el array con un elemento condicional
        // por ID (NULL si la condición no se cumple) y compactarlo con
        // ARRAY_CONSTRUCT_COMPACT, que descarta los NULL.
        const items = [];
        for (const [id, conds] of dict.entries()) {
            items.push("IFF(" + conds.join(" AND ") + ", " + id + ", NULL)");
        }

        let sql = "ARRAY_CONSTRUCT_COMPACT(" + CRLF;
        sql += items.map(s => "        " + s).join("," + CRLF) + CRLF;
        sql += "    )";
        return sql;
    }

    // BigQuery (comportamiento original, sin cambios)
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

    // ---- SELECT final: cruce ROW_ID x COLUMN_ID + agregación ----
    // BigQuery usa UNNEST(array) AS alias; Snowflake no soporta esa forma
    // sobre un array literal -> se usa LATERAL FLATTEN, que devuelve el
    // valor en una columna "VALUE" (de tipo VARIANT) que hay que castear.
    sql += "SELECT" + CRLF;
    if (Provider.key() === "snowflake") {
        sql += "    ROW_ID_F.VALUE::INTEGER AS ROW_ID," + CRLF;
        sql += "    COLUMN_ID_F.VALUE::INTEGER AS COLUMN_ID," + CRLF;
        sql += "    SUM(IMPORTE) AS IMPORTE" + CRLF;
        sql += "FROM TAGGED," + CRLF;
        sql += "LATERAL FLATTEN(input => ROW_IDS) AS ROW_ID_F," + CRLF;
        sql += "LATERAL FLATTEN(input => COLUMN_IDS) AS COLUMN_ID_F" + CRLF;
    } else {
        sql += "    ROW_ID," + CRLF;
        sql += "    COLUMN_ID," + CRLF;
        sql += "    SUM(IMPORTE) AS IMPORTE" + CRLF;
        sql += "FROM TAGGED," + CRLF;
        sql += "UNNEST(ROW_IDS) AS ROW_ID," + CRLF;
        sql += "UNNEST(COLUMN_IDS) AS COLUMN_ID" + CRLF;
    }
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

/**
 * Convierte las filas ya normalizadas de Provider/SF.runQuery (array de
 * objetos { COLUMNA: valor }) al mismo formato de texto que devuelve
 * BigQuery y que ya saben leer parseJsonValueTriples() y
 * jsonTo3MatricesCore(): ambos escanean el texto buscando el literal
 * `"v":` una vez por celda, en el mismo orden fila a fila / columna a
 * columna del SELECT — no hace falta que sea JSON válido de verdad, solo
 * que contenga exactamente esos tokens y nada más. Así Snowflake reutiliza
 * los dos parsers existentes sin tocarlos.
 */
function snowflakeRowsToPseudoBqJson(rows) {
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

async function executeSQL(sql) {
    if (Provider.key() === "snowflake") {
        const rows = await SF.runQuery(sql);
        return snowflakeRowsToPseudoBqJson(rows);
    }

    // BigQuery (comportamiento original, sin cambios)
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

async function jsonPaintValues(context, json, reportId) {
    const triples = parseJsonValueTriples(json);
    const resultSheetName = await getDracoResultSheetName(context, reportId);
    await ensureDracoResultSheetExists(context, resultSheetName);
    const sheet = context.workbook.worksheets.getItem(resultSheetName);

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
async function actualizarInformeFixedCore(reportIdOverride) {
    const reportId = reportIdOverride !== undefined ? reportIdOverride : activeReportIdOrNull();
    let sql;

    // 1) LoadReportDefinition + BuildSQL_Fixed + escritura de A1
    await Excel.run(async (context) => {
        const editReportGrid = await getEditReportGrid(context, reportId);
        const relGrid = await window.SemanticModelStore.getModelGrid("MODEL_RELATIONSHIP");
        const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");
        const atributesGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
        const resultSheetName = resultSheetNameFromGrid(editReportGrid, reportId);
        await ensureDracoResultSheetExists(context, resultSheetName);
        const csvGrid = await getFormulaGrid(context, resultSheetName);

        loadReportDefinition(editReportGrid, reportId);

        sql = await buildSQLFixed(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid);

        await context.sync();
    });

    // 2) ExecuteSQL contra BigQuery
    const json = await executeSQL(sql);

    // [Punto 8] SQL y JSON generados ya NO se escriben en A1/B1 de la hoja
    // de resultados: se escriben en EDIT_REPORT!X1 (SQL) e Y1 (JSON).
    await Excel.run(async (context) => {
        const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
        editReportSheet.getRange("X1").values = [[sql]];

        const EXCEL_CELL_CHAR_LIMIT = 32000; // límite real de Excel: 32767
        const jsonForCell = json.length > EXCEL_CELL_CHAR_LIMIT
            ? json.substring(0, EXCEL_CELL_CHAR_LIMIT) + " ...(truncado, JSON completo en la consola F12)"
            : json;
        editReportSheet.getRange("Y1").values = [[jsonForCell]];

        await context.sync();
    });

    // 3) JSON_PaintValues
    await Excel.run(async (context) => {
        await jsonPaintValues(context, json, reportId);
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
/**
 * Escribe el mensaje de un error en EDIT_REPORT!X1, para diagnosticar
 * fallos que antes solo quedaban en la consola (F12) y dejaban la celda
 * en blanco sin más pista de qué había pasado ni por qué no se llegó a
 * hacer la llamada al proveedor de datos.
 */
async function surfaceErrorToSheet(error) {
    try {
        await Excel.run(async (context) => {
            const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
            editReportSheet.getRange("X1").values = [["ERROR: " + (error && error.message ? error.message : String(error))]];
            await context.sync();
        });
    } catch (e2) {
        console.error("Además, no se ha podido escribir el error en EDIT_REPORT!X1:", e2);
    }
}

async function actualizarInformeFixed(event) {
    try {
        await actualizarInformeFixedCore();
    } catch (error) {
        console.error("Error al actualizar el informe (fixed):", error);
        await surfaceErrorToSheet(error);
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

/**
 * Interpreta la cadena guardada en la columna "Valor" de un filtro.
 * Formato actual: JSON { mode:"list", values|items, ranges,
 * excludeValues|excludeItems, excludeRanges }. Formatos de versiones
 * anteriores del diálogo ("values"/"range"/"mixed" con
 * include/valuesInclude/rangeInclude) y el formato "antiguo" (una cadena
 * simple, tratada como igualdad de un único valor) se siguen
 * interpretando igual, por compatibilidad. Misma lógica que
 * parseFilterValue de filterModal.js, duplicada aquí a propósito para que
 * commands.js no dependa de que filterModal.js se haya cargado antes.
 */
function parseStoredFilterValue(raw) {
    const s = String(raw || "").trim();
    if (s === "") return null;

    if (s[0] === "{") {
        try {
            const parsed = JSON.parse(s);
            if (parsed && typeof parsed === "object" && parsed.mode) return parsed;
        } catch (e) {
            // No era JSON: se trata como valor simple (ver abajo).
        }
    }

    return { mode: "values", include: true, values: [s] };
}

function sqlLiteralForFilter(atributesGrid, dimension, attributeName, rawValue) {
    const tipo = getAttributeType(atributesGrid, dimension, attributeName);
    if (UNQUOTED_TYPES.includes(tipo)) return String(rawValue);
    return "'" + String(rawValue).replace(/'/g, "''") + "'";
}

/**
 * Condición para una lista de valores (IN/NOT IN, o "=" / "<>" si es uno
 * solo), dado el campo SQL ya resuelto ("alias.ATRIBUTO").
 */
function buildValuesCondition(atributesGrid, dimension, attributeName, campo, values, include) {
    const vals = (values || []).filter(v => String(v).trim() !== "");
    if (vals.length === 0) return "";

    if (vals.length === 1) {
        const lit = sqlLiteralForFilter(atributesGrid, dimension, attributeName, vals[0]);
        return campo + (include ? " = " : " <> ") + lit;
    }

    const lista = vals.map(v => sqlLiteralForFilter(atributesGrid, dimension, attributeName, v)).join(", ");
    return campo + (include ? " IN (" : " NOT IN (") + lista + ")";
}

/**
 * Condición para un rango (BETWEEN/NOT BETWEEN), dado el campo SQL ya
 * resuelto ("alias.ATRIBUTO").
 */
function buildRangeCondition(atributesGrid, dimension, attributeName, campo, from, to, include) {
    const desde = sqlLiteralForFilter(atributesGrid, dimension, attributeName, from);
    const hasta = sqlLiteralForFilter(atributesGrid, dimension, attributeName, to);
    const cond = campo + " BETWEEN " + desde + " AND " + hasta;
    return include ? cond : ("NOT (" + cond + ")");
}

/**
 * Condición para un filtro "simple": un único atributo real, con selección
 * de varios valores (IN/NOT IN), un rango (BETWEEN/NOT BETWEEN), ambos a
 * la vez (modo "mixed": el campo cumple el filtro si cumple los valores O
 * el rango, cada uno con su propio incluir/excluir), o el modo actual del
 * diálogo ("list": ver buildListFilterCondition). Con un solo valor en
 * modo "values" se genera "=" / "<>", igual que antes.
 */
function buildSimpleFilterCondition(atributesGrid, relGrid, dimension, attributeName, filter) {
    const campo = getTableAlias(relGrid, dimension) + "." + attributeName;

    if (filter.mode === "list") {
        return buildListFilterCondition(atributesGrid, dimension, attributeName, campo, filter);
    }

    if (filter.mode === "range") {
        return buildRangeCondition(atributesGrid, dimension, attributeName, campo, filter.from, filter.to, filter.include);
    }

    if (filter.mode === "mixed") {
        const valuesCond = buildValuesCondition(atributesGrid, dimension, attributeName, campo, filter.values, filter.valuesInclude);
        const rangeCond = buildRangeCondition(atributesGrid, dimension, attributeName, campo, filter.from, filter.to, filter.rangeInclude);
        const partes = [valuesCond, rangeCond].filter(Boolean);
        if (partes.length === 0) return "";
        if (partes.length === 1) return partes[0];
        return "(" + partes.join(CRLF + "   OR ") + ")";
    }

    return buildValuesCondition(atributesGrid, dimension, attributeName, campo, filter.values, filter.include);
}

/**
 * Condición para el modo "list" del diálogo de filtro actual: valores y
 * rangos sueltos, incluidos y excluidos, combinables libremente entre sí
 * (ver cabecera de js/filterDialog.js). Se cumple si se satisface
 * cualquiera de los incluidos (valores O rangos), sin caer en ninguno de
 * los excluidos (ni valores NI rangos). Si no hay nada incluido, el filtro
 * es solo de exclusión (basta con no caer en lo excluido).
 */
function buildListFilterCondition(atributesGrid, dimension, attributeName, campo, filter) {
    const partes = [];

    const includedParts = [];
    const incValuesCond = buildValuesCondition(atributesGrid, dimension, attributeName, campo, filter.values, true);
    if (incValuesCond) includedParts.push(incValuesCond);
    (filter.ranges || []).forEach(r => {
        const c = buildRangeCondition(atributesGrid, dimension, attributeName, campo, r.from, r.to, true);
        if (c) includedParts.push(c);
    });
    if (includedParts.length > 0) {
        partes.push(includedParts.length === 1 ? includedParts[0] : ("(" + includedParts.join(CRLF + "   OR ") + ")"));
    }

    const excValuesCond = buildValuesCondition(atributesGrid, dimension, attributeName, campo, filter.excludeValues, false);
    if (excValuesCond) partes.push(excValuesCond);
    (filter.excludeRanges || []).forEach(r => {
        const c = buildRangeCondition(atributesGrid, dimension, attributeName, campo, r.from, r.to, false);
        if (c) partes.push(c);
    });

    if (partes.length === 0) return "";
    if (partes.length === 1) return partes[0];
    return "(" + partes.join(CRLF + "   AND ") + ")";
}

/**
 * Condición para un filtro de JERARQUÍA con varios miembros seleccionados.
 * Los miembros pueden pertenecer a niveles (atributos reales) distintos de
 * la jerarquía, así que se agrupan por atributo y se unen con OR: "está en
 * el continente Europa, o en el país España, o...". Con "Excluir" se niega
 * el conjunto completo. El modo "list" admite miembros incluidos Y
 * excluidos a la vez (ver buildHierarchyListCondition).
 */
function buildHierarchyFilterCondition(atributesGrid, relGrid, dimension, filter) {
    if (filter.mode === "list") {
        return buildHierarchyListCondition(atributesGrid, relGrid, dimension, filter);
    }

    const items = filter.items || [];
    if (items.length === 0) return "";

    const porAtributo = new Map();
    for (const it of items) {
        if (!porAtributo.has(it.attribute)) porAtributo.set(it.attribute, []);
        porAtributo.get(it.attribute).push(it.value);
    }

    const partes = [];
    for (const [attr, values] of porAtributo.entries()) {
        const cond = buildSimpleFilterCondition(atributesGrid, relGrid, dimension, attr,
            { mode: "values", include: true, values });
        if (cond) partes.push(cond);
    }

    if (partes.length === 0) return "";

    const combinado = partes.length === 1 ? partes[0] : ("(" + partes.join(CRLF + "   OR ") + ")");
    return filter.include ? combinado : ("NOT " + combinado);
}

/**
 * Condición de jerarquía para el modo "list": miembros incluidos
 * (items) O excluidos (excludeItems), cada grupo agrupado por atributo
 * real igual que el modo clásico, y ambos grupos combinados con AND
 * (cumple los incluidos Y no cae en ninguno de los excluidos).
 */
function buildHierarchyListCondition(atributesGrid, relGrid, dimension, filter) {
    const buildGroupedCondition = (items, include) => {
        if (!items || items.length === 0) return "";

        const porAtributo = new Map();
        for (const it of items) {
            if (!porAtributo.has(it.attribute)) porAtributo.set(it.attribute, []);
            porAtributo.get(it.attribute).push(it.value);
        }

        const partes = [];
        for (const [attr, values] of porAtributo.entries()) {
            const cond = buildSimpleFilterCondition(atributesGrid, relGrid, dimension, attr,
                { mode: "values", include: true, values });
            if (cond) partes.push(cond);
        }

        if (partes.length === 0) return "";
        const combinado = partes.length === 1 ? partes[0] : ("(" + partes.join(CRLF + "   OR ") + ")");
        return include ? combinado : ("NOT " + combinado);
    };

    const incCond = buildGroupedCondition(filter.items, true);
    const excCond = buildGroupedCondition(filter.excludeItems, false);
    const partes = [incCond, excCond].filter(Boolean);

    if (partes.length === 0) return "";
    if (partes.length === 1) return partes[0];
    return "(" + partes.join(CRLF + "   AND ") + ")";
}

function buildWhere(atributesGrid, relGrid) {
    // Filtros "vacíos" (sin valor seleccionado en el taskpane) no deben
    // añadirse al WHERE: se ignoran por completo, como si no existieran.
    const activeFilters = ReportState.Filters.filter(f => String(f.Value).trim() !== "");

    if (activeFilters.length === 0) return "";

    const condiciones = [];

    for (const f of activeFilters) {
        const filter = parseStoredFilterValue(f.Value);
        if (!filter) continue;

        // Un filtro de jerarquía con selección múltiple trae "items" y/o
        // (modo "list") "excludeItems" (cada uno con su propio atributo
        // real); todo lo demás (dimensión plana, o jerarquía "antigua" de
        // un solo valor) usa f.AttributeName.
        const esJerarquia = (filter.items && filter.items.length)
            || (filter.excludeItems && filter.excludeItems.length);
        const cond = esJerarquia
            ? buildHierarchyFilterCondition(atributesGrid, relGrid, f.Dimension, filter)
            : buildSimpleFilterCondition(atributesGrid, relGrid, f.Dimension, f.AttributeName, filter);

        if (cond) condiciones.push(cond);
    }

    if (condiciones.length === 0) return "";

    return "WHERE" + CRLF + condiciones.join(CRLF + "AND ");
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
function computeGroupingSetDimensionsAndHierarchies(relGrid) {
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

    return { dimensions, hierarchies, hierarchiesAcum, combinations };
}

function buildConfigSets(relGrid) {
    const { dimensions, hierarchiesAcum, combinations } = computeGroupingSetDimensionsAndHierarchies(relGrid);

    const tuples = combinations.map(item => groupingSetTupleToColumnList(item, dimensions, hierarchiesAcum));
    return groupingSetTuplesToText(tuples);
}

/**
 * Convierte UNA combinación del "odómetro" (generateHierarchyCombinations)
 * en la lista de columnas ("alias.attr") que le corresponden. Extraído de
 * buildConfigSets para poder reutilizarlo también al añadir los conjuntos
 * extra de subtotales (buildConfigSetsWithSubtotals).
 */
function groupingSetTupleToColumnList(item, dimensions, hierarchiesAcum) {
    const cols = [];
    for (let i = 0; i < item.length; i++) {
        for (let j = 1; j <= item[i]; j++) {
            if (item[i] > 0) {
                if (i === 0) {
                    cols.push(dimensions[j]);
                } else {
                    cols.push(dimensions[j + hierarchiesAcum[i - 1]]);
                }
            }
            // item[i] === 0 no ocurre nunca en el flujo normal (generateHierarchyCombinations
            // nunca baja de 1): se deja sin aportar columnas, en vez del literal "[TOTAL]"
            // que tenía el código original (nunca llegaba a ejecutarse).
        }
    }
    return cols;
}

// Une una lista de tuplas (cada una: array de "alias.attr") en el texto
// final "GROUP BY GROUPING SETS ( (...), (...), ... )".
function groupingSetTuplesToText(tuples) {
    let texto = "";
    tuples.forEach((cols, idx) => {
        texto += idx === 0 ? "GROUP BY GROUPING SETS ( ( " : ",( ";
        texto += cols.join(", ");
        texto += " )";
    });
    texto += " )";
    return texto;
}

/**
 * Nombre de la columna GROUPING(...) para un campo dado, usada tanto en
 * buildSelectWithSubtotals como en buildFinalSelectWithSubtotals.
 */
function subtotalFlagName(attributeName) {
    return "IS_SUBTOTAL_" + String(attributeName).toUpperCase().replace(/[^A-Z0-9_]/g, "_");
}

/**
 * Variante de buildConfigSets() que AÑADE conjuntos extra de agrupación
 * para los campos marcados con "X" en EDIT_REPORT columnas L (eje de
 * ReportState.Columns) o R (eje de ReportState.Rows) -ver loadColumns/
 * loadRows-. Los conjuntos que ya generaba buildConfigSets (jerarquías)
 * se mantienen exactamente igual; esto solo AÑADE combinaciones nuevas,
 * nunca quita ni modifica las existentes.
 *
 * Semántica de una marca en el campo situado en la posición i de su eje:
 * añade UN conjunto de agrupación con el "prefijo" de campos de ESE MISMO
 * eje anteriores a i (puede quedar vacío = total general de ese eje),
 * combinado con el OTRO eje a su detalle COMPLETO. Así, marcar dos campos
 * consecutivos de un eje genera dos subtotales distintos (uno por cada
 * nivel), en vez de todas las combinaciones cruzadas posibles.
 */
function buildConfigSetsWithSubtotals(relGrid) {
    const { dimensions, hierarchiesAcum, combinations } = computeGroupingSetDimensionsAndHierarchies(relGrid);
    const baseTuples = combinations.map(item => groupingSetTupleToColumnList(item, dimensions, hierarchiesAcum));

    const colsAxisFields = ReportState.Columns.map(c => getTableAlias(relGrid, c.Dimension) + "." + c.AttributeName);
    const rowsAxisFields = ReportState.Rows.map(r => getTableAlias(relGrid, r.Dimension) + "." + r.AttributeName);

    const extraTuples = [];

    ReportState.Columns.forEach((c, idx) => {
        if (c.Subtotal) {
            extraTuples.push([...colsAxisFields.slice(0, idx), ...rowsAxisFields]);
        }
    });
    ReportState.Rows.forEach((r, idx) => {
        if (r.Subtotal) {
            extraTuples.push([...colsAxisFields, ...rowsAxisFields.slice(0, idx)]);
        }
    });

    return groupingSetTuplesToText([...baseTuples, ...extraTuples]);
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

    // Una columna por medida (antes: "IMPORTE" fijo, lo que descartaba
    // cualquier medida adicional aunque ya viniera calculada en REPORT_DATA).
    for (const m of ReportState.Measures) {
        sql += "," + CRLF + "    " + m.Name;
    }
    sql += CRLF + CRLF + "FROM REPORT_DATA";

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

/* =======================================================================
 * SUBTOTALES "NO JERARQUICOS" (EDIT_REPORT columnas L/R = subtotal,
 * M/S = orden -- de momento M/S no se usan aqui, solo subtotal).
 * Estas funciones NUEVAS solo se usan cuando hay AL MENOS UN campo con
 * Subtotal=true en cualquiera de los dos ejes (ver buildSQL). Si no hay
 * ninguno, buildSQL sigue llamando a las funciones de siempre (buildSelect,
 * buildFinalSelect, buildFinalWhere, buildConfigSets) sin ningun cambio,
 * asi que el SQL generado es exactamente igual que antes de este cambio.
 * ===================================================================== */

function hasAnySubtotalMarked() {
    return (ReportState.Columns || []).some(c => c.Subtotal) || (ReportState.Rows || []).some(r => r.Subtotal);
}

/**
 * Igual que buildSelect(), pero anadiendo un GROUPING(campo) por cada
 * campo de AMBOS ejes (no solo los marcados con X): un campo situado
 * DESPUES de otro marcado, dentro del mismo eje, tambien puede quedar
 * excluido de un grouping set concreto (ver buildConfigSetsWithSubtotals),
 * asi que necesita su propio indicador para poder etiquetarse como
 * "TOTAL" en buildFinalSelectWithSubtotals.
 */
function buildSelectWithSubtotals(measuresGrid, relGrid) {
    let sql = buildSelect(measuresGrid, relGrid);

    for (const c of ReportState.Columns) {
        sql += "," + CRLF + "    GROUPING(" + getTableAlias(relGrid, c.Dimension) + "." + c.AttributeName + ") AS " + subtotalFlagName(c.AttributeName);
    }
    for (const r of ReportState.Rows) {
        sql += "," + CRLF + "    GROUPING(" + getTableAlias(relGrid, r.Dimension) + "." + r.AttributeName + ") AS " + subtotalFlagName(r.AttributeName);
    }

    return sql;
}

// Necesita CAST(... AS STRING) para poder mostrar 'TOTAL' en su lugar?
function attributeNeedsStringCast(atributesGrid, dimension, attributeName) {
    const tipo = getAttributeType(atributesGrid, dimension, attributeName);
    return UNQUOTED_TYPES.includes(tipo);
}

function fieldFinalExpressionWithSubtotal(atributesGrid, field) {
    const flag = subtotalFlagName(field.AttributeName);
    const rawExpr = attributeNeedsStringCast(atributesGrid, field.Dimension, field.AttributeName)
        ? ("CAST(" + field.AttributeName + " AS STRING)")
        : field.AttributeName;
    return "CASE WHEN " + flag + " = 1 THEN 'TOTAL' ELSE " + rawExpr + " END AS " + field.AttributeName;
}

/**
 * ORDER BY interno de un DENSE_RANK para UN eje (ReportState.Columns o
 * ReportState.Rows), colocando sus filas de subtotal justo antes o despues
 * del grupo de detalle al que resumen, segun "Mostrar subtotales arriba"
 * (EDIT_REPORT!D4 -> subtotalsOnTop). Si ese eje no tiene NINGUN campo
 * marcado, devuelve exactamente lo mismo que el bucle original de
 * buildFinalSelect (misma lista de nombres de campo, sin nada anadido).
 */
function buildAxisOrderByWithSubtotals(fields, subtotalsOnTop) {
    if (!fields.some(f => f.Subtotal)) {
        return fields.filter(f => f.Hierarchy > 0).map(f => f.AttributeName);
    }

    const topVal = subtotalsOnTop ? 0 : 1;
    const bottomVal = subtotalsOnTop ? 1 : 0;
    const parts = [];

    // Si el PRIMER campo del eje esta marcado, su fila de "total general del
    // eje" (todos los campos de este eje colapsados a la vez) se ordena la
    // primera (o la ultima) de todas.
    if (fields.length > 0 && fields[0].Subtotal) {
        parts.push("CASE WHEN " + subtotalFlagName(fields[0].AttributeName) + " = 1 THEN " + topVal + " ELSE " + bottomVal + " END");
    }

    fields.forEach((f, idx) => {
        if (idx > 0) {
            // El subtotal de ESTE campo (si esta marcado) se intercala junto
            // al grupo del campo anterior.
            parts.push(subtotalFlagName(f.AttributeName) + (subtotalsOnTop ? " DESC" : " ASC"));
        }
        if (f.Hierarchy > 0) parts.push(f.AttributeName);
    });

    return parts;
}

/**
 * Igual que buildFinalSelect(), pero:
 *   - Los ORDER BY internos de ROW_ID/COLUMN_ID usan
 *     buildAxisOrderByWithSubtotals (arriba) en vez de la lista simple.
 *   - Cada campo se etiqueta 'TOTAL' cuando su GROUPING(...) = 1
 *     (fieldFinalExpressionWithSubtotal), en vez de mostrarse tal cual.
 */
function buildFinalSelectWithSubtotals(atributesGrid, subtotalsOnTop) {
    let sql = "SELECT" + CRLF + CRLF;

    const rowOrderParts = buildAxisOrderByWithSubtotals(ReportState.Columns, subtotalsOnTop);
    sql += "    DENSE_RANK() OVER (" + CRLF + "        ORDER BY";
    rowOrderParts.forEach((part, idx) => {
        sql += idx === 0 ? (CRLF + "            " + part) : ("," + CRLF + "            " + part);
    });
    sql += CRLF + "    ) AS ROW_ID," + CRLF + CRLF;

    const colOrderParts = buildAxisOrderByWithSubtotals(ReportState.Rows, subtotalsOnTop);
    sql += "    DENSE_RANK() OVER (" + CRLF + "        ORDER BY";
    colOrderParts.forEach((part, idx) => {
        sql += idx === 0 ? (CRLF + "            " + part) : ("," + CRLF + "            " + part);
    });
    sql += CRLF + "    ) AS COLUMN_ID," + CRLF + CRLF;

    let first = true;
    for (const c of ReportState.Columns) {
        sql += first ? ("    " + fieldFinalExpressionWithSubtotal(atributesGrid, c)) : ("," + CRLF + "    " + fieldFinalExpressionWithSubtotal(atributesGrid, c));
        first = false;
    }
    for (const r of ReportState.Rows) {
        sql += "," + CRLF + "    " + fieldFinalExpressionWithSubtotal(atributesGrid, r);
    }

    for (const m of ReportState.Measures) {
        sql += "," + CRLF + "    " + m.Name;
    }
    sql += CRLF + CRLF + "FROM REPORT_DATA";

    return sql;
}

/**
 * Igual que buildFinalWhere(), pero el eje que tenga algun campo marcado
 * con subtotal NO excluye sus filas "todo NULL": esas son precisamente los
 * totales/subtotales que ahora queremos conservar. El otro eje (si no
 * tiene marcas) conserva el filtrado de siempre, sin ningun cambio.
 */
function buildFinalWhereWithSubtotals() {
    let sql = "";
    let currentDim = "";
    let first = true;

    const colsHaveSubtotal = ReportState.Columns.some(c => c.Subtotal);
    const rowsHaveSubtotal = ReportState.Rows.some(r => r.Subtotal);

    if (!colsHaveSubtotal) {
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
        if (currentDim !== "") sql += CRLF + ")";
    }

    currentDim = "";

    if (!rowsHaveSubtotal) {
        for (const r of ReportState.Rows) {
            if (r.Hierarchy > 0) {
                if (currentDim !== r.Dimension.toUpperCase()) {
                    if (currentDim !== "") {
                        sql += CRLF + CRLF + ")" + CRLF + CRLF + "AND NOT (" + CRLF;
                    } else {
                        sql += (sql.trim() !== "" ? (CRLF + CRLF + "AND NOT (" + CRLF) : ("WHERE NOT (" + CRLF));
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
        if (currentDim !== "") sql += CRLF + ")";
    }

    return sql;
}

/**
 * BuildSQL: ensambla el flujo "Dinamico" completo. Si NINGUN campo tiene
 * marcado un subtotal (columnas L/R de EDIT_REPORT), usa exactamente las
 * mismas funciones de siempre (SQL identico al de antes de este cambio).
 * Si hay al menos uno, usa las variantes "WithSubtotals" de arriba.
 */
function buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop) {
    let sql = "";
    const withSubtotals = hasAnySubtotalMarked();

    sql += "WITH REPORT_DATA AS (" + CRLF + CRLF;
    sql += (withSubtotals ? buildSelectWithSubtotals(measuresGrid, relGrid) : buildSelect(measuresGrid, relGrid)) + CRLF + CRLF;
    sql += buildFrom(measuresGrid) + CRLF + CRLF;
    sql += buildJoins(relGrid) + CRLF + CRLF;
    sql += buildWhere(atributesGrid, relGrid) + CRLF + CRLF;
    sql += (withSubtotals ? buildConfigSetsWithSubtotals(relGrid) : buildConfigSets(relGrid)) + CRLF + CRLF;
    sql += ")" + CRLF + CRLF;
    sql += (withSubtotals ? buildFinalSelectWithSubtotals(atributesGrid, !!subtotalsOnTop) : buildFinalSelect()) + CRLF + CRLF;
    sql += (withSubtotals ? buildFinalWhereWithSubtotals() : buildFinalWhere()) + CRLF + CRLF;
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

// Estado de contraído/expandido, último JSON e indicadores +/- pintados:
// TODO ESTO ANTES ERA GLOBAL AL LIBRO (un único informe asumido). Con
// varios informes en el mismo libro, cada uno necesita su propia copia —
// si no, refrescar el informe 2 "olvidaba" qué nodos tenía contraídos el
// informe 1 y dejaba sus indicadores +/- apuntando a nodos que ya no
// existían. Se guardan en mapas indexados por reportId (0 = sin
// ReportStore/informe, por compatibilidad).
const DracoCollapseStateByReport = new Map(); // reportId -> { rows:{signature,collapsed}, cols:{signature,collapsed} }
const DracoLastJsonByReport = new Map();      // reportId -> último JSON de BigQuery de ESE informe
const DracoIndicatorMapByReport = new Map();  // reportId -> Map("row_col" -> { axis, nodeKey })

function dracoStateKey(reportId) {
    return reportId || 0;
}

function getDracoCollapseState(reportId) {
    const key = dracoStateKey(reportId);
    if (!DracoCollapseStateByReport.has(key)) {
        DracoCollapseStateByReport.set(key, {
            rows: { signature: null, collapsed: new Set() },
            cols: { signature: null, collapsed: new Set() }
        });
    }
    return DracoCollapseStateByReport.get(key);
}

function getDracoIndicatorMap(reportId) {
    const key = dracoStateKey(reportId);
    if (!DracoIndicatorMapByReport.has(key)) DracoIndicatorMapByReport.set(key, new Map());
    return DracoIndicatorMapByReport.get(key);
}

// Hojas de resultados en las que ya se han enganchado los listeners de
// clic (+/-) y reconocimiento de miembros (antes era un único booleano:
// solo se enganchaban en la PRIMERA hoja de resultados usada en la sesión,
// así que el informe 2, si se pinta en otra hoja, se quedaba sin clic
// interactivo). Ahora es un Set de nombres de hoja: se engancha en cada
// hoja de resultados distinta que se vaya usando.
const DracoHandlerRegisteredSheets = new Set();
let DracoEditReportHandlerRegistered = false; // evita registrar el listener de EDIT_REPORT!A5 (picker) más de una vez — INDEPENDIENTE de lo anterior: no depende de que exista ninguna hoja de resultados ni de que se haya refrescado nunca
let DracoSuppressChangeEvents = false; // true mientras jsonTo3Matrices pinta celdas (evita que el reconocimiento de miembros reaccione a nuestras propias escrituras)

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

function resetDracoCollapseIfAxisChanged(axis, signature, reportId) {
    const st = getDracoCollapseState(reportId)[axis];
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
        subtotalsOnTop: false,
        overwriteFormats: true,
        autoFitColumns: true
    };
    try {
        if (!window.ReportStore) return defaults;
        return Object.assign({}, defaults, window.ReportStore.getActiveReportProperties());
    } catch (e) {
        console.warn("No se pudieron leer las propiedades del informe activo, se usan valores por defecto:", e);
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
    const totalDimFilas = ReportState.ColumnCount; // nº de entradas EDIT_REPORT del eje Filas (jerarquías expandidas a 1 fila por nivel)
    const totalDimCols = ReportState.RowCount;      // nº de entradas EDIT_REPORT del eje Columnas
    // Igual que en jsonTo3Matrices: cada COLUMNA FÍSICA corresponde a una
    // entrada con NIVEL/flag === 1 (inicio de un campo nuevo); los niveles
    // siguientes de esa misma jerarquía (flag > 1) se pintan en la MISMA
    // columna física, así que no deben generar una entrada nueva en la
    // tabla devuelta (si no, con jerarquías multinivel + campos simples
    // mezclados, las columnas físicas quedaban desalineadas con el campo
    // correcto).
    const fields = [];
    if (axis === "rows") {
        let iAux = 0;
        for (let i = 1; i <= totalDimFilas; i++) {
            const flag = Number(cellValue(editReportGrid, i + 14, 10)); // J
            if (flag === 1) {
                iAux++;
                fields[iAux] = { dim: cellValue(editReportGrid, i + 14, 8), attr: cellValue(editReportGrid, i + 14, 9) };
            }
        }
    } else {
        let iAux = 0;
        for (let i = 1; i <= totalDimCols; i++) {
            const flag = Number(cellValue(editReportGrid, i + 14, 16)); // P
            if (flag === 1) {
                iAux++;
                fields[iAux] = { dim: cellValue(editReportGrid, i + 14, 14), attr: cellValue(editReportGrid, i + 14, 15) };
            }
        }
    }
    return fields;
}

async function convertAxisStaticFormulas(axis, makeStatic) {
    const reportId = activeReportIdOrNull();
    const rangeName = axis === "rows" ? dracoRangeNames(reportId).rows : dracoRangeNames(reportId).cols;
    // DracoCollapseState/DracoIndicatorMap usan "rows"/"cols" como clave de
    // eje (no "columns"), a diferencia del parámetro `axis` de esta función.
    const stateAxis = axis === "rows" ? "rows" : "cols";
    DracoSuppressChangeEvents = true;
    try {

    if (makeStatic) {
        // El eje pasa a Estático: se quita el glifo +/- de sus celdas (más
        // abajo) y deja de tener sentido "contraer/expandir", así que hay
        // que OLVIDAR tanto los indicadores ya registrados de ese eje como
        // su estado de contraído — si no, quedan entradas obsoletas en
        // DracoIndicatorMap que siguen apuntando a esas mismas posiciones
        // de celda (ahora con fórmula EPM_VALUE, sin jerarquía), y una
        // petición desde EDIT_REPORT!T1:V1 sobre esa misma posición (ver
        // handleDracoEditReportExpandCollapseRequest) las tomaría por
        // buenas e intentaría un repintado/"refresco" con el último JSON,
        // aunque el eje ya no tenga jerarquías desplegables.
        const indicatorMap = getDracoIndicatorMap(reportId);
        for (const [key, meta] of indicatorMap) {
            if (meta.axis === stateAxis) indicatorMap.delete(key);
        }
        const st = getDracoCollapseState(reportId)[stateAxis];
        st.signature = null;
        st.collapsed = new Set();
    }

    await Excel.run(async (context) => {
        const editReportGrid = await getEditReportGrid(context);
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
function dracoColToLetters(col) {    let s = "";
    while (col > 0) {
        const rem = (col - 1) % 26;
        s = String.fromCharCode(65 + rem) + s;
        col = Math.floor((col - 1) / 26);
    }
    return s;
}

/**
 * Limpia comentarios (notas) que pudieran haber quedado de versiones
 * anteriores del add-in. Ya NO se anade ningun tooltip de
 * "Clic para expandir/contraer" sobre las celdas.
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
        console.warn("No se pudieron limpiar los comentarios previos:", e);
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
    const reportId = activeReportIdOrNull();
    const resultSheetName = await getDracoResultSheetName(context, reportId);
    const sheet = context.workbook.worksheets.getItem(resultSheetName);
    const cell = sheet.getRange(addr);
    cell.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

    const rangeNames = dracoRangeNames(reportId);
    const rowsNamed = context.workbook.names.getItemOrNullObject(rangeNames.rows);
    const colsNamed = context.workbook.names.getItemOrNullObject(rangeNames.cols);
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

    const editReportGrid = await getEditReportGrid(context, reportId);
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
            const resultSheetName = await getDracoResultSheetName(context);
            const sheet = context.workbook.worksheets.getItem(resultSheetName);
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
            const resultSheetName = await getDracoResultSheetName(context);
            const sheet = context.workbook.worksheets.getItem(resultSheetName);
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
 * Comprueba si una dirección (celda o rango) de la hoja de resultados cae
 * dentro del rango con nombre Draco_001_Rows y actualiza SOLO la etiqueta
 * del botón "Crear informe" del ribbon: "Editar informe" si la selección
 * está dentro de Draco_001_Rows, o de vuelta a "Crear informe" si no lo
 * está o si el rango no existe (getItemOrNullObject). No cambia la
 * Action del botón ni ninguna otra lógica: reutiliza el mismo helper
 * requestRibbonLabelUpdate que ya usan "Pausar refresco" y "Reconocimiento
 * de miembros" (definido más abajo en este archivo).
 */
async function updateCrearInformeLabelForAddress(addr) {
    let withinRowsRange = false;

    try {
        await Excel.run(async (context) => {
            const reportId = activeReportIdOrNull();
            const resultSheetName = await getDracoResultSheetName(context, reportId);
            const sheet = context.workbook.worksheets.getItemOrNullObject(resultSheetName);
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) return;

            const cell = sheet.getRange(addr);
            cell.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

            const rowsNamed = context.workbook.names.getItemOrNullObject(dracoRangeNames(reportId).rows);
            rowsNamed.load("isNullObject");
            await context.sync();

            // El rango puede no existir todavía (p.ej. sin informe creado nunca):
            // en ese caso se deja withinRowsRange en false -> "Crear informe".
            if (rowsNamed.isNullObject) return;

            const r = rowsNamed.getRange();
            r.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
            await context.sync();

            withinRowsRange =
                cell.rowIndex < r.rowIndex + r.rowCount &&
                cell.rowIndex + cell.rowCount > r.rowIndex &&
                cell.columnIndex < r.columnIndex + r.columnCount &&
                cell.columnIndex + cell.columnCount > r.columnIndex;
        });
    } catch (e) {
        console.warn("[Draco] No se pudo comprobar la selección frente al rango de filas del informe activo:", e);
    }

    await requestRibbonLabelUpdate(
        "CrearInformeButton",
        withinRowsRange ? "Editar informe" : "Crear informe",
        "EdicionGroup"
    );
}

/**
 * onSelectionChanged: dispara la comprobación anterior con la dirección
 * seleccionada. Es un listener más, independiente de
 * handleDracoMemberRecognitionSelection (no altera su comportamiento).
 */
async function handleDracoRibbonLabelSelection(eventArgs) {
    try {
        let addr = (eventArgs && eventArgs.address) || "";
        if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
        if (!addr) return;

        await updateCrearInformeLabelForAddress(addr);
    } catch (e) {
        console.error("Error al actualizar la etiqueta de Crear/Editar informe:", e);
    }
}

/* =================================================================
 * "Añadir filtro": rango con nombre sobre CUALQUIER celda del libro que
 * representa un filtro de dimensión, independiente del filtro de la zona
 * "Filtros" del diseño del informe (que vive dentro de EDIT_REPORT/
 * ReportStore). El botón "Añadir filtro" del taskpane (ver
 * openAddFilterRangeModal/createFilterRangeFromModal en taskpane.js) crea
 * el rango con nombre sobre la celda activa; aquí se localiza ese rango al
 * hacer clic sobre él, se abre el mismo selector filterDialog.html que usa
 * FilterModal, y se pinta el resultado en la celda.
 *
 * Convención de nombre: Draco_Filter_<DIM>_<CAMPO>_<sufijo>, donde sufijo
 * es el nº de informe con 3 dígitos (001, 002…) o "all" si el filtro se
 * creó para "Todos los informes". El nombre solo sirve para identificar
 * el rango en el Administrador de nombres de Excel: los metadatos reales
 * (dimensión, campo, jerarquía sí/no, informe/"todos" y el último filtro
 * aplicado) viven en FilterRangeStore (roaming settings), indexados por
 * ese mismo nombre — así no hace falta "parsear" el nombre para
 * recuperarlos, y el nombre puede sanearse (mayúsculas, sin acentos ni
 * espacios) sin perder información.
 * ================================================================= */

function sanitizeDracoNamePart(s) {
    return String(s || "")
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // quita acentos
        .replace(/[^A-Za-z0-9]+/g, "_")
        .replace(/^_+|_+$/g, "")
        .toUpperCase() || "X";
}

function dracoFilterRangeName(dim, name, suffix) {
    return `Draco_Filter_${sanitizeDracoNamePart(dim)}_${sanitizeDracoNamePart(name)}_${suffix}`;
}

/**
 * Busca, entre los rangos con nombre registrados en FilterRangeStore, si
 * alguno apunta EXACTAMENTE a la celda indicada (misma hoja, misma
 * dirección). Devuelve {rangeName, meta} o null si esa celda no es
 * ninguno de nuestros rangos de filtro.
 */
async function locateDracoFilterRangeAtAddress(context, sheetName, addr) {
    if (!window.FilterRangeStore) return null;
    const names = window.FilterRangeStore.listNames();
    if (names.length === 0) return null;

    const candidates = names.map(n => {
        const item = context.workbook.names.getItemOrNullObject(n);
        item.load("isNullObject");
        return { name: n, item };
    });
    await context.sync();

    const alive = candidates.filter(c => !c.item.isNullObject);
    alive.forEach(c => {
        c.range = c.item.getRange();
        c.range.load("address");
    });
    await context.sync();

    const target = (sheetName + "!" + addr).toUpperCase();
    for (const c of alive) {
        if (String(c.range.address || "").toUpperCase() === target) {
            return { rangeName: c.name, meta: window.FilterRangeStore.get(c.name) };
        }
    }
    return null;
}

/**
 * Inverso de locateDracoFilterRangeAtAddress: dado el NOMBRE de un rango
 * de filtro, devuelve en qué hoja y celda está ahora mismo ({sheetName,
 * addr}), o null si el nombre ya no existe en el libro (p.ej. se borró la
 * hoja). Lo usa el taskpane para poder reabrir el selector con doble clic
 * sobre la etiqueta "bloqueada" de la zona Filtros, sin que el usuario
 * tenga que ir a buscar la celda en Excel.
 */
async function resolveDracoFilterRangeAddress(rangeName) {
    let result = null;
    await Excel.run(async (context) => {
        const item = context.workbook.names.getItemOrNullObject(rangeName);
        item.load("isNullObject");
        await context.sync();
        if (item.isNullObject) return;

        const range = item.getRange();
        range.load("address");
        const sheet = range.worksheet;
        sheet.load("name");
        await context.sync();

        result = { addr: String(range.address).split("!").pop(), sheetName: sheet.name };
    });
    return result;
}

/**
 * Construye los "segmentos" de texto (valor o rango, incluido o excluido)
 * a partir del objeto filtro que devuelve filterDialog.js (mode:"list",
 * ver cabecera de js/filterModal.js). Mismo orden que describeFilter():
 * valores incluidos, rangos incluidos, valores excluidos, rangos
 * excluidos.
 */
function buildDracoFilterCellSegments(filter) {
    if (!filter) return [];

    const incValues = filter.items ? filter.items.map(it => it.value) : (filter.values || []);
    const excValues = filter.excludeItems ? filter.excludeItems.map(it => it.value) : (filter.excludeValues || []);
    const incRanges = filter.ranges || [];
    const excRanges = filter.excludeRanges || [];

    const segments = [];
    incValues.forEach(v => segments.push({ text: String(v), excluded: false }));
    incRanges.forEach(r => segments.push({ text: `[${r.from || "…"}-${r.to || "…"}]`, excluded: false }));
    excValues.forEach(v => segments.push({ text: String(v), excluded: true }));
    excRanges.forEach(r => segments.push({ text: `[${r.from || "…"}-${r.to || "…"}]`, excluded: true }));
    return segments;
}

/**
 * Pinta en la celda el resumen del filtro (segmentos separados por ", "),
 * coloreando en granate SOLO los tramos excluidos.
 *
 * IMPORTANTE: Excel.Range NO tiene un método getCharacters() (eso es de
 * Word/VBA, no de la Excel JS API). El formato de texto por tramos dentro
 * de una misma celda ("rich text runs") se hace con
 * Range.setCellProperties([[{ textRuns: [...] }]]), disponible desde
 * ExcelApi 1.18 (liberado feb-2025) vía Range.setCellProperties/
 * getCellProperties + CellPropertiesInternal.textRuns. Si el Excel del
 * usuario no soporta ese requirement set, se colorea la celda ENTERA en
 * granate cuando haya alguna exclusión — la mejor aproximación posible
 * sin coloreado por tramos.
 */
async function paintDracoFilterCell(context, sheet, addr, filter) {
    const EXCLUDED_COLOR = "#A80000";
    const DEFAULT_COLOR = "#000000";
    const segments = buildDracoFilterCellSegments(filter);
    const cell = sheet.getRange(addr);

    if (segments.length === 0) {
        cell.values = [[""]];
        cell.format.font.color = DEFAULT_COLOR;
        await context.sync();
        return;
    }

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.18")) {
        try {
            const textRuns = [];
            segments.forEach((seg, i) => {
                textRuns.push({
                    text: seg.text,
                    font: { color: seg.excluded ? EXCLUDED_COLOR : DEFAULT_COLOR }
                });
                if (i < segments.length - 1) {
                    // separador entre segmentos, siempre en color normal
                    textRuns.push({ text: ", ", font: { color: DEFAULT_COLOR } });
                }
            });

            cell.setCellProperties([[{ textRuns }]]);
            await context.sync();
            return;
        } catch (err) {
            console.warn("[Draco] Añadir filtro: coloreado por tramos no disponible en este Excel, se colorea la celda completa.", err);
        }
    }

    // Fallback: ExcelApi < 1.18 (o setCellProperties ha fallado) -> celda completa
    const fullText = segments.map(s => s.text).join(", ");
    cell.values = [[fullText]];
    cell.format.font.color = segments.some(s => s.excluded) ? EXCLUDED_COLOR : DEFAULT_COLOR;
    await context.sync();
}

/**
 * Abre filterDialog.html — el MISMO diálogo de selección de valores que
 * usa la zona "Filtros" del taskpane a través de FilterModal.open — como
 * diálogo independiente de Office, igual que openMemberRecognitionPicker,
 * para que funcione también desde el runtime oculto de comandos (sin
 * depender de que el taskpane esté abierto). Al aceptar, pinta el
 * resultado en la celda y actualiza FilterRangeStore.
 */
async function openDracoFilterRangePicker(addr, sheetName, rangeName, meta) {
    let items = [];
    try {
        const sql = await window.ExcelService.buildFilterValuesSQL(meta.dim, meta.name);
        if (!sql) {
            console.warn("[Draco] Añadir filtro: no se ha encontrado el atributo o jerarquía.", meta);
            return;
        }
        const json = await window.ExcelService.executeSQL(sql);
        items = parseMemberJsonTree(json);
    } catch (err) {
        console.error("[Draco] Añadir filtro: error cargando los valores del filtro:", err);
        return;
    }

    const dialogUrl = new URL("filterDialog.html", window.location.href).href;

    await new Promise((resolve) => {
        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 65, width: 48, displayInIframe: false },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error(
                        "[Draco] Añadir filtro: displayDialogAsync ha fallado:",
                        asyncResult.error && asyncResult.error.code,
                        asyncResult.error && asyncResult.error.message
                    );
                    resolve();
                    return;
                }

                const dialog = asyncResult.value;
                let settled = false;
                const closeDialog = () => { try { dialog.close(); } catch (e) { /* ya cerrado */ } };

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                    let payload;
                    try {
                        payload = JSON.parse(arg.message);
                    } catch (err) {
                        console.error("[Draco] Añadir filtro: mensaje del diálogo no es JSON válido:", err);
                        return;
                    }

                    if (payload.type === "ready") {
                        dialog.messageChild(JSON.stringify({
                            items,
                            fieldData: { dim: meta.dim, name: meta.name, isHierarchy: meta.isHierarchy },
                            currentFilter: meta.filter || null,
                            initialSearch: ""
                        }));
                        return;
                    }

                    if (payload.type === "apply") {
                        settled = true;
                        closeDialog();
                        try {
                            DracoSuppressChangeEvents = true; // evita reabrir el selector por nuestra propia escritura
                            await Excel.run(async (context) => {
                                const sheet = context.workbook.worksheets.getItem(sheetName);
                                await paintDracoFilterCell(context, sheet, addr, payload.filter);
                            });
                            await window.FilterRangeStore.set(rangeName, Object.assign({}, meta, { filter: payload.filter }));
                        } catch (err) {
                            console.error("[Draco] Añadir filtro: error pintando el filtro en la celda:", err);
                        } finally {
                            DracoSuppressChangeEvents = false;
                        }
                        // Si el taskpane está abierto (mismo contexto que
                        // commands.js), refresca la zona "Filtros" para que
                        // la etiqueta bloqueada muestre el nuevo resumen sin
                        // esperar a que el usuario cambie de informe y vuelva.
                        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.loadDesignFromSheet) {
                            try { await TaskPaneApp.loadDesignFromSheet(); } catch (e) { /* taskpane sin informe abierto: se ignora */ }
                        }
                        resolve();
                        return;
                    }

                    if (payload.type === "cancel") {
                        settled = true;
                        closeDialog();
                        resolve();
                    }
                });

                dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                    if (!settled) resolve();
                });
            }
        );
    });
}

/**
 * Aproximación al doble clic (misma limitación de plataforma documentada
 * en handleDracoMemberRecognitionSelection: Office.js no expone un evento
 * de doble clic real sobre una celda): al SELECCIONAR una celda que sea
 * uno de los rangos con nombre creados con "Añadir filtro", se abre
 * directamente el selector de valores.
 */
async function handleDracoFilterRangeSelection(eventArgs) {
    try {
        if (DracoSuppressChangeEvents) return;

        let addr = (eventArgs && eventArgs.address) || "";
        if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
        if (!addr) return;

        let located = null;
        let sheetName = "";

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            const cell = sheet.getRange(addr);
            cell.load(["rowCount", "columnCount"]);
            await context.sync();

            sheetName = sheet.name;
            if (cell.rowCount !== 1 || cell.columnCount !== 1) return; // solo celda única

            located = await locateDracoFilterRangeAtAddress(context, sheetName, addr);
        });

        if (!located) return;
        await openDracoFilterRangePicker(addr, sheetName, located.rangeName, located.meta);
    } catch (e) {
        console.error("[Draco] Error en el selector de 'Añadir filtro' (clic sobre la celda):", e);
    }
}

const DracoFilterRangeHandlerRegisteredSheets = new Set();
let DracoFilterRangeOnAddedRegistered = false;

/**
 * Engancha handleDracoFilterRangeSelection en TODAS las hojas del libro
 * (los rangos de "Añadir filtro" pueden vivir en cualquier hoja, no solo
 * en la hoja de resultados de un informe) y en las hojas que se añadan
 * después. Se puede llamar varias veces sin problema (cada hoja solo se
 * engancha una vez).
 */
async function ensureDracoFilterRangeHandlersRegistered() {
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            sheets.items.forEach(sheet => {
                if (DracoFilterRangeHandlerRegisteredSheets.has(sheet.name)) return;
                sheet.onSelectionChanged.add(handleDracoFilterRangeSelection);
                DracoFilterRangeHandlerRegisteredSheets.add(sheet.name);
            });

            if (!DracoFilterRangeOnAddedRegistered) {
                sheets.onAdded.add(async (e) => {
                    try {
                        await Excel.run(async (ctx) => {
                            const sheet = ctx.workbook.worksheets.getItem(e.worksheetId);
                            sheet.load("name");
                            await ctx.sync();
                            if (!DracoFilterRangeHandlerRegisteredSheets.has(sheet.name)) {
                                sheet.onSelectionChanged.add(handleDracoFilterRangeSelection);
                                DracoFilterRangeHandlerRegisteredSheets.add(sheet.name);
                                await ctx.sync();
                            }
                        });
                    } catch (err) {
                        console.warn("[Draco] No se pudo enganchar 'Añadir filtro' en la hoja nueva:", err);
                    }
                });
                DracoFilterRangeOnAddedRegistered = true;
            }

            await context.sync();
        });
        console.log("[Draco] Listeners de 'Añadir filtro' registrados en todas las hojas.");
    } catch (e) {
        console.warn("[Draco] No se pudieron registrar los listeners de 'Añadir filtro':", e);
    }
}

/* ---------------------------------------------------------------------
 * Registro del clic en cualquier celda de las tablas de Filas de Draco
 * (rangos con nombre Draco_001_Rows, Draco_002_Rows, ...): al hacer clic
 * izquierdo sobre una celda que caiga dentro de alguno de esos rangos, se
 * clasifica el clic en una de 3 zonas según el offsetX (en puntos, desde
 * la esquina izquierda de la celda) -Izquierda / Letra / Derecha-, y
 * además se comprueba si el texto de la celda empieza por el glifo "▾"
 * (nodo de jerarquía EXPANDIDO) para escribir "Jerarquia: SI/NO". Todo
 * el resultado se escribe en A51 de la MISMA hoja donde se ha hecho clic.
 *
 * ACCIÓN REAL: si el clic cae en la zona "Letra" (sobre el icono) Y el
 * texto empieza por "▾" o por "▸" (nodo con indicador +/-, expandido o
 * contraído), se dispara de verdad el expandir/contraer de ese nodo
 * (toggleDracoCollapseAtCell, la misma lógica que antes solo se podía
 * pedir rellenando EDIT_REPORT!T1:V1) y se repinta el informe. El
 * resultado de esa acción también se anota en A51 ("Accion: ...").
 *
 * Usa onSingleClicked (ExcelApi 1.10), igual que la versión anterior
 * sobre A50: se dispara con cada clic (aunque la celda ya estuviera
 * seleccionada) y trae offsetX/offsetY.
 *
 * NOTA: solo se considera "Jerarquia: SI" (para el texto informativo de
 * A51) cuando el texto empieza exactamente por "▾" (expandido); un nodo
 * contraído empieza por "▸" y cuenta como "NO" ahí. Esto es solo el
 * indicador de diagnóstico — para decidir si HAY que disparar la acción
 * de expandir/contraer se admiten los 2 glifos (▾ Y ▸), ya que un nodo
 * contraído también se puede expandir haciendo clic en su icono.
 *
 * IMPORTANTE — la clasificación Izquierda/Letra/Derecha sigue siendo una
 * APROXIMACIÓN, no una medida exacta (ver notas de estimateFirstLetterBoundsPt
 * más abajo).
 * ------------------------------------------------------------------- */

/**
 * Busca, entre todos los nombres del libro que cumplan el patrón
 * "Draco_<n>_Rows", cuál (si alguno) contiene la celda indicada EN LA
 * MISMA HOJA del clic. Devuelve {rangeName, reportId, level} o null si el
 * clic no cayó dentro de ninguna tabla de Filas de Draco.
 * `level` es la posición (1-based) de esa columna dentro del eje de
 * Filas (columna 1 = primer nivel de la jerarquía).
 */
async function findDracoRowsNamedRangeForCell(context, worksheetId, addr) {
    const sheet = context.workbook.worksheets.getItem(worksheetId);
    sheet.load("name");
    const cell = sheet.getRange(addr);
    cell.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    const names = context.workbook.names;
    names.load("items/name");
    await context.sync();

    if (cell.rowCount !== 1 || cell.columnCount !== 1) return null;

    const rowsNames = names.items
        .map(n => n.name)
        .filter(n => /^Draco_\d+_Rows$/i.test(n));
    if (rowsNames.length === 0) return null;

    const candidates = rowsNames.map(name => {
        const item = context.workbook.names.getItem(name);
        const range = item.getRange();
        range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
        const ws = range.worksheet;
        ws.load("name");
        return { name, range, ws };
    });
    await context.sync();

    for (const c of candidates) {
        if (c.ws.name !== sheet.name) continue;
        const within =
            cell.rowIndex >= c.range.rowIndex &&
            cell.rowIndex < c.range.rowIndex + c.range.rowCount &&
            cell.columnIndex >= c.range.columnIndex &&
            cell.columnIndex < c.range.columnIndex + c.range.columnCount;
        if (within) {
            const m = c.name.match(/^Draco_(\d+)_Rows$/i);
            return {
                rangeName: c.name,
                reportId: m ? Number(m[1]) : null,
                level: (cell.columnIndex - c.range.columnIndex) + 1,
                sheetName: sheet.name
            };
        }
    }
    return null;
}

// Contexto de canvas reutilizado para medir texto (evita crear un
// <canvas> nuevo en cada clic).
let _a50MeasureCtx = null;
function getA50MeasureCtx() {
    if (!_a50MeasureCtx) {
        const canvas = document.createElement("canvas");
        _a50MeasureCtx = canvas.getContext("2d");
    }
    return _a50MeasureCtx;
}

// 1 punto = 4/3 px a 96 dpi (estándar CSS/canvas) -> para volver de px a pt:
function pxToPt(px) { return px * (72 / 96); }

/** Construye el string de fuente CSS para el canvas a partir de los
 *  datos de Excel.Font. Entrecomillamos el nombre por si tiene espacios
 *  (p. ej. "Segoe UI"). */
function buildA50CanvasFont(fontSizePt, fontName, bold, italic) {
    const weight = bold ? "bold " : "";
    const style = italic ? "italic " : "";
    const size = fontSizePt || 11;
    const name = fontName || "Calibri";
    return `${style}${weight}${size}pt "${name}"`;
}

/** Pequeño margen izquierdo interno que Excel deja siempre dentro de la
 *  celda antes de empezar a pintar el texto, incluso con indentLevel 0.
 *  Es un valor aproximado (no hay API que lo exponga). */
const A50_CELL_LEFT_PADDING_PT = 1;

/**
 * Devuelve el rango [start, end] en puntos (relativo a la esquina
 * izquierda de la celda) donde se estima que cae la "primera letra",
 * usando el ancho de una "D" mayúscula como referencia.
 *
 * OJO: el ancho de la "D" se mide con la fuente DE LA CELDA (A50), pero
 * el ancho del indentado se mide con la fuente del estilo "Normal" del
 * libro — es la que usa Excel internamente como unidad de medida del
 * indentado (y también del ancho de columna en "caracteres"), NO la
 * fuente concreta que tenga puesta la celda. Con fuentes condensadas
 * (p.ej. "Aptos Narrow") calcular el indentado con la fuente de la
 * celda da un resultado muy desplazado a la derecha.
 *
 * CALIBRACIÓN (revisada): la regla "3 espacios por nivel" quedaba
 * desplazada hacia la derecha frente a clics reales sobre el icono
 * ▾/▸ (con indent creciente el desfase se acumulaba y el clic pasaba a
 * clasificarse como "Izquierda" en vez de "Letra"). Excel define el
 * "carácter estándar" (el mismo que usa para el ancho de columna en
 * "caracteres") como el ancho del dígito "0" en la fuente Normal del
 * libro, y cada nivel de IndentLevel equivale a 1 carácter estándar
 * (NO a 3 espacios). Con datos reales de clic se comprobó que esta
 * unidad encaja mucho mejor: indent1 ≈ +5.2pt e indent2 ≈ +5.6pt sobre
 * el nivel anterior, frente al ancho del dígito "0" en Aptos Narrow
 * 11pt (≈5.5pt aprox.) — mientras que 3 espacios daban ≈8.25pt, muy
 * por encima de lo observado.
 */
function estimateFirstLetterBoundsPt(cellFontName, cellFontSizePt, cellBold, cellItalic, indentLevel, normalFontName, normalFontSizePt) {
    const ctx = getA50MeasureCtx();

    // Ancho de la "D" -> con la fuente de la celda.
    ctx.font = buildA50CanvasFont(cellFontSizePt, cellFontName, cellBold, cellItalic);
    const dWidthPt = pxToPt(ctx.measureText("D").width);

    // Ancho del "carácter estándar" de Excel para el indentado -> ancho
    // del dígito "0" con la fuente "Normal" del libro (sin negrita/
    // cursiva, es el estilo base). Es la misma unidad que usa Excel
    // para el ancho de columna en "caracteres".
    ctx.font = buildA50CanvasFont(normalFontSizePt, normalFontName, false, false);
    const stdCharWidthPt = pxToPt(ctx.measureText("0").width);

    // 1 nivel de indentado = 1 carácter estándar de la fuente "Normal".
    const indentWidthPt = (indentLevel || 0) * stdCharWidthPt;

    const start = A50_CELL_LEFT_PADDING_PT + indentWidthPt;
    const end = start + dWidthPt;
    return { start, end, dWidthPt, indentWidthPt, stdCharWidthPt };
}

// Detección "manual" de doble clic: Office.js NO expone un evento nativo
// de doble clic sobre una celda (a diferencia de Workbook_SheetBeforeDoubleClick
// en VBA) — Excel.Worksheet solo ofrece onSingleClicked, onChanged y
// onSelectionChanged (ver comentarios de handleDracoMemberRecognitionChanged
// más arriba, donde ya se documentaba esta misma limitación).
//
// Lo simulamos guardando, por cada hoja, sobre qué celda y en qué instante
// cayó el ÚLTIMO clic simple. Si el siguiente clic llega en menos de
// DRACO_DOUBLE_CLICK_MS Y es sobre la MISMA celda, lo contamos como doble
// clic. Tras detectarlo, se borra el registro (para que un 3er clic rápido
// no cuente como "doble clic" otra vez contra el 2º).
const DracoLastClickByWorksheet = new Map(); // worksheetId -> { addr, time }
const DRACO_DOUBLE_CLICK_MS = 500;

// Formatea un timestamp (ms epoch) como HH:MM:SS.mmm, para poder leer a
// simple vista en la hoja el instante de cada clic (trazas A53 en
// adelante de handleDracoRowsSingleClick).
function formatDracoClickTime(ms) {
    if (!ms && ms !== 0) return "";
    const d = new Date(ms);
    const pad = (n, len) => String(n).padStart(len || 2, "0");
    return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}.${pad(d.getMilliseconds(), 3)}`;
}

// Devuelve toda la info del par de clics (no solo el booleano) para
// poder trazarla: hora del clic anterior (click1), hora del actual
// (click2), diferencia en ms y si cuenta como doble clic.
function detectDracoDoubleClick(worksheetId, addr) {
    const now = Date.now();
    const last = DracoLastClickByWorksheet.get(worksheetId);
    const isDouble = !!(last && last.addr === addr && (now - last.time) < DRACO_DOUBLE_CLICK_MS);
    const info = {
        isDouble,
        click1Time: last ? last.time : null,
        click1Addr: last ? last.addr : null,
        click2Time: now,
        click2Addr: addr,
        deltaMs: last ? (now - last.time) : null
    };
    if (isDouble) {
        DracoLastClickByWorksheet.delete(worksheetId);
    } else {
        DracoLastClickByWorksheet.set(worksheetId, { addr, time: now });
    }
    return info;
}

async function handleDracoRowsSingleClick(eventArgs) {
    try {
        let addr = (eventArgs && eventArgs.address) || "";
        if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
        addr = addr.replace(/\$/g, "").toUpperCase();
        if (!addr) return;

        const offsetX = eventArgs.offsetX;
        const offsetY = eventArgs.offsetY;
        const clickInfo = detectDracoDoubleClick(eventArgs.worksheetId, addr);
        const isDoubleClick = clickInfo.isDouble;

        await Excel.run(async (context) => {
            const located = await findDracoRowsNamedRangeForCell(context, eventArgs.worksheetId, addr);
            if (!located) return; // el clic no cayó dentro de ninguna tabla de Filas de Draco

            const sheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
            const cell = sheet.getRange(addr);
            cell.load([
                "values",
                "format/font/name",
                "format/font/size",
                "format/font/bold",
                "format/font/italic",
                "format/indentLevel",
                "format/columnWidth"
            ]);

            // Fuente del estilo "Normal" del libro: unidad de medida del indentado.
            const normalStyle = context.workbook.styles.getItem("Normal");
            normalStyle.load(["font/name", "font/size"]);

            await context.sync();

            const cellText = String((cell.values && cell.values[0] && cell.values[0][0]) || "");
            // Solo cuenta como "SI" si empieza EXACTAMENTE por el glifo de
            // expandido ("▾"); un nodo contraído ("▸") cuenta como "NO".
            const jerarquiaTexto = cellText.indexOf("▾") === 0 ? "SI" : "NO";

            const bounds = estimateFirstLetterBoundsPt(
                cell.format.font.name,
                cell.format.font.size,
                cell.format.font.bold,
                cell.format.font.italic,
                cell.format.indentLevel,
                normalStyle.font.name,
                normalStyle.font.size
            );

            let resultado;
            if (offsetX < bounds.start) {
                resultado = "Izquierda";
            } else if (offsetX > bounds.end) {
                resultado = "Derecha";
            } else {
                resultado = "Letra";
            }

            // Además de "SI" (expandido), un nodo contraído ("▸") también
            // tiene un icono con el que se puede interactuar: cualquiera de
            // los 2 glifos cuenta como "tiene indicador +/- clicable".
            const hasHierarchyGlyph = cellText.indexOf("▾") === 0 || cellText.indexOf("▸") === 0;

            let accionTexto = "";
            // Variables de traza (ver bloque A53:A58 más abajo): reflejan
            // por qué rama pasó este clic, aunque no sea doble clic o no
            // llegue a entrar en la comprobación de Estáticos.
            let staticCheckEntered = false;
            let rowsStaticTrace = null;
            let colsStaticTrace = null;
            let fieldLocatedTrace = null;

            if (resultado === "Letra" && hasHierarchyGlyph) {
                // El clic ha caído sobre el icono ▾/▸: expandir/contraer el
                // nodo, igual que si se hubiera rellenado EDIT_REPORT!T1:V1.
                const toggled = await toggleDracoCollapseAtCell(context, located.sheetName, addr);
                accionTexto = toggled
                    ? " | Accion: expandir/contraer ejecutado"
                    : " | Accion: sin indicador +/- valido en esa celda";
            } else if (isDoubleClick) {
                // Doble clic "real" fuera del icono +/-: equivalente a
                // Workbook_SheetBeforeDoubleClick + GetDimAttrRows/
                // GetDimAttrCols + Helpvalue del VBA. Allí Helpvalue
                // escribía dimname/attr/valor/dirección en
                // EDIT_REPORT!A1:A5 y algo aguas abajo leía esas celdas
                // para abrir el buscador de miembros. Aquí se hace lo
                // mismo pero sin pasar por ninguna celda física: se
                // localiza el campo con locateDracoAxisField (la misma
                // función que ya usa el "reconocimiento de miembros" de
                // onChanged/onSelectionChanged) y se abre el picker
                // directamente con el resultado en memoria.
                staticCheckEntered = true;
                const reportId = activeReportIdOrNull();
                const editReportGrid = await getEditReportGrid(context, reportId);
                const rowsStatic = String(cellValue(editReportGrid, 12, 8)).trim().toUpperCase() === "X";
                const colsStatic = String(cellValue(editReportGrid, 12, 14)).trim().toUpperCase() === "X";
                rowsStaticTrace = rowsStatic;
                colsStaticTrace = colsStatic;

                // Igual que en VBA (RowsStatic Y ColsStatic ambos "X"):
                // el buscador de doble clic solo actúa si TODO el diseño
                // del informe (filas y columnas) está marcado como
                // Estático.
                if (rowsStatic && colsStatic) {
                    const fieldLocated = await locateDracoAxisField(context, addr);
                    fieldLocatedTrace = fieldLocated;
                    if (fieldLocated) {
                        accionTexto = " | Accion: buscador de miembros abierto (doble clic)";
                        await openMemberRecognitionPicker(addr, fieldLocated, cellText);
                    } else {
                        accionTexto = " | Accion: doble clic fuera de Filas/Columnas del informe";
                    }
                } else {
                    accionTexto = " | Accion: doble clic ignorado (informe no 100% Estatico)";
                }
            }

            console.log(
                `[Draco] Clic en ${located.rangeName} (nivel ${located.level}) -> offsetX=${offsetX.toFixed(2)}pt | ` +
                `letra estimada entre ${bounds.start.toFixed(2)}pt y ${bounds.end.toFixed(2)}pt -> ${resultado} | Jerarquia: ${jerarquiaTexto}${accionTexto}`
            );

            const target = sheet.getRange("A51");
            target.values = [[`${resultado} | Jerarquia: ${jerarquiaTexto}${accionTexto}`]];

            // A52: volcado de diagnóstico para poder calibrar el cálculo.
            const debugRange = sheet.getRange("A52");
            debugRange.values = [[
                `Rango=${located.rangeName} Nivel=${located.level} | X=${offsetX.toFixed(2)} Y=${offsetY.toFixed(2)} | ` +
                `letra=[${bounds.start.toFixed(2)},${bounds.end.toFixed(2)}] | indent=${cell.format.indentLevel} ` +
                `(carácter=${bounds.stdCharWidthPt.toFixed(2)}pt) | colW=${cell.format.columnWidth.toFixed(2)}pt | ` +
                `celda=${cell.format.font.name} ${cell.format.font.size}pt | Normal=${normalStyle.font.name} ${normalStyle.font.size}pt | texto="${cellText}"`
            ]];

            // A53:A58 — traza detallada de cada clic, para depurar por qué
            // no se abre el buscador de miembros. Se reescribe en TODOS
            // los clics (simples y dobles), así que siempre se ve el
            // último estado, incluyendo el clic1 (anterior) y clic2
            // (actual) que se comparan para decidir si hay doble clic.
            if (isDoubleClick) {
                console.log(`[Draco] Doble clic detectado en ${addr} (${located.rangeName}).`);
            }

            const staticInfo = staticCheckEntered
                ? `SI | RowsStatic=${rowsStaticTrace ? "X" : "(no)"} ColsStatic=${colsStaticTrace ? "X" : "(no)"}`
                : "NO (no fue doble clic fuera del icono +/-)";

            const campoInfo = fieldLocatedTrace
                ? `${fieldLocatedTrace.axis} nivel ${fieldLocatedTrace.level} -> Dim=${fieldLocatedTrace.dim} Attr=${fieldLocatedTrace.attr || "(sin attr)"}`
                : (staticCheckEntered && rowsStaticTrace && colsStaticTrace
                    ? "(no encontrado por locateDracoAxisField)"
                    : "(no evaluado)");

            const traceRows = [
                [`Click1 (anterior)=${clickInfo.click1Time ? formatDracoClickTime(clickInfo.click1Time) + " @" + clickInfo.click1Addr : "(ninguno todavía)"}`],
                [`Click2 (actual)=${formatDracoClickTime(clickInfo.click2Time)} @${clickInfo.click2Addr} | Delta=${clickInfo.deltaMs === null ? "-" : clickInfo.deltaMs + "ms"} (umbral=${DRACO_DOUBLE_CLICK_MS}ms)`],
                [`EsDobleClick=${isDoubleClick ? "SI" : "NO"} | SobreIconoJerarquia=${hasHierarchyGlyph ? "SI" : "NO"} (resultado=${resultado})`],
                [`EntraPorEstaticos=${staticInfo}`],
                [`CampoLocalizado=${campoInfo}`],
                [`AccionFinal=${accionTexto ? accionTexto.replace(/^ \| Accion: /, "") : "(ninguna)"}`]
            ];
            const traceRange = sheet.getRange("A53:A58");
            traceRange.values = traceRows;

            await context.sync();
        });
    } catch (e) {
        console.error("[Draco] Error registrando el clic en Draco_*_Rows:", e);
    }
}

const DracoRowsClickHandlerRegisteredSheets = new Set();
let DracoRowsClickOnAddedRegistered = false;

/**
 * Engancha handleDracoRowsSingleClick en TODAS las hojas del libro (los
 * rangos Draco_XXX_Rows pueden estar en cualquiera) y en las hojas que se
 * añadan después. Se puede llamar varias veces sin problema (cada hoja
 * solo se engancha una vez).
 */
async function ensureDracoRowsClickLoggerRegistered() {
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            sheets.items.forEach(sheet => {
                if (DracoRowsClickHandlerRegisteredSheets.has(sheet.name)) return;
                sheet.onSingleClicked.add(handleDracoRowsSingleClick);
                DracoRowsClickHandlerRegisteredSheets.add(sheet.name);
            });

            if (!DracoRowsClickOnAddedRegistered) {
                sheets.onAdded.add(async (e) => {
                    try {
                        await Excel.run(async (ctx) => {
                            const sheet = ctx.workbook.worksheets.getItem(e.worksheetId);
                            sheet.load("name");
                            await ctx.sync();
                            if (!DracoRowsClickHandlerRegisteredSheets.has(sheet.name)) {
                                sheet.onSingleClicked.add(handleDracoRowsSingleClick);
                                DracoRowsClickHandlerRegisteredSheets.add(sheet.name);
                                await ctx.sync();
                            }
                        });
                    } catch (err) {
                        console.warn("[Draco] No se pudo enganchar el listener de Draco_*_Rows en la hoja nueva:", err);
                    }
                });
                DracoRowsClickOnAddedRegistered = true;
            }

            await context.sync();
        });
        console.log("[Draco] Listener de clic en rangos Draco_*_Rows registrado en todas las hojas.");
    } catch (e) {
        console.warn("[Draco] No se pudo registrar el listener de clic en Draco_*_Rows (¿host sin soporte de ExcelApi 1.10?):", e);
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
                                const resultSheetName = await getDracoResultSheetName(context);
                                const sheet = context.workbook.worksheets.getItem(resultSheetName);
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

        console.log("[Draco] handleEditReportMemberPickerRequest: onChanged en EDIT_REPORT, address =", eventArgs.address);

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

            console.log("[Draco] handleEditReportMemberPickerRequest: A1:A5 =", { dimension, attribute, searchValue, addr, request });

            if (request !== "X") {
                console.log("[Draco] handleEditReportMemberPickerRequest: A5 no es 'X' (valor real: '" + request + "'), se ignora.");
                return;
            }

            // Consumimos la petición
            sheet.getRange("A5").clear(Excel.ClearApplyTo.contents);
            await context.sync();

            if (!dimension || !attribute || !addr) {
                console.warn("[Draco] handleEditReportMemberPickerRequest: falta dimension/attribute/addr, no se abre el picker.", { dimension, attribute, addr });
                return;
            }

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



/**
 * Alterna (o fuerza) el estado colapsado/expandido del nodo de jerarquía
 * cuya celda de indicador +/- (el icono ▾/▸ fusionado en el texto) está en
 * targetAddr, dentro de la hoja sheetName. Es la lógica común para las 2
 * formas de disparar un expandir/contraer:
 *   1) Rellenando EDIT_REPORT!T1:V1 (ver handleDracoEditReportExpandCollapseRequest).
 *   2) Haciendo clic directamente sobre el icono ▾/▸ en una celda de
 *      Draco_XXX_Rows (ver handleDracoRowsSingleClick).
 *
 * flag: "C"/"CONTRAER" fuerza contraído, "E"/"EXPANDIR" fuerza expandido,
 * cualquier otro valor (incluido undefined) alterna el estado actual.
 *
 * Devuelve true si se ha encontrado un indicador +/- válido en esa celda y
 * se ha alternado/forzado su estado (y repintado el informe); false si la
 * celda no tiene un indicador asociado (p.ej. eje Estático, celda sin
 * jerarquía, o la hoja/informe no existen).
 */
async function toggleDracoCollapseAtCell(context, sheetName, targetAddr, flag) {
    const reportId = reportIdForResultSheet(sheetName);
    const indicatorMap = getDracoIndicatorMap(reportId);
    if (indicatorMap.size === 0) return false;

    const resultSheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    resultSheet.load("isNullObject");
    await context.sync();
    if (resultSheet.isNullObject) return false;

    const targetRange = resultSheet.getRange(targetAddr);
    targetRange.load(["rowCount", "columnCount", "rowIndex", "columnIndex", "formulas"]);
    await context.sync();

    if (targetRange.rowCount !== 1 || targetRange.columnCount !== 1) return false;

    // Una celda de eje Estático es una fórmula EPM_VALUE (sin glifo +/-).
    const formula = targetRange.formulas[0][0];
    const isStaticFormula = typeof formula === "string" && /^=\s*EPM_VALUE\s*\(/i.test(formula);

    const key = (targetRange.rowIndex + 1) + "_" + (targetRange.columnIndex + 1);
    const meta = indicatorMap.get(key);
    if (!meta || isStaticFormula) return false;

    const st = getDracoCollapseState(reportId)[meta.axis];
    if (flag === "C" || flag === "CONTRAER") {
        st.collapsed.add(meta.nodeKey);
    } else if (flag === "E" || flag === "EXPANDIR") {
        st.collapsed.delete(meta.nodeKey);
    } else if (st.collapsed.has(meta.nodeKey)) {
        st.collapsed.delete(meta.nodeKey);
    } else {
        st.collapsed.add(meta.nodeKey);
    }

    const lastJson = DracoLastJsonByReport.get(dracoStateKey(reportId));
    if (lastJson) {
        await jsonTo3Matrices(context, lastJson, reportId);
    }
    return true;
}

// Celdas de control de EDIT_REPORT para pedir un expandir/contraer sin
// pasar por onSelectionChanged (ver handleDracoEditReportExpandCollapseRequest).
// No se usan D15/E15/F15 (como se planteó al principio) porque esa zona
// (fila >=15, columnas C:F) está reservada al primer filtro del diseño
// (ver isDesignOwnedCell y el comentario "C15:F.. -> filtros" en
// reportStore.js): aunque el JS ya no las lee para los filtros (gana
// siempre el JSON), reutilizarlas aquí podría chocar con el XLAM si en
// algún momento vuelve a escribir físicamente ahí. T1/U1/V1 están fuera
// de cualquier zona ya usada (picker A1:A5, B1, D1/D4/D5/D6/E1/G1, X1/Y1,
// H10/N10/H12/N12 y toda la fila >=15 de C a S).
const DRACO_EXPAND_COLLAPSE_SHEET_CELL = "T1"; // nombre de la pestaña de resultados
const DRACO_EXPAND_COLLAPSE_TARGET_CELL = "U1"; // celda con el indicador +/- a tocar
const DRACO_EXPAND_COLLAPSE_FLAG_CELL = "V1"; // "E"/"EXPANDIR" o "C"/"CONTRAER" (opcional)

/**
 * Reemplaza al antiguo handleDracoSelectionChanged (expandía/contraía con
 * solo seleccionar la celda del indicador +/-). Ahora hace falta rellenar
 * explícitamente, en EDIT_REPORT:
 *   T1 -> nombre de la pestaña de resultados (p.ej. "CSV_RESULT")
 *   U1 -> celda de esa pestaña donde está el indicador +/- a expandir o
 *         contraer (p.ej. "B7")
 *   V1 -> opcional: "E"/"EXPANDIR" o "C"/"CONTRAER". Si se omite o trae un
 *         valor no reconocido, se alterna (mismo comportamiento que el
 *         clic de antes).
 *
 * Se dispara con onChanged de EDIT_REPORT al tocar T1, U1 o V1 (o un
 * rango que las incluya, p.ej. si se pegan las 3 a la vez), pero solo
 * actúa si T1 Y U1 tienen valor. Al terminar, limpia T1:V1 (con
 * DracoSuppressChangeEvents activo, para no reaccionar a nuestra propia
 * escritura) y así la misma petición puede repetirse más adelante sin
 * quedarse "pegada" en las celdas.
 */
async function handleDracoEditReportExpandCollapseRequest(eventArgs) {
    try {
        if (DracoSuppressChangeEvents) return;
        if (!eventArgs || !eventArgs.address) return;

        let addr = String(eventArgs.address);
        if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
        if (!addr) return;

        const touchedRange = parseAddressRange(addr);
        const controlCells = [DRACO_EXPAND_COLLAPSE_SHEET_CELL, DRACO_EXPAND_COLLAPSE_TARGET_CELL, DRACO_EXPAND_COLLAPSE_FLAG_CELL]
            .map(parseAddress);
        const touchesControlZone = controlCells.some(p =>
            p.row >= touchedRange.r1 && p.row <= touchedRange.r2 &&
            p.col >= touchedRange.c1 && p.col <= touchedRange.c2
        );
        if (!touchesControlZone) return;

        await Excel.run(async (context) => {
            const editReport = context.workbook.worksheets.getItem("EDIT_REPORT");
            const ctrl = editReport.getRange(
                DRACO_EXPAND_COLLAPSE_SHEET_CELL + ":" + DRACO_EXPAND_COLLAPSE_FLAG_CELL
            );
            ctrl.load("values");
            await context.sync();

            const sheetName = String(ctrl.values[0][0] || "").trim();
            const targetAddr = String(ctrl.values[0][1] || "").trim();
            const flag = String(ctrl.values[0][2] || "").trim().toUpperCase();

            if (!sheetName || !targetAddr) return; // hacen falta las 2

            console.log("[Draco] Petición de expandir/contraer desde EDIT_REPORT:", { sheetName, targetAddr, flag });

            const toggled = await toggleDracoCollapseAtCell(context, sheetName, targetAddr, flag);
            if (!toggled) {
                console.warn("[Draco] EDIT_REPORT!U1 no apunta a un indicador +/- válido:", sheetName, targetAddr);
            }

            // Consumimos la petición (igual que A5 con el Member Picker) para
            // que T1/U1 rellenas no sigan disparando el handler indefinidamente.
            DracoSuppressChangeEvents = true;
            editReport.getRange(
                DRACO_EXPAND_COLLAPSE_SHEET_CELL + ":" + DRACO_EXPAND_COLLAPSE_FLAG_CELL
            ).clear(Excel.ClearApplyTo.contents);
            await context.sync();
        });
    } catch (e) {
        console.error("[Draco] Error al expandir/contraer desde EDIT_REPORT!T1:V1:", e);
    } finally {
        DracoSuppressChangeEvents = false;
    }
}

/**
 * Registra (una sola vez) los listeners de EDIT_REPORT que no dependen de
 * que exista ninguna hoja de resultados: el picker de A5 (Member Picker)
 * y la petición de expandir/contraer de T1:V1. Es INDEPENDIENTE de
 * registerDracoSelectionHandler/DracoHandlerRegistered a propósito: antes
 * el picker estaba dentro de esa misma función y compartía su flag, lo
 * que significaba que si CSV_RESULT (la hoja de resultados) todavía no
 * existía —p.ej. sesión recién abierta, sin haber pulsado nunca
 * "Actualizar"— el picker de A5 NO se registraba, aunque EDIT_REPORT sí
 * existiera y el usuario ya estuviera rellenando A1/A2/A4/A5. Con esta
 * función aparte, ensureDracoHandlersRegistered puede engancharla sin
 * depender de que exista CSV_RESULT.
 */
async function registerEditReportPickerHandler(context) {
    if (DracoEditReportHandlerRegistered) return;

    const editReport = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
    editReport.load("isNullObject");
    await context.sync();

    if (editReport.isNullObject) {
        console.log("[Draco] registerEditReportPickerHandler: EDIT_REPORT no existe todavía.");
        return;
    }

    editReport.onChanged.add(handleEditReportMemberPickerRequest);
    editReport.onChanged.add(handleDracoEditReportExpandCollapseRequest);
    await context.sync();

    DracoEditReportHandlerRegistered = true;
    console.log("[Draco] Listeners de EDIT_REPORT!A5 (Member Picker) y T1:V1 (expandir/contraer) registrados.");
}

// sheetName es el nombre de la hoja de resultados donde vive `sheet`: cada
// informe puede pintarse en una hoja distinta (ver resultSheetNameFromGrid/
// getDracoResultSheetName), así que los listeners de clic (+/-) y
// reconocimiento de miembros se enganchan POR HOJA, no una sola vez para
// todo el libro — si no, un segundo informe en otra hoja se quedaba sin
// clic interactivo (nunca se registraba nada ahí).
async function registerDracoSelectionHandler(context, sheet, sheetName) {
    if (DracoHandlerRegisteredSheets.has(sheetName)) return;

    sheet.onSelectionChanged.add(handleDracoMemberRecognitionSelection);
    sheet.onSelectionChanged.add(handleDracoRibbonLabelSelection);
    sheet.onChanged.add(handleDracoMemberRecognitionChanged);

    // NUEVO: petición de apertura del Member Picker desde EDIT_REPORT
    await registerEditReportPickerHandler(context);

    await context.sync();
    DracoHandlerRegisteredSheets.add(sheetName);
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
    // El picker de EDIT_REPORT!A5 no depende de que exista CSV_RESULT ni de
    // que se haya refrescado nunca: se intenta registrar siempre, aparte.
    try {
        await Excel.run(async (context) => {
            await registerEditReportPickerHandler(context);
        });
    } catch (e) {
        console.warn("[Draco] No se pudo registrar el listener de EDIT_REPORT!A5 de forma temprana:", e);
    }

    try {
        await Excel.run(async (context) => {
            const resultSheetName = await getDracoResultSheetName(context);
            if (DracoHandlerRegisteredSheets.has(resultSheetName)) return;
            const sheet = context.workbook.worksheets.getItemOrNullObject(resultSheetName);
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) {
                console.log("[Draco] ensureDracoHandlersRegistered: " + resultSheetName + " no existe todavía (sin refresco previo).");
                return;
            }
            await registerDracoSelectionHandler(context, sheet, resultSheetName);
            console.log("[Draco] Listeners de " + resultSheetName + " registrados de forma temprana (sin necesidad de refrescar).");
        });
    } catch (e) {
        console.warn("[Draco] No se pudieron registrar los listeners de forma temprana:", e);
    }
}

/**
 * Recorre físicamente las filas de EDIT_REPORT de un eje (H/I/J para
 * Filas, N/O/P para Columnas) empezando en la fila 15, IGUAL que
 * loadRows/loadColumns, pero sin descartar las filas "MEASURE": las
 * incluye en la lista, en el orden físico real, marcadas con
 * isMeasure=true. Esto es necesario porque loadRows/loadColumns cuentan
 * MEASURE aparte (ReportState.MeasureCount) sin que quede constancia de
 * EN QUÉ POSICIÓN del eje iba esa fila MEASURE, y jsonTo3Matrices
 * necesita esa posición para pintar correctamente cuando MEASURE no es
 * la última fila del eje (ver computeAxisPaintPlan).
 */
function buildDracoAxisLevels(editReportGrid, dimCol, attrCol, hierCol) {
    const levels = [];
    let R = 15;
    while (String(cellValue(editReportGrid, R, dimCol)).trim() !== "") {
        const dim = String(cellValue(editReportGrid, R, dimCol)).trim();
        levels.push({
            isMeasure: dim.toUpperCase() === "MEASURE",
            dim,
            attr: String(cellValue(editReportGrid, R, attrCol)).trim(),
            flag: Number(cellValue(editReportGrid, R, hierCol))
        });
        R++;
    }
    return levels;
}

/**
 * Firma de un eje (para expandir/contraer) calculada a partir de los
 * levels ya leídos (en vez de releer EDIT_REPORT con un `count` que
 * excluía las filas MEASURE, lo que producía firmas que no reflejaban
 * la fila MEASURE ni su posición).
 */
function computeDracoAxisSignatureFromLevels(levels) {
    return levels
        .map(l => (l.isMeasure ? "MEASURE" : String(l.dim).toUpperCase()) + "." + String(l.attr).toUpperCase() + ":" + (l.isMeasure ? "" : String(l.flag)))
        .join("|");
}

/**
 * A partir de los levels físicos de un eje (dimensiones reales + filas
 * MEASURE intercaladas), calcula todo lo que jsonTo3Matrices necesita
 * para pintar:
 *  - flagsReal / fieldsReal / iauxReal: iguales a los antiguos
 *    flagsFilas/fieldsFilas/flagsColumnas/fieldsColumnas pero indexados
 *    1..N SOLO sobre las dimensiones reales (excluyendo MEASURE), que es
 *    el mismo orden en que aparecen los valores V[i] del FACT/dict.
 *  - flag1OrdinalToIaux: para remapear el `field` que devuelve
 *    filterAndCompactDracoAxis (que cuenta solo dimensiones reales con
 *    NIVEL=1) a la columna/fila física real ya pintada (iAux), que SÍ
 *    incluye el hueco de las filas MEASURE.
 *  - measureLevels: una entrada por cada fila MEASURE del eje, con la
 *    columna/fila física (iAux) que le corresponde y la etiqueta a
 *    pintar (el nombre de medida, ej. "IMPORTE").
 *  - totalIaux: nº total de columnas/filas físicas pintadas en este eje
 *    (dimensiones reales de NIVEL=1 + filas MEASURE), para
 *    Draco_001_Rows/Cols.
 */
function computeAxisPaintPlan(levels) {
    const flagsReal = [];
    const fieldsReal = [];
    const iauxReal = [];
    const flag1OrdinalToIaux = [];
    // Una entrada por medida, TODAS compartiendo el mismo iAux (una única
    // fila/columna física "extra" para el grupo de medidas del eje, no una
    // por medida): measureLevels[k].ordinal (0..N-1) es la posición de esa
    // medida DENTRO del grupo, usada luego para repartirla en su propio
    // hueco físico perpendicular al eje (ver measureCount en jsonTo3Matrices).
    const measureLevels = [];

    let iAux = 0;
    let dimIdx = 0;
    let ordinal = 0;
    let measureIaux = null;

    for (const level of levels) {
        if (level.isMeasure) {
            if (measureIaux === null) {
                iAux++;
                measureIaux = iAux;
            }
            measureLevels.push({ iAux: measureIaux, ordinal: measureLevels.length, label: (level.attr || level.dim || "").trim() || "MEASURE" });
            continue;
        }
        dimIdx++;
        const flag = level.flag;
        if (flag === 1) {
            iAux++;
            ordinal++;
            flag1OrdinalToIaux[ordinal] = iAux;
        }
        flagsReal[dimIdx] = flag;
        fieldsReal[dimIdx] = { dim: level.dim, attr: level.attr };
        iauxReal[dimIdx] = iAux;
    }

    return { flagsReal, fieldsReal, iauxReal, flag1OrdinalToIaux, measureLevels, totalIaux: iAux, dimCount: dimIdx };
}

/**
 * JSON_To_3_Matrices: traducción literal, incluyendo el mismo cruce
 * "invertido" H/N que usa el resto del módulo (total_dim_filas=ColumnCount
 * leído de columnas H/I/J; total_dim_cols=RowCount leído de columnas N/O/P).
 */
// reportId es opcional: por defecto, el informe activo (mismo
// comportamiento que antes de que existiera "Refrescar todos"). Se pasa
// explícito cuando se pinta UN informe concreto que no tiene por qué ser
// el activo en el taskpane (ver actualizarTodosCore).
async function jsonTo3Matrices(context, json, reportId) {
    DracoSuppressChangeEvents = true; // evita que el reconocimiento de miembros reaccione a este pintado
    try {
        return await jsonTo3MatricesCore(context, json, reportId);
    } finally {
        DracoSuppressChangeEvents = false;
    }
}

async function jsonTo3MatricesCore(context, json, reportIdOverride) {
    const reportId = reportIdOverride !== undefined ? reportIdOverride : activeReportIdOrNull();
    DracoLastJsonByReport.set(dracoStateKey(reportId), json); // cache para poder repintar en un toggle +/- sin re-consultar BigQuery

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
    const editReportGrid = await getEditReportGrid(context, reportId);

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
    // Se lee el eje completo (incluidas las filas MEASURE, si las hay) con
    // buildDracoAxisLevels/computeAxisPaintPlan: iterar 1..totalDimFilas y
    // leer directamente la fila (i+14) — como se hacía antes — asume que
    // TODAS las filas físicas del eje son dimensiones reales; en cuanto
    // aparece una fila MEASURE intercalada, esa asunción deja de ser
    // cierta y desalinea el resto de dimensiones (y de paso nunca llega a
    // pintarse la propia etiqueta de la medida). Por eso ahora se recorre
    // el eje físico completo una vez (con las filas MEASURE incluidas) y
    // se deriva de ahí tanto la correspondencia {dim, attr, NIVEL} de cada
    // dimensión real como la columna/fila física de cada etiqueta MEASURE.
    const levelsFilas = buildDracoAxisLevels(editReportGrid, 8, 9, 10);      // H/I/J
    const levelsColumnas = buildDracoAxisLevels(editReportGrid, 14, 15, 16); // N/O/P

    const planFilas = computeAxisPaintPlan(levelsFilas);
    const planColumnas = computeAxisPaintPlan(levelsColumnas);

    // ---- Varias medidas en el MISMO eje (Σ Medidas del taskpane): el
    // taskpane garantiza que todas las medidas están en un único eje
    // (nunca repartidas entre Filas y Columnas), así que como mucho uno de
    // los dos "measureLevels" de abajo tiene contenido. Ese eje se pinta
    // measureCount veces más ancho/alto (una fila/columna física física
    // por medida), en vez de una sola celda con el valor de la última
    // medida (como pasaba antes de este cambio).
    const measureCount = Math.max(1, ReportState.Measures.length);
    const measuresOnRowsAxis = planFilas.measureLevels.length > 0;      // eje "Filas" físico (H/I/J)
    const measuresOnColsAxis = planColumnas.measureLevels.length > 0;   // eje "Columnas" físico (N/O/P)
    // Índice (0-based) del primer campo de medida dentro de cada fila del
    // FACT: tras ROW_ID, COLUMN_ID y las dimensiones de ambos ejes, en el
    // mismo orden en que buildSelect/buildFinalSelect las escriben.
    const measureFieldBase = 2 + ReportState.ColumnCount + ReportState.RowCount;
    // logicalId 1..N -> posición física 1..N*measureCount (measureCount
    // huecos consecutivos por cada logicalId, uno por medida); si el eje no
    // lleva medidas, o measureCount es 1, devuelve el mismo logicalId.
    const explodeForMeasures = (logicalId, isMeasureAxis, ordinal) =>
        isMeasureAxis ? (logicalId - 1) * measureCount + ordinal + 1 : logicalId;

    const flagsFilas = planFilas.flagsReal;
    const fieldsFilas = planFilas.fieldsReal; // {dim, attr} por posición — solo hace falta si el eje es Estático (fórmulas EPM_VALUE)

    const flagsColumnas = planColumnas.flagsReal;
    const fieldsColumnas = planColumnas.fieldsReal;

    // ---- Expandir/Contraer: si el eje (misma lista de campos) no ha
    // cambiado desde el último refresco, se respeta qué nodos estaban
    // contraídos; si ha cambiado, se resetea (todo expandido) para ESE eje.
    const rowsSignature = computeDracoAxisSignatureFromLevels(levelsFilas);
    const colsSignature = computeDracoAxisSignatureFromLevels(levelsColumnas);
    resetDracoCollapseIfAxisChanged("rows", rowsSignature, reportId);
    resetDracoCollapseIfAxisChanged("cols", colsSignature, reportId);
    const dracoCollapseState = getDracoCollapseState(reportId);

    // "Suprimir ceros en filas/columnas" (propiedades del informe): se quita
    // del diccionario, ANTES de compactar, cualquier ROW_ID/COLUMN_ID cuyo
    // valor sea siempre cero (o vacío) en todo el FACT.
    if (reportProps.suppressZeroRows || reportProps.suppressZeroCols) {
        const rowsWithValue = new Set();
        const colsWithValue = new Set();
        for (let i = 0; i < filas; i++) {
            const f = fact[i];
            // Con varias medidas se suprime la fila/columna solo si TODAS
            // sus medidas son cero/vacío (antes solo miraba la última).
            let anyNonZero = false;
            for (let mIdx = 0; mIdx < measureCount; mIdx++) {
                const n = Number(f[measureFieldBase + mIdx]);
                if (!isNaN(n) && n !== 0) { anyNonZero = true; break; }
            }
            if (anyNonZero) {
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

    const rowsFilter = filterAndCompactDracoAxis(rowDict, totalDimFilas, dracoCollapseState.rows.collapsed, flagsFilas, { subtotalsOnTop: reportProps.subtotalsOnTop });
    const colsFilter = filterAndCompactDracoAxis(colDict, totalDimCols, dracoCollapseState.cols.collapsed, flagsColumnas, { subtotalsOnTop: reportProps.subtotalsOnTop });

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

    const resultSheetName = resultSheetNameFromGrid(editReportGrid, reportId);
    await ensureDracoResultSheetExists(context, resultSheetName);
    const sheet = context.workbook.worksheets.getItem(resultSheetName);

    // Rectángulo que ocupaba la ÚLTIMA pintada de ESTE informe (unión de
    // los rangos con nombre Draco_<id>_Rows/Cols/Values de la ejecución
    // anterior, si existen). Antes se limpiaba directamente TODO el
    // "usedRange" de la hoja (ver más abajo) porque cada informe tenía su
    // propia pestaña en exclusiva; ahora que varios informes pueden
    // compartir una misma pestaña (ver addReport()/reportStore.
    // createReport: ya no se crea una hoja nueva por informe, se usa la
    // hoja activa en el momento de crearlo), limpiar TODO el usedRange
    // borraría también lo pintado de OTROS informes de esa misma hoja. Se
    // calcula ANTES de limpiar los rangos con nombre porque clear() no
    // borra la definición del nombre ni su dirección, solo el contenido.
    const rangeNamesForClear = dracoRangeNames(reportId);
    const prevNamedItems = [rangeNamesForClear.rows, rangeNamesForClear.cols, rangeNamesForClear.values]
        .map(n => context.workbook.names.getItemOrNullObject(n));
    prevNamedItems.forEach(it => it.load("isNullObject"));
    await context.sync();

    const prevRanges = prevNamedItems.filter(it => !it.isNullObject).map(it => it.getRange());
    prevRanges.forEach(r => r.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]));
    if (prevRanges.length > 0) await context.sync();

    let prevBounds = null;
    for (const r of prevRanges) {
        const top = r.rowIndex, left = r.columnIndex;
        const bottom = r.rowIndex + r.rowCount, right = r.columnIndex + r.columnCount;
        if (!prevBounds) {
            prevBounds = { top, left, bottom, right };
        } else {
            prevBounds.top = Math.min(prevBounds.top, top);
            prevBounds.left = Math.min(prevBounds.left, left);
            prevBounds.bottom = Math.max(prevBounds.bottom, bottom);
            prevBounds.right = Math.max(prevBounds.right, right);
        }
    }

    // Antes de refrescar: borrar formato Y contenido de los rangos con
    // nombre Draco_<id>_Rows / Draco_<id>_Cols / Draco_<id>_Values DE ESTE
    // INFORME (y solo de este) de la ejecución anterior (si existen), para
    // no arrastrar colores/bordes de una tabla previa más grande o con
    // otra forma — y, sobre todo, para NO tocar los rangos de otros
    // informes (antes, con un nombre compartido "Draco_001_*", refrescar
    // un informe borraba lo pintado de otro).
    await clearDracoNamedRanges(context, reportId);

    // Limpiar cualquier resto de la ejecución anterior que hubiera quedado
    // FUERA de esos rangos con nombre (p.ej. fórmulas EPM_VALUE residuales
    // de una tabla previa más grande, y su formato: relleno, fuente,
    // bordes...), pero SOLO dentro del rectángulo que ocupaba la ejecución
    // anterior de ESTE informe (prevBounds, calculado arriba) — nunca toda
    // la hoja, que ahora puede tener pintados otros informes si comparten
    // pestaña. Los indicadores +/- van fusionados en las propias celdas de
    // Rows/Cols (no en una columna aparte), así que quedan cubiertos por
    // este mismo rectángulo. Se preserva siempre la fila 1 (A1=SQL,
    // B1=JSON), que se escribe aparte. Solo se reinicia el mapa de
    // indicadores DE ESTE informe (getDracoIndicatorMap ya lo crea vacío
    // si no existía); los de otros informes no se tocan.
    const dracoIndicatorMap = getDracoIndicatorMap(reportId);
    dracoIndicatorMap.clear();

    if (prevBounds) {
        const firstDataRow = Math.max(prevBounds.top, 1); // índice 0-based: fila 2 en adelante
        if (prevBounds.bottom > firstDataRow) {
            sheet.getRangeByIndexes(
                firstDataRow, prevBounds.left,
                prevBounds.bottom - firstDataRow, prevBounds.right - prevBounds.left
            ).clear(Excel.ClearApplyTo.all);
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

        // Una celda por cada medida (antes: solo la última, f[totalCampos-1]):
        // si el eje con medidas es el de Filas, se reparten en filas físicas
        // consecutivas (misma columna); si es el de Columnas, en columnas
        // físicas consecutivas (misma fila). Con una sola medida, measureCount
        // es 1 y esto se comporta exactamente igual que antes.
        for (let mIdx = 0; mIdx < measureCount; mIdx++) {
            const physRow = explodeForMeasures(newRowId, measuresOnRowsAxis, mIdx);
            const physCol = explodeForMeasures(newColId, measuresOnColsAxis, mIdx);
            const row = physRow + rowsOffRow;
            const col = physCol + colsOffCol;
            const value = coerceCellLiteral(String(f[measureFieldBase + mIdx]));
            factCells.set(row + "_" + col, { row, col, value });
        }
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

        // Si las medidas están en ESTE eje, cada dimensión real se repite en
        // las measureCount filas físicas de su grupo (misma columna, una
        // fila por medida); si no, se pinta una sola vez, como antes.
        const repeatCount = measuresOnRowsAxis ? measureCount : 1;

        for (let i = 1; i <= totalDimFilas; i++) {
            const flag = flagsFilas[i];
            const iAux = planFilas.iauxReal[i];

            if (String(V[i]).toLowerCase().indexOf("null") !== 0) {
                const indent = flag === 1 ? 0 : Math.max(0, flag - 1);
                const text = coerceCellLiteral(V[i]);
                // Eje Estático: se escribe como fórmula EPM_VALUE editable
                // (mismo formato que lee el flujo Fijo), no como texto plano.
                const cellVal = rowsStatic
                    ? buildEpmValueFormula(fieldsFilas[i].dim, fieldsFilas[i].attr, text)
                    : text;
                // Sobrescribe si ya había una entrada (mismo comportamiento que
                // el bucle secuencial original: el último nivel no-nulo gana).
                // OJO: coerceCellLiteral puede devolver un Number (p.ej. cuando
                // la dimensión es INTEGER, como YEAR), y Number no tiene
                // .trim() — sin el String(...) esto lanzaba una excepción a
                // media pintura (se pintaban los valores del FACT, pero se
                // interrumpía antes de pintar/formatear este eje).
                const isTotal = String(text).trim().toUpperCase() === "TOTAL";
                for (let mIdx = 0; mIdx < repeatCount; mIdx++) {
                    const row = explodeForMeasures(Number(V[0]), measuresOnRowsAxis, mIdx) + rowsOffRow;
                    const col = iAux + rowsOffCol;
                    filasCells.set(row + "_" + col, { row, col, value: cellVal, indent, field: iAux, isTotal });
                }
            }
        }

        // Filas MEASURE del eje (p.ej. "MEASURE" -> "IMPORTE"): no vienen del
        // dict (no son un valor variable por ROW_ID); se pintan como etiqueta
        // en su propia columna física, una por cada medida del grupo, en la
        // fila física que le corresponde a ESA medida (measuresOnRowsAxis).
        for (const m of planFilas.measureLevels) {
            const row = explodeForMeasures(Number(V[0]), measuresOnRowsAxis, m.ordinal) + rowsOffRow;
            const col = m.iAux + rowsOffCol;
            filasCells.set(row + "_" + col, { row, col, value: m.label, indent: 0, field: m.iAux, isTotal: false });
        }
    }

    // Fusionar el indicador +/- en la misma celda del nivel (nodos con
    // hijos), en vez de una columna aparte: "− España" / "+ España".
    // No aplica si el eje es Estático (esas celdas ya son fórmulas EPM_VALUE
    // editables a mano; anteponer un glifo de texto las rompería), ni
    // tampoco si "Sobrescribir formatos" está desactivado (D5): en ese
    // caso solo se escriben los valores, sin icono de jerarquía.
    if (!rowsStatic && reportProps.overwriteFormats) {
        // it.field (de filterAndCompactDracoAxis) es el ordinal entre las
        // dimensiones reales de NIVEL=1 (no cuenta las filas MEASURE);
        // se remapea a la columna física real (iAux) con flag1OrdinalToIaux,
        // que sí tiene en cuenta el hueco que dejan las filas MEASURE.
        const byLogicalKey = new Map(rowsFilter.indicators.map(it => [it.newId + "_" + planFilas.flag1OrdinalToIaux[it.field], it]));
        for (const [physKey, cell] of filasCells) {
            const physRowId = cell.row - rowsOffRow;
            // Si las medidas están en este eje, la fila física está
            // "explotada" (measureCount filas por cada ROW_ID lógico): hay
            // que deshacer eso para encontrar el nodo de jerarquía real.
            const newId = measuresOnRowsAxis ? Math.floor((physRowId - 1) / measureCount) + 1 : physRowId;
            const ind = byLogicalKey.get(newId + "_" + cell.field);
            if (ind) {
                const glyph = ind.collapsed ? "▸" : "▾";
                cell.value = glyph + " " + cell.value;
                dracoIndicatorMap.set(physKey, { axis: "rows", nodeKey: ind.nodeKey });
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

        // Análogo a "repeatCount" en FILAS, pero para columnas físicas.
        const repeatCount = measuresOnColsAxis ? measureCount : 1;

        for (let i = 1; i <= totalDimCols; i++) {
            const flag = flagsColumnas[i];
            const iAux = planColumnas.iauxReal[i];

            if (String(V[i]).toLowerCase().indexOf("null") !== 0) {
                const indent = flag === 1 ? 0 : Math.max(0, flag - 1);
                const text = coerceCellLiteral(V[i]);
                const cellVal = colsStatic
                    ? buildEpmValueFormula(fieldsColumnas[i].dim, fieldsColumnas[i].attr, text)
                    : text;
                // Ver comentario equivalente en el bloque de FILAS: text puede
                // ser un Number (dimensión INTEGER) y no tiene .trim().
                const isTotal = String(text).trim().toUpperCase() === "TOTAL";
                for (let mIdx = 0; mIdx < repeatCount; mIdx++) {
                    const row = iAux + colsOffRow;
                    const col = explodeForMeasures(Number(V[0]), measuresOnColsAxis, mIdx) + colsOffCol;
                    columnasCells.set(row + "_" + col, { row, col, value: cellVal, indent, field: iAux, isTotal });
                }
            }
        }

        // Filas MEASURE del eje columnas (p.ej. "MEASURE" -> "IMPORTE"/
        // "CANTIDAD"): una etiqueta por medida, en su propia columna física
        // dentro del grupo de columnas que le corresponde a ESE COLUMN_ID
        // (measuresOnColsAxis), en la fila que le corresponde a la medida
        // (respeta la posición configurada en EDIT_REPORT: por encima o por
        // debajo de ESCENARIO, según dónde esté la fila MEASURE en N/O/P).
        for (const m of planColumnas.measureLevels) {
            const row = m.iAux + colsOffRow;
            const col = explodeForMeasures(Number(V[0]), measuresOnColsAxis, m.ordinal) + colsOffCol;
            columnasCells.set(row + "_" + col, { row, col, value: m.label, indent: 0, field: m.iAux, isTotal: false });
        }
    }

    // Igual que en FILAS: el glifo +/- se fusiona en la propia celda
    // (salvo eje Estático, y salvo "Sobrescribir formatos" desactivado,
    // ver comentario arriba).
    if (!colsStatic && reportProps.overwriteFormats) {
        const byLogicalKey = new Map(colsFilter.indicators.map(it => [it.newId + "_" + planColumnas.flag1OrdinalToIaux[it.field], it]));
        for (const [physKey, cell] of columnasCells) {
            const physColId = cell.col - colsOffCol;
            const newId = measuresOnColsAxis ? Math.floor((physColId - 1) / measureCount) + 1 : physColId;
            const ind = byLogicalKey.get(newId + "_" + cell.field);
            if (ind) {
                const glyph = ind.collapsed ? "▸" : "▾";
                cell.value = glyph + " " + cell.value;
                dracoIndicatorMap.set(physKey, { axis: "cols", nodeKey: ind.nodeKey });
            }
        }
    }

    await writeCellBlock(context, sheet, columnasCells);
    if (reportProps.overwriteFormats) {
        await writeIndentAndColorRuns(context, sheet, columnasCells, "row", colsOffRow);
    }

    // ---- Punto 7: fondo RGB(255,255,204) en cabeceras "Total" y en los
    // valores de fila/columna de total (gateado por "Sobrescribir
    // formatos", igual que el resto de formato). ----
    if (reportProps.overwriteFormats) {
        await applyDracoTotalHighlight(context, sheet, filasCells, columnasCells, factCells);
    }

    /* -------------------------------------------------------------
     * 4) RANGOS CON NOMBRE Draco_001_Rows / Draco_001_Cols / Draco_001_Values
     *    + formato general (Segoe UI 9, número en Values, bordes finos),
     *    salvo que "Sobrescribir formatos" esté desactivado en las
     *    propiedades del informe (entonces solo se (re)definen los rangos).
     * ----------------------------------------------------------- */
    // Nº de COLUMNAS/FILAS físicamente pintadas: NO es totalDimFilas/totalDimCols
    // (eso cuenta cada NIVEL de una jerarquía como una entrada distinta, ya que
    // MODEL_HIER se expande a varias filas en EDIT_REPORT), sino el nº de
    // campos "de primer nivel" (flag/NIVEL === 1): cada campo -aunque sea una
    // jerarquía completa con varios niveles- se pinta en UNA sola columna
    // (los niveles siguientes solo añaden indentación dentro de esa misma
    // columna). Usar totalDimFilas/totalDimCols aquí generaba rangos con
    // nombre (Draco_001_Rows/Cols) más anchos/altos de lo realmente pintado.
    // Incluye tanto las dimensiones reales de NIVEL=1 como las filas MEASURE
    // (planFilas/planColumnas.totalIaux ya suma ambas), porque una fila
    // MEASURE también ocupa su propia columna/fila física pintada.
    const paintedFilasCols = planFilas.totalIaux;
    const paintedColsRows = planColumnas.totalIaux;

    // El ANCHO/ALTO de valores también crece ×measureCount en el eje que
    // lleva las medidas (measuresOnRowsAxis/measuresOnColsAxis): maxRowId/
    // maxColId son IDs lógicos (una combinación de dimensiones = 1 ID), pero
    // cada ID ocupa ahora measureCount filas/columnas físicas si ese es el
    // eje con medidas — los rangos con nombre deben reflejar ese ancho real
    // pintado, o quedarían más estrechos/bajos que la tabla real.
    const physicalMaxRowId = measuresOnRowsAxis ? maxRowId * measureCount : maxRowId;
    const physicalMaxColId = measuresOnColsAxis ? maxColId * measureCount : maxColId;

    await applyDracoNamedRanges(context, sheet, {
        RRows, RCols,
        totalDimFilas: paintedFilasCols,
        totalDimCols: paintedColsRows,
        maxRowId: physicalMaxRowId, maxColId: physicalMaxColId,
        applyVisualFormat: reportProps.overwriteFormats
    }, reportId);

    // 6) Registrar (una sola vez por hoja) los listeners de clic/edición:
    //    resuelven los indicadores +/- pintados arriba y el "Reconocimiento
    //    de miembros" sobre las celdas de Draco_<id>_Rows/Draco_<id>_Cols
    //    de ESTA hoja.
    await registerDracoSelectionHandler(context, sheet, resultSheetName);

    // 7) Autoajustar ancho de columnas (propiedades del informe: D6), solo
    // de los rangos con nombre Draco_<id>_Rows y Draco_<id>_Cols de ESTE
    // informe (donde se pintan sus filas/columnas), no de toda la hoja.
    if (reportProps.autoFitColumns) {
        try {
            const rn = dracoRangeNames(reportId);
            const rowsRangeName = context.workbook.names.getItemOrNullObject(rn.rows);
            const colsRangeName = context.workbook.names.getItemOrNullObject(rn.cols);
            rowsRangeName.load("isNullObject");
            colsRangeName.load("isNullObject");
            await context.sync();

            if (!rowsRangeName.isNullObject) rowsRangeName.getRange().format.autofitColumns();
            if (!colsRangeName.isNullObject) colsRangeName.getRange().format.autofitColumns();

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

/**
 * Punto 7: fondo RGB(255,255,204) para las celdas de cabecera (filas o
 * columnas) cuyo texto sea "Total" (subtotales/total general generados
 * por SQL con 'TOTAL'), y para las celdas de VALORES que caigan en una
 * fila o columna marcada como total.
 */
async function applyDracoTotalHighlight(context, sheet, filasCells, columnasCells, factCells) {
    const TOTAL_FILL = "#FFFFCC"; // RGB(255,255,204)

    const totalRows = new Set();
    const totalCols = new Set();
    const headerCells = [];

    for (const cell of filasCells.values()) {
        if (cell.isTotal) {
            headerCells.push(cell);
            totalRows.add(cell.row);
        }
    }
    for (const cell of columnasCells.values()) {
        if (cell.isTotal) {
            headerCells.push(cell);
            totalCols.add(cell.col);
        }
    }

    if (totalRows.size === 0 && totalCols.size === 0) return;

    for (const cell of headerCells) {
        sheet.getRangeByIndexes(cell.row - 1, cell.col - 1, 1, 1).format.fill.color = TOTAL_FILL;
    }

    for (const cell of factCells.values()) {
        if (totalRows.has(cell.row) || totalCols.has(cell.col)) {
            sheet.getRangeByIndexes(cell.row - 1, cell.col - 1, 1, 1).format.fill.color = TOTAL_FILL;
        }
    }

    await context.sync();
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
async function clearDracoNamedRanges(context, reportId) {
    const rn = dracoRangeNames(reportId);
    const names = [rn.rows, rn.cols, rn.values];
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
async function applyDracoNamedRanges(context, sheet, dims, reportId) {
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
    const rn = dracoRangeNames(reportId);
    const defs = [
        { name: rn.rows, range: rowsRange },
        { name: rn.cols, range: colsRange },
        { name: rn.values, range: valuesRange }
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
/**
 * Núcleo de Actualizar_informe(), sin manejo de `event`. reportId es
 * opcional (por defecto, el informe activo); se pasa explícito desde
 * "Refrescar todos" para poder recorrer varios informes sin depender de
 * cuál esté activo en el taskpane.
 */
async function actualizarInformeCore(reportIdOverride) {
    const reportId = reportIdOverride !== undefined ? reportIdOverride : activeReportIdOrNull();
    let sql;

    await Excel.run(async (context) => {
        const editReportGrid = await getEditReportGrid(context, reportId);
        const relGrid = await window.SemanticModelStore.getModelGrid("MODEL_RELATIONSHIP");
        const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");
        const atributesGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");

        loadReportDefinition(editReportGrid, reportId);

        // EDIT_REPORT!D4 = "Mostrar subtotales arriba" (Propiedades del
        // informe). Solo tiene efecto real cuando además hay algún campo
        // marcado con subtotal en L/R (ver hasAnySubtotalMarked/buildSQL).
        const subtotalsOnTop = String(cellValue(editReportGrid, 4, 4)).trim().toUpperCase() === "X";

        sql = buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop);

        console.log("BuildSQL ->", sql);

        // [Punto 8] SQL ya no se escribe en A1 de la hoja de resultados:
        // se escribe en EDIT_REPORT!X1.
        const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
        editReportSheet.getRange("X1").values = [[sql]];

        await context.sync();
    });

    const json = await executeSQL(sql);

    console.log("JSON de BigQuery ->", json);

    await Excel.run(async (context) => {
        const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
        const EXCEL_CELL_CHAR_LIMIT = 32000; // límite real de Excel: 32767
        const jsonForCell = json.length > EXCEL_CELL_CHAR_LIMIT
            ? json.substring(0, EXCEL_CELL_CHAR_LIMIT) + " ...(truncado, JSON completo en la consola F12)"
            : json;
        // [Punto 8] JSON ya no se escribe en B1 de la hoja de resultados:
        // se escribe en EDIT_REPORT!Y1.
        editReportSheet.getRange("Y1").values = [[jsonForCell]];
        await context.sync();
    });

    await Excel.run(async (context) => {
        await jsonTo3Matrices(context, json, reportId);
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
        await surfaceErrorToSheet(error);
    } finally {
        if (event) {
            event.completed();
        }
    }
}

/**
 * Núcleo de Actualizar(): dispatcher según EDIT_REPORT!H12/N12 (mismo
 * criterio "Fijo" vs "Dinámico" de siempre), pero reutilizable para UN
 * informe concreto (reportId opcional, por defecto el activo). La usan
 * tanto actualizar() (botón "Actualizar", un único informe: el activo)
 * como actualizarTodosCore() (recorre todos los informes).
 */
async function actualizarUnInforme(reportIdOverride) {
    const reportId = reportIdOverride !== undefined ? reportIdOverride : activeReportIdOrNull();
    const isFixed = await Excel.run(async (context) => {
        const grid = await getEditReportGrid(context, reportId);
        const h12 = String(cellValue(grid, 12, 8)).trim().toUpperCase();
        const n12 = String(cellValue(grid, 12, 14)).trim().toUpperCase();
        return h12 === "X" && n12 === "X";
    });

    if (isFixed) {
        await actualizarInformeFixedCore(reportId);
    } else {
        await actualizarInformeCore(reportId);
    }
}

/**
 * Traducción de Actualizar(): dispatcher según EDIT_REPORT!H12/N12.
 * @param {Office.AddinCommands.Event} event
 */
async function actualizar(event) {
    try {
        await actualizarUnInforme();
    } catch (error) {
        console.error("Error al actualizar el informe:", error);
        await surfaceErrorToSheet(error);
    } finally {
        if (event) {
            event.completed();
        }
    }
}

/**
 * Ejecuta N tareas asíncronas con un límite de concurrencia (para no
 * lanzar, p.ej., 20 consultas a BigQuery/Snowflake a la vez sin control).
 */
async function runWithConcurrency(items, limit, worker) {
    let next = 0;
    async function runNext() {
        while (next < items.length) {
            const idx = next++;
            await worker(items[idx], idx);
        }
    }
    const workers = [];
    for (let i = 0; i < Math.min(limit, items.length); i++) workers.push(runNext());
    await Promise.all(workers);
}

/**
 * "Refrescar todos los informes": recorre TODOS los informes guardados en
 * ReportStore (no solo el activo) y los actualiza:
 *   1) Construye el SQL de cada uno (síncrono/local, en serie: no hace
 *      falta red y así no se pisan los unos a los otros al leer su grid).
 *   2) Lanza las consultas a BigQuery/Snowflake EN PARALELO (con un tope
 *      de concurrencia): es la parte lenta (red) y cada informe es
 *      independiente, así que no hay motivo para esperarlas una a una.
 *   3) Pinta cada resultado en Excel EN SERIE: la API de Excel no admite
 *      escrituras concurrentes seguras sobre el mismo libro (Excel.run/
 *      context.sync no están pensados para solaparse así).
 * Un fallo en un informe concreto (SQL, consulta o pintado) no interrumpe
 * a los demás: se registra en consola y se sigue con el resto.
 */
async function actualizarTodosCore(concurrency) {
    if (!window.ReportStore) return;
    const reports = window.ReportStore.listReports();
    if (!reports || reports.length === 0) return;

    const dynamicJobs = [];
    const fixedReports = [];

    // 1) Clasificar cada informe (Fijo/Dinámico, según H12/N12 de su
    //    propio grid) y, para los Dinámicos, construir ya su SQL: es
    //    síncrono/local (no toca red), así que se hace en serie sin
    //    problema — no hay motivo para paralelizar esta parte.
    for (const r of reports) {
        try {
            const reportId = r.id;
            const isFixed = await Excel.run(async (context) => {
                const grid = await getEditReportGrid(context, reportId);
                const h12 = String(cellValue(grid, 12, 8)).trim().toUpperCase();
                const n12 = String(cellValue(grid, 12, 14)).trim().toUpperCase();
                return h12 === "X" && n12 === "X";
            });

            if (isFixed) {
                fixedReports.push({ reportId, reportName: r.name });
            } else {
                let sql;
                await Excel.run(async (context) => {
                    const editReportGrid = await getEditReportGrid(context, reportId);
                    const relGrid = await window.SemanticModelStore.getModelGrid("MODEL_RELATIONSHIP");
                    const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");
                    const atributesGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
                    loadReportDefinition(editReportGrid, reportId);
                    const subtotalsOnTop = String(cellValue(editReportGrid, 4, 4)).trim().toUpperCase() === "X";
                    sql = buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop);
                });
                dynamicJobs.push({ reportId, reportName: r.name, sql });
            }
        } catch (err) {
            console.error(`[Draco] Error preparando el informe "${r.name}" (id ${r.id}):`, err);
        }
    }

    // 2) Consultas de los informes "Dinámicos" a BigQuery/Snowflake EN
    //    PARALELO (tope de concurrencia por defecto 4): es la parte lenta
    //    (red, sin tocar Excel) y cada informe es independiente.
    await runWithConcurrency(dynamicJobs, concurrency || 4, async (job) => {
        try {
            job.json = await executeSQL(job.sql);
        } catch (err) {
            job.error = err;
            console.error(`[Draco] Error consultando datos del informe "${job.reportName}":`, err);
        }
    });

    // 3) Pintar los "Dinámicos" en Excel EN SERIE: la API de Excel no
    //    admite escrituras concurrentes seguras sobre el mismo libro.
    for (const job of dynamicJobs) {
        if (job.error || !job.json) continue;
        try {
            await Excel.run(async (context) => {
                const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
                const EXCEL_CELL_CHAR_LIMIT = 32000;
                const jsonForCell = job.json.length > EXCEL_CELL_CHAR_LIMIT
                    ? job.json.substring(0, EXCEL_CELL_CHAR_LIMIT) + " ...(truncado, JSON completo en la consola F12)"
                    : job.json;
                editReportSheet.getRange("Y1").values = [[jsonForCell]];
                await jsonTo3Matrices(context, job.json, job.reportId);
            });
        } catch (err) {
            console.error(`[Draco] Error pintando el informe "${job.reportName}":`, err);
        }
    }

    // 4) Informes en modo "Fijo" (ambos ejes Estático): flujo aparte,
    //    todo-en-uno (aquí incluso construir el SQL necesita leer la hoja
    //    de resultados en Excel, no solo el pintado), así que van en SERIE
    //    entre sí — no solo el pintado, sino el flujo completo.
    for (const f of fixedReports) {
        try {
            await actualizarInformeFixedCore(f.reportId);
        } catch (err) {
            console.error(`[Draco] Error actualizando el informe "${f.reportName}" (modo fijo):`, err);
        }
    }
}

async function actualizarTodos(event) {
    try {
        await actualizarTodosCore();
    } catch (error) {
        console.error("Error al actualizar todos los informes:", error);
        await surfaceErrorToSheet(error);
    } finally {
        if (event) {
            event.completed();
        }
    }
}

// Exponer las funciones de actualización para poder llamarlas también desde
// el taskpane (p.ej. tras guardar el diseño), no solo desde el ribbon.
window.ReportActions = {
    actualizar, actualizarInforme, actualizarInformeFixed, actualizarTodos,
    toggleRefreshPaused, toggleMemberRecognition, openReportProperties, openFieldOptions,
    convertAxisStaticFormulas, ensureDracoHandlersRegistered,
    // "Añadir filtro" (rango con nombre sobre una celda cualquiera del libro)
    dracoFilterRangeName, ensureDracoFilterRangeHandlersRegistered,
    resolveDracoFilterRangeAddress, openDracoFilterRangePicker,
    // Registro del clic en Draco_XXX_Rows -> escribe zona + Jerarquia SI/NO en A51
    ensureDracoRowsClickLoggerRegistered
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

async function requestRibbonLabelUpdate(controlId, label, groupId) {
    try {
        if (Office.ribbon && Office.ribbon.requestUpdate) {
            await Office.ribbon.requestUpdate({
                tabs: [{
                    id: "DracoTab",
                    groups: [{ id: groupId || "EdicionGroup", controls: [{ id: controlId, label }] }]
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

/**
 * Escribe "X" (ON) o "" (OFF) en EDIT_REPORT!B1, para que el resto del
 * flujo (SQL/reconocimiento) pueda leer el estado directamente de la
 * hoja si lo necesita, además del guardado en Office roaming settings.
 * Si EDIT_REPORT no existe todavía no falla: simplemente no escribe nada.
 */
async function writeMemberRecognitionFlagToSheet(nowOn) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
            sheet.load("isNullObject");
            await context.sync();

            if (sheet.isNullObject) {
                console.log("[Draco] writeMemberRecognitionFlagToSheet: EDIT_REPORT no existe todavía, no se escribe B1.");
                return;
            }

            sheet.getRange("B1").values = [[nowOn ? "X" : ""]];
            await context.sync();
        });
    } catch (e) {
        console.warn("[Draco] No se pudo escribir el estado de Reconocimiento de miembros en EDIT_REPORT!B1:", e);
    }
}

async function toggleMemberRecognition(event) {
    try {
        const nowOn = await toggleDracoSetting("draco_memberRecognition");
        console.log("Reconocimiento de miembros:", nowOn ? "activado" : "desactivado");

        await writeMemberRecognitionFlagToSheet(nowOn);

        await requestRibbonLabelUpdate(
            "ReconocimientoMiembrosButton",
            nowOn ? "Reconocimiento de miembros (ON)" : "Reconocimiento de miembros (OFF)",
            "EdicionGroup"
        );
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

async function abrirDistribuirValores(event) {
    try {
        console.log("[Draco] abrirDistribuirValores: shared runtime ¿activo? ->", typeof TaskPaneApp !== "undefined");
        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.openDistributeValuesModal) {
            TaskPaneApp.openDistributeValuesModal();
            console.log("[Draco] openDistributeValuesModal() ejecutado directamente (shared runtime OK).");
        } else {
            console.warn("[Draco] TaskPaneApp NO existe en este contexto: usando fallback de settings (¿manifest sin Runtimes recargado?).");
            const settings = Office.context.document.settings;
            settings.set("draco_pendingAction", "distributeValues");
            await new Promise((resolve) => settings.saveAsync(resolve));
        }
        if (Office.addin && Office.addin.showAsTaskpane) {
            await Office.addin.showAsTaskpane();
            console.log("[Draco] showAsTaskpane() resuelto sin error.");
        } else {
            console.warn("[Draco] Office.addin.showAsTaskpane no está disponible en este runtime.");
        }
    } catch (error) {
        console.error("Error al abrir Distribuir valores:", error);
    } finally {
        if (event) event.completed();
    }
}

/**
 * Botón de ribbon "Añadir filtro" (ver AnadirFiltroButton en manifest.xml):
 * abre addFilterRange.html como diálogo INDEPENDIENTE de Office
 * (Office.context.ui.displayDialogAsync), sin depender del taskpane ni de
 * shared runtime — exactamente el mismo patrón que openDracoFilterRangePicker
 * usa para el selector de valores. Antes intentaba
 * Office.addin.showAsTaskpane(), pero esa API SOLO existe cuando el
 * manifest declara <Runtimes> (shared runtime), lo cual no es el caso aquí;
 * por eso el botón no hacía nada visible.
 *
 * Flujo:
 *  1) Captura la celda seleccionada AHORA MISMO en Excel (antes de abrir el
 *     diálogo: una ventana de diálogo no tiene acceso a Excel.run).
 *  2) Carga dimensiones/jerarquías (ExcelService.readDim2Data) e informes
 *     (ReportStore.listReports) — ninguno de los dos necesita Excel.run,
 *     leen de Office roaming settings, así que es seguro hacerlo aquí.
 *  3) Abre el diálogo y le envía esos datos en cuanto avisa "ready".
 *  4) Al recibir "apply", crea/reemplaza el rango con nombre
 *     Draco_Filter_<DIM>_<CAMPO>_<sufijo> sobre la celda capturada en el
 *     paso 1 (misma lógica que tenía TaskPaneApp.createFilterRangeFromModal)
 *     y lo guarda en FilterRangeStore.
 */
async function abrirAnadirFiltro(event) {
    try {
        let addr = "";
        let sheetName = "";
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load("address");
            await context.sync();

            sheetName = sheet.name;
            // Esquina superior izquierda si hay varias celdas seleccionadas.
            const rawAddress = String(selectedRange.address).split("!").pop();
            addr = rawAddress.split(":")[0];
        });

        let dimensions = [];
        let dimensionsError = "";
        try {
            const result = await window.ExcelService.readDim2Data();
            if (result && result.error) {
                dimensionsError = result.error;
            } else {
                dimensions = (result && result.data) || [];
            }
        } catch (err) {
            console.error("[Draco] Añadir filtro: error cargando dimensiones:", err);
            dimensionsError = "Error al cargar los campos del modelo semántico.";
        }

        const reports = (window.ReportStore ? window.ReportStore.listReports() : []) || [];

        const dialogUrl = new URL("addFilterRange.html", window.location.href).href;

        await new Promise((resolve) => {
            Office.context.ui.displayDialogAsync(
                dialogUrl,
                { height: 45, width: 32, displayInIframe: false },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error(
                            "[Draco] Añadir filtro: displayDialogAsync ha fallado:",
                            asyncResult.error && asyncResult.error.code,
                            asyncResult.error && asyncResult.error.message
                        );
                        resolve();
                        return;
                    }

                    const dialog = asyncResult.value;
                    let settled = false;
                    const closeDialog = () => { try { dialog.close(); } catch (e) { /* ya cerrado */ } };

                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                        let payload;
                        try {
                            payload = JSON.parse(arg.message);
                        } catch (err) {
                            console.error("[Draco] Añadir filtro: mensaje del diálogo no es JSON válido:", err);
                            return;
                        }

                        if (payload.type === "ready") {
                            if (dimensionsError) {
                                dialog.messageChild(JSON.stringify({ type: "error", message: dimensionsError }));
                            } else {
                                dialog.messageChild(JSON.stringify({ type: "data", dimensions, reports }));
                            }
                            return;
                        }

                        if (payload.type === "apply") {
                            settled = true;
                            closeDialog();
                            try {
                                const fieldData = payload.fieldData;
                                const reportValue = payload.reportValue;
                                const isAll = !reportValue || reportValue === "all";
                                const reportId = isAll ? null : Number(reportValue);
                                const suffix = isAll ? "all" : pad3(reportId);
                                const rangeName = dracoFilterRangeName(fieldData.dim, fieldData.name, suffix);

                                await Excel.run(async (context) => {
                                    const sheet = context.workbook.worksheets.getItem(sheetName);
                                    const cell = sheet.getRange(addr);

                                    // Si ya existía un rango con este mismo nombre (misma
                                    // dimensión/campo + mismo informe), se borra y se
                                    // vuelve a crear sobre la NUEVA celda seleccionada
                                    // (así "Añadir filtro" también sirve para mover un
                                    // filtro existente).
                                    const existing = context.workbook.names.getItemOrNullObject(rangeName);
                                    existing.load("isNullObject");
                                    await context.sync();
                                    if (!existing.isNullObject) {
                                        existing.delete();
                                        await context.sync();
                                    }

                                    context.workbook.names.add(rangeName, cell);
                                    await context.sync();
                                });

                                await window.FilterRangeStore.set(rangeName, {
                                    dim: fieldData.dim,
                                    name: fieldData.name,
                                    isHierarchy: !!fieldData.isHierarchy,
                                    reportId: reportId,
                                    filter: null
                                });

                                // Por si el usuario ha creado el filtro en una hoja nueva
                                // que todavía no tuviera enganchado el listener de clic,
                                // y para refrescar la zona "Filtros" si el taskpane ya
                                // estaba abierto (mismo contexto que commands.js).
                                if (window.ReportActions && window.ReportActions.ensureDracoFilterRangeHandlersRegistered) {
                                    await window.ReportActions.ensureDracoFilterRangeHandlersRegistered();
                                }
                                if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.loadDesignFromSheet) {
                                    try { await TaskPaneApp.loadDesignFromSheet(); } catch (e) { /* taskpane sin informe abierto: se ignora */ }
                                }
                            } catch (err) {
                                console.error("[Draco] Añadir filtro: error al crear el rango de filtro:", err);
                            }
                            resolve();
                            return;
                        }

                        if (payload.type === "cancel") {
                            settled = true;
                            closeDialog();
                            resolve();
                        }
                    });

                    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                        if (!settled) resolve();
                    });
                }
            );
        });
    } catch (error) {
        console.error("Error al abrir 'Añadir filtro':", error);
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
    Office.actions.associate("abrirModeloSemantico", abrirModeloSemantico);
    Office.actions.associate("guardarModeloSemantico", guardarModeloSemantico);
    Office.actions.associate("writeHolaInA1", writeHolaInA1);
    Office.actions.associate("actualizarInformeFixed", actualizarInformeFixed);
    Office.actions.associate("actualizarInforme", actualizarInforme);
    Office.actions.associate("actualizar", actualizar);
    Office.actions.associate("comingSoon", comingSoon);
    Office.actions.associate("toggleRefreshPaused", toggleRefreshPaused);
    Office.actions.associate("toggleMemberRecognition", toggleMemberRecognition);
    Office.actions.associate("openReportProperties", openReportProperties);
    Office.actions.associate("openFieldOptions", openFieldOptions);
    Office.actions.associate("abrirDistribuirValores", abrirDistribuirValores);
    Office.actions.associate("abrirAnadirFiltro", abrirAnadirFiltro);
    Office.actions.associate("guardarExcelEnBucket", guardarExcelEnBucket);
    Office.actions.associate("abrirDesdeBucket", abrirDesdeBucket);
} catch (e) {
    console.warn("Office.actions.associate no disponible en este contexto:", e);
}
