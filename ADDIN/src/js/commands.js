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

    // NOTA: existía aquí un registro automático y global (todas las hojas)
    // de un listener de onSelectionChanged para "Añadir filtro"
    // (ensureDracoFilterRangeHandlersRegistered / handleDracoFilterRangeSelection).
    // Se ha eliminado por completo del fichero.
});

/* ==========================================================================
 * LOG DE DIAGNÓSTICO EN EDIT_REPORT!AA1:AA151
 * --------------------------------------------------------------------------
 * dracoLog(...) escribe en consola y en un buffer en memoria que se vuelca
 * (con debounce, en UN solo Excel.run) a EDIT_REPORT!AA1:AA151, la línea
 * más reciente arriba. AA1 lleva el BUILD cargado: si después de publicar
 * en GitHub Pages no ves el build nuevo en AA1, Excel sigue usando la
 * versión en caché (y cualquier "arreglo" no se está probando de verdad).
 * El picker (memberPicker.html) también manda sus logs aquí por mensaje.
 * Para desactivarlo: DRACO_LOG_ENABLED = false.
 * ========================================================================== */
const DRACO_BUILD = "2026-10-28-merge";
const DRACO_LOG_ENABLED = true;
const DRACO_LOG_SHEET = "EDIT_REPORT";
const DRACO_LOG_COL = "AA";
const DRACO_LOG_MAX = 150;
const DracoLogLines = [];
let DracoLogFlushTimer = null;
let DracoLogFlushing = false;
let DracoLogDirty = false;

function dracoErrMsg(e) {
    if (!e) return "(error vacío)";
    if (typeof e === "string") return e;
    return (e.code ? e.code + ": " : "") + (e.message || String(e));
}

function dracoLog(...args) {
    const d = new Date();
    const ts = d.toTimeString().slice(0, 8) + "." + String(d.getMilliseconds()).padStart(3, "0");
    const msg = args.map((a) => {
        if (typeof a === "string") return a;
        try { return JSON.stringify(a); } catch (e) { return String(a); }
    }).join(" ");
    console.log("[DracoLog]", ts, msg);
    if (!DRACO_LOG_ENABLED) return;
    DracoLogLines.unshift(ts + " | " + msg.slice(0, 600));
    if (DracoLogLines.length > DRACO_LOG_MAX) DracoLogLines.length = DRACO_LOG_MAX;
    DracoLogDirty = true;
    scheduleDracoLogFlush();
}
window.dracoLog = dracoLog;

function scheduleDracoLogFlush() {
    if (DracoLogFlushTimer) return;
    DracoLogFlushTimer = setTimeout(flushDracoLog, 700);
}

async function flushDracoLog() {
    DracoLogFlushTimer = null;
    if (DracoLogFlushing) { scheduleDracoLogFlush(); return; }
    if (!DracoLogDirty) return;
    DracoLogFlushing = true;
    DracoLogDirty = false;
    const snapshot = DracoLogLines.slice();
    try {
        await Excel.run(async (context) => {
            const sh = context.workbook.worksheets.getItemOrNullObject(DRACO_LOG_SHEET);
            sh.load("isNullObject");
            await context.sync();
            if (sh.isNullObject) return;
            const values = [["LOG DRACO | build " + DRACO_BUILD + " | más reciente arriba"]];
            for (let i = 0; i < DRACO_LOG_MAX; i++) values.push([snapshot[i] || ""]);
            const rng = sh.getRange(DRACO_LOG_COL + "1:" + DRACO_LOG_COL + (DRACO_LOG_MAX + 1));
            rng.numberFormat = values.map(() => ["@"]); // texto: que nada se interprete como fórmula/fecha
            rng.values = values;
            await context.sync();
        });
    } catch (e) {
        console.warn("[DracoLog] No se pudo volcar el log a EDIT_REPORT:", e);
    } finally {
        DracoLogFlushing = false;
        if (DracoLogDirty) scheduleDracoLogFlush();
    }
}

Office.onReady(() => {
    dracoLog("Add-in cargado. BUILD=" + DRACO_BUILD + " | url=" + window.location.pathname);
});

/**
 * Asegura que exista la hoja técnica EDIT_REPORT (estado del diseño del
 * informe: filtros/filas/columnas), igual que ensureCoreModelSheets() en
 * semantic_model.js — pero ese fichero solo se carga en el taskpane
 * (semantic_model.html), no en el runtime de comandos del ribbon
 * (commands.html), así que aquí se repite la misma lógica para que
 * abrirModeloSemantico() (botón del ribbon) también la cree si hace
 * falta. Se crea vacía si no existe, salvo D5 y D6 que llevan "X", y se
 * deja oculta.
 */
async function ensureEditReportSheetFromRibbon() {
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;

            let editSheet = sheets.getItemOrNullObject("EDIT_REPORT");
            await context.sync();

            if (editSheet.isNullObject) {
                editSheet = sheets.add("EDIT_REPORT");
                editSheet.getRange("D5").values = [["X"]];
                editSheet.getRange("D6").values = [["X"]];
            }

            editSheet.visibility = Excel.SheetVisibility.hidden;
            await context.sync();
        });
    } catch (e) {
        console.error("[Draco] No se pudo asegurar la hoja EDIT_REPORT al abrir el modelo semántico:", e);
    }
}

/**
 * Botón de ribbon "Abrir modelo semántico" (ModeloAbrirButton).
 * Abre directamente el diálogo independiente de importación LookML
 * (Office.context.ui.displayDialogAsync), sin depender de que el taskpane
 * del modelo semántico esté abierto ni de ningún popup dentro de él.
 * Antes de abrir el diálogo se asegura la hoja EDIT_REPORT (ver
 * ensureEditReportSheetFromRibbon), igual que hace el taskpane del
 * modelo semántico en cuanto se abre.
 * @param {Office.AddinCommands.Event} event
 */
function abrirModeloSemantico(event) {
    try {
        ensureEditReportSheetFromRibbon().finally(() => {
            // LkmlOpenBridge (js/lkmlOpenBridge.js) abre openSemanticModel.html
            // y, cuando el usuario elige un fichero y confirma, guarda el
            // modelo importado en SemanticModelStore desde este mismo runtime
            // de comandos (que sí tiene acceso a Office.context.document.settings).
            window.LkmlOpenBridge.openOpenLkmlDialog();
        });
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
 * Botones de ribbon "Editar informes" (CrearInformeButton), "Conexión"
 * (ConexionButton) y "Editar modelos semántico" (ModeloCrearButton).
 *
 * Antes eran ShowTaskpane (3 paneles acoplados distintos: Taskpane/
 * ConexionPane/ModeloPane). Con Shared Runtime solo puede existir UN
 * task pane en todo el manifest — y NINGÚN <Action> del Host con
 * <Runtimes> debe llevar <TaskpaneId> (documentación oficial: "In the
 * ancestor <Host> element that has a descendant <Runtime> element that
 * is set to a long lifetime, none of the descendant <Action> elements
 * should have a <TaskpaneID> child element"). Así que los 3 son ahora
 * ExecuteFunction: usan Office.addin.showAsTaskpane() (API que requiere
 * precisamente Shared Runtime) para asegurar que el panel está visible,
 * y luego mostrarVistaEpm() (definida en taskpane.html, al final del
 * body) para enseñar solo la sección correspondiente — login.html,
 * semantic_model.html y taskpane.html YA NO SON PÁGINAS APARTE: su
 * contenido vive fusionado dentro de taskpane.html como 3 <div>
 * ocultables (#view-informes / #view-conexion / #view-modelo).
 *
 * IMPORTANTE (probado y descartado): se intentó primero abrir
 * login.html/semantic_model.html como DIÁLOGOS independientes
 * (Office.context.ui.displayDialogAsync) en vez de fusionarlos — no
 * funciona, porque Office.js no permite abrir un diálogo DESDE DENTRO
 * de otro diálogo, y tanto login.js (el flujo OAuth) como
 * semantic_model.js (abrir/guardar LookML) abren sus propios diálogos
 * hijos. Por eso la fusión en un único taskpane es la única vía viable
 * con Shared Runtime.
 */
function mostrarVistaInformes(event) {
    try {
        Office.addin.showAsTaskpane()
            .then(() => {
                if (window.mostrarVistaEpm) window.mostrarVistaEpm("view-informes");
            })
            .catch((e) => console.error("[Draco] Error mostrando la vista Editar informes:", e));
    } catch (error) {
        console.error("Error al mostrar la vista de Editar informes:", error);
    } finally {
        if (event) event.completed();
    }
}

function mostrarVistaConexion(event) {
    try {
        Office.addin.showAsTaskpane()
            .then(() => {
                if (window.mostrarVistaEpm) window.mostrarVistaEpm("view-conexion");
            })
            .catch((e) => console.error("[Draco] Error mostrando la vista Conexión:", e));
    } catch (error) {
        console.error("Error al mostrar la vista de Conexión:", error);
    } finally {
        if (event) event.completed();
    }
}

function mostrarVistaEditarModelo(event) {
    try {
        // Misma comprobación/creación de EDIT_REPORT que ya hacía
        // abrirModeloSemantico (botón "Abrir modelo semántico"): antes la
        // garantizaba semantic_model.js al cargar el panel; se asegura
        // aquí desde el ribbon igual que en ese otro botón.
        ensureEditReportSheetFromRibbon().finally(() => {
            Office.addin.showAsTaskpane()
                .then(() => {
                    if (window.mostrarVistaEpm) window.mostrarVistaEpm("view-modelo");
                })
                .catch((e) => console.error("[Draco] Error mostrando la vista Editar modelo semántico:", e));
        });
    } catch (error) {
        console.error("Error al mostrar la vista de Editar modelo semántico:", error);
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
        // SEGURIDAD: ya no se pasa el token por la URL (ver
        // BQ.handleDialogMessage/requestSessionFromOpener en bigquery.js).
        Office.context.ui.displayDialogAsync(url.href, { height: 55, width: 40, displayInIframe: false }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("No se pudo abrir el diálogo del bucket:", asyncResult.error.message);
                return;
            }
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                if (window.BQ && BQ.handleDialogMessage(dialog, arg.message)) return;
                dialog.close();
            });
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
 *   - ProjectId: el de facturación de la conexión o, si no hay, el de la
 *     tabla consultada (BQ.jobProjectFor).
 *   - Token: BQ.getToken() (sessionStorage; la hoja "CONFIG" no existe).
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
 * "Mostrar ID y texto" en una JERARQUÍA (ver ReportStore._expandAxis): la
 * jerarquía de textos se añade tras la de IDs con este sufijo en el NOMBRE de
 * jerarquía (columna K/Q de EDIT_REPORT). Sus niveles comparten grupo de
 * GROUPING SETS con los niveles de ID (no son niveles nuevos) y no cuentan
 * como profundidad al contraer/expandir.
 */
const DRACO_HIER_TEXT_SUFFIX = "~TXT";
// Columna de orden de EDIT_REPORT (M/S) con la marca "HIDE": "Solo texto".
function dracoIsHiddenMarker(v) {
    return String(v === null || v === undefined ? "" : v).toUpperCase().indexOf("HIDE") !== -1;
}
function dracoIsHierTextName(name) {
    const n = String(name || "");
    return n.length > DRACO_HIER_TEXT_SUFFIX.length && n.endsWith(DRACO_HIER_TEXT_SUFFIX);
}

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
    Measures: [],
    // Medidas arrastradas a la zona "Filtros" (ver loadFilters /
    // appendLockedFilterRangesToState / applyFilterMeasuresToDynamicReport):
    // no van en Filters (no afectan al WHERE) ni se cuentan en Rows/Columns.
    FilterMeasureNames: []
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

// Etiqueta "00X - <fase>" para BusyIndicator.update() durante un
// refresco (ver actualizarInformeCore/actualizarInformeFixedCore/
// jsonTo3MatricesCore/actualizarTodosCore): identifica de qué informe es
// cada fase cuando puede haber más de uno de por medio (p.ej. "Actualizar
// todos"). reportId puede venir vacío/null (informe activo sin resolver
// todavía); en ese caso se omite el prefijo en vez de mostrar "000 - ...".
function busyStepLabel(reportId, phase) {
    return reportId ? (pad3(reportId) + " - " + phase) : phase;
}

// Acumula las líneas de draco_perfMark de UN refresco (o de "Actualizar
// todos" entero) para volcarlas en EDIT_REPORT!Z1 al terminar — así no
// hace falta abrir la consola del navegador (F12): un add-in de Excel
// (JavaScript/Office.js) no tiene ninguna "ventana Inmediato" como el
// editor de VBA, así que la celda es el sustituto más parecido. Depth
// funciona igual que el refCount de BusyIndicator: "Actualizar todos"
// llama a beginDracoPerfLog() una vez para TODO el lote, y cada informe
// individual (actualizarInformeCore/actualizarInformeFixedCore) puede
// seguir llamando a begin/end sin que el registro se reinicie o se
// escriba a medias entre uno y otro — solo se vacía en el primer begin()
// (depth 0->1) y solo se escribe en el último end() (depth 1->0).
let DracoPerfLog = [];
let DracoPerfLogDepth = 0;

function beginDracoPerfLog() {
    DracoPerfLogDepth++;
    if (DracoPerfLogDepth === 1) DracoPerfLog = [];
}

async function endDracoPerfLog() {
    DracoPerfLogDepth = Math.max(0, DracoPerfLogDepth - 1);
    if (DracoPerfLogDepth > 0) return; // todavía dentro de otra llamada (p.ej. "Actualizar todos")
    if (DracoPerfLog.length === 0) return;

    const text = DracoPerfLog.join("\n");
    try {
        // Mismo guardado que writeDracoPlanningDebug (W1): se protege con
        // DracoSuppressChangeEvents para que esta propia escritura no
        // dispare handleDracoPlanningValueChanged sobre sí misma.
        DracoSuppressChangeEvents = true;
        await Excel.run(async (context) => {
            const editReportSheet = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
            editReportSheet.load("isNullObject");
            await context.sync();
            if (editReportSheet.isNullObject) return;
            editReportSheet.getRange("Z1").values = [[draco_debugCell(text)]];
            await context.sync();
        });
    } catch (e) {
        console.error("[Draco] No se pudo escribir el registro de tiempos en EDIT_REPORT!Z1:", e);
    } finally {
        DracoSuppressChangeEvents = false;
    }
}

/**
 * Instrumentación de tiempos (consola F12 Y EDIT_REPORT!Z1, ver
 * begin/endDracoPerfLog): marca cuánto ha tardado la fase que acaba de
 * terminar, desde el checkpoint anterior (performance.now() en tPrev), y
 * devuelve el checkpoint ACTUAL para encadenar la siguiente llamada.
 * Pensado para responder "¿qué es lo que de verdad tarda -consulta,
 * parseo del JSON, o cada parte del pintado-?" con datos reales en vez
 * de suposiciones. No afecta al indicador visual (BusyIndicator.update,
 * que es un texto para el usuario) ni al comportamiento del refresco.
 */
function draco_perfMark(reportId, phase, tPrev) {
    const now = performance.now();
    const line = `${busyStepLabel(reportId, phase)}: ${(now - tPrev).toFixed(0)} ms`;
    console.log(`[Draco][perf] ${line}`);
    DracoPerfLog.push(line);
    return now;
}

// Mismo patrón de profundidad que DracoPerfLogDepth/BusyIndicator.refCount:
// "Actualizar todos" pone el cálculo en manual UNA sola vez para TODO el
// lote (llama a draco_beginManualCalculation antes del bucle), y cada
// informe individual (actualizarInformeCore/actualizarInformeFixedCore)
// puede seguir llamando a begin/end sin que se restaure a mitad de camino
// -solo se restaura de verdad en el ÚLTIMO end() (depth 1->0)-.
let DracoCalcModeDepth = 0;
let DracoCalcModePrevMode = null;

/**
 * Pone el cálculo de Excel en Manual mientras dura el refresco: con
 * cálculo Automático (lo normal), CADA sync() que cambia una celda puede
 * disparar un recálculo de TODO el libro (fórmulas, EPM_VALUE, lo que
 * dependa de lo que sea) — esto es clásico de cualquier macro VBA seria
 * (Application.Calculation = xlCalculationManual) y explica la varianza
 * tan grande vista en operaciones minúsculas (2 celdas a veces tardando
 * más que 432: el coste no es el tamaño de lo que escribimos, es el
 * recálculo que dispara). Se guarda el modo ANTERIOR para restaurarlo
 * exacto al acabar (ver draco_endManualCalculation) — nunca se fuerza a
 * "Automático" a pelo, por si el usuario ya lo tenía en Manual él mismo.
 */
async function draco_beginManualCalculation() {
    DracoCalcModeDepth++;
    if (DracoCalcModeDepth > 1) return; // ya en manual por una llamada exterior (p.ej. "Actualizar todos")
    try {
        await Excel.run(async (context) => {
            const app = context.workbook.application;
            app.load("calculationMode");
            await context.sync();
            DracoCalcModePrevMode = app.calculationMode;
            if (DracoCalcModePrevMode !== Excel.CalculationMode.manual) {
                app.calculationMode = Excel.CalculationMode.manual;
                await context.sync();
            }
        });
    } catch (e) {
        console.warn("[Draco] No se pudo poner el cálculo en manual:", e);
    }
}

/**
 * Restaura el modo de cálculo anterior. SIEMPRE se llama desde un
 * finally (ver actualizarInformeCore/actualizarInformeFixedCore/
 * actualizarTodosCore) para que, aunque el refresco falle a mitad de
 * camino, el libro nunca se quede encallado en Manual — eso sería mucho
 * peor que el problema que se está arreglando (dejaría de recalcularse
 * TODO lo demás del libro sin que el usuario supiera por qué).
 */
async function draco_endManualCalculation() {
    DracoCalcModeDepth = Math.max(0, DracoCalcModeDepth - 1);
    if (DracoCalcModeDepth > 0) return; // todavía dentro de otra llamada exterior (p.ej. "Actualizar todos")
    const prevMode = DracoCalcModePrevMode;
    DracoCalcModePrevMode = null;
    if (!prevMode || prevMode === Excel.CalculationMode.manual) return; // no había nada que restaurar
    try {
        await Excel.run(async (context) => {
            context.workbook.application.calculationMode = prevMode;
            await context.sync();
        });
    } catch (e) {
        console.warn("[Draco] No se pudo restaurar el modo de cálculo:", e);
    }
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
    // Mantiene al día las columnas de "ID y texto" (y las jerarquías) con el
    // modelo semántico actual aunque el informe no se haya vuelto a guardar.
    if (window.ReportStore && reportId && window.ReportStore.refreshExpansion) {
        try { await window.ReportStore.refreshExpansion(reportId); }
        catch (e) { console.warn("[Draco] No se pudo actualizar la expansión del diseño:", e); }
    }
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
    // Medidas arrastradas a la zona "Filtros": no son un filtro real (no
    // hay una dimensión/tabla real que consultar para sus miembros, ver
    // también el cambio en taskpane.js que ya no permite doble clic sobre
    // ellas), así que NO se añaden a ReportState.Filters — pero tampoco se
    // descartan sin más: se recogen aquí para que el informe DINÁMICO
    // (applyFilterMeasuresToDynamicReport, llamado tras
    // loadReportDefinition) las añada al final de ReportState.Measures,
    // como si estuvieran en la última posición de Columnas. El modo Fijo
    // (buildSQLFixed) no las usa en absoluto.
    ReportState.FilterMeasureNames = [];

    let R = 15;

    while (String(cellValue(editReportGrid, R, 3)).trim() !== "") {
        const dimension = String(cellValue(editReportGrid, R, 3)).trim();
        if (dimension.toUpperCase() === "MEASURE") {
            ReportState.FilterMeasureNames.push(String(cellValue(editReportGrid, R, 4)).trim());
        } else {
            ReportState.FilterCount++;
            ReportState.Filters.push({
                Dimension: dimension,
                AttributeName: String(cellValue(editReportGrid, R, 4)).trim(),
                Value: String(cellValue(editReportGrid, R, 5)).trim()
            });
        }
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
        // Informe vinculado: hereda (bloqueados) los filtros del botón de su principal
        const linkMain = dracoLinkMainId(reportId);
        const appliesToThisReport = isAll
            || (reportId !== null && reportId !== undefined && Number(meta.reportId) === Number(reportId))
            || (linkMain !== null && Number(meta.reportId) === linkMain);
        if (!appliesToThisReport) continue;

        // Igual que en loadFilters(): una MEDIDA en un filtro "bloqueado"
        // (creado con "Añadir filtro") tampoco participa del WHERE.
        if (String(meta.dim).toUpperCase() === "MEASURE") {
            if (!ReportState.FilterMeasureNames) ReportState.FilterMeasureNames = [];
            ReportState.FilterMeasureNames.push(meta.name);
            continue;
        }

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
            // Columnas N/O/P de EDIT_REPORT = zona "columns" del taskpane
            // (ver ReportStore.getReportGrid: expandedCols -> N..S). La zona
            // hace falta para la clave de fieldOptions (formato/agregación).
            ReportState.Measures.push({
                Name: String(cellValue(editReportGrid, R, 15)).trim(),
                Zone: "columns"
            });
        } else {
            ReportState.RowCount++;

            // "Solo texto": el ID (orden = "HIDE", col. S) se consulta pero no se pinta,
            // así que no ocupa columna física.
            const hiddenRow = dracoIsHiddenMarker(cellValue(editReportGrid, R, 19));
            if (Number(cellValue(editReportGrid, R, 16)) === 1 && !hiddenRow) {
                ReportState.AttrRowCount++;
            }

            const hierRaw = String(cellValue(editReportGrid, R, 16)).trim();

            ReportState.Rows.push({
                Position: Number(cellValue(editReportGrid, R, 13)),
                Dimension: dim,
                AttributeName: String(cellValue(editReportGrid, R, 15)).trim(),
                Hierarchy: hierRaw === "" ? 0 : Number(hierRaw),
                // Columna Q: nombre de jerarquía; con sufijo "~TXT" es la
                // jerarquía de TEXTOS de un "ID y texto" (ver DRACO_HIER_TEXT_SUFFIX).
                IsTextCompanion: dracoIsHierTextName(cellValue(editReportGrid, R, 17)),
                Hidden: hiddenRow,
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
            // Columnas H/I/J de EDIT_REPORT = zona "rows" del taskpane (ver
            // ReportStore.getReportGrid: expandedRows -> H..M).
            ReportState.Measures.push({
                Name: String(cellValue(editReportGrid, R, 9)).trim(),
                Zone: "rows"
            });
        } else {
            const hiddenCol = dracoIsHiddenMarker(cellValue(editReportGrid, R, 13)); // columna M
            if (Number(cellValue(editReportGrid, R, 10)) === 1 && !hiddenCol) {
                ReportState.AttrColumnCount++;
            }

            ReportState.ColumnCount++;

            const hierRaw = String(cellValue(editReportGrid, R, 10)).trim();

            ReportState.Columns.push({
                Position: Number(cellValue(editReportGrid, R, 7)),
                Dimension: dim,
                AttributeName: String(cellValue(editReportGrid, R, 9)).trim(),
                Hierarchy: hierRaw === "" ? 0 : Number(hierRaw),
                // Columna K: nombre de jerarquía (ver loadRows / DRACO_HIER_TEXT_SUFFIX).
                IsTextCompanion: dracoIsHierTextName(cellValue(editReportGrid, R, 11)),
                Hidden: hiddenCol,
                // Columna L de EDIT_REPORT: "X" = generar subtotal para este
                // campo (ver buildConfigSetsWithSubtotals).
                Subtotal: String(cellValue(editReportGrid, R, 12)).trim().toUpperCase() === "X"
            });
        }

        R++;
    }
}

function loadReportDefinition(editReportGrid, reportId) {
    const effectiveReportId = reportId !== undefined ? reportId : activeReportIdOrNull();
    ReportState.ReportId = effectiveReportId;
    loadFilters(editReportGrid);
    appendLockedFilterRangesToState(effectiveReportId);
    loadRows(editReportGrid);
    loadColumns(editReportGrid);
    attachDracoMeasureOptions();
    attachDracoCalcs();
    applyDracoRuntimeFilters();
}

/**
 * Filtros de TIEMPO DE EJECUCIÓN (barra de filtros Draco de PLANNING, «Filtrar directamente»): una lista
 * [{ dim, attr, value }] en window.__dracoRuntimeFilters que se añade a los filtros del informe; si el informe ya
 * filtra el mismo campo, el de la barra lo sustituye. En Excel no se usa (la lista no existe).
 */
function applyDracoRuntimeFilters() {
    const list = (typeof window !== "undefined" && Array.isArray(window.__dracoRuntimeFilters)) ? window.__dracoRuntimeFilters : [];
    list.forEach(f => {
        if (!f || !f.dim || !f.attr || f.value === undefined || f.value === null) return;
        // f.rep: solo para ESE informe de un grupo vinculado (sin rep: para todos los informes de la tabla)
        if (f.rep !== undefined && f.rep !== null && f.rep !== "" && Number(f.rep) !== Number(ReportState.ReportId)) return;
        // f.value: un valor, o una lista (filtro de varios valores)
        const values = (Array.isArray(f.value) ? f.value : [f.value]).map(String).filter(v => v !== "");
        if (!values.length) return;
        const D = String(f.dim).toUpperCase(), A = String(f.attr).toUpperCase();
        ReportState.Filters = (ReportState.Filters || []).filter(x => !(String(x.Dimension).toUpperCase() === D && String(x.AttributeName).toUpperCase() === A));
        ReportState.Filters.push({ Dimension: f.dim, AttributeName: f.attr, Value: JSON.stringify({ mode: "list", values }) });
    });
}

/* ---------------------------------------------------------------------
 * CÁLCULOS ENTRE MIEMBROS (ver js/calcFormula.js)
 *
 * ReportState.Calcs: un elemento por atributo de Filas/Columnas que
 * tenga cálculos definidos en el informe:
 *   { dim, attr, axis: "Columns"|"Rows", idx, calcs: [calc + {ast, refs}] }
 * (axis "Columns" = columnas H.. de EDIT_REPORT = zona Filas del
 * taskpane, que se pinta en filas físicas; "Rows" = zona Columnas).
 * Solo se usan en el flujo Dinámico.
 * ------------------------------------------------------------------- */
function attachDracoCalcs() {
    ReportState.Calcs = [];
    ReportState.CalcOrder = {};
    if (!window.DracoCalc) return;

    let fieldOptions = {};
    try {
        const report = (window.ReportStore && ReportState.ReportId !== null && ReportState.ReportId !== undefined)
            ? window.ReportStore.getReport(ReportState.ReportId) : null;
        fieldOptions = (report && report.design && report.design.fieldOptions) || {};
    } catch (e) { fieldOptions = {}; }

    const seen = new Set();
    const collect = (fields, axis) => {
        fields.forEach((fld, idx) => {
            const k = String(fld.AttributeName).toUpperCase();
            if (seen.has(k)) return; // mismo atributo dos veces: solo el primero
            const calcs = window.DracoCalc.getCalcs(fieldOptions, fld.Dimension, fld.AttributeName)
                .map(c => {
                    const a = window.DracoCalc.analyze(c.expr, c.mode);
                    return a.ok ? Object.assign({}, c, { ast: a.ast, refs: a.refs }) : null;
                })
                .filter(Boolean);
            if (calcs.length === 0) return;
            seen.add(k);
            ReportState.Calcs.push({ dim: fld.Dimension, attr: fld.AttributeName, axis, idx, calcs });
        });
    };
    collect(ReportState.Columns, "Columns");
    collect(ReportState.Rows, "Rows");
}

// Literal de texto SQL según motor (BigQuery escapa con \, el resto
// duplicando la comilla).
function dracoSqlString(v) {
    const t = String(v);
    if (Provider.key && Provider.key() === "bigquery") {
        return "'" + t.replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'";
    }
    return "'" + t.replace(/'/g, "''") + "'";
}

// Parte del ORDER BY de ROW_ID/COLUMN_ID: un atributo con cálculos se
// ordena primero por su orden de cálculo (reales = 0, cálculos al final
// o al principio) y después por su valor ORIGINAL (no por el texto, que
// para atributos numéricos ordenaría "10" antes que "2").
function dracoCalcOrderPart(part) {
    const o = ReportState.CalcOrder && ReportState.CalcOrder[part];
    return o || part;
}

/**
 * PIEZAS DE UN KPI (Tablas Draco en PLANNING: bloque «KPI Draco» y $$NOMBRE$$ en los textos).
 * Un KPI es un informe con UNA medida y, como mucho, filtros (sin campos en Filas ni Columnas). En vez de un SQL
 * completo, devuelve sus piezas para que PLANNING junte TODOS los KPIs de un informe en una sola consulta con
 * agregación condicional (una lectura por tabla de hechos, ver DracoTables.kpiSql):
 *   { measure, agg: "SUM"|"AVG"|"COUNT"|"COUNT_DISTINCT"|"MIN"|"MAX"|<otra tal cual>, field: "f.CAMPO",
 *     from: "FROM <tabla de hechos> f", joins: "<JOINs de las dimensiones usadas>", conds: [<condiciones>] }
 * Las condiciones pueden llevar marcadores $VARIABLE$ (se sustituyen al ejecutar).
 * Requiere ReportState cargado (loadReportDefinition + applyFilterMeasuresToDynamicReport).
 */
function buildDracoKpiParts(relGrid, measuresGrid, atributesGrid) {
    const dims = [...ReportState.Columns, ...ReportState.Rows];
    if (dims.length) throw new Error("Un KPI no lleva campos en Filas ni en Columnas: quítalos y deja solo la medida y los filtros.");
    if (ReportState.Measures.length !== 1) throw new Error(ReportState.Measures.length ? "Un KPI tiene UNA sola medida: deja solo una en «Filtros»." : "Añade la medida del KPI: arrástrala a «Filtros».");
    const m = ReportState.Measures[0];
    const R = buscarMedida(measuresGrid, m.Name);
    if (R === 0) throw new Error(`La medida «${m.Name}» no está en el modelo semántico.`);
    const field = String(cellValue(measuresGrid, R, 6)).trim();
    const modelAgg = String(cellValue(measuresGrid, R, 7)).trim();
    const agg = normalizeDracoAggregation(m.Aggregation) || normalizeDracoAggregation(modelAgg) || modelAgg || "SUM";
    const parts = {
        measure: m.Name,
        agg,
        scale: m.Scale || 1,
        fmtKind: m.FmtKind || "",
        decimals: m.FmtDecimals,
        field: "f." + field,
        from: buildFrom(measuresGrid, m.Name, relGrid),
        joins: buildJoins(relGrid),
        conds: buildFilterConditions(atributesGrid, relGrid),
        provider: Provider.key()
    };
    // Agregación de excepción: columnas de referencia y JOIN que faltan (ver DracoTables.kpiSql en PLANNING)
    if (m.Exc) {
        const ref = dracoExcRef(relGrid, m.Exc);
        parts.exc = { agg: m.Exc.agg, cols: ref.cols, joins: ref.joins.map(j => j.trim()) };
    }
    return parts;
}

/**
 * CTEs de los cálculos, encadenadas detrás de REPORT_DATA. Devuelve
 * { sql: ",\nDRC_...", last: "nombre de la última CTE" } o null.
 *
 * Por cada atributo con cálculos (paso i):
 *   - DRC_HID_i (solo si hay cálculos por miembro): la misma consulta que
 *     REPORT_DATA pero quitando los filtros de la dimensión del atributo
 *     y quedándose solo con los miembros referenciados. Así un cálculo
 *     fijo funciona aunque esos miembros no estén en el filtro.
 *   - DRC_SRC_i: REPORT_DATA (o el paso anterior) marcado 'P' + DRC_HID_i
 *     marcado 'H'.
 *   - DRC_STEP_i: las filas del paso anterior (el atributo pasa a texto) +
 *     una fila por cálculo y combinación del resto de campos, con:
 *       · por miembro: SUM(CASE WHEN atributo = 'X' ...) operados según
 *         la fórmula (miembro visible si está; si no, el leído aparte).
 *       · por posición: NULL (el valor lo pone la fórmula de Excel).
 *     DRC_S_i guarda el valor original (orden) y DRC_O_i la posición del
 *     cálculo (0 = miembro real).
 */
function buildDracoCalcCtes(relGrid, measuresGrid, atributesGrid, withSubtotals) {
    const specs = ReportState.Calcs || [];
    if (specs.length === 0) return null;

    const STR = Provider.stringType();
    const dimCols = [...ReportState.Columns.map(c => c.AttributeName), ...ReportState.Rows.map(r => r.AttributeName)];
    const flagCols = withSubtotals
        ? [...ReportState.Columns.map(c => subtotalFlagName(c.AttributeName)), ...ReportState.Rows.map(r => subtotalFlagName(r.AttributeName))]
        : [];
    const measureCols = ReportState.Measures.map(m => m.Name);
    const extras = [];           // [{attr, s, o}] de pasos anteriores
    const castedAttrs = new Set(); // atributos ya convertidos a texto
    let prev = "REPORT_DATA";
    let sql = "";

    specs.forEach((spec, n) => {
        const i = n + 1;
        const A = spec.attr;
        const flagA = withSubtotals ? subtotalFlagName(A) : null;
        const S = "DRC_S" + i, O = "DRC_O" + i;
        const curCols = [...dimCols, ...flagCols, ...extras.flatMap(e => [e.s, e.o]), ...measureCols];

        const memberRefs = [];
        spec.calcs.filter(c => c.mode === "member").forEach(c => c.refs.forEach(r => {
            if (memberRefs.indexOf(String(r)) === -1) memberRefs.push(String(r));
        }));
        const hasH = memberRefs.length > 0;
        let src = prev;

        if (hasH) {
            // --- DRC_HID_i: consulta auxiliar con los miembros referenciados
            const savedFilters = ReportState.Filters;
            let conds;
            try {
                ReportState.Filters = (savedFilters || []).filter(f =>
                    String(f.Dimension).toUpperCase() !== String(spec.dim).toUpperCase());
                conds = buildFilterConditions(atributesGrid, relGrid);
            } finally {
                ReportState.Filters = savedFilters;
            }
            const tipo = getAttributeType(atributesGrid, spec.dim, A);
            const lits = memberRefs.map(r => UNQUOTED_TYPES.includes(tipo) && !isNaN(Number(r)) ? String(Number(r)) : dracoSqlString(r));
            conds.push(getTableAlias(relGrid, spec.dim) + "." + A + " IN (" + lits.join(", ") + ")");

            sql += "," + CRLF + "DRC_HID_" + i + " AS (" + CRLF
                + (withSubtotals ? buildSelectWithSubtotals(measuresGrid, relGrid) : buildSelect(measuresGrid, relGrid)) + CRLF
                + buildFrom(measuresGrid, null, relGrid) + CRLF
                + buildJoins(relGrid) + CRLF
                + "WHERE" + CRLF + conds.join(CRLF + "AND ") + CRLF
                + (withSubtotals ? buildConfigSetsWithSubtotals(relGrid) : buildConfigSets(relGrid)) + CRLF
                + ")";

            // --- DRC_SRC_i
            const fromPrev = curCols.map(c => "p." + c + " AS " + c).join(", ");
            const fromHid = curCols.map(c => {
                const e = extras.find(x => x.s === c || x.o === c);
                if (e) return (c === e.s ? "h." + e.attr : "0") + " AS " + c;
                if (castedAttrs.has(c)) return "CAST(h." + c + " AS " + STR + ") AS " + c;
                return "h." + c + " AS " + c;
            }).join(", ");
            sql += "," + CRLF + "DRC_SRC_" + i + " AS (" + CRLF
                + "    SELECT 'P' AS DRC_SRC, " + fromPrev + " FROM " + prev + " p" + CRLF
                + "    UNION ALL" + CRLF
                + "    SELECT 'H' AS DRC_SRC, " + fromHid + " FROM DRC_HID_" + i + " h" + CRLF
                + ")";
            src = "DRC_SRC_" + i;
        }

        // --- DRC_STEP_i
        const realBranch = "    SELECT "
            + curCols.map(c => (c === A ? "CAST(p." + A + " AS " + STR + ")" : "p." + c) + " AS " + c).join(", ")
            + ", p." + A + " AS " + S + ", 0 AS " + O + CRLF
            + "    FROM " + prev + " p";

        const groupCols = curCols.filter(c => c !== A && c !== flagA && measureCols.indexOf(c) === -1);
        const memberCond = (r) => "CAST(s." + A + " AS " + STR + ") = " + dracoSqlString(r);

        const calcBranches = spec.calcs.map((calc, ci) => {
            const ord = 1 + ci; // los cálculos van siempre detrás de los miembros
            const isMember = calc.mode === "member";
            const cols = curCols.map(c => {
                if (c === A) return dracoSqlString(calc.name) + " AS " + c;
                if (c === flagA) return "0 AS " + c;
                if (measureCols.indexOf(c) !== -1) {
                    if (!isMember) return "NULL AS " + c;
                    const expr = window.DracoCalc.toSql(calc.ast, (r) => {
                        const vis = "SUM(CASE WHEN " + (hasH ? "s.DRC_SRC = 'P' AND " : "") + memberCond(r) + " THEN s." + c + " END)";
                        if (!hasH) return "COALESCE(" + vis + ", 0)";
                        const hid = "SUM(CASE WHEN s.DRC_SRC = 'H' AND " + memberCond(r) + " THEN s." + c + " END)";
                        return "COALESCE(" + vis + ", " + hid + ", 0)";
                    });
                    return expr + " AS " + c;
                }
                return "s." + c + " AS " + c;
            });
            const where = ["s." + A + " IS NOT NULL"];
            if (!isMember && hasH) where.push("s.DRC_SRC = 'P'");
            const having = (isMember && calc.refs.length)
                ? "HAVING COUNT(CASE WHEN CAST(s." + A + " AS " + STR + ") IN (" + calc.refs.map(r => dracoSqlString(r)).join(", ") + ") THEN 1 END) > 0"
                : "";
            return "    SELECT " + cols.join(", ") + ", NULL AS " + S + ", " + ord + " AS " + O + CRLF
                + "    FROM " + src + " s" + CRLF
                + "    WHERE " + where.join(" AND ") + CRLF
                + (groupCols.length ? "    GROUP BY " + groupCols.map(c => "s." + c).join(", ") + CRLF : "")
                + (having ? "    " + having + CRLF : "");
        });

        sql += "," + CRLF + "DRC_STEP_" + i + " AS (" + CRLF
            + realBranch + CRLF
            + calcBranches.map(b => "    UNION ALL" + CRLF + b).join("")
            + ")";

        extras.push({ attr: A, s: S, o: O });
        castedAttrs.add(A);
        ReportState.CalcOrder[A] = O + ", " + S + ", " + A;
        prev = "DRC_STEP_" + i;
    });

    return { sql, last: prev };
}

/* ---------------------------------------------------------------------
 * Opciones de campo de las MEDIDAS (panel "Opciones de campo" del
 * taskpane, ver buildMeasureOptionsForm en taskpane.js): formato de
 * número (Número / Porcentaje + nº de decimales) y agregación. Viven por
 * informe en report.design.fieldOptions con clave "zona|MEASURE|nombre".
 * Se cuelgan de cada ReportState.Measures[k] como .Aggregation (para el
 * SQL/DAX) y .NumberFormat (para el pintado), así buildSQL/
 * jsonTo3Matrices no necesitan saber de qué informe se trata.
 * ------------------------------------------------------------------- */

// Agregaciones que ofrece el desplegable. Todas funcionan tal cual en el
// flujo Dinámico: GROUPING SETS calcula cada nivel (detalle, subtotales y
// nodos de jerarquía) directamente sobre las filas de la tabla de hechos,
// nunca agregando agregados, así que AVG / COUNT DISTINCT dan el valor
// correcto también en los subtotales.
const DRACO_SQL_AGGREGATIONS = {
    SUM: field => "SUM(" + field + ")",
    AVG: field => "AVG(" + field + ")",
    COUNT: field => "COUNT(" + field + ")",
    COUNT_DISTINCT: field => "COUNT(DISTINCT " + field + ")",
    MIN: field => "MIN(" + field + ")",
    MAX: field => "MAX(" + field + ")"
};

function normalizeDracoAggregation(value) {
    const v = String(value || "").trim().toUpperCase().replace(/\s+/g, "_");
    if (v === "AVERAGE") return "AVG";
    if (v === "COUNTDISTINCT" || v === "DISTINCTCOUNT") return "COUNT_DISTINCT";
    return DRACO_SQL_AGGREGATIONS[v] ? v : "";
}

function dracoMeasureFieldOptions(reportId, zone, measureName) {
    try {
        const report = (window.ReportStore && reportId !== null && reportId !== undefined)
            ? window.ReportStore.getReport(reportId) : null;
        const all = (report && report.design && report.design.fieldOptions) || {};
        return all[`${zone}|MEASURE|${measureName}`] || null;
    } catch (e) {
        return null;
    }
}

/**
 * Formato de Excel para una medida a partir de sus opciones:
 *   - "Número":     #,##0  /  #,##0.00 ...  (separador de miles)
 *   - "Porcentaje": 0%     /  0.00% ...
 * con tantos decimales como diga "Nº de decimales". numberFormat (no
 * numberFormatLocal) usa siempre la sintaxis inglesa ("," miles, "."
 * decimales) y Excel muestra los separadores del idioma del usuario, por
 * eso ya no hace falta elegir separador decimal. Devuelve null si la
 * medida no tiene opciones guardadas (se deja el formato general).
 *
 * Compatibilidad: informes guardados con la versión anterior del panel
 * tienen numberFormat "#,##0", "#,##0.00", "0.00%" o "#,##0.00 €": todo lo
 * que lleva "%" se trata como porcentaje y el resto como número.
 */
function dracoNumberFormatFromOptions(opts) {
    if (!opts) return null;
    const kindRaw = String(opts.numberFormat || "number");
    const kind = (kindRaw === "percent" || kindRaw.indexOf("%") !== -1) ? "percent" : "number";
    let decimals = Number(opts.decimals);
    if (!isFinite(decimals) || decimals < 0) decimals = kind === "percent" ? 0 : 2;
    decimals = Math.min(10, Math.floor(decimals));
    const dec = decimals > 0 ? "." + "0".repeat(decimals) : "";
    return kind === "percent" ? "0" + dec + "%" : "#,##0" + dec;
}

function decorateDracoMeasure(m) {
    if (!m) return;
    if (!m.Zone) m.Zone = "filters";
    const opts = dracoMeasureFieldOptions(ReportState.ReportId, m.Zone, m.Name);
    m.Aggregation = opts ? normalizeDracoAggregation(opts.aggregation) : "";
    m.NumberFormat = dracoNumberFormatFromOptions(opts);
    // Factor de escala (Opciones de campo): el valor se divide en la consulta (1 = sin escala)
    const sc = opts ? Number(opts.scale) : 1;
    m.Scale = (isFinite(sc) && sc > 1) ? sc : 1;
    // Formato "en bruto" (para los KPIs de PLANNING: número / porcentaje y decimales)
    m.FmtKind = opts ? ((String(opts.numberFormat || "").indexOf("%") !== -1 || opts.numberFormat === "percent") ? "percent" : "number") : "";
    m.FmtDecimals = (opts && isFinite(Number(opts.decimals))) ? Number(opts.decimals) : null;
    // Agregación de excepción (null si no tiene)
    m.Exc = normalizeDracoExc(opts);
}

// "(expr) / factor" si hay factor de escala
function dracoScaled(expr, scale) {
    const f = Number(scale);
    return (isFinite(f) && f > 1) ? "(" + expr + ") / " + f : expr;
}

/**
 * Factor de escala de una medida por su NOMBRE. Sirve para los modos Fijo e Híbrido, donde la medida puede no estar
 * en ReportState.Measures (la pintada como MEASURE en el modo Fijo, o la de «Filtros», que solo se añade a Measures
 * en el modo dinámico): primero la de ReportState.Measures (ya decorada) y si no, sus Opciones de campo en cualquier
 * zona (filtros, columnas, filas). 1 = sin escala.
 */
function dracoMeasureScale(measureName) {
    if (!measureName) return 1;
    const up = String(measureName).toUpperCase();
    const m = (ReportState.Measures || []).find(x => String(x.Name).toUpperCase() === up);
    if (m && Number(m.Scale) > 1) return Number(m.Scale);
    let all = {};
    try {
        const report = (window.ReportStore && ReportState.ReportId !== null && ReportState.ReportId !== undefined)
            ? window.ReportStore.getReport(ReportState.ReportId) : null;
        all = (report && report.design && report.design.fieldOptions) || {};
    } catch (e) { all = {}; }
    for (const zone of ["filters", "columns", "rows"]) {
        const key = Object.keys(all).find(k => k.toUpperCase() === (zone + "|MEASURE|" + up).toUpperCase());
        const sc = key && all[key] ? Number(all[key].scale) : 1;
        if (isFinite(sc) && sc > 1) return sc;
    }
    return 1;
}

/* ---------------------------------------------------------------------
 * AGREGACIÓN DE EXCEPCIÓN (como en BPC), en las Opciones de campo de una medida:
 *   excAggregation: "MIN" | "MAX" | "AVG" | "SUM" | "COUNT"
 *   excDims:        [{ dim, attr }]   dimensiones de referencia
 * La medida se agrega primero (con su agregación normal) POR CADA VALOR de las dimensiones de referencia
 * —y por todo lo demás como siempre— y sobre esos valores se aplica la agregación de excepción.
 * Ej.: mínimo mensual = referencia «tiempo · mes» + MIN.
 * Si una medida no la tiene, el SQL que se genera es EXACTAMENTE el de siempre.
 * ------------------------------------------------------------------- */
const DRACO_EXC_AGGS = ["MIN", "MAX", "AVG", "SUM", "COUNT"];

function normalizeDracoExc(opts) {
    if (!opts) return null;
    const agg = String(opts.excAggregation || "").toUpperCase();
    const dims = Array.isArray(opts.excDims) ? opts.excDims.filter(d => d && d.dim && d.attr) : [];
    return (DRACO_EXC_AGGS.indexOf(agg) !== -1 && dims.length) ? { agg, dims: dims.map(d => ({ dim: String(d.dim), attr: String(d.attr) })) } : null;
}

/** Agregación de excepción de una medida por su NOMBRE (la de ReportState.Measures o sus Opciones de campo). */
function dracoMeasureExc(measureName) {
    if (!measureName) return null;
    const up = String(measureName).toUpperCase();
    const m = (ReportState.Measures || []).find(x => String(x.Name).toUpperCase() === up);
    if (m && m.Exc) return m.Exc;
    let all = {};
    try {
        const report = (window.ReportStore && ReportState.ReportId !== null && ReportState.ReportId !== undefined)
            ? window.ReportStore.getReport(ReportState.ReportId) : null;
        all = (report && report.design && report.design.fieldOptions) || {};
    } catch (e) { all = {}; }
    for (const zone of ["filters", "columns", "rows"]) {
        const key = Object.keys(all).find(k => k.toUpperCase() === (zone + "|MEASURE|" + up).toUpperCase());
        const exc = key ? normalizeDracoExc(all[key]) : null;
        if (exc) return exc;
    }
    return null;
}

/** JOIN de UNA dimensión, con el mismo texto que genera buildJoins para ella. */
function dracoJoinForDimension(relGrid, dimension) {
    const lastRow = lastRowInColumnValues(relGrid, 1);
    for (let R = 2; R <= lastRow; R++) {
        if (String(cellValue(relGrid, R, 2)).trim().toUpperCase() !== String(dimension).trim().toUpperCase()) continue;
        return CRLF
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
    return "";
}

/** Columnas de las dimensiones de referencia y los JOIN que faltan (los de dimensiones que el informe no usa). */
function dracoExcRef(relGrid, exc) {
    const cols = [], joins = [];
    exc.dims.forEach(d => {
        const c = getTableAlias(relGrid, d.dim) + "." + d.attr;
        if (cols.indexOf(c) === -1) cols.push(c);
        if (!dimensionIsUsed(d.dim)) {
            const j = dracoJoinForDimension(relGrid, d.dim);
            if (j && joins.indexOf(j) === -1) joins.push(j);
        }
    });
    return { cols, joins };
}

/** "expr" de una medida sin su " AS nombre" final (getMeasureSQL). */
function dracoMeasureExpr(measuresGrid, m) {
    const full = getMeasureSQL(measuresGrid, m.Name, m.Aggregation, m.Scale);
    const i = full.lastIndexOf(" AS ");
    return i === -1 ? full : full.slice(0, i);
}

function attachDracoMeasureOptions() {
    (ReportState.Measures || []).forEach(decorateDracoMeasure);
}

/**
 * SOLO para el informe DINÁMICO (buildSQL/actualizarInformeCore,
 * actualizarTodosCore): añade al final de ReportState.Measures las
 * medidas que se hubieran arrastrado a la zona "Filtros" (tanto locales
 * del taskpane como filtros "bloqueados" de Añadir filtro), recogidas por
 * loadFilters()/appendLockedFilterRangesToState() en
 * ReportState.FilterMeasureNames. Se llama DESPUÉS de
 * loadReportDefinition() (que ya ha corrido loadColumns(), la última
 * función que toca ReportState.Measures), así que quedan siempre en
 * última posición — "como si estuvieran en la última posición de
 * Columnas". No se llama en ningún punto del modo Fijo
 * (buildSQLFixed/actualizarInformeFixedCore): ese flujo no se toca en
 * absoluto, tal y como estaba.
 */
function applyFilterMeasuresToDynamicReport() {
    (ReportState.FilterMeasureNames || []).forEach(name => {
        if (!name) return;
        ReportState.MeasureCount = (ReportState.MeasureCount || 0) + 1;
        const m = { Name: name, Zone: "filters" };
        decorateDracoMeasure(m);
        ReportState.Measures.push(m);
    });
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

// Tipos DATA_TYPE (MODEL_ATRIBUTES) que NO van entre comillas en SQL —
// mismo criterio "isNumeric" que ya usa semantic_model.js al pintar el
// editor de campos, para que ambos sitios coincidan.
const DRACO_PLANNING_NUMERIC_TYPES = ["INTEGER", "FLOAT", "NUMERIC", "BIGNUMERIC"];

/**
 * DATA_TYPE (MODEL_ATRIBUTES) de la dimensión "dim" tal como se usa en el
 * writeback de Draco Planning: a diferencia de getAttributeType (que
 * necesita DIMENSION + ATRIBUTO exactos), aquí solo se conoce el nombre de
 * la dimensión tal como aparece en Filtros/Filas/Columnas — para una
 * dimensión "plana" de un único atributo (caso típico de las columnas de
 * DRACO_PLANIFICACION: CANAL, CUENTA, PRODUCTO...) el nombre de esa
 * dimensión Y el de su atributo suelen coincidir, así que se prueba esa
 * coincidencia exacta primero; si no la hay, se cae al atributo marcado
 * como IS_KEY ("X") de esa dimensión. Devuelve "STRING" si no se
 * encuentra nada (mismo comportamiento por defecto que getAttributeType).
 */
function resolveDracoPlanningDimensionType(atributesGrid, dim) {
    const lastRow = lastRowInColumnValues(atributesGrid, 1);
    const dimUpper = String(dim).trim().toUpperCase();
    let keyRowType = null;

    for (let R = 2; R <= lastRow; R++) {
        const rowDim = String(cellValue(atributesGrid, R, 2)).trim().toUpperCase();
        if (rowDim !== dimUpper) continue;

        const rowAttr = String(cellValue(atributesGrid, R, 3)).trim().toUpperCase();
        const rowType = String(cellValue(atributesGrid, R, 9)).trim().toUpperCase();
        const isKey = String(cellValue(atributesGrid, R, 10)).trim().toUpperCase() === "X";

        if (rowAttr === dimUpper) return rowType; // coincidencia exacta atributo==dimensión: la más fiable
        if (isKey && keyRowType === null) keyRowType = rowType;
    }

    return keyRowType !== null ? keyRowType : "STRING";
}

function dracoPlanningValueIsNumericType(dataType) {
    return DRACO_PLANNING_NUMERIC_TYPES.includes(String(dataType || "").trim().toUpperCase());
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

            // Solo las EPM_VALUE son condiciones de la línea; una EPM_FORMULA
            // (línea calculada) u otra fórmula del usuario no entra en el SQL.
            if (cellHasFormula(csvGrid, R, Col) && /EPM_VALUE\s*\(/i.test(String(cellFormula(csvGrid, R, Col)))) {

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

            // Solo las EPM_VALUE son condiciones de la línea; una EPM_FORMULA
            // (línea calculada) u otra fórmula del usuario no entra en el SQL.
            if (cellHasFormula(csvGrid, R, Col) && /EPM_VALUE\s*\(/i.test(String(cellFormula(csvGrid, R, Col)))) {

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

// measureField: nombre real del campo en la tabla de hechos (FACT_FIELD
// de MODEL_MEASURES) para la medida encontrada — pintada como MEASURE,
// filtrada, o del diseño de EDIT_REPORT, en ese orden (ver buildSQLFixed).
// Sin ningún literal por defecto: si no se encuentra ninguna medida en
// ningún sitio, se deja tal cual venga (SUM(f.) sería un SQL claramente
// inválido, y así se detecta a simple vista en vez de sumar en silencio
// un campo que no es).
function buildSelectBase(relGrid, rowsDefs, colDefs, measureField, scale, excCols) {
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
    // agregación de excepción: dimensiones de referencia (DRC_R_j)
    (excCols || []).forEach((c, j) => { sql += "    " + c + " AS DRC_R_" + j + "," + CRLF; });
    sql += "    " + dracoScaled("SUM(f." + (measureField || "") + ")", scale) + " AS IMPORTE";

    return sql;
}

// measureNameOverride: nombre de medida a usar para el FROM (ver
// buildSQLFixed) en vez de ReportState.Measures[0].Name — para que el
// modo Fijo use la MISMA medida (y por tanto la misma tabla) que la
// celda MEASURE realmente pintada, no necesariamente la primera del
// diseño en EDIT_REPORT. Si no se pasa, se mantiene el comportamiento de
// siempre (usado tal cual por el modo Dinámico).
function buildFrom(measuresGrid, measureNameOverride, relGrid) {
    const firstMeasure = ReportState.Measures && ReportState.Measures[0];
    const measureName = measureNameOverride || (firstMeasure && firstMeasure.Name);

    // Informe SIN medidas (solo dimensiones en Filas y/o Columnas): no hay
    // medida de la que sacar la tabla de hechos, así que se usa la de la
    // relación (MODEL_RELATIONSHIP: FACT_PROJECT/DATASET/TABLE) de la
    // primera dimensión usada — la misma tabla "f" contra la que
    // buildJoins hace los JOIN, así que se listan los miembros que tienen
    // datos, igual que cuando hay medida.
    if (!measureName && relGrid) {
        const lastRow = lastRowInColumnValues(relGrid, 1);
        for (let R = 2; R <= lastRow; R++) {
            if (dimensionIsUsed(cellValue(relGrid, R, 2))) {
                return "FROM " + Provider.qualify(
                    cellValue(relGrid, R, 3),
                    cellValue(relGrid, R, 4),
                    cellValue(relGrid, R, 5)
                ) + " f";
            }
        }
        throw new Error("No se encuentra la tabla de hechos de las dimensiones del informe (MODEL_RELATIONSHIP).");
    }
    if (!measureName) {
        throw new Error("El informe no tiene ninguna medida ni dimensión que consultar.");
    }

    const R = buscarMedida(measuresGrid, measureName);

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

function buildGroupByBase(relGrid, rowsDefs, colDefs, excCols) {
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

    (excCols || []).forEach((c, j) => { if (!dict.has("DRC_R|" + j)) dict.set("DRC_R|" + j, c); });

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

    // BigQuery
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
    const rowsDefsRaw = await readRowDefinitions(context, editReportGrid, csvGrid);
    const colDefsRaw = await readColumnDefinitions(context, editReportGrid, csvGrid);

    // Diagnóstico: abre las herramientas de desarrollador (F12) del panel de
    // tareas / comandos para ver exactamente qué Dimension/Attribute/Value
    // se ha leído de las fórmulas EPM_VALUE de CSV_RESULT.
    console.log("readRowDefinitions ->", rowsDefsRaw);
    console.log("readColumnDefinitions ->", colDefsRaw);

    // La celda MEASURE pintada (Paso 1/2: =EPM_VALUE("MEASURE","<nombre>","","<nombre>"))
    // no es una condición de filtro para ROW_ID/COLUMN_ID: es el nombre de
    // la medida a usar en el SELECT. Puede estar en Filas o en Columnas
    // (nunca en ambas a la vez), así que se busca en las dos; el resto de
    // rowsDefs/colDefs (los que SÍ son dimensiones reales) se queda igual.
    const isMeasureDef = (d) => String(d.Dimension).toUpperCase() === "MEASURE";
    const measureDef = rowsDefsRaw.find(isMeasureDef) || colDefsRaw.find(isMeasureDef);
    const rowsDefs = rowsDefsRaw.filter((d) => !isMeasureDef(d));
    const colDefs = colDefsRaw.filter((d) => !isMeasureDef(d));

    // Nombre de la medida, por orden de prioridad:
    //   1) la celda MEASURE pintada (Filas o Columnas);
    //   2) una medida arrastrada a la zona "Filtros" (ReportState.
    //      FilterMeasureNames, ver loadFilters/appendLockedFilterRangesToState
    //      — no participa del WHERE, pero sigue siendo la medida del
    //      informe si es la única que hay);
    //   3) la primera del diseño de EDIT_REPORT (ReportState.Measures[0],
    //      mismo criterio que ya usaba buildFrom).
    // Con eso se busca en MODEL_MEASURES el campo REAL de la tabla de
    // hechos (FACT_FIELD, columna F): el nombre de la medida (p.ej.
    // "IMPORTE") no tiene por qué coincidir con el nombre físico de la
    // columna. Si no se encuentra ahí, se usa el propio nombre de medida
    // tal cual.
    const measureName = (measureDef && measureDef.AttributeName)
        || (ReportState.FilterMeasureNames && ReportState.FilterMeasureNames[0])
        || (ReportState.Measures[0] && ReportState.Measures[0].Name);
    const measureRow = measureName ? buscarMedida(measuresGrid, measureName) : 0;
    const measureField = (measureRow && String(cellValue(measuresGrid, measureRow, 6)).trim())
        || measureName;

    const rowIdsArray = buildIdArray(atributesGrid, rowsDefs, "ROW_ID");
    const columnIdsArray = buildIdArray(atributesGrid, colDefs, "COLUMN_ID");

    // Filtros del informe (zona "Filtros") + filtros "bloqueados" (Añadir
    // filtro, tanto los de ESTE informe como los globales de "Todos los
    // informes") — ReportState.Filters ya los trae todos juntos (los
    // carga loadReportDefinition, llamado antes de esta función). Se
    // combinan con AND justo detrás del bloque OR de buildBaseWhere: ese
    // OR ya poda a las filas que podrían encajar en ALGUNA cabecera de
    // fila/columna (optimización para no escanear toda la tabla de
    // hechos); los filtros de aquí son una condición añadida de negocio,
    // así que van en AND, no dentro del propio OR.
    const filterConditions = buildFilterConditions(atributesGrid, relGrid);

    // Agregación de excepción de la medida: BASE lleva también las dimensiones de referencia (DRC_R_j) y cada
    // celda se calcula en dos pasos (suma por celda y referencia; después la excepción por celda).
    const exc = dracoMeasureExc(measureName);
    const excRef = exc ? dracoExcRef(relGrid, exc) : null;
    const excCols = excRef ? excRef.cols : [];
    const rNames = excCols.map((c, j) => "DRC_R_" + j);

    let sql = "";

    // ---- CTE BASE: se escanea la tabla de hechos UNA sola vez ----
    sql += "WITH BASE AS (" + CRLF + CRLF;
    sql += buildSelectBase(relGrid, rowsDefs, colDefs, measureField, dracoMeasureScale(measureName), excCols) + CRLF + CRLF;
    sql += buildFrom(measuresGrid, measureName) + CRLF + CRLF;
    sql += buildJoins(relGrid) + (excRef ? excRef.joins.join("") : "") + CRLF + CRLF;
    sql += buildBaseWhere(atributesGrid, relGrid, rowsDefs, colDefs);
    for (const cond of filterConditions) {
        sql += CRLF + "AND " + cond;
    }
    sql += CRLF + CRLF;
    sql += buildGroupByBase(relGrid, rowsDefs, colDefs, excCols) + CRLF;
    sql += ")," + CRLF + CRLF;

    // ---- CTE TAGGED: etiqueta cada fila de BASE con los ROW_ID/COLUMN_ID
    //      que cumple, sin volver a escanear BASE por cada uno ----
    sql += "TAGGED AS (" + CRLF + CRLF;
    sql += "    SELECT" + CRLF;
    sql += "        IMPORTE," + CRLF + CRLF;
    rNames.forEach(n => { sql += "        " + n + "," + CRLF; });
    if (rNames.length) sql += CRLF;
    sql += indentBlock(rowIdsArray, "        ") + " AS ROW_IDS," + CRLF + CRLF;
    sql += indentBlock(columnIdsArray, "        ") + " AS COLUMN_IDS" + CRLF + CRLF;
    sql += "    FROM BASE" + CRLF;
    sql += ")" + CRLF + CRLF;

    // ---- SELECT final: cruce ROW_ID x COLUMN_ID + agregación ----
    // (con agregación de excepción: suma por celda y dimensiones de referencia y, encima, la excepción)
    if (exc) sql += "SELECT" + CRLF + "    ROW_ID," + CRLF + "    COLUMN_ID," + CRLF + "    " + exc.agg + "(IMPORTE) AS IMPORTE" + CRLF + "FROM (" + CRLF;
    sql += "SELECT" + CRLF;
    sql += "    ROW_ID," + CRLF;
    sql += "    COLUMN_ID," + CRLF;
    rNames.forEach(n => { sql += "    " + n + "," + CRLF; });
    sql += "    SUM(IMPORTE) AS IMPORTE" + CRLF;
    sql += "FROM TAGGED," + CRLF;
    sql += "UNNEST(ROW_IDS) AS ROW_ID," + CRLF;
    sql += "UNNEST(COLUMN_IDS) AS COLUMN_ID" + CRLF;
    if (exc) {
        sql += "GROUP BY ROW_ID, COLUMN_ID" + rNames.map(n => ", " + n).join("") + CRLF;
        sql += ") X" + CRLF;
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
 * Convierte filas ya normalizadas (array de objetos { COLUMNA: valor }, p. ej.
 * las de Provider.runQuery) al mismo formato de texto que devuelve BigQuery y
 * que ya saben leer parseJsonValueTriples() y jsonTo3MatricesCore(): ambos
 * escanean el texto buscando el literal `"v":` una vez por celda, en el mismo
 * orden fila a fila / columna a columna del SELECT. Lo usa Tablas Draco
 * (Planning), que ejecuta con su propia conexión.
 */
function rowsToPseudoBqJson(rows) {
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

/**
 * @param {string} sql
 * @param {string} [billingProjectId] Proyecto de facturación a usar para el job
 *   de BigQuery (jobs.query). Si no se indica, se usa BQ.getBillingProject()
 *   (config de la conexión) y, en su defecto, el proyecto de la tabla del SQL
 *   antes. Quien llama y SÍ sabe a qué proyecto pertenece la tabla destino
 *   (p.ej. executeDracoPlanningInserts, a partir del INSERT INTO `proyecto.dataset.tabla`)
 *   debe pasarlo aquí para que el job se facture ahí y no en un proyecto sin
 *   billing (ver executeDracoPlanningInserts).
 */
async function executeSQL(sql, billingProjectId) {
    // Variables ($NOMBRE$ en las líneas EPM_VALUE de un eje Estático, ver
    // js/variablesStore.js): el SQL se genera con el marcador y se sustituye
    // por el valor AQUÍ, justo antes de ejecutarlo. Así un SQL guardado (Tablas
    // Draco) vale para cualquier valor de la variable.
    if (typeof sql === "string" && window.DracoVariables) {
        sql = window.DracoVariables.resolveSql(sql, Provider.key());
    }

    // BigQuery
    const token = BQ.getToken();

    if (!token) {
        throw new Error("No hay una sesión activa de BigQuery. Inicia sesión en el panel primero.");
    }

    const projectId = BQ.jobProjectFor(sql, billingProjectId);

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
//
// NULL -> celda en blanco: BigQuery devuelve los NULL sin comillas
// ({"v": null}, con el JSON "bonito" -> el texto extraído empieza por
// "null" seguido de espacios/saltos de línea), y rowsToPseudoBqJson genera
// ese mismo literal. Antes se pintaba tal cual
// ("null" en la celda). Las cabeceras de Filas/Columnas ya descartaban esos
// valores antes de llegar aquí (indexOf("null") === 0), así que esto solo
// cambia el pintado de VALORES (medidas), que es lo que se buscaba.
function isDracoNullLiteral(text) {
    if (text === null || text === undefined) return true;
    const t = String(text).trim().toLowerCase();
    return t === "null" || t === "undefined";
}

/**
 * SEGURIDAD (auditoría F-15): las celdas de depuración de EDIT_REPORT
 * (W1 y Z1 de planificación, X1 con la SQL, Y1 con el JSON de resultados)
 * quedaban guardadas DENTRO del .xlsx y viajaban al compartir el libro.
 * Ahora solo se rellenan si DracoConfig.debugWorkbookCells = true; si no,
 * se escriben vacías (así además se borra lo que dejaron versiones anteriores).
 */
function draco_debugCell(value) {
    const on = typeof DracoConfig !== "undefined" && DracoConfig.debugWorkbookCells === true;
    return on ? value : "";
}

function coerceCellLiteral(text) {
    if (isDracoNullLiteral(text)) return "";
    text = String(text);
    if (text.trim() !== "" && !isNaN(text)) {
        return Number(text);
    }
    return draco_neutralizeFormula(text);
}

/**
 * SEGURIDAD (auditoría F-02, inyección de fórmulas): un texto que viene de
 * BigQuery y empieza por = + - @, tabulador o retorno se escribe con un
 * apóstrofo delante para que Excel lo trate como TEXTO y no lo ejecute como
 * fórmula (p. ej. =HYPERLINK(...) o =WEBSERVICE(...) metidos en un dato).
 * Solo afecta a los DATOS pintados; las fórmulas propias de Draco
 * (EPM_VALUE, EPM_FORMULA) se escriben por otras rutas y no pasan por aquí.
 */
function draco_neutralizeFormula(text) {
    if (typeof text !== "string" || text === "") return text;
    return /^[=+\-@\t\r]/.test(text) ? "'" + text : text;
}

// reportId -> Map("fila_columna" -> valor), con el valor que jsonPaintValues
// pintó en cada celda de Draco_XXX_Values en el ÚLTIMO refresco — es
// "la verdad" contra la que se compara al guardar planificación, para
// mandar al INSERT solo las celdas cuyo valor de verdad cambió (ver
// filterDracoPlanningRealChanges). Se limpia y rellena entero en cada
// refresco (jsonPaintValues), no se va acumulando entre refrescos.
//
// SOLO se guarda para informes de planificación (reportProperties.
// planningReport) — para el resto no hace falta esta "tabla" en memoria
// en absoluto, y así no se guarda ni se retiene sin necesidad.
const DracoPlanningBaselineValues = new Map();

// reportId -> Map("row_col" -> {dim, attr, value}), UNA entrada por cada
// celda de cabecera de Filas/Columnas (Draco_XXX_Rows / Draco_XXX_Cols)
// pintada en el ÚLTIMO refresco. "value" es el valor CANÓNICO tal como
// vino en el JSON de BigQuery (antes de coerceCellLiteral y de que Excel
// lo escriba en la celda) — la fuente de verdad para el writeback
// (INSERT), independientemente de lo que Excel decida guardar/mostrar en
// esa celda después (p.ej. reinterpretar "2025-08-01" como fecha y
// reformatearla según el locale). Se reconstruye entera en cada refresco
// (jsonTo3MatricesCore), igual que DracoPlanningBaselineValues. Ver
// dracoPlanningCrossItemsForCellDynamic, que la consulta en vez de leer
// la celda de Excel.
const DracoAxisCrossValues = new Map();

async function jsonPaintValues(context, json, reportId) {
    let tPerf = performance.now(); // instrumentación de tiempos, ver draco_perfMark
    const triples = parseJsonValueTriples(json);
    tPerf = draco_perfMark(reportId, `Parseo JSON manual (${triples.length} celdas)`, tPerf);
    const resultSheetName = await getDracoResultSheetName(context, reportId);
    await ensureDracoResultSheetExists(context, resultSheetName);
    const sheet = context.workbook.worksheets.getItem(resultSheetName);

    const report = window.ReportStore ? window.ReportStore.getReport(reportId) : null;
    const isPlanningReport = !!(report && report.reportProperties && report.reportProperties.planningReport);
    const baseline = isPlanningReport ? new Map() : null;

    // Antes: UN range.getRangeByIndexes(...).values=[[...]] POR CELDA (un
    // informe "Fijo"/planificación grande -p.ej. 500 filas x 24 meses =
    // 12.000 celdas, ver comentario en handleDracoPlanningValueChanged-
    // podía significar miles de objetos Range y miles de operaciones en
    // cola, aunque solo hubiera un context.sync() final): eso es lo que
    // de verdad hacía lento el pintado aquí, y al tardar tanto podía dar
    // la sensación de que el indicador "Actualizando" se quedaba colgado
    // sin ocultarse (en realidad seguía trabajando). Ahora se lee el
    // rectángulo que cubre todas las celdas de una vez (como antes hacía
    // writeCellBlock para los informes dinámicos), se cambian en memoria
    // solo las que trae el JSON -preservando cualquier fórmula/etiqueta
    // fija que el usuario tenga en los huecos, que en "Fijo" sí puede
    // haberla- y se escribe de una sola vez.
    if (triples.length > 0) {
        let minRow = Infinity, maxRow = -Infinity, minCol = Infinity, maxCol = -Infinity;
        for (const t of triples) {
            if (t.row < minRow) minRow = t.row;
            if (t.row > maxRow) maxRow = t.row;
            if (t.col < minCol) minCol = t.col;
            if (t.col > maxCol) maxCol = t.col;
        }
        const numRows = maxRow - minRow + 1;
        const numCols = maxCol - minCol + 1;

        // Igual que en writeCellBlock: si el rectángulo saliera enorme y
        // disperso (pocas celdas reales en un área gigante, poco probable
        // aquí pero por seguridad), se cae al camino celda a celda de
        // toda la vida en vez de reservar/leer un array gigantesco.
        if (numRows * numCols > 50000 && numRows * numCols > triples.length * 50) {
            for (const t of triples) {
                const range = sheet.getRangeByIndexes(t.row - 1, t.col - 1, 1, 1);
                const literal = coerceCellLiteral(t.text);
                range.values = [[literal]];
                if (baseline) baseline.set(t.row + "_" + t.col, literal);
            }
        } else {
            const range = sheet.getRangeByIndexes(minRow - 1, minCol - 1, numRows, numCols);
            range.load("values");
            await context.sync();

            const grid = range.values.map(r => r.slice());
            for (const t of triples) {
                const literal = coerceCellLiteral(t.text);
                grid[t.row - minRow][t.col - minCol] = literal;
                if (baseline) baseline.set(t.row + "_" + t.col, literal);
            }
            range.values = grid;
        }
    }

    if (isPlanningReport) {
        DracoPlanningBaselineValues.set(reportId, baseline);
    } else {
        DracoPlanningBaselineValues.delete(reportId); // por si dejó de ser de planificación desde el último refresco
    }

    await context.sync();
    draco_perfMark(reportId, "Escribir valores en la hoja", tPerf);
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
    beginDracoPerfLog();
    const tRefreshStart = performance.now(); // TOTAL de extremo a extremo, ver más abajo
    let tPerf = tRefreshStart; // instrumentación de tiempos, ver draco_perfMark

    await draco_beginManualCalculation();
    tPerf = draco_perfMark(reportId, "Poner cálculo en Manual", tPerf);

    // PLANIFICACIÓN > "refrescar" hace lo mismo que "Guardar
    // planificación": vuelca en EDIT_REPORT!A127 y devuelve su color
    // original a las celdas modificadas a mano pendientes, y vacía el
    // registro (ver flushDracoPlanningModifiedCells). Se hace ANTES del
    // refresco porque este va a borrar y repintar entero
    // Draco_<id>_Values de todas formas.
    await flushDracoPlanningModifiedCells();
    tPerf = draco_perfMark(reportId, "Vaciar cambios de planificación pendientes", tPerf);

    // "Actualizando…" (shape flotante sobre la hoja, ver busyindicator.js):
    // se muestra ANTES del primer Excel.run y solo se quita en el finally,
    // así cubre TODO el refresco (SQL, consulta y pintado), no solo la
    // parte visible. show()/hide() llevan su propio refCount, así que da
    // igual que taskpane.js (runAutoSaveAndRefresh) también lo muestre a
    // la vez: no se quita hasta que TODAS las llamadas a show() tengan su
    // hide(). if (window.BusyIndicator) por si esta página (commands.html
    // vs taskpane.html) no lo tiene cargado.
    if (window.BusyIndicator) await window.BusyIndicator.show(busyStepLabel(reportId, "Generando SQL"));
    tPerf = draco_perfMark(reportId, "Mostrar indicador (crear forma)", tPerf);

    // Ver comentario junto a DracoSuppressPlanningPaintCount: mientras
    // dura todo este refresco (que borra y reescribe Draco_<id>_Values),
    // handleDracoPlanningValueChanged no debe pintar esas celdas como si
    // fueran una edición manual del usuario.
    await beginSuppressPlanningPaint();
    tPerf = draco_perfMark(reportId, "Candado cruzado de planificación (begin)", tPerf);
    try {

    // 1) LoadReportDefinition + BuildSQL_Fixed + escritura de A1
    await Excel.run(async (context) => {
        let tSub = performance.now(); // sub-marcas dentro de "Generando SQL"
        const editReportGrid = await getEditReportGrid(context, reportId);
        tSub = draco_perfMark(reportId, "  Leer EDIT_REPORT (usedRange)", tSub);

        const relGrid = await window.SemanticModelStore.getModelGrid("MODEL_RELATIONSHIP");
        const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");
        const atributesGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
        const resultSheetName = resultSheetNameFromGrid(editReportGrid, reportId);
        await ensureDracoResultSheetExists(context, resultSheetName);
        const csvGrid = await getFormulaGrid(context, resultSheetName);
        tSub = draco_perfMark(reportId, "  Leer hoja de resultados (fórmulas)", tSub);

        loadReportDefinition(editReportGrid, reportId);

        sql = await buildSQLFixed(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid);
        tSub = draco_perfMark(reportId, "  loadReportDefinition + buildSQLFixed (JS puro)", tSub);

        await context.sync();
        draco_perfMark(reportId, "  Escribir/sincronizar", tSub);
    });
    tPerf = draco_perfMark(reportId, "Generando SQL", tPerf);

    // 2) ExecuteSQL contra BigQuery
    if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(reportId, "Ejecutando consulta"));
    const json = await executeSQL(sql);
    tPerf = draco_perfMark(reportId, "Ejecutando consulta (red)", tPerf);

    // [Punto 8] SQL y JSON generados ya NO se escriben en A1/B1 de la hoja
    // de resultados: se escriben en EDIT_REPORT!X1 (SQL) e Y1 (JSON).
    await Excel.run(async (context) => {
        const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
        editReportSheet.getRange("X1").values = [[draco_debugCell(typeof sql === "string" ? sql : sql.text)]];

        const EXCEL_CELL_CHAR_LIMIT = 32000; // límite real de Excel: 32767
        const jsonForCell = json.length > EXCEL_CELL_CHAR_LIMIT
            ? json.substring(0, EXCEL_CELL_CHAR_LIMIT) + " ...(truncado, JSON completo en la consola F12)"
            : json;
        editReportSheet.getRange("Y1").values = [[draco_debugCell(jsonForCell)]];

        await context.sync();
    });
    tPerf = draco_perfMark(reportId, "Escribir SQL/JSON en EDIT_REPORT", tPerf);

    // 3) JSON_PaintValues
    if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(reportId, "Pintando resultados"));
    await Excel.run(async (context) => {
        await jsonPaintValues(context, json, reportId);
    });
    tPerf = draco_perfMark(reportId, "Pintando resultados (total, ver desglose arriba)", tPerf);

    // Líneas EPM_FORMULA (fórmula, formato, barras) y, con "Sobrescribir
    // formatos en informes estáticos", cabeceras y valores con formato.
    try {
        await Excel.run(async (context) => {
            const editReportGrid = await getEditReportGrid(context, reportId);
            loadReportDefinition(editReportGrid, reportId);
            await finishDracoFixedReport(context, reportId, editReportGrid);
        });
    } catch (e) {
        console.warn("[Draco] No se pudieron aplicar las fórmulas/formato del informe estático:", e);
    }
    tPerf = draco_perfMark(reportId, "Líneas EPM_FORMULA y formato estático", tPerf);

    } finally {
        await endSuppressPlanningPaint();
        tPerf = draco_perfMark(reportId, "Candado cruzado de planificación (end)", tPerf);
        await draco_endManualCalculation();
        tPerf = draco_perfMark(reportId, "Restaurar cálculo automático", tPerf);
        if (window.BusyIndicator) await window.BusyIndicator.hide();
        draco_perfMark(reportId, "Ocultar indicador (borrar forma)", tPerf);
        draco_perfMark(reportId, "══ TOTAL refresco completo ══", tRefreshStart);
        await endDracoPerfLog();
    }
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
            editReportSheet.getRange("X1").values = [[draco_debugCell("ERROR: " + (error && error.message ? error.message : String(error)))]];
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
 * ===============  MODO HÍBRIDO: un eje Estático + uno Dinámico  ===========
 * ==========================================================================
 * Tres modos según EDIT_REPORT!H12 (Filas Estático) / N12 (Columnas
 * Estático), ver dracoReportMode:
 *   - "fixed"   : los dos Estáticos  -> buildSQLFixed + jsonPaintValues
 *   - "dynamic" : ninguno Estático   -> buildSQL (GROUPING SETS) + jsonTo3Matrices
 *   - "hybrid"  : solo uno Estático  -> buildSQLHybrid + jsonPaintHybrid (aquí)
 *
 * Eje ESTÁTICO: igual que en el modo Fijo. Cada línea (fila o columna de la
 * hoja) la define el usuario con fórmulas EPM_VALUE (condiciones
 * dimensión/atributo/valor, y opcionalmente su propia MEDIDA); se lee de la
 * hoja (readRowDefinitions / readColumnDefinitions) y NO se repinta nunca.
 * Puede ser asimétrico, tener huecos, separadores o fórmulas propias del
 * usuario entre líneas: nada de eso se toca.
 *
 * Eje DINÁMICO: sale de los datos en cada refresco (miembros distintos de
 * sus atributos, ordenados), y se pinta como fórmulas EPM_VALUE —igual que
 * un eje Estático— para que planificación, "Añadir filtro" y el paso a
 * Estático lo lean igual que siempre. Las medidas en este eje se reparten
 * por miembro en la posición en que están en el diseño.
 *
 * SQL (BigQuery):
 *   BASE   : atributos del eje estático + del dinámico + una SUM por
 *            medida, filtrado por (condiciones estáticas en OR) AND filtros.
 *   TAGGED : cada fila de BASE etiquetada con las líneas estáticas que cumple.
 *   AGG    : SUM por línea estática x miembro dinámico.
 *   final  : LINE_ID, DYN_ID (DENSE_RANK del miembro dinámico), P_ID (rango
 *            del prefijo anterior a la medida, para ordenar), atributos
 *            dinámicos y una columna por medida.
 * Como Fijo, agrega con SUM (los subtotales del eje dinámico no aplican:
 * se muestran sus miembros de detalle).
 * ========================================================================== */

function dracoReportMode(editReportGrid) {
    const rowsStatic = String(cellValue(editReportGrid, 12, 8)).trim().toUpperCase() === "X";
    const colsStatic = String(cellValue(editReportGrid, 12, 14)).trim().toUpperCase() === "X";
    if (rowsStatic && colsStatic) return "fixed";
    if (rowsStatic || colsStatic) return "hybrid";
    return "dynamic";
}

// Todos los valores "v" del JSON en orden (mismo parser que jsonTo3MatricesCore).
function extractDracoJsonValues(json) {
    const valores = [];
    let pos = 0;
    while (true) {
        const idx = json.indexOf('"v":', pos);
        if (idx === -1) break;
        let ini = idx + 4;
        while (json.charAt(ini) === " ") ini++;
        let fin, texto;
        if (json.charAt(ini) === '"') {
            ini++;
            fin = json.indexOf('"', ini);
            texto = json.substring(ini, fin);
        } else {
            fin = json.indexOf("}", ini);
            texto = json.substring(ini, fin).trim();
        }
        valores.push(texto);
        pos = fin + 1;
    }
    return valores;
}

// Propiedades de UN informe concreto (no del activo) con sus valores por defecto.
function dracoReportPropsFor(reportId) {
    const defaults = getDracoReportProperties();
    try {
        const report = window.ReportStore ? window.ReportStore.getReport(reportId) : null;
        return Object.assign({}, defaults, (report && report.reportProperties) || {});
    } catch (e) {
        return defaults;
    }
}

// Formato de número de una medida por su nombre (Opciones de campo de la
// medida, en cualquier zona), o "#,##0.00" si no tiene.
function dracoMeasureNumberFormatByName(reportId, name) {
    if (!name) return "#,##0.00";
    const m = (ReportState.Measures || []).find(x => String(x.Name).toUpperCase() === String(name).toUpperCase());
    if (m && m.NumberFormat) return m.NumberFormat;
    for (const zone of ["rows", "columns", "filters"]) {
        const f = dracoNumberFormatFromOptions(dracoMeasureFieldOptions(reportId, zone, name));
        if (f) return f;
    }
    return "#,##0.00";
}

/**
 * Campos de un eje en el orden del diseño (EDIT_REPORT desde la fila 15;
 * Filas = H..K, Columnas = N..Q: Dim, Atributo, Nivel, Jerarquía):
 *   {type:"attr", dim, attr}
 *   {type:"hier", dim, name, levels:[{attr, nivel}]}  (una fila por nivel visible)
 *   {type:"measure"}  (todas las medidas del eje = UN nivel, en la posición de la primera)
 */
function readDracoAxisFields(editReportGrid, isColumnsAxis) {
    const dimCol = isColumnsAxis ? 14 : 8;
    const fields = [];
    let R = 15;
    while (String(cellValue(editReportGrid, R, dimCol)).trim() !== "") {
        const dim = String(cellValue(editReportGrid, R, dimCol)).trim();
        const attr = String(cellValue(editReportGrid, R, dimCol + 1)).trim();
        const nivel = Number(cellValue(editReportGrid, R, dimCol + 2)) || 1;
        const hier = String(cellValue(editReportGrid, R, dimCol + 3)).trim();
        const hidden = dracoIsHiddenMarker(cellValue(editReportGrid, R, dimCol + 5));
        const last = fields[fields.length - 1];
        if (dim.toUpperCase() === "MEASURE") {
            if (!fields.some(f => f.type === "measure")) fields.push({ type: "measure" });
        } else if (hier) {
            if (last && last.type === "hier" && last.name.toUpperCase() === hier.toUpperCase() && last.dim.toUpperCase() === dim.toUpperCase()) {
                last.levels.push({ attr, nivel });
            } else {
                fields.push({ type: "hier", dim, name: hier, levels: [{ attr, nivel }], hidden });
            }
        } else if (!fields.some(f => f.type === "attr" && f.dim.toUpperCase() === dim.toUpperCase() && f.attr.toUpperCase() === attr.toUpperCase())) {
            // "Solo texto": el texto sustituye al ID oculto que lo precede
            const idPartner = (last && last.type === "attr" && last.hidden && !hidden
                && last.dim.toUpperCase() === dim.toUpperCase()) ? last : null;
            fields.push({ type: "attr", dim, attr, hidden, idPartner });
        }
        R++;
    }
    fields.forEach(f => { if (f.type === "hier") f.levels.sort((a, b) => a.nivel - b.nivel); });
    return fields;
}

async function buildSQLHybrid(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid, reportId) {
    const rowsStatic = String(cellValue(editReportGrid, 12, 8)).trim().toUpperCase() === "X";
    const staticIsRows = rowsStatic;               // si no, el estático es Columnas
    const dynamicIsColumns = staticIsRows;
    const props = dracoReportPropsFor(reportId);
    const nodesOnTop = props.hierarchyNodesOnTop !== false;

    // --- Eje estático: definiciones de la hoja (fórmulas EPM_VALUE)
    const defsRaw = staticIsRows
        ? await readRowDefinitions(context, editReportGrid, csvGrid)
        : await readColumnDefinitions(context, editReportGrid, csvGrid);
    const isMeasureDef = (d) => String(d.Dimension).toUpperCase() === "MEASURE";
    const linesMap = new Map(); // R -> {id, defs, measure}
    for (const d of defsRaw) {
        if (!linesMap.has(d.R)) linesMap.set(d.R, { id: d.R, defs: [], measure: null });
        const line = linesMap.get(d.R);
        if (isMeasureDef(d)) line.measure = d.AttributeName;
        else line.defs.push(d);
    }
    const lines = Array.from(linesMap.values()).sort((a, b) => a.id - b.id);

    // Líneas EPM_FORMULA del eje estático (no van al SQL; se rellenan al pintar)
    const formulaLines = [];
    try {
        const st = parseAddress(String(cellValue(editReportGrid, 10, staticIsRows ? 8 : 14)).trim());
        const last = staticIsRows ? lastRowInColumnFormulas(csvGrid, st.col) : lastColInRowFormulas(csvGrid, st.row);
        for (let i = staticIsRows ? st.row : st.col; i <= last; i++) {
            const r = staticIsRows ? i : st.row, c = staticIsRows ? st.col : i;
            if (cellHasFormula(csvGrid, r, c) && /EPM_FORMULA\s*\(/i.test(String(cellFormula(csvGrid, r, c)))) formulaLines.push(i);
        }
    } catch (e) { /* sin ancla: no hay líneas */ }

    // --- Eje dinámico: campos del diseño (atributos, jerarquías, medidas)
    const fields = readDracoAxisFields(editReportGrid, dynamicIsColumns);
    const dynCols = [];          // [{dim, attr}] columnas SQL (sin repetir)
    const colIdx = (dim, attr) => {
        let i = dynCols.findIndex(c => c.attr.toUpperCase() === attr.toUpperCase());
        if (i === -1) { dynCols.push({ dim, attr }); i = dynCols.length - 1; }
        return i;
    };
    fields.forEach(f => {
        if (f.type === "attr") f.col = colIdx(f.dim, f.attr);
        if (f.type === "hier") f.levels.forEach(l => { l.col = colIdx(f.dim, l.attr); });
    });
    // "ID y texto" en una jerarquía: la jerarquía de textos (sufijo ~TXT) se
    // cuelga de la de IDs (hostField) y comparte su profundidad/agrupación.
    fields.forEach(f => {
        if (f.type !== "hier" || !dracoIsHierTextName(f.name)) return;
        const hostName = f.name.slice(0, f.name.length - DRACO_HIER_TEXT_SUFFIX.length).toUpperCase();
        const host = fields.find(h => h.type === "hier" && h !== f && h.dim.toUpperCase() === f.dim.toUpperCase() && h.name.toUpperCase() === hostName);
        if (host) f.hostField = host;
    });
    const measureLevelIdx = fields.findIndex(f => f.type === "measure");
    const dynZone = dynamicIsColumns ? "columns" : "rows";
    const dynMeasures = ReportState.Measures.filter(m => m.Zone === dynZone).map(m => m.Name);

    const defaultMeasure = (ReportState.FilterMeasureNames && ReportState.FilterMeasureNames[0])
        || (ReportState.Measures[0] && ReportState.Measures[0].Name) || null;
    const neededMeasures = [];
    const addM = (n) => { if (n && neededMeasures.indexOf(n) === -1) neededMeasures.push(n); };
    if (dynMeasures.length) dynMeasures.forEach(addM);
    else { lines.forEach(l => addM(l.measure)); addM(defaultMeasure); }

    const measureCols = neededMeasures.map((name, k) => {
        const R = buscarMedida(measuresGrid, name);
        if (R === 0) throw new Error(`La medida «${name}» no está en el modelo semántico.`);
        return { name, alias: "M_" + (k + 1), field: String(cellValue(measuresGrid, R, 6)).trim(), scale: dracoMeasureScale(name) };
    });

    const plan = {
        staticIsRows, dynamicIsColumns,
        lines: lines.map(l => ({ id: l.id, measure: l.measure || defaultMeasure })),
        formulaLines,
        fields, dynCols, measureLevelIdx, dynMeasures,
        measureCols: measureCols.map(m => ({ name: m.name, alias: m.alias })),
        hasP: measureLevelIdx > 0 && dynMeasures.length > 0
    };

    if (lines.length === 0) return { sql: null, plan };

    const TRUE_EXPR = "TRUE";

    // --- BASE: atributos estáticos (de las condiciones) + dinámicos + medidas
    const selectCols = new Map();
    const addCol = (dim, attr) => {
        const key = (dim + "|" + attr).toUpperCase();
        if (!selectCols.has(key)) selectCols.set(key, getTableAlias(relGrid, dim) + "." + attr);
    };
    lines.forEach(l => l.defs.forEach(d => addCol(d.Dimension, d.AttributeName)));
    dynCols.forEach(a => addCol(a.dim, a.attr));
    const dynNames = dynCols.map(a => a.attr);

    const lineCond = (l) => l.defs.length
        ? l.defs.map(d => d.AttributeName + "=" + sqlValue(atributesGrid, d.Dimension, d.AttributeName, d.Value)).join(" AND ")
        : TRUE_EXPR;
    const prune = [];
    if (!lines.some(l => l.defs.length === 0)) {
        const seen = new Set();
        lines.forEach(l => l.defs.forEach(d => {
            const c = getTableAlias(relGrid, d.Dimension) + "." + d.AttributeName + "=" + sqlValue(atributesGrid, d.Dimension, d.AttributeName, d.Value);
            if (!seen.has(c.toUpperCase())) { seen.add(c.toUpperCase()); prune.push(c); }
        }));
    }
    const where = [];
    if (prune.length) where.push("(" + CRLF + "      " + prune.join(CRLF + "   OR ") + CRLF + ")");
    buildFilterConditions(atributesGrid, relGrid).forEach(c => where.push(c));

    // Agregación de excepción: BASE lleva también las dimensiones de referencia de esas medidas (DRC_R_j); el
    // AGG de siempre no las usa y cada medida con excepción se calcula aparte (AGG_E_n) y se cruza en AGG_X.
    const excMs = measureCols.filter(m => (m.exc = dracoMeasureExc(m.name)));
    const rCols = [], rJoins = [];
    excMs.forEach(m => {
        const ref = dracoExcRef(relGrid, m.exc);
        m.rNames = ref.cols.map(c => { let j = rCols.indexOf(c); if (j === -1) { rCols.push(c); j = rCols.length - 1; } return "DRC_R_" + j; });
        ref.joins.forEach(j => { if (rJoins.indexOf(j) === -1) rJoins.push(j); });
    });
    const rAll = rCols.map((c, j) => "DRC_R_" + j);

    let sql = "WITH BASE AS (" + CRLF + CRLF;
    sql += "SELECT" + CRLF;
    sql += [...selectCols.values(), ...rCols.map((c, j) => c + " AS DRC_R_" + j), ...measureCols.map(m => dracoScaled("SUM(f." + m.field + ")", m.scale) + " AS " + m.alias)]
        .map(x => "    " + x).join("," + CRLF) + CRLF + CRLF;
    sql += buildFrom(measuresGrid, measureCols.length ? measureCols[0].name : null, relGrid) + CRLF + CRLF;
    sql += buildJoins(relGrid) + rJoins.join("") + CRLF + CRLF;
    if (where.length) sql += "WHERE" + CRLF + where.join(CRLF + "AND ") + CRLF + CRLF;
    if (selectCols.size || rCols.length) sql += "GROUP BY" + CRLF + [...selectCols.values(), ...rCols].map(x => "    " + x).join("," + CRLF) + CRLF;
    sql += ")," + CRLF + CRLF;

    const mSel = measureCols.map(m => m.alias);

    // GROUP BY de AGG: con jerarquías en el eje dinámico, GROUPING SETS con
    // una combinación por profundidad de cada jerarquía (nodos + hojas,
    // como el modo Dinámico); sin jerarquías, GROUP BY normal.
    const hierFields = fields.filter(f => f.type === "hier" && !f.hostField);
    const plainCols = fields.filter(f => f.type === "attr").map(f => dynNames[f.col]);
    const groupBy = (lineRef, colRef, extra) => {
        extra = extra || [];
        if (hierFields.length === 0) {
            return "GROUP BY " + [lineRef, ...dynNames.map(colRef), ...extra].join(", ");
        }
        let combos = [[]];
        hierFields.forEach(h => {
            const next = [];
            const texts = fields.filter(t => t.hostField === h);
            combos.forEach(c => {
                for (let d = 1; d <= h.levels.length; d++) {
                    next.push([...c, h.levels.slice(0, d).map(l => dynNames[l.col]),
                        ...texts.map(t => t.levels.slice(0, d).map(l => dynNames[l.col]))]);
                }
            });
            combos = next;
        });
        const sets = combos.map(c => {
            const cols = [];
            [...plainCols, ...c.flat()].forEach(n => { if (cols.indexOf(n) === -1) cols.push(n); });
            return "(" + [lineRef, ...cols.map(colRef), ...extra].join(", ") + ")";
        });
        return "GROUP BY GROUPING SETS (" + CRLF + sets.map(x => "        " + x).join("," + CRLF) + CRLF + "    )";
    };

    // con excepción y jerarquías, AGG lleva el GROUPING() de cada columna dinámica para cruzar con AGG_E_n
    const gFlags = (excMs.length && hierFields.length) ? dynNames : [];
    const gSel = (colRef) => gFlags.map((n, i) => ", GROUPING(" + colRef(n) + ") AS DRC_G_" + i).join("");
    let excSrc = null;      // { lineRef, colRef, mRef, from } para AGG_E_n
    {
        const structs = lines.map(l => "STRUCT(" + l.id + " AS LINE_ID, (" + lineCond(l) + ") AS COND)");
        sql += "TAGGED AS (" + CRLF
            + "    SELECT " + [...dynNames, ...mSel, ...rAll].map(x => x + ", ").join("")
            + "ARRAY(" + CRLF + "        SELECT LINE_ID FROM UNNEST([" + CRLF
            + structs.map(x => "            " + x).join("," + CRLF) + CRLF
            + "        ]) WHERE COND" + CRLF + "    ) AS LINE_IDS" + CRLF
            + "    FROM BASE" + CRLF + ")," + CRLF + CRLF;
        excSrc = { lineRef: "LINE_ID", colRef: n => n, mRef: a => a, from: "    FROM TAGGED, UNNEST(LINE_IDS) AS LINE_ID" + CRLF };
        sql += "AGG AS (" + CRLF
            + "    SELECT LINE_ID" + dynNames.map(n => ", " + n).join("") + gSel(n => n) + measureCols.map(m => ", SUM(" + m.alias + ") AS " + m.alias).join("") + CRLF
            + "    FROM TAGGED, UNNEST(LINE_IDS) AS LINE_ID" + CRLF
            + "    " + groupBy("LINE_ID", n => n) + CRLF
            + ")" + CRLF + CRLF;
    }

    // --- Agregación de excepción: AGG_E_n (la medida sumada por línea, columnas dinámicas y referencia; encima, la
    //     excepción por línea y columnas dinámicas) y AGG_X = AGG con esas medidas tomadas de su AGG_E_n.
    let aggName = "AGG";
    if (excMs.length) {
        sql = sql.replace(/\)\r?\n\r?\n$/, ")," + CRLF + CRLF);
        const S = excSrc, keyN = [...dynNames, ...gFlags.map((n, i) => "DRC_G_" + i)];
        excMs.forEach((m, k) => {
            sql += "AGG_E_" + (k + 1) + " AS (" + CRLF
                + "    SELECT LINE_ID" + keyN.map(n => ", " + n).join("") + ", " + m.exc.agg + "(DRC_V) AS DRC_V" + CRLF
                + "    FROM (" + CRLF
                + "    SELECT " + S.lineRef + " AS LINE_ID" + dynNames.map(n => ", " + S.colRef(n) + " AS " + n).join("") + gSel(S.colRef)
                + ", SUM(" + S.mRef(m.alias) + ") AS DRC_V" + CRLF
                + S.from
                + "    " + groupBy(S.lineRef, S.colRef, m.rNames.map(S.colRef)) + CRLF
                + "    ) Q" + CRLF
                + "    GROUP BY LINE_ID" + keyN.map(n => ", " + n).join("") + CRLF
                + ")," + CRLF + CRLF;
        });
        sql += "AGG_X AS (" + CRLF
            + "    SELECT a.LINE_ID" + dynNames.map(n => ", a." + n).join("")
            + measureCols.map(m => { const k = excMs.indexOf(m); return k === -1 ? ", a." + m.alias : ", e" + (k + 1) + ".DRC_V AS " + m.alias; }).join("") + CRLF
            + "    FROM AGG a" + CRLF
            + excMs.map((m, k) => "    LEFT JOIN AGG_E_" + (k + 1) + " e" + (k + 1) + " ON a.LINE_ID = e" + (k + 1) + ".LINE_ID"
                + keyN.map(n => " AND a." + n + " IS NOT DISTINCT FROM e" + (k + 1) + "." + n).join("") + CRLF).join("")
            + ")" + CRLF + CRLF;
        aggName = "AGG_X";
    }

    // --- Orden del eje dinámico: atributos por su valor; en jerarquías,
    //     el nodo (niveles inferiores NULL) delante o detrás de sus hijos
    //     según "Mostrar nodos de jerarquía arriba".
    const nullFirst = (n) => "CASE WHEN " + n + " IS NULL THEN " + (nodesOnTop ? "0 ELSE 1" : "1 ELSE 0") + " END";
    const orderOf = (flds) => {
        const parts = [];
        flds.forEach(f => {
            if (f.type === "attr") parts.push(dynNames[f.col]);
            // La jerarquía de textos ("ID y texto") no ordena: va por sus IDs.
            if (f.type === "hier" && !f.hostField) f.levels.forEach((l, i) => {
                if (i > 0) parts.push(nullFirst(dynNames[l.col]));
                parts.push(dynNames[l.col]);
            });
        });
        return parts;
    };
    const rank = (parts) => parts.length ? "DENSE_RANK() OVER (ORDER BY " + parts.join(", ") + ")" : "1";
    sql += "SELECT" + CRLF
        + "    LINE_ID," + CRLF
        + "    " + rank(orderOf(fields)) + " AS DYN_ID," + CRLF
        + "    " + (plan.hasP ? rank(orderOf(fields.slice(0, measureLevelIdx))) : "1") + " AS P_ID"
        + [...dynNames, ...mSel].map(x => "," + CRLF + "    " + x).join("") + CRLF
        + "FROM " + aggName + CRLF
        + "ORDER BY LINE_ID, DYN_ID";

    return { sql, plan };
}

/**
 * Estilo de cabecera de Draco (el mismo que el modo Dinámico: Segoe UI 9,
 * relleno/texto por nivel con DRACO_LEVEL_PALETTE, sangría por
 * profundidad de jerarquía y borde exterior fino) para un conjunto de
 * celdas de cabecera {row, col, indent}.
 *   axis "col": cabecera de FILAS (el campo lo da la columna).
 *   axis "row": cabecera de COLUMNAS (el campo lo da la fila).
 */
async function applyDracoHeaderStyle(context, sheet, cells, axis, fieldOffset) {
    if (!cells || cells.size === 0) return;
    await writeIndentAndColorRuns(context, sheet, cells, axis, fieldOffset, true);
    let r1 = Infinity, r2 = -Infinity, c1 = Infinity, c2 = -Infinity;
    for (const c of cells.values()) {
        r1 = Math.min(r1, c.row); r2 = Math.max(r2, c.row);
        c1 = Math.min(c1, c.col); c2 = Math.max(c2, c.col);
    }
    const box = sheet.getRangeByIndexes(r1 - 1, c1 - 1, r2 - r1 + 1, c2 - c1 + 1);
    box.format.font.bold = false;
    [Excel.BorderIndex.edgeTop, Excel.BorderIndex.edgeBottom, Excel.BorderIndex.edgeLeft, Excel.BorderIndex.edgeRight]
        .forEach(edge => {
            const b = box.format.borders.getItem(edge);
            b.style = Excel.BorderLineStyle.continuous;
            b.weight = Excel.BorderWeight.thin;
            b.color = DRACO_BORDER_COLOR;
        });
}

/**
 * Pinta el resultado híbrido: cabecera del eje dinámico (fórmulas
 * EPM_VALUE, con el estilo del modo Dinámico y jerarquías con sangría) y
 * valores SOLO en las celdas línea-estática x miembro dinámico.
 */
async function jsonPaintHybrid(context, json, reportId, plan, editReportGrid) {
    let tPerf = performance.now();
    const props = dracoReportPropsFor(reportId);

    const resultSheetName = await getDracoResultSheetName(context, reportId);
    await ensureDracoResultSheetExists(context, resultSheetName);
    const sheet = context.workbook.worksheets.getItem(resultSheetName);
    const rn = dracoRangeNames(reportId);
    const dynName = plan.dynamicIsColumns ? rn.cols : rn.rows;

    // --- Anclas
    const tryParse = (raw) => { const t = String(raw || "").trim(); if (!t) return null; try { return parseAddress(t); } catch (e) { return null; } };
    const RRows = tryParse(cellValue(editReportGrid, 10, 8));
    const RCols = tryParse(cellValue(editReportGrid, 10, 14));
    let anchor = plan.dynamicIsColumns ? RCols : RRows;
    if (!anchor) {
        const st = plan.dynamicIsColumns ? RRows : RCols;
        if (!st) throw new Error("No se encuentra la posición del eje Estático (EDIT_REPORT!H10/N10).");
        anchor = plan.dynamicIsColumns
            ? { row: st.row, col: st.col + Math.max(1, ReportState.AttrColumnCount) }
            : { row: st.row + Math.max(1, ReportState.AttrRowCount), col: st.col };
    }

    // --- Parseo
    const nDyn = plan.dynCols.length;
    const nM = plan.measureCols.length;
    const per = 3 + nDyn + nM;
    const vals = json ? extractDracoJsonValues(json) : [];
    const recs = [];
    for (let i = 0; i + per <= vals.length; i += per) {
        recs.push({
            line: Number(vals[i]), dyn: Number(vals[i + 1]), p: Number(vals[i + 2]),
            attrs: vals.slice(i + 3, i + 3 + nDyn),
            values: vals.slice(i + 3 + nDyn, i + per)
        });
    }
    tPerf = draco_perfMark(reportId, `Híbrido: parseo (${recs.length} filas)`, tPerf);

    // --- Miembros dinámicos y líneas físicas del eje dinámico
    const members = new Map();
    recs.forEach(r => { if (!members.has(r.dyn)) members.set(r.dyn, { p: r.p, attrs: r.attrs }); });
    const tuples = [];
    if (plan.dynMeasures.length) {
        for (const [dyn, mem] of members) plan.dynMeasures.forEach((m, mIdx) => tuples.push({ dyn, mIdx, mem }));
        tuples.sort((a, b) => (a.mem.p - b.mem.p) || (a.mIdx - b.mIdx) || (a.dyn - b.dyn));
    } else {
        for (const [dyn, mem] of members) tuples.push({ dyn, mIdx: -1, mem });
        tuples.sort((a, b) => a.dyn - b.dyn);
    }
    if (members.size === 0 && nDyn === 0 && !plan.dynMeasures.length) tuples.push({ dyn: 1, mIdx: -1, mem: { attrs: [] } });
    const tupleIndex = new Map(tuples.map((t, j) => [t.dyn + "_" + t.mIdx, j]));
    // "Solo texto": los campos ocultos (ID) se consultan pero no se pintan en la cabecera.
    const paintFields = plan.fields.filter(f => !f.hidden);
    const headerDepth = paintFields.length;
    const valCell = (lineId, j) => plan.dynamicIsColumns ? { r: lineId, c: anchor.col + j } : { r: anchor.row + j, c: lineId };
    const firstDyn = plan.dynamicIsColumns ? anchor.col : anchor.row;
    const nTuples = tuples.length;

    // --- Borrar lo pintado la vez anterior
    const prevDyn = context.workbook.names.getItemOrNullObject(dynName);
    const prevVals = context.workbook.names.getItemOrNullObject(rn.values);
    prevDyn.load("name"); prevVals.load("name");
    await context.sync();
    let prevDynRange = null, prevValsRange = null;
    if (!prevDyn.isNullObject) { prevDynRange = prevDyn.getRange(); prevDynRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "worksheet/name"]); }
    if (!prevVals.isNullObject) { prevValsRange = prevVals.getRange(); prevValsRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "worksheet/name"]); }
    await context.sync();
    const clearMode = props.overwriteFormats ? Excel.ClearApplyTo.all : Excel.ClearApplyTo.contents;
    if (prevDynRange && prevDynRange.worksheet.name === resultSheetName) prevDynRange.clear(clearMode);
    if (prevValsRange && prevValsRange.worksheet.name === resultSheetName) {
        const pv = prevValsRange;
        const toClear = [...plan.lines, ...(plan.formulaLines || []).map(id => ({ id }))];
        for (const l of toClear) {
            if (plan.dynamicIsColumns) {
                if (l.id - 1 >= pv.rowIndex && l.id - 1 < pv.rowIndex + pv.rowCount) {
                    sheet.getRangeByIndexes(l.id - 1, pv.columnIndex, 1, pv.columnCount).clear(Excel.ClearApplyTo.contents);
                }
            } else if (l.id - 1 >= pv.columnIndex && l.id - 1 < pv.columnIndex + pv.columnCount) {
                sheet.getRangeByIndexes(pv.rowIndex, l.id - 1, pv.rowCount, 1).clear(Excel.ClearApplyTo.contents);
            }
        }
    }
    if (!prevDyn.isNullObject) prevDyn.delete();
    if (!prevVals.isNullObject) prevVals.delete();

    // --- Cabecera dinámica: una línea por campo; las jerarquías en UNA
    //     línea, con el nodo más profundo informado y sangría por nivel.
    const headerCells = new Map();
    if (headerDepth > 0 && nTuples > 0) {
        const rowsN = plan.dynamicIsColumns ? headerDepth : nTuples;
        const colsN = plan.dynamicIsColumns ? nTuples : headerDepth;
        const grid = [];
        for (let r = 0; r < rowsN; r++) grid.push(new Array(colsN).fill(""));
        const isNull = (v) => v === undefined || isDracoNullLiteral(v);
        tuples.forEach((t, j) => {
            paintFields.forEach((fld, k) => {
                let f = "", indent = 0;
                if (fld.type === "measure") {
                    const mName = t.mIdx >= 0 ? plan.dynMeasures[t.mIdx] : "";
                    if (mName) f = buildEpmValueFormula("MEASURE", mName, "", mName);
                } else if (fld.type === "attr") {
                    const v = t.mem.attrs[fld.col];
                    if (!isNull(v)) {
                        // Solo texto: atributo y valor del ID, descripción = texto
                        f = fld.idPartner
                            ? buildEpmValueFormula(fld.dim, fld.idPartner.attr, t.mem.attrs[fld.idPartner.col], v)
                            : buildEpmValueFormula(fld.dim, fld.attr, v);
                    }
                } else {
                    // Jerarquía de textos ("ID y texto"): misma profundidad que la de IDs
                    // (un texto NULL no la acorta), el valor es el texto de ese nivel.
                    const depthFld = fld.hostField || fld;
                    let depth = 0;
                    while (depth < depthFld.levels.length && !isNull(t.mem.attrs[depthFld.levels[depth].col])) depth++;
                    if (depth > 0) {
                        const lv = fld.levels[Math.min(depth, fld.levels.length) - 1];
                        const v = t.mem.attrs[lv.col];
                        if (!fld.hostField || !isNull(v)) {
                            const hl = (fld.hostField && fld.hostField.hidden) ? fld.hostField.levels[Math.min(depth, fld.hostField.levels.length) - 1] : null;
                            f = hl
                                ? buildEpmValueFormula(fld.dim, hl.attr, t.mem.attrs[hl.col], v)
                                : buildEpmValueFormula(fld.dim, lv.attr, v);
                        }
                        indent = depth - 1;
                    }
                }
                const r = plan.dynamicIsColumns ? k : j, c = plan.dynamicIsColumns ? j : k;
                grid[r][c] = f;
                const row = anchor.row + r, col = anchor.col + c;
                headerCells.set(row + "_" + col, { row, col, indent });
            });
        });
        const hdr = sheet.getRangeByIndexes(anchor.row - 1, anchor.col - 1, rowsN, colsN);
        hdr.formulas = grid;
        context.workbook.names.add(dynName, hdr);
        if (props.overwriteFormats) {
            // Mismo estilo que el modo Dinámico (antes: blanco y negrita)
            await applyDracoHeaderStyle(context, sheet, headerCells,
                plan.dynamicIsColumns ? "row" : "col",
                plan.dynamicIsColumns ? anchor.row - 1 : anchor.col - 1);
        }
    }

    // --- Valores: solo las celdas línea x miembro
    const measureColIdx = new Map(plan.measureCols.map((m, k) => [m.name, k]));
    const lineMeasure = new Map(plan.lines.map(l => [l.id, l.measure]));
    const cells = [];
    for (const r of recs) {
        if (plan.dynMeasures.length) {
            plan.dynMeasures.forEach((mName, mIdx) => {
                const j = tupleIndex.get(r.dyn + "_" + mIdx);
                const k = measureColIdx.get(mName);
                if (j === undefined || k === undefined) return;
                const pos = valCell(r.line, j);
                cells.push({ r: pos.r, c: pos.c, value: coerceCellLiteral(r.values[k]), measure: mName });
            });
        } else {
            const j = tupleIndex.get(r.dyn + "_-1");
            const mName = lineMeasure.get(r.line);
            const k = measureColIdx.get(mName);
            if (j === undefined || k === undefined) continue;
            const pos = valCell(r.line, j);
            cells.push({ r: pos.r, c: pos.c, value: coerceCellLiteral(r.values[k]), measure: mName });
        }
    }

    const lineIds = [...plan.lines.map(l => l.id), ...(plan.formulaLines || [])];
    if (lineIds.length && nTuples > 0) {
        const minL = Math.min(...lineIds), maxL = Math.max(...lineIds);
        const rect = plan.dynamicIsColumns
            ? sheet.getRangeByIndexes(minL - 1, firstDyn - 1, maxL - minL + 1, nTuples)
            : sheet.getRangeByIndexes(firstDyn - 1, minL - 1, nTuples, maxL - minL + 1);
        context.workbook.names.add(rn.values, rect);
    }

    // Escritura por línea estática (sin pisar nada entre líneas). El
    // formato de los valores (medida, centrado) se aplica con "Sobrescribir
    // formatos" o con "Sobrescribir formatos en informes estáticos".
    const formatValues = props.overwriteFormats || props.overwriteStaticFormats;
    const byLine = new Map();
    for (const c of cells) {
        const key = plan.dynamicIsColumns ? c.r : c.c;
        if (!byLine.has(key)) byLine.set(key, []);
        byLine.get(key).push(c);
    }
    const baseline = props.planningReport ? new Map() : null;
    for (const [key, list] of byLine) {
        const along = (c) => plan.dynamicIsColumns ? c.c : c.r;
        const lo = Math.min(...list.map(along)), hi = Math.max(...list.map(along));
        const n = hi - lo + 1;
        const values = new Array(n).fill("");
        const formats = new Array(n).fill("#,##0.00");
        list.forEach(c => {
            values[along(c) - lo] = c.value;
            formats[along(c) - lo] = dracoMeasureNumberFormatByName(reportId, c.measure);
            if (baseline) baseline.set(c.r + "_" + c.c, c.value);
        });
        const rg = plan.dynamicIsColumns
            ? sheet.getRangeByIndexes(key - 1, lo - 1, 1, n)
            : sheet.getRangeByIndexes(lo - 1, key - 1, n, 1);
        rg.values = plan.dynamicIsColumns ? [values] : values.map(v => [v]);
        if (formatValues) {
            rg.numberFormat = plan.dynamicIsColumns ? [formats] : formats.map(f => [f]);
            rg.format.font.name = DRACO_FONT_NAME;
            rg.format.font.size = DRACO_FONT_SIZE;
            rg.format.horizontalAlignment = Excel.HorizontalAlignment.center;
        }
    }
    if (baseline) DracoPlanningBaselineValues.set(reportId, baseline);
    else DracoPlanningBaselineValues.delete(reportId);

    // Eje estático: líneas EPM_FORMULA (fórmula en cada miembro dinámico,
    // con su formato y barras) y, opcionalmente, cabeceras con estilo Draco.
    const staticLines = await scanDracoStaticAxis(context, sheet, editReportGrid, plan.staticIsRows ? "rows" : "cols");
    DracoCalcCellsByReport.set(dracoStateKey(reportId), new Set());
    if (nTuples > 0) {
        const targets = [];
        for (let j = 0; j < nTuples; j++) targets.push(firstDyn + j);
        applyDracoFormulaLines(sheet, reportId, staticLines, plan.staticIsRows, targets);
    }
    if (props.overwriteStaticFormats) {
        await formatDracoStaticAxis(context, sheet, editReportGrid, plan.staticIsRows ? "rows" : "cols", staticLines);
    }

    // El autoajuste de columnas NO se hace aquí: durante el refresco el
    // cálculo está en manual y las fórmulas EPM_VALUE de la cabecera aún no
    // muestran su texto, así que Excel ajustaría a celdas "vacías" (columna
    // enana). Se hace al final, con el cálculo ya restaurado (ver
    // autofitDracoHybrid en actualizarInformeHybridCore).
    await context.sync();
    draco_perfMark(reportId, `Híbrido: pintado (${nTuples} miembros dinámicos, ${cells.length} valores)`, tPerf);
}

/* ---------------------------------------------------------------------
 * "Sobrescribir formatos en informes estáticos": formato de las celdas
 * que define el usuario (ejes Estáticos).
 * ------------------------------------------------------------------- */

// Argumentos de una fórmula =EPM_VALUE("DIM","ATTR","VALOR","DISPLAY") (sin resolver referencias)
function parseEpmValueArgs(formula) {
    const f = String(formula || "");
    const m = /EPM_VALUE\s*\(([\s\S]*)\)\s*$/i.exec(f);
    if (!m) return null;
    const args = [];
    let cur = "", inStr = false;
    const body = m[1];
    for (let i = 0; i < body.length; i++) {
        const ch = body[i];
        if (ch === '"') {
            if (inStr && body[i + 1] === '"') { cur += '"'; i++; continue; }
            inStr = !inStr;
            continue;
        }
        if (!inStr && (ch === "," || ch === ";")) { args.push(cur.trim()); cur = ""; continue; }
        cur += ch;
    }
    args.push(cur.trim());
    return args;
}

// Separa los argumentos de una llamada f(a, b, ...) respetando comillas
// ("" escapadas), paréntesis y llaves. `formula` es la fórmula completa
// (con o sin "=" / "@"); devuelve null si no es una llamada a `fnName`.
function splitDracoFunctionArgs(formula, fnName) {
    const f = String(formula || "");
    const re = new RegExp("^\\s*=\\s*@?\\s*" + fnName + "\\s*\\(", "i");
    const m = re.exec(f);
    if (!m) return null;
    const args = [];
    let cur = "", depth = 0, inStr = false;
    for (let i = m[0].length; i < f.length; i++) {
        const ch = f[i];
        if (inStr) {
            cur += ch;
            if (ch === '"') {
                if (f[i + 1] === '"') { cur += '"'; i++; } else inStr = false;
            }
            continue;
        }
        if (ch === '"') { inStr = true; cur += ch; continue; }
        if (ch === "(" || ch === "{") { depth++; cur += ch; continue; }
        if (ch === ")" || ch === "}") {
            if (depth === 0) { args.push(cur.trim()); return args; }
            depth--; cur += ch; continue;
        }
        if (depth === 0 && (ch === "," || ch === ";")) { args.push(cur.trim()); cur = ""; continue; }
        cur += ch;
    }
    args.push(cur.trim());
    return args;
}

const unquoteDraco = (a) => {
    const t = String(a || "").trim();
    return (t.length >= 2 && t[0] === '"' && t[t.length - 1] === '"') ? t.slice(1, -1).replace(/""/g, '"') : t;
};

// "barra|invertir|izq=40|dcha=25|neg=#F0436A|pos=#70D06A" -> opciones (o null)
function parseDracoBarSpec(spec) {
    const parts = String(spec || "").split("|").map(x => x.trim()).filter(Boolean);
    if (!parts.length) return null;
    const o = { display: null, invert: false, left: 42, right: 30, neg: "#F0436A", pos: "#70D06A" };
    for (const p of parts) {
        const low = p.toLowerCase().replace(/\s+/g, " ");
        const kv = /^([a-z]+)\s*=\s*(.+)$/.exec(low);
        if (low === "barra" || low === "barras") o.display = "bar";
        else if (low === "solo barra" || low === "solobarra" || low === "solo barras") o.display = "barOnly";
        else if (low === "invertir") o.invert = true;
        else if (kv && kv[1] === "izq") o.left = Math.max(0, Math.min(90, Number(kv[2]) || 0));
        else if (kv && kv[1] === "dcha") o.right = Math.max(0, Math.min(90, Number(kv[2]) || 0));
        else if (kv && kv[1] === "neg") o.neg = p.split("=")[1].trim();
        else if (kv && kv[1] === "pos") o.pos = p.split("=")[1].trim();
    }
    return o.display ? o : null;
}

/**
 * Iconos del 4.º parámetro de EPM_FORMULA (se combinan con las barras), separados por "|":
 *   semaforo (o pelotitas)  -> 3 semáforos (círculos rojo / ámbar / verde)
 *   señales                 -> 3 señales (rombo rojo / triángulo ámbar / círculo verde)
 *   cortes=a,b              -> cortes FIJOS del icono intermedio (>= a) y del superior (>= b);
 *                              sin cortes: 33 % y 67 % entre el mínimo y el máximo, como Excel
 *   invertir-iconos         -> el verde para los valores bajos (p. ej. costes)
 *   solo-icono              -> oculta el número
 * Devuelve { style: "traffic"|"signs", cuts: [a,b]|null, invert, only } o null.
 */
function parseDracoIconSpec(spec) {
    const parts = String(spec || "").split("|").map(x => x.trim()).filter(Boolean);
    let style = null, cuts = null, invert = false, only = false;
    for (const p of parts) {
        const low = p.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ");
        if (low === "semaforo" || low === "semaforos" || low === "pelotitas") style = "traffic";
        else if (low === "senales" || low === "senal") style = "signs";
        else if (low === "invertir-iconos" || low === "invertir iconos") invert = true;
        else if (low === "solo-icono" || low === "solo icono") only = true;
        else if (low.indexOf("cortes=") === 0) {
            const v = p.split("=")[1].split(/[,;]/).map(x => Number(String(x).trim().replace(",", ".")));
            if (v.length === 2 && v.every(isFinite)) cuts = v[0] <= v[1] ? v : [v[1], v[0]];
        }
    }
    return style ? { style, cuts, invert, only } : null;
}

/**
 * Recorre las cabeceras de un eje Estático (desde EDIT_REPORT!H10 si es
 * Filas, N10 si es Columnas) y devuelve sus líneas:
 *   { line, isFormula, measure, cells:[{row,col,indent}],
 *     formula: { r1c1, format, bar } }   (solo en líneas EPM_FORMULA)
 * Una línea es EPM_VALUE (datos) o EPM_FORMULA (calculada). Igual que
 * readRowDefinitions, la línea se corta en la primera celda que no sea EPM.
 */
async function scanDracoStaticAxis(context, sheet, editReportGrid, axis) {
    const isRows = axis === "rows";
    const anchorRaw = String(cellValue(editReportGrid, 10, isRows ? 8 : 14)).trim();
    if (!anchorRaw) return [];
    let anchor;
    try { anchor = parseAddress(anchorRaw); } catch (e) { return []; }

    const fields = readDracoAxisFields(editReportGrid, !isRows);
    const levelOf = new Map();
    fields.forEach(f => { if (f.type === "hier") f.levels.forEach((l, i) => levelOf.set(f.dim.toUpperCase() + "|" + l.attr.toUpperCase(), i)); });

    const depth = Math.max(1, isRows ? ReportState.AttrColumnCount : ReportState.AttrRowCount);
    const used = sheet.getUsedRange(true);
    used.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();
    const span = isRows
        ? Math.max(0, used.rowIndex + used.rowCount - anchor.row + 1)
        : Math.max(0, used.columnIndex + used.columnCount - anchor.col + 1);
    if (span === 0) return [];

    const area = isRows
        ? sheet.getRangeByIndexes(anchor.row - 1, anchor.col - 1, span, depth)
        : sheet.getRangeByIndexes(anchor.row - 1, anchor.col - 1, depth, span);
    area.load(["formulas", "formulasR1C1"]);
    await context.sync();

    const lines = [];
    for (let i = 0; i < span; i++) {
        let line = null;
        for (let k = 0; k < depth; k++) {
            const a1 = isRows ? area.formulas[i][k] : area.formulas[k][i];
            if (typeof a1 !== "string") break;
            const row = isRows ? anchor.row + i : anchor.row + k;
            const col = isRows ? anchor.col + k : anchor.col + i;
            if (/^\s*=\s*@?\s*EPM_FORMULA\s*\(/i.test(a1)) {
                const r1c1 = isRows ? area.formulasR1C1[i][k] : area.formulasR1C1[k][i];
                const args = splitDracoFunctionArgs(r1c1, "EPM_FORMULA") || [];
                if (!line) line = { line: isRows ? row : col, isFormula: true, measure: null, cells: [] };
                line.isFormula = true;
                line.formula = {
                    r1c1: (args[1] || "").trim(),
                    format: unquoteDraco(args[2]),
                    bar: parseDracoBarSpec(unquoteDraco(args[3])),
                    icons: parseDracoIconSpec(unquoteDraco(args[3]))
                };
                line.cells.push({ row, col, indent: 0 });
                break;
            }
            if (!/^\s*=\s*@?\s*EPM_VALUE\s*\(/i.test(a1)) break;
            const args = parseEpmValueArgs(a1) || [];
            const dim = String(args[0] || "").toUpperCase();
            const attr = String(args[1] || "").toUpperCase();
            if (!line) line = { line: isRows ? row : col, isFormula: false, measure: null, cells: [] };
            if (dim === "MEASURE") line.measure = args[1] || null;
            line.cells.push({ row, col, indent: levelOf.get(dim + "|" + attr) || 0 });
        }
        if (line) lines.push(line);
    }
    return lines;
}

/**
 * Da a las cabeceras (EPM_VALUE y EPM_FORMULA) de un eje Estático el
 * estilo del modo Dinámico. Devuelve las líneas (ver scanDracoStaticAxis).
 */
async function formatDracoStaticAxis(context, sheet, editReportGrid, axis, scanned) {
    const lines = scanned || await scanDracoStaticAxis(context, sheet, editReportGrid, axis);
    const cells = new Map();
    lines.forEach(l => l.cells.forEach(c => cells.set(c.row + "_" + c.col, c)));
    if (cells.size) {
        const anchor = parseAddress(String(cellValue(editReportGrid, 10, axis === "rows" ? 8 : 14)).trim());
        await applyDracoHeaderStyle(context, sheet, cells, axis === "rows" ? "col" : "row", axis === "rows" ? anchor.col - 1 : anchor.row - 1);
    }
    return lines;
}

/**
 * Líneas EPM_FORMULA de un eje Estático: copia su fórmula a cada celda de
 * valores de la línea (formulasR1C1 = mismas referencias relativas que
 * copiar y pegar la celda de cabecera), aplica el formato de número si se
 * indicó y las barras de datos si se pidieron. Las celdas quedan fuera de
 * la planificación.
 *   lines   : líneas del eje (scanDracoStaticAxis), solo se usan las isFormula
 *   isRows  : true si el eje Estático es Filas (las celdas van a lo ancho)
 *   targets : posiciones (columnas si isRows, filas si no) donde escribir
 */
function applyDracoFormulaLines(sheet, reportId, lines, isRows, targets) {
    const set = DracoCalcCellsByReport.get(dracoStateKey(reportId)) || new Set();
    const sorted = Array.from(new Set(targets)).sort((a, b) => a - b);
    // tramos contiguos de destino
    const runs = [];
    for (const t of sorted) {
        const last = runs[runs.length - 1];
        if (last && t === last.to + 1) last.to = t; else runs.push({ from: t, to: t });
    }

    for (const l of lines) {
        if (!l.isFormula || !l.formula || !l.formula.r1c1) continue;
        const f = "=" + l.formula.r1c1;
        const fmt = l.formula.format;
        const bar = l.formula.bar;
        const ranges = runs.map(run => {
            const n = run.to - run.from + 1;
            const rg = isRows
                ? sheet.getRangeByIndexes(l.line - 1, run.from - 1, 1, n)
                : sheet.getRangeByIndexes(run.from - 1, l.line - 1, n, 1);
            rg.formulasR1C1 = isRows ? [new Array(n).fill(f)] : new Array(n).fill(0).map(() => [f]);
            for (let t = run.from; t <= run.to; t++) set.add(isRows ? (l.line + "_" + t) : (t + "_" + l.line));
            rg.conditionalFormats.clearAll();
            return { rg, run };
        });

        // Formato de número (si se indicó). Con "barra", relleno con espacios
        // para dejar el número junto al extremo hacia el que crece su barra.
        let numFmt = fmt || "";
        if (bar && bar.display === "bar") {
            const base = numFmt || "#,##0.00";
            numFmt = base.indexOf(";") !== -1 ? base : "* " + base + "\\ ;\\ -" + base + "* ;* " + base + "\\ ";
        }
        if (numFmt) ranges.forEach(({ rg }) => { rg.numberFormat = [[numFmt]]; });

        // Iconos (semáforo / señales): formato condicional de Excel, con o sin barras
        const icons = l.formula.icons;
        if (icons && ranges.length) {
            ranges.forEach(({ rg }) => {
                const cf = rg.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
                const is = cf.iconSet;
                is.style = icons.style === "signs" ? Excel.IconSet.threeSigns : Excel.IconSet.threeTrafficLights1;
                is.reverseIconOrder = !!icons.invert;
                is.showIconOnly = !!icons.only;
                const type = icons.cuts ? Excel.ConditionalFormatIconRuleType.number : Excel.ConditionalFormatIconRuleType.percent;
                const v = icons.cuts || [33, 67];
                is.criteria = [
                    {},
                    { type, operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual, formula: "=" + v[0] },
                    { type, operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual, formula: "=" + v[1] }
                ];
            });
        }

        if (!bar || ranges.length === 0) continue;
        const addr = ({ run }) => isRows
            ? "$" + dracoColumnLetter(run.from) + "$" + l.line + ":$" + dracoColumnLetter(run.to) + "$" + l.line
            : "$" + dracoColumnLetter(l.line) + "$" + run.from + ":$" + dracoColumnLetter(l.line) + "$" + run.to;
        const list = ranges.slice(0, 200).map(addr).join(",");
        const barOnly = bar.display === "barOnly";
        const L = barOnly ? 0.04 : bar.left / 100;
        const R = barOnly ? 0.04 : bar.right / 100;
        const mn = "MIN(0,MIN(" + list + "))";
        const mx = "MAX(0,MAX(" + list + "))";
        const spanE = "MAX(0.000001," + mx + "-" + mn + ")";
        const ln = "IF(" + mn + "<0," + L + ",0)";
        const lp = "IF(" + mx + ">0," + R + ",0)";
        const barW = "MAX(0.05,1-" + ln + "-" + lp + ")";
        const lo = "=" + mn + "-" + spanE + "*" + ln + "/" + barW;
        const hi = "=" + mx + "+" + spanE + "*" + lp + "/" + barW;
        const neg = bar.invert ? bar.pos : bar.neg;
        const pos = bar.invert ? bar.neg : bar.pos;
        ranges.forEach(({ rg }) => {
            const cf = rg.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
            const db = cf.dataBar;
            db.showDataBarOnly = barOnly;
            db.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
            db.axisFormat = Excel.ConditionalDataBarAxisFormat.automatic;
            db.axisColor = "#000000";
            db.positiveFormat.fillColor = pos;
            db.positiveFormat.gradientFill = false;
            db.negativeFormat.matchPositiveFillColor = false;
            db.negativeFormat.fillColor = neg;
            db.lowerBoundRule = { type: Excel.ConditionalFormatRuleType.formula, formula: lo };
            db.upperBoundRule = { type: Excel.ConditionalFormatRuleType.formula, formula: hi };
        });
    }
    DracoCalcCellsByReport.set(dracoStateKey(reportId), set);
}

/**
 * Modo Estático, tras pintar los valores:
 *   - líneas EPM_FORMULA (siempre): fórmula, formato y barras;
 *   - "Sobrescribir formatos en informes estáticos": cabeceras con estilo
 *     Draco y valores con el formato de su medida.
 * En el cruce de una fila EPM_FORMULA con una columna EPM_FORMULA manda
 * la de la columna (p. ej. la Variación de la fila Margen).
 */
async function finishDracoFixedReport(context, reportId, editReportGrid) {
    const props = dracoReportPropsFor(reportId);
    const resultSheetName = await getDracoResultSheetName(context, reportId);
    const sheet = context.workbook.worksheets.getItem(resultSheetName);
    let rowLines = await scanDracoStaticAxis(context, sheet, editReportGrid, "rows");
    let colLines = await scanDracoStaticAxis(context, sheet, editReportGrid, "cols");
    DracoCalcCellsByReport.set(dracoStateKey(reportId), new Set());

    if (props.overwriteStaticFormats) {
        await formatDracoStaticAxis(context, sheet, editReportGrid, "rows", rowLines);
        await formatDracoStaticAxis(context, sheet, editReportGrid, "cols", colLines);
        const defaultMeasure = (ReportState.FilterMeasureNames && ReportState.FilterMeasureNames[0])
            || (ReportState.Measures[0] && ReportState.Measures[0].Name) || null;
        const dataCols = colLines.filter(c => !c.isFormula);
        const cols = dataCols.map(c => c.line).sort((a, b) => a - b);
        for (const rl of rowLines.filter(r => !r.isFormula)) {
            const fmtAt = (idx) => dracoMeasureNumberFormatByName(reportId,
                rl.measure || (dataCols.find(c => c.line === cols[idx]) || {}).measure || defaultMeasure);
            let start = 0;
            for (let i = 1; i <= cols.length; i++) {
                const breakRun = i === cols.length || cols[i] !== cols[i - 1] + 1 || fmtAt(i) !== fmtAt(i - 1);
                if (breakRun) {
                    const rg = sheet.getRangeByIndexes(rl.line - 1, cols[start] - 1, 1, cols[i - 1] - cols[start] + 1);
                    rg.numberFormat = [[fmtAt(start)]];
                    rg.format.font.name = DRACO_FONT_NAME;
                    rg.format.font.size = DRACO_FONT_SIZE;
                    rg.format.horizontalAlignment = Excel.HorizontalAlignment.center;
                    start = i;
                }
            }
        }
    }

    // Filas EPM_FORMULA: en las columnas de datos. Columnas EPM_FORMULA: en
    // todas las filas (datos y calculadas).
    applyDracoFormulaLines(sheet, reportId, rowLines, true, colLines.filter(c => !c.isFormula).map(c => c.line));
    const before = DracoCalcCellsByReport.get(dracoStateKey(reportId));
    applyDracoFormulaLines(sheet, reportId, colLines, false, rowLines.map(r => r.line));
    if (before) before.forEach(x => DracoCalcCellsByReport.get(dracoStateKey(reportId)).add(x));
    await context.sync();
}

async function actualizarInformeHybridCore(reportIdOverride) {
    const reportId = reportIdOverride !== undefined ? reportIdOverride : activeReportIdOrNull();
    let built;
    beginDracoPerfLog();
    const tRefreshStart = performance.now();
    let tPerf = tRefreshStart;

    await draco_beginManualCalculation();
    await flushDracoPlanningModifiedCells();
    if (window.BusyIndicator) await window.BusyIndicator.show(busyStepLabel(reportId, "Generando SQL"));
    await beginSuppressPlanningPaint();
    try {
        let editReportGridForPaint;
        await Excel.run(async (context) => {
            const editReportGrid = await getEditReportGrid(context, reportId);
            editReportGridForPaint = editReportGrid;
            const relGrid = await window.SemanticModelStore.getModelGrid("MODEL_RELATIONSHIP");
            const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");
            const atributesGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
            const resultSheetName = resultSheetNameFromGrid(editReportGrid, reportId);
            await ensureDracoResultSheetExists(context, resultSheetName);
            const csvGrid = await getFormulaGrid(context, resultSheetName);

            loadReportDefinition(editReportGrid, reportId);
            built = await buildSQLHybrid(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid, reportId);
            await context.sync();
        });
        tPerf = draco_perfMark(reportId, "Híbrido: generando SQL", tPerf);

        let json = "";
        if (built.sql) {
            if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(reportId, "Ejecutando consulta"));
            json = await executeSQL(built.sql);
            tPerf = draco_perfMark(reportId, "Híbrido: ejecutando consulta (red)", tPerf);
        }

        await Excel.run(async (context) => {
            const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
            editReportSheet.getRange("X1").values = [[draco_debugCell(built.sql || "(Híbrido: el eje Estático no tiene ninguna línea EPM_VALUE todavía)")]];
            const LIMIT = 32000;
            editReportSheet.getRange("Y1").values = [[draco_debugCell(json.length > LIMIT ? json.substring(0, LIMIT) + " ...(truncado)" : json)]];
            await context.sync();
        });

        if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(reportId, "Pintando resultados"));
        await Excel.run(async (context) => {
            await jsonPaintHybrid(context, json, reportId, built.plan, editReportGridForPaint);
        });
    } finally {
        await endSuppressPlanningPaint();
        await draco_endManualCalculation();
        // Con el cálculo ya restaurado, las cabeceras EPM_VALUE muestran su
        // texto: ahora sí se puede autoajustar.
        try { await autofitDracoHybrid(reportId, built && built.plan); } catch (e) { console.warn("[Draco] Autoajuste híbrido:", e); }
        if (window.BusyIndicator) await window.BusyIndicator.hide();
        draco_perfMark(reportId, "══ TOTAL refresco híbrido ══", tRefreshStart);
        await endDracoPerfLog();
    }
}

/**
 * Autoajusta las columnas de lo pintado por el híbrido (cabecera del eje
 * dinámico y, si las Filas son dinámicas, también su bloque de valores),
 * DESPUÉS de recalcular, para que las fórmulas EPM_VALUE ya muestren su
 * texto y no se queden las columnas estrechas.
 */
async function autofitDracoHybrid(reportId, plan) {
    if (!plan || !dracoReportPropsFor(reportId).autoFitColumns) return;
    await Excel.run(async (context) => {
        try { context.workbook.application.calculate(Excel.CalculationType.recalculate); } catch (e) { /* host sin calculate */ }
        await context.sync();
        const rn = dracoRangeNames(reportId);
        const items = [plan.dynamicIsColumns ? rn.cols : rn.rows, rn.values]
            .map(n => context.workbook.names.getItemOrNullObject(n));
        items.forEach(it => it.load("name"));
        await context.sync();
        items.forEach(it => { if (!it.isNullObject) it.getRange().format.autofitColumns(); });
        await context.sync();
    });
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
// aggregationOverride (opcional): la agregación elegida en "Opciones de
// campo" de la medida (SUM/AVG/COUNT/COUNT_DISTINCT/MIN/MAX). Vacía =
// la del modelo semántico (MODEL_MEASURES!G), como hasta ahora.
function getMeasureSQL(measuresGrid, measureName, aggregationOverride, scale) {
    const R = buscarMedida(measuresGrid, measureName);
    if (R === 0) return "";
    const aggregation = String(cellValue(measuresGrid, R, 7)).trim(); // G: AGGREGATION
    const field = String(cellValue(measuresGrid, R, 6)).trim();       // F: FACT_FIELD
    // Override de Opciones de campo, si no la del modelo. Si la del modelo
    // es una de las conocidas se pasa por el mismo mapeo (así
    // COUNT_DISTINCT, que el editor del modelo permite elegir, genera
    // COUNT(DISTINCT ...) en vez de un "COUNT_DISTINCT(...)" inválido);
    // cualquier otra se deja tal cual, como siempre.
    const effective = normalizeDracoAggregation(aggregationOverride) || normalizeDracoAggregation(aggregation);
    if (effective) {
        return dracoScaled(DRACO_SQL_AGGREGATIONS[effective]("f." + field), scale) + " AS " + measureName;
    }
    return dracoScaled(aggregation + "(f." + field + ")", scale) + " AS " + measureName;
}

/**
 * BuildSelect: columnas del eje Columnas, luego eje Filas, luego medidas.
 */
function buildSelect(measuresGrid, relGrid) {
    let sql = "SELECT" + CRLF;
    let first = true;

    for (const c of ReportState.Columns) {
        if (!first) sql += "," + CRLF;
        sql += "    " + dracoAxisCol(relGrid, c) + (c.SqlAttr ? " AS " + c.AttributeName : "");
        first = false;
    }

    for (const r of ReportState.Rows) {
        if (!first) sql += "," + CRLF;
        sql += "    " + dracoAxisCol(relGrid, r) + (r.SqlAttr ? " AS " + r.AttributeName : "");
        first = false;
    }

    for (const m of ReportState.Measures) {
        if (!first) sql += "," + CRLF;
        sql += "    " + getMeasureSQL(measuresGrid, m.Name, m.Aggregation, m.Scale);
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

/**
 * Para el writeback de Draco Planning: si un filtro (Zona "Filtros" de un
 * informe) representa EXACTAMENTE un valor incluido — un único valor/item,
 * sin exclusiones (excludeValues/excludeItems/excludeRanges, ni el
 * include:false del formato antiguo) y sin rangos — devuelve ese valor
 * (el canónico: values[0], o item.value si viene del selector de
 * miembros/jerarquía). Con 0 valores, 2+ valores, exclusiones o un rango,
 * es ambiguo qué escribir en el INSERT, así que devuelve undefined: esa
 * dimensión sigue contando como "sin resolver" (falta), igual que antes.
 */
function singleDracoFilterValueForWriteback(rawValue) {
    const filter = parseStoredFilterValue(rawValue);
    if (!filter) return undefined;

    const hasExclusions =
        (filter.excludeValues && filter.excludeValues.length) ||
        (filter.excludeItems && filter.excludeItems.length) ||
        (filter.excludeRanges && filter.excludeRanges.length) ||
        filter.include === false; // formato antiguo "values" con NOT
    if (hasExclusions) return undefined;

    if (filter.ranges && filter.ranges.length) return undefined; // rango: no hay un único valor discreto

    const values = filter.values || [];
    const items = filter.items || [];
    if (values.length + items.length !== 1) return undefined; // 0 o 2+: ambiguo

    if (values.length === 1) return values[0];
    const it = items[0];
    return (it && typeof it === "object") ? (it.value !== undefined ? it.value : it.display) : it;
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

/**
 * Condiciones de los filtros del informe (zona "Filtros" del taskpane) Y
 * de los filtros "bloqueados" (botón "Añadir filtro", tanto los atados a
 * ESTE informe como los de "Todos los informes") — ReportState.Filters ya
 * los trae todos juntos, cargados por loadReportDefinition
 * (loadFilters + appendLockedFilterRangesToState). Devuelve un array de
 * condiciones SQL sueltas (sin "WHERE" ni "AND" delante) — cada llamante
 * decide cómo combinarlas: buildWhere (modo Dinámico) las usa como ÚNICO
 * WHERE; buildSQLFixed (modo Fijo) las añade con AND al bloque OR ya
 * existente (ver comentario en buildSQLFixed).
 */
function buildFilterConditions(atributesGrid, relGrid) {
    // Filtros "vacíos" (sin valor seleccionado en el taskpane) no deben
    // añadirse al WHERE: se ignoran por completo, como si no existieran.
    const activeFilters = ReportState.Filters.filter(f => String(f.Value).trim() !== "");
    if (activeFilters.length === 0) return [];

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

    return condiciones;
}

function buildWhere(atributesGrid, relGrid) {
    const condiciones = buildFilterConditions(atributesGrid, relGrid);
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
    // ---- Campos "lógicos" de cada eje ----
    // Los niveles de la jerarquía de TEXTOS de un "ID y texto" (IsTextCompanion)
    // no son niveles nuevos: cada uno se agrupa SIEMPRE junto al nivel de ID
    // de su misma dimensión y NIVEL (si no, el "odómetro" generaría
    // combinaciones de ID de un nivel con texto de otro). Se cuelgan de ese
    // campo ("extra") y se emiten juntos en cada conjunto de agrupación.
    function logicalFields(fields) {
        const out = [];
        for (const f of fields) {
            if (f.IsTextCompanion) {
                let host = null;
                for (let k = out.length - 1; k >= 0; k--) {
                    if (out[k].f.Dimension === f.Dimension && Number(out[k].f.Hierarchy) === Number(f.Hierarchy)) {
                        host = out[k];
                        break;
                    }
                }
                if (host) {
                    host.extra.push(f);
                    continue;
                }
                // sin nivel de ID al que pegarse: se trata como un campo normal
            }
            out.push({ f, extra: [] });
        }
        return out;
    }

    const logicalRows = logicalFields(ReportState.Rows);
    const logicalCols = logicalFields(ReportState.Columns);

    // ---- DIMENSIONS: Filas primero, luego Columnas ----
    // Cada posición es la columna del campo o, si lleva texto pegado, un
    // array [columna, columnaTexto] (ver groupingSetTupleToColumnList).
    const dimensions = [null]; // índice 1-based, dimensions[0] sin usar

    for (const l of [...logicalRows, ...logicalCols]) {
        const cols = [dracoAxisCol(relGrid, l.f), ...l.extra.map(e => dracoAxisCol(relGrid, e))];
        dimensions.push(cols.length === 1 ? cols[0] : cols);
    }

    // ---- HIERARCHIES: longitud de cada tramo de "misma JERARQUÍA real" ----
    // Un tramo NO es simplemente "misma Dimension consecutiva": eso
    // confundía dos atributos PLANOS sueltos de una misma dimensión (p.ej.
    // Región y Ciudad, cada uno con NIVEL=1 -ver loadRows/loadColumns y
    // ReportStore.saveDesign/expandAxis, que graba NIVEL=1 fijo para
    // cualquier campo que NO sea una jerarquía) con una jerarquía real
    // desplegada (NIVEL 1,2,3... consecutivos de ReportStore.saveDesign/
    // expandAxis para un campo isHierarchy). Con la condición anterior
    // (solo Dimension), dos atributos sueltos de la misma dimensión
    // generaban subtotales/anidado de jerarquía que nadie pidió. Ahora un
    // tramo solo continúa si, ADEMÁS de compartir Dimension, el NIVEL es
    // exactamente el siguiente consecutivo (1,2,3...); dos campos con
    // NIVEL=1 seguidos (aunque sean de la misma dimensión) rompen el tramo
    // y se tratan como dos tramos de longitud 1 -igual que si fueran de
    // dimensiones distintas-, sin generar ningún subtotal automático.
    function pushHierarchyRunLengths(logical, out) {
        let lastDim = "";
        let lastLevel = 0;
        let contador = 0;
        for (const l of logical) {
            const f = l.f;
            const nivel = Number(f.Hierarchy) || 0;
            if (contador > 0 && f.Dimension === lastDim && nivel === lastLevel + 1) {
                contador++;
            } else {
                if (contador > 0) out.push(contador);
                contador = 1;
            }
            lastDim = f.Dimension;
            lastLevel = nivel;
        }
        if (contador > 0) out.push(contador);
    }

    const hierarchies = [];
    pushHierarchyRunLengths(logicalRows, hierarchies);
    pushHierarchyRunLengths(logicalCols, hierarchies);

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
    // Sin campos en Filas ni Columnas (solo medidas): total general, sin
    // GROUP BY (un "GROUPING SETS ( () )" no lo admiten todos los motores).
    if (ReportState.Rows.length === 0 && ReportState.Columns.length === 0) return "";
    return groupingSetTuplesToText(configSetTuples(relGrid));
}

// Combinaciones de GROUPING SETS (listas de columnas) de buildConfigSets
function configSetTuples(relGrid) {
    const { dimensions, hierarchiesAcum, combinations } = computeGroupingSetDimensionsAndHierarchies(relGrid);
    return combinations.map(item => groupingSetTupleToColumnList(item, dimensions, hierarchiesAcum));
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
                const dimEntry = i === 0 ? dimensions[j] : dimensions[j + hierarchiesAcum[i - 1]];
                // Un campo con texto pegado ("ID y texto") aporta sus dos columnas.
                if (Array.isArray(dimEntry)) cols.push(...dimEntry);
                else cols.push(dimEntry);
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
    return groupingSetTuplesToText(configSetTuplesWithSubtotals(relGrid));
}

// Combinaciones de GROUPING SETS (listas de columnas) de buildConfigSetsWithSubtotals
function configSetTuplesWithSubtotals(relGrid) {
    const { dimensions, hierarchiesAcum, combinations } = computeGroupingSetDimensionsAndHierarchies(relGrid);
    const baseTuples = combinations.map(item => groupingSetTupleToColumnList(item, dimensions, hierarchiesAcum));

    const colsAxisFields = ReportState.Columns.map(c => dracoAxisCol(relGrid, c));
    const rowsAxisFields = ReportState.Rows.map(r => dracoAxisCol(relGrid, r));

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

    return [...baseTuples, ...extraTuples];
}

/**
 * BuildFinalSelect: traducción literal, incluyendo el cruce ROW_ID/COLUMN_ID
 * "invertido" tal cual está en tu VBA (ROW_ID se calcula a partir del eje
 * Columnas y COLUMN_ID a partir del eje Filas) — se deja igual, según lo
 * acordado.
 */
// "DENSE_RANK() OVER (ORDER BY ...) AS <alias>" a partir de las partes
// del ORDER BY; si el eje está vacío (no hay ningún campo en él), una
// constante 1: todo el resultado cae en una única fila/columna lógica
// (DENSE_RANK sin ORDER BY no es portable).
function buildDenseRankExpression(orderParts, alias) {
    if (orderParts) orderParts = orderParts.map(dracoCalcOrderPart);
    if (!orderParts || orderParts.length === 0) {
        return "    1 AS " + alias;
    }
    return "    DENSE_RANK() OVER (" + CRLF + "        ORDER BY" + CRLF
        + orderParts.map(p => "            " + p).join("," + CRLF) + CRLF
        + "    ) AS " + alias;
}

function buildFinalSelect() {
    let sql = "SELECT" + CRLF + CRLF;

    sql += buildDenseRankExpression(ReportState.Columns.filter(c => c.Hierarchy > 0).map(c => c.AttributeName), "ROW_ID") + "," + CRLF + CRLF;
    sql += buildDenseRankExpression(ReportState.Rows.filter(r => r.Hierarchy > 0).map(r => r.AttributeName), "COLUMN_ID");

    // Campos de Columnas, de Filas y una columna por medida (antes:
    // "IMPORTE" fijo). Con lista común para que un eje vacío o un informe
    // sin medidas no deje comas sueltas.
    const items = [
        ...ReportState.Columns.map(c => c.AttributeName),
        ...ReportState.Rows.map(r => r.AttributeName),
        ...ReportState.Measures.map(m => m.Name)
    ];
    if (items.length > 0) sql += "," + CRLF + CRLF + items.map(it => "    " + it).join("," + CRLF);
    sql += CRLF + CRLF + "FROM REPORT_DATA";

    return sql;
}

/**
 * BuildFinalWhere: excluye de cada eje las filas de detalle que no
 * correspondan al grouping-set de totales de esa jerarquía (IS NULL).
 */
function buildFinalWhere() {
    return buildFinalWhereFromAxes([ReportState.Columns, ReportState.Rows]);
}

/**
 * Para cada eje recibido, agrupa sus campos (Hierarchy > 0) en tramos
 * consecutivos de la misma dimensión y genera "NOT ( a IS NULL AND b IS
 * NULL ... )" por tramo, todo unido con AND. Mismo resultado que el bucle
 * original de buildFinalWhere cuando los dos ejes tienen campos, pero sin
 * dejar paréntesis/AND sueltos si alguno (o los dos) está vacío.
 */
function buildFinalWhereFromAxes(axes) {
    const groups = [];
    for (const fields of axes) {
        let currentDim = "";
        let current = null;
        for (const f of fields) {
            if (!(f.Hierarchy > 0)) continue;
            const dim = String(f.Dimension).toUpperCase();
            if (!current || dim !== currentDim) {
                current = [];
                groups.push(current);
                currentDim = dim;
            }
            current.push(f.AttributeName + " IS NULL");
        }
    }
    if (groups.length === 0) return "";
    return groups
        .map((g, idx) => (idx === 0 ? "WHERE NOT (" : "AND NOT (") + CRLF + "    " + g.join(CRLF + "    AND ") + CRLF + ")")
        .join(CRLF + CRLF);
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
        sql += "," + CRLF + "    GROUPING(" + dracoAxisCol(relGrid, c) + ") AS " + subtotalFlagName(c.AttributeName);
    }
    for (const r of ReportState.Rows) {
        sql += "," + CRLF + "    GROUPING(" + dracoAxisCol(relGrid, r) + ") AS " + subtotalFlagName(r.AttributeName);
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
        ? ("CAST(" + field.AttributeName + " AS " + Provider.stringType() + ")")
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

    sql += buildDenseRankExpression(buildAxisOrderByWithSubtotals(ReportState.Columns, subtotalsOnTop), "ROW_ID") + "," + CRLF + CRLF;
    sql += buildDenseRankExpression(buildAxisOrderByWithSubtotals(ReportState.Rows, subtotalsOnTop), "COLUMN_ID");

    const items = [
        ...ReportState.Columns.map(c => fieldFinalExpressionWithSubtotal(atributesGrid, c)),
        ...ReportState.Rows.map(r => fieldFinalExpressionWithSubtotal(atributesGrid, r)),
        ...ReportState.Measures.map(m => m.Name)
    ];
    if (items.length > 0) sql += "," + CRLF + CRLF + items.map(it => "    " + it).join("," + CRLF);
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
    const colsHaveSubtotal = ReportState.Columns.some(c => c.Subtotal);
    const rowsHaveSubtotal = ReportState.Rows.some(r => r.Subtotal);
    return buildFinalWhereFromAxes([
        colsHaveSubtotal ? [] : ReportState.Columns,
        rowsHaveSubtotal ? [] : ReportState.Rows
    ]);
}

/**
 * BuildSQL: ensambla el flujo "Dinamico" completo. Si NINGUN campo tiene
 * marcado un subtotal (columnas L/R de EDIT_REPORT), usa exactamente las
 * mismas funciones de siempre (SQL identico al de antes de este cambio).
 * Si hay al menos uno, usa las variantes "WithSubtotals" de arriba.
 */
/**
 * REPORT_DATA del modo dinámico cuando alguna medida tiene agregación de excepción. Mismas columnas que el
 * REPORT_DATA de siempre, así que el resto del SQL (cálculos, SELECT final) no cambia:
 *   DRC_BASE    = la consulta de siempre + por cada columna de los GROUPING SETS su valor (DRC_K_i) y su
 *                 GROUPING() (DRC_G_i), que identifican la fila (conjunto + valores);
 *   DRC_EXC_n   = por cada medida con excepción: la medida agregada en cada conjunto AMPLIADO con las
 *                 dimensiones de referencia y, encima, la agregación de excepción por (conjunto + valores);
 *   REPORT_DATA = DRC_BASE con la columna de esas medidas tomada de su DRC_EXC_n.
 * Una dimensión de referencia que ya está en el informe se agrupa como COALESCE(c, c): así es otro elemento de
 * agrupación y GROUPING(c) sigue diciendo si c está en el conjunto (si no, en los subtotales no casaría).
 */
function buildDracoExcReportData(relGrid, measuresGrid, atributesGrid, withSubtotals) {
    const hasDims = ReportState.Rows.length > 0 || ReportState.Columns.length > 0;
    const tuples = hasDims ? (withSubtotals ? configSetTuplesWithSubtotals(relGrid) : configSetTuples(relGrid)) : [];
    const keys = [];
    tuples.forEach(t => t.forEach(c => { if (keys.indexOf(c) === -1) keys.push(c); }));
    const keyCols = keys.map((k, i) => "    GROUPING(" + k + ") AS DRC_G_" + i + "," + CRLF + "    " + k + " AS DRC_K_" + i);
    const keyNames = keys.map((k, i) => "DRC_G_" + i + ", DRC_K_" + i);
    const from = buildFrom(measuresGrid, null, relGrid);
    const joins = buildJoins(relGrid);
    const where = buildWhere(atributesGrid, relGrid);

    let sql = "WITH DRC_BASE AS (" + CRLF + CRLF;
    sql += (withSubtotals ? buildSelectWithSubtotals(measuresGrid, relGrid) : buildSelect(measuresGrid, relGrid))
        + (keyCols.length ? "," + CRLF + keyCols.join("," + CRLF) : "") + CRLF + CRLF;
    sql += from + CRLF + CRLF + joins + CRLF + CRLF + where + CRLF + CRLF;
    sql += (hasDims ? groupingSetTuplesToText(tuples) : "") + CRLF + CRLF + ")";

    const excMs = ReportState.Measures.filter(m => m.Exc);
    excMs.forEach((m, n) => {
        const ref = dracoExcRef(relGrid, m.Exc);
        const refExpr = ref.cols.map(c => keys.indexOf(c) !== -1 ? "COALESCE(" + c + ", " + c + ")" : c);
        const seen = new Set(), sets = [];
        tuples.forEach(t => { const x = t.concat(refExpr), k = x.slice().sort().join("|"); if (!seen.has(k)) { seen.add(k); sets.push(x); } });
        sql += "," + CRLF + "DRC_EXC_" + (n + 1) + " AS (" + CRLF + CRLF
            + "SELECT" + CRLF + (keyNames.length ? "    " + keyNames.join("," + CRLF + "    ") + "," + CRLF : "")
            + "    " + m.Exc.agg + "(DRC_V) AS DRC_V" + CRLF
            + "FROM (" + CRLF + CRLF
            + "SELECT" + CRLF + (keyCols.length ? keyCols.join("," + CRLF) + "," + CRLF : "")
            + "    " + dracoMeasureExpr(measuresGrid, m) + " AS DRC_V" + CRLF + CRLF
            + from + CRLF + CRLF + joins + ref.joins.join("") + CRLF + CRLF + where + CRLF + CRLF
            + (hasDims ? groupingSetTuplesToText(sets) : "GROUP BY " + refExpr.join(", ")) + CRLF + CRLF
            + ") DRC_X" + CRLF
            + (keyNames.length ? "GROUP BY " + keyNames.join(", ") + CRLF : "") + CRLF + ")";
    });

    const cols = [
        ...ReportState.Columns.map(c => "b." + c.AttributeName),
        ...ReportState.Rows.map(r => "b." + r.AttributeName),
        ...ReportState.Measures.map(m => {
            const n = excMs.indexOf(m);
            return n === -1 ? "b." + m.Name : "e" + (n + 1) + ".DRC_V AS " + m.Name;
        })
    ];
    if (withSubtotals) {
        ReportState.Columns.forEach(c => cols.push("b." + subtotalFlagName(c.AttributeName)));
        ReportState.Rows.forEach(r => cols.push("b." + subtotalFlagName(r.AttributeName)));
    }
    sql += "," + CRLF + "REPORT_DATA AS (" + CRLF + CRLF + "SELECT" + CRLF + cols.map(c => "    " + c).join("," + CRLF) + CRLF + "FROM DRC_BASE b";
    excMs.forEach((m, n) => {
        const e = "e" + (n + 1);
        if (!keys.length) { sql += CRLF + "CROSS JOIN DRC_EXC_" + (n + 1) + " " + e; return; }
        sql += CRLF + "LEFT JOIN DRC_EXC_" + (n + 1) + " " + e + CRLF + "    ON "
            + keys.map((k, i) => "b.DRC_G_" + i + " = " + e + ".DRC_G_" + i + " AND b.DRC_K_" + i + " IS NOT DISTINCT FROM " + e + ".DRC_K_" + i)
                .join(CRLF + "   AND ");
    });
    sql += CRLF + CRLF + ")";
    return sql;
}

/* Columna física de un campo de los ejes: su atributo real (SqlAttr si se le ha dado otro nombre, ver dracoDedupAxisNames). */
function dracoAxisCol(relGrid, f) {
    return getTableAlias(relGrid, f.Dimension) + "." + (f.SqlAttr || f.AttributeName);
}
/**
 * Dos campos de los ejes con el MISMO nombre de atributo (p. ej. ACCOUNT.DESCRIPTION y COMPANY.DESCRIPTION) darían
 * dos columnas iguales en REPORT_DATA y el motor lo rechaza ("duplicate column names"). Al segundo y siguientes se les
 * da un nombre propio (DESCRIPTION_COMPANY) que usan todas las consultas; la columna se sigue leyendo con su nombre
 * real (SqlAttr). El pintado vuelve a cargar la definición, así que los encabezados no cambian. Sin repetidos, el SQL
 * es exactamente el de siempre.
 */
function dracoDedupAxisNames() {
    const seen = new Set();
    [...(ReportState.Columns || []), ...(ReportState.Rows || [])].forEach(f => {
        if (!f || !f.AttributeName || f.SqlAttr) { if (f && f.AttributeName) seen.add(String(f.AttributeName).toUpperCase()); return; }
        const up = String(f.AttributeName).toUpperCase();
        if (!seen.has(up)) { seen.add(up); return; }
        const base = String(f.AttributeName) + "_" + String(f.Dimension || "").replace(/[^A-Za-z0-9_]/g, "_");
        let alias = base, n = 2;
        while (seen.has(alias.toUpperCase())) alias = base + "_" + (n++);
        f.SqlAttr = f.AttributeName;
        f.AttributeName = alias;
        seen.add(alias.toUpperCase());
    });
}

function buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop) {
    dracoDedupAxisNames();
    if (ReportState.Rows.length === 0 && ReportState.Columns.length === 0 && ReportState.Measures.length === 0) {
        throw new Error("Añade al menos un campo a Filas o Columnas.");
    }
    let sql = "";
    const withSubtotals = hasAnySubtotalMarked();

    if ((ReportState.Measures || []).some(m => m.Exc)) {
        // Alguna medida con agregación de excepción: REPORT_DATA se compone (ver buildDracoExcReportData)
        sql += buildDracoExcReportData(relGrid, measuresGrid, atributesGrid, withSubtotals);
    } else {
    sql += "WITH REPORT_DATA AS (" + CRLF + CRLF;
    sql += (withSubtotals ? buildSelectWithSubtotals(measuresGrid, relGrid) : buildSelect(measuresGrid, relGrid)) + CRLF + CRLF;
    sql += buildFrom(measuresGrid, null, relGrid) + CRLF + CRLF;
    sql += buildJoins(relGrid) + CRLF + CRLF;
    sql += buildWhere(atributesGrid, relGrid) + CRLF + CRLF;
    sql += (withSubtotals ? buildConfigSetsWithSubtotals(relGrid) : buildConfigSets(relGrid)) + CRLF + CRLF;
    sql += ")";
    }
    // Cálculos entre miembros (si el informe tiene): CTEs encadenadas y el
    // SELECT final lee de la última en vez de REPORT_DATA.
    ReportState.CalcOrder = {};
    const calcCtes = buildDracoCalcCtes(relGrid, measuresGrid, atributesGrid, withSubtotals);
    if (calcCtes) sql += calcCtes.sql;
    sql += CRLF + CRLF;
    let finalSelect = withSubtotals ? buildFinalSelectWithSubtotals(atributesGrid, !!subtotalsOnTop) : buildFinalSelect();
    if (calcCtes) finalSelect = finalSelect.replace(/FROM REPORT_DATA$/, "FROM " + calcCtes.last);
    sql += finalSelect + CRLF + CRLF;
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

// true mientras handleDracoPickerFlagRequest está procesando de verdad una
// petición DC/REC (desde que empieza su Excel.run hasta que limpia T2:V2).
// El VBA del XLAM no espera a que el add-in termine antes de escribir el
// siguiente doble clic en T2:V2 ("el filtrado fino se hace aquí", ver
// comentario más arriba), así que varios dobles clics rápidos disparan
// varios onChanged casi seguidos. Antes, cada uno de ellos entraba a hacer
// su PROPIO Excel.run + findDracoRowsNamedRangeForCell/resolveDracoFieldForAxisLevel
// (varias vueltas context.sync cada uno) ANTES de llegar al único guard que
// existía (DracoMemberPickerOpen, dentro de openMemberRecognitionPicker) —
// con Shared Runtime todo esto compite por el mismo hilo JS único del
// panel+ribbon, así que una ráfaga de N clics amontona N cadenas de
// Excel.run en la misma cola, y el add-in entero (incluida la ventana del
// picker que sí llegó a abrirse) se queda esperando a que esa cola drene.
// Esto es lo que de verdad explica "se pilla con clics rápidos y tarda
// minutos en las consultas siguientes", no la velocidad de BigQuery. Este
// guard corta la ráfaga ANTES de tocar Excel: solo la primera petición
// DC/REC en curso se procesa; las que llegan mientras tanto se ignoran de
// inmediato (sin Excel.run, sin coste), y en cuanto la primera limpia
// T2:V2 (picker cerrado), un doble clic nuevo vuelve a procesarse normal.
let DracoPickerFlagRequestBusy = false;

// Candado DEDICADO y COMPLETAMENTE INDEPENDIENTE de DracoSuppressChangeEvents
// (a propósito: reutilizar ese flag, que ya protege muchas otras cosas —
// picker, expandir/contraer, etc. — para esto acabó rompiendo otras partes
// del add-in). Solo lo usa handleDracoPlanningValueChanged (marcar en cian
// una celda tocada a mano de un informe de planificación).
//
// Es un CONTADOR (no un simple booleano) porque "Actualizar todos" envuelve
// también, uno a uno, a actualizarInformeFixedCore: con un booleano simple,
// el primer informe fijo procesado lo pondría a "false" en su propio
// finally y reactivaría el pintado ANTES de que terminaran de refrescarse
// los siguientes informes de ese mismo "Actualizar todos". Mientras el
// contador sea > 0, sigue deshabilitado.
//
// Y (herencia del runtime clásico, antes de migrar a Shared Runtime) lleva
// ADEMÁS un contador espejo en EDIT_REPORT!Z2: cuando el ribbon
// (commands.html) y el taskpane (taskpane.html) eran dos procesos JS
// separados, cada uno tenía su PROPIA copia de esta variable en memoria,
// y el contador local por sí solo no bastaba para avisar al otro proceso
// de que había un refresco en curso. Con Shared Runtime ya no hay dos
// procesos que sincronizar, así que el contador local ya sería
// suficiente — se mantiene el espejo en Z2 por ahora porque otros flujos
// (p.ej. el XLAM en VBA) pueden seguir leyéndolo; revisar si sigue
// haciendo falta al limpiar EDIT_REPORT.
let DracoSuppressPlanningPaintCount = 0;
const DRACO_PLANNING_SUPPRESS_CELL = "Z2";

// PLANIFICACIÓN > registro EN MEMORIA (solo dura la sesión, no se
// persiste en el libro) de las celdas de Draco_<id>_Values que el
// usuario ha tocado a mano desde el último "Guardar planificación" o
// "Actualizar", junto con el color que tenían ANTES de pintarse en cian
// (ver handleDracoPlanningValueChanged). Se consulta/vacía en
// flushDracoPlanningModifiedCells, que escribe el listado en
// EDIT_REPORT!A127, devuelve a cada celda su color original y limpia
// este registro — se llama tanto desde guardarPlanificacion como desde
// el arranque de cualquier refresco (actualizarInformeCore/
// actualizarInformeFixedCore/actualizarTodosCore).
//
// Estructura: Map<reportId, Map<direccion, { sheetName, address, color }>>.
// Se indexa también por reportId (aunque el volcado a A127 es global,
// de todos los informes a la vez) porque handleDracoPlanningValueChanged
// ya conoce el reportId de la hoja tocada y así una misma dirección de
// celda en dos informes/hojas distintos no se pisa entre sí.
const DracoPlanningModifiedCells = new Map();
const DRACO_PLANNING_LOG_CELL = "A127";

/**
 * Vuelca en EDIT_REPORT!A127 (como JSON) todas las celdas de
 * planificación modificadas a mano registradas en
 * DracoPlanningModifiedCells, les devuelve el color que tenían antes de
 * pintarse en cian, y vacía el registro. La llaman tanto
 * guardarPlanificacion (botón "Guardar") como el arranque de cualquier
 * refresco (para que un refresco no deje el registro con celdas que ya
 * no reflejan lo que hay pintado en la hoja).
 *
 * No hace nada (ni siquiera toca A127) si no hay ninguna celda
 * registrada, para no pisar un listado previo si esta función se llama
 * varias veces seguidas sin cambios de por medio (p.ej. "Actualizar
 * todos", que internamente vuelve a llamar a esta función una vez por
 * cada informe en modo "Fijo").
 */
async function flushDracoPlanningModifiedCells() {
    if (DracoPlanningModifiedCells.size === 0) return;

    // Aplanar todo lo registrado (de todos los informes tocados) en una
    // única lista para el JSON de A127, y agrupar por hoja para repintar
    // con el mínimo de context.sync().
    const flat = [];
    const bySheet = new Map();
    for (const [reportId, bucket] of DracoPlanningModifiedCells.entries()) {
        for (const entry of bucket.values()) {
            flat.push({ reportId, sheet: entry.sheetName, address: entry.address, color: entry.color || "" });
            if (!bySheet.has(entry.sheetName)) bySheet.set(entry.sheetName, []);
            bySheet.get(entry.sheetName).push(entry);
        }
    }

    try {
        await Excel.run(async (context) => {
            const editReportSheet = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
            editReportSheet.load("isNullObject");
            await context.sync();

            if (!editReportSheet.isNullObject) {
                const EXCEL_CELL_CHAR_LIMIT = 32000; // límite real de Excel: 32767
                let jsonForCell = JSON.stringify(flat);
                if (jsonForCell.length > EXCEL_CELL_CHAR_LIMIT) {
                    jsonForCell = jsonForCell.substring(0, EXCEL_CELL_CHAR_LIMIT) + " ...(truncado)";
                }
                editReportSheet.getRange(DRACO_PLANNING_LOG_CELL).values = [[jsonForCell]];
            }

            for (const [sheetName, entries] of bySheet.entries()) {
                const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                sheet.load("isNullObject");
                await context.sync();
                if (sheet.isNullObject) continue;

                for (const entry of entries) {
                    const cell = sheet.getRange(entry.address);
                    if (entry.color) {
                        cell.format.fill.color = entry.color; // color que tenía antes de marcarse
                    } else {
                        cell.format.fill.clear(); // no tenía relleno (color "" / automático)
                    }
                }
                await context.sync();
            }
        });
    } catch (e) {
        console.error("[Draco] Error volcando en EDIT_REPORT!A127 y/o restaurando el color de las celdas de planificación modificadas:", e);
    }

    // Vaciar SIEMPRE el registro, incluso si algo ha fallado al escribir
    // o repintar, para no arrastrar un estado inconsistente a la
    // siguiente edición del usuario.
    DracoPlanningModifiedCells.clear();
}

async function beginSuppressPlanningPaint() {
    DracoSuppressPlanningPaintCount++;
    try {
        await Excel.run(async (context) => {
            // Antes: 1 sync para saber si EDIT_REPORT existe + 1 sync más
            // para leer el candado (Z2) = 2 vueltas de red solo para LEER,
            // antes incluso de escribir nada. Se cargan las dos cosas
            // juntas y se resuelven en un único sync (si EDIT_REPORT no
            // existiera, el .getRange() de abajo fallaría al resolverse,
            // pero el try/catch de fuera ya lo cubre igual que antes).
            const editReport = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
            editReport.load("isNullObject");
            const cell = editReport.getRange(DRACO_PLANNING_SUPPRESS_CELL);
            cell.load("values");
            await context.sync();
            if (editReport.isNullObject) return;

            const current = Number(cell.values && cell.values[0] && cell.values[0][0]) || 0;
            cell.values = [[String(current + 1)]];
            await context.sync();
        });
    } catch (e) {
        console.warn("[Draco] No se pudo incrementar el candado cruzado de planificación (EDIT_REPORT!Z2):", e);
    }
}

async function endSuppressPlanningPaint() {
    DracoSuppressPlanningPaintCount = Math.max(0, DracoSuppressPlanningPaintCount - 1);
    try {
        await Excel.run(async (context) => {
            // Misma optimización que beginSuppressPlanningPaint: 1 sync
            // para leer (existencia + candado) en vez de 2.
            const editReport = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
            editReport.load("isNullObject");
            const cell = editReport.getRange(DRACO_PLANNING_SUPPRESS_CELL);
            cell.load("values");
            await context.sync();
            if (editReport.isNullObject) return;

            const current = Number(cell.values && cell.values[0] && cell.values[0][0]) || 0;
            cell.values = [[String(Math.max(0, current - 1))]];
            await context.sync();
        });
    } catch (e) {
        console.warn("[Draco] No se pudo decrementar el candado cruzado de planificación (EDIT_REPORT!Z2):", e);
    }
}

// true mientras ya hay un diálogo de "buscador de miembros" abierto. Se
// usa como candado para evitar que se abran dos diálogos a la vez para
// una misma petición (p.ej. clic sobre el icono de jerarquía y una
// petición REC/DC de EDIT_REPORT!T2:V2 casi simultáneas).
//
// Con Shared Runtime (manifest con <Runtimes>), el panel y el ribbon
// comparten un único proceso/memoria JS, así que este flag local ya
// basta: no puede haber un segundo proceso con su propio
// DracoMemberPickerOpen en `false` compitiendo por abrir otro diálogo.
// ANTES (runtime clásico, sin <Runtimes> en el manifest) hacía falta
// ADEMÁS un candado visible entre procesos, escrito en la propia hoja
// (EDIT_REPORT!Z1, vía tryClaimDracoPickerLock/releaseDracoPickerLock).
// Se ha quitado al migrar a Shared Runtime: ya no hay dos procesos que
// sincronizar.
let DracoMemberPickerOpen = false;

/**
 * HISTÓRICO (antes de migrar a Shared Runtime): el botón "Reconocimiento
 * de miembros" corría en un runtime separado del panel (FunctionFile
 * clásico: commands.html y taskpane.html eran dos procesos JS distintos),
 * y Office.context.document.settings es una caché LOCAL de cada runtime
 * — así que cuando toggleMemberRecognition hacía settings.set(...) en el
 * runtime del ribbon, el del panel (donde vive el reconocimiento) seguía
 * viendo el valor antiguo. Con Shared Runtime ya no aplica (un único
 * proceso, una única caché), pero se deja esta función leyendo de
 * EDIT_REPORT!B1 igualmente: sigue siendo válida y no hay motivo para
 * quitarla ahora que ya no es estrictamente necesaria.
 */
async function isDracoMemberRecognitionActive(context) {
    try {
        const editReport = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
        editReport.load("isNullObject");
        await context.sync();
        if (editReport.isNullObject) {
            // Sin EDIT_REPORT no hay fuente de verdad en la hoja: se cae
            // al valor cacheado localmente (mejor que nada).
            return !!Office.context.document.settings.get("draco_memberRecognition");
        }
        const flagCell = editReport.getRange("B1");
        flagCell.load("values");
        await context.sync();
        return String((flagCell.values && flagCell.values[0] && flagCell.values[0][0]) || "").trim().toUpperCase() === "X";
    } catch (e) {
        console.warn("[Draco] No se pudo leer EDIT_REPORT!B1 (Reconocimiento de miembros), se usa el valor cacheado:", e);
        return !!Office.context.document.settings.get("draco_memberRecognition");
    }
}

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
    // "Mostrar nodos de jerarquía arriba": igual que subtotalsOnTop, pero
    // para los nodos-resumen que vienen de una JERARQUÍA real (ver
    // computeGroupingSetDimensionsAndHierarchies), no de un subtotal manual
    // (columnas L/R "X" de EDIT_REPORT). Antes las dos cosas compartían el
    // mismo interruptor (subtotalsOnTop); ahora son independientes.
    const hierarchyNodesOnTop = opts.hierarchyNodesOnTop !== false;

    // Posiciones (1..count) de la jerarquía de TEXTOS de un "ID y texto": no
    // son niveles (su valor puede ser NULL de verdad si el miembro no tiene
    // descripción), así que no cuentan para la profundidad del nodo.
    const companion = opts.companion || [];

    const items = [];
    for (const V of dict.values()) {
        let deepest = 0;
        for (let i = 1; i <= count; i++) {
            if (companion[i]) continue;
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

    // Longitud del tramo de JERARQUÍA real (NIVEL ascendente 1,2,3... —
    // ver computeGroupingSetDimensionsAndHierarchies) al que pertenece cada
    // posición del eje; 1 si es un campo plano (o el único nivel de una
    // jerarquía de un solo nivel). `flags` ya trae el NIVEL (Hierarchy) de
    // cada posición (ver computeAxisPaintPlan/flagsReal). Un campo plano
    // SIEMPRE se guarda con NIVEL=1 (ver ReportStore.saveDesign/expandAxis),
    // así que un tramo de longitud 1 solo puede producir un nodo-resumen a
    // través de la marca manual "Subtotal" (columnas L/R), nunca de una
    // jerarquía real — de ahí que sirva para distinguir el origen del nodo.
    const tramoLen = new Array(count + 2).fill(1);
    {
        let i = 1;
        while (i <= count) {
            let j = i;
            while (j < count && Number(flags[j + 1]) === Number(flags[j]) + 1 && !!companion[j + 1] === !!companion[j]) j++;
            const len = j - i + 1;
            for (let k = i; k <= j; k++) tramoLen[k] = len;
            i = j + 1;
        }
    }

    // Reordena los nodos-resumen según corresponda: los que vienen de una
    // jerarquía real usan "Mostrar nodos de jerarquía arriba"; los que
    // vienen de un subtotal marcado a mano (columnas L/R) usan "Mostrar
    // subtotales arriba". Por defecto (ambas propiedades a "arriba") se
    // respeta el orden natural del resultado (el resumen sale justo antes
    // que su detalle, que es como suele venir el GROUPING SETS); cuando la
    // propiedad correspondiente está desactivada, ese nodo-resumen concreto
    // pasa a ir DESPUÉS de todos sus descendientes.
    items.sort((a, b) => {
        const len = Math.min(a.deepest, b.deepest);
        let samePrefix = true;
        for (let i = 1; i <= len; i++) {
            if (String(a.V[i]) !== String(b.V[i])) { samePrefix = false; break; }
        }
        if (samePrefix && a.deepest !== b.deepest) {
            const boundaryField = len + 1;
            const onTop = tramoLen[boundaryField] > 1 ? hierarchyNodesOnTop : subtotalsOnTop;
            if (onTop) return a.oldId - b.oldId; // orden natural: el resumen ya viene antes
            // Uno es ancestro (resumen) del otro: el ancestro va DESPUÉS.
            return a.deepest < b.deepest ? 1 : -1;
        }
        return a.oldId - b.oldId; // sin relación de parentesco directa: orden natural
    });

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
 * arriba, nodos de jerarquía arriba, sobrescribir formatos, autoajustar
 * columnas), guardadas por el modal "Propiedades del informe" del taskpane
 * en Office roaming settings. Accesible desde cualquier contexto (taskpane
 * o commands.html/ribbon).
 */
function getDracoReportProperties() {
    const defaults = {
        reportName: "Report 001",
        suppressZeroRows: false,
        suppressZeroCols: false,
        subtotalsOnTop: false,
        hierarchyNodesOnTop: true,
        overwriteFormats: true,
        overwriteStaticFormats: false,
        autoFitColumns: true,
        planningReport: false
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
// ya saben leer para el flujo Fijo). `display` es opcional: si no se pasa,
// se usa el mismo texto que `value` (caso normal, Valor === Display). Se
// pasa distinto solo para la fila/columna MEASURE (ver
// convertAxisStaticFormulas): una medida no tiene "valor de miembro", así
// que Valor queda vacío y Display lleva el nombre de la medida.
function buildEpmValueFormula(dim, attr, value, display) {
    const esc = (s) => String(s === null || s === undefined ? "" : s).replace(/"/g, '""');
    const disp = display !== undefined ? display : value;
    return '=EPM_VALUE("' + esc(dim) + '","' + esc(attr) + '","' + esc(value) + '","' + esc(disp) + '")';
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
//
// CORREGIDO: antes esta tabla guardaba UN SOLO {dim, attr} por columna
// física (iAux), tomado siempre de la primera fila de EDIT_REPORT de ese
// campo (la de NIVEL/flag=1). Eso está bien para campos simples, pero en
// una JERARQUÍA multinivel, cada profundidad (flag=1,2,3...) puede tener
// su propio Attr distinto (p.ej. "NIVEL1", "NIVEL2"...) aunque compartan
// columna física — y al convertir a Estático se aplicaba SIEMPRE el Attr
// del nivel 1 a todas las filas pintadas de esa columna, sin importar su
// profundidad real (de ahí que toda la jerarquía se pintase con NIVEL1).
// Ahora se guarda, por columna física, un {dim, attr} POR CADA flag/nivel
// (byFlag), para poder elegir el correcto según la profundidad real de
// cada celda pintada (ver resolveAxisFieldForFlag).
//
// CORREGIDO (2): las filas/columnas MEASURE ya NO se saltan — antes se
// dejaban sin entrada en esta tabla ("las filas MEASURE no pintan como
// Estáticas"), así que la celda con el nombre de la medida (p.ej.
// "IMPORTE") se quedaba siempre en texto plano. Ahora se les da su propio
// hueco físico (igual que measureLevels en computeAxisPaintPlan) con
// {dim:"MEASURE", attr:<nombre de la medida>}, y el bucle de conversión
// más abajo genera para ellas =EPM_VALUE("MEASURE","<nombre>","","<nombre>")
// (Valor vacío: una medida no tiene "valor de miembro", solo nombre).
function buildAxisFieldLevelsTable(editReportGrid, axis) {
    const [dimCol, attrCol, hierCol, jerCol] = axis === "rows" ? [8, 9, 10, 11] : [14, 15, 16, 17];
    const levels = buildDracoAxisLevels(editReportGrid, dimCol, attrCol, hierCol, jerCol);
    const columns = [];
    let iAux = 0;
    let current = null;
    // Grupo MEASURE: TODAS las filas/columnas MEASURE del eje comparten un
    // único hueco físico (igual que measureLevels en computeAxisPaintPlan),
    // en la posición en la que aparezca la PRIMERA de ellas — no
    // necesariamente la última. attr = nombre de la medida (columna I/O,
    // ver loadRows/loadColumns), dim = "MEASURE" fijo (no lvl.dim, para no
    // arrastrar mayúsculas/minúsculas tal y como se haya escrito en la
    // hoja).
    let measureIaux = null;
    for (const lvl of levels) {
        if (lvl.isMeasure) {
            if (measureIaux === null) {
                iAux++;
                measureIaux = iAux;
                columns[measureIaux] = { byFlag: {}, maxFlag: 0 };
            }
            columns[measureIaux].byFlag[1] = { dim: "MEASURE", attr: lvl.attr, jerarquia: "" };
            columns[measureIaux].maxFlag = 1;
            continue;
        }
        if (lvl.hidden) continue; // "Solo texto": el ID no tiene columna
        const flag = lvl.flag || 1;
        if (flag === 1 || !current) {
            iAux++;
            current = { byFlag: {}, maxFlag: 0 };
            columns[iAux] = current;
        }
        current.byFlag[flag] = { dim: lvl.dim, attr: lvl.attr, jerarquia: lvl.jerarquia };
        if (flag > current.maxFlag) current.maxFlag = flag;
    }
    return columns;
}

// Dada la entrada de buildAxisFieldLevelsTable para UNA columna física y
// el "flag" (nivel de profundidad, 1-based = indentLevel+1) de la celda
// pintada, devuelve el {dim, attr} del nivel correspondiente. Si ese
// nivel exacto no está definido (p.ej. una fila con más indentado del
// que tiene el eje configurado), usa el nivel más profundo conocido; si
// tampoco hay nada, cae al nivel 1.
function resolveAxisFieldForFlag(columnEntry, flag) {
    if (!columnEntry) return null;
    if (columnEntry.byFlag[flag]) return columnEntry.byFlag[flag];
    if (flag > columnEntry.maxFlag && columnEntry.byFlag[columnEntry.maxFlag]) {
        return columnEntry.byFlag[columnEntry.maxFlag];
    }
    return columnEntry.byFlag[1] || null;
}

// Lee, para cada celda de `range` (mismo orden que range.values), su
// format.indentLevel actual, en UNA sola llamada a la API (en vez de un
// cell.load por celda). Devuelve una matriz [row][col] de números.
async function readIndentLevelsForRange(context, range) {
    try {
        const result = range.getCellProperties({ format: { indentLevel: true } });
        await context.sync();
        return result.value.map(row =>
            row.map(c => (c && c.format && typeof c.format.indentLevel === "number") ? c.format.indentLevel : 0)
        );
    } catch (err) {
        // Fallback para Excel sin soporte de getCellProperties (< ExcelApi
        // 1.9): una llamada de load por celda, todas en el mismo sync.
        console.warn("[Draco] getCellProperties no disponible, usando fallback celda a celda para indentLevel:", err);
        range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
        await context.sync();
        const cells = [];
        for (let r = 0; r < range.rowCount; r++) {
            const rowCells = [];
            for (let c = 0; c < range.columnCount; c++) {
                const cell = range.getCell(r, c);
                cell.load("format/indentLevel");
                rowCells.push(cell);
            }
            cells.push(rowCells);
        }
        await context.sync();
        return cells.map(row => row.map(cell => cell.format.indentLevel || 0));
    }
}


/**
 * "Solo texto": las celdas de texto de un eje con IDs ocultos NO se pueden
 * convertir in situ entre texto y EPM_VALUE (el ID no está en la hoja), así
 * que se repinta desde el último resultado. Requiere el diseño ya guardado
 * con el nuevo Estático/Dinámico. Devuelve true si el eje tiene IDs ocultos
 * (se haya repintado o no, porque aún no hubiese resultado).
 */
async function repaintAxisIfHiddenIds(axis) {
    const reportId = activeReportIdOrNull();
    let hasHidden = false;
    await Excel.run(async (context) => {
        const editReportGrid = await getEditReportGrid(context);
        const [dimCol, attrCol, hierCol] = axis === "rows" ? [8, 9, 10] : [14, 15, 16];
        hasHidden = buildDracoAxisLevels(editReportGrid, dimCol, attrCol, hierCol).some(l => l.hidden);
        if (!hasHidden) return;
        const lastJson = DracoLastJsonByReport.get(dracoStateKey(reportId));
        if (lastJson) await jsonTo3Matrices(context, lastJson, reportId);
    });
    return hasHidden;
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
        const fieldLevels = buildAxisFieldLevelsTable(editReportGrid, axis);

        const namedRange = context.workbook.names.getItemOrNullObject(rangeName);
        namedRange.load("isNullObject");
        await context.sync();
        if (namedRange.isNullObject) {
            // Todavía no hay tabla pintada: nada que convertir ahora mismo.
            return;
        }

        const range = namedRange.getRange();
        range.load(["values", "text", "formulas", "rowCount", "columnCount"]);
        await context.sync();

        // Profundidad real (indentLevel) de cada celda ya pintada: para una
        // jerarquía, cada fila/columna puede estar a un nivel distinto
        // aunque compartan columna/fila física, y el Attr correcto (p.ej.
        // "NIVEL1" vs "NIVEL2") depende de esa profundidad, no solo de la
        // posición del campo dentro del eje.
        const indentLevels = await readIndentLevelsForRange(context, range);

        const GLYPH_PREFIX = /^[▸▾]\s+/; // indicador +/- fusionado (solo aplica en eje Dinámico)
        const EPM_RE = /^=\s*EPM_VALUE\s*\(/i;
        const EPM_VALOR_RE = /EPM_VALUE\s*\(\s*"(?:[^"]|"")*"\s*,\s*"(?:[^"]|"")*"\s*,\s*"((?:[^"]|"")*)"/i;
        // Solo para la fila/columna MEASURE (ver comentario en
        // buildAxisFieldLevelsTable): su fórmula lleva el Valor (3er
        // argumento) vacío y el nombre real en el Display (4º argumento),
        // al revés que un campo normal.
        const EPM_MEASURE_RE = /^=\s*EPM_VALUE\s*\(\s*"MEASURE"/i;
        const EPM_DISPLAY_RE = /EPM_VALUE\s*\(\s*"(?:[^"]|"")*"\s*,\s*"(?:[^"]|"")*"\s*,\s*"(?:[^"]|"")*"\s*,\s*"((?:[^"]|"")*)"/i;

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

                // Columna/fila física del campo dentro del eje: en Filas
                // crece por COLUMNA, en Columnas crece por FILA (misma
                // orientación que al pintar).
                const iAux = (axis === "rows" ? c : r) + 1;
                const columnEntry = fieldLevels[iAux];
                // Profundidad real de ESTA celda (1-based, como el flag de
                // EDIT_REPORT): indentLevel 0 -> nivel 1, indentLevel 1 ->
                // nivel 2, etc.
                const flag = (indentLevels[r][c] || 0) + 1;
                const field = resolveAxisFieldForFlag(columnEntry, flag);
                const isEpmFormula = typeof currentFormula === "string" && EPM_RE.test(currentFormula);

                if (makeStatic) {
                    if (isEpmFormula || !field) {
                        rowOut.push(currentFormula);
                        continue;
                    }
                    // Texto TAL COMO SE VE en la celda (.text), no el valor
                    // crudo (.values) — para una celda con formato de fecha,
                    // .values da el número de serie de Excel (p.ej. 46054),
                    // no la fecha; .text da lo que de verdad está pintado.
                    const text = String(range.text[r][c]).replace(GLYPH_PREFIX, "");
                    if (field.dim === "MEASURE") {
                        // Una medida no tiene "valor de miembro": Valor
                        // vacío, Display = nombre de la medida.
                        rowOut.push(buildEpmValueFormula("MEASURE", field.attr, "", text));
                    } else {
                        rowOut.push(buildEpmValueFormula(field.dim, field.attr, text));
                    }
                    changed = true;
                } else {
                    if (!isEpmFormula) {
                        rowOut.push(currentFormula);
                        continue;
                    }
                    const isMeasureFormula = EPM_MEASURE_RE.test(currentFormula);
                    const re = isMeasureFormula ? EPM_DISPLAY_RE : EPM_VALOR_RE;
                    const match = currentFormula.match(re);
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
    cell.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "format/indentLevel"]);

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

    const indentLevel = cell.format.indentLevel || 0;
    return resolveDracoFieldForAxisLevel(context, reportId, axis, level, indentLevel);
}

/**
 * Núcleo de locateDracoAxisField, extraído para poder reutilizarlo desde
 * flujos que YA conocen (reportId, axis, level) por otra vía —por
 * ejemplo, el clic multi-informe de handleDracoRowsSingleClick, que
 * puede caer en la "zona de crecimiento" de un informe (fuera todavía
 * del rango con nombre real), donde el chequeo de pertenencia por rango
 * de locateDracoAxisField no serviría—.
 *
 * `indentLevel` es el format.indentLevel (0-based) de la celda pulsada;
 * para una celda vacía (típicamente la zona de crecimiento, donde aún no
 * hay ninguna fila/columna pintada) se puede pasar 0 sin problema: cae
 * en el primer nivel del campo, que es el comportamiento razonable al
 * "crear" una fila/columna nueva por ese hueco.
 * Devuelve {axis, level, dim, attr} o null si ese nivel no corresponde a
 * ningún campo real del diseño del informe.
 */
async function resolveDracoFieldForAxisLevel(context, reportId, axis, level, indentLevel) {
    const editReportGrid = await getEditReportGrid(context, reportId);
    loadReportDefinition(editReportGrid);
    const fieldLevels = buildAxisFieldLevelsTable(editReportGrid, axis);
    const columnEntry = fieldLevels[level];
    // Profundidad real de la celda pulsada (1-based, como el flag de
    // EDIT_REPORT): antes se usaba siempre el Attr del primer nivel del
    // campo (p.ej. "NIVEL1"), aunque el clic fuera sobre un nodo más
    // profundo de la jerarquía (NIVEL2, NIVEL3...).
    const flag = (indentLevel || 0) + 1;
    const field = resolveAxisFieldForFlag(columnEntry, flag);
    if (!field || !field.dim) return null;

    // Igual que el VBA original (GetDimAttrRows/GetDimAttrCols): si el
    // campo es una JERARQUÍA, el buscador de miembros debe abrir el
    // árbol COMPLETO de la jerarquía (todos los niveles), no la lista
    // plana de un único nivel — así que aquí se usa el nombre de la
    // Jerarquía en vez del Attr del nivel concreto. Attr (NIVEL1/NIVEL2)
    // solo se usa para la fórmula EPM_VALUE final, no para buscar.
    // La jerarquía de TEXTOS de un "ID y texto" (sufijo ~TXT) no existe en el
    // modelo: se busca por el atributo de texto del nivel.
    const searchAttr = (field.jerarquia && !dracoIsHierTextName(field.jerarquia)) ? field.jerarquia : field.attr;

    return { axis, level, dim: field.dim, attr: searchAttr, levelAttr: field.attr };
}

// NOTA: aquí existía handleDracoMemberRecognitionChanged, la función
// legada de "reconocimiento de miembros" que reaccionaba con onChanged
// directamente sobre la hoja de resultados. Ya no estaba enganchada a
// ningún evento (ver registerDracoSelectionHandler) porque esa detección
// la hace ahora el XLAM en VBA, avisando por EDIT_REPORT!T2:V2 con
// V2="REC" (ver handleDracoPickerFlagRequest / runDracoMemberRecognitionAction).
// Se ha eliminado por completo. locateDracoAxisField, que solo usaba
// esta función, ha quedado sin ninguna llamada.

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

// NOTA: aquí existían handleDracoFilterRangeSelection y
// ensureDracoFilterRangeHandlersRegistered (onSelectionChanged sobre
// TODAS las hojas, para abrir el selector de "Añadir filtro" al
// seleccionar la celda del rango). Se han eliminado por completo: ya no
// queda ningún listener de onSelectionChanged en el add-in.
// locateDracoFilterRangeAtAddress se conserva (por si se reutiliza desde
// otro punto de entrada), pero ha quedado sin ninguna llamada.

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
 * Regla de reparto del hueco entre dos informes consecutivos del mismo
 * eje (sin margen: todo el hueco es de alguien):
 *   - El informe ANTERIOR (más arriba/izquierda) se queda con TODO el
 *     hueco hasta el siguiente informe (hasta la celda justo antes de
 *     donde empieza el siguiente).
 *   - El informe SIGUIENTE nunca crece hacia atrás (hacia el hueco).
 *   - Delante del PRIMER informe de un eje no crece nadie.
 *   - El ÚLTIMO informe de un eje en una hoja (el que no tiene otro
 *     informe después en esa misma familia de columnas/filas) se queda
 *     con TODO lo que sigue, hasta el final absoluto de la hoja.
 */

const DRACO_LAST_ROW_INDEX = 1048576 - 1;   // última fila de Excel (0-based)
const DRACO_LAST_COLUMN_INDEX = 16384 - 1;  // última columna de Excel (0-based)

/**
 * Devuelve, para una hoja dada, todos los rangos con nombre
 * "Draco_<n>_<axisSuffix>" ("Rows" o "Cols") definidos en ESA hoja, con
 * sus límites (rowIndex/columnIndex/rowCount/columnCount, 0-based) ya
 * cargados.
 */
async function collectDracoAxisCandidates(context, sheet, axisSuffix) {
    const names = context.workbook.names;
    names.load("items/name");
    await context.sync();

    const pattern = new RegExp(`^Draco_(\\d+)_${axisSuffix}$`, "i");
    const matching = names.items.filter(n => pattern.test(n.name));
    if (matching.length === 0) return [];

    const loaded = matching.map(n => {
        const range = context.workbook.names.getItem(n.name).getRange();
        range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
        const ws = range.worksheet;
        ws.load("name");
        return { name: n.name, range, ws };
    });
    await context.sync();

    return loaded
        .filter(c => c.ws.name === sheet.name)
        .map(c => ({
            name: c.name,
            reportId: Number(c.name.match(pattern)[1]),
            rowIndex: c.range.rowIndex,
            rowCount: c.range.rowCount,
            columnIndex: c.range.columnIndex,
            columnCount: c.range.columnCount
        }));
}

/**
 * Localiza a qué informe Draco_<n>_<axisSuffix> pertenece —o PODRÍA
 * pertenecer, ampliando su rango con nombre— una celda concreta, sin
 * invadir jamás el hueco/rango de otro informe (ver reglas en
 * ver comentario sobre el reparto del hueco más arriba).
 *
 * axisSuffix: "Rows" (el eje que varía es la FILA; la celda debe caer
 * dentro de las mismas columnas que el informe) o "Cols" (el eje que
 * varía es la COLUMNA; la celda debe caer dentro de las mismas filas).
 *
 * Devuelve null si la celda no pertenece ni podría llegar a pertenecer a
 * ningún informe de ese eje en esa hoja. Si pertenece, devuelve:
 *   {
 *     rangeName, reportId, sheetName,
 *     exact,          // true = ya está dentro del rango con nombre actual
 *     axisStart,      // 0-based: inicio actual del rango en el eje que varía
 *     axisCount,      // 0-based: tamaño actual del rango en ese eje
 *     requiredCount,  // tamaño que DEBERÍA tener el rango para incluir la
 *                     // celda (== axisCount si exact=true)
 *     crossIndex, crossCount // límites (0-based) del eje que NO varía
 *                            // (columnas para _Rows, filas para _Cols)
 *   }
 */
async function findDracoAxisRangeOrGrowthZoneForCell(context, worksheetId, addr, axisSuffix) {
    const sheet = context.workbook.worksheets.getItem(worksheetId);
    sheet.load("name");
    const cell = sheet.getRange(addr);
    cell.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();

    if (cell.rowCount !== 1 || cell.columnCount !== 1) return null;

    const candidates = await collectDracoAxisCandidates(context, sheet, axisSuffix);
    if (candidates.length === 0) return null;

    const isRows = axisSuffix.toLowerCase() === "rows";
    const cellAxisIndex = isRows ? cell.rowIndex : cell.columnIndex;
    const cellCrossIndex = isRows ? cell.columnIndex : cell.rowIndex;

    // Solo son candidatos "de la misma familia" los que comparten con la
    // celda pulsada el eje que NO varía (mismas columnas para _Rows,
    // mismas filas para _Cols) — cada uno con SUS PROPIOS límites, no se
    // exige que todos los informes tengan exactamente el mismo ancho.
    const sameFamily = candidates.filter(c => {
        const crossStart = isRows ? c.columnIndex : c.rowIndex;
        const crossCount = isRows ? c.columnCount : c.rowCount;
        return cellCrossIndex >= crossStart && cellCrossIndex < crossStart + crossCount;
    });
    if (sameFamily.length === 0) return null;

    // 1) Coincidencia EXACTA: la celda ya está dentro de algún rango.
    for (const c of sameFamily) {
        const start = isRows ? c.rowIndex : c.columnIndex;
        const count = isRows ? c.rowCount : c.columnCount;
        if (cellAxisIndex >= start && cellAxisIndex < start + count) {
            return {
                rangeName: c.name,
                reportId: c.reportId,
                sheetName: sheet.name,
                exact: true,
                axisStart: start,
                axisCount: count,
                requiredCount: count,
                crossIndex: isRows ? c.columnIndex : c.rowIndex,
                crossCount: isRows ? c.columnCount : c.rowCount,
                cellCrossIndex
            };
        }
    }

    // 2) Sin coincidencia exacta: el hueco entre dos informes consecutivos
    //    (o entre el último informe y el final de la hoja) es SIEMPRE del
    //    informe ANTERIOR — nunca queda un tramo sin reclamar entre ellos,
    //    y el informe SIGUIENTE nunca crece hacia atrás.
    const sorted = sameFamily.slice().sort((a, b) => {
        const sa = isRows ? a.rowIndex : a.columnIndex;
        const sb = isRows ? b.rowIndex : b.columnIndex;
        return sa - sb;
    });
    const MAX_AXIS_INDEX = isRows ? DRACO_LAST_ROW_INDEX : DRACO_LAST_COLUMN_INDEX;

    for (let i = 0; i < sorted.length; i++) {
        const cur = sorted[i];
        const start = isRows ? cur.rowIndex : cur.columnIndex;
        const count = isRows ? cur.rowCount : cur.columnCount;
        const end = start + count - 1; // último índice (0-based) ya ocupado por este informe

        // Si la celda está a la altura de este informe o antes de él, no
        // le pertenece a ÉL (o ya habría salido en el paso 1 si estaba
        // dentro): seguimos mirando el siguiente informe de la lista.
        if (cellAxisIndex <= end) continue;

        const next = sorted[i + 1];
        const isLast = !next;
        // Sin margen: todo el hueco hasta el siguiente informe (o hasta
        // el final absoluto de la hoja si es el último) es de "cur".
        const zoneEnd = isLast
            ? MAX_AXIS_INDEX
            : (isRows ? next.rowIndex : next.columnIndex) - 1;

        if (cellAxisIndex <= zoneEnd) {
            return {
                rangeName: cur.name,
                reportId: cur.reportId,
                sheetName: sheet.name,
                exact: false,
                axisStart: start,
                axisCount: count,
                requiredCount: (cellAxisIndex - start) + 1,
                crossIndex: isRows ? cur.columnIndex : cur.rowIndex,
                crossCount: isRows ? cur.columnCount : cur.rowCount,
                cellCrossIndex
            };
        }
        // No debería llegar aquí (con el reparto sin margen, cualquier
        // celda posterior a "cur" y anterior al siguiente informe cae
        // siempre dentro de su zona) — se deja como salvaguarda por si
        // sameFamily contuviera candidatos con huecos inconsistentes.
    }

    return null;
}

/**
 * Amplía (o crea) el rango con nombre `rangeName` para que cubra desde
 * `axisStart` hasta incluir `requiredCount` filas/columnas en el eje que
 * varía, conservando intacto el eje que no varía (crossIndex/crossCount).
 * Sigue el mismo patrón (borrar + volver a crear el nombre) que ya usa
 * applyDracoNamedRanges al repintar un informe.
 */
async function growDracoNamedRange(context, sheetName, rangeName, axisSuffix, axisStart, requiredCount, crossIndex, crossCount) {
    const isRows = axisSuffix.toLowerCase() === "rows";
    const sheet = context.workbook.worksheets.getItem(sheetName);

    const newRange = isRows
        ? sheet.getRangeByIndexes(axisStart, crossIndex, requiredCount, crossCount)
        : sheet.getRangeByIndexes(crossIndex, axisStart, crossCount, requiredCount);

    const existing = context.workbook.names.getItemOrNullObject(rangeName);
    existing.load("isNullObject");
    await context.sync();
    if (!existing.isNullObject) {
        existing.delete();
        await context.sync();
    }
    context.workbook.names.add(rangeName, newRange);
    await context.sync();
}

/**
 * Busca, entre todos los nombres del libro que cumplan el patrón
 * "Draco_<n>_Rows" o "Draco_<n>_Cols", cuál (si alguno) contiene la
 * celda indicada EN LA MISMA HOJA del clic —o podría llegar a
 * contenerla ampliando su rango con nombre, sin invadir el hueco de
 * otro informe (ver findDracoAxisRangeOrGrowthZoneForCell)—.
 *
 * Devuelve {rangeName, reportId, axis, level, sheetName, exact, growth}
 * o null si el clic no cae ni podría llegar a caer en ninguna tabla de
 * Filas/Columnas de Draco. `axis` es "rows" o "columns" (mismo
 * vocabulario que locateDracoAxisField). `level` es la posición
 * (1-based) dentro del eje que NO varía (columna para _Rows, fila para
 * _Cols). `growth` es null si `exact` es true; si `exact` es false,
 * trae todo lo necesario para ampliar de verdad el rango con nombre
 * (ver growDracoNamedRange) en cuanto el usuario confirme un valor en
 * el picker.
 */
async function findDracoRowsNamedRangeForCell(context, worksheetId, addr) {
    let found = await findDracoAxisRangeOrGrowthZoneForCell(context, worksheetId, addr, "Rows");
    let axis = "rows";
    if (!found) {
        found = await findDracoAxisRangeOrGrowthZoneForCell(context, worksheetId, addr, "Cols");
        axis = "columns";
    }
    if (!found) return null;

    return {
        rangeName: found.rangeName,
        reportId: found.reportId,
        axis,
        // Posición (1-based) de la celda dentro del eje que NO varía
        // (columna dentro de _Rows, fila dentro de _Cols) — igual que
        // antes calculaba el "within" exacto, pero también válido en la
        // zona de crecimiento (el eje que no varía nunca se amplía).
        level: (found.cellCrossIndex - found.crossIndex) + 1,
        sheetName: found.sheetName,
        exact: found.exact,
        growth: found.exact ? null : {
            rangeName: found.rangeName,
            axisSuffix: axis === "rows" ? "Rows" : "Cols",
            sheetName: found.sheetName,
            axisStart: found.axisStart,
            requiredCount: found.requiredCount,
            crossIndex: found.crossIndex,
            crossCount: found.crossCount
        }
    };
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

// NOTA: antes había aquí una detección "manual" de doble clic (Office.js
// no expone un evento nativo de doble clic sobre una celda). Se ha
// quitado porque el toggle del icono de jerarquía ▾/▸ se resuelve con el
// clic SIMPLE (ver más abajo, rama "resultado === 'Letra' && hasHierarchyGlyph"):
// nunca se llegó a usar el resultado del doble clic para esa acción.
async function handleDracoRowsSingleClick(eventArgs) {
    try {
        let addr = (eventArgs && eventArgs.address) || "";
        if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
        addr = addr.replace(/\$/g, "").toUpperCase();
        if (!addr) return;

        const offsetX = eventArgs.offsetX;
        const offsetY = eventArgs.offsetY;

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);

            const located = await findDracoRowsNamedRangeForCell(context, eventArgs.worksheetId, addr);
            if (!located) {
                // Cae fuera de cualquier Draco_XXX_Rows/Cols (y fuera
                // también de la zona de crecimiento de cualquiera de
                // ellos): no hay nada que hacer.
                return;
            }
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

            if (resultado === "Letra" && hasHierarchyGlyph) {
                // El clic ha caído sobre el icono ▾/▸: expandir/contraer el
                // nodo, igual que si se hubiera rellenado EDIT_REPORT!T1:V1.
                const toggled = await toggleDracoCollapseAtCell(context, located.sheetName, addr);
                accionTexto = toggled
                    ? " | Accion: expandir/contraer ejecutado"
                    : " | Accion: sin indicador +/- valido en esa celda";
            }
            // ELIMINADO: doble clic "real" fuera del icono +/- que abría el
            // buscador de miembros (resolveDracoFieldForAxisLevel +
            // openMemberRecognitionPicker). El clic en Draco_XXX_Rows/Cols
            // ahora SOLO expande/contrae jerarquía (rama de arriba); ya no
            // hace nada más.

            console.log(
                `[Draco] Clic en ${located.rangeName} (nivel ${located.level}) -> offsetX=${offsetX.toFixed(2)}pt | ` +
                `letra estimada entre ${bounds.start.toFixed(2)}pt y ${bounds.end.toFixed(2)}pt -> ${resultado} | Jerarquia: ${jerarquiaTexto}${accionTexto}`
            );

            await context.sync();
        });
    } catch (e) {
        console.error("[Draco] Error registrando el clic en Draco_*_Rows:", e);
    }
}

/**
 * Se dispara con onChanged en cada hoja de resultados (mismo registro por
 * hoja que handleDracoRowsSingleClick, ver ensureDracoRowsClickLoggerRegistered
 * más abajo): por cada celda del bloque tocado (una sola celda al teclear,
 * o varias a la vez con arrastrar/pegar/relleno) que caiga dentro de un
 * rango con nombre Draco_<n>_Values Y ese informe tenga marcada la
 * propiedad "Informe de planificación" (Propiedades del informe >
 * planningReport), pinta su fondo en RGB(223,255,255) — marca visual de
 * "editada a mano, pendiente de guardar planificación".
 *
 * Se ignora sin hacer nada mientras:
 *   - DracoSuppressChangeEvents esté activo: son nuestras propias
 *     escrituras (p.ej. la fórmula EPM_VALUE que deja el picker, o el
 *     propio jsonTo3Matrices pintando la tabla), no una edición manual.
 *   - DracoSuppressPlanningPaintCount > 0: hay un refresco en curso
 *     (Actualizar / Actualizar todos / Guardar planificación, ver
 *     beginSuppressPlanningPaint/endSuppressPlanningPaint) — si no,
 *     CADA refresco dejaría todo el informe pintado en cian, como si el
 *     usuario lo hubiera editado a mano entero.
 *
 * Con Shared Runtime solo hay un proceso registrando esto (antes, con el
 * runtime clásico, el ribbon y el panel podían llegar a registrar este
 * mismo onChanged cada uno por su cuenta, marcando la celda dos veces
 * seguidas — inofensivo pero redundante). Al ser un único proceso, este
 * controlador se registra una sola vez, sin duplicados.
 *
 * ORDEN DE TRABAJO (así se piden las cosas a Excel, de menos a más caro):
 *   1) Informes de planificación cuya hoja de resultados sea ESTA hoja —
 *      puro JS contra ReportStore (localStorage), CERO idas y vueltas a
 *      Excel. Si no hay ninguno, se corta aquí: la inmensa mayoría de
 *      ediciones en hojas sin ningún informe de planificación cuestan UNA
 *      sola ida y vuelta (la de abajo, para tener sheet.name), no cuatro.
 *   2) Solo el rango Draco_<id>_Values de ESOS candidatos en concreto
 *      (por nombre exacto), no la colección completa de nombres del
 *      libro — evita traer y filtrar en el cliente decenas de nombres
 *      que no pintan nada aquí.
 *   3) Intersección con la dirección modificada — en memoria.
 *   4) Solo entonces se lee el formato y se escribe, y solo de las
 *      celdas que de verdad han caído dentro de un rango de planificación.
 *
 * El color que tenía cada celda ANTES de pintarse se guarda en
 * DracoPlanningModifiedCells (si no estaba ya registrada: si el usuario
 * la toca varias veces seguidas mientras sigue en cian, no se debe
 * "olvidar" el color original de antes de la primera edición, y tampoco
 * hace falta releerlo/reescribirlo si ya está en cian). Esa misma
 * estructura la vacía/usa flushDracoPlanningModifiedCells, tanto al
 * guardar planificación como al arrancar cualquier refresco.
 */
// Diagnóstico temporal: escribe en EDIT_REPORT!W1 en qué paso se para
// handleDracoPlanningValueChanged, para poder ver qué pasa sin DevTools
// (basta con mirar esa celda después de editar). Se protege con
// DracoSuppressChangeEvents para que la propia escritura no se dispare
// a sí misma en bucle (EDIT_REPORT también tiene este mismo onChanged
// enganchado).
async function writeDracoPlanningDebug(msg) {
    try {
        DracoSuppressChangeEvents = true;
        await Excel.run(async (context) => {
            const editReportSheet = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
            editReportSheet.load("isNullObject");
            await context.sync();
            if (editReportSheet.isNullObject) return;
            editReportSheet.getRange("W1").values = [[draco_debugCell(new Date().toLocaleTimeString() + " | " + msg)]];
            await context.sync();
        });
    } catch (e) {
        console.error("[Draco] No se pudo escribir el diagnóstico en EDIT_REPORT!W1:", e);
    } finally {
        DracoSuppressChangeEvents = false;
    }
}

/**
 * Contexto de planificación de UN informe, resuelto una sola vez y
 * reutilizado para todas sus celdas modificadas — lo usan tanto el
 * chequeo de "Planificable" (buildDracoPlanningMeasureCheckers) como la
 * validación de "Dimensiones obligatorias" (validateDracoPlanningMandatoryDimensions)
 * y la tabla guardada (writeDracoPlanningSavedTable), para no repetir
 * tres veces la misma resolución de medida/cruce.
 *
 * Expone:
 *   - resolveMeasure(r, c) -> {measureName, zoneId} | null
 *     Medida aplicable a una celda del área de Draco_XXX_Values: FIJA
 *     (arrastrada a "Filtros") o VARIABLE por fila/columna (puesta como
 *     "MEASURE" en el diseño). zoneId es "filters"/"rows"/"columns" —
 *     hace falta para construir la clave de fieldOptions.
 *   - getCrossItems(addr) -> [{dim, attr, value}, ...]
 *     El "cruce" de esa celda (Estático primero, Dinámico si no hay
 *     fórmulas EPM_VALUE que leer).
 *   - isMeasurePlanificable(zoneId, measureName) -> boolean
 *   - dimensionsBehaviorForMeasure(zoneId, measureName) -> objeto
 *     {dim: {required, mode}, ...} — el de la medida si tiene
 *     "Personalizar" marcado (Opciones de campo), si no el del informe
 *     (Propiedades del informe).
 *   - measureLabel, filtersText: para la tabla guardada.
 */
// atributesGrid (opcional): grid de MODEL_ATRIBUTES del modelo de ESTE
// informe. Solo lo necesita getCrossItemsForWriteback (ver más abajo) para
// convertir dimensiones de tipo DATE al formato canónico; getCrossItems
// (la versión "de presentación", usada en diagnósticos/UI) no lo usa y
// sigue funcionando igual si se omite.
async function buildDracoPlanningReportContext(context, reportId, atributesGrid) {
    const editReportGrid = await getEditReportGrid(context, reportId);
    loadReportDefinition(editReportGrid, reportId); // rellena ReportState (Filters/FilterMeasureNames)

    const report = window.ReportStore ? window.ReportStore.getReport(reportId) : null;
    const fieldOptions = (report && report.design) ? (report.design.fieldOptions || {}) : {};
    const reportLevelBehavior = (report && report.reportProperties && report.reportProperties.dimensionsBehavior) || {};

    // Medida: FIJA (filtro) o VARIABLE (filas/columnas) — se resuelve una
    // sola vez por informe, no por celda.
    let fixedMeasure = null;
    let measureIAuxRows = null, measureIAuxCols = null, rowsRangeForMeasure = null, colsRangeForMeasure = null;

    if (ReportState.FilterMeasureNames && ReportState.FilterMeasureNames.length > 0) {
        fixedMeasure = { measureName: ReportState.FilterMeasureNames[0], zoneId: "filters" };
    } else {
        const fieldLevelsRows = buildAxisFieldLevelsTable(editReportGrid, "rows");
        const fieldLevelsCols = buildAxisFieldLevelsTable(editReportGrid, "columns");
        const names = dracoRangeNames(reportId);

        const rowsNameObj = context.workbook.names.getItemOrNullObject(names.rows);
        rowsRangeForMeasure = rowsNameObj.getRangeOrNullObject();
        rowsRangeForMeasure.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
        const colsNameObj = context.workbook.names.getItemOrNullObject(names.cols);
        colsRangeForMeasure = colsNameObj.getRangeOrNullObject();
        colsRangeForMeasure.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
        await context.sync();

        const findMeasureIAux = (fieldLevels) => {
            for (let iAux = 1; iAux < fieldLevels.length; iAux++) {
                const col = fieldLevels[iAux];
                if (col && col.byFlag[1] && String(col.byFlag[1].dim).toUpperCase() === "MEASURE") return iAux;
            }
            return null;
        };
        measureIAuxRows = findMeasureIAux(fieldLevelsRows);
        measureIAuxCols = findMeasureIAux(fieldLevelsCols);
    }

    const resolveMeasure = (r, c) => {
        if (fixedMeasure) return fixedMeasure;

        if (measureIAuxRows !== null && rowsRangeForMeasure && !rowsRangeForMeasure.isNullObject) {
            const localRow = r - rowsRangeForMeasure.rowIndex;
            const localCol = measureIAuxRows - 1;
            if (localRow >= 0 && localRow < rowsRangeForMeasure.rowCount && localCol >= 0 && localCol < rowsRangeForMeasure.columnCount) {
                const raw = rowsRangeForMeasure.values[localRow][localCol];
                if (raw !== "" && raw !== null && raw !== undefined) {
                    return { measureName: String(raw).replace(DRACO_GLYPH_PREFIX, ""), zoneId: "rows" };
                }
            }
        }
        if (measureIAuxCols !== null && colsRangeForMeasure && !colsRangeForMeasure.isNullObject) {
            const localCol = c - colsRangeForMeasure.columnIndex;
            const localRow = measureIAuxCols - 1;
            if (localRow >= 0 && localRow < colsRangeForMeasure.rowCount && localCol >= 0 && localCol < colsRangeForMeasure.columnCount) {
                const raw = colsRangeForMeasure.values[localRow][localCol];
                if (raw !== "" && raw !== null && raw !== undefined) {
                    return { measureName: String(raw).replace(DRACO_GLYPH_PREFIX, ""), zoneId: "columns" };
                }
            }
        }
        return null;
    };

    // Cruce: Estático primero (fórmulas EPM_VALUE); si no hay ninguna,
    // Dinámico (texto ya pintado).
    let crossCtx = null;
    let useDynamic = false;
    try {
        crossCtx = await loadDracoPlanningCrossContext(context, reportId);
        if (crossCtx.rowDefs.length === 0 && crossCtx.colDefs.length === 0) crossCtx = null;
    } catch (e) {
        crossCtx = null;
    }
    if (!crossCtx) {
        try {
            crossCtx = await loadDracoPlanningCrossContextDynamic(context, reportId);
            useDynamic = true;
        } catch (e) {
            crossCtx = null;
        }
    }

    const getCrossItems = (addr) => {
        if (!crossCtx) return [];
        return useDynamic
            ? dracoPlanningCrossItemsForCellDynamic(crossCtx, addr)
            : dracoPlanningCrossItemsForCell(crossCtx, addr);
    };

    // Versión "para escritura" de getCrossItems: en vez del valor pensado
    // para MOSTRAR (Display — p.ej. "01/08/2025" con el formato de fecha
    // del locale del usuario), usa el valor CANÓNICO que coincide con lo
    // que ya hay guardado en BigQuery (p.ej. "2025-08-01"), para que el
    // INSERT case contra la dimensión FECHA (o cualquier otra columna
    // DATE) igual que lo hacen las consultas normales del informe. Ver
    // dracoPlanningCrossItemsForCell / …Dynamic para el detalle de cada
    // caso (Estático/Dinámico).
    const getCrossItemsForWriteback = (addr) => {
        if (!crossCtx) return [];
        return useDynamic
            ? dracoPlanningCrossItemsForCellDynamic(crossCtx, addr, atributesGrid)
            : dracoPlanningCrossItemsForCell(crossCtx, addr, { preferCanonical: true });
    };

    const isMeasurePlanificable = (zoneId, measureName) => {
        if (!measureName) return false;
        const opts = fieldOptions[`${zoneId}|MEASURE|${measureName}`];
        return !!(opts && opts.planificable);
    };

    const dimensionsBehaviorForMeasure = (zoneId, measureName) => {
        if (zoneId && measureName) {
            const opts = fieldOptions[`${zoneId}|MEASURE|${measureName}`];
            if (opts && opts.dimensionsBehaviorCustom && opts.dimensionsBehavior) return opts.dimensionsBehavior;
        }
        return reportLevelBehavior;
    };

    return {
        resolveMeasure,
        getCrossItems,
        getCrossItemsForWriteback,
        isMeasurePlanificable,
        dimensionsBehaviorForMeasure,
        measureLabel: crossCtx ? crossCtx.measureLabel : "",
        filtersText: crossCtx ? crossCtx.filtersText : ""
    };
}

/**
 * Para cada informe candidato (ya filtrado a "de planificación" + esta
 * hoja), decide qué medida aplica a cada celda concreta y si esa medida
 * está marcada como "Planificable" — dato que vive POR INFORME (panel
 * "Opciones de campo" de "Editar informes", ver buildMeasureOptionsForm
 * en taskpane.js), NO en el modelo semántico: dos informes distintos con
 * la misma medida pueden tener cada uno su propio "Planificable".
 *
 * Devuelve Map<reportId, (r, c) => boolean>.
 */
async function buildDracoPlanningMeasureCheckers(context, reportIds) {
    const checkers = new Map();

    for (const reportId of reportIds) {
        try {
            const ctx = await buildDracoPlanningReportContext(context, reportId);
            checkers.set(reportId, (r, c) => {
                const m = ctx.resolveMeasure(r, c);
                if (!m) return false; // no se pudo determinar la medida: por seguridad, no marcar
                return ctx.isMeasurePlanificable(m.zoneId, m.measureName);
            });
        } catch (e) {
            console.warn(`[Draco] No se pudo resolver la medida del informe ${reportId} para comprobar "Planificable":`, e);
            checkers.set(reportId, () => false); // por seguridad, no marcar si falla la comprobación
        }
    }

    return checkers;
}

async function handleDracoPlanningValueChanged(eventArgs) {
    try {
        if (DracoSuppressChangeEvents) return;
        if (DracoSuppressPlanningPaintCount > 0) return;
        if (!eventArgs || !eventArgs.address) return;

        let addr = String(eventArgs.address);
        if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
        addr = addr.replace(/\$/g, "").toUpperCase();
        if (!addr) return;

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
            sheet.load("name");

            const range = sheet.getRange(addr);
            range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
            await context.sync();

            // Corta el bucle real que provoca el propio diagnóstico: escribir
            // en EDIT_REPORT!W1 (o Z1, el registro de tiempos de
            // draco_perfMark, ver endDracoPerfLog) dispara este mismo
            // onChanged (el aviso de DracoSuppressChangeEvents no llega a
            // tiempo, porque el evento de "cambió" puede notificarse después
            // de haberlo ya quitado). Cortar aquí, por dirección exacta, no
            // depende de ninguna temporización.
            if (sheet.name === "EDIT_REPORT" && (addr === "W1" || addr === "Z1")) return;

            // Límite de seguridad: un pegado/relleno realmente descomunal
            // no se recorre celda a celda. 20.000 da margen de sobra para
            // planificaciones grandes de verdad (p.ej. 500 líneas x 24
            // meses = 12.000 celdas en un único pegado) sin dejar de cortar
            // ante algo claramente fuera de lo normal (pegar el libro
            // entero, etc.).
            const totalCells = range.rowCount * range.columnCount;
            if (totalCells > 20000) {
                console.warn(`[Draco] Cambio de ${totalCells} celdas de golpe en ${sheet.name}!${addr}: se omite el marcado en cian (demasiadas celdas de una vez).`);
                await writeDracoPlanningDebug(`${sheet.name}!${addr}: cortado, ${totalCells} celdas de golpe (>20000)`);
                return;
            }

            // 1) Informes de planificación — filtrado en JS puro primero
            // (sin Excel), luego resuelto con getDracoResultSheetName (la
            // MISMA función que usa el resto del código, con su fallback a
            // EDIT_REPORT!D1) — confirmado con datos reales que
            // report.design.resultSheetName puede venir "undefined" (no
            // garantizado como se asumía antes), así que no basta con
            // mirarlo directamente.
            const planningReportIds = [];
            const debugAllReports = []; // para el diagnóstico si no hay candidatos
            if (window.ReportStore) {
                const store = window.ReportStore.getAllReports() || {};
                for (const idStr of Object.keys(store)) {
                    const report = store[idStr];
                    const isPlanning = !!(report && report.reportProperties && report.reportProperties.planningReport);
                    debugAllReports.push(`#${idStr}(plan=${isPlanning})`);
                    if (isPlanning) planningReportIds.push(Number(idStr));
                }
            }
            if (planningReportIds.length === 0) {
                await writeDracoPlanningDebug(`${sheet.name}!${addr}: 0 informes de planificación en todo el libro. Informes vistos: ${debugAllReports.join(", ") || "(ninguno)"}`);
                return; // ningún informe está marcado como "de planificación"
            }

            const candidateReportIds = [];
            const debugSheetNames = [];
            for (const id of planningReportIds) {
                const resultSheetName = await getDracoResultSheetName(context, id);
                debugSheetNames.push(`#${id}->"${resultSheetName}"`);
                if (resultSheetName === sheet.name) candidateReportIds.push(id);
            }
            if (candidateReportIds.length === 0) {
                await writeDracoPlanningDebug(`${sheet.name}!${addr}: hay informe(s) de planificación (${planningReportIds.join(",")}) pero ninguno resuelve a esta hoja. Hoja de cada uno: ${debugSheetNames.join(", ")}`);
                return; // esta hoja no tiene ningún informe de planificación
            }

            // 2) Rango Draco_<id>_Values de CADA candidato, por nombre
            // exacto — no la colección completa de nombres del libro.
            const namedLookups = candidateReportIds.map(id => {
                const nameObj = context.workbook.names.getItemOrNullObject(`Draco_${pad3(id)}_Values`);
                const r = nameObj.getRangeOrNullObject();
                r.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
                return { reportId: id, range: r };
            });
            await context.sync();

            const candidates = namedLookups
                .filter(n => !n.range.isNullObject)
                .map(n => ({
                    reportId: n.reportId,
                    rowIndex: n.range.rowIndex,
                    rowCount: n.range.rowCount,
                    columnIndex: n.range.columnIndex,
                    columnCount: n.range.columnCount
                }));
            if (candidates.length === 0) {
                await writeDracoPlanningDebug(`${sheet.name}!${addr}: informes candidatos ${candidateReportIds.join(",")} pero ninguno tiene rango con nombre Draco_XXX_Values (¿no se ha refrescado nunca?)`);
                return;
            }

            // 3) Intersección con la dirección modificada — en memoria.
            const hits = []; // {r, c, reportId}
            for (let dr = 0; dr < range.rowCount; dr++) {
                for (let dc = 0; dc < range.columnCount; dc++) {
                    const r = range.rowIndex + dr;
                    const c = range.columnIndex + dc;
                    const match = candidates.find(cand =>
                        r >= cand.rowIndex && r < cand.rowIndex + cand.rowCount &&
                        c >= cand.columnIndex && c < cand.columnIndex + cand.columnCount
                    );
                    if (match) hits.push({ r, c, reportId: match.reportId });
                }
            }
            if (hits.length === 0) {
                const rangos = candidates.map(c => `#${c.reportId}[fila ${c.rowIndex + 1}-${c.rowIndex + c.rowCount}, col ${c.columnIndex + 1}-${c.columnIndex + c.columnCount}]`).join(", ");
                await writeDracoPlanningDebug(`${sheet.name}!${addr} (fila ${range.rowIndex + 1}, col ${range.columnIndex + 1}) no cae dentro de ningún Draco_XXX_Values. Rangos existentes: ${rangos}`);
                return;
            }

            // 3.5) Además de que el INFORME sea de planificación, la MEDIDA
            // de cada celda tiene que estar marcada "Planificable" (tick
            // nuevo en MODEL_MEASURES) — puede ser una medida fija en la
            // zona "Filtros", o variar por fila/columna si está puesta en
            // el diseño de Filas/Columnas.
            const measureCheckers = await buildDracoPlanningMeasureCheckers(
                context,
                Array.from(new Set(hits.map(h => h.reportId)))
            );
            const hitsAntesDeMedida = hits.length;
            const hitsFiltrados = hits.filter(h => {
                const checker = measureCheckers.get(h.reportId);
                return checker ? checker(h.r, h.c) : false;
            });
            hits.length = 0;
            // Las celdas de cálculos entre miembros (Variación...) no son
            // datos de la base: nunca se planifican ni se escriben.
            hits.push(...hitsFiltrados.filter(h => !isDracoCalcCell(h.reportId, h.r, h.c)));
            if (hits.length === 0) {
                await writeDracoPlanningDebug(`${sheet.name}!${addr}: ${hitsAntesDeMedida} celda(s) en zona de planificación, pero ninguna con medida marcada "Planificable" en MODEL_MEASURES (o son celdas de cálculo).`);
                return;
            }

            // 4) Leer dirección + color ACTUAL solo de las celdas que han
            // caído dentro de un rango de planificación, todas de una vez.
            const cellRanges = hits.map(h => sheet.getRangeByIndexes(h.r, h.c, 1, 1));
            cellRanges.forEach(cr => cr.load(["address", "format/fill/color"]));
            await context.sync();

            cellRanges.forEach((cr, i) => {
                const reportId = hits[i].reportId;
                let cellAddr = String(cr.address);
                if (cellAddr.indexOf("!") !== -1) cellAddr = cellAddr.split("!").pop();
                cellAddr = cellAddr.replace(/\$/g, "").toUpperCase();

                if (!DracoPlanningModifiedCells.has(reportId)) {
                    DracoPlanningModifiedCells.set(reportId, new Map());
                }
                const bucket = DracoPlanningModifiedCells.get(reportId);
                if (!bucket.has(cellAddr)) {
                    bucket.set(cellAddr, {
                        sheetName: sheet.name,
                        address: cellAddr,
                        color: cr.format.fill.color || ""
                    });
                }
                cr.format.fill.color = "#DFFFFF"; // RGB(223,255,255)
            });
            await context.sync();

            console.log(`[Draco] ${hits.length} celda(s) de planificación marcada(s) en cian en ${sheet.name} (bloque ${addr}).`);
            await writeDracoPlanningDebug(`${sheet.name}!${addr}: OK, ${hits.length} celda(s) marcada(s) en cian (informe ${hits[0].reportId}).`);
        });
    } catch (e) {
        console.error("[Draco] Error marcando en cian las celdas de planificación modificadas:", e);
        await writeDracoPlanningDebug("EXCEPCIÓN: " + (e && e.message ? e.message : String(e)));
    }
}

const DracoRowsClickHandlerRegisteredSheets = new Set();
let DracoRowsClickOnAddedRegistered = false;

/**
 * Engancha, en TODAS las hojas del libro (los rangos Draco_XXX_Rows/Cols/
 * Values pueden estar en cualquiera) y en las que se añadan después:
 *   - handleDracoRowsSingleClick (onSingleClicked): clic en el icono +/-
 *     de jerarquía.
 *   - handleDracoPlanningValueChanged (onChanged): marca en cian una
 *     celda de Draco_XXX_Values editada a mano en un informe de
 *     planificación.
 * Se puede llamar varias veces sin problema (cada hoja solo se engancha
 * una vez, para los dos listeners a la vez).
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
                sheet.onChanged.add(handleDracoPlanningValueChanged);
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
                                sheet.onChanged.add(handleDracoPlanningValueChanged);
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
        console.log("[Draco] Listeners de clic (Draco_*_Rows) y cambio (Draco_*_Values) registrados en todas las hojas.");
    } catch (e) {
        console.warn("[Draco] No se pudieron registrar los listeners de clic/cambio en Draco_* (¿host sin soporte de ExcelApi 1.10?):", e);
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

/* ==========================================================================
 * BUSCADOR DE MIEMBROS (diálogo de Office) — versión v4
 * --------------------------------------------------------------------------
 * Problemas de la versión anterior que explican el "se queda en Cargando…"
 * y el "a veces sí, a veces no":
 *
 *  1) CARRERA en memberPicker.html: mandaba "ready" ANTES de registrar su
 *     handler de DialogParentMessageReceived (addHandlerAsync es asíncrono).
 *     Si la consulta ya estaba resuelta (dimensión pequeña o caché de
 *     BigQuery = las RÁPIDAS), el host contestaba al instante y el mensaje
 *     se perdía -> "Cargando…" para siempre. Las lentas sí funcionaban.
 *     Ahora el diálogo registra el handler primero y repite "ready" hasta
 *     recibir los datos; el host reenvía en cada "ready" hasta recibir
 *     "loaded" (acuse de recibo).
 *  2) Los items iban en UN solo messageChild: con dimensiones grandes el
 *     mensaje es enorme. Ahora van troceados (begin + chunks, con número de
 *     ronda para descartar restos de envíos anteriores).
 *  3) handleDracoPickerFlagRequest mantenía su Excel.run abierto TODO el
 *     rato que el diálogo estaba abierto, con un flag "busy" que IGNORABA
 *     los dobles clics que llegaban mientras tanto, y al terminar BORRABA
 *     T2:V2... borrando la petición del siguiente doble clic. Justo el caso
 *     "elijo valor y doble clic en la celda de al lado -> no pasa nada".
 *     Ahora T2:V2 se lee y se limpia AL MOMENTO (Excel.run corto), la
 *     petición va a una cola de un hueco ("gana la última") y, si hay un
 *     picker abierto, se cierra para abrir el nuevo.
 *  4) Reabrir un diálogo justo después de cerrar otro puede fallar con
 *     12007 ("ya hay un diálogo abierto"): antes se perdía en silencio.
 *     Ahora se espera un poco y se reintenta.
 *  5) Caché de valores por proveedor+dim+attr (5 min) y deduplicación de
 *     consultas en curso: doble clic en otra celda del mismo campo = al
 *     instante.
 * ========================================================================== */

let DracoCurrentPicker = null;          // { id, close(reason) } del diálogo abierto
let DracoPickerSeq = 0;
let DracoLastDialogClosedAt = 0;
const DracoPickerItemsCache = new Map();    // key -> { at, items }
const DracoPickerItemsInflight = new Map(); // key -> Promise<items>
const DRACO_PICKER_CACHE_TTL_MS = 5 * 60 * 1000;
const DRACO_PICKER_QUERY_TIMEOUT_MS = 120 * 1000;
const DRACO_PICKER_CHUNK_CHARS = 16000;

function dracoSleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

function dracoPickerCacheKey(located) {
    let prov = "";
    try { prov = (window.Provider && typeof Provider.key === "function") ? Provider.key() : ""; } catch (e) { /* nada */ }
    return prov + "|" + String(located.dim).toUpperCase() + "|" + String(located.attr).toUpperCase();
}

/** Vacía la caché de valores del picker (p.ej. tras cambiar de modelo/conexión). */
function clearDracoPickerCache() {
    DracoPickerItemsCache.clear();
    dracoLog("Caché del picker vaciada.");
}
window.clearDracoPickerCache = clearDracoPickerCache;

async function loadDracoPickerItems(located, tag) {
    const key = dracoPickerCacheKey(located);
    const cached = DracoPickerItemsCache.get(key);
    if (cached && Date.now() - cached.at < DRACO_PICKER_CACHE_TTL_MS) {
        dracoLog(tag, "valores desde caché (" + cached.items.length + " items) " + key);
        return cached.items;
    }

    let p = DracoPickerItemsInflight.get(key);
    if (p) {
        dracoLog(tag, "reutilizo consulta ya en curso para " + key);
    } else {
        p = (async () => {
            const t0 = Date.now();
            const sql = await window.ExcelService.buildFilterValuesSQL(located.dim, located.attr);
            dracoLog(tag, "SQL construida en " + (Date.now() - t0) + " ms" + (sql ? "" : " -> VACÍA (dim/attr no encontrados en MODEL_ATRIBUTES/MODEL_HIER)"));
            if (!sql) {
                throw new Error("No se ha encontrado " + located.dim + " / " + located.attr + " en el modelo (atributo o jerarquía).");
            }
            const t1 = Date.now();
            let json = await window.ExcelService.executeSQL(sql);
            if (typeof json !== "string") json = JSON.stringify(json);
            const items = parseMemberJsonTree(json);
            dracoLog(tag, "consulta OK en " + (Date.now() - t1) + " ms | respuesta " + json.length + " chars | " + items.length + " items");
            DracoPickerItemsCache.set(key, { at: Date.now(), items });
            return items;
        })();
        DracoPickerItemsInflight.set(key, p);
        p.then(() => DracoPickerItemsInflight.delete(key), () => DracoPickerItemsInflight.delete(key));
    }

    let timer;
    const timeout = new Promise((_, rej) => {
        timer = setTimeout(() => rej(new Error("La consulta de valores ha tardado más de " + (DRACO_PICKER_QUERY_TIMEOUT_MS / 1000) + " s.")), DRACO_PICKER_QUERY_TIMEOUT_MS);
    });
    try {
        return await Promise.race([p, timeout]);
    } finally {
        clearTimeout(timer);
    }
}

function displayDracoDialogOnce(url, options) {
    return new Promise((resolve) => {
        try {
            Office.context.ui.displayDialogAsync(url, options, (res) => resolve(res));
        } catch (e) {
            resolve({ status: Office.AsyncResultStatus.Failed, error: { code: "EXC", message: dracoErrMsg(e) } });
        }
    });
}

/**
 * Abre el buscador de miembros y devuelve una promesa que se resuelve
 * cuando el diálogo se ha cerrado DEL TODO (elegido/cancelado/X/sustituido
 * por otra petición), con la fórmula ya escrita si se eligió un valor.
 */
async function openMemberRecognitionPicker(addr, located, initialSearch) {
    const id = ++DracoPickerSeq;
    const tag = "[picker#" + id + "]";
    const t0 = Date.now();
    dracoLog(tag, "abrir para " + addr + " -> " + located.dim + "." + located.attr +
        " (informe " + located.reportId + (located.growth ? ", zona crecimiento" : "") + ") texto='" + (initialSearch || "") + "'");

    // Si había otro abierto, se cierra (gana la última petición).
    if (DracoCurrentPicker) {
        dracoLog(tag, "había otro picker abierto (#" + DracoCurrentPicker.id + "): se cierra");
        await DracoCurrentPicker.close("sustituido por picker#" + id);
    }

    let resolveDone;
    const done = new Promise((r) => { resolveDone = r; });
    let finished = false;
    let dialogRef = null;
    let dialogReadySeen = false;
    let delivered = false;
    let round = 0;
    let itemsPayload = null;
    let itemsError = null;
    let readyWatchdog = null;

    const finish = async (reason, after) => {
        if (finished) return done;
        finished = true;
        clearTimeout(readyWatchdog);
        if (dialogRef) {
            try { dialogRef.close(); } catch (e) { /* ya cerrado */ }
            DracoLastDialogClosedAt = Date.now();
        }
        if (DracoCurrentPicker && DracoCurrentPicker.id === id) DracoCurrentPicker = null;
        if (after) {
            try { await after(); } catch (e) { dracoLog(tag, "ERROR tras cerrar: " + dracoErrMsg(e)); }
        }
        DracoMemberPickerOpen = !!DracoCurrentPicker;
        dracoLog(tag, "cerrado: " + reason + " (" + (Date.now() - t0) + " ms en total)");
        resolveDone();
        return done;
    };

    DracoCurrentPicker = { id, close: (reason) => finish(reason) };
    DracoMemberPickerOpen = true;

    let lastSendAt = 0;
    let loggedWaiting = false;
    const sendIfReady = (trigger) => {
        if (finished || !dialogRef || !dialogReadySeen || delivered) return;
        // Los "ready" repetidos no reenvían si se acaba de mandar (el
        // diálogo puede estar todavía montando la lista); "resend" sí.
        if (trigger === "ready" && round > 0 && Date.now() - lastSendAt < 4000) return;
        if (!itemsPayload && !itemsError) {
            if (!loggedWaiting) dracoLog(tag, "diálogo listo, la consulta aún no ha terminado");
            loggedWaiting = true;
            return;
        }
        round++;
        lastSendAt = Date.now();
        try {
            if (itemsError) {
                dialogRef.messageChild(JSON.stringify({ type: "error", round, message: itemsError }));
                dracoLog(tag, "error enviado al diálogo (ronda " + round + ", " + trigger + ")");
                return;
            }
            const body = JSON.stringify(itemsPayload.items);
            const parts = [];
            for (let i = 0; i < body.length; i += DRACO_PICKER_CHUNK_CHARS) parts.push(body.slice(i, i + DRACO_PICKER_CHUNK_CHARS));
            dialogRef.messageChild(JSON.stringify({
                type: "begin", round, total: parts.length,
                initialSearch: itemsPayload.initialSearch, title: itemsPayload.title
            }));
            parts.forEach((part, seq) => dialogRef.messageChild(JSON.stringify({ type: "chunk", round, seq, part })));
            dracoLog(tag, "enviados " + itemsPayload.items.length + " items en " + parts.length +
                " trozos (" + body.length + " chars), ronda " + round + " (" + trigger + ")");
        } catch (e) {
            dracoLog(tag, "ERROR en messageChild: " + dracoErrMsg(e));
        }
    };

    // Consulta en paralelo con la apertura del diálogo.
    loadDracoPickerItems(located, tag).then(
        (items) => {
            itemsPayload = { items, initialSearch: initialSearch || "", title: located.dim + " / " + located.attr };
            sendIfReady("consulta terminada");
        },
        (err) => {
            itemsError = dracoErrMsg(err);
            dracoLog(tag, "ERROR en la consulta: " + itemsError);
            sendIfReady("consulta con error");
        }
    );

    // Deja respirar a Office si se acaba de cerrar otro diálogo (evita 12007).
    const sinceClose = Date.now() - DracoLastDialogClosedAt;
    if (sinceClose < 400) await dracoSleep(400 - sinceClose);
    if (finished) return done;

    const dialogUrl = new URL("memberPicker.html", window.location.href).href +
        "?v=" + encodeURIComponent(DRACO_BUILD) + "&pid=" + id;

    let res = null;
    for (let attempt = 1; attempt <= 6; attempt++) {
        res = await displayDracoDialogOnce(dialogUrl, { height: 55, width: 28, displayInIframe: false });
        if (res.status !== Office.AsyncResultStatus.Failed) break;
        const code = res.error && res.error.code;
        dracoLog(tag, "displayDialogAsync falló (intento " + attempt + "): " + code + " " + (res.error && res.error.message));
        if (code !== 12007 || finished) break; // 12007 = ya hay un diálogo abierto -> reintentar
        await dracoSleep(250 * attempt);
    }
    if (!res || res.status === Office.AsyncResultStatus.Failed) {
        return finish("no se pudo abrir el diálogo");
    }

    dialogRef = res.value;
    if (finished) { // sustituido mientras se abría
        try { dialogRef.close(); } catch (e) { /* nada */ }
        DracoLastDialogClosedAt = Date.now();
        return done;
    }
    dracoLog(tag, "diálogo abierto en " + (Date.now() - t0) + " ms");

    readyWatchdog = setTimeout(() => {
        if (!dialogReadySeen && !finished) {
            dracoLog(tag, "AVISO: 15 s sin 'ready' del diálogo. ¿memberPicker.html antiguo en caché o office.js sin cargar?");
        }
    }, 15000);

    dialogRef.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        let payload;
        try { payload = JSON.parse(arg.message); } catch (e) {
            dracoLog(tag, "mensaje no JSON del diálogo: " + String(arg.message).slice(0, 200));
            return;
        }
        if (!payload) return;

        switch (payload.type) {
            case "log":
                dracoLog(tag, "[diálogo] " + payload.msg);
                return;
            case "ready":
                if (!dialogReadySeen) dracoLog(tag, "'ready' del diálogo (build diálogo=" + (payload.build || "?") + ") a los " + (Date.now() - t0) + " ms");
                dialogReadySeen = true;
                sendIfReady("ready");
                return;
            case "resend":
                dracoLog(tag, "el diálogo pide reenvío (ronda " + payload.round + " incompleta)");
                delivered = false;
                sendIfReady("resend");
                return;
            case "loaded":
                if (!delivered) dracoLog(tag, "el diálogo confirma datos pintados (" + payload.count + " items) a los " + (Date.now() - t0) + " ms");
                delivered = true;
                return;
            case "select":
                dracoLog(tag, "valor elegido: " + payload.attribute + " = " + payload.value);
                finish("valor elegido", async () => {
                    await Excel.run(async (context) => {
                        if (located.growth) {
                            const g = located.growth;
                            dracoLog(tag, "ampliando " + g.rangeName + " para incluir " + addr);
                            await growDracoNamedRange(
                                context, g.sheetName, g.rangeName, g.axisSuffix,
                                g.axisStart, g.requiredCount, g.crossIndex, g.crossCount
                            );
                        }
                        const resultSheetName = await getDracoResultSheetName(context, located.reportId);
                        const sheet = context.workbook.worksheets.getItem(resultSheetName);
                        sheet.getRange(addr).values = [[buildEpmValueFormula(located.dim, payload.attribute, payload.value)]];
                        await context.sync();
                    });
                    dracoLog(tag, "EPM_VALUE escrita en " + addr);
                });
                return;
            case "cancel":
                finish("cancelado por el usuario");
                return;
            default:
                dracoLog(tag, "mensaje desconocido del diálogo: " + payload.type);
        }
    });

    dialogRef.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        // 12006 = cerrado con la X; 12002/12003 = no se pudo cargar la página.
        finish("evento del diálogo " + (arg && arg.error));
    });

    return done;
}



// NOTA: aquí existía handleEditReportMemberPickerRequest, el mecanismo
// antiguo del Member Picker que leía EDIT_REPORT!A1:A5 (flag "X" en A5).
// Ya estaba desactivado (sin enganchar a ningún onChanged) y ha sido
// eliminado. El único punto de entrada del buscador de miembros es ahora
// EDIT_REPORT!T2:V2 (ver handleDracoPickerFlagRequest).



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
        const linkedIds = dracoLinkedOf(reportId);
        if (linkedIds.length) await dracoClearLinkedAreas(context, linkedIds);   // antes: el principal puede crecer
        await jsonTo3Matrices(context, lastJson, reportId);
        if (linkedIds.length) {
            await dracoPaintLinkedChain(context, reportId, linkedIds,
                linkedIds.map(id => DracoLastJsonByReport.get(dracoStateKey(id))));
        }
    }
    return true;
}

// NOTA: aquí existían las celdas de control EDIT_REPORT!T1:V1 y el
// handler handleDracoEditReportExpandCollapseRequest (onChanged), que
// permitían pedir un expandir/contraer sin pasar por clic directo sobre
// el icono ▾/▸. Se han eliminado por completo. El expandir/contraer del
// icono de jerarquía sigue funcionando igual que antes por CLIC SIMPLE
// (ver handleDracoRowsSingleClick -> toggleDracoCollapseAtCell), que es
// la función que se conserva.

// Celdas de control de EDIT_REPORT para pedir, desde el VBA del XLAM, la
// apertura del buscador de miembros — tanto por reconocimiento al
// teclear como por doble clic real cancelado (Workbook_SheetBeforeDoubleClick
// + Cancel = True, ya que Office.js no reenvía ese segundo clic). Sustituye
// al antiguo picker directo de A5 (ya eliminado): ahora hay un
// ÚNICO punto de entrada (T2:V2) con dos valores posibles en V2.
const DRACO_PICKER_SHEET_CELL = "T2"; // hoja donde ocurrió (Sh.Name en VBA)
const DRACO_PICKER_TARGET_CELL = "U2"; // celda afectada (Target.Address)
const DRACO_PICKER_FLAG_CELL = "V2"; // "REC" = reconocimiento de miembros (tecleado) | "DC" = doble clic

/**
 * Núcleo de la acción de "doble clic real" (fuera del icono +/- de
 * expandir/contraer): dado el `located` que ya ha resuelto
 * findDracoRowsNamedRangeForCell (informe/eje/nivel realmente afectado,
 * exacto o en zona de crecimiento), resuelve el campo (dim/attr) y abre
 * el buscador de miembros. Extraído tal cual de handleDracoRowsSingleClick
 * (misma lógica, sin ningún cambio de comportamiento) para poder
 * invocarlo TAMBIÉN desde runDracoSimulatedDoubleClick, que dispara este
 * mismo camino cuando el VBA del XLAM cancela un doble clic nativo y lo
 * pide en su lugar rellenando EDIT_REPORT!T2:V2 (ver
 * handleDracoPickerFlagRequest más abajo).
 *
 * Devuelve { fieldLocated, accionTexto } para que quien llame pueda
 * trazarlo igual que antes.
 */
async function runDracoDoubleClickAction(context, addr, located, indentLevel, cellText) {
    let accionTexto;
    let fieldLocated = null;
    let pickerTarget = null;

    fieldLocated = await resolveDracoFieldForAxisLevel(
        context, located.reportId, located.axis, located.level, indentLevel
    );
    if (fieldLocated) {
        // Se adjunta el reportId realmente clicado (no el "activo" del
        // taskpane) para que openMemberRecognitionPicker escriba en la
        // hoja de resultados correcta, y —si el clic cayó en la zona de
        // crecimiento, fuera todavía del rango con nombre real—
        // `growth`, para ampliar de verdad ese rango con nombre en
        // cuanto el usuario confirme un valor en el picker.
        fieldLocated.reportId = located.reportId;
        fieldLocated.growth = located.growth || null;
        accionTexto = located.exact
            ? " | Accion: buscador de miembros abierto (doble clic)"
            : " | Accion: buscador de miembros abierto (doble clic, ampliará el rango si se confirma un valor)";
        // v4: ya NO se abre aquí (con el Excel.run abierto); se devuelve
        // el destino y lo abre runDracoPickerLoop fuera de Excel.run.
        pickerTarget = { addr, fieldLocated, initialSearch: cellText };
    } else {
        accionTexto = " | Accion: doble clic fuera de Filas/Columnas del informe";
    }

    return { fieldLocated, accionTexto, pickerTarget };
}

/**
 * Ejecuta el reconocimiento de miembros (misma lógica que tenía, antes de
 * eliminarse, la función legada handleDracoMemberRecognitionChanged al
 * detectar un valor tecleado en una celda de Draco_XXX_Rows/Cols) sobre
 * sheetName!targetAddr. Ya no hay
 * ningún onChanged enganchado directamente a las hojas de resultados
 * para esto: la detección de "se ha tecleado algo" la hace el XLAM en
 * VBA, que rellena EDIT_REPORT!T2:V2 con V2="REC", igual que ya hacía
 * para el doble clic con V2="DC" (ver handleDracoPickerFlagRequest, que
 * llama a esta función).
 */
async function runDracoMemberRecognitionAction(context, sheetName, targetAddr, tag) {
    tag = tag || "[REC]";
    let addr = String(targetAddr || "");
    if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
    addr = addr.replace(/\$/g, "").toUpperCase();
    if (!addr) {
        dracoLog(tag, "REC: sin celda (EDIT_REPORT!U2 vacío).");
        return;
    }

    const recognitionActive = await isDracoMemberRecognitionActive(context);
    if (!recognitionActive) {
        dracoLog(tag, "REC: EDIT_REPORT!B1 desactivado, se ignora la petición para", sheetName, addr);
        return;
    }

    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load(["isNullObject", "id"]);
    await context.sync();
    if (sheet.isNullObject) {
        dracoLog(tag, "REC: la hoja indicada en EDIT_REPORT!T2 no existe:", sheetName);
        return;
    }

    const cell = sheet.getRange(addr);
    cell.load(["rowCount", "columnCount", "values", "formulas"]);
    cell.format.load("indentLevel");
    await context.sync();

    // Varias celdas a la vez (pegado/relleno): se ignora.
    if (cell.rowCount !== 1 || cell.columnCount !== 1) {
        dracoLog(tag, "REC: varias celdas a la vez (pegado/relleno), se ignora:", sheetName, addr);
        return;
    }

    const value = cell.values[0][0];
    const formula = cell.formulas[0][0];
    if (value === "" || value === null || value === undefined) {
        dracoLog(tag, "REC: valor vacío, se ignora:", sheetName, addr);
        return;
    }
    // Ya es EPM_VALUE (tecleada, pegada, o escrita por nosotros mismos
    // hace un instante): salir para no reabrir el buscador en bucle.
    if (typeof formula === "string" && /^=\s*EPM_VALUE\s*\(/i.test(formula)) {
        dracoLog(tag, "REC: la celda ya es una fórmula EPM_VALUE, se ignora:", sheetName, addr);
        return;
    }
    const currentText = String(value);

    const located = await findDracoRowsNamedRangeForCell(context, sheet.id, addr);
    if (!located) {
        dracoLog(tag, "REC: fuera de cualquier Draco_XXX_Rows/Cols:", sheetName, addr);
        return;
    }

    const fieldLocated = await resolveDracoFieldForAxisLevel(
        context, located.reportId, located.axis, located.level, cell.format.indentLevel
    );
    if (!fieldLocated) {
        dracoLog(tag, `REC: informe ${located.reportId}, eje ${located.axis}, nivel ${located.level} sin dim/attr configurado, se ignora.`);
        return;
    }

    // Igual que en el doble clic: se adjunta el reportId realmente
    // afectado y, si la celda cae en zona de crecimiento (fuera todavía
    // del rango con nombre real), la info para ampliarlo de verdad en
    // cuanto se confirme un valor en el picker.
    fieldLocated.reportId = located.reportId;
    fieldLocated.growth = located.growth || null;

    dracoLog(tag, `REC: abriendo picker (${fieldLocated.dim} / ${fieldLocated.attr}) para ${sheetName}!${addr}...`);
    return { addr, fieldLocated, initialSearch: currentText };
}

/**
 * Ejecuta el equivalente de un doble clic "real" (fuera del icono +/-)
 * sobre sheetName!targetAddr, llamando a la MISMA lógica que usa el clic
 * físico (findDracoRowsNamedRangeForCell + runDracoDoubleClickAction). Si
 * la celda no pertenece (ni podría ampliar) ningún Draco_XXX_Rows/Cols,
 * simplemente no hace nada (igual que hacía antes el "FUERA DE RANGO" del
 * doble clic físico) — el VBA solo necesita acertar de forma aproximada,
 * el filtrado fino se hace aquí.
 */
async function runDracoSimulatedDoubleClick(context, sheetName, targetAddr, tag) {
    tag = tag || "[DC]";
    let addr = String(targetAddr || "");
    if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
    addr = addr.replace(/\$/g, "").toUpperCase();
    if (!addr) {
        dracoLog(tag, "DC sin celda (EDIT_REPORT!U2 vacío).");
        return;
    }

    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load(["isNullObject", "id"]);
    await context.sync();
    if (sheet.isNullObject) {
        dracoLog(tag, "DC: la hoja indicada en EDIT_REPORT!T2 no existe:", sheetName);
        return;
    }

    const located = await findDracoRowsNamedRangeForCell(context, sheet.id, addr);
    if (!located) {
        dracoLog(tag, "DC fuera de cualquier Draco_*_Rows/Cols (ni zona de crecimiento): " + sheetName + "!" + addr);
        return;
    }

    const cell = sheet.getRange(addr);
    cell.load(["values", "format/indentLevel"]);
    await context.sync();
    const cellText = String((cell.values && cell.values[0] && cell.values[0][0]) || "");

    const result = await runDracoDoubleClickAction(context, addr, located, cell.format.indentLevel, cellText);
    dracoLog(tag, `DC en ${sheetName}!${addr} (${located.rangeName}, eje ${located.axis}, nivel ${located.level}, sangría ${cell.format.indentLevel})` +
        (result.fieldLocated ? ` -> ${result.fieldLocated.dim}.${result.fieldLocated.attr}` : " -> SIN dim/attr para ese nivel (resolveDracoFieldForAxisLevel = null)"));
    return result.pickerTarget;
}

/**
 * onChanged de EDIT_REPORT. El VBA del XLAM escribe en T2 (hoja), U2
 * (celda) y V2 ("DC" doble clic | "REC" reconocimiento al teclear).
 *
 * v4: se lee y se LIMPIA T2:V2 al momento (Excel.run corto, serializado),
 * se deja la petición en una cola de un solo hueco (gana la última) y se
 * procesa en runDracoPickerLoop SIN mantener ningún Excel.run abierto
 * mientras el diálogo está en pantalla. Así un doble clic que llega con
 * un picker abierto ya no se pierde: cierra ese picker y abre el nuevo.
 */
let DracoPickerPending = null;
let DracoPickerLoopRunning = false;
let DracoPickerReadChain = Promise.resolve();
let DracoPickerReqSeq = 0;

async function handleDracoPickerFlagRequest(eventArgs) {
    if (!eventArgs || !eventArgs.address) return;

    let addr = String(eventArgs.address);
    if (addr.indexOf("!") !== -1) addr = addr.split("!").pop();
    if (!addr) return;

    const touchedRange = parseAddressRange(addr);
    const controlCells = [DRACO_PICKER_SHEET_CELL, DRACO_PICKER_TARGET_CELL, DRACO_PICKER_FLAG_CELL]
        .map(parseAddress);
    const touchesControlZone = controlCells.some(p =>
        p.row >= touchedRange.r1 && p.row <= touchedRange.r2 &&
        p.col >= touchedRange.c1 && p.col <= touchedRange.c2
    );
    if (!touchesControlZone) return;

    // Se anota AHORA si estábamos pintando (para descartar REC provocados
    // por nuestras propias escrituras), antes de encolar la lectura.
    const suppressedAtEvent = DracoSuppressChangeEvents;

    DracoPickerReadChain = DracoPickerReadChain
        .then(() => readAndClearDracoPickerRequest(addr, suppressedAtEvent))
        .catch((e) => dracoLog("ERROR leyendo/limpiando EDIT_REPORT!T2:V2: " + dracoErrMsg(e)));
    return DracoPickerReadChain;
}

async function readAndClearDracoPickerRequest(eventAddr, suppressedAtEvent) {
    let req = null;
    await Excel.run(async (context) => {
        const ctrl = context.workbook.worksheets.getItem("EDIT_REPORT")
            .getRange(DRACO_PICKER_SHEET_CELL + ":" + DRACO_PICKER_FLAG_CELL);
        ctrl.load("values");
        await context.sync();

        const sheetName = String(ctrl.values[0][0] || "").trim();
        const targetAddr = String(ctrl.values[0][1] || "").trim();
        const flag = String(ctrl.values[0][2] || "").trim().toUpperCase();

        if (!sheetName && !targetAddr && !flag) return; // vacío: nuestra propia limpieza (o ya procesado)

        if (!sheetName || !targetAddr || (flag !== "DC" && flag !== "REC")) {
            dracoLog("T2:V2 incompleto o flag desconocido (evento " + eventAddr + "), NO se toca: [" +
                sheetName + "] [" + targetAddr + "] [" + flag + "]");
            return;
        }

        ctrl.clear(Excel.ClearApplyTo.contents);
        await context.sync();
        req = { id: ++DracoPickerReqSeq, sheetName, targetAddr, flag };
    });
    if (!req) return;

    const tag = "[req#" + req.id + "]";
    if (req.flag === "REC" && suppressedAtEvent) {
        dracoLog(tag, "REC " + req.sheetName + "!" + req.targetAddr + " ignorado (lo provocó un repintado del add-in)");
        return;
    }

    dracoLog(tag, "recibido " + req.flag + " " + req.sheetName + "!" + req.targetAddr);
    if (DracoPickerPending) dracoLog("[req#" + DracoPickerPending.id + "] descartado: sustituido por req#" + req.id);
    DracoPickerPending = req;

    if (DracoCurrentPicker) DracoCurrentPicker.close("nuevo " + req.flag + " (req#" + req.id + ")");
    runDracoPickerLoop();
}

async function runDracoPickerLoop() {
    if (DracoPickerLoopRunning) return;
    DracoPickerLoopRunning = true;
    try {
        while (DracoPickerPending) {
            const req = DracoPickerPending;
            DracoPickerPending = null;
            const tag = "[req#" + req.id + "]";
            try {
                const t0 = Date.now();
                let target = null;
                await Excel.run(async (context) => {
                    target = req.flag === "DC"
                        ? await runDracoSimulatedDoubleClick(context, req.sheetName, req.targetAddr, tag)
                        : await runDracoMemberRecognitionAction(context, req.sheetName, req.targetAddr, tag);
                });
                dracoLog(tag, "resuelto en " + (Date.now() - t0) + " ms" + (target ? "" : " -> no se abre picker"));
                if (!target) continue;
                if (DracoPickerPending) {
                    dracoLog(tag, "no se abre: ya ha llegado una petición más nueva");
                    continue;
                }
                await openMemberRecognitionPicker(target.addr, target.fieldLocated, target.initialSearch);
            } catch (e) {
                dracoLog(tag, "ERROR procesando la petición: " + dracoErrMsg(e));
            }
        }
    } finally {
        DracoPickerLoopRunning = false;
    }
}

/**
 * Registra (una sola vez) los listeners de EDIT_REPORT que no dependen de
 * que exista ninguna hoja de resultados: la petición de expandir/contraer
 * de T1:V1 y el buscador de miembros (T2:V2, REC/DC). Es INDEPENDIENTE de
 * registerDracoSelectionHandler/DracoHandlerRegistered a propósito: antes
 * el picker estaba dentro de esa misma función y compartía su flag, lo
 * que significaba que si CSV_RESULT (la hoja de resultados) todavía no
 * existía —p.ej. sesión recién abierta, sin haber pulsado nunca
 * "Actualizar"— el picker NO se registraba, aunque EDIT_REPORT sí
 * existiera. Con esta función aparte, ensureDracoHandlersRegistered puede
 * engancharla sin depender de que exista CSV_RESULT.
 */
let DracoEditReportHandlerRegistering = false;
async function registerEditReportPickerHandler(context) {
    // v4: el flag se reserva ANTES del primer await para que dos llamadas
    // casi simultáneas no registren el handler dos veces.
    if (DracoEditReportHandlerRegistered || DracoEditReportHandlerRegistering) return;
    DracoEditReportHandlerRegistering = true;
    try {
        await registerEditReportPickerHandlerInner(context);
    } finally {
        DracoEditReportHandlerRegistering = false;
    }
}

async function registerEditReportPickerHandlerInner(context) {

    const editReport = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
    editReport.load("isNullObject");
    await context.sync();

    if (editReport.isNullObject) {
        console.log("[Draco] registerEditReportPickerHandler: EDIT_REPORT no existe todavía.");
        return;
    }

    // Único punto que abre el buscador de miembros: reacciona a
    // EDIT_REPORT!V2 pasando a "REC" (reconocimiento al escribir) o "DC"
    // (doble clic), ambos rellenados por el XLAM en VBA en T2 (hoja) /
    // U2 (celda) / V2 (REC|DC).
    editReport.onChanged.add(handleDracoPickerFlagRequest);
    await context.sync();

    DracoEditReportHandlerRegistered = true;
    dracoLog("Listener de EDIT_REPORT!T2:V2 (DC/REC) registrado.");
}

// NOTA: aquí existía handleDracoPlanningValueChanged (onChanged de cada
// hoja de resultados), que marcaba en cian las celdas de
// Draco_<id>_Values editadas a mano en informes de planificación. Se ha
// eliminado por completo, junto con su enganche en
// registerDracoSelectionHandler. El resto de la infraestructura de
// planificación (DracoPlanningModifiedCells, flushDracoPlanningModifiedCells,
// beginSuppressPlanningPaint, guardarPlanificacion) se conserva intacta
// porque la usan otros flujos (guardar planificación / refresco), pero ya
// no se alimenta automáticamente: ninguna celda se marcará en cian al
// editarla a mano.

// sheetName es el nombre de la hoja de resultados donde vive `sheet`: cada
// informe puede pintarse en una hoja distinta (ver resultSheetNameFromGrid/
// getDracoResultSheetName), así que los listeners de clic (+/-) y
// reconocimiento de miembros se enganchan POR HOJA, no una sola vez para
// todo el libro — si no, un segundo informe en otra hoja se quedaba sin
// clic interactivo (nunca se registraba nada ahí).
async function registerDracoSelectionHandler(context, sheet, sheetName) {
    if (DracoHandlerRegisteredSheets.has(sheetName)) return;

    // El marcado en cian de planificación (handleDracoPlanningValueChanged)
    // NO se engancha aquí: vive en ensureDracoRowsClickLoggerRegistered,
    // registrado una vez para TODAS las hojas del libro (no solo la de
    // resultados de este informe) desde el arranque, junto al clic de
    // jerarquía. La detección de reconocimiento/doble clic la sigue
    // haciendo el XLAM en VBA, avisando por EDIT_REPORT!T2:V2 (V2="REC"
    // o "DC"); el único sitio que abre el picker es
    // handleDracoPickerFlagRequest, enganchado una vez a EDIT_REPORT
    // (ver registerEditReportPickerHandler), no aquí por cada hoja.

    // Petición de apertura del Member Picker desde EDIT_REPORT
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
function buildDracoAxisLevels(editReportGrid, dimCol, attrCol, hierCol, jerCol) {
    const levels = [];
    let R = 15;
    while (String(cellValue(editReportGrid, R, dimCol)).trim() !== "") {
        const dim = String(cellValue(editReportGrid, R, dimCol)).trim();
        levels.push({
            isMeasure: dim.toUpperCase() === "MEASURE",
            dim,
            attr: String(cellValue(editReportGrid, R, attrCol)).trim(),
            flag: Number(cellValue(editReportGrid, R, hierCol)),
            // Nombre de la JERARQUÍA (col K=11 en Filas / Q=17 en Columnas,
            // ver ReportStore.saveDesign/getReportGrid), NO el atributo del
            // nivel: todos los niveles de una misma jerarquía comparten este
            // mismo nombre (p.ej. "H_CUENTA"), a diferencia de `attr`
            // (NIVEL1/NIVEL2...), que sí cambia por nivel. Cadena vacía para
            // campos que no son jerarquía, o si no se pide esta columna.
            jerarquia: jerCol ? String(cellValue(editReportGrid, R, jerCol)).trim() : "",
            // Jerarquía de TEXTOS de un "ID y texto" (ver DRACO_HIER_TEXT_SUFFIX)
            isTextCompanion: jerCol ? dracoIsHierTextName(cellValue(editReportGrid, R, jerCol)) : false,
            // "Solo texto": columna de ID que se consulta pero no se pinta
            hidden: dracoIsHiddenMarker(cellValue(editReportGrid, R, hierCol + 3))
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
    const companionReal = []; // companionReal[i] = true si la posición i es de la jerarquía de textos
    const flag1OrdinalToIaux = [];
    // Ordinales de nivel 1 ocultos ("Solo texto") pendientes de apuntar a la
    // columna del texto que los sustituye (la siguiente que se pinte).
    let pendingHiddenOrdinals = [];
    // "Solo texto" + eje Estático: la fórmula EPM_VALUE de la columna de TEXTO
    // lleva el ID (atributo y valor) y el texto solo como descripción;
    // idPartnerReal[i] = posición del ID oculto al que sustituye el texto i.
    const idPartnerReal = [];
    const hiddenDims = []; // { idx, dim, flag, claimed }
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
        if (level.hidden) {
            hiddenDims.push({ idx: dimIdx, dim: level.dim, flag: level.flag, claimed: false });
            // Solo se consulta: sin columna física (iauxReal = null => no se pinta).
            // Su +/- (si es nodo de jerarquía) se pinta en la columna del texto.
            if (flag === 1) {
                ordinal++;
                pendingHiddenOrdinals.push(ordinal);
            }
            flagsReal[dimIdx] = flag;
            companionReal[dimIdx] = !!level.isTextCompanion;
            fieldsReal[dimIdx] = { dim: level.dim, attr: level.attr };
            iauxReal[dimIdx] = null;
            continue;
        }
        if (flag === 1) {
            iAux++;
            ordinal++;
            flag1OrdinalToIaux[ordinal] = iAux;
            pendingHiddenOrdinals.forEach(o => { flag1OrdinalToIaux[o] = iAux; });
            pendingHiddenOrdinals = [];
        }
        flagsReal[dimIdx] = flag;
        companionReal[dimIdx] = !!level.isTextCompanion;
        fieldsReal[dimIdx] = { dim: level.dim, attr: level.attr };
        iauxReal[dimIdx] = iAux;

        // ¿Es el texto que sustituye a un ID oculto?
        let partner = null;
        if (level.isTextCompanion) {
            for (let k = hiddenDims.length - 1; k >= 0; k--) {
                const h = hiddenDims[k];
                if (!h.claimed && h.dim === level.dim && h.flag === flag) { partner = h; break; }
            }
        } else {
            const prev = hiddenDims[hiddenDims.length - 1];
            if (prev && !prev.claimed && prev.idx === dimIdx - 1 && prev.dim === level.dim) partner = prev;
        }
        if (partner) {
            partner.claimed = true;
            idPartnerReal[dimIdx] = partner.idx;
        }
    }

    return { flagsReal, companionReal, idPartnerReal, fieldsReal, iauxReal, flag1OrdinalToIaux, measureLevels, totalIaux: iAux, dimCount: dimIdx };
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

// jsonTo3MatricesCore (y sus funciones auxiliares: clearDracoNamedRanges,
// writeCellBlock, writeIndentAndColorRuns, applyDracoNamedRanges,
// applyDracoTotalHighlight, el autofit de columnas) borran y escriben en
// varias tandas — de ahí el parpadeo "primero borra, luego va
// escribiendo" que reporta el usuario. suspendScreenUpdatingUntilNextSync()
// solo dura hasta el PRÓXIMO sync, así que hay que rearmarlo antes de
// cada sync que vaya a producir un cambio visible.
//
// OJO (probado y descartado): NO llamar a esto antes de TODOS los
// context.sync() de estas funciones, incluidos los que solo leen
// propiedades (.load(...) sin ninguna escritura de por medio) — la propia
// documentación de Office avisa de que llamarlo repetidamente (p.ej. en
// bucle) provoca MÁS parpadeo, no menos, y añade lentitud perceptible. Se
// llama solo delante de los sync que van justo después de un borrado o
// una escritura real (clear/.values=/.format...); los que solo cargan
// datos para leerlos no lo necesitan y se dejan tal cual.
function draco_suspendScreenUpdating(context) {
    try {
        context.application.suspendScreenUpdatingUntilNextSync();
    } catch (e) {
        // Excel/host sin soporte (API set < 1.9 de ExcelApi): se ignora y
        // el refresco simplemente vuelve a parpadear como antes, sin
        // romper nada.
    }
}

async function jsonTo3MatricesCore(context, json, reportIdOverride) {
    const reportId = reportIdOverride !== undefined ? reportIdOverride : activeReportIdOrNull();
    let tPerf = performance.now(); // instrumentación de tiempos, ver draco_perfMark
    DracoLastJsonByReport.set(dracoStateKey(reportId), json); // cache para poder repintar en un toggle +/- sin re-consultar BigQuery

    // Se reconstruye entera en cada refresco (no se va acumulando entre
    // refrescos) — ver comentario en la declaración de DracoAxisCrossValues.
    const axisCrossValues = new Map();
    DracoAxisCrossValues.set(reportId, axisCrossValues);

    // Baseline del IMPORTE (DracoPlanningBaselineValues): hasta ahora SOLO
    // se rellenaba en jsonPaintValues (el refresco "Fixed" antiguo), pero
    // los informes con Filas/Columnas jerárquicas se pintan por AQUÍ
    // (factCells, más abajo) y jsonPaintValues nunca llega a ejecutarse
    // para ellos — así que el baseline se quedaba siempre vacío, y el
    // writeback insertaba el valor ABSOLUTO tecleado en vez del delta
    // contra el último refresco. Se rellena igual que jsonPaintValues:
    // SOLO si es informe de planificación, keyed "row_col" -> valor.
    const report_ = window.ReportStore ? window.ReportStore.getReport(reportId) : null;
    const isPlanningReport_ = !!(report_ && report_.reportProperties && report_.reportProperties.planningReport);
    const factBaseline = isPlanningReport_ ? new Map() : null;

    // Definición del informe que se va a pintar (ReportState): se recarga
    // aquí para que corresponda SIEMPRE a reportId. Antes se pintaba con
    // lo que hubiera dejado en ReportState el último loadReportDefinition,
    // que en "Actualizar todos" (todas las consultas se preparan primero y
    // se pintan después) o en un +/- de un informe que no es el último
    // diseñado era el de OTRO informe. Además aporta las opciones de campo
    // de cada medida (formato), que el pintado necesita.
    const editReportGrid = await getEditReportGrid(context, reportId);
    loadReportDefinition(editReportGrid, reportId);
    applyFilterMeasuresToDynamicReport();

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
    tPerf = draco_perfMark(reportId, `Parseo JSON manual (${valores.length} valores)`, tPerf);

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
    // (editReportGrid ya se ha leído al principio de la función)

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
    const levelsFilas = buildDracoAxisLevels(editReportGrid, 8, 9, 10, 11);      // H/I/J (+K: jerarquía)
    const levelsColumnas = buildDracoAxisLevels(editReportGrid, 14, 15, 16, 17); // N/O/P (+Q: jerarquía)

    const planFilas = computeAxisPaintPlan(levelsFilas);
    const planColumnas = computeAxisPaintPlan(levelsColumnas);

    // Ancla de cada eje (EDIT_REPORT!H10 = Filas, N10 = Columnas). El
    // taskpane (RangeAxis.addressOf) deja la dirección VACÍA cuando el eje
    // no tiene ningún campo (alto/ancho 0) y antes eso hacía fallar
    // parseAddress -> no se pintaba nada hasta tener campos en los DOS
    // ejes. Ahora, si falta una, se deduce de la otra con la misma regla
    // de RangeAxis (Columnas empieza justo a la derecha de las cabeceras de
    // Filas; Filas empieza justo debajo de las cabeceras de Columnas), y
    // si faltan las dos (solo medidas en Filtros), se usa la celda ancla
    // del informe (EDIT_REPORT!E1) o A1.
    const tryParseAddress = (raw) => {
        const t = String(raw === null || raw === undefined ? "" : raw).trim();
        if (t === "") return null;
        try { return parseAddress(t); } catch (e) { return null; }
    };
    let RRows = tryParseAddress(cellValue(editReportGrid, 10, 8));  // EDIT_REPORT!H10
    let RCols = tryParseAddress(cellValue(editReportGrid, 10, 14)); // EDIT_REPORT!N10
    if (!RRows && !RCols) {
        const anchor = tryParseAddress(cellValue(editReportGrid, 1, 5)) || { row: 1, col: 1 }; // EDIT_REPORT!E1
        RRows = { row: anchor.row, col: anchor.col };
        RCols = { row: anchor.row, col: anchor.col };
    } else if (!RCols) {
        RCols = { row: RRows.row - planColumnas.totalIaux, col: RRows.col + planFilas.totalIaux };
    } else if (!RRows) {
        RRows = { row: RCols.row + planColumnas.totalIaux, col: RCols.col - planFilas.totalIaux };
    }
    RRows.row = Math.max(1, RRows.row); RRows.col = Math.max(1, RRows.col);
    RCols.row = Math.max(1, RCols.row); RCols.col = Math.max(1, RCols.col);

    // INFORME VINCULADO (ver dracoPaintLinkedChain): se pinta pegado al principal y
    // SIN las cabeceras del eje que comparte con él (filas a la izquierda o
    // columnas arriba: esas solo están en el principal). Solo activo mientras se
    // pinta un vinculado; en cualquier otro informe esto no hace nada.
    const linkPaint = (DracoLinkPaint && Number(DracoLinkPaint.reportId) === Number(reportId)) ? DracoLinkPaint : null;
    if (linkPaint) {
        RRows = { row: linkPaint.row, col: Math.max(1, linkPaint.col - planFilas.totalIaux) };
        RCols = { row: Math.max(1, linkPaint.row - planColumnas.totalIaux), col: linkPaint.col };
    }

    const rowsOffRow = RRows.row - 1;
    const rowsOffCol = RRows.col - 1;
    const colsOffRow = RCols.row - 1;
    const colsOffCol = RCols.col - 1;

    // ---- Medidas arrastradas a la zona "Filtros" (ver loadFilters/
    // appendLockedFilterRangesToState + applyFilterMeasuresToDynamicReport,
    // que ya las añadió al FINAL de ReportState.Measures antes de llamar a
    // buildSQL): siguen reservando su propio hueco físico en el bloque de
    // VALORES (measuresOnRowsAxis/measuresOnColsAxis + measureCount, más
    // abajo, se calculan sobre measureLevels.length y por tanto ya cuentan
    // con ellas), con el mismo mIdx que le toca por su posición en
    // ReportState.Measures, y su VALOR se pinta con normalidad (si no, un
    // informe cuya única medida esté en Filtros se quedaría sin ningún
    // dato que mostrar). Lo único que NO se pinta es su CABECERA (la
    // etiqueta con el nombre de la medida), así que no deben sumar una
    // fila/columna más al bloque de CABECERA (planFilas/planColumnas.
    // totalIaux, que es lo que fija el alto/ancho de Draco_<id>_Rows/Cols):
    // si el eje YA tenía alguna medida real, comparten su mismo hueco de
    // cabecera (ya contado); si el eje no tenía NINGUNA medida real (todas
    // las medidas del informe están en Filtros), no se reserva ninguna
    // fila/columna de cabecera para ellas — sencillamente no hay ninguna
    // etiqueta que pintar ahí. Se tratan como si estuvieran en la ÚLTIMA
    // posición del eje que YA lleva las medidas (Filas o Columnas, el que
    // sea: el taskpane garantiza que todas las medidas reales están en uno
    // solo); si ninguno de los dos ejes tiene medidas reales, se usa
    // Columnas por defecto (solo a efectos de measuresOnColsAxis/
    // measureCount; no reserva cabecera, ver arriba).
    const filterMeasureNames = ReportState.FilterMeasureNames || [];
    if (filterMeasureNames.length > 0) {
        const targetPlan = planFilas.measureLevels.length > 0 ? planFilas : planColumnas;
        // Si ya había medidas reales, se reutiliza su iAux (ya contado en
        // totalIaux); si no, no hace falta ningún iAux válido: las
        // entradas "hidden" nunca llegan a leer .iAux (los dos bucles que
        // pintan CABECERAS de medida las saltan antes de usarlo; el bucle
        // que pinta VALORES no usa .iAux en absoluto).
        const measureIaux = targetPlan.measureLevels.length > 0 ? targetPlan.measureLevels[0].iAux : null;
        filterMeasureNames.forEach(name => {
            targetPlan.measureLevels.push({
                iAux: measureIaux,
                ordinal: targetPlan.measureLevels.length,
                label: name,
                hidden: true // no se pinta su cabecera (el VALOR sí se pinta, ver el bucle de factCells más abajo)
            });
        });
    }

    // ---- Varias medidas en el MISMO eje (Σ Medidas del taskpane): el
    // taskpane garantiza que todas las medidas están en un único eje
    // (nunca repartidas entre Filas y Columnas), así que como mucho uno de
    // los dos "measureLevels" de abajo tiene contenido. Ese eje se pinta
    // measureCount veces más ancho/alto (una fila/columna física física
    // por medida), en vez de una sola celda con el valor de la última
    // medida (como pasaba antes de este cambio).
    // realMeasureCount = 0 -> informe solo con dimensiones: se pintan las
    // cabeceras de Filas/Columnas y ningún valor. measureCount se mantiene
    // en >= 1 porque solo se usa para "explotar" el eje que lleva medidas.
    const realMeasureCount = ReportState.Measures.length;
    const measureCount = Math.max(1, realMeasureCount);
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
    if (realMeasureCount > 0 && (reportProps.suppressZeroRows || reportProps.suppressZeroCols)) {
        const rowsWithValue = new Set();
        const colsWithValue = new Set();
        // Cálculos por posición: llegan del SQL sin valor (la fórmula de
        // Excel se escribe al pintar), así que no deben suprimirse por
        // "cero/vacío".
        const positionCalcChecks = (ReportState.Calcs || []).map(spec => ({
            fi: 2 + (spec.axis === "Columns" ? spec.idx : ReportState.ColumnCount + spec.idx),
            names: new Set(spec.calcs.filter(c => c.mode === "position").map(c => String(c.name).trim()))
        })).filter(x => x.names.size > 0);
        for (let i = 0; i < filas; i++) {
            const f = fact[i];
            // Con varias medidas se suprime la fila/columna solo si TODAS
            // sus medidas son cero/vacío (antes solo miraba la última).
            let anyNonZero = positionCalcChecks.some(x => x.names.has(String(f[x.fi]).trim()));
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

    const rowsFilter = filterAndCompactDracoAxis(rowDict, totalDimFilas, dracoCollapseState.rows.collapsed, flagsFilas, { companion: planFilas.companionReal, subtotalsOnTop: reportProps.subtotalsOnTop, hierarchyNodesOnTop: reportProps.hierarchyNodesOnTop });
    const colsFilter = filterAndCompactDracoAxis(colDict, totalDimCols, dracoCollapseState.cols.collapsed, flagsColumnas, { companion: planColumnas.companionReal, subtotalsOnTop: reportProps.subtotalsOnTop, hierarchyNodesOnTop: reportProps.hierarchyNodesOnTop });

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
    tPerf = draco_perfMark(reportId, "Construir estructuras en memoria (fact/filas/cols)", tPerf);
    const rangeNamesForClear = dracoRangeNames(reportId);
    const prevNamedItems = [rangeNamesForClear.rows, rangeNamesForClear.cols, rangeNamesForClear.values]
        .map(n => context.workbook.names.getItemOrNullObject(n));
    prevNamedItems.forEach(it => it.load("isNullObject"));
    await context.sync();

    const prevRanges = prevNamedItems.filter(it => !it.isNullObject).map(it => it.getRange());
    prevRanges.forEach(r => r.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]));
    if (prevRanges.length > 0) {
        await context.sync();
    }

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

    // Antes de refrescar: borrar contenido de los rangos con nombre
    // Draco_<id>_Rows / Draco_<id>_Cols / Draco_<id>_Values DE ESTE
    // INFORME (y solo de este) de la ejecución anterior (si existen), para
    // no arrastrar colores/bordes de una tabla previa más grande o con
    // otra forma — y, sobre todo, para NO tocar los rangos de otros
    // informes (antes, con un nombre compartido "Draco_001_*", refrescar
    // un informe borraba lo pintado de otro). Si "Sobrescribir formatos"
    // está desactivado, solo se borra el CONTENIDO (Contents), no el
    // formato — antes se borraba siempre con "all" sin mirar esta opción,
    // así que un informe con overwriteFormats=false igualmente perdía el
    // formato que el usuario hubiera dejado a mano en la ejecución anterior.
    //
    // OJO — TODO EL RESTO DE ESTA FUNCIÓN (borrados, FACT, FILAS,
    // COLUMNAS, resaltado de totales, fuente/número/bordes, rangos con
    // nombre) queda en cola sin sincronizar hasta el ÚNICO sync grande de
    // más abajo (justo antes de "Registrar listeners"). Antes cada paso
    // tenía su propio context.sync(), y cada uno de esos "viajes" a Excel
    // parece tener un coste fijo bastante alto e irregular en este
    // entorno (2 celdas tardaban a veces lo mismo que 432) — exactamente
    // el mismo motivo por el que en VBA se hacía todo con
    // Application.ScreenUpdating=False de un tirón. Aquí no hace falta
    // reconsultar si los 3 nombres existen (isNullObject) para borrarlos
    // NI para redefinirlos más abajo: ya lo sabemos por prevNamedItems,
    // cargado un poco más arriba — antes se volvía a preguntar dos veces
    // más (una en clearDracoNamedRanges, otra en applyDracoNamedRanges),
    // cada una con su propio sync.
    const clearMode = reportProps.overwriteFormats ? Excel.ClearApplyTo.all : Excel.ClearApplyTo.contents;
    for (const it of prevNamedItems) {
        if (!it.isNullObject) {
            const prevRange = it.getRange();
            prevRange.clear(clearMode);
            // Barras de datos de los cálculos (formato condicional): se
            // quitan siempre que se sobrescriben formatos.
            if (reportProps.overwriteFormats) prevRange.conditionalFormats.clearAll();
        }
    }

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
            ).clear(clearMode);
        }
    }
    tPerf = draco_perfMark(reportId, "Borrar rangos anteriores (en cola, sin sync)", tPerf);

    /* -------------------------------------------------------------
     * 1) FACT — construir el bloque completo en memoria y escribirlo
     *    de una sola vez con un único range.values = [...]. Las filas
     *    cuyo ROW_ID/COLUMN_ID original haya quedado oculto por una
     *    jerarquía contraída se descartan, y los IDs restantes se
     *    remapean a la numeración compactada 1..N.
     * ----------------------------------------------------------- */
    const factCells = new Map(); // "row_col" -> {row, col, value}
    const calcEntries = (ReportState.Calcs && ReportState.Calcs.length) ? [] : null;
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
        //
        // Las medidas en Filtros (hidden) SÍ escriben su valor aquí — solo
        // se omite su CABECERA (label de medida, más abajo): si se omitiera
        // también el valor, un informe cuya ÚNICA medida está en Filtros
        // quedaría con la tabla completamente en blanco (sin ningún dato
        // que mostrar), que no es lo que se pedía — "no pinte esa
        // fila/columna" se refiere a la etiqueta de cabecera, no al dato.
        for (let mIdx = 0; mIdx < realMeasureCount; mIdx++) {
            const physRow = explodeForMeasures(newRowId, measuresOnRowsAxis, mIdx);
            const physCol = explodeForMeasures(newColId, measuresOnColsAxis, mIdx);
            const row = physRow + rowsOffRow;
            const col = physCol + colsOffCol;
            // NULL (o sin valor) -> celda en blanco, ver coerceCellLiteral.
            const value = coerceCellLiteral(f[measureFieldBase + mIdx]);
            factCells.set(row + "_" + col, { row, col, value, mIdx });
            if (factBaseline) factBaseline.set(row + "_" + col, value);
            if (calcEntries) calcEntries.push({ row, col, mIdx, f });
        }
    }

    // Cálculos entre miembros: fórmulas de Excel (por posición), celdas
    // excluidas de la planificación y líneas para formato/barras.
    const calcPaint = calcEntries
        ? resolveDracoCalcCells(calcEntries, factCells, factBaseline, ReportState.ColumnCount)
        : null;
    DracoCalcCellsByReport.set(dracoStateKey(reportId), calcPaint ? calcPaint.cellSet : new Set());

    await writeCellBlock(context, sheet, factCells, true);
    if (calcPaint) {
        for (const fm of calcPaint.formulas) {
            sheet.getRangeByIndexes(fm.row - 1, fm.col - 1, 1, 1).formulas = [[fm.formula]];
        }
    }
    tPerf = draco_perfMark(reportId, `Escribir FACT (${factCells.size} celdas, en cola)`, tPerf);

    if (isPlanningReport_) {
        DracoPlanningBaselineValues.set(reportId, factBaseline);
    } else {
        DracoPlanningBaselineValues.delete(reportId); // por si dejó de ser de planificación desde el último refresco
    }

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
            if (iAux === null) continue; // "Solo texto": el ID no se pinta

            if (String(V[i]).toLowerCase().indexOf("null") !== 0) {
                const indent = flag === 1 ? 0 : Math.max(0, flag - 1);
                const text = coerceCellLiteral(V[i]);
                // Eje Estático: se escribe como fórmula EPM_VALUE editable
                // (mismo formato que lee el flujo Fijo), no como texto plano.
                const pI = planFilas.idPartnerReal[i];
                const cellVal = rowsStatic
                    ? (pI ? buildEpmValueFormula(fieldsFilas[pI].dim, fieldsFilas[pI].attr, coerceCellLiteral(V[pI]), text)
                          : buildEpmValueFormula(fieldsFilas[i].dim, fieldsFilas[i].attr, text))
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
                    // Valor canónico (antes de Excel) para el writeback —
                    // ver DracoAxisCrossValues. Mismas coordenadas físicas
                    // que la celda que se acaba de pintar.
                    axisCrossValues.set(row + "_" + col, pI
                        ? { dim: fieldsFilas[pI].dim, attr: fieldsFilas[pI].attr, value: coerceCellLiteral(V[pI]) }
                        : { dim: fieldsFilas[i].dim, attr: fieldsFilas[i].attr, value: text });
                }
            }
        }

        // Filas MEASURE del eje (p.ej. "MEASURE" -> "IMPORTE"): no vienen del
        // dict (no son un valor variable por ROW_ID); se pintan como etiqueta
        // en su propia columna física, una por cada medida del grupo, en la
        // fila física que le corresponde a ESA medida (measuresOnRowsAxis).
        for (const m of planFilas.measureLevels) {
            if (m.hidden) continue; // medida en Filtros: no se pinta su cabecera
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

    const skipFilasHdr = !!(linkPaint && linkPaint.axis === "rows");   // vinculado por filas: sin cabeceras de fila
    if (!skipFilasHdr) await writeCellBlock(context, sheet, filasCells, true);
    let filasRunCount = 0;
    if (reportProps.overwriteFormats && !skipFilasHdr) {
        filasRunCount = await writeIndentAndColorRuns(context, sheet, filasCells, "col", rowsOffCol, true);
    }
    tPerf = draco_perfMark(reportId, `Escribir+formatear FILAS (${filasCells.size} celdas, ${filasRunCount} tramos, en cola)`, tPerf);

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
            if (iAux === null) continue; // "Solo texto": el ID no se pinta

            if (String(V[i]).toLowerCase().indexOf("null") !== 0) {
                const indent = flag === 1 ? 0 : Math.max(0, flag - 1);
                const text = coerceCellLiteral(V[i]);
                const pI = planColumnas.idPartnerReal[i];
                const cellVal = colsStatic
                    ? (pI ? buildEpmValueFormula(fieldsColumnas[pI].dim, fieldsColumnas[pI].attr, coerceCellLiteral(V[pI]), text)
                          : buildEpmValueFormula(fieldsColumnas[i].dim, fieldsColumnas[i].attr, text))
                    : text;
                // Ver comentario equivalente en el bloque de FILAS: text puede
                // ser un Number (dimensión INTEGER) y no tiene .trim().
                const isTotal = String(text).trim().toUpperCase() === "TOTAL";
                for (let mIdx = 0; mIdx < repeatCount; mIdx++) {
                    const row = iAux + colsOffRow;
                    const col = explodeForMeasures(Number(V[0]), measuresOnColsAxis, mIdx) + colsOffCol;
                    columnasCells.set(row + "_" + col, { row, col, value: cellVal, indent, field: iAux, isTotal });
                    // Valor canónico (antes de Excel) para el writeback —
                    // ver DracoAxisCrossValues. Mismas coordenadas físicas
                    // que la celda que se acaba de pintar.
                    axisCrossValues.set(row + "_" + col, pI
                        ? { dim: fieldsColumnas[pI].dim, attr: fieldsColumnas[pI].attr, value: coerceCellLiteral(V[pI]) }
                        : { dim: fieldsColumnas[i].dim, attr: fieldsColumnas[i].attr, value: text });
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
            if (m.hidden) continue; // medida en Filtros: no se pinta su cabecera
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

    const skipColsHdr = !!(linkPaint && linkPaint.axis === "cols");    // vinculado por columnas: sin cabeceras de columna
    if (!skipColsHdr) await writeCellBlock(context, sheet, columnasCells, true);
    let columnasRunCount = 0;
    if (reportProps.overwriteFormats && !skipColsHdr) {
        columnasRunCount = await writeIndentAndColorRuns(context, sheet, columnasCells, "row", colsOffRow, true);
    }
    tPerf = draco_perfMark(reportId, `Escribir+formatear COLUMNAS (${columnasCells.size} celdas, ${columnasRunCount} tramos, en cola)`, tPerf);

    // ---- Punto 7: fondo RGB(255,255,204) en cabeceras "Total" y en los
    // valores de fila/columna de total (gateado por "Sobrescribir
    // formatos", igual que el resto de formato). ----
    if (reportProps.overwriteFormats) {
        await applyDracoTotalHighlight(context, sheet, filasCells, columnasCells, factCells, true, linkPaint ? linkPaint.axis : null);
        tPerf = draco_perfMark(reportId, "Resaltado de totales (en cola)", tPerf);
    }

    // ══ ÚNICO sync grande para TODO lo anterior (borrados + FACT + FILAS
    // + COLUMNAS + resaltado de totales) ══ — antes cada uno de esos pasos
    // tenía su propio context.sync(); ahora es UN solo viaje a Excel para
    // todo el bloque de escritura de datos. El "Formato general" que
    // viene justo después (fuente/número/bordes + rangos con nombre) SÍ
    // necesita sus propios 2 sync (los bordes lo exigen, ver comentario
    // en applyDracoNamedRanges), pero ya no arrastra ningún sync previo
    // sin fusionar.
    draco_suspendScreenUpdating(context);
    await context.sync();
    tPerf = draco_perfMark(reportId, "── Sync grande: borrado + FACT + FILAS + COLUMNAS + totales ──", tPerf);

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

    const namedRanges = await applyDracoNamedRanges(context, sheet, {
        RRows, RCols,
        totalDimFilas: skipFilasHdr ? 0 : paintedFilasCols,   // el vinculado no registra el rango de la cabecera compartida
        totalDimCols: skipColsHdr ? 0 : paintedColsRows,
        maxRowId: physicalMaxRowId, maxColId: physicalMaxColId,
        applyVisualFormat: reportProps.overwriteFormats,
        hasValues: realMeasureCount > 0
    }, reportId, prevNamedItems);
    tPerf = draco_perfMark(reportId, "Rangos con nombre + formato general (fuente/número/bordes)", tPerf);

    // Formato de número POR MEDIDA (Opciones de campo: Número/Porcentaje +
    // nº de decimales), solo con "Sobrescribir formatos" activado — igual
    // que el resto del formato. Va después de applyDracoNamedRanges, que
    // deja el formato general por defecto (#,##0.00) en todo el bloque.
    if (reportProps.overwriteFormats && namedRanges.valuesRange && realMeasureCount > 0) {
        try {
            const applied = applyDracoMeasureNumberFormats(namedRanges.valuesRange, {
                rows: physicalMaxRowId, cols: physicalMaxColId,
                measureCount: realMeasureCount,
                measuresOnRowsAxis, measuresOnColsAxis
            });
            if (applied) {
                draco_suspendScreenUpdating(context);
                await context.sync();
            }
        } catch (e) {
            console.warn("[Draco] No se pudo aplicar el formato de número de las medidas:", e);
        }
    }
    if (reportProps.overwriteFormats && calcPaint && calcPaint.lines.size > 0) {
        try {
            applyDracoCalcFormats(sheet, calcPaint);
            draco_suspendScreenUpdating(context);
            await context.sync();
        } catch (e) {
            console.warn("[Draco] No se pudo aplicar el formato de los cálculos:", e);
        }
    }
    tPerf = draco_perfMark(reportId, "Formato de número por medida", tPerf);

    // 6) Registrar (una sola vez por hoja) los listeners de clic/edición:
    //    resuelven los indicadores +/- pintados arriba y el "Reconocimiento
    //    de miembros" sobre las celdas de Draco_<id>_Rows/Draco_<id>_Cols
    //    de ESTA hoja.
    await registerDracoSelectionHandler(context, sheet, resultSheetName);
    tPerf = draco_perfMark(reportId, "Registrar listeners de la hoja", tPerf);

    // 7) Autoajustar ancho de columnas (propiedades del informe: D6), solo
    // de los rangos con nombre Draco_<id>_Rows y Draco_<id>_Cols de ESTE
    // informe (donde se pintan sus filas/columnas), no de toda la hoja.
    // Reutiliza los MISMOS objetos Range que acaba de construir/devolver
    // applyDracoNamedRanges (rowsRange/colsRange): antes se volvían a
    // buscar por nombre con su propio load+sync, un viaje de red entero
    // solo para recuperar algo que ya teníamos en memoria.
    if (reportProps.autoFitColumns && (namedRanges.rowsRange || namedRanges.colsRange)) {
        if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(reportId, "Ajustando columnas"));
        try {
            if (namedRanges.rowsRange) namedRanges.rowsRange.format.autofitColumns();
            if (namedRanges.colsRange) namedRanges.colsRange.format.autofitColumns();

            draco_suspendScreenUpdating(context);
            await context.sync();
            draco_perfMark(reportId, "Ajustar ancho de columnas", tPerf);
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
async function applyDracoTotalHighlight(context, sheet, filasCells, columnasCells, factCells, deferSync, skipHeaderAxis) {
    const TOTAL_FILL = "#FFFFCC"; // RGB(255,255,204)

    const totalRows = new Set();
    const totalCols = new Set();
    const headerCells = [];

    // skipHeaderAxis ("rows"/"cols", informes vinculados): esas cabeceras no se pintan
    // en este informe (solo en su principal), así que no se colorean; sus filas/columnas
    // de total de VALORES sí.
    for (const cell of filasCells.values()) {
        if (cell.isTotal) {
            if (skipHeaderAxis !== "rows") headerCells.push(cell);
            totalRows.add(cell.row);
        }
    }
    for (const cell of columnasCells.values()) {
        if (cell.isTotal) {
            if (skipHeaderAxis !== "cols") headerCells.push(cell);
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

    // deferSync=true (ver jsonTo3MatricesCore): se deja en cola, el
    // llamador hace un único sync grande al final.
    if (!deferSync) {
        draco_suspendScreenUpdating(context);
        await context.sync();
    }
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
async function writeIndentAndColorRuns(context, sheet, cellsMap, axis, fieldOffset, deferSync) {
    if (cellsMap.size === 0) return 0;

    const groupKeyName = axis === "col" ? "col" : "row";
    const runKeyName = axis === "col" ? "row" : "col";

    const groups = new Map();
    for (const c of cellsMap.values()) {
        const g = c[groupKeyName];
        if (!groups.has(g)) groups.set(g, []);
        groups.get(g).push(c);
    }

    let runCount = 0; // instrumentación: nº de range.format aplicados (ver draco_perfMark)
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
                runCount++;

                runStart = k;
            }
        }
    }

    // deferSync=true (ver jsonTo3MatricesCore): se deja en cola, el
    // llamador hace un único sync grande al final.
    if (!deferSync) {
        draco_suspendScreenUpdating(context);
        await context.sync();
    }
    return runCount;
}

/**
 * Borra (formato + contenido) los rangos de la ejecución anterior a los
 * que apuntaban los nombres Draco_001_Rows/Cols/Values, si existen.
 */
/**
 * Borra el contenido de los 3 rangos con nombre de un informe
 * (Rows/Cols/Values) antes de repintarlo. Si overwriteFormats es false
 * (opción "Sobrescribir formatos" desactivada en las propiedades del
 * informe), solo se borra el CONTENIDO (Contents) y se respeta el
 * formato que el usuario haya dejado a mano — igual que el "if" que ya
 * usa jsonTo3MatricesCore más abajo para decidir si repinta formato o no;
 * antes esta limpieza no miraba esa opción para nada y borraba "all"
 * siempre, así que el formato se perdía igualmente en la propia limpieza,
 * aunque luego no se repintara.
 */
async function clearDracoNamedRanges(context, reportId, overwriteFormats) {
    const rn = dracoRangeNames(reportId);
    const names = [rn.rows, rn.cols, rn.values];
    const items = names.map(n => context.workbook.names.getItemOrNullObject(n));
    items.forEach(it => it.load("isNullObject"));
    await context.sync();

    const clearMode = overwriteFormats ? Excel.ClearApplyTo.all : Excel.ClearApplyTo.contents;
    let anyToClear = false;
    for (const it of items) {
        if (!it.isNullObject) {
            it.getRange().clear(clearMode);
            anyToClear = true;
        }
    }
    if (anyToClear) {
        draco_suspendScreenUpdating(context);
        await context.sync();
    }
}

// reportId -> Set("fila_col", 1-based) de celdas de cálculos del último pintado
const DracoCalcCellsByReport = new Map();

function isDracoCalcCell(reportId, r0, c0) {
    const set = DracoCalcCellsByReport.get(dracoStateKey(reportId));
    return !!(set && set.has((r0 + 1) + "_" + (c0 + 1)));
}

function dracoColumnLetter(col) { // 1-based
    let s = "";
    while (col > 0) {
        const m = (col - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        col = Math.floor((col - 1) / 26);
    }
    return s;
}

/**
 * Recorre las celdas de valores pintadas y, para cada atributo con
 * cálculos:
 *   - agrupa las celdas por "todo lo demás igual" (resto de campos +
 *     medida), separando miembros reales de cálculos;
 *   - en los cálculos por POSICIÓN, escribe la fórmula de Excel que
 *     apunta a las celdas del 1º, 2º... miembro real visible de ese mismo
 *     grupo (en el orden en que se ven);
 *   - quita las celdas de cálculo de la línea base de planificación.
 * Devuelve { formulas, cellSet, lines: Map(calcId -> {calc, onRows, cells}) }.
 */
function resolveDracoCalcCells(entries, factCells, factBaseline, columnCount) {
    const formulas = [];
    const cellSet = new Set();
    const lines = new Map();

    for (const spec of ReportState.Calcs) {
        const onRows = spec.axis === "Columns";            // H.. -> filas físicas
        const fi = 2 + (onRows ? spec.idx : columnCount + spec.idx);
        const byName = new Map(spec.calcs.map(c => [String(c.name).trim(), c]));
        const groups = new Map(); // clave -> { reals: [], calcs: [] }

        for (const e of entries) {
            const member = String(e.f[fi] === undefined ? "" : e.f[fi]).trim();
            const keyParts = [];
            for (let k = 2; k < 2 + ReportState.ColumnCount + ReportState.RowCount; k++) {
                if (k !== fi) keyParts.push(String(e.f[k]).trim());
            }
            keyParts.push("m" + e.mIdx);
            const key = keyParts.join("\u0001");
            if (!groups.has(key)) groups.set(key, { reals: [], calcs: [] });
            const calc = byName.get(member);
            if (calc) groups.get(key).calcs.push({ e, calc });
            else if (member !== "" && member.toLowerCase().indexOf("null") !== 0) groups.get(key).reals.push(e);
        }

        for (const g of groups.values()) {
            g.reals.sort((a, b) => onRows ? (a.row - b.row) : (a.col - b.col));
            for (const { e, calc } of g.calcs) {
                const id = e.row + "_" + e.col;
                cellSet.add(id);
                if (factBaseline) factBaseline.delete(id);
                if (!lines.has(calc.id)) lines.set(calc.id, { calc, onRows, cells: [] });
                lines.get(calc.id).cells.push({ row: e.row, col: e.col });

                if (calc.mode === "position") {
                    const formula = window.DracoCalc.toExcel(calc.ast, (pos) => {
                        const target = g.reals[pos - 1];
                        return target ? dracoColumnLetter(target.col) + target.row : null;
                    });
                    const cell = factCells.get(id);
                    if (cell) cell.value = "";
                    if (formula) formulas.push({ row: e.row, col: e.col, formula });
                }
            }
        }
    }
    return { formulas, cellSet, lines };
}

/**
 * Formato de los cálculos: número/porcentaje con sus decimales y, si se
 * pide, barras de datos (rojo/verde con eje, como en "Variation Detail").
 * La escala de las barras es común a TODAS las celdas del cálculo y deja
 * a cada lado el % de ancho indicado para el número.
 */
function applyDracoCalcFormats(sheet, calcPaint) {
    for (const { calc, onRows, cells } of calcPaint.lines.values()) {
        // Celdas del cálculo agrupadas en líneas contiguas (una fila física
        // si el atributo está en Filas, una columna si está en Columnas).
        const byLine = new Map();
        for (const c of cells) {
            const k = onRows ? c.row : c.col;
            if (!byLine.has(k)) byLine.set(k, []);
            byLine.get(k).push(onRows ? c.col : c.row);
        }
        const ranges = [];
        for (const [k, list] of byLine) {
            list.sort((a, b) => a - b);
            let start = list[0], prev = list[0];
            const flush = (a, b) => ranges.push(onRows
                ? { r: k, c: a, h: 1, w: b - a + 1 }
                : { r: a, c: k, h: b - a + 1, w: 1 });
            for (let i = 1; i < list.length; i++) {
                if (list[i] !== prev + 1) { flush(start, prev); start = list[i]; }
                prev = list[i];
            }
            flush(start, prev);
        }

        const fmt = window.DracoCalc.numberFormat(calc);
        for (const rg of ranges) {
            sheet.getRangeByIndexes(rg.r - 1, rg.c - 1, rg.h, rg.w).numberFormat = [[fmt]];
        }

        if (calc.display !== "bar" && calc.display !== "barOnly") continue;

        const addr = (rg) => "$" + dracoColumnLetter(rg.c) + "$" + rg.r + ":$" + dracoColumnLetter(rg.c + rg.w - 1) + "$" + (rg.r + rg.h - 1);
        const list = ranges.slice(0, 200).map(addr).join(",");
        const barOnly = calc.display === "barOnly";
        const L = barOnly ? 0.04 : Math.max(0, Math.min(0.9, Number(calc.labelLeft) / 100 || 0));
        const R = barOnly ? 0.04 : Math.max(0, Math.min(0.9, Number(calc.labelRight) / 100 || 0));
        const mn = "MIN(0,MIN(" + list + "))";
        const mx = "MAX(0,MAX(" + list + "))";
        const span = "MAX(0.000001," + mx + "-" + mn + ")";
        const ln = "IF(" + mn + "<0," + L + ",0)";
        const lp = "IF(" + mx + ">0," + R + ",0)";
        const bar = "MAX(0.05,1-" + ln + "-" + lp + ")";
        const lo = "=" + mn + "-" + span + "*" + ln + "/" + bar;
        const hi = "=" + mx + "+" + span + "*" + lp + "/" + bar;
        const neg = calc.invertColors ? calc.colorPos : calc.colorNeg;
        const pos = calc.invertColors ? calc.colorNeg : calc.colorPos;

        for (const rg of ranges) {
            const cf = sheet.getRangeByIndexes(rg.r - 1, rg.c - 1, rg.h, rg.w)
                .conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
            const db = cf.dataBar;
            db.showDataBarOnly = barOnly;
            db.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
            db.axisFormat = Excel.ConditionalDataBarAxisFormat.automatic;
            db.axisColor = "#000000";
            db.positiveFormat.fillColor = pos || "#70D06A";
            db.positiveFormat.gradientFill = false;
            db.negativeFormat.matchPositiveFillColor = false;
            db.negativeFormat.fillColor = neg || "#F0436A";
            db.lowerBoundRule = { type: Excel.ConditionalFormatRuleType.formula, formula: lo };
            db.upperBoundRule = { type: Excel.ConditionalFormatRuleType.formula, formula: hi };
        }
    }
}

/**
 * Aplica el formato de número de cada medida (ReportState.Measures[k].
 * NumberFormat, ver dracoNumberFormatFromOptions) sobre el bloque de
 * valores ya definido. Las medidas están todas en un mismo eje y
 * "explotadas" (measureCount filas/columnas físicas consecutivas por
 * cada ID lógico, ver explodeForMeasures en jsonTo3MatricesCore), así que
 * la medida de una celda es (fila % measureCount) si van en Filas o
 * (columna % measureCount) si van en Columnas. Deja las operaciones en
 * cola (sin sync); devuelve true si ha encolado alguna.
 */
function applyDracoMeasureNumberFormats(valuesRange, layout) {
    const DEFAULT_FMT = "#,##0.00";
    const { rows, cols, measureCount, measuresOnRowsAxis, measuresOnColsAxis } = layout;
    if (!valuesRange || rows <= 0 || cols <= 0 || measureCount <= 0) return false;

    const fmts = [];
    for (let k = 0; k < measureCount; k++) {
        const m = ReportState.Measures[k];
        fmts.push((m && m.NumberFormat) || DEFAULT_FMT);
    }
    // Nada que cambiar respecto al formato por defecto ya aplicado.
    if (fmts.every(f => f === DEFAULT_FMT)) return false;

    // Todas iguales (o una sola medida): una única asignación.
    if (fmts.every(f => f === fmts[0]) || (!measuresOnRowsAxis && !measuresOnColsAxis)) {
        valuesRange.numberFormat = [[fmts[0]]];
        return true;
    }

    const fmtFor = (r, c) => fmts[measuresOnRowsAxis ? (r % measureCount) : (c % measureCount)];

    // Bloques razonables: matriz completa en una sola operación.
    if (rows * cols <= 200000) {
        const matrix = [];
        for (let r = 0; r < rows; r++) {
            const line = new Array(cols);
            for (let c = 0; c < cols; c++) line[c] = fmtFor(r, c);
            matrix.push(line);
        }
        valuesRange.numberFormat = matrix;
        return true;
    }

    // Bloques enormes: fila a fila (medidas en Filas) o columna a columna
    // (medidas en Columnas), solo las que no usan el formato por defecto.
    if (measuresOnRowsAxis) {
        for (let r = 0; r < rows; r++) {
            const f = fmts[r % measureCount];
            if (f !== DEFAULT_FMT) valuesRange.getRow(r).numberFormat = [[f]];
        }
    } else {
        for (let c = 0; c < cols; c++) {
            const f = fmts[c % measureCount];
            if (f !== DEFAULT_FMT) valuesRange.getColumn(c).numberFormat = [[f]];
        }
    }
    return true;
}

/**
 * Crea/actualiza los 3 rangos con nombre a partir de la última tabla
 * pintada, y aplica: fuente Segoe UI 9 en los tres, número 2 decimales +
 * separador de miles + centrado en Draco_001_Values, y un borde fino
 * RGB(13,23,42) alrededor de cada uno de los 3 rangos.
 *
 * existingItems (opcional): los mismos 3 NamedItem (rows/cols/values, en
 * ese orden) que jsonTo3MatricesCore ya comprobó al principio del
 * refresco (prevNamedItems) — se reutilizan para no volver a preguntar
 * si existen (antes: 1 sync más aquí, redundante con esa comprobación
 * previa). Si no se pasa, se comporta como si no existiera ninguno
 * (add() directo; si el nombre ya existiera de verdad, Excel lo rechaza
 * — por eso conviene pasar siempre existingItems desde el refresco
 * normal, que sí los tiene).
 *
 * Devuelve {rowsRange, colsRange, valuesRange} (o todo null si no había
 * datos) para que el autoajuste de columnas, justo después, pueda
 * reutilizar estos MISMOS objetos Range sin tener que volver a buscarlos
 * por nombre (otro sync que se ahorra).
 */
async function applyDracoNamedRanges(context, sheet, dims, reportId, existingItems) {
    const { RRows, RCols, totalDimFilas, totalDimCols, maxRowId, maxColId, applyVisualFormat } = dims;
    const doFormat = applyVisualFormat !== false; // por defecto, sí formatear (comportamiento previo)
    const hasValues = dims.hasValues !== false;   // false = informe sin medidas (solo cabeceras)
    let tPerf = performance.now(); // instrumentación de tiempos, ver draco_perfMark

    // Cada rango se crea solo si tiene tamaño: un informe con campos solo
    // en Filas (o solo en Columnas), o sin medidas, tiene alguno de los 3
    // vacíos. Antes, si faltaba cualquiera, no se creaba NINGUNO — y sin
    // rangos con nombre el siguiente refresco no sabía qué borrar.
    const rowsRange = (maxRowId > 0 && totalDimFilas > 0)
        ? sheet.getRangeByIndexes(RRows.row - 1, RRows.col - 1, maxRowId, totalDimFilas) : null;
    const colsRange = (maxColId > 0 && totalDimCols > 0)
        ? sheet.getRangeByIndexes(RCols.row - 1, RCols.col - 1, totalDimCols, maxColId) : null;
    const valuesRange = (hasValues && maxRowId > 0 && maxColId > 0)
        ? sheet.getRangeByIndexes(RRows.row - 1, RCols.col - 1, maxRowId, maxColId) : null;

    const rn = dracoRangeNames(reportId);
    const defs = [
        { name: rn.rows, range: rowsRange },
        { name: rn.cols, range: colsRange },
        { name: rn.values, range: valuesRange }
    ].filter(d => d.range);

    if (defs.length === 0) {
        // Nada pintado: se borran los nombres antiguos (si los había) para
        // que no apunten a una tabla que ya no existe.
        if (existingItems && existingItems.some(it => !it.isNullObject)) {
            existingItems.forEach(it => { if (!it.isNullObject) it.delete(); });
            await context.sync();
        }
        return { rowsRange: null, colsRange: null, valuesRange: null };
    }
    const presentRanges = defs.map(d => d.range);

    // "Sobrescribir formatos" (propiedades del informe) desactivado: se
    // conservan el color/fuente/bordes que el usuario haya tocado a mano,
    // y solo se (re)definen los 3 rangos con nombre para que apunten al
    // tamaño actual de la tabla.
    if (doFormat) {
        // ---- Fuente Segoe UI 9 en los tres rangos ----
        for (const r of presentRanges) {
            r.format.font.name = DRACO_FONT_NAME;
            r.format.font.size = DRACO_FONT_SIZE;
        }

        // ---- Valores: 2 decimales + separador de miles, centrado ----
        // (formato por defecto; las medidas con formato propio en
        // "Opciones de campo" lo sustituyen después, ver
        // applyDracoMeasureNumberFormats)
        if (valuesRange) {
            valuesRange.numberFormat = [["#,##0.00"]];
            valuesRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;
        }

        // ---- Bordes: primero se elimina cualquier borde existente del rango
        //      (externo + interior, restos de refrescos anteriores con otra
        //      forma/tamaño de tabla) y SOLO DESPUÉS se pinta el borde exterior
        //      fino nuevo. Hacerlo en el mismo lote sin sync intermedio hace
        //      que Excel no aplique bien el cambio (probado).
        //
        // DIAGNÓSTICO TEMPORAL: se han separado aquí en 4 sync (en vez de
        // los 2 habituales) para medir por separado bordes vs. nombres —
        // esto es MÁS LENTO a propósito, solo para saber dónde pesan de
        // verdad los ~700 ms observados. Una vez confirmado, se vuelve a
        // fusionar en 2 (quitar bordes + borrar nombres juntos, pintar
        // bordes + crear nombres juntos), como estaba antes de este mark.
        const ALL_BORDER_EDGES = [
            Excel.BorderIndex.edgeTop, Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft, Excel.BorderIndex.edgeRight,
            Excel.BorderIndex.insideHorizontal, Excel.BorderIndex.insideVertical
        ];
        const OUTER_BORDER_EDGES = [
            Excel.BorderIndex.edgeTop, Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft, Excel.BorderIndex.edgeRight
        ];

        // 1a: quitar bordes (+ fuente/número, que no necesitan sync propio).
        for (const r of presentRanges) {
            for (const edge of ALL_BORDER_EDGES) {
                r.format.borders.getItem(edge).style = Excel.BorderLineStyle.none;
            }
        }
        draco_suspendScreenUpdating(context);
        await context.sync();
        tPerf = draco_perfMark(reportId, "  [diag] Quitar bordes (+fuente/número)", tPerf);

        // 1b: borrar los 3 nombres antiguos (si existían).
        if (existingItems) {
            existingItems.forEach(it => { if (!it.isNullObject) it.delete(); });
        }
        await context.sync();
        tPerf = draco_perfMark(reportId, "  [diag] Borrar nombres antiguos", tPerf);

        // 2a: pintar el borde exterior fino, color RGB(13,23,42).
        for (const r of presentRanges) {
            for (const edge of OUTER_BORDER_EDGES) {
                const border = r.format.borders.getItem(edge);
                border.style = Excel.BorderLineStyle.continuous;
                border.weight = Excel.BorderWeight.thin;
                border.color = DRACO_BORDER_COLOR;
            }
        }
        draco_suspendScreenUpdating(context);
        await context.sync();
        tPerf = draco_perfMark(reportId, "  [diag] Pintar bordes", tPerf);

        // 2b: crear los 3 nombres nuevos.
        defs.forEach(d => context.workbook.names.add(d.name, d.range));
        await context.sync();
        draco_perfMark(reportId, "  [diag] Crear nombres nuevos", tPerf);
    } else {
        // Sin "Sobrescribir formatos": solo hay que (re)definir los 3
        // nombres, sin tocar fuente/número/bordes. Sigue haciendo falta
        // borrar antes de añadir (Excel no deja redefinir un nombre ya
        // existente con add() directamente), pero ya sin la comprobación
        // de existencia redundante (existingItems ya la trae hecha).
        let anyDeleted = false;
        if (existingItems) {
            existingItems.forEach(it => {
                if (!it.isNullObject) { it.delete(); anyDeleted = true; }
            });
        }
        if (anyDeleted) {
            draco_suspendScreenUpdating(context);
            await context.sync();
        }
        defs.forEach(d => context.workbook.names.add(d.name, d.range));
        draco_suspendScreenUpdating(context);
        await context.sync();
        draco_perfMark(reportId, "  Redefinir rangos con nombre (sin formato)", tPerf);
    }

    return { rowsRange, colsRange, valuesRange };
}

/**
 * Escribe un conjunto de celdas {row,col,value} en el MÍNIMO número de
 * llamadas a la API de Excel posible: agrupa por bounding box, LEE el
 * contenido actual de ese rectángulo (para no pisar nada fuera de las
 * celdas concretas que tocan — mismo comportamiento que escribir celda a
 * celda sin ClearContents previo) y hace una única escritura de vuelta.
 *
 * NOTA sobre fechas: esta función YA NO intenta evitar que Excel
 * reinterprete un texto tipo "2025-08-01" como fecha (no se toca
 * numberFormat — se dejó expresamente así). Para el writeback de Draco
 * Planning, el valor de verdad de cada dimensión se lee de
 * DracoAxisCrossValues (el JSON tal como vino de BigQuery), no de lo que
 * haya quedado pintado en la celda — así da igual cómo Excel decida
 * mostrarlo. Ver dracoPlanningCrossItemsForCellDynamic.
 */
async function writeCellBlock(context, sheet, cellsMap, deferSync) {
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
        if (!deferSync) {
            draco_suspendScreenUpdating(context);
            await context.sync();
        }
        return;
    }

    // El área que ocupa este bloque ya viene SIEMPRE recién borrada
    // (clearDracoNamedRanges + limpieza de prevBounds, justo antes de
    // llamar a esta función para factCells/filasCells/columnasCells — ver
    // jsonTo3MatricesCore) o nunca se ha usado (celda en blanco por
    // defecto): cualquier celda del rectángulo que no esté en cellsMap YA
    // es "" ahora mismo. Antes se leía el rectángulo entero (range.load
    // + sync) solo para volver a copiar esos mismos huecos en blanco —
    // una vuelta de red entera desperdiciada. Se elimina esa lectura y se
    // construye el array ya con "" de partida: de 2 sync() por llamada a
    // 1 (y esta función se llama 3 veces por refresco: fact/filas/cols).
    // OJO: esto NO vale para un área con contenido ajeno que deba
    // conservarse en los huecos (por eso jsonPaintValues, que sí puede
    // convivir con fórmulas/etiquetas fijas del usuario en los huecos,
    // sigue leyendo antes de escribir).
    const grid = Array.from({ length: numRows }, () => new Array(numCols).fill(""));
    for (const c of cellsMap.values()) {
        grid[c.row - minRow][c.col - minCol] = c.value;
    }

    const range = sheet.getRangeByIndexes(minRow - 1, minCol - 1, numRows, numCols);
    range.values = grid;
    // deferSync=true (ver jsonTo3MatricesCore): se deja en cola sin
    // sincronizar — el llamador hace UN solo sync grande al final, junto
    // con FILAS/COLUMNAS/resaltado/formato/rangos con nombre, en vez de
    // uno por cada bloque.
    if (!deferSync) {
        draco_suspendScreenUpdating(context);
        await context.sync();
    }
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
    beginDracoPerfLog();
    const tRefreshStart = performance.now(); // TOTAL de extremo a extremo, ver más abajo
    let tPerf = tRefreshStart; // instrumentación de tiempos, ver draco_perfMark

    await draco_beginManualCalculation();
    tPerf = draco_perfMark(reportId, "Poner cálculo en Manual", tPerf);

    // PLANIFICACIÓN > ver comentario igual en actualizarInformeFixedCore.
    await flushDracoPlanningModifiedCells();
    tPerf = draco_perfMark(reportId, "Vaciar cambios de planificación pendientes", tPerf);

    if (window.BusyIndicator) await window.BusyIndicator.show(busyStepLabel(reportId, "Generando SQL"));
    tPerf = draco_perfMark(reportId, "Mostrar indicador (crear forma)", tPerf);

    // Ver comentario junto a DracoSuppressPlanningPaintCount.
    await beginSuppressPlanningPaint();
    tPerf = draco_perfMark(reportId, "Candado cruzado de planificación (begin)", tPerf);
    try {

    await Excel.run(async (context) => {
        let tSub = performance.now(); // sub-marcas dentro de "Generando SQL"
        const editReportGrid = await getEditReportGrid(context, reportId);
        tSub = draco_perfMark(reportId, "  Leer EDIT_REPORT (usedRange)", tSub);

        const relGrid = await window.SemanticModelStore.getModelGrid("MODEL_RELATIONSHIP");
        const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");
        const atributesGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
        tSub = draco_perfMark(reportId, "  Leer modelo semántico (en memoria, sin Excel)", tSub);

        loadReportDefinition(editReportGrid, reportId);
        applyFilterMeasuresToDynamicReport();

        // EDIT_REPORT!D4 = "Mostrar subtotales arriba" (Propiedades del
        // informe). Solo tiene efecto real cuando además hay algún campo
        // marcado con subtotal en L/R (ver hasAnySubtotalMarked/buildSQL).
        const subtotalsOnTop = String(cellValue(editReportGrid, 4, 4)).trim().toUpperCase() === "X";

        sql = buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop);
        tSub = draco_perfMark(reportId, "  loadReportDefinition + buildSQL (JS puro)", tSub);

        const queryText = typeof sql === "string" ? sql : sql.text;
        console.log("BuildSQL ->", queryText);

        // [Punto 8] SQL ya no se escribe en A1 de la hoja de resultados:
        // se escribe en EDIT_REPORT!X1.
        const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
        editReportSheet.getRange("X1").values = [[draco_debugCell(queryText)]];

        await context.sync();
        draco_perfMark(reportId, "  Escribir X1 (sync)", tSub);
    });
    tPerf = draco_perfMark(reportId, "Generando SQL", tPerf);

    if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(reportId, "Ejecutando consulta"));
    const json = await executeSQL(sql);
    tPerf = draco_perfMark(reportId, "Ejecutando consulta (red)", tPerf);

    console.log("JSON de BigQuery ->", json);

    await Excel.run(async (context) => {
        const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
        const EXCEL_CELL_CHAR_LIMIT = 32000; // límite real de Excel: 32767
        const jsonForCell = json.length > EXCEL_CELL_CHAR_LIMIT
            ? json.substring(0, EXCEL_CELL_CHAR_LIMIT) + " ...(truncado, JSON completo en la consola F12)"
            : json;
        // [Punto 8] JSON ya no se escribe en B1 de la hoja de resultados:
        // se escribe en EDIT_REPORT!Y1.
        editReportSheet.getRange("Y1").values = [[draco_debugCell(jsonForCell)]];
        await context.sync();
    });
    tPerf = draco_perfMark(reportId, "Escribir SQL/JSON en EDIT_REPORT", tPerf);

    if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(reportId, "Pintando resultados"));
    await Excel.run(async (context) => {
        await jsonTo3Matrices(context, json, reportId);
    });
    tPerf = draco_perfMark(reportId, "Pintando resultados (total, ver desglose arriba)", tPerf);

    } finally {
        await endSuppressPlanningPaint();
        tPerf = draco_perfMark(reportId, "Candado cruzado de planificación (end)", tPerf);
        await draco_endManualCalculation();
        tPerf = draco_perfMark(reportId, "Restaurar cálculo automático", tPerf);
        if (window.BusyIndicator) await window.BusyIndicator.hide();
        draco_perfMark(reportId, "Ocultar indicador (borrar forma)", tPerf);
        draco_perfMark(reportId, "══ TOTAL refresco completo ══", tRefreshStart);
        await endDracoPerfLog();
    }
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
/* =========================================================================
 * INFORMES VINCULADOS
 *
 * Un informe vinculado comparte con su informe PRINCIPAL el eje de Filas
 * (se pinta a su derecha) o el de Columnas (se pinta debajo), y hereda sus
 * filtros del botón «Añadir filtro» y las variables. Sus filtros de diseño,
 * sus medidas y el otro eje son propios (p. ej. el principal filtra 2025 y el
 * vinculado 2026), y se planifica por separado, como un informe más.
 *
 *   design.link = { mainId, axis: "rows" | "cols" }   (en el VINCULADO)
 *
 * El eje compartido solo se pinta en el principal. Como las filas (o
 * columnas) de cada informe pueden salir distintas, al refrescar se ejecutan
 * todas las consultas del grupo, se fusionan sus combinaciones
 * (dracoMergeLinkedResults: orden alfabético dentro de su padre) y se pinta
 * primero el principal y después cada vinculado, pegado a lo ya pintado.
 * Refrescar cualquiera de ellos refresca el grupo entero.
 *
 * De momento solo informes DINÁMICOS (sin filas ni columnas fijas).
 * Botón: crearInformeVinculado (registrado con Office.actions.associate;
 * basta con añadir el control en el manifest).
 * ======================================================================= */
let DracoLinkPaint = null; // { reportId, axis, row, col } solo mientras se pinta un vinculado (ver jsonTo3MatricesCore)

function dracoLinkOf(reportId) {
    if (!window.ReportStore || reportId === null || reportId === undefined) return null;
    const r = window.ReportStore.getReport(reportId);
    const l = r && r.design && r.design.link;
    if (!l || l.mainId === undefined || l.mainId === null) return null;
    // si el principal ya no existe (o apunta a sí mismo), deja de estar vinculado
    if (Number(l.mainId) === Number(reportId) || !window.ReportStore.getReport(l.mainId)) return null;
    return { mainId: Number(l.mainId), axis: l.axis === "cols" ? "cols" : "rows" };
}
function dracoLinkMainId(reportId) { const l = dracoLinkOf(reportId); return l ? l.mainId : null; }
/** Vinculados de un principal, en orden de creación (= orden en que se pintan). */
function dracoLinkedOf(mainId) {
    if (!window.ReportStore || mainId === null || mainId === undefined) return [];
    return window.ReportStore.listReports().map(r => r.id)
        .filter(id => { const l = dracoLinkOf(id); return !!l && l.mainId === Number(mainId); });
}
/** Principal del grupo al que pertenece un informe (él mismo si es principal), o null. */
function dracoLinkGroupMain(reportId) {
    const m = dracoLinkMainId(reportId);
    if (m !== null) return m;
    return dracoLinkedOf(reportId).length ? Number(reportId) : null;
}

/** Campos del eje compartido, copiados del principal (incluidas sus opciones de campo). */
function dracoLinkedAxisFields(main, linked, axis) {
    const md = (main && main.design) || {}, ld = (linked && linked.design) || {};
    const zone = axis === "cols" ? "columns" : "rows";
    const copy = x => JSON.parse(JSON.stringify(x || []));
    const fo = {};
    Object.keys(ld.fieldOptions || {}).forEach(k => { if (k.indexOf(zone + "|") !== 0) fo[k] = ld.fieldOptions[k]; });
    Object.keys(md.fieldOptions || {}).forEach(k => { if (k.indexOf(zone + "|") === 0) fo[k] = md.fieldOptions[k]; });
    const out = { fieldOptions: fo, resultSheetName: md.resultSheetName || ld.resultSheetName || "" };
    if (axis === "cols") { out.columns = copy(md.columns); out.expandedCols = copy(md.expandedCols); out.colsStatic = !!md.colsStatic; }
    else { out.rows = copy(md.rows); out.expandedRows = copy(md.expandedRows); out.rowsStatic = !!md.rowsStatic; }
    return out;
}
/** El vinculado sigue siempre el eje del principal (si se cambia en el principal, se cambia en él). */
async function dracoSyncLinkedDesigns(mainId, linkedIds) {
    const main = window.ReportStore.getReport(mainId);
    for (const id of linkedIds) {
        const l = dracoLinkOf(id), linked = window.ReportStore.getReport(id);
        if (!l || !linked) continue;
        const fields = dracoLinkedAxisFields(main, linked, l.axis);
        const cur = {};
        Object.keys(fields).forEach(k => { cur[k] = (linked.design || {})[k]; });
        if (JSON.stringify(cur) !== JSON.stringify(fields)) await window.ReportStore.updateDesignFields(id, fields);
    }
}
/** El vinculado usa el MISMO estado de +/- que el principal en el eje compartido (mismas filas visibles). */
function dracoShareCollapseState(mainId, linkedId, axis) {
    const k = axis === "cols" ? "cols" : "rows";
    getDracoCollapseState(linkedId)[k] = getDracoCollapseState(mainId)[k];
}

/**
 * INFORMES VINCULADOS — fusión de los resultados de un grupo (principal + vinculados)
 * que comparten el eje de Filas ("rows") o de Columnas ("cols").
 *
 * parts[i] = { json, nRowDims, nColDims } (i = 0 es el principal). Cada json es el
 * pseudo-JSON de BigQuery que devuelve executeSQL: rows[].f[].v con los campos
 * [ROW_ID, COLUMN_ID, dims de Filas..., dims de Columnas..., medidas...].
 *
 * La posición de cada fila/columna en la hoja sale SOLO de ROW_ID/COLUMN_ID
 * (DENSE_RANK sobre los campos del eje), así que fusionar = unir las combinaciones
 * del eje compartido de todos los informes, ordenarlas y renumerarlas IGUAL en
 * todos, y añadir a cada informe (con medidas vacías) las combinaciones que no
 * tenga. Así el pintado de siempre las coloca en la misma fila/columna física.
 *
 * Orden: alfabético (natural: "2" antes que "10") dentro de su padre (el orden
 * es por niveles, de izquierda a derecha). Se respeta lo que ya hace el SQL del
 * principal: un nivel ordenado descendente sigue descendente, y los nodos /
 * subtotales (null o "TOTAL") van delante o detrás según salían. Si no se puede
 * fusionar (JSON ilegible, ejes distintos), devuelve los JSON tal cual.
 */
function dracoMergeLinkedResults(parts, axis) {
    const out = parts.map(p => p.json);
    if (!parts || parts.length < 2) return out;
    const idIdx = axis === "cols" ? 1 : 0;
    const from = p => (axis === "cols" ? 2 + p.nRowDims : 2);
    const count = p => (axis === "cols" ? p.nColDims : p.nRowDims);
    const n = count(parts[0]);
    if (!(n > 0) || parts.some(p => count(p) !== n)) return out;
    const parsed = parts.map(p => { try { const o = JSON.parse(p.json); return o && typeof o === "object" ? o : null; } catch (e) { return null; } });
    if (parsed.some(o => !o)) return out;

    const val = c => (c && c.v !== undefined && c.v !== null ? c.v : null);
    const kind = v => (v === null || v === "null" ? "N" : (String(v).toUpperCase() === "TOTAL" ? "T" : ""));
    const coll = (typeof Intl !== "undefined" && Intl.Collator) ? new Intl.Collator("es", { numeric: true, sensitivity: "base" }) : null;
    const cmpText = (x, y) => { const a = String(x), b = String(y); const c = coll ? coll.compare(a, b) : 0; return c || (a < b ? -1 : a > b ? 1 : 0); };
    const keyOf = (f, p) => JSON.stringify(f.slice(from(p), from(p) + n).map(val));

    // Combinaciones y su ID original en cada informe
    const tuples = new Map(); // key -> { vals, ids[] }
    parsed.forEach((o, pi) => (o.rows || []).forEach(r => {
        const f = r.f || [];
        const k = keyOf(f, parts[pi]);
        let t = tuples.get(k);
        if (!t) { t = { vals: JSON.parse(k), ids: [] }; tuples.set(k, t); }
        const id = Number(val(f[idIdx]));
        if (isFinite(id) && (t.ids[pi] === undefined || id < t.ids[pi])) t.ids[pi] = id;
    }));

    // Deducir del SQL (principal primero) el sentido de cada nivel y dónde van nulos/TOTAL
    const dirVote = new Array(n).fill(0), spVote = { N: 0, T: 0 };
    parsed.forEach((o, pi) => {
        const seq = [...tuples.values()].filter(t => t.ids[pi] !== undefined).sort((a, b) => a.ids[pi] - b.ids[pi]);
        const w = pi === 0 ? 1000 : 1; // manda el principal
        for (let i = 1; i < seq.length; i++) {
            const a = seq[i - 1].vals, b = seq[i].vals;
            let d = 0; while (d < n && a[d] === b[d]) d++;
            if (d >= n) continue;
            const ka = kind(a[d]), kb = kind(b[d]);
            if (ka && !kb) spVote[ka] += w;          // especial delante
            else if (kb && !ka) spVote[kb] -= w;     // especial detrás
            else if (!ka && !kb) dirVote[d] += w * Math.sign(cmpText(b[d], a[d]));
        }
    });
    const spFirst = { N: spVote.N >= 0, T: spVote.T > 0 }; // sin pistas: nodos (null) delante, "TOTAL" detrás
    const cmp = (A, B) => {
        for (let i = 0; i < n; i++) {
            const x = A.vals[i], y = B.vals[i];
            if (x === y) continue;
            const kx = kind(x), ky = kind(y);
            if (kx || ky) {
                if (kx && ky) return kx === ky ? 0 : (kx === "N" ? -1 : 1);
                return kx ? (spFirst[kx] ? -1 : 1) : (spFirst[ky] ? 1 : -1);
            }
            const c = cmpText(x, y);
            if (c) return dirVote[i] < 0 ? -c : c;
        }
        return 0;
    };
    const all = [...tuples.entries()].sort((a, b) => cmp(a[1], b[1]));
    const newId = new Map(all.map(([k], i) => [k, String(i + 1)]));

    return parsed.map((o, pi) => {
        const p = parts[pi];
        const rows = (o.rows || []).map(r => {
            const f = (r.f || []).slice();
            f[idIdx] = { v: newId.get(keyOf(f, p)) };
            return Object.assign({}, r, { f });
        });
        // Combinaciones que no tiene este informe: fila con medidas vacías (con la
        // otra coordenada de su primera fila, para no inventar columnas/filas nuevas)
        const tpl = rows[0];
        if (tpl) {
            const mFrom = 2 + p.nRowDims + p.nColDims;
            for (const [k, t] of all) {
                if (t.ids[pi] !== undefined) continue;
                const f = tpl.f.map((c, j) => (j >= mFrom ? { v: null } : c));
                f[idIdx] = { v: newId.get(k) };
                t.vals.forEach((v, j) => { f[from(p) + j] = { v }; });
                rows.push({ f });
            }
        }
        // Mismo texto que executeSQL ('{"v": ...}', con espacio): el pintado (jsonTo3MatricesCore)
        // no usa JSON.parse, recorre el texto buscando '"v":' y cuenta con ese espacio.
        return JSON.stringify(Object.assign({}, o, { rows, totalRows: String(rows.length) })).replace(/"v":/g, '"v": ');
    });
}

/** Borra lo pintado por los vinculados (y sus rangos con nombre) ANTES de repintar el principal,
 *  que puede crecer y ocupar su sitio: si no, al repintarse después borrarían celdas del principal. */
async function dracoClearLinkedAreas(context, ids) {
    for (const id of ids) {
        const rep = window.ReportStore ? window.ReportStore.getReport(id) : null;
        const props = (rep && rep.reportProperties) || {};
        await clearDracoNamedRanges(context, id, props.overwriteFormats !== false);
        const rn = dracoRangeNames(id);
        const items = [rn.rows, rn.cols, rn.values].map(n => context.workbook.names.getItemOrNullObject(n));
        items.forEach(it => it.load("isNullObject"));
        await context.sync();
        let any = false;
        items.forEach(it => { if (!it.isNullObject) { it.delete(); any = true; } });
        if (any) await context.sync();
    }
}
/** Rangos pintados (rows/cols/values) de unos informes, con su posición. */
async function dracoPaintedRanges(context, ids) {
    const items = [];
    for (const id of ids) {
        const rn = dracoRangeNames(id);
        [["rows", rn.rows], ["cols", rn.cols], ["values", rn.values]].forEach(([kind, n]) =>
            items.push({ id, kind, it: context.workbook.names.getItemOrNullObject(n) }));
    }
    items.forEach(x => x.it.load("isNullObject"));
    await context.sync();
    const live = items.filter(x => !x.it.isNullObject).map(x => {
        const r = x.it.getRange();
        r.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
        return Object.assign(x, { r });
    });
    await context.sync();
    return live;
}
/** Pinta cada vinculado pegado a lo ya pintado (principal + vinculados anteriores), con una
 *  columna (o fila) vacía de separación y alineado con el eje del principal. */
async function dracoPaintLinkedChain(context, mainId, linkedIds, jsons) {
    for (let i = 0; i < linkedIds.length; i++) {
        const id = linkedIds[i], json = jsons[i], l = dracoLinkOf(id);
        if (!json || !l) continue;
        const live = await dracoPaintedRanges(context, [mainId, ...linkedIds.slice(0, i)]);
        const mainOf = k => { const x = live.find(y => y.id === mainId && y.kind === k); return x ? x.r : null; };
        let row, col;
        if (l.axis === "rows") {
            const base = mainOf("rows") || mainOf("values");
            if (!base) continue;                                   // el principal no tiene filas: nada con lo que alinear
            row = base.rowIndex + 1;
            col = Math.max(...live.map(x => x.r.columnIndex + x.r.columnCount)) + 2;
        } else {
            const base = mainOf("cols") || mainOf("values");
            if (!base) continue;
            col = base.columnIndex + 1;
            row = Math.max(...live.map(x => x.r.rowIndex + x.r.rowCount)) + 2;
        }
        dracoShareCollapseState(mainId, id, l.axis);
        DracoLinkPaint = { reportId: id, axis: l.axis, row, col };
        try { await jsonTo3Matrices(context, json, id); }
        finally { DracoLinkPaint = null; }
    }
}
/** SQL de un informe dinámico (lo mismo que hace actualizarInformeCore) y el tamaño de sus ejes. */
async function dracoBuildDynamicQuery(reportId) {
    let sql = null, nRowDims = 0, nColDims = 0, empty = false;
    await Excel.run(async (context) => {
        const editReportGrid = await getEditReportGrid(context, reportId);
        const relGrid = await window.SemanticModelStore.getModelGrid("MODEL_RELATIONSHIP");
        const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");
        const atributesGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
        loadReportDefinition(editReportGrid, reportId);
        applyFilterMeasuresToDynamicReport();
        const subtotalsOnTop = String(cellValue(editReportGrid, 4, 4)).trim().toUpperCase() === "X";
        sql = buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop);
        // OJO: en ReportState, "Columns" son las Filas de la hoja y "Rows" las Columnas (herencia del VBA)
        nRowDims = ReportState.ColumnCount;
        nColDims = ReportState.RowCount;
        empty = !(ReportState.Measures && ReportState.Measures.length);
    });
    return { sql, nRowDims, nColDims, empty };
}
/** Refresca un grupo vinculado (sin indicador ni cálculo manual: eso lo pone quien llama). */
async function dracoRefreshLinkGroup(mainId) {
    const linkedIds = dracoLinkedOf(mainId);
    const ids = [mainId, ...linkedIds];
    for (const id of ids) {
        const mode = await Excel.run(async (context) => dracoReportMode(await getEditReportGrid(context, id)));
        if (mode !== "dynamic") {
            const r = window.ReportStore.getReport(id);
            throw new Error(`Los informes vinculados funcionan, de momento, solo con informes dinámicos: «${r ? r.name : id}» tiene ${mode === "fixed" ? "filas y columnas fijas" : "un eje fijo"}.`);
        }
    }
    await dracoSyncLinkedDesigns(mainId, linkedIds);
    const parts = [], painted = [];
    for (const id of ids) {
        if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(id, "Generando SQL"));
        const q = await dracoBuildDynamicQuery(id);
        // Un vinculado recién creado (aún sin medidas) no se consulta ni se pinta: si no, su consulta
        // vacía haría fallar el refresco de todo el grupo, también el del principal.
        if (id !== mainId && q.empty) continue;
        if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(id, "Ejecutando consulta"));
        q.json = await executeSQL(q.sql);
        parts.push(q);
        if (id !== mainId) painted.push(id);
    }
    const axis = linkedIds.length ? dracoLinkOf(linkedIds[0]).axis : "rows";
    const merged = painted.length ? dracoMergeLinkedResults(parts, axis) : parts.map(p => p.json);
    if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(mainId, "Pintando resultados"));
    await Excel.run(async (context) => {
        await dracoClearLinkedAreas(context, linkedIds);
        await jsonTo3Matrices(context, merged[0], mainId);
        await dracoPaintLinkedChain(context, mainId, painted, merged.slice(1));
    });
}
async function actualizarGrupoVinculado(mainId) {
    await draco_beginManualCalculation();
    await flushDracoPlanningModifiedCells();
    if (window.BusyIndicator) await window.BusyIndicator.show(busyStepLabel(mainId, "Informes vinculados"));
    await beginSuppressPlanningPaint();
    try {
        await dracoRefreshLinkGroup(mainId);
    } finally {
        await endSuppressPlanningPaint();
        await draco_endManualCalculation();
        if (window.BusyIndicator) await window.BusyIndicator.hide();
    }
}

/** Crea un informe vinculado a mainId por su eje de Filas ("rows") o Columnas ("cols"). Devuelve su id. */
async function dracoCreateLinkedReport(mainId, axis) {
    axis = axis === "cols" ? "cols" : "rows";
    const main = window.ReportStore.getReport(mainId);
    if (!main) throw new Error("No se encuentra el informe principal.");
    if (dracoLinkOf(mainId)) throw new Error("Ese informe ya es un vinculado: elige su informe principal.");
    const prev = dracoLinkedOf(mainId);
    if (prev.length && dracoLinkOf(prev[0]).axis !== axis) {
        throw new Error(`Los vinculados de «${main.name}» comparten sus ${dracoLinkOf(prev[0]).axis === "cols" ? "columnas" : "filas"}: todos los de un mismo principal tienen que vincularse por el mismo eje.`);
    }
    const sheet = (main.design && main.design.resultSheetName) || "";
    const rep = await window.ReportStore.createReport(main.semanticModelName || "", sheet);   // queda como informe activo
    await window.ReportStore.updateDesignFields(rep.id, Object.assign({ link: { mainId: Number(mainId), axis } },
        dracoLinkedAxisFields(main, rep, axis)));
    if (window.ReportStore.renameReport) await window.ReportStore.renameReport(rep.id, `${main.name} · vinculado ${prev.length + 1}`);
    return rep.id;
}

/**
 * Comando de cinta "crearInformeVinculado": pide el informe principal y el eje a
 * vincular, crea el vinculado y abre el diseñador para que se le añadan sus
 * medidas, columnas (o filas) y filtros propios. Para activarlo basta con añadir
 * al manifest un botón con <FunctionName>crearInformeVinculado</FunctionName>.
 */
let dracoLinkDialogOpen = false;
async function crearInformeVinculado(event) {
    if (dracoLinkDialogOpen || !window.ReportStore) { if (event) event.completed(); return; }
    dracoLinkDialogOpen = true;
    try {
        const reports = window.ReportStore.listReports()
            .filter(r => !dracoLinkOf(r.id))
            .map(r => {
                const linked = dracoLinkedOf(r.id);
                return { id: r.id, name: r.name, lockedAxis: linked.length ? dracoLinkOf(linked[0]).axis : null, linkedCount: linked.length };
            });
        const url = new URL("linkDialog.html", window.location.href).href;
        let createdId = null;
        await new Promise((resolve) => {
            Office.context.ui.displayDialogAsync(url, { height: 55, width: 35, displayInIframe: false }, (res) => {
                if (res.status === Office.AsyncResultStatus.Failed) { console.error("Informe vinculado: displayDialogAsync:", res.error && res.error.message); resolve(); return; }
                const dialog = res.value;
                const reply = (o) => { try { dialog.messageChild(JSON.stringify(o)); } catch (e) { console.error(e); } };
                const done = () => { try { dialog.close(); } catch (e) { /* ya cerrado */ } resolve(); };
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                    let p;
                    try { p = JSON.parse(arg.message); } catch (e) { return; }
                    if (p.type === "ready") { reply({ type: "init", reports, active: activeReportIdOrNull() }); return; }
                    if (p.type === "create") {
                        try { createdId = await dracoCreateLinkedReport(Number(p.mainId), p.axis); done(); }
                        catch (e) { reply({ type: "error", message: (e && e.message) || String(e) }); }
                        return;
                    }
                    if (p.type === "close") done();
                });
                dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve());
            });
        });
        if (createdId !== null) {
            try {
                await Office.addin.showAsTaskpane();
                if (window.mostrarVistaEpm) window.mostrarVistaEpm("view-informes");
            } catch (e) { console.warn("[Draco] Informe vinculado creado; no se pudo abrir el diseñador:", e); }
        }
    } catch (error) {
        console.error("Error al crear el informe vinculado:", error);
        await surfaceErrorToSheet(error);
    } finally {
        dracoLinkDialogOpen = false;
        if (event) event.completed();
    }
}

async function actualizarUnInforme(reportIdOverride) {
    const reportId = reportIdOverride !== undefined ? reportIdOverride : activeReportIdOrNull();
    // Informes vinculados: refrescar uno refresca el grupo (primero el principal)
    const groupMain = dracoLinkGroupMain(reportId);
    if (groupMain !== null) return await actualizarGrupoVinculado(groupMain);
    // Tres modos (ver dracoReportMode): Fijo, Híbrido o Dinámico.
    const mode = await Excel.run(async (context) => {
        const grid = await getEditReportGrid(context, reportId);
        return dracoReportMode(grid);
    });

    if (mode === "fixed") {
        await actualizarInformeFixedCore(reportId);
    } else if (mode === "hybrid") {
        await actualizarInformeHybridCore(reportId);
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
 * lanzar, p.ej., 20 consultas a BigQuery a la vez sin control).
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
 *   2) Lanza las consultas a BigQuery EN PARALELO (con un tope
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
    const tRefreshStart = performance.now(); // TOTAL de extremo a extremo, ver más abajo

    // Una sola vez para TODO el lote (ver comentario en
    // draco_beginManualCalculation): cada informe individual, dentro del
    // bucle, puede seguir llamando a su propio begin/end sin que se
    // restaure a mitad de camino (profundidad, igual que BusyIndicator).
    await draco_beginManualCalculation();

    // PLANIFICACIÓN > ver comentario igual en actualizarInformeFixedCore.
    // (actualizarInformeFixedCore también lo llama por su cuenta más
    // abajo, una vez por cada informe "Fijo" de este mismo "Actualizar
    // todos"; a partir de la primera llamada el registro ya está vacío,
    // así que esas llamadas repetidas no hacen nada.)
    await flushDracoPlanningModifiedCells();

    // Un único show()/hide() para TODO el lote: como show()/hide() llevan
    // refCount, cada informe individual (actualizarInformeFixedCore, o el
    // jsonTo3Matrices de los dinámicos, más abajo) puede seguir mostrando
    // el suyo sin que el indicador desaparezca entre uno y otro — solo se
    // quita de verdad cuando ESTE hide() (el último en llamarse) lleva el
    // contador a 0.
    if (window.BusyIndicator) await window.BusyIndicator.show("Actualizando todos");
    beginDracoPerfLog();

    // Ver comentario junto a DracoSuppressPlanningPaintCount: cubre TODO
    // "Refrescar todos" (dinámicos + fijos, que a su vez llama a
    // actualizarInformeFixedCore y suma su propia cuenta al contador).
    await beginSuppressPlanningPaint();
    try {

    const dynamicJobs = [];
    const fixedReports = [];
    const linkGroups = [];   // principales con informes vinculados: se refrescan al final, como grupo

    // 1) Clasificar cada informe (Fijo/Dinámico, según H12/N12 de su
    //    propio grid) y, para los Dinámicos, construir ya su SQL: es
    //    síncrono/local (no toca red), así que se hace en serie sin
    //    problema — no hay motivo para paralelizar esta parte.
    for (const r of reports) {
        try {
            const reportId = r.id;
            if (dracoLinkMainId(reportId) !== null) continue;                     // vinculado: va con su principal
            if (dracoLinkedOf(reportId).length) { linkGroups.push(reportId); continue; }
            if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(reportId, "Clasificando informe"));
            const mode = await Excel.run(async (context) => {
                const grid = await getEditReportGrid(context, reportId);
                return dracoReportMode(grid);
            });

            if (mode === "fixed" || mode === "hybrid") {
                fixedReports.push({ reportId, reportName: r.name, mode });
            } else {
                let sql;
                if (window.BusyIndicator) await window.BusyIndicator.update(busyStepLabel(reportId, "Generando SQL"));
                await Excel.run(async (context) => {
                    const editReportGrid = await getEditReportGrid(context, reportId);
                    const relGrid = await window.SemanticModelStore.getModelGrid("MODEL_RELATIONSHIP");
                    const measuresGrid = await window.SemanticModelStore.getModelGrid("MODEL_MEASURES");
                    const atributesGrid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
                    loadReportDefinition(editReportGrid, reportId);
                    applyFilterMeasuresToDynamicReport();
                    const subtotalsOnTop = String(cellValue(editReportGrid, 4, 4)).trim().toUpperCase() === "X";
                    sql = buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop);
                });
                dynamicJobs.push({ reportId, reportName: r.name, sql });
            }
        } catch (err) {
            console.error(`[Draco] Error preparando el informe "${r.name}" (id ${r.id}):`, err);
        }
    }

    // 2) Consultas de los informes "Dinámicos" a BigQuery EN
    //    PARALELO (tope de concurrencia por defecto 4): es la parte lenta
    //    (red, sin tocar Excel) y cada informe es independiente. Al ir
    //    varias a la vez, el indicador no puede mostrar un "en curso" por
    //    cada una a la vez (un solo shape, un solo texto): se actualiza al
    //    EMPEZAR cada consulta (así se ve cuál acaba de arrancar) y se
    //    lleva la cuenta de cuántas han terminado ya, sin bajar la
    //    concurrencia para conseguirlo.
    let queriesDone = 0;
    const totalQueries = dynamicJobs.length;
    await runWithConcurrency(dynamicJobs, concurrency || 4, async (job) => {
        if (window.BusyIndicator) {
            await window.BusyIndicator.update(
                busyStepLabel(job.reportId, "Ejecutando consulta") + ` (${queriesDone}/${totalQueries} hechas)`
            );
        }
        const tJob = performance.now(); // tiempo propio de ESTE informe, no encadenado (van en paralelo)
        try {
            job.json = await executeSQL(job.sql);
            draco_perfMark(job.reportId, "Ejecutando consulta (red, en paralelo)", tJob);
        } catch (err) {
            job.error = err;
            console.error(`[Draco] Error consultando datos del informe "${job.reportName}":`, err);
        } finally {
            queriesDone++;
        }
    });

    // 3) Pintar los "Dinámicos" en Excel EN SERIE: la API de Excel no
    //    admite escrituras concurrentes seguras sobre el mismo libro.
    let paintedCount = 0;
    for (const job of dynamicJobs) {
        if (job.error || !job.json) continue;
        try {
            if (window.BusyIndicator) {
                await window.BusyIndicator.update(
                    busyStepLabel(job.reportId, "Pintando resultados") + ` (${paintedCount + 1}/${totalQueries})`
                );
            }
            await Excel.run(async (context) => {
                const editReportSheet = context.workbook.worksheets.getItem("EDIT_REPORT");
                const EXCEL_CELL_CHAR_LIMIT = 32000;
                const jsonForCell = job.json.length > EXCEL_CELL_CHAR_LIMIT
                    ? job.json.substring(0, EXCEL_CELL_CHAR_LIMIT) + " ...(truncado, JSON completo en la consola F12)"
                    : job.json;
                editReportSheet.getRange("Y1").values = [[draco_debugCell(jsonForCell)]];
                await jsonTo3Matrices(context, job.json, job.reportId);
            });
            paintedCount++;
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
            if (f.mode === "hybrid") await actualizarInformeHybridCore(f.reportId);
            else await actualizarInformeFixedCore(f.reportId);
        } catch (err) {
            console.error(`[Draco] Error actualizando el informe "${f.reportName}" (modo ${f.mode === "hybrid" ? "híbrido" : "fijo"}):`, err);
        }
    }


    // Informes vinculados: cada principal con sus vinculados, como grupo
    for (const mainId of linkGroups) {
        try {
            await dracoRefreshLinkGroup(mainId);
        } catch (err) {
            console.error(`[Draco] Error actualizando los informes vinculados del informe ${mainId}:`, err);
        }
    }

    } finally {
        await endSuppressPlanningPaint();
        await draco_endManualCalculation();
        if (window.BusyIndicator) await window.BusyIndicator.hide();
        draco_perfMark(null, "══ TOTAL 'Actualizar todos' completo ══", tRefreshStart);
        await endDracoPerfLog();
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
    convertAxisStaticFormulas, repaintAxisIfHiddenIds, ensureDracoHandlersRegistered,
    // "Añadir filtro" (rango con nombre sobre una celda cualquiera del libro)
    dracoFilterRangeName,
    resolveDracoFilterRangeAddress, openDracoFilterRangePicker,
    // Registro del clic en Draco_XXX_Rows -> escribe zona + Jerarquia SI/NO en A51
    ensureDracoRowsClickLoggerRegistered,
    // Averiguar a qué informe pertenece una hoja (ver "seguir informe
    // activo según la selección" en taskpane.js).
    reportIdForResultSheet
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
 * PLANIFICACIÓN > Guardar (botón del ribbon "PlanifGuardarButton"): de
 * momento es un dummy para que el manifest no falle al pulsar el botón
 * (Office exige que toda <Action xsi:type="ExecuteFunction"> tenga una
 * función registrada con Office.actions.associate, o el ribbon lanza un
 * error al hacer clic). Sustituir por la lógica real (INSERT con los
 * cambios marcados en el informe de planificación, revisión en
 * EDIT_REPORT!V4 y ejecución) cuando esté lista.
 * ------------------------------------------------------------------- */
// Tabla legible (informe | celda | color anterior | cruce | medida |
// filtros) que escribe guardarPlanificacion justo antes de vaciar el
// registro — a diferencia del JSON de A127 (que también se escribe en
// cada refresco, pensado para que lo lea el XLAM), esta tabla es solo
// para "Guardar" y en formato humano: una fila por celda modificada.
//
// El "cruce" (qué Dimensión/Valor corresponde a esa celda en cada eje)
// sale de readRowDefinitions/readColumnDefinitions — las mismas que ya
// usa buildSQLFixed para leer las fórmulas EPM_VALUE pintadas en las
// cabeceras de fila/columna: para una celda de Draco_<id>_Values, su fila
// física coincide con la fila de sus cabeceras de FILA, y su columna
// física con la columna de sus cabeceras de COLUMNA, así que basta
// filtrar esas dos listas por R === fila / R === columna de la celda.
const DRACO_PLANNING_SAVED_TABLE_CELL = "A130"; // A130:F... hacia abajo

// Texto legible ("valor1/valor2" o "NO valor1/valor2") de un filtro
// guardado, reutilizando el mismo parseo que ya usa buildFilterConditions.
function summarizeDracoFilterValue(rawValue) {
    const filter = parseStoredFilterValue(rawValue);
    if (!filter) return String(rawValue || "");

    const list = (filter.items && filter.items.length) ? filter.items
        : (filter.excludeItems && filter.excludeItems.length) ? filter.excludeItems
        : null;
    if (list) {
        const texto = list.map(it => (it && (it.display || it.value)) || it).join("/");
        return (filter.excludeItems && filter.excludeItems.length && !filter.items) ? "NO " + texto : texto;
    }
    if (filter.values && filter.values.length) {
        return (filter.include === false ? "NO " : "") + filter.values.join("/");
    }
    return String(rawValue || "");
}

// Para un informe concreto: {rowDefs, colDefs, measureLabel, filtersText}
// — una sola vez por informe (no por celda), reutilizado luego para cada
// celda modificada de ESE informe.
async function loadDracoPlanningCrossContext(context, reportId) {
    const editReportGrid = await getEditReportGrid(context, reportId);
    const resultSheetName = resultSheetNameFromGrid(editReportGrid, reportId);
    const csvGrid = await getFormulaGrid(context, resultSheetName);

    loadReportDefinition(editReportGrid, reportId); // rellena ReportState (incluye Filters)

    // Variables ($X$) -> su valor real: la planificación escribe con él.
    const resolveVars = (defs) => window.DracoVariables
        ? defs.map(d => Object.assign({}, d, { Value: window.DracoVariables.resolveValue(d.Value) }))
        : defs;
    const rowDefs = resolveVars(await readRowDefinitions(context, editReportGrid, csvGrid));
    const colDefs = resolveVars(await readColumnDefinitions(context, editReportGrid, csvGrid));

    const isMeasureDef = (d) => String(d.Dimension).toUpperCase() === "MEASURE";
    const measureDef = rowDefs.find(isMeasureDef) || colDefs.find(isMeasureDef);
    const measureLabel = measureDef ? (measureDef.Display || measureDef.AttributeName || "") : "";

    const filtersText = (ReportState.Filters || [])
        .filter(f => String(f.Value).trim() !== "")
        .map(f => `${f.Dimension}=${summarizeDracoFilterValue(f.Value)}`)
        .join(", ");

    return {
        rowDefs: rowDefs.filter(d => !isMeasureDef(d)),
        colDefs: colDefs.filter(d => !isMeasureDef(d)),
        measureLabel,
        filtersText
    };
}

// Items estructurados {dim, attr, value} de una celda concreta (Estático:
// hay fórmulas EPM_VALUE que leer). dracoPlanningCrossTextForCell (más
// abajo) es solo un formateador de esto a texto — la validación de
// dimensiones obligatorias usa esta versión estructurada directamente.
//
// options.preferCanonical: por defecto (false/omitido) se prioriza Display
// (la etiqueta pensada para MOSTRAR al usuario, p.ej. "01/08/2025" con el
// formato de fecha de su locale) — así se ha comportado siempre esta
// función, usada en diagnósticos/UI. Con preferCanonical:true (usado por
// getCrossItemsForWriteback) se prioriza Value — el valor CANÓNICO que ya
// usa el resto del informe para consultar/unir contra BigQuery (p.ej.
// "2025-08-01"), imprescindible al escribir de vuelta: insertar el Display
// de una fecha en la tabla de hechos hace que luego no la encuentre el
// JOIN contra la dimensión (formato de fecha distinto al guardado).
function dracoPlanningCrossItemsForCell(crossCtx, cellAddr, options) {
    const preferCanonical = !!(options && options.preferCanonical);
    const { row, col } = parseAddress(cellAddr);
    return [
        ...crossCtx.rowDefs.filter(d => d.R === row),
        ...crossCtx.colDefs.filter(d => d.R === col)
    ].map(d => ({
        dim: d.Dimension,
        attr: d.AttributeName,
        value: preferCanonical ? (d.Value || d.Display) : (d.Display || d.Value)
    }));
}

// "Dim X Valor Y, Dim Z Valor W" para una celda concreta, a partir de las
// listas de fila/columna ya cargadas (loadDracoPlanningCrossContext).
// Solo encuentra algo si el eje correspondiente está en Estático (hay
// fórmulas EPM_VALUE que leer) — para Dinámico, ver
// dracoPlanningCrossTextForCellDynamic más abajo.
function dracoPlanningCrossTextForCell(crossCtx, cellAddr) {
    return dracoPlanningCrossItemsForCell(crossCtx, cellAddr)
        .map(it => `${it.dim}.${it.attr}=${it.value}`).join(", ");
}

/**
 * Igual que loadDracoPlanningCrossContext, pero para ejes en DINÁMICO: no
 * hay ninguna fórmula EPM_VALUE que leer (el eje solo tiene el texto ya
 * calculado), así que se usa el MISMO procedimiento que
 * convertAxisStaticFormulas usa para saber "esto es Dimensión X, Valor Y"
 * a partir de una celda de cabecera sin fórmula:
 *   1) iAux = columna física (en Filas) o fila física (en Columnas) del
 *      campo dentro del eje.
 *   2) indentLevel real de ESA celda concreta -> flag (profundidad).
 *   3) fieldLevels[iAux] (construido desde el DISEÑO en EDIT_REPORT, no
 *      depende de si el eje es Dinámico o Estático) + flag ->
 *      resolveAxisFieldForFlag -> {dim, attr}.
 *   4) el VALOR es el propio texto ya pintado en esa celda, sin el glifo
 *      ▸/▾ de expandir/contraer.
 */
async function loadDracoPlanningCrossContextDynamic(context, reportId) {
    const editReportGrid = await getEditReportGrid(context, reportId);
    loadReportDefinition(editReportGrid, reportId); // rellena ReportState (incluye Filters)

    const fieldLevelsRows = buildAxisFieldLevelsTable(editReportGrid, "rows");
    const fieldLevelsCols = buildAxisFieldLevelsTable(editReportGrid, "columns");

    const names = dracoRangeNames(reportId);

    const rowsNameObj = context.workbook.names.getItemOrNullObject(names.rows);
    const rowsRange = rowsNameObj.getRangeOrNullObject();
    rowsRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);

    const colsNameObj = context.workbook.names.getItemOrNullObject(names.cols);
    const colsRange = colsNameObj.getRangeOrNullObject();
    colsRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);

    await context.sync();

    const rowsIndent = rowsRange.isNullObject ? null : await readIndentLevelsForRange(context, rowsRange);
    const colsIndent = colsRange.isNullObject ? null : await readIndentLevelsForRange(context, colsRange);

    const isMeasureDim = (dim) => String(dim).toUpperCase() === "MEASURE";
    const findMeasureLabel = (fieldLevels) => {
        for (let iAux = 1; iAux < fieldLevels.length; iAux++) {
            const col = fieldLevels[iAux];
            if (!col) continue;
            for (const flagKey of Object.keys(col.byFlag)) {
                const entry = col.byFlag[flagKey];
                if (entry && isMeasureDim(entry.dim)) return entry.attr;
            }
        }
        return "";
    };
    const measureLabel = findMeasureLabel(fieldLevelsRows) || findMeasureLabel(fieldLevelsCols);

    const filtersText = (ReportState.Filters || [])
        .filter(f => String(f.Value).trim() !== "")
        .map(f => `${f.Dimension}=${summarizeDracoFilterValue(f.Value)}`)
        .join(", ");

    return { reportId, rowsRange, rowsIndent, fieldLevelsRows, colsRange, colsIndent, fieldLevelsCols, measureLabel, filtersText };
}

const DRACO_GLYPH_PREFIX = /^[▸▾]\s+/;

// Tipos DATA_TYPE (MODEL_ATRIBUTES / esquema BigQuery) que representan una
// fecha/instante — para estos, el eje Dinámico debe convertir el valor
// crudo de Excel (número de serie) a ISO 8601, no usar el texto mostrado
// (que sale formateado según el locale del usuario, p.ej. "01/08/2025",
// y no coincide con lo que hay guardado en la tabla, p.ej. "2025-08-01").
const DRACO_PLANNING_DATE_TYPES = ["DATE", "DATETIME", "TIMESTAMP"];

function dracoPlanningValueIsDateType(dataType) {
    return DRACO_PLANNING_DATE_TYPES.includes(String(dataType || "").trim().toUpperCase());
}

// Convierte un número de serie de fecha de Excel (días desde 1899-12-30,
// incluyendo el "bug" del año bisiesto de 1900 que Excel hereda de Lotus)
// a "YYYY-MM-DD". 25569 = días entre el epoch de Excel y el de Unix
// (1970-01-01) — fórmula estándar, válida para cualquier fecha moderna.
function excelSerialToISODate(serial) {
    const utcMillis = Math.round((Number(serial) - 25569) * 86400 * 1000);
    const d = new Date(utcMillis);
    const yyyy = d.getUTCFullYear();
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const dd = String(d.getUTCDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
}

// Items estructurados {dim, attr, value} — versión Dinámica (sin fórmulas
// EPM_VALUE, lee el texto ya pintado). Mismo papel que
// dracoPlanningCrossItemsForCell pero para el otro caso.
//
// El valor de cada dimensión sale PRIMERO de DracoAxisCrossValues — el
// valor tal como vino en el JSON de BigQuery al último refresco, guardado
// en jsonTo3MatricesCore ANTES de que Excel toque la celda. Así da igual
// lo que Excel decida guardar/mostrar ahí (p.ej. reinterpretar una fecha
// ISO y reformatearla según el locale del usuario): la fuente de verdad
// para el writeback es el JSON, no la celda.
//
// Solo si no hay nada en la caché para esa celda (p.ej. la sesión se abrió
// con datos ya pintados de una versión anterior, antes de que existiera
// esta caché, y aún no se ha refrescado) se cae al comportamiento antiguo:
// atributesGrid (opcional, pasado por getCrossItemsForWriteback) permite
// reconocer un valor numérico que en realidad es una fecha (DATE/DATETIME/
// TIMESTAMP en MODEL_ATRIBUTES) y convertirlo desde el número de serie de
// Excel; si tampoco aplica, se usa el texto tal cual lo ve el usuario.
function dracoPlanningCrossItemsForCellDynamic(ctx, cellAddr, atributesGrid) {
    const { row, col } = parseAddress(cellAddr); // 1-based
    const items = [];
    const cache = ctx.reportId !== undefined ? DracoAxisCrossValues.get(ctx.reportId) : null;

    const resolveValue = (cacheKey, dim, currentValue, text) => {
        const cached = cache ? cache.get(cacheKey) : null;
        if (cached) return cached.value;
        if (atributesGrid && typeof currentValue === "number") {
            const dataType = resolveDracoPlanningDimensionType(atributesGrid, dim);
            if (dracoPlanningValueIsDateType(dataType)) return excelSerialToISODate(currentValue);
        }
        return text;
    };

    // Cruce de FILA: mismo ROW absoluto, recorriendo las columnas del
    // rango Draco_XXX_Rows (cada columna = un nivel de jerarquía).
    if (ctx.rowsRange && !ctx.rowsRange.isNullObject) {
        const rr = ctx.rowsRange;
        const localRow = (row - 1) - rr.rowIndex;
        if (localRow >= 0 && localRow < rr.rowCount) {
            for (let c = 0; c < rr.columnCount; c++) {
                const currentValue = rr.values[localRow][c];
                if (currentValue === "" || currentValue === null || currentValue === undefined) continue;
                const iAux = c + 1;
                const flag = (ctx.rowsIndent[localRow][c] || 0) + 1;
                const field = resolveAxisFieldForFlag(ctx.fieldLevelsRows[iAux], flag);
                if (!field || String(field.dim).toUpperCase() === "MEASURE") continue;
                const text = String(rr.text[localRow][c]).replace(DRACO_GLYPH_PREFIX, "");
                const physCol = rr.columnIndex + c + 1; // physical row = row (ver comentario en el módulo)
                const cacheKey = row + "_" + physCol;
                items.push({ dim: field.dim, attr: field.attr, value: resolveValue(cacheKey, field.dim, currentValue, text) });
            }
        }
    }

    // Cruce de COLUMNA: mismo COL absoluto, recorriendo las filas del
    // rango Draco_XXX_Cols.
    if (ctx.colsRange && !ctx.colsRange.isNullObject) {
        const cr = ctx.colsRange;
        const localCol = (col - 1) - cr.columnIndex;
        if (localCol >= 0 && localCol < cr.columnCount) {
            for (let r = 0; r < cr.rowCount; r++) {
                const currentValue = cr.values[r][localCol];
                if (currentValue === "" || currentValue === null || currentValue === undefined) continue;
                const iAux = r + 1;
                const flag = (ctx.colsIndent[r][localCol] || 0) + 1;
                const field = resolveAxisFieldForFlag(ctx.fieldLevelsCols[iAux], flag);
                if (!field || String(field.dim).toUpperCase() === "MEASURE") continue;
                const text = String(cr.text[r][localCol]).replace(DRACO_GLYPH_PREFIX, "");
                const physRow = cr.rowIndex + r + 1; // physical col = col (ver comentario en el módulo)
                const cacheKey = physRow + "_" + col;
                items.push({ dim: field.dim, attr: field.attr, value: resolveValue(cacheKey, field.dim, currentValue, text) });
            }
        }
    }

    return items;
}

function dracoPlanningCrossTextForCellDynamic(ctx, cellAddr) {
    return dracoPlanningCrossItemsForCellDynamic(ctx, cellAddr)
        .map(it => `${it.dim}.${it.attr}=${it.value}`).join(", ");
}

async function writeDracoPlanningSavedTable() {
    if (DracoPlanningModifiedCells.size === 0) return;

    try {
        await Excel.run(async (context) => {
            const editReportSheet = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
            editReportSheet.load("isNullObject");
            await context.sync();
            if (editReportSheet.isNullObject) return;

            const rows = [];
            for (const [reportId, bucket] of DracoPlanningModifiedCells.entries()) {
                if (bucket.size === 0) continue;

                let ctx = null;
                let crossDiag = "";
                try {
                    ctx = await buildDracoPlanningReportContext(context, reportId);
                } catch (e) {
                    console.warn(`[Draco] No se pudo resolver el contexto de planificación del informe ${reportId} para la tabla guardada:`, e);
                    crossDiag = "EXCEPCIÓN: " + (e && e.message ? e.message : String(e));
                    ctx = null;
                }

                const report = window.ReportStore ? window.ReportStore.getReport(reportId) : null;
                const reportName = (report && report.name) || `Informe ${reportId}`;

                for (const entry of bucket.values()) {
                    const items = ctx ? ctx.getCrossItems(entry.address) : [];
                    const crossText = items.map(it => `${it.dim}.${it.attr}=${it.value}`).join(", ");

                    // Valor actual (el importe tecleado) y diferencia contra
                    // el valor de referencia del último refresco (ver
                    // filterDracoPlanningRealChanges, que ya los deja
                    // guardados en la propia entrada). Si no hay referencia
                    // (informe nunca refrescado en esta sesión), la
                    // diferencia se deja vacía — no hay contra qué comparar.
                    const currentNum = Number(entry.currentValue);
                    const baselineNum = Number(entry.baselineValue);
                    const diff = (entry.baselineValue !== undefined && !isNaN(currentNum) && !isNaN(baselineNum))
                        ? currentNum - baselineNum
                        : "";

                    rows.push([
                        reportName,
                        entry.sheetName + "!" + entry.address,
                        entry.color || "",
                        crossText,
                        ctx ? ctx.measureLabel : "",
                        ctx ? ctx.filtersText : "",
                        entry.currentValue !== undefined ? entry.currentValue : "",
                        diff,
                        crossDiag
                    ]);
                }
            }
            if (rows.length === 0) return;

            const startRow = Number(DRACO_PLANNING_SAVED_TABLE_CELL.replace(/\D/g, "")); // 130
            editReportSheet.getRangeByIndexes(startRow - 1, 0, 1, 9).values =
                [["Informe", "Celda", "Color anterior", "Cruce", "Medida", "Filtros", "Valor", "Diferencia vs JSON", "Diagnóstico"]];
            editReportSheet.getRangeByIndexes(startRow, 0, rows.length, 9).values = rows;

            await context.sync();
        });
    } catch (e) {
        console.error("[Draco] Error escribiendo la tabla de celdas de planificación guardadas:", e);
    }
}

/**
 * Compara el valor ACTUAL de cada celda marcada en cian contra su valor
 * de referencia (el que pintó el último refresco, guardado en
 * DracoPlanningBaselineValues) y descarta del registro las que en
 * realidad no cambiaron de valor (el usuario la tocó, pero la dejó
 * igual que estaba) — así lo que llegue después a la tabla guardada / al
 * MERGE es solo lo que de verdad hace falta escribir.
 *
 * Si una celda no tiene valor de referencia (el informe nunca se
 * refrescó en esta sesión, por ejemplo), se mantiene tal cual: sin
 * referencia no hay forma de saber si cambió, así que por seguridad se
 * asume que sí.
 */
async function filterDracoPlanningRealChanges() {
    const allCells = []; // {reportId, address, sheetName}
    for (const [reportId, bucket] of DracoPlanningModifiedCells.entries()) {
        for (const entry of bucket.values()) {
            allCells.push({ reportId, address: entry.address, sheetName: entry.sheetName });
        }
    }
    if (allCells.length === 0) return;

    try {
        await Excel.run(async (context) => {
            const cellRanges = allCells.map(c =>
                context.workbook.worksheets.getItem(c.sheetName).getRange(c.address)
            );
            cellRanges.forEach(cr => cr.load("values"));
            await context.sync();

            let descartadas = 0;
            allCells.forEach((c, i) => {
                const { row, col } = parseAddress(c.address);
                const baseline = DracoPlanningBaselineValues.get(c.reportId);
                const baselineValue = baseline ? baseline.get(row + "_" + col) : undefined;
                const currentValue = cellRanges[i].values[0][0];

                if (baseline !== undefined && baselineValue !== undefined && String(baselineValue) === String(currentValue)) {
                    const bucket = DracoPlanningModifiedCells.get(c.reportId);
                    if (bucket) {
                        bucket.delete(c.address);
                        descartadas++;
                    }
                    return;
                }

                // Se queda: se guardan valor actual y de referencia en la
                // propia entrada, para no releerlos en writeDracoPlanningSavedTable.
                const bucket = DracoPlanningModifiedCells.get(c.reportId);
                const entry = bucket ? bucket.get(c.address) : null;
                if (entry) {
                    entry.currentValue = currentValue;
                    entry.baselineValue = baselineValue; // undefined si no había referencia
                }
            });

            // Informes que se quedaron sin ninguna celda de verdad modificada.
            for (const [reportId, bucket] of DracoPlanningModifiedCells.entries()) {
                if (bucket.size === 0) DracoPlanningModifiedCells.delete(reportId);
            }

            if (descartadas > 0) {
                console.log(`[Draco] ${descartadas} celda(s) marcada(s) en cian pero con el mismo valor que el último refresco: descartadas antes de guardar.`);
            }
        });
    } catch (e) {
        console.warn("[Draco] No se pudo comparar el valor actual contra el de referencia; se guardan todas las celdas marcadas por si acaso:", e);
    }
}

/**
 * Comprueba que cada celda modificada de planificación tenga valor para
 * TODAS las dimensiones marcadas como obligatorias (Propiedades del
 * informe, o el criterio propio de su medida si tiene "Personalizar"
 * marcado — ver buildDracoPlanningReportContext.dimensionsBehaviorForMeasure).
 *
 * De paso (para no recorrer las celdas dos veces ni recalcular el
 * contexto por informe otra vez), monta también la fila que le
 * correspondería a cada celda en el futuro INSERT/MERGE — ver `rows` en
 * el resultado. Si una dimensión NO obligatoria no aparece en la celda,
 * su valor es `null` (criterio "NULL") o el marcador `"__REPARTO__"`
 * (criterio "lineal"/"proporcional" — de momento solo un aviso, el
 * reparto de verdad se ataca después, aparte).
 *
 * Reutiliza buildDracoPlanningReportContext (una vez por informe, no por
 * celda) — el cruce/medida ya sale calculado una sola vez ahí, así que
 * esto no añade ninguna ida y vuelta a Excel extra aparte de la que ya
 * hacía falta para la tabla guardada.
 *
 * Devuelve {valid, errors, rows, dimNames}:
 *   - errors: array de {reportName, cell, missingDims} — una entrada por
 *     cada celda a la que le falte alguna dimensión obligatoria.
 *   - rows: array de {reportId, reportName, address, dimsValues,
 *     measureName, value} — SOLO si valid es true (si falta algo
 *     obligatorio en cualquier celda, no se monta ninguna fila).
 *   - dimNames: nombres de todas las dimensiones usadas en `rows` (para
 *     saber qué columnas lleva el INSERT/MERGE).
 */
async function validateDracoPlanningMandatoryDimensions() {
    const errors = [];
    const rows = [];
    const dimNamesSet = new Set();
    if (DracoPlanningModifiedCells.size === 0) return { valid: true, errors, rows, dimNames: [] };

    try {
        await Excel.run(async (context) => {
            for (const [reportId, bucket] of DracoPlanningModifiedCells.entries()) {
                if (bucket.size === 0) continue;

                const report = window.ReportStore ? window.ReportStore.getReport(reportId) : null;
                const reportName = (report && report.name) || `Informe ${reportId}`;
                const modelName = report ? report.semanticModelName : "";

                // MODEL_ATRIBUTES del modelo de ESTE informe — hace falta YA
                // (antes de construir el contexto) porque getCrossItemsForWriteback
                // la usa para saber qué dimensiones son DATE y convertir su
                // valor al formato canónico (ver más abajo), no al texto tal
                // como lo muestra Excel según el locale del usuario.
                const atributesGrid = (modelName && window.SemanticModelStore)
                    ? await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES", modelName)
                    : null;

                let ctx;
                try {
                    ctx = await buildDracoPlanningReportContext(context, reportId, atributesGrid);
                } catch (e) {
                    console.warn(`[Draco] No se pudo validar las dimensiones obligatorias del informe ${reportId} (se deja pasar, sin contexto no se puede comprobar):`, e);
                    continue;
                }

                // Dimensiones resueltas por la zona "Filtros" (no por Filas/
                // Columnas) con un único valor concreto — ReportState.Filters
                // ya quedó relleno para ESTE informe por loadReportDefinition
                // (dentro de buildDracoPlanningReportContext). Una sola vez
                // por informe, no por celda: el filtro es el mismo para
                // todas las celdas modificadas de este informe.
                const filterValuesMap = new Map();
                for (const f of (ReportState.Filters || [])) {
                    const val = singleDracoFilterValueForWriteback(f.Value);
                    if (val !== undefined) filterValuesMap.set(String(f.Dimension).toUpperCase(), val);
                }

                // MODEL_MEASURES del modelo de ESTE informe, una sola vez
                // (no por celda) — de aquí sale la tabla de hechos real
                // (FACT_PROJECT.FACT_DATASET.FACT_TABLE) y el campo real de
                // cada medida (FACT_FIELD, p.ej. "IMPORTE"/"CANTIDAD") — el
                // destino real del INSERT, no un log genérico.
                const measuresGrid = (modelName && window.SemanticModelStore)
                    ? await window.SemanticModelStore.getModelGrid("MODEL_MEASURES", modelName)
                    : null;

                for (const entry of bucket.values()) {
                    const { row, col } = parseAddress(entry.address); // 1-based
                    const measure = ctx.resolveMeasure(row - 1, col - 1); // resolveMeasure espera 0-based
                    const behavior = measure
                        ? ctx.dimensionsBehaviorForMeasure(measure.zoneId, measure.measureName)
                        : (report && report.reportProperties ? (report.reportProperties.dimensionsBehavior || {}) : {});

                    const items = ctx.getCrossItemsForWriteback(entry.address);
                    const presentMap = new Map(items.map(it => [String(it.dim).toUpperCase(), it.value]));

                    // "Presente" = está en Filas/Columnas de ESTA celda, O
                    // resuelto por un filtro de la zona "Filtros" con un
                    // único valor (ver filterValuesMap arriba) — una
                    // dimensión puede estar fijada por filtro sin aparecer
                    // nunca en el cruce de fila/columna.
                    const hasValue = (upperDim) => presentMap.has(upperDim) || filterValuesMap.has(upperDim);
                    const valueFor = (upperDim) => presentMap.has(upperDim) ? presentMap.get(upperDim) : filterValuesMap.get(upperDim);

                    const requiredDims = Object.keys(behavior).filter(dim => behavior[dim] && behavior[dim].required);
                    const missing = requiredDims.filter(dim => !hasValue(String(dim).toUpperCase()));

                    if (missing.length > 0) {
                        errors.push({
                            reportName,
                            cell: entry.sheetName + "!" + entry.address,
                            missingDims: missing
                        });
                        continue; // esta celda falla: no se monta fila para ella
                    }

                    // Todas las obligatorias están — montar el valor de CADA
                    // dimensión configurada (esté o no en esta celda).
                    const dimsValues = {};
                    const dimTypes = {}; // dim -> DATA_TYPE (MODEL_ATRIBUTES), para no citar numéricos entre comillas
                    for (const dim of Object.keys(behavior)) {
                        dimNamesSet.add(dim);
                        const upperDim = dim.toUpperCase();
                        if (hasValue(upperDim)) {
                            dimsValues[dim] = valueFor(upperDim);
                        } else {
                            const cfg = behavior[dim] || {};
                            dimsValues[dim] = (cfg.mode === "lineal" || cfg.mode === "proporcional")
                                ? "__REPARTO__"
                                : null; // "null" (o sin criterio explícito): NULL
                        }
                        dimTypes[dim] = atributesGrid ? resolveDracoPlanningDimensionType(atributesGrid, dim) : "STRING";
                    }

                    const measureName = (measure && measure.measureName) || ctx.measureLabel || "";
                    const measureRow = measuresGrid ? buscarMedida(measuresGrid, measureName) : 0;
                    const factField = (measureRow && String(cellValue(measuresGrid, measureRow, 6)).trim()) || measureName;
                    const factTable = measureRow
                        ? Provider.qualify(
                            cellValue(measuresGrid, measureRow, 3),
                            cellValue(measuresGrid, measureRow, 4),
                            cellValue(measuresGrid, measureRow, 5)
                        )
                        : "";

                    // DRACO_PLANIFICACION es "append-only" (se lee con
                    // SUM(IMPORTE) — ver la consulta del informe): insertar
                    // el valor ABSOLUTO que el usuario tecleó sumaría por
                    // encima de lo que ya había, en vez de dejarlo en ese
                    // número. Hay que insertar la DIFERENCIA contra el
                    // último valor conocido (baseline, el que puso el
                    // último refresco — ver filterDracoPlanningRealChanges),
                    // igual que ya se calcula (solo para mostrar) en
                    // writeDracoPlanningSavedTable. Sin baseline (informe
                    // nunca refrescado en esta sesión, o valor no numérico)
                    // se inserta el valor tal cual: no hay nada de lo que
                    // restar.
                    const currentNum = Number(entry.currentValue);
                    const baselineNum = Number(entry.baselineValue);
                    const hasBaseline = entry.baselineValue !== undefined && !isNaN(baselineNum);
                    const deltaValue = (hasBaseline && !isNaN(currentNum))
                        ? currentNum - baselineNum
                        : entry.currentValue;

                    rows.push({
                        reportId,
                        reportName,
                        address: entry.sheetName + "!" + entry.address,
                        dimsValues,
                        dimTypes,        // dim -> DATA_TYPE (MODEL_ATRIBUTES): decide si el literal SQL va entre comillas
                        factTable,       // tabla de hechos real (vacía si no se pudo resolver la medida en MODEL_MEASURES)
                        factField,       // columna real de la medida (p.ej. "IMPORTE"), no un genérico "measure_name"
                        value: deltaValue
                    });
                }
            }
        });
    } catch (e) {
        console.error("[Draco] Error validando dimensiones obligatorias:", e);
        errors.push({ reportName: "", cell: "", missingDims: ["(error interno: " + (e && e.message ? e.message : String(e)) + ")"] });
    }

    return { valid: errors.length === 0, errors, rows: errors.length === 0 ? rows : [], dimNames: Array.from(dimNamesSet) };
}

// Celda donde se escribe el INSERT generado — lejos de la tabla de A130
// hacia abajo, para que no choquen aunque haya muchas celdas modificadas.
const DRACO_PLANNING_INSERT_SQL_CELL = "J1";

// Celda justo al lado de la anterior, donde se pinta el RESULTADO de
// ejecutar ese INSERT contra el proveedor activo (respuesta cruda del
// proveedor si va bien, o el mensaje de error si falla) — ver
// writeDracoPlanningInsertResult, llamada desde guardarPlanificacion justo
// después de executeDracoPlanningInserts.
const DRACO_PLANNING_INSERT_RESULT_CELL = "K1";

// Literal SQL para el valor de una dimensión: NULL, la palabra suelta
// "reparto" (marcador, aposta NO es SQL válido — el reparto de verdad se
// ataca después, aparte) o el valor — entre comillas simples si su
// DATA_TYPE (MODEL_ATRIBUTES) es de texto, o tal cual (sin comillas) si es
// numérico (INTEGER/FLOAT/NUMERIC/BIGNUMERIC — ver
// resolveDracoPlanningDimensionType), para no meter un STRING en una
// columna INT64/NUMERIC de BigQuery.
function dracoPlanningSqlLiteralForDimValue(value, dataType) {
    if (value === null || value === undefined) return "NULL";
    if (value === "__REPARTO__") return "reparto";

    if (dracoPlanningValueIsNumericType(dataType)) {
        const s = String(value).trim().replace(",", "."); // por si llega con coma decimal
        const num = Number(s);
        return (s === "" || isNaN(num)) ? "NULL" : num;
    }

    return "'" + String(value).replace(/'/g, "''") + "'";
}

function dracoPlanningSqlLiteralForMeasureValue(value) {
    if (value === null || value === undefined || value === "") return "NULL";
    const num = Number(value);
    return isNaN(num) ? "NULL" : num;
}

/**
 * Genera el texto del INSERT directo (Opción A: columnas fijas por
 * dimensión) contra la TABLA DE HECHOS real (no un log propio): tabla y
 * columna de medida salen de MODEL_MEASURES (FACT_PROJECT.FACT_DATASET.
 * FACT_TABLE / FACT_FIELD, ya resueltos por fila en
 * validateDracoPlanningMandatoryDimensions), así que la columna de
 * importe se llama tal cual está en el modelo (p.ej. "IMPORTE",
 * "CANTIDAD"), no un "measure_name"/"value" genérico. Se escribe en
 * EDIT_REPORT!J1 — de momento solo el TEXTO, no se ejecuta contra ningún
 * proveedor todavía. Las dimensiones con criterio "lineal"/"proporcional"
 * que faltaban en su celda llevan la palabra "reparto" en vez de un
 * valor real — se sustituirá por el reparto de verdad más adelante.
 *
 * Si hay filas de distintas tablas de hechos (informes con medidas de
 * tablas distintas) se genera un INSERT por cada tabla. Si hay más de
 * una medida involucrada para la MISMA tabla (p.ej. IMPORTE y CANTIDAD),
 * se añaden ambas columnas y cada fila deja NULL la que no le
 * corresponde a ella.
 */
// Agrupa las filas ya validadas por tabla de hechos real y construye un
// INSERT por cada una (columnas fijas por dimensión + una columna por
// cada medida realmente usada en esa tabla). Reutilizada tanto para
// escribir el texto en EDIT_REPORT!J1 como para ejecutarlo de verdad
// (executeDracoPlanningInserts) — una sola vez se decide cómo se
// construye el INSERT, no dos.
/**
 * Extrae el proyecto GCP de una tabla ya cualificada por Provider.qualify()
 * (`proyecto.dataset.tabla`, con o sin backticks). Esto es lo
 * que permite facturar el INSERT contra el MISMO proyecto donde vive la
 * tabla de hechos real, en vez de un proyecto de facturación fijo.
 */
function projectIdFromQualifiedTable(qualifiedTable) {
    const clean = String(qualifiedTable || "").replace(/`/g, "").trim();
    const parts = clean.split(".");
    return parts.length >= 3 ? parts[0] : null;
}

function buildDracoPlanningInsertStatements(rows, dimNames) {
    const byTable = new Map(); // factTable -> {rows: [...], factFields: Set}
    for (const r of rows) {
        const table = r.factTable || "(tabla de hechos no resuelta)";
        if (!byTable.has(table)) byTable.set(table, { rows: [], factFields: new Set() });
        const bucket = byTable.get(table);
        bucket.rows.push(r);
        bucket.factFields.add(r.factField);
    }

    const statements = [];
    for (const [table, bucket] of byTable.entries()) {
        const factFields = Array.from(bucket.factFields);
        const columns = [...dimNames, ...factFields];

        const tuples = bucket.rows.map(r => {
            const dimVals = dimNames.map(d => dracoPlanningSqlLiteralForDimValue(r.dimsValues[d], r.dimTypes && r.dimTypes[d]));
            const measureVals = factFields.map(f =>
                f === r.factField ? dracoPlanningSqlLiteralForMeasureValue(r.value) : "NULL"
            );
            return "    (" + [...dimVals, ...measureVals].join(", ") + ")";
        });

        statements.push({
            table,
            sql:
`INSERT INTO ${table}
  (${columns.join(", ")})
VALUES
${tuples.join(",\n")};`
        });
    }

    return statements;
}

async function writeDracoPlanningMergeSql(rows, dimNames) {
    if (rows.length === 0) return;

    const sql = buildDracoPlanningInsertStatements(rows, dimNames).map(s => s.sql).join("\n\n");

    try {
        await Excel.run(async (context) => {
            const editReportSheet = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
            editReportSheet.load("isNullObject");
            await context.sync();
            if (editReportSheet.isNullObject) return;
            editReportSheet.getRange(DRACO_PLANNING_INSERT_SQL_CELL).values = [[sql]];
            await context.sync();
        });
    } catch (e) {
        console.error("[Draco] Error escribiendo el INSERT de planificación en EDIT_REPORT:", e);
    }
}

/**
 * Ejecuta de verdad, contra el proveedor activo (BigQuery, ver
 * executeSQL), cada INSERT generado por buildDracoPlanningInsertStatements
 * — uno por tabla de hechos distinta. Se para en el primero que falle
 * (no sigue con los demás) y devuelve el error tal cual lo dio el
 * proveedor, para poder mostrarlo.
 *
 * BigQuery no lanzan excepción si el SQL falla — devuelven un
 * texto/JSON con "error"/"errors" dentro, con código 200 igualmente — por
 * eso hace falta mirar el contenido de la respuesta, no solo capturar
 * try/catch.
 */
async function executeDracoPlanningInserts(rows, dimNames) {
    if (rows.length === 0) return { success: true, error: "", responses: [] };

    const statements = buildDracoPlanningInsertStatements(rows, dimNames);
    const responses = []; // respuesta cruda de CADA statement ejecutado con éxito, en orden

    for (const stmt of statements) {
        try {
            // Factura el INSERT contra el proyecto de SU tabla de hechos
            // (p.ej. "draco-506807"), no contra un proyecto fijo sin billing.
            const billingProjectId = projectIdFromQualifiedTable(stmt.table);
            const responseText = await executeSQL(stmt.sql, billingProjectId);
            let parsed = null;
            try { parsed = JSON.parse(responseText); } catch (e) { /* respuesta no era JSON, se asume OK */ }

            if (parsed) {
                if (parsed.error) {
                    return { success: false, error: parsed.error.message || JSON.stringify(parsed.error), responses };
                }
                if (Array.isArray(parsed.errors) && parsed.errors.length > 0) {
                    return { success: false, error: parsed.errors.map(e => e.message || JSON.stringify(e)).join("; "), responses };
                }
            }

            responses.push(responseText);
        } catch (e) {
            return { success: false, error: e && e.message ? e.message : String(e), responses };
        }
    }

    return { success: true, error: "", responses };
}

/**
 * Pinta en EDIT_REPORT!K1 (al lado del texto del INSERT en J1, ver
 * DRACO_PLANNING_INSERT_RESULT_CELL) el resultado de ejecutarlo: la
 * respuesta cruda del proveedor si fue bien (una por statement/tabla de
 * hechos, separadas por una línea en blanco), o el mensaje de error si
 * falló — así se ve de un vistazo, sin mirar la consola, que el INSERT
 * realmente se ha lanzado y qué contestó el proveedor.
 */
async function writeDracoPlanningInsertResult(insertResult) {
    const text = insertResult.success
        ? (insertResult.responses && insertResult.responses.length
            ? insertResult.responses.join("\n\n")
            : "(sin filas que insertar)")
        : "ERROR: " + insertResult.error;

    try {
        await Excel.run(async (context) => {
            const editReportSheet = context.workbook.worksheets.getItemOrNullObject("EDIT_REPORT");
            editReportSheet.load("isNullObject");
            await context.sync();
            if (editReportSheet.isNullObject) return;
            editReportSheet.getRange(DRACO_PLANNING_INSERT_RESULT_CELL).values = [[text]];
            await context.sync();
        });
    } catch (e) {
        console.error("[Draco] Error escribiendo el resultado del INSERT de planificación en EDIT_REPORT:", e);
    }
}

/**
 * Popup pequeño (planningValidationBadge.html) que confirma el resultado
 * de validateDracoPlanningMandatoryDimensions al pulsar "Guardar
 * planificación" — mismo patrón que showMemberRecognitionBadge
 * (displayDialogAsync + auto-cierre): "fire and forget", no bloquea a
 * quien llama. Si es válido, se cierra solo a los ~1,4s (como el de
 * Reconocimiento de miembros); si faltan dimensiones, se queda abierto
 * (hay una lista que leer) hasta que el usuario haga clic.
 *
 * titulo/subtitulo son opcionales — si no se pasan, usa los textos por
 * defecto de validación de dimensiones; para el resultado de ejecutar el
 * INSERT de verdad se pasan unos distintos (ver guardarPlanificacion).
 */
function showPlanningValidationBadge(isValid, detalle, titulo, subtitulo) {
    try {
        const params = new URLSearchParams({
            valid: isValid ? "1" : "0",
            detalle: detalle || ""
        });
        if (titulo) params.set("titulo", titulo);
        if (subtitulo) params.set("subtitulo", subtitulo);
        const dialogUrl = new URL(
            `planningValidationBadge.html?${params.toString()}`,
            window.location.href
        ).href;

        Office.context.ui.displayDialogAsync(
            dialogUrl,
            isValid ? { height: 15, width: 22, displayInIframe: false } : { height: 35, width: 32, displayInIframe: false },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.warn(
                        "[Draco] No se pudo mostrar el aviso de validación de planificación:",
                        asyncResult.error && asyncResult.error.message
                    );
                    return;
                }
                const dialog = asyncResult.value;
                let closed = false;
                const closeOnce = () => {
                    if (closed) return;
                    closed = true;
                    try { dialog.close(); } catch (e) { /* ya cerrado */ }
                };
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, closeOnce);
                dialog.addEventHandler(Office.EventType.DialogEventReceived, closeOnce);
            }
        );
    } catch (e) {
        console.warn("[Draco] Error mostrando el aviso de validación de planificación:", e);
    }
}

async function guardarPlanificacion(event) {
    try {
        console.log("Guardar planificación: ejecuta el INSERT contra el proveedor activo y refresca los informes afectados.");

        // Solo con permiso de escritura y sobre modelos de Draco Planning (ver DracoPlanningGate).
        if (window.DracoPlanningGate && !DracoPlanningGate.canWrite()) {
            showPlanningValidationBadge(false, DracoPlanningGate.check("").reason,
                "Planificación no disponible", "Permiso de solo lectura");
            return;
        }

        // Descarta las celdas marcadas en cian que en realidad no cambiaron
        // de valor respecto al último refresco (ver comentario de la
        // función) — antes de comprobar nada ni de mandar nada al INSERT.
        await filterDracoPlanningRealChanges();

        // Dimensiones obligatorias: si a alguna celda le falta alguna, se
        // avisa y NO se sigue (ni tabla guardada, ni vaciado del registro
        // — las celdas se quedan tal cual, en cian, para poder corregirlas).
        // De paso, si todo va bien, ya vienen montadas las filas del INSERT
        // (validation.rows) — ver la función para el detalle.
        const validation = await validateDracoPlanningMandatoryDimensions();
        if (window.DracoPlanningGate && validation.rows && window.ReportStore) {
            const rejected = Array.from(new Set(validation.rows.map(r => r.reportId)))
                .map(id => window.ReportStore.getReport(id))
                .filter(rep => rep && !DracoPlanningGate.check(rep.semanticModelName).ok);
            if (rejected.length) {
                showPlanningValidationBadge(false, DracoPlanningGate.check(rejected[0].semanticModelName).reason,
                    "Planificación no disponible", "Modelo no válido para planificar");
                return;
            }
        }
        if (!validation.valid) {
            const detalle = validation.errors
                .map(e => `${e.reportName} · ${e.cell}: ${e.missingDims.join(", ")}`)
                .join("\n");
            showPlanningValidationBadge(false, detalle);
            return;
        }
        showPlanningValidationBadge(true, "");

        // INSERT (Opción A: columnas fijas por dimensión) — se escribe el
        // texto en EDIT_REPORT!J1 (para poder revisarlo/depurarlo) Y se
        // ejecuta de verdad contra el proveedor activo. Las dimensiones
        // con criterio "lineal"/"proporcional" que faltaban llevan la
        // palabra "reparto" en vez de un valor (el reparto de verdad se
        // ataca aparte, cuando el caso principal funcione bien — con
        // "reparto" en el texto, ejecutarlo fallaría por sintaxis, así que
        // de momento ese caso concreto no debería llegar a "Guardar" sin
        // que antes se revise a mano).
        await writeDracoPlanningMergeSql(validation.rows, validation.dimNames);

        const insertResult = await executeDracoPlanningInserts(validation.rows, validation.dimNames);

        // Se pinta el resultado (respuesta del proveedor o error) en la
        // celda de al lado del INSERT (K1) SIEMPRE, vaya bien o mal, para
        // poder comprobar de un vistazo que el "Guardar" realmente ha
        // llegado a ejecutarse contra el proveedor.
        await writeDracoPlanningInsertResult(insertResult);

        if (!insertResult.success) {
            showPlanningValidationBadge(
                false,
                insertResult.error,
                "Error al guardar en la tabla de hechos",
                "No se ha podido ejecutar el INSERT:"
            );
            return; // no se vacía el registro: las celdas se quedan en cian para poder reintentar
        }
        showPlanningValidationBadge(true, "", "Guardado", "Los datos se han escrito en la tabla de hechos");

        // Tabla legible (celda | color anterior) — SOLO en "Guardar", antes
        // de que flushDracoPlanningModifiedCells vacíe el registro.
        await writeDracoPlanningSavedTable();

        // Qué informes hay que refrescar, ANTES de vaciar el registro
        // (flushDracoPlanningModifiedCells lo deja vacío).
        const affectedReportIds = Array.from(new Set(validation.rows.map(r => r.reportId)));

        // Vuelca en EDIT_REPORT!A127 las celdas de Draco_<id>_Values
        // modificadas a mano desde el último guardado/refresco (junto con
        // el color que tenían antes de marcarse en cian), les devuelve
        // ese color original, y vacía el registro en memoria. Ver
        // handleDracoPlanningValueChanged / flushDracoPlanningModifiedCells.
        await flushDracoPlanningModifiedCells();

        // Refrescar cada informe afectado, para que se vea ya el dato
        // recién guardado (actualizarUnInforme decide Fijo/Dinámico según
        // el diseño de cada uno).
        for (const reportId of affectedReportIds) {
            try {
                await actualizarUnInforme(reportId);
            } catch (e) {
                console.error(`[Draco] Error refrescando el informe ${reportId} tras guardar planificación:`, e);
            }
        }
    } catch (e) {
        console.error("[Draco] Error en guardarPlanificacion:", e);
    } finally {
        if (event) {
            event.completed();
        }
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

/**
 * Popup pequeño y auto-cierre (recognitionBadge.html) que confirma
 * visualmente el nuevo estado ON/OFF de "Reconocimiento de miembros" al
 * pulsar el botón del ribbon. Es "fire and forget": no bloquea al que
 * llama (no se espera aquí dentro), y si por lo que sea falla al abrir
 * (host sin soporte de displayDialogAsync, etc.) simplemente no se ve
 * nada — la etiqueta del propio botón del ribbon ya refleja el estado.
 */
function showMemberRecognitionBadge(nowOn) {
    try {
        const dialogUrl = new URL(
            `recognitionBadge.html?state=${nowOn ? "on" : "off"}`,
            window.location.href
        ).href;

        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 15, width: 20, displayInIframe: false },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.warn(
                        "[Draco] No se pudo mostrar el badge de Reconocimiento de miembros:",
                        asyncResult.error && asyncResult.error.message
                    );
                    return;
                }
                const dialog = asyncResult.value;
                let closed = false;
                const closeOnce = () => {
                    if (closed) return;
                    closed = true;
                    try { dialog.close(); } catch (e) { /* ya cerrado */ }
                };
                // El propio HTML se auto-cierra pasado un momento (o al
                // clicarlo) y avisa aquí vía messageParent para que el
                // host cierre de verdad el diálogo.
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, closeOnce);
                // 12006 = el usuario lo cerró él mismo con la X.
                dialog.addEventHandler(Office.EventType.DialogEventReceived, closeOnce);
            }
        );
    } catch (e) {
        console.warn("[Draco] Error mostrando el badge de Reconocimiento de miembros:", e);
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

        showMemberRecognitionBadge(nowOn);
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
// Candado contra invocaciones simultáneas: si el botón "Añadir filtro" se
// dispara dos veces seguidas (doble clic, o el propio host mandando el
// evento de clic repetido — se ha visto en Excel Online), sin esto se
// llamaba a displayDialogAsync() dos veces. Office.js solo lleva bien UN
// diálogo activo a la vez: la segunda llamada concurrente deja el canal de
// mensajes hecho un lío (de ahí que a veces salieran 2 ventanas, y a veces
// una de ellas se quedara sin recibir los datos). Con el candado, un
// segundo clic mientras ya hay uno abierto simplemente se ignora.
let dracoFilterDialogOpen = false;

/* ---------------------------------------------------------------------
 * VARIABLES (js/variablesStore.js + variablesDialog.html)
 *
 * Comando de cinta "abrirVariables": abre la ventana de variables (nombre,
 * campo y valor elegido de la lista de valores del campo). Registrado con
 * Office.actions.associate, así que en el addin basta con añadir un botón
 * al manifest que llame a "abrirVariables". En Tablas Draco (PLANNING) lo
 * abre el botón «Variables» del diseñador, con el mismo diálogo.
 * ------------------------------------------------------------------- */
let dracoVariablesDialogOpen = false;

async function dracoModelFieldList() {
    const grid = await window.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
    const out = [];
    for (let R = 2; String(cellValue(grid, R, 2)).trim() !== ""; R++) {
        const dim = String(cellValue(grid, R, 2)).trim(), attr = String(cellValue(grid, R, 3)).trim();
        if (attr) out.push({ dim, attr, label: dim.toLowerCase() + " · " + attr.toLowerCase() });
    }
    return out.sort((a, b) => a.label.localeCompare(b.label));
}

async function dracoFieldMembers(dim, attr) {
    const sql = await window.ExcelService.buildFilterValuesSQL(dim, attr);
    if (!sql) throw new Error("No se encuentra el campo en el modelo semántico.");
    const json = await window.ExcelService.executeSQL(sql);
    return Array.from(new Set(extractDracoJsonValues(json).filter(v => !isDracoNullLiteral(v)).map(v => String(v)))).slice(0, 5000);
}

async function abrirVariables(event) {
    if (dracoVariablesDialogOpen || !window.DracoVariables) { if (event) event.completed(); return; }
    dracoVariablesDialogOpen = true;
    try {
        let fields = [];
        try { fields = await dracoModelFieldList(); } catch (e) { console.warn("[Draco] Variables: campos del modelo:", e); }
        const url = new URL("variablesDialog.html", window.location.href).href;
        await new Promise((resolve) => {
            Office.context.ui.displayDialogAsync(url, { height: 70, width: 55, displayInIframe: false }, (res) => {
                if (res.status === Office.AsyncResultStatus.Failed) { console.error("Variables: displayDialogAsync:", res.error && res.error.message); resolve(); return; }
                const dialog = res.value;
                const reply = (o) => { try { dialog.messageChild(JSON.stringify(o)); } catch (e) { console.error(e); } };
                const done = () => { try { dialog.close(); } catch (e) { /* cerrado */ } resolve(); };
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                    let p;
                    try { p = JSON.parse(arg.message); } catch (e) { return; }
                    if (p.type === "ready") { reply({ type: "init", variables: window.DracoVariables.list(), fields }); return; }
                    if (p.type === "requestMembers") {
                        try { reply({ type: "members", dim: p.dim, attr: p.attr, members: await dracoFieldMembers(p.dim, p.attr) }); }
                        catch (e) { reply({ type: "members", dim: p.dim, attr: p.attr, members: [], error: (e && e.message) || String(e) }); }
                        return;
                    }
                    if (p.type === "save") {
                        await window.DracoVariables.saveAll(p.variables || []);
                        // En Tablas Draco el "Excel falso" recalcula al terminar cada Excel.run: así las
                        // cabeceras $VARIABLE$ muestran ya su nuevo valor. En Excel no hace nada.
                        try { await Excel.run(async (context) => { await context.sync(); }); } catch (e) { /* nada */ }
                        return;
                    }
                    if (p.type === "close") done();
                });
                dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve());
            });
        });
    } finally {
        dracoVariablesDialogOpen = false;
        if (event) event.completed();
    }
}

/**
 * Clic derecho sobre celdas de Filas/Columnas -> "Filtrar valores
 * seleccionados": los valores (ID) de las celdas seleccionadas pasan a la
 * zona Filtros del informe, como lista de valores del campo de cada celda
 * (si ese campo ya estaba en Filtros se sustituye su selección).
 * Celdas EPM_VALUE (estático) usan su atributo/valor (el ID); el resto, el
 * campo del nivel de la celda y su texto.
 */
async function filtrarValoresSeleccionados(event) {
    try {
        const groups = new Map(); // "DIM|ATTR" -> {dim, attr, values:Set}
        let reportId = null;
        let problem = "";
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("id");
            const sel = context.workbook.getSelectedRange();
            sel.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "values", "formulas"]);
            await context.sync();
            if (sel.rowCount * sel.columnCount > 5000) { problem = "Selección demasiado grande."; return; }

            const indents = await readIndentLevelsForRange(context, sel);
            const first = sel.address.split("!").pop().split(":")[0].replace(/\$/g, "");
            const located = await findDracoRowsNamedRangeForCell(context, sheet.id, first);
            if (!located) { problem = "La selección no está en Filas/Columnas de un informe Draco."; return; }
            reportId = located.reportId;

            const fieldCache = new Map();
            for (let r = 0; r < sel.rowCount; r++) {
                for (let c = 0; c < sel.columnCount; c++) {
                    const value = sel.values[r][c];
                    if (value === "" || value === null || value === undefined) continue;
                    let dim, attr, v = value;
                    const fx = sel.formulas[r][c];
                    const args = (typeof fx === "string" && /^=\s*EPM_VALUE\s*\(/i.test(fx)) ? parseEpmValueArgs(fx) : null;
                    if (args && args.length >= 3) {
                        dim = args[0]; attr = args[1]; v = args[2];
                    } else {
                        const level = located.axis === "rows" ? located.level + c : located.level + r;
                        const key = level + "|" + indents[r][c];
                        if (!fieldCache.has(key)) {
                            fieldCache.set(key, await resolveDracoFieldForAxisLevel(context, located.reportId, located.axis, level, indents[r][c]));
                        }
                        const f = fieldCache.get(key);
                        if (!f) continue;
                        dim = f.dim; attr = f.levelAttr || f.attr;
                    }
                    if (!dim || !attr || String(dim).toUpperCase() === "MEASURE") continue;
                    const g = dim + "|" + attr;
                    if (!groups.has(g)) groups.set(g, { dim, attr, values: new Set() });
                    groups.get(g).values.add(String(v));
                }
            }
        });

        if (problem || !groups.size) {
            console.warn("[Draco] Filtrar valores seleccionados:", problem || "ninguna celda de campo seleccionada.");
            if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.setAutoStatus) {
                TaskPaneApp.setAutoStatus(problem || "Selecciona celdas de Filas/Columnas para filtrar.");
            }
            return;
        }
        if (typeof TaskPaneApp === "undefined" || !TaskPaneApp.setFilterValues) {
            console.warn("[Draco] Filtrar valores seleccionados: taskpane no disponible.");
            return;
        }
        await TaskPaneApp.setFilterValues(reportId, Array.from(groups.values()).map(g => ({ dim: g.dim, attr: g.attr, values: Array.from(g.values) })));
        if (Office.addin && Office.addin.showAsTaskpane) { try { await Office.addin.showAsTaskpane(); } catch (e) { /* nada */ } }
    } catch (error) {
        console.error("Error al filtrar valores seleccionados:", error);
    } finally {
        if (event) event.completed();
    }
}

/**
 * Clic derecho sobre una celda de Filas/Columnas -> acciones sobre SU campo
 * (totales, mostrar ID/descripción, opciones de jerarquía). Averigua el
 * informe y el campo de la celda activa y se lo pasa al taskpane.
 */
async function dracoFieldActionFromCell(action) {
    let target = null;
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("id");
        const sel = context.workbook.getSelectedRange();
        sel.load("address");
        await context.sync();
        const first = sel.address.split("!").pop().split(":")[0].replace(/\$/g, "");
        const located = await findDracoRowsNamedRangeForCell(context, sheet.id, first);
        if (!located) return;
        const cell = sheet.getRange(first);
        cell.format.load("indentLevel");
        await context.sync();
        const f = await resolveDracoFieldForAxisLevel(context, located.reportId, located.axis, located.level, cell.format.indentLevel);
        if (!f || String(f.dim).toUpperCase() === "MEASURE") return;
        target = { reportId: located.reportId, axis: located.axis, dim: f.dim, attr: f.attr, levelAttr: f.levelAttr };
    });
    if (!target) {
        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.setAutoStatus) TaskPaneApp.setAutoStatus("La celda no es un campo de Filas/Columnas de un informe Draco.");
        return;
    }
    if (typeof TaskPaneApp === "undefined" || !TaskPaneApp.fieldActionFromCell) return;
    await TaskPaneApp.fieldActionFromCell(target, action);
    if (action === "options" && Office.addin && Office.addin.showAsTaskpane) { try { await Office.addin.showAsTaskpane(); } catch (e) { /* nada */ } }
}
async function campoTotalesMostrar(event) {
    try { await dracoFieldActionFromCell("totals:show"); }
    catch (error) { console.error("Error en campoTotalesMostrar:", error); }
    finally { if (event) event.completed(); }
}
async function campoTotalesOcultar(event) {
    try { await dracoFieldActionFromCell("totals:hide"); }
    catch (error) { console.error("Error en campoTotalesOcultar:", error); }
    finally { if (event) event.completed(); }
}
async function campoMostrarID(event) {
    try { await dracoFieldActionFromCell("mode:ID"); }
    catch (error) { console.error("Error en campoMostrarID:", error); }
    finally { if (event) event.completed(); }
}
async function campoMostrarIDTexto(event) {
    try { await dracoFieldActionFromCell("mode:ID_TEXT"); }
    catch (error) { console.error("Error en campoMostrarIDTexto:", error); }
    finally { if (event) event.completed(); }
}
async function campoMostrarTexto(event) {
    try { await dracoFieldActionFromCell("mode:TEXT"); }
    catch (error) { console.error("Error en campoMostrarTexto:", error); }
    finally { if (event) event.completed(); }
}
async function campoOpciones(event) {
    try { await dracoFieldActionFromCell("options"); }
    catch (error) { console.error("Error en campoOpciones:", error); }
    finally { if (event) event.completed(); }
}

async function abrirAnadirFiltro(event) {
    if (dracoFilterDialogOpen) {
        console.warn("[Draco] Añadir filtro: ya hay un diálogo abierto; se ignora este segundo clic.");
        if (event) event.completed();
        return;
    }
    dracoFilterDialogOpen = true;
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
                                // En Tablas Draco (PLANNING) una tabla = UN informe: el
                                // diálogo no pregunta a qué informe aplica el filtro.
                                const singleReport = !!window.DracoEngineApi;
                                dialog.messageChild(JSON.stringify({ type: "data", dimensions, reports, singleReport }));
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

                                // Para refrescar la zona "Filtros" si el taskpane ya
                                // estaba abierto (mismo contexto que commands.js).
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
        dracoFilterDialogOpen = false;
        if (event) event.completed();
    }
}

/* =======================================================================
 * "Copiar informe" / "Pegar informe" / "Eliminar informe" (ribbon, grupo
 * InformeGroup del manifiesto). Con Shared Runtime (ver <Runtimes> en el
 * manifiesto), estas funciones corren en el MISMO contexto JS que
 * taskpane.js (vivo aunque el panel esté oculto, lifetime="long"), así
 * que pueden llamar directamente a TaskPaneApp -mismo patrón que
 * openReportProperties/openFieldOptions, más arriba-, con fallback si
 * por lo que sea no existiera (host sin Shared Runtime).
 * ===================================================================== */

/**
 * "Copiar informe": copia el diseño (filtros/filas/columnas/Estático/
 * opciones de campo) y las Propiedades del informe ACTIVO (el que el
 * taskpane tiene abierto ahora mismo) a un "portapapeles" en Office
 * roaming settings (ReportStore.copyReportToClipboard), listo para
 * "Pegar informe". No toca nada de Excel ni abre el panel (como copiar
 * al portapapeles de Windows, no da ningún aviso si todo va bien).
 */
async function copiarInforme(event) {
    try {
        if (!window.ReportStore) return;
        const reportId = activeReportIdOrNull();
        if (!reportId) {
            console.warn("[Draco] Copiar informe: no hay ningún informe activo.");
            return;
        }
        const copied = await window.ReportStore.copyReportToClipboard(reportId);
        if (copied) {
            console.log(`[Draco] Informe "${copied.name}" copiado al portapapeles.`);
            if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.setAutoStatus) {
                TaskPaneApp.setAutoStatus(`Informe "${copied.name}" copiado`);
            }
        }
    } catch (error) {
        console.error("Error al copiar el informe:", error);
    } finally {
        if (event) event.completed();
    }
}

/**
 * "Pegar informe": crea un informe NUEVO en la hoja activa AHORA MISMO
 * (igual que "Añadir informe"), con el diseño y las propiedades del
 * último informe copiado (ReportStore.pasteReportFromClipboard) — como
 * si se hubiera creado a mano en la celda seleccionada y luego se le
 * hubieran puesto esas propiedades. Se deja el taskpane mostrando ya ese
 * informe y, si tiene diseño (filas/columnas), se refresca de inmediato
 * para que el resultado aparezca pintado sin más pasos.
 */
async function pegarInforme(event) {
    try {
        if (!window.ReportStore) return;

        // Se muestra el panel YA (si estaba oculto): cualquier alert() de
        // aquí en adelante necesita que la ventana esté visible, o el
        // usuario nunca la vería (Excel puede dejarla colgada si no).
        if (Office.addin && Office.addin.showAsTaskpane) await Office.addin.showAsTaskpane();

        if (!window.ReportStore.hasReportClipboard()) {
            alert("No hay ningún informe copiado. Usa antes 'Copiar informe'.");
            return;
        }

        // Hoja activa AHORA MISMO: si hay Shared Runtime se reutiliza
        // captureActiveEditContext (además dejar EDIT_REPORT!D1/E1 al día,
        // por si algo más los necesita); si no, se lee directo con Excel.run.
        let activeSheetName = "";
        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.captureActiveEditContext) {
            const editContext = await TaskPaneApp.captureActiveEditContext();
            activeSheetName = editContext ? editContext.activeSheetName : "";
        }
        if (!activeSheetName) {
            await Excel.run(async (context) => {
                const activeSheet = context.workbook.worksheets.getActiveWorksheet();
                activeSheet.load("name");
                await context.sync();
                activeSheetName = activeSheet.name;
            });
        }

        const pasted = await window.ReportStore.pasteReportFromClipboard(activeSheetName);
        if (!pasted) {
            alert("No se pudo pegar el informe.");
            return;
        }
        console.log(`[Draco] Informe "${pasted.name}" pegado en "${activeSheetName}" (id ${pasted.id}).`);

        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.populateReportSelector) {
            // populateReportSelector() relee el informe ACTIVO de
            // ReportStore (que ya es el recién pegado: createReport lo
            // marca activo) y refresca el desplegable "Informe"; el resto
            // de pasos son los mismos que al elegirlo a mano ahí (ver
            // onReportSelectorChange).
            await TaskPaneApp.populateReportSelector();
            await TaskPaneApp.onReportSelectorChange(String(pasted.id));
            if (TaskPaneApp.setAutoStatus) {
                TaskPaneApp.setAutoStatus(`Informe "${pasted.name}" pegado en "${activeSheetName}"`);
            }
        }

        // Si el diseño copiado ya tiene filas/columnas, se pinta de
        // inmediato (si está vacío, actualizarUnInforme no tiene nada que
        // hacer y no pasa nada por llamarlo igualmente).
        await actualizarUnInforme(pasted.id);
    } catch (error) {
        console.error("Error al pegar el informe:", error);
        await surfaceErrorToSheet(error);
    } finally {
        if (event) event.completed();
    }
}

/**
 * "Eliminar informe": borra la DEFINICIÓN del informe ACTIVO
 * (ReportStore.deleteReport) — filtros/filas/columnas/propiedades, todo
 * lo que vive en Office roaming settings — preguntando antes si también
 * se quiere borrar el CONTENIDO ya pintado en la hoja (los 3 rangos con
 * nombre Draco_<id>_Rows/Cols/Values, ver clearDracoNamedRanges). La
 * definición se borra SIEMPRE; el contenido de la hoja solo si se
 * responde que sí (clear de solo contenido, no de formato).
 */
async function eliminarInforme(event) {
    try {
        if (!window.ReportStore) return;
        const reportId = activeReportIdOrNull();
        if (!reportId) {
            console.warn("[Draco] Eliminar informe: no hay ningún informe activo.");
            return;
        }
        const report = window.ReportStore.getReport(reportId);
        const reportName = report ? report.name : ("Informe " + reportId);

        // Hace falta el panel VISIBLE antes de preguntar: showConfirmDialog
        // es un modal propio del taskpane (no un confirm() nativo, que
        // Excel bloquea/ignora dentro del WebView -ver el comentario en
        // taskpane.js, junto a showConfirmDialog-), así que si el panel
        // estuviera oculto el usuario nunca vería la pregunta.
        if (Office.addin && Office.addin.showAsTaskpane) await Office.addin.showAsTaskpane();

        let clearContent = false;
        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.showConfirmDialog) {
            clearContent = await TaskPaneApp.showConfirmDialog(
                `¿Quieres borrar también el contenido que "${reportName}" tiene pintado en la hoja?`,
                { title: "Eliminar informe", confirmLabel: "Sí, borrar contenido" }
            );
        } else {
            // Sin Shared Runtime (host antiguo, improbable): sin forma
            // fiable de preguntar, se opta por la opción más segura -no
            // tocar el contenido de la hoja- y solo se borra la definición.
            console.warn("[Draco] Eliminar informe: TaskPaneApp no disponible, no se preguntará por el contenido (se conserva).");
        }

        if (clearContent) {
            await Excel.run(async (context) => {
                // false -> Excel.ClearApplyTo.contents (ver clearDracoNamedRanges):
                // borra solo el contenido, no el formato de la hoja.
                await clearDracoNamedRanges(context, reportId, false);
            });
        }

        await window.ReportStore.deleteReport(reportId);
        console.log(`[Draco] Informe "${reportName}" (id ${reportId}) eliminado.` + (clearContent ? " Contenido de la hoja borrado." : ""));

        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.populateReportSelector) {
            await TaskPaneApp.populateReportSelector();
            if (TaskPaneApp.currentReportId) {
                await TaskPaneApp.onReportSelectorChange(String(TaskPaneApp.currentReportId));
            } else if (TaskPaneApp.showEmptyState) {
                // No queda ningún informe: se limpia también el estado en
                // memoria del taskpane (igual que addReport() al crear el
                // primero desde cero), para no dejar filtros/filas/
                // columnas del informe recién eliminado colgando en
                // this.state sin ningún currentReportId al que pertenezcan.
                if (TaskPaneApp.state) {
                    TaskPaneApp.state.filters = [];
                    TaskPaneApp.state.rows = [];
                    TaskPaneApp.state.columns = [];
                    TaskPaneApp.state.rowsStatic = false;
                    TaskPaneApp.state.colsStatic = false;
                    TaskPaneApp.state.fieldOptions = {};
                }
                TaskPaneApp.showEmptyState();
            }
            if (TaskPaneApp.setAutoStatus) {
                TaskPaneApp.setAutoStatus(`Informe "${reportName}" eliminado` + (clearContent ? " (contenido borrado)" : ""));
            }
        }
    } catch (error) {
        console.error("Error al eliminar el informe:", error);
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
    Office.actions.associate("mostrarVistaInformes", mostrarVistaInformes);
    Office.actions.associate("mostrarVistaConexion", mostrarVistaConexion);
    Office.actions.associate("mostrarVistaEditarModelo", mostrarVistaEditarModelo);
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
    Office.actions.associate("filtrarValoresSeleccionados", filtrarValoresSeleccionados);
    Office.actions.associate("campoTotalesMostrar", campoTotalesMostrar);
    Office.actions.associate("campoTotalesOcultar", campoTotalesOcultar);
    Office.actions.associate("campoMostrarID", campoMostrarID);
    Office.actions.associate("campoMostrarIDTexto", campoMostrarIDTexto);
    Office.actions.associate("campoMostrarTexto", campoMostrarTexto);
    Office.actions.associate("campoOpciones", campoOpciones);
    Office.actions.associate("abrirVariables", abrirVariables);
    Office.actions.associate("guardarPlanificacion", guardarPlanificacion);
    Office.actions.associate("guardarExcelEnBucket", guardarExcelEnBucket);
    Office.actions.associate("abrirDesdeBucket", abrirDesdeBucket);
    Office.actions.associate("copiarInforme", copiarInforme);
    Office.actions.associate("pegarInforme", pegarInforme);
    Office.actions.associate("eliminarInforme", eliminarInforme);
    Office.actions.associate("crearInformeVinculado", crearInformeVinculado);   // botón pendiente de añadir al manifest
} catch (e) {
    console.warn("Office.actions.associate no disponible en este contexto:", e);
}
