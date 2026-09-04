/**
 * GcsSaveBridge
 * ------------------------------------------------------------------------
 * El diálogo "Guardar en bucket" (saveBucket.html) es una ventana de
 * Office independiente y, como tal, NO tiene acceso a Office.context.document
 * (no puede leer el .xlsx activo). Por eso el diálogo solo deja elegir
 * proyecto, bucket y nombre de fichero, y envía esa elección con
 * Office.context.ui.messageParent(); es quien lo abrió (el ribbon, vía
 * js/commands.js, o el taskpane, vía js/taskpane.js) quien de verdad tiene
 * acceso al documento y hace la subida real con js/gcsExport.js.
 *
 * Se centraliza aquí para que ambos puntos de entrada (ribbon y taskpane)
 * abran el mismo diálogo y suban el fichero de la misma forma.
 */
(function () {

    /**
     * Abre el diálogo de selección de bucket.
     * @param {(selection: {bucket: string, objectName: string}) => void} onSelected
     *        Se llama cuando el usuario confirma con una elección válida.
     *        El diálogo ya se ha cerrado en ese momento.
     */
    function openSaveBucketDialog(onSelected) {
        const suggested = (window.GCS && GCS.getSuggestedFileName && GCS.getSuggestedFileName()) || "Informe_EPM.xlsx";
        const url = new URL("saveBucket.html", window.location.href);
        url.searchParams.set("name", suggested);
        // Ver BQ.getSessionQueryParams(): el diálogo se abre en su propia
        // ventana y puede no compartir localStorage con este runtime, así
        // que le pasamos el token vigente por la URL para que no pida
        // conectarse de nuevo estando ya conectado.
        const sessionParams = window.BQ ? BQ.getSessionQueryParams() : "";
        if (sessionParams) {
            new URLSearchParams(sessionParams).forEach((value, key) => url.searchParams.set(key, value));
        }

        Office.context.ui.displayDialogAsync(
            url.href,
            { height: 55, width: 40, displayInIframe: false },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("No se pudo abrir el diálogo de guardado en bucket:", asyncResult.error.message);
                    return;
                }

                const dialog = asyncResult.value;

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                    let payload = null;
                    try {
                        payload = JSON.parse(arg.message);
                    } catch (e) {
                        payload = null;
                    }

                    dialog.close();

                    if (!payload || payload.cancelled) return;
                    if (!payload.bucket || !payload.objectName) return;

                    onSelected({ bucket: payload.bucket, objectName: payload.objectName });
                });
            }
        );
    }

    window.GcsSaveBridge = { openSaveBucketDialog };

})();
