/**
 * LkmlSaveBridge
 * ------------------------------------------------------------------------
 * El diálogo "Guardar modelo semántico (.lkml)" (saveSemanticModel.html)
 * es una ventana de Office independiente y, como tal, NO tiene acceso al
 * modelo de objetos de Excel (no puede hacer Excel.run). Por eso, cuando
 * el usuario pulsa "Guardar" allí, el diálogo solo construye el texto
 * LookML y lo envía con Office.context.ui.messageParent(); es quien lo
 * abrió (el ribbon, vía js/commands.js, o el taskpane, vía
 * semantic_model.js) quien de verdad tiene acceso a Excel y escribe ese
 * texto en la celda EDIT_REPORT!G1.
 *
 * Se centraliza aquí para que ambos puntos de entrada (ribbon y taskpane)
 * abran el mismo diálogo y escriban en la misma celda de la misma forma.
 */
(function () {

    function openSaveLkmlDialog(models, activeModel) {

        const query = `models=${encodeURIComponent(JSON.stringify(models || {}))}`
            + `&active=${encodeURIComponent(activeModel || "")}`;
        const dialogUrl = new URL(`saveSemanticModel.html?${query}`, window.location.href).href;

        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 45, width: 35, displayInIframe: false },
            (asyncResult) => {

                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Error al abrir el diálogo de guardado de modelo semántico:", asyncResult.error);
                    return;
                }

                const dialog = asyncResult.value;

                // El diálogo envía { content } cuando el usuario pulsa "Guardar".
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {

                    let payload = null;
                    try {
                        payload = JSON.parse(arg.message);
                    } catch (e) {
                        payload = null;
                    }

                    if (!payload || typeof payload.content !== "string") {
                        dialog.close();
                        return;
                    }

                    try {
                        await Excel.run(async (context) => {
                            const sheet = context.workbook.worksheets.getItem("EDIT_REPORT");
                            sheet.getRange("G1").values = [[payload.content]];
                            await context.sync();
                        });
                        dialog.close();
                    } catch (err) {
                        console.error("Error al escribir el LookML en EDIT_REPORT!G1:", err);
                        // Avisamos al diálogo del error para que el usuario lo vea y pueda
                        // reintentar, en vez de cerrarlo sin más explicación.
                        try {
                            dialog.messageChild(JSON.stringify({ error: (err && err.message) || String(err) }));
                        } catch (e2) {
                            console.error("Además, no se ha podido avisar al diálogo del error:", e2);
                            dialog.close();
                        }
                    }

                });

            }
        );

    }

    window.LkmlSaveBridge = { openSaveLkmlDialog };

})();
