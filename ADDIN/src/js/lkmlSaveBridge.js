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

        // Los modelos ya NO van en la URL: un modelo real (decenas de campos y
        // jerarquías) son decenas de KB y Office rechazaba el diálogo con
        // «URL demasiado larga». El diálogo los pide al arrancar
        // (MODELS_REQUEST) y se le mandan por mensajes, troceados.
        const modelsJson = JSON.stringify(models || {});
        const query = `active=${encodeURIComponent(activeModel || "")}&byMessage=1`;
        const dialogUrl = new URL(`saveSemanticModel.html?${query}`, window.location.href).href;
        const CHUNK = 20000;

        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 75, width: 45, displayInIframe: false },
            (asyncResult) => {

                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Error al abrir el diálogo de guardado de modelo semántico:", asyncResult.error);
                    return;
                }

                const dialog = asyncResult.value;

                // El diálogo envía { content } cuando el usuario pulsa "Guardar".
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                    // Petición del token de Git de la sesión (ver js/gitSession.js, F-06)
                    if (window.GitSession && GitSession.handleDialogMessage(dialog, arg.message)) return;

                    let payload = null;
                    try {
                        payload = JSON.parse(arg.message);
                    } catch (e) {
                        payload = null;
                    }

                    // El diálogo pide los modelos semánticos: se mandan en trozos
                    if (payload && payload.type === "draco-save-models-request") {
                        const n = Math.max(1, Math.ceil(modelsJson.length / CHUNK));
                        for (let i = 0; i < n; i++) {
                            dialog.messageChild(JSON.stringify({
                                type: "draco-save-models-chunk", i, n,
                                data: modelsJson.slice(i * CHUNK, (i + 1) * CHUNK)
                            }));
                        }
                        return;
                    }

                    if (!payload || typeof payload.content !== "string") {
                        dialog.close();
                        return;
                    }

                    try {
                        await Excel.run(async (context) => {
                            const sheet = context.workbook.worksheets.getItem("EDIT_REPORT");
                            // Texto, nunca fórmula (auditoría F-02)
                            const content = /^[=+\-@\t\r]/.test(payload.content) ? "'" + payload.content : payload.content;
                            sheet.getRange("G1").values = [[content]];
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
