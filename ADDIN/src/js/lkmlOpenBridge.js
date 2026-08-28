/**
 * LkmlOpenBridge
 * ------------------------------------------------------------------------
 * El diálogo "Abrir modelo semántico (.lkml)" (openSemanticModel.html) es
 * una ventana de Office independiente y, como tal, NO tiene acceso a
 * Office.context.document.settings (SemanticModelStore, ver
 * js/semanticModelStore.js). Por eso, cuando el usuario elige un fichero
 * .lkml, indica el nombre del modelo y pulsa "Seleccionar", el diálogo solo
 * lee el fichero, lo parsea con LkmlImport (ver js/lkmlImport.js) y envía
 * el resultado con Office.context.ui.messageParent(); es quien lo abrió (el
 * ribbon, vía js/commands.js, o el taskpane, vía semantic_model.js) quien
 * de verdad tiene acceso a SemanticModelStore y guarda el modelo semántico
 * importado.
 *
 * Se centraliza aquí para que ambos puntos de entrada (ribbon y taskpane)
 * abran el mismo diálogo y guarden el modelo de la misma forma. Simétrico a
 * js/lkmlSaveBridge.js.
 */
(function () {

    /**
     * @param {(modelName: string) => void} [onImported] Callback opcional,
     * invocado tras guardar con éxito el modelo importado (el taskpane lo
     * usa para refrescar su UI; el ribbon no necesita pasar nada).
     */
    function openOpenLkmlDialog(onImported) {

        const dialogUrl = new URL("openSemanticModel.html", window.location.href).href;

        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 72, width: 45, displayInIframe: false },
            (asyncResult) => {

                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Error al abrir el diálogo de apertura de modelo semántico:", asyncResult.error);
                    return;
                }

                const dialog = asyncResult.value;

                // El diálogo envía { modelName, model } cuando el usuario
                // pulsa "Seleccionar" y ha podido leer/parsear el fichero.
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {

                    let payload = null;
                    try {
                        payload = JSON.parse(arg.message);
                    } catch (e) {
                        payload = null;
                    }

                    if (!payload || !payload.modelName || !payload.model) {
                        dialog.close();
                        return;
                    }

                    try {

                        await window.SemanticModelStore.saveModel(payload.modelName, payload.model);
                        await window.SemanticModelStore.setActiveModelName(payload.modelName);

                        dialog.close();

                        if (typeof onImported === "function") {
                            onImported(payload.modelName);
                        }

                    } catch (err) {
                        console.error("Error al guardar el modelo semántico importado:", err);
                        // Avisamos al diálogo del error para que el usuario lo vea y
                        // pueda reintentar, en vez de cerrarlo sin más explicación.
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

    window.LkmlOpenBridge = { openOpenLkmlDialog };

})();
