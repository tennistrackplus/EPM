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

        const dialogUrl = new URL("openSemanticModel.html?v=" + encodeURIComponent((typeof DRACO_BUILD !== "undefined" ? DRACO_BUILD : "1")), window.location.href).href;

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
                    // Petición del token de Git de la sesión (ver js/gitSession.js, F-06)
                    if (window.GitSession && GitSession.handleDialogMessage(dialog, arg.message)) return;

                    let payload = null;
                    try {
                        payload = JSON.parse(arg.message);
                    } catch (e) {
                        payload = null;
                    }

                    // Pestaña Databricks del diálogo: listados de catálogos,
                    // esquemas y vistas. Se hacen AQUÍ (con la sesión de
                    // Databricks del panel/cinta): el token nunca va al diálogo.
                    if (payload && payload.dbxRequest) {
                        let reply;
                        try {
                            reply = { dbxResponse: payload.dbxRequest, result: await dbxList(payload.op, payload.args || {}) };
                        } catch (err) {
                            reply = { dbxResponse: payload.dbxRequest, error: (err && err.message) || String(err) };
                        }
                        try { dialog.messageChild(JSON.stringify(reply)); } catch (e) { console.error("No se pudo responder al diálogo:", e); }
                        return;
                    }

                    // Abrir una Metric View creada por Draco: se lee aquí su
                    // modelo (propiedad draco.model) y se guarda como cualquier
                    // otro modelo importado.
                    if (payload && payload.dbxOpen && payload.modelName) {
                        try {
                            const o = payload.dbxOpen;
                            const r = await window.DracoMetricView.readModel(o.catalog, o.schema, o.view);
                            payload.model = r.model;
                            // Lo que no se ha podido traer se enseña en el diálogo antes de cerrarlo
                            if (r.warnings && r.warnings.length) {
                                try { dialog.messageChild(JSON.stringify({ info: "Importado con avisos: " + r.warnings.slice(0, 4).join(" · ")
                                    + (r.warnings.length > 4 ? ` (y ${r.warnings.length - 4} más)` : "") })); } catch (e) { /* sin aviso */ }
                                await new Promise(res => setTimeout(res, 4000));
                            }
                        } catch (err) {
                            console.error("No se pudo abrir la Metric View:", err);
                            try { dialog.messageChild(JSON.stringify({ error: (err && err.message) || String(err) })); } catch (e2) { dialog.close(); }
                            return;
                        }
                    }

                    if (!payload || !payload.modelName || !payload.model) {
                        dialog.close();
                        return;
                    }

                    try {

                        await window.SemanticModelStore.saveModel(payload.modelName, payload.model);
                        await window.SemanticModelStore.setActiveModelName(payload.modelName);

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
                        return;
                    }

                    dialog.close();

                    // Fuera del try anterior: un fallo al refrescar la UI no debe
                    // tratarse como "no se ha podido guardar" (el diálogo ya está
                    // cerrado y el aviso se perdía sin que el usuario viera nada).
                    if (typeof onImported === "function") {
                        try {
                            await onImported(payload.modelName);
                        } catch (uiErr) {
                            console.error("Modelo importado, pero falló el refresco del panel:", uiErr);
                        }
                    }

                });

            }
        );

    }

    /** Listados para la pestaña Databricks del diálogo (solo nombres). */
    async function dbxList(op, args) {
        if (!window.DBX || !(window.Provider && Provider.isDatabricks && Provider.isDatabricks())) {
            throw new Error("Activa antes una conexión de Databricks en la pestaña Conexión.");
        }
        if (op === "catalogs") return await DBX.listCatalogs();
        if (op === "schemas") return await DBX.listSchemas(String(args.catalog || ""));
        if (op === "views") return await DBX.listMetricViews(String(args.catalog || ""), String(args.schema || ""));
        throw new Error("Operación no válida.");
    }

    window.LkmlOpenBridge = { openOpenLkmlDialog };

})();
