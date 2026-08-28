/**
 * Lógica para comandos ejecutados en segundo plano por Excel
 */
Office.onReady(() => {
    // Handshake completado para comandos
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
        const dialogUrl = new URL("openSemanticModel.html", window.location.href).href;

        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 70, width: 45, displayInIframe: false },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Error al abrir el diálogo de apertura de modelo semántico:", asyncResult.error);
                }
            }
        );
    } catch (error) {
        console.error("Error al abrir el diálogo de apertura de modelo semántico:", error);
    } finally {
        if (event) event.completed();
    }

}

/**
 * Botón de ribbon "Guardar modelo semántico" (ModeloGuardarButton).
 * Abre directamente el diálogo independiente de exportación a LookML.
 * Los modelos semánticos viven en Office.context.document.settings
 * (SemanticModelStore), a los que este fichero de comandos SÍ tiene
 * acceso (a diferencia del propio diálogo), así que se le pasan ya
 * "aplanados" en JSON por querystring.
 * @param {Office.AddinCommands.Event} event
 */
function guardarModeloSemantico(event) {

    try {
        const models = window.SemanticModelStore.getAllModels();
        const active = window.SemanticModelStore.getActiveModelName();

        const query = `models=${encodeURIComponent(JSON.stringify(models))}&active=${encodeURIComponent(active)}`;
        const dialogUrl = new URL(`saveSemanticModel.html?${query}`, window.location.href).href;

        Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 75, width: 45, displayInIframe: false },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Error al abrir el diálogo de guardado de modelo semántico:", asyncResult.error);
                }
            }
        );
    } catch (error) {
        console.error("Error al abrir el diálogo de guardado de modelo semántico:", error);
    } finally {
        if (event) event.completed();
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

// Asociar el nombre del comando del manifiesto con la función JavaScript
Office.actions.associate("hidePane", hidePane);
Office.actions.associate("abrirModeloSemantico", abrirModeloSemantico);
Office.actions.associate("guardarModeloSemantico", guardarModeloSemantico);