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

// Asociar el nombre del comando del manifiesto con la función JavaScript
Office.actions.associate("hidePane", hidePane);