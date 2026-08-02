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

/**
 * Función que maneja el botón 'Escribir HOLA' (WriteHolaButton) definido en el manifiesto
 * @param {Office.AddinCommands.Event} event
 */
async function writeHolaInA1(event) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("A1");
            
            range.values = [["HOLA"]];
            
            await context.sync();
        });
    } catch (error) {
        console.error("Error al escribir en la celda A1:", error);
    } finally {
        // OBLIGATORIO: Informar a Excel que la función finalizó para no bloquear el runtime
        if (event) {
            event.completed();
        }
    }
}

// Asociar el nombre del comando del manifiesto con la función JavaScript
Office.actions.associate("hidePane", hidePane);
Office.actions.associate("writeHolaInA1", writeHolaInA1);