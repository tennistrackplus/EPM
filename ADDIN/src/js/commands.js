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

// Asociar el nombre del comando del manifiesto con la función JavaScript
Office.actions.associate("hidePane", hidePane);
Office.actions.associate("writeHolaInA1", writeHolaInA1);