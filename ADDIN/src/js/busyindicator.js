/**
 * busyIndicator.js — indicador visual "Actualizando…" como un shape
 * flotante DENTRO de la propia hoja de Excel (no en el taskpane, ver
 * setAutoStatus en taskpane.js para ese otro indicador de texto), con un
 * icono giratorio que se anima mientras dura la operación y se borra solo
 * al terminar.
 * ------------------------------------------------------------------------
 * Uso:
 *   await BusyIndicator.show();              // usa la etiqueta por defecto
 *   await BusyIndicator.show("Guardando");   // etiqueta personalizada
 *   ... trabajo largo (Excel.run, fetch a BigQuery, etc.) ...
 *   await BusyIndicator.hide();
 *
 * show()/hide() llevan la cuenta de cuántas operaciones lo tienen abierto
 * a la vez (refCount): si dos procesos distintos llaman a show() solapados,
 * el shape no desaparece hasta que AMBOS hayan llamado a hide(). Así no
 * hace falta coordinar manualmente quién "es el dueño" del indicador.
 *
 * LIMITACIÓN IMPORTANTE (API de Office): no existe en Excel JS API un
 * equivalente a Application.StatusBar de VBA, ni tampoco una forma de leer
 * la posición de scroll/ventana visible (no hay "ActiveWindow.VisibleRange").
 * Por eso "abajo a la izquierda" aquí es una posición fija en PUNTOS desde
 * la esquina superior izquierda de LA HOJA (A1), no de lo que se ve en
 * pantalla en cada momento: si el usuario ha hecho scroll lejos de A1, el
 * indicador puede quedar fuera de la vista. Ajusta INDICATOR_TOP más abajo
 * si tus hojas suelen verse con más o menos scroll por defecto. Si
 * necesitas garantía de visibilidad siempre, la alternativa robusta sigue
 * siendo el texto en el propio taskpane (setAutoStatus).
 */
(function () {

    const SHAPE_NAME = "EPM_BusyIndicator";
    const FRAME_MS = 150;

    // Spinner "círculo giratorio" — se ve como un circulito rotando, más
    // vistoso que | / - \ y con buen soporte de fuente en Windows/Mac.
    const FRAMES = ["◐", "◓", "◑", "◒"];
    // Alternativa más suave (puntos Braille tipo CLI): descomenta para
    // probarla en vez de la de arriba.
    // const FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"];

    const INDICATOR_LEFT = 8;    // puntos desde el borde izquierdo de la hoja
    const INDICATOR_TOP = 520;   // puntos desde el borde superior de la hoja
    const INDICATOR_WIDTH = 140;
    const INDICATOR_HEIGHT = 24;

    let refCount = 0;
    let timerId = null;
    let frameIdx = 0;
    let label = "Actualizando";
    let sheetName = null; // hoja donde se creó el shape (se fija en show())

    function frameText() {
        return `${FRAMES[frameIdx]}  ${label}…`;
    }

    async function createShape(context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await context.sync();
        sheetName = sheet.name;

        const shape = sheet.shapes.addTextBox(frameText());
        shape.name = SHAPE_NAME;
        shape.left = INDICATOR_LEFT;
        shape.top = INDICATOR_TOP;
        shape.width = INDICATOR_WIDTH;
        shape.height = INDICATOR_HEIGHT;
        shape.lockAspectRatio = false;

        shape.fill.setSolidColor("#323130");
        shape.fill.transparency = 0.08;
        shape.lineFormat.visible = false;

        shape.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
        shape.textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;
        shape.textFrame.textRange.font.color = "#FFFFFF";
        shape.textFrame.textRange.font.size = 11;
        shape.textFrame.textRange.font.bold = true;

        await context.sync();
    }

    async function findShape(context) {
        if (!sheetName) return null;
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        const shapes = sheet.shapes;
        shapes.load("items/name");
        await context.sync();
        if (sheet.isNullObject) return null; // la hoja se borró mientras tanto
        return shapes.items.find(s => s.name === SHAPE_NAME) || null;
    }

    async function tick() {
        frameIdx = (frameIdx + 1) % FRAMES.length;
        try {
            await Excel.run(async (context) => {
                const shape = await findShape(context);
                if (!shape) return; // hide() ya lo quitó, o el usuario lo borró a mano
                shape.textFrame.textRange.text = frameText();
                await context.sync();
            });
        } catch (err) {
            // No se interrumpe la operación en curso por esto: si un tick
            // falla (p.ej. la hoja activa cambió a una protegida), se
            // reintenta en el siguiente; hide() limpia igualmente al acabar.
            console.warn("[BusyIndicator] No se pudo animar el indicador:", err);
        }
    }

    /**
     * Muestra el indicador (crea el shape si no existía ya). Si ya estaba
     * visible por otra operación en curso, solo incrementa el contador de
     * referencias — no reinicia la animación ni la etiqueta.
     */
    async function show(customLabel) {
        refCount++;
        if (customLabel) label = customLabel;
        if (refCount > 1) return; // ya lo estaba mostrando otra llamada

        try {
            await Excel.run(async (context) => {
                const existing = await findShape(context);
                if (!existing) await createShape(context);
            });
        } catch (err) {
            console.warn("[BusyIndicator] No se pudo crear el indicador:", err);
            return; // sin shape no tiene sentido animar
        }

        frameIdx = 0;
        if (!timerId) timerId = setInterval(tick, FRAME_MS);
    }

    /**
     * Oculta el indicador (borra el shape) — solo cuando NINGUNA operación
     * en curso lo sigue necesitando (refCount llega a 0).
     */
    async function hide() {
        refCount = Math.max(0, refCount - 1);
        if (refCount > 0) return;

        if (timerId) { clearInterval(timerId); timerId = null; }

        try {
            await Excel.run(async (context) => {
                const shape = await findShape(context);
                if (shape) {
                    shape.delete();
                    await context.sync();
                }
            });
        } catch (err) {
            console.warn("[BusyIndicator] No se pudo eliminar el indicador:", err);
        } finally {
            sheetName = null;
        }
    }

    window.BusyIndicator = { show, hide };

})();
