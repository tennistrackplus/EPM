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
 * Por eso, en vez de una posición fija en puntos desde A1 (que puede caer
 * fuera de lo que el usuario ve en pantalla según el tamaño de su ventana
 * — así falló la primera versión de este archivo, el shape se creaba pero
 * quedaba invisible más abajo del borde inferior de la ventana), el
 * indicador se ancla justo debajo de la CELDA ACTIVA: al dispararse la
 * actualización (doble clic en un filtro, editar el diseño...) esa celda
 * casi siempre está dentro de lo que el usuario tiene visible, así que es
 * la mejor aproximación disponible a "donde está mirando ahora mismo". Si
 * aun así no apareciera, revisa la consola del navegador/Excel: show()
 * registra con console.error cualquier fallo al crear el shape (p.ej. si
 * el host de Excel no soporta la API de Shapes, ExcelApi 1.9).
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

        // Antes: posición fija (INDICATOR_LEFT/INDICATOR_TOP) medida desde
        // A1. Con una ventana de Excel que no llegue a mostrar esa fila, el
        // shape se crea igualmente pero queda FUERA de lo que se ve en
        // pantalla — sin ningún error, simplemente no se ve. La celda
        // activa, en cambio, casi siempre está dentro de lo que el usuario
        // tiene visible en el momento de disparar la actualización (doble
        // clic en un filtro, editar el diseño, etc.), así que anclamos el
        // indicador justo debajo de ella en vez de a un punto fijo de la
        // hoja.
        const activeCell = context.workbook.getActiveCell();
        activeCell.load(["left", "top", "height"]);

        await context.sync();
        sheetName = sheet.name;

        const left = Math.max(4, activeCell.left);
        const top = activeCell.top + activeCell.height + 4;

        const shape = sheet.shapes.addTextBox(frameText());
        shape.name = SHAPE_NAME;
        shape.left = left;
        shape.top = top;
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

    // Duración mínima que se deja el indicador visible una vez creado,
    // aunque la operación real termine antes: si el refresco tarda p.ej.
    // 100ms, el shape se crearía y se borraría casi en el mismo instante,
    // dando la sensación de que nunca llegó a aparecer.
    const MIN_VISIBLE_MS = 400;
    let shownAt = 0;

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
            console.error("[BusyIndicator] No se pudo crear el indicador:", err);
            return; // sin shape no tiene sentido animar
        }

        shownAt = Date.now();
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

        const elapsed = Date.now() - shownAt;
        if (elapsed < MIN_VISIBLE_MS) {
            await new Promise(resolve => setTimeout(resolve, MIN_VISIBLE_MS - elapsed));
        }

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
