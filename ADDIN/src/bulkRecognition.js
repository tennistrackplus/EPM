/**
 * ============================================================
 * DRACO — RECONOCIMIENTO MASIVO DE MIEMBROS (pegar varias celdas)
 * ============================================================
 * El reconocimiento individual (teclear UN valor -> buscador) sigue igual
 * y NO pasa por aquí. Este flujo es aparte, con su propio diálogo
 * (memberPickerBulk.html); de commands.js solo REUTILIZA funciones, sin
 * cambiarlas:
 *   isDracoMemberRecognitionActive, findDracoRowsNamedRangeForCell,
 *   resolveDracoFieldForAxisLevel, dracoStaticAxisTarget,
 *   dracoAddStaticLevelAndLocate, loadDracoPickerItems,
 *   buildEpmValueFormula, growDracoNamedRange, getDracoResultSheetName,
 *   dracoRebuildStaticNames, displayDracoDialogOnce.
 *
 * Flujo (con el reconocimiento de miembros activo):
 *  1. Se pega un rango. Se quedan las celdas con valor que NO son fórmula
 *     y caen en Filas/Columnas de un informe (o en un eje Estático).
 *  2. Ventana: «Se han pegado N valores… ¿continuar?».
 *  3. Se va por grupos (una columna de Filas o una fila de Columnas =
 *     una dimensión). Si el grupo es un nivel nuevo (p. ej. una columna a
 *     la izquierda del eje Estático), se pide la dimensión con el mismo
 *     selector que el reconocimiento individual.
 *  4. Celda a celda: si el valor coincide con UN solo miembro (exacto por
 *     ID, texto o «ID - texto»; o, si no hay exacto, un único resultado al
 *     filtrar como el buscador), se escribe la EPM_VALUE sin preguntar. Si
 *     no, se selecciona la celda en la hoja y se muestra el buscador ya
 *     filtrado por su valor: elegir, «Omitir celda» o «Detener».
 *  5. La ventana enseña «Cargando…» mientras trabaja y solo muestra la
 *     lista cuando hace falta elegir; si todo se reconoce, se cierra sola.
 * ============================================================
 */
(function () {
    "use strict";

    const TAG = "[BULK]";
    const CHUNK = 16000;
    let active = false;                 // hay un reconocimiento masivo en marcha

    /** ¿Está en marcha? (commands.js ignora mientras tanto las peticiones individuales) */
    function isActive() { return active; }

    // ---------------------------------------------------------------
    // Ventana (una sola para todo el proceso; se cierra y se vuelve a
    // abrir solo si hay que pedir una dimensión, porque Office solo
    // permite un diálogo a la vez).
    // ---------------------------------------------------------------
    function createWindow() {
        const w = {
            dlg: null, ready: false, closedByUser: false, displayMode: "ID_TEXT",
            waiters: [], lastItems: null, round: 0, sentKey: ""
        };
        const deliver = (msg) => {
            const i = w.waiters.findIndex(x => x.match(msg));
            if (i >= 0) { const x = w.waiters.splice(i, 1)[0]; x.resolve(msg); }
        };
        w.open = async () => {
            const sinceClose = Date.now() - (typeof DracoLastDialogClosedAt === "number" ? DracoLastDialogClosedAt : 0);
            if (sinceClose < 500) await dracoSleep(500 - sinceClose);
            const url = new URL("memberPickerBulk.html", window.location.href).href + "?v=" + encodeURIComponent(DRACO_BUILD);
            let res = null;
            for (let attempt = 1; attempt <= 6; attempt++) {
                res = await displayDracoDialogOnce(url, { height: 60, width: 30, displayInIframe: false });
                if (res.status !== Office.AsyncResultStatus.Failed) break;
                const code = res.error && res.error.code;
                dracoLog(TAG, "no se pudo abrir la ventana (intento " + attempt + "): " + code);
                if (code !== 12007) break;
                await dracoSleep(300 * attempt);
            }
            if (!res || res.status === Office.AsyncResultStatus.Failed) return false;
            w.dlg = res.value;
            w.ready = false;
            w.sentKey = "";
            const readyP = new Promise((resolve) => w.waiters.push({ match: m => m.type === "ready", resolve }));
            w.dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                let m = null;
                try { m = JSON.parse(arg.message); } catch (e) { return; }
                if (!m) return;
                if (m.type === "log") { dracoLog(TAG, "[ventana] " + m.msg); return; }
                if (m.type === "ready") {
                    if (m.displayMode) w.displayMode = m.displayMode;
                    if (w.ready) return;
                    w.ready = true;
                }
                if (m.type === "displayMode") { w.displayMode = m.displayMode || w.displayMode; return; }
                if (m.type === "resend") { if (w.lastItems && w.lastItems.fieldKey === m.fieldKey) w.sendItems(w.lastItems, true); return; }
                if (m.type === "stop") { w.closedByUser = true; }
                deliver(m);
            });
            w.dlg.addEventHandler(Office.EventType.DialogEventReceived, () => {
                // Cerrada con la X: se detiene todo
                w.closedByUser = true;
                w.dlg = null;
                deliver({ type: "stop", byClose: true });
            });
            const ok = await Promise.race([readyP.then(() => true), dracoSleep(20000).then(() => false)]);
            if (!ok) dracoLog(TAG, "la ventana no ha respondido en 20 s");
            w.send({ type: "hello" });
            return ok;
        };
        w.send = (obj) => {
            if (!w.dlg) return;
            try { w.dlg.messageChild(JSON.stringify(obj)); } catch (e) { dracoLog(TAG, "messageChild: " + dracoErrMsg(e)); }
        };
        /** Espera el siguiente mensaje que cumpla match (o «stop»). */
        w.wait = (match) => new Promise((resolve) => {
            if (w.closedByUser) return resolve({ type: "stop" });
            w.waiters.push({ match: (m) => m.type === "stop" || match(m), resolve });
        });
        w.status = (text, progress) => w.send({ type: "status", text, progress: progress || "" });
        w.sendItems = (pack, force) => {
            if (!w.dlg || (!force && w.sentKey === pack.fieldKey)) return;
            w.lastItems = pack;
            w.sentKey = pack.fieldKey;
            const body = JSON.stringify(pack.items);
            const parts = [];
            for (let i = 0; i < body.length; i += CHUNK) parts.push(body.slice(i, i + CHUNK));
            const round = ++w.round;
            w.send({ type: "begin", round, total: parts.length, fieldKey: pack.fieldKey, title: pack.title });
            parts.forEach((part, seq) => w.send({ type: "chunk", round, seq, part }));
        };
        w.close = () => {
            if (w.dlg) {
                try { w.dlg.close(); } catch (e) { /* ya cerrada */ }
                w.dlg = null;
                if (typeof DracoLastDialogClosedAt === "number") DracoLastDialogClosedAt = Date.now();
            }
        };
        return w;
    }

    // ---------------------------------------------------------------
    // Utilidades
    // ---------------------------------------------------------------
    function colLetter(n) {          // 1 -> A
        let s = "";
        while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
        return s;
    }
    const addrOf = (row0, col0) => colLetter(col0 + 1) + (row0 + 1);
    const norm = (v) => String(v === null || v === undefined ? "" : v).trim().toLowerCase();

    /**
     * Miembros que encajan con el texto de la celda:
     *  - exactos: ID, texto o «ID - texto» iguales (sin mayúsculas/espacios);
     *  - si no hay exactos, los que salen al filtrar como el buscador (contiene).
     */
    function findMatches(items, text) {
        const t = norm(text);
        if (!t) return { exact: [], contains: [] };
        const exact = [], contains = [];
        for (const it of items) {
            const id = norm(it.value), lbl = norm(it.label), txt = norm(it.text);
            if (id === t || (lbl && lbl === t) || txt === t || (lbl && norm(it.value + " - " + it.label) === t)) exact.push(it);
            else if (id.indexOf(t) !== -1 || (lbl && lbl.indexOf(t) !== -1) || txt.indexOf(t) !== -1) contains.push(it);
        }
        return { exact, contains };
    }

    /** Lo que se ve en la celda según el modo del buscador (el valor es siempre el ID). */
    function shownFor(it, mode) {
        const id = String(it.value), txt = it.label ? String(it.label) : "";
        if (!txt || mode === "ID") return undefined;
        if (mode === "TEXT") return txt;
        return id + " - " + txt;
    }

    const fieldKeyOf = (f) => [f.reportId, f.dim, f.attr].join("|");

    /**
     * Destino de UNA celda, con la misma lógica que el reconocimiento
     * individual (runDracoMemberRecognitionAction), pero sin seleccionar
     * la celda ni abrir nada: { fieldLocated } | { special } | null.
     */
    async function resolveCell(context, sheet, addr, text, indent) {
        const located = await findDracoRowsNamedRangeForCell(context, sheet.id, addr);
        if (!located) {
            const st = await dracoStaticAxisTarget(context, sheet, addr, text, TAG);
            if (st && st.fieldLocated) st.fieldLocated.isStatic = true;
            return st || null;
        }
        const f = await resolveDracoFieldForAxisLevel(context, located.reportId, located.axis, located.level, indent || 0);
        if (!f) return null;
        f.reportId = located.reportId;
        f.growth = located.growth || null;
        return { addr, fieldLocated: f, axis: located.axis };
    }

    /** Escribe la EPM_VALUE como el buscador individual (ampliando el rango si hace falta). */
    async function writeMember(sheetName, addr, f, attribute, value, shown) {
        await Excel.run(async (context) => {
            if (f.growth) {
                const g = f.growth;
                await growDracoNamedRange(context, g.sheetName, g.rangeName, g.axisSuffix,
                    g.axisStart, g.requiredCount, g.crossIndex, g.crossCount);
            }
            const resultSheetName = await getDracoResultSheetName(context, f.reportId);
            const sheet = context.workbook.worksheets.getItem(resultSheetName || sheetName);
            sheet.getRange(addr).values = [[buildEpmValueFormula(f.dim, attribute, value, shown)]];
            await context.sync();
            // Eje Estático sin rangos con nombre: hay que rehacerlos para que la siguiente celda se encuentre
            if (f.isStatic) {
                try { await dracoRebuildStaticNames(context, f.reportId); }
                catch (e) { dracoLog(TAG, "no se pudieron ajustar los rangos: " + dracoErrMsg(e)); }
            }
        });
    }

    // ---------------------------------------------------------------
    // Proceso principal
    // ---------------------------------------------------------------
    /**
     * Punto de entrada (desde handleDracoSheetEdited en commands.js al
     * pegar/rellenar un rango). No hace nada si el reconocimiento de
     * miembros está desactivado o ninguna celda pegada es de un informe.
     */
    async function run(sheetName, rangeAddr) {
        if (active) { dracoLog(TAG, "ya hay un reconocimiento masivo en marcha: se ignora " + sheetName + "!" + rangeAddr); return; }
        active = true;
        const w = createWindow();
        const touchedReports = new Set();
        const stats = { auto: 0, chosen: 0, skipped: 0, total: 0 };
        const itemsByField = new Map();   // fieldKey -> miembros (una consulta/parseo por campo)
        try {
            // 1) Celdas candidatas y grupos (columna de Filas / fila de Columnas)
            let sheetId = null, groups = [], needMeasure = false;
            await Excel.run(async (context) => {
                if (!(await isDracoMemberRecognitionActive(context))) { dracoLog(TAG, "reconocimiento desactivado"); return; }
                const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                sheet.load(["isNullObject", "id"]);
                await context.sync();
                if (sheet.isNullObject) return;
                sheetId = sheet.id;
                const rng = sheet.getRange(rangeAddr);
                rng.load(["values", "formulas", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
                let props = null;
                try { props = rng.getCellProperties({ format: { indentLevel: true } }); } catch (e) { props = null; }
                await context.sync();
                if (rng.rowCount * rng.columnCount > 5000) {
                    dracoLog(TAG, "rango demasiado grande (" + rng.rowCount * rng.columnCount + " celdas): se ignora");
                    return;
                }
                const cells = [];
                for (let r = 0; r < rng.rowCount; r++) {
                    for (let c = 0; c < rng.columnCount; c++) {
                        const v = rng.values[r][c], fml = rng.formulas[r][c];
                        if (v === "" || v === null || v === undefined) continue;
                        if (typeof fml === "string" && /^\s*=/.test(fml)) continue;   // fórmulas: no se tocan
                        const indent = props && props.value && props.value[r] && props.value[r][c] && props.value[r][c].format
                            ? (props.value[r][c].format.indentLevel || 0) : 0;
                        cells.push({ row0: rng.rowIndex + r, col0: rng.columnIndex + c, addr: addrOf(rng.rowIndex + r, rng.columnIndex + c), text: String(v), indent });
                    }
                }
                if (!cells.length) return;

                // Un grupo por columna (eje Filas) o por fila (eje Columnas): se
                // averigua con la primera celda de cada columna y de cada fila.
                const byCol = new Map(), byRow = new Map();
                cells.forEach(cl => {
                    if (!byCol.has(cl.col0)) byCol.set(cl.col0, []);
                    byCol.get(cl.col0).push(cl);
                    if (!byRow.has(cl.row0)) byRow.set(cl.row0, []);
                    byRow.get(cl.row0).push(cl);
                });
                const taken = new Set();
                const probe = async (list, axisWanted) => {
                    const first = list.find(cl => !taken.has(cl.addr));
                    if (!first) return;
                    const t = await resolveCell(context, sheet, first.addr, first.text, first.indent);
                    if (!t) return;
                    if (t.special === "needMeasure") { needMeasure = true; return; }
                    const axis = t.special === "needDimension" ? t.axis : (t.axis || (t.fieldLocated && t.fieldLocated.axis) || null);
                    if (axis !== axisWanted) return;   // esta columna/fila es del otro eje: la recoge el otro recorrido
                    const members = list.filter(cl => !taken.has(cl.addr));
                    members.forEach(cl => taken.add(cl.addr));
                    groups.push({ axis: axisWanted, first: t, cells: members,
                        reportId: t.special ? t.reportId : t.fieldLocated.reportId,
                        pos: axisWanted === "rows" ? first.col0 : first.row0 });
                };
                for (const [, list] of [...byCol.entries()].sort((a, b) => a[0] - b[0])) await probe(list, "rows");
                for (const [, list] of [...byRow.entries()].sort((a, b) => a[0] - b[0])) await probe(list, "columns");
            });

            if (needMeasure && !groups.length) {
                dracoInfoNotice("Falta una medida", "Es necesario que exista al menos una medida en filas, columnas o filtros.", "warn");
                return;
            }
            if (!groups.length) { dracoLog(TAG, "ninguna celda pegada es de un informe: no se hace nada"); return; }
            stats.total = groups.reduce((a, g) => a + g.cells.length, 0);
            dracoLog(TAG, `${stats.total} valores pegados en ${sheetName}!${rangeAddr}, ${groups.length} grupo(s)`);

            // 2) Confirmación
            if (typeof DracoCurrentPicker !== "undefined" && DracoCurrentPicker) await DracoCurrentPicker.close("reconocimiento masivo");
            await ensureDracoFnMode();
            if (!(await w.open())) { dracoInfoNotice("Reconocimiento de miembros", "No se pudo abrir la ventana de reconocimiento.", "warn"); return; }
            const dims = groups.length === 1 ? "una dimensión" : groups.length + " dimensiones";
            w.send({ type: "confirm", text: `Se han pegado ${stats.total} valores en ${dims} del informe. ¿Seguro que quieres reconocerlos?` });
            const conf = await w.wait(m => m.type === "confirm");
            if (conf.type !== "confirm" || !conf.ok) { dracoLog(TAG, "cancelado en la confirmación"); return; }

            // 3) Grupo a grupo
            let done = 0;
            for (const g of groups) {
                if (w.closedByUser) break;
                const progressOf = () => `${done}/${stats.total}`;

                // Nivel nuevo: se pide la dimensión (mismo selector que el individual)
                if (g.first.special === "needDimension") {
                    w.close();
                    const fl = await dracoAddStaticLevelAndLocate(g.first);
                    if (!fl) {
                        dracoLog(TAG, "no se eligió dimensión: se omite el grupo");
                        stats.skipped += g.cells.length; done += g.cells.length;
                        if (!(await w.open())) break;
                        continue;
                    }
                    if (!(await w.open())) break;
                }

                for (const cl of g.cells) {
                    if (w.closedByUser) break;
                    w.status(`Reconociendo ${cl.addr} («${cl.text}»)…`, progressOf());

                    // Destino actual de la celda (puede haber crecido el rango o añadido un nivel)
                    let target = null;
                    await Excel.run(async (context) => {
                        const sheet = context.workbook.worksheets.getItem(sheetName);
                        sheet.load("id");
                        await context.sync();
                        target = await resolveCell(context, sheet, cl.addr, cl.text, cl.indent);
                    });
                    if (!target || !target.fieldLocated) {
                        dracoLog(TAG, `${cl.addr}: sin dimensión resoluble, se omite`);
                        stats.skipped++; done++;
                        continue;
                    }
                    const f = target.fieldLocated;
                    touchedReports.add(f.reportId);

                    // Miembros del campo (caché de la sesión: una consulta por campo)
                    let items = itemsByField.get(fieldKeyOf(f));
                    try {
                        if (!items) {
                            w.status(`Cargando valores de ${f.dim} / ${f.attr}…`, progressOf());
                            items = await loadDracoPickerItems(f, TAG);
                            itemsByField.set(fieldKeyOf(f), items);
                        }
                    } catch (e) {
                        w.send({ type: "error", message: dracoErrMsg(e) });
                        await w.wait(() => false);
                        break;
                    }

                    const { exact, contains } = findMatches(items, cl.text);
                    const unique = exact.length === 1 ? exact[0] : (exact.length === 0 && contains.length === 1 ? contains[0] : null);
                    if (unique) {
                        await writeMember(sheetName, cl.addr, f, unique.attribute, unique.value, shownFor(unique, w.displayMode));
                        dracoLog(TAG, `${cl.addr}: «${cl.text}» -> ${unique.attribute}=${unique.value} (automático)`);
                        stats.auto++; done++;
                        continue;
                    }

                    // Sin coincidencia única: se selecciona la celda y se pregunta
                    try {
                        await Excel.run(async (context) => {
                            const sheet = context.workbook.worksheets.getItem(sheetName);
                            sheet.activate();
                            sheet.getRange(cl.addr).select();
                            await context.sync();
                        });
                    } catch (e) { /* no se pudo seleccionar: se pregunta igual */ }
                    const fieldKey = fieldKeyOf(f);
                    w.sendItems({ fieldKey, items, title: f.dim + " / " + f.attr });
                    const seq = ++w.round;
                    w.send({ type: "ask", seq, fieldKey, cell: cl.addr, text: cl.text, initialSearch: cl.text, progress: `${done + 1}/${stats.total}` });
                    const ans = await w.wait(m => (m.type === "select" || m.type === "skip") && m.seq === seq);
                    if (ans.type === "stop") break;
                    if (ans.type === "select") {
                        // Lo que se ve, tal cual lo eligió el usuario en el buscador (como el individual)
                        const shown = ans.display && String(ans.display) !== String(ans.value) ? ans.display : undefined;
                        await writeMember(sheetName, cl.addr, f, ans.attribute, ans.value, shown);
                        dracoLog(TAG, `${cl.addr}: «${cl.text}» -> ${ans.attribute}=${ans.value} (elegido)`);
                        stats.chosen++;
                    } else {
                        dracoLog(TAG, `${cl.addr}: omitida`);
                        stats.skipped++;
                    }
                    done++;
                }
            }

            // 4) Rangos con nombre de los informes estáticos tocados (una vez al final)
            if (touchedReports.size) {
                w.status("Terminando…", "");
                await Excel.run(async (context) => {
                    for (const rid of touchedReports) {
                        try { await dracoRebuildStaticNames(context, rid); } catch (e) { /* no estático o sin cambios */ }
                    }
                });
            }
            const rest = stats.total - stats.auto - stats.chosen - stats.skipped;
            dracoLog(TAG, `fin: ${stats.auto} automáticos, ${stats.chosen} elegidos, ${stats.skipped} omitidos` + (rest > 0 ? `, ${rest} sin procesar (detenido)` : ""));
            if (stats.auto + stats.chosen + stats.skipped + rest > 0) {
                const msg = `${stats.auto + stats.chosen} de ${stats.total} reconocidos` +
                    (stats.skipped ? `, ${stats.skipped} omitidos` : "") + (rest > 0 ? `, ${rest} sin procesar` : "") + ".";
                dracoInfoNotice("Reconocimiento de miembros", msg, rest > 0 || stats.skipped ? "warn" : "info");
            }
        } catch (e) {
            dracoLog(TAG, "ERROR: " + dracoErrMsg(e));
            try { w.send({ type: "error", message: dracoErrMsg(e) }); await dracoSleep(2500); } catch (e2) { /* nada */ }
        } finally {
            w.close();
            active = false;
        }
    }

    window.DracoBulkRecognition = { run, isActive, _findMatches: findMatches };
})();
