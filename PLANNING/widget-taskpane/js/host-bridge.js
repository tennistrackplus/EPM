/**
 * host-bridge.js — sustituye a "office.js" dentro del taskpane copiado.
 *
 * El add-in real llama a Excel.run/context.sync y a Office.context.* para
 * leer y escribir celdas de la hoja de Excel y para persistir su estado
 * (settings). Aquí, en vez de eso, cada operación se traduce a llamadas
 * sobre window.parent.WidgetTableEditor (el motor de rejilla del widget) o
 * sobre el propio CONFIG_JSON del widget — así taskpane.js, commands.js,
 * reportStore.js y filterModal.js se cargan CASI SIN CAMBIOS respecto al
 * add-in original: misma lógica de arrastrar y soltar, misma generación de
 * SQL, mismos botones y colores (mismo CSS).
 *
 * Dos decisiones de diseño importantes:
 *
 * 1) Hojas: el add-in usa varias hojas ocultas de apoyo (EDIT_REPORT,
 *    CSV_RESULT, MODEL_HIER...) además de la hoja visible con el informe.
 *    Aquí solo existe UNA hoja "visible" de verdad (la rejilla del widget,
 *    "Hoja1"); cualquier otro nombre de hoja se respalda con un almacén en
 *    memoria APARTE (editor.state.hiddenSheets), para que escribir en
 *    EDIT_REPORT!D1 (bookkeeping interno) no sobreescriba una celda D1 que
 *    el usuario esté usando de verdad en su informe.
 *
 * 2) Repintado: Excel repinta la pantalla mientras hay JS ejecutando
 *    (los cambios se ven en cuanto se hacen). Nuestra rejilla necesita que
 *    alguien llame explícitamente a renderGrid(); como el patrón universal
 *    del add-in es Excel.run(async context => {...}), aquí se repinta
 *    automáticamente al terminar CADA Excel.run.
 *
 * Cobertura: acotada a la superficie de la API de Excel realmente usada
 * por commands.js/taskpane.js. Si aparece algo no cubierto, se lanza un
 * error explícito ("host-bridge: <método> no implementado") en vez de
 * fallar en silencio, para poder ampliar el shim sobre casos reales.
 */
(function () {

    // CSS del overlay de diálogo (ver displayDialogAsync más abajo) —
    // inyectado aquí para no tener que tocar taskpane.html.
    const style = document.createElement("style");
    style.textContent = `
        .wte-dialog-overlay {
            position: fixed; inset: 0; z-index: 9999;
            background: rgba(15, 23, 42, 0.45);
            display: flex; align-items: center; justify-content: center;
        }
        .wte-dialog-frame {
            width: min(720px, 92vw); height: min(560px, 88vh);
            border: none; border-radius: 10px;
            background: #fff; box-shadow: 0 20px 60px rgba(0,0,0,0.35);
        }
    `;
    document.head.appendChild(style);

    // Promesa que lkml-bootstrap.js resuelve (window.__wteSignalModelsReady())
    // en cuanto termina de registrar los modelos semánticos — Office.onReady
    // espera a esto antes de dejar arrancar taskpane.js/commands.js (ver
    // más abajo). Si por lo que sea nadie la resuelve nunca (p.ej. esta
    // página no es el taskpane principal, como filterDialog.html/
    // addFilterRange.html), se cae a los 4s a un timeout de seguridad para
    // no dejar la página colgada para siempre.
    window.__wteModelsReadyPromise = new Promise((resolve) => {
        window.__wteSignalModelsReady = resolve;
        setTimeout(resolve, 6000); // margen sobrado: como mucho hay UNA petición de red (el cubo ya activo, si lo hay)
    });

    // Si esta página se cargó como diálogo hijo (ver displayDialogAsync más
    // abajo), su id viaja en la URL — se lee aquí, de forma síncrona, para
    // que esté disponible ANTES de que corra el Office.onReady del hijo
    // (que es una microtarea y se dispara antes que el evento "load" del
    // iframe en el padre).
    try {
        const params = new URLSearchParams(window.location.search);
        if (params.has("wteDialogId")) window.__wteDialogId = parseInt(params.get("wteDialogId"), 10);
    } catch (e) { /* no debería fallar nunca */ }

    function host() {
        if (!window.parent || !window.parent.WidgetTableEditor) {
            throw new Error("host-bridge: no se encuentra WidgetTableEditor en la ventana padre.");
        }
        return window.parent.WidgetTableEditor;
    }

    function notImplemented(name) {
        return function () { throw new Error(`host-bridge: ${name} no implementado todavía.`); };
    }

    // ------------------------------------------------------------
    // Direcciones A1 -> {r1,c1,r2,c2} (0-based, igual que el estado del
    // widget). Soporta "A1", "A1:C5" y, si trae hoja, "Hoja1!A1:C5".
    // ------------------------------------------------------------
    function colLettersToIndex(letters) {
        let n = 0;
        for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
        return n - 1;
    }

    function parseCellRef(ref) {
        const m = ref.match(/^([A-Za-z]+)(\d+)$/);
        if (!m) throw new Error(`host-bridge: referencia de celda no válida "${ref}".`);
        return { c: colLettersToIndex(m[1].toUpperCase()), r: parseInt(m[2], 10) - 1 };
    }

    function parseAddress(address) {
        const clean = String(address).split("!").pop().trim();
        const [a, b] = clean.split(":");
        const start = parseCellRef(a);
        const end = b ? parseCellRef(b) : start;
        return {
            r1: Math.min(start.r, end.r), c1: Math.min(start.c, end.c),
            r2: Math.max(start.r, end.r), c2: Math.max(start.c, end.c)
        };
    }

    function colIndexToLetters(idx) {
        let n = idx + 1, s = "";
        while (n > 0) { const rem = (n - 1) % 26; s = String.fromCharCode(65 + rem) + s; n = Math.floor((n - 1) / 26); }
        return s;
    }

    function addressFor(r1, c1, r2, c2) {
        const a = `${colIndexToLetters(c1)}${r1 + 1}`;
        const b = `${colIndexToLetters(c2)}${r2 + 1}`;
        return a === b ? a : `${a}:${b}`;
    }

    // ------------------------------------------------------------
    // Almacén de celdas/fusiones por hoja. "Hoja1" (o sin nombre) usa la
    // rejilla REAL del widget; cualquier otro nombre (EDIT_REPORT,
    // CSV_RESULT, MODEL_HIER...) usa un almacén aparte en memoria, para no
    // pisar el informe pintado con bookkeeping interno del add-in.
    // ------------------------------------------------------------
    function isMainSheet(name) { return !name || name === "Hoja1"; }

    function backingCells(name) {
        const editor = host();
        if (isMainSheet(name)) return editor.state.cells;
        if (!editor.state.hiddenSheets) editor.state.hiddenSheets = {};
        if (!editor.state.hiddenSheets[name]) editor.state.hiddenSheets[name] = {};
        return editor.state.hiddenSheets[name];
    }

    function backingMerges(name) {
        const editor = host();
        if (isMainSheet(name)) return editor.state.merges;
        if (!editor.state.hiddenSheetMerges) editor.state.hiddenSheetMerges = {};
        if (!editor.state.hiddenSheetMerges[name]) editor.state.hiddenSheetMerges[name] = [];
        return editor.state.hiddenSheetMerges[name];
    }

    // ------------------------------------------------------------
    // Rango de Excel (fake) — respaldado por el almacén de la hoja
    // correspondiente (ver arriba).
    // ------------------------------------------------------------
    function makeRange(r1, c1, r2, c2, sheetName) {
        sheetName = sheetName || "Hoja1";
        const cells = backingCells(sheetName);

        function getCellObj(r, c) { return cells[`${r}_${c}`] || {}; }
        function writeCellObj(r, c, value, extra) {
            const key = `${r}_${c}`;
            if (!cells[key]) cells[key] = {};
            if (value !== undefined) cells[key].v = value;
            if (extra) Object.assign(cells[key], extra);
        }
        function clearCellObj(r, c) { delete cells[`${r}_${c}`]; }
        function collectStyleFlags() { return getCellObj(r1, c1); }
        function forEachCell(fn) { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) fn(r, c); }

        const range = {
            address: `${sheetName}!${addressFor(r1, c1, r2, c2)}`,
            rowIndex: r1,
            columnIndex: c1,
            rowCount: r2 - r1 + 1,
            columnCount: c2 - c1 + 1,

            get worksheet() { return makeWorksheet(sheetName); },

            get values() {
                const out = [];
                for (let r = r1; r <= r2; r++) {
                    const row = [];
                    for (let c = c1; c <= c2; c++) row.push(getCellObj(r, c).v ?? "");
                    out.push(row);
                }
                return out;
            },
            set values(v) {
                if (!Array.isArray(v)) return;
                v.forEach((row, ri) => {
                    if (!Array.isArray(row)) return;
                    row.forEach((val, ci) => {
                        writeCellObj(r1 + ri, c1 + ci, val === null || val === undefined ? "" : String(val));
                    });
                });
            },

            numberFormat: undefined, // no se usa (los valores ya se formatean como texto antes de escribir)

            get formulas() {
                // No implementamos fórmulas EPM_VALUE reales: se devuelve el
                // mismo texto que .values, así el patrón /^=\s*EPM_VALUE\(/i
                // que usa el add-in para detectar "ya está resuelta" da
                // siempre falso y el reconocimiento de miembros puede seguir.
                return range.values;
            },

            format: {
                font: {
                    get bold() { return !!collectStyleFlags().b; },
                    set bold(v) { forEachCell((r, c) => writeCellObj(r, c, undefined, { b: v ? 1 : 0 })); },
                    get italic() { return !!collectStyleFlags().i; },
                    set italic(v) { forEachCell((r, c) => writeCellObj(r, c, undefined, { i: v ? 1 : 0 })); },
                    get underline() { return collectStyleFlags().u ? "Single" : "None"; },
                    set underline(v) { forEachCell((r, c) => writeCellObj(r, c, undefined, { u: (v && v !== "None") ? 1 : 0 })); },
                    get color() { return collectStyleFlags().col || "#000000"; },
                    set color(v) { forEachCell((r, c) => writeCellObj(r, c, undefined, { col: v })); },
                    size: 11,
                    name: "Calibri"
                },
                fill: {
                    get color() { return collectStyleFlags().bg || "#FFFFFF"; },
                    set color(v) { forEachCell((r, c) => writeCellObj(r, c, undefined, { bg: v })); },
                    clear() { forEachCell((r, c) => writeCellObj(r, c, undefined, { bg: "#FFFFFF" })); }
                },
                borders: {
                    getItem(edge) {
                        const map = { EdgeTop: "bt", EdgeRight: "br", EdgeBottom: "bb", EdgeLeft: "bl" };
                        const key = map[edge];
                        // Clave del bug de los bordes "gruesos"/en rejilla: en
                        // Excel de verdad, EdgeTop/Bottom/Left/Right sobre un
                        // RANGO pintan una única línea en el PERÍMETRO exterior
                        // de ese rango, no en cada celda interior. La versión
                        // anterior aplicaba el borde a TODAS las celdas del
                        // rango con forEachCell(), así que un solo "borde
                        // exterior fino" acababa dibujando una rejilla
                        // completa por dentro (cada fila/columna interior
                        // sumaba su propia línea, doblando el grosor visual).
                        // InsideHorizontal/InsideVertical no están soportados
                        // (no los usa el flujo principal): no-op seguro.
                        function boundaryCell() {
                            if (edge === "EdgeBottom") return { r: r2, c: c1 };
                            if (edge === "EdgeRight") return { r: r1, c: c2 };
                            return { r: r1, c: c1 };
                        }
                        function forEdgeCells(fn) {
                            if (edge === "EdgeTop") { for (let c = c1; c <= c2; c++) fn(r1, c); }
                            else if (edge === "EdgeBottom") { for (let c = c1; c <= c2; c++) fn(r2, c); }
                            else if (edge === "EdgeLeft") { for (let r = r1; r <= r2; r++) fn(r, c1); }
                            else if (edge === "EdgeRight") { for (let r = r1; r <= r2; r++) fn(r, c2); }
                        }
                        function current() {
                            if (!key) return null;
                            const bc = boundaryCell();
                            const b = getCellObj(bc.r, bc.c)[key];
                            return (b && typeof b === "object") ? b : null;
                        }
                        function update(patch) {
                            if (!key) return;
                            forEdgeCells((r, c) => {
                                const existing = getCellObj(r, c)[key];
                                const base = (existing && typeof existing === "object") ? existing : { style: "continuous", color: "#1a1f2b", weight: "Thin" };
                                writeCellObj(r, c, undefined, { [key]: Object.assign({}, base, patch) });
                            });
                        }
                        return {
                            get style() { const b = current(); return b && b.style === "continuous" ? "Continuous" : "None"; },
                            set style(v) { update({ style: (v && v !== "None") ? "continuous" : "none" }); },
                            get color() { const b = current(); return (b && b.color) || "#1a1f2b"; },
                            set color(v) { update({ color: v }); },
                            get weight() { const b = current(); return (b && b.weight) || "Thin"; },
                            set weight(v) { update({ weight: v }); }
                        };
                    }
                },
                get indentLevel() { return collectStyleFlags().ind || 0; },
                set indentLevel(v) { forEachCell((r, c) => writeCellObj(r, c, undefined, { ind: v || 0 })); },
                get horizontalAlignment() {
                    const al = collectStyleFlags().al;
                    if (al === "center") return "Center";
                    if (al === "right") return "Right";
                    if (al === "left") return "Left";
                    return "General";
                },
                set horizontalAlignment(v) {
                    const map = { Center: "center", Left: "left", Right: "right", General: "" };
                    const al = Object.prototype.hasOwnProperty.call(map, v) ? map[v] : "";
                    forEachCell((r, c) => writeCellObj(r, c, undefined, { al }));
                },
                columnWidth: undefined,
                autofitColumns() { /* nuestra rejilla ya usa un ancho por columna razonable; no-op */ },
                autofitRows() { /* no-op */ }
            },

            load() { return range; },
            clear() { forEachCell((r, c) => clearCellObj(r, c)); return range; },
            select() { /* no-op: no hay selección de usuario real que mover */ },

            merge(across) {
                const merges = backingMerges(sheetName);
                if (across) {
                    for (let r = r1; r <= r2; r++) merges.push({ r, c: c1, rowSpan: 1, colSpan: c2 - c1 + 1 });
                } else {
                    merges.push({ r: r1, c: c1, rowSpan: r2 - r1 + 1, colSpan: c2 - c1 + 1 });
                }
            },
            unmerge() {
                const merges = backingMerges(sheetName);
                const filtered = merges.filter(m => !(m.r >= r1 && m.r <= r2 && m.c >= c1 && m.c <= c2));
                merges.length = 0;
                merges.push(...filtered);
            },

            getRangeByIndexes(startRow, startCol, rowCount, colCount) {
                return makeRange(r1 + startRow, c1 + startCol, r1 + startRow + rowCount - 1, c1 + startCol + colCount - 1, sheetName);
            },
            getCell(rr, cc) { return makeRange(r1 + rr, c1 + cc, r1 + rr, c1 + cc, sheetName); },
            getEntireColumn: notImplemented("range.getEntireColumn"),
            getEntireRow: notImplemented("range.getEntireRow"),

            rows: { collapsed: false },
            columns: { collapsed: false }
        };
        return range;
    }

    // ------------------------------------------------------------
    // Nombres definidos (rangos con nombre) — usados para asociar filtros
    // ad-hoc a un rango de la hoja (ver addFilterRange). Se guardan dentro
    // del propio estado del widget para que viajen con el resto del diseño.
    // ------------------------------------------------------------
    function namedRangesStore() {
        const editor = host();
        if (!editor.state.namedRanges) editor.state.namedRanges = {};
        return editor.state.namedRanges;
    }

    const namesApi = {
        load() { return this; },
        get items() {
            return Object.keys(namedRangesStore()).map(name => ({ name }));
        },
        getItem(name) {
            const stored = namedRangesStore()[name];
            if (!stored) throw new Error(`host-bridge: nombre definido "${name}" no existe.`);
            return { load() { return this; }, name, getRange: () => makeRange(stored.r1, stored.c1, stored.r2, stored.c2, stored.sheet) };
        },
        getItemOrNullObject(name) {
            const stored = namedRangesStore()[name];
            return {
                load() { return this; },
                name,
                isNullObject: !stored,
                getRange: () => stored ? makeRange(stored.r1, stored.c1, stored.r2, stored.c2, stored.sheet) : null,
                delete: () => { delete namedRangesStore()[name]; }
            };
        },
        add(name, range) {
            namedRangesStore()[name] = {
                r1: range.rowIndex, c1: range.columnIndex,
                r2: range.rowIndex + range.rowCount - 1, c2: range.columnIndex + range.columnCount - 1,
                sheet: (range.address || "").split("!")[0] || "Hoja1"
            };
        }
    };

    // ------------------------------------------------------------
    // Eventos de hoja (onChanged / onSelectionChanged / onSingleClicked).
    // El add-in real los usa para el reconocimiento de miembros (onChanged
    // al confirmar un valor tecleado, onSelectionChanged al entrar en una
    // celda vacía) y para expandir/contraer jerarquías con un clic
    // (onSingleClicked). Se disparan desde WidgetTableEditor (ver
    // fireTaskpaneEvent en widget-table-editor.js) a través de
    // window.__fireExcelEvent, expuesta más abajo.
    // ------------------------------------------------------------
    const eventHandlers = { onChanged: [], onSelectionChanged: [], onSingleClicked: [] };

    function makeEventApi(type) {
        return {
            add(handler) { if (eventHandlers[type]) eventHandlers[type].push(handler); return { remove: () => {} }; },
            remove(handler) { if (eventHandlers[type]) eventHandlers[type] = eventHandlers[type].filter(h => h !== handler); }
        };
    }

    window.__fireExcelEvent = function (type, r, c) {
        const address = `Hoja1!${addressFor(r, c, r, c)}`;
        const args = { address, worksheetId: "Hoja1", source: "Draco" };
        (eventHandlers[type] || []).slice().forEach(handler => {
            try {
                const result = handler(args);
                if (result && typeof result.catch === "function") result.catch(err => console.error(`host-bridge: error en handler de ${type}:`, err));
            } catch (err) {
                console.error(`host-bridge: error en handler de ${type}:`, err);
            }
        });
    };

    // ------------------------------------------------------------
    // Hoja de cálculo (fake). "Hoja1" es la rejilla real y visible;
    // cualquier otro nombre es una hoja interna oculta (ver backingCells).
    // ------------------------------------------------------------
    function makeWorksheet(name) {
        name = name || "Hoja1";
        return {
            name,
            load() { return this; },
            onChanged: makeEventApi("onChanged"),
            onSelectionChanged: makeEventApi("onSelectionChanged"),
            onSingleClicked: makeEventApi("onSingleClicked"),
            comments: { load() {}, items: [] },
            getRange(address) {
                if (!address) {
                    if (isMainSheet(name)) {
                        const editor = host();
                        return makeRange(0, 0, editor.state.rows - 1, editor.state.cols - 1, name);
                    }
                    return makeRange(0, 0, 999, 199, name);
                }
                const a = parseAddress(address);
                return makeRange(a.r1, a.c1, a.r2, a.c2, name);
            },
            getRangeByIndexes(row, col, rowCount, colCount) {
                return makeRange(row, col, row + rowCount - 1, col + colCount - 1, name);
            },
            getCell(row, col) { return makeRange(row, col, row, col, name); },
            getUsedRangeOrNullObject() {
                const cells = backingCells(name);
                const keys = Object.keys(cells);
                if (!keys.length) return { isNullObject: true, load() { return this; } };
                let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
                keys.forEach(k => {
                    const [r, c] = k.split("_").map(Number);
                    if (r < r1) r1 = r; if (c < c1) c1 = c;
                    if (r > r2) r2 = r; if (c > c2) c2 = c;
                });
                return Object.assign({ isNullObject: false }, makeRange(r1, c1, r2, c2, name));
            },
            getUsedRange() {
                const r = this.getUsedRangeOrNullObject();
                if (r.isNullObject) throw new Error("host-bridge: la hoja está vacía (getUsedRange).");
                return r;
            }
        };
    }

    const worksheetsApi = {
        load() { return this; },
        get items() { return [makeWorksheet("Hoja1")]; },
        onAdded: makeEventApi("onChanged"), // nunca se crean hojas nuevas de verdad: no-op seguro
        getActiveWorksheet() { return makeWorksheet("Hoja1"); },
        getItem(name) { return makeWorksheet(name); },
        getItemOrNullObject(name) { return Object.assign({ isNullObject: false }, makeWorksheet(name)); },
        add(name) { return makeWorksheet(name); }
    };

    // ------------------------------------------------------------
    // Excel.run — al terminar cada lote se repinta la rejilla real (si
    // hubo cambios en la hoja principal, se ven inmediatamente).
    // ------------------------------------------------------------
    const fakeContext = {
        workbook: {
            worksheets: worksheetsApi,
            names: namesApi,
            getSelectedRange() {
                const editor = host();
                const sel = editor.selection || { r1: 0, c1: 0, r2: 0, c2: 0 };
                const r1 = Math.min(sel.r1, sel.r2), c1 = Math.min(sel.c1, sel.c2);
                const r2 = Math.max(sel.r1, sel.r2), c2 = Math.max(sel.c1, sel.c2);
                return makeRange(r1, c1, r2, c2, "Hoja1");
            }
        },
        sync() { return Promise.resolve(); }
    };

    window.Excel = {
        HorizontalAlignment: { general: "General", left: "Left", center: "Center", right: "Right" },
        BorderIndex: {
            edgeTop: "EdgeTop", edgeBottom: "EdgeBottom", edgeLeft: "EdgeLeft", edgeRight: "EdgeRight",
            insideHorizontal: "InsideHorizontal", insideVertical: "InsideVertical"
        },
        BorderLineStyle: { continuous: "Continuous", none: "None" },
        BorderWeight: { thin: "Thin", medium: "Medium", thick: "Thick" },
        ClearApplyTo: { all: "All", contents: "Contents", formats: "Formats" },
        run(callback) {
            return Promise.resolve()
                .then(() => callback(fakeContext))
                .then(result => {
                    try { const editor = host(); editor.markDirty(); editor.renderGrid(); } catch (e) { /* el host puede no estar listo aún */ }
                    return result;
                });
        }
    };

    // ------------------------------------------------------------
    // Office.js — settings (persistencia) + diálogos (vía postMessage
    // con un <iframe> propio en lugar de una ventana de Office real)
    // ------------------------------------------------------------
    function settingsStore() {
        const editor = host();
        if (!editor.state.taskpaneSettings) editor.state.taskpaneSettings = {};
        return editor.state.taskpaneSettings;
    }

    let dialogSeq = 0;
    const openDialogs = {};

    window.addEventListener("message", (ev) => {
        const msg = ev.data;
        if (!msg || msg.__wteDialog === undefined) return;
        const entry = openDialogs[msg.__wteDialog];
        if (!entry) return;
        if (msg.type === "childMessage") {
            (entry.handlers.DialogMessageReceived || []).forEach(h => h({ message: msg.payload }));
        }
    });

    window.Office = {
        HostType: { Excel: "Excel" },
        AsyncResultStatus: { Failed: "failed", Succeeded: "succeeded" },
        EventType: {
            DialogMessageReceived: "DialogMessageReceived",
            DialogEventReceived: "DialogEventReceived",
            DialogParentMessageReceived: "DialogParentMessageReceived"
        },
        actions: {
            // El add-in intenta registrar acciones de la cinta real de
            // Excel (Office.actions.associate); aquí no hay cinta, así
            // que se ignoran en silencio en vez de lanzar.
            associate() {}
        },

        onReady(callback) {
            // Espera a que lkml-bootstrap.js termine de registrar los
            // modelos semánticos (descarga real desde GitHub, tarda más
            // que esto) ANTES de dejar correr taskpane.js/commands.js. Sin
            // esto, taskpane.js arranca casi al instante y decide "no hay
            // modelo/informe activo" antes de que el bootstrap termine,
            // perdiendo aparentemente el diseño e informe ya guardados.
            const ready = window.__wteModelsReadyPromise || Promise.resolve();
            return ready.then(() => {
                callback && callback({ host: "Excel" });
                return { host: "Excel" };
            });
        },

        context: {
            document: {
                settings: {
                    get(key) { return Object.prototype.hasOwnProperty.call(settingsStore(), key) ? settingsStore()[key] : null; },
                    set(key, value) { settingsStore()[key] = value; },
                    remove(key) { delete settingsStore()[key]; },
                    saveAsync(callback) {
                        // No hay una persistencia intermedia real: el propio
                        // widget guarda taskpaneSettings junto con el resto
                        // de su estado al pulsar "Guardar widget".
                        Promise.resolve().then(() => callback && callback({ status: "succeeded" }));
                    }
                }
            },
            ui: {
                // Abre el diálogo hijo en un <iframe> superpuesto en vez de
                // una ventana de Office real; el resto del contrato
                // (addEventHandler, messageChild, close) se mantiene igual.
                displayDialogAsync(url, options, callback) {
                    const id = ++dialogSeq;
                    // El id se pasa por la URL (no basta con asignarlo tras
                    // el evento "load" del iframe): Office.onReady del hijo
                    // se dispara en una microtarea, ANTES de que "load"
                    // llegue a ejecutarse, así que el hijo mandaba "ready"
                    // sin saber todavía su propio id y el padre lo descartaba.
                    const sep = url.includes("?") ? "&" : "?";
                    const framedUrl = `${url}${sep}wteDialogId=${id}`;
                    const overlay = document.createElement("div");
                    overlay.className = "wte-dialog-overlay";
                    overlay.innerHTML = `<iframe class="wte-dialog-frame" src="${framedUrl}"></iframe>`;
                    document.body.appendChild(overlay);
                    const frame = overlay.querySelector("iframe");

                    openDialogs[id] = { overlay, frame, handlers: {} };

                    const dialog = {
                        close() {
                            overlay.remove();
                            delete openDialogs[id];
                        },
                        messageChild(msg) {
                            frame.contentWindow.postMessage({ __wteDialog: id, type: "parentMessage", payload: msg }, "*");
                        },
                        addEventHandler(eventType, handler) {
                            if (!openDialogs[id]) return;
                            if (!openDialogs[id].handlers[eventType]) openDialogs[id].handlers[eventType] = [];
                            openDialogs[id].handlers[eventType].push(handler);
                        }
                    };

                    Promise.resolve().then(() => callback && callback({ status: "succeeded", value: dialog }));
                },
                // Lado HIJO (dentro del propio diálogo): manda un mensaje al padre.
                messageParent(msg) {
                    const id = window.__wteDialogId;
                    window.parent.postMessage({ __wteDialog: id, type: "childMessage", payload: msg }, "*");
                },
                addHandlerAsync(eventType, handler) {
                    window.addEventListener("message", (ev) => {
                        const m = ev.data;
                        if (!m || m.type !== "parentMessage") return;
                        handler({ message: m.payload });
                    });
                }
            }
        }
    };

})();

/**
 * Superficie de la API de Excel usada por commands.js (medida con grep
 * antes de escribir este shim), para saber qué falta si algo lanza
 * "host-bridge: ... no implementado":
 *   getRange, worksheets.getItem/getItemOrNullObject/getActiveWorksheet/add,
 *   getRangeByIndexes, names.getItem/getItemOrNullObject/add,
 *   format.font (bold/italic/underline/color), format.indentLevel,
 *   format.fill, getUsedRangeOrNullObject, format.borders,
 *   format.autofitColumns, getCell, format.horizontalAlignment,
 *   format.columnWidth, getSelectedRange, onChanged/onSelectionChanged/
 *   onSingleClicked.
 */
