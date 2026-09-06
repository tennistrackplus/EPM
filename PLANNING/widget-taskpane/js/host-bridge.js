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
 * Cobertura: se ha acotado a la superficie de la API de Excel realmente
 * usada por commands.js (ver comentario al final del fichero con el grep
 * que se usó para medirla). Si aparece algo no cubierto, se lanza un error
 * explícito ("host-bridge: <método> no implementado") en vez de fallar en
 * silencio, para poder ampliar el shim sobre casos reales.
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
    // widget). Soporta "A1", "A1:C5" y, por si acaso, "Hoja1!A1:C5".
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
    // Rango de Excel (fake) — respaldado por el estado del widget.
    // ------------------------------------------------------------
    function makeRange(r1, c1, r2, c2) {
        const editor = host();

        function collectStyleFlags() {
            // Devuelve el estilo "predominante" de la primera celda del
            // rango, que es lo que casi siempre se consulta tras un load().
            return editor.getCell(r1, c1) || {};
        }

        function forEachCell(fn) {
            for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) fn(r, c);
        }

        const range = {
            address: addressFor(r1, c1, r2, c2),
            rowIndex: r1,
            columnIndex: c1,
            rowCount: r2 - r1 + 1,
            columnCount: c2 - c1 + 1,

            get worksheet() { return makeWorksheet("Hoja1"); },

            get values() {
                const out = [];
                for (let r = r1; r <= r2; r++) {
                    const row = [];
                    for (let c = c1; c <= c2; c++) row.push(editor.getCell(r, c).v ?? "");
                    out.push(row);
                }
                return out;
            },
            set values(v) {
                if (!Array.isArray(v)) return;
                v.forEach((row, ri) => {
                    if (!Array.isArray(row)) return;
                    row.forEach((val, ci) => {
                        editor.writeCell(r1 + ri, c1 + ci, val === null || val === undefined ? "" : String(val));
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
                    set bold(v) { forEachCell((r, c) => editor.writeCell(r, c, editor.getCell(r, c).v ?? "", { b: v ? 1 : 0 })); },
                    get italic() { return !!collectStyleFlags().i; },
                    set italic(v) { forEachCell((r, c) => editor.writeCell(r, c, editor.getCell(r, c).v ?? "", { i: v ? 1 : 0 })); },
                    get underline() { return collectStyleFlags().u ? "Single" : "None"; },
                    set underline(v) { forEachCell((r, c) => editor.writeCell(r, c, editor.getCell(r, c).v ?? "", { u: (v && v !== "None") ? 1 : 0 })); },
                    get color() { return collectStyleFlags().col || "#000000"; },
                    set color(v) { forEachCell((r, c) => editor.writeCell(r, c, editor.getCell(r, c).v ?? "", { col: v })); },
                    size: 11,
                    name: "Calibri"
                },
                fill: {
                    get color() { return collectStyleFlags().bg || "#FFFFFF"; },
                    set color(v) { forEachCell((r, c) => editor.writeCell(r, c, editor.getCell(r, c).v ?? "", { bg: v })); },
                    clear() { forEachCell((r, c) => editor.writeCell(r, c, editor.getCell(r, c).v ?? "", { bg: "#FFFFFF" })); }
                },
                borders: {
                    getItem(edge) {
                        const map = { EdgeTop: "bt", EdgeRight: "br", EdgeBottom: "bb", EdgeLeft: "bl" };
                        const key = map[edge];
                        return {
                            get style() { return key && collectStyleFlags()[key] ? "Continuous" : "None"; },
                            set style(v) { if (key) forEachCell((r, c) => editor.writeCell(r, c, editor.getCell(r, c).v ?? "", { [key]: (v && v !== "None") ? 1 : 0 })); },
                            color: "#000000",
                            weight: "Thin"
                        };
                    }
                },
                indentLevel: 0,
                horizontalAlignment: "General",
                columnWidth: undefined,
                autofitColumns() { /* nuestra rejilla ya usa un ancho por columna razonable; no-op */ },
                autofitRows() { /* no-op */ }
            },

            load() { return range; },
            clear() { forEachCell((r, c) => editor.clearRegion(r, c, r, c)); return range; },
            select() { /* no-op: no hay selección de usuario real que mover */ },

            merge(across) {
                if (across) {
                    for (let r = r1; r <= r2; r++) editor.state.merges.push({ r, c: c1, rowSpan: 1, colSpan: c2 - c1 + 1 });
                } else {
                    editor.state.merges.push({ r: r1, c: c1, rowSpan: r2 - r1 + 1, colSpan: c2 - c1 + 1 });
                }
                editor.markDirty();
            },
            unmerge() {
                editor.state.merges = editor.state.merges.filter(m =>
                    !(m.r >= r1 && m.r <= r2 && m.c >= c1 && m.c <= c2));
                editor.markDirty();
            },

            getRangeByIndexes(startRow, startCol, rowCount, colCount) {
                return makeRange(r1 + startRow, c1 + startCol, r1 + startRow + rowCount - 1, c1 + startCol + colCount - 1);
            },
            getCell(rr, cc) { return makeRange(r1 + rr, c1 + cc, r1 + rr, c1 + cc); },
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
            return { getRange: () => makeRange(stored.r1, stored.c1, stored.r2, stored.c2) };
        },
        getItemOrNullObject(name) {
            const stored = namedRangesStore()[name];
            return {
                isNullObject: !stored,
                getRange: () => stored ? makeRange(stored.r1, stored.c1, stored.r2, stored.c2) : null,
                delete: () => { delete namedRangesStore()[name]; }
            };
        },
        add(name, range) {
            namedRangesStore()[name] = { r1: range.rowIndex, c1: range.columnIndex, r2: range.rowIndex + range.rowCount - 1, c2: range.columnIndex + range.columnCount - 1 };
        }
    };

    // ------------------------------------------------------------
    // Hoja de cálculo (fake) — solo hay UNA "hoja", la rejilla del widget.
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // Eventos de hoja (onChanged / onSelectionChanged / onSingleClicked).
    // El add-in real los usa para el reconocimiento de miembros (onChanged
    // al confirmar un valor tecleado, onSelectionChanged al entrar en una
    // celda vacía) y para expandir/contraer jerarquías con un clic
    // (onSingleClicked). Aquí se disparan desde WidgetTableEditor (ver
    // fireTaskpaneEvent en widget-table-editor.js) a través de
    // window.__fireExcelEvent, expuesta más abajo.
    // ------------------------------------------------------------
    const eventHandlers = { onChanged: [], onSelectionChanged: [], onSingleClicked: [] };

    function makeEventApi(type) {
        return {
            add(handler) { eventHandlers[type].push(handler); return { remove: () => {} }; },
            remove(handler) { eventHandlers[type] = eventHandlers[type].filter(h => h !== handler); }
        };
    }

    window.__fireExcelEvent = function (type, r, c) {
        const address = addressFor(r, c, r, c);
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

    function makeWorksheet(name) {
        const editor = host();
        return {
            name: name || "Hoja1",
            load() { return this; },
            onChanged: makeEventApi("onChanged"),
            onSelectionChanged: makeEventApi("onSelectionChanged"),
            onSingleClicked: makeEventApi("onSingleClicked"),
            comments: { load() {}, items: [] },
            getRange(address) {
                if (!address) return makeRange(0, 0, editor.state.rows - 1, editor.state.cols - 1);
                const a = parseAddress(address);
                return makeRange(a.r1, a.c1, a.r2, a.c2);
            },
            getRangeByIndexes(row, col, rowCount, colCount) {
                return makeRange(row, col, row + rowCount - 1, col + colCount - 1);
            },
            getCell(row, col) { return makeRange(row, col, row, col); },
            getUsedRangeOrNullObject() {
                const keys = Object.keys(editor.state.cells);
                if (!keys.length) return { isNullObject: true };
                let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
                keys.forEach(k => {
                    const [r, c] = k.split("_").map(Number);
                    if (r < r1) r1 = r; if (c < c1) c1 = c;
                    if (r > r2) r2 = r; if (c > c2) c2 = c;
                });
                return Object.assign({ isNullObject: false }, makeRange(r1, c1, r2, c2));
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
        getActiveWorksheet() { return makeWorksheet("Hoja1"); },
        getItem(name) { return makeWorksheet(name); },
        getItemOrNullObject(name) { return Object.assign({ isNullObject: false }, makeWorksheet(name)); },
        add(name) { return makeWorksheet(name); }
    };

    // ------------------------------------------------------------
    // Excel.run — sin lote/lote real: cada operación ya se aplica al
    // vuelo sobre el estado del widget, así que sync() no tiene nada
    // pendiente que confirmar.
    // ------------------------------------------------------------
    const fakeContext = {
        workbook: { worksheets: worksheetsApi, names: namesApi },
        sync() { return Promise.resolve(); }
    };

    window.Excel = {
        run(callback) { return Promise.resolve().then(() => callback(fakeContext)); }
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

        onReady(callback) {
            Promise.resolve().then(() => callback && callback({ host: "Excel" }));
            return Promise.resolve({ host: "Excel" });
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
                    const overlay = document.createElement("div");
                    overlay.className = "wte-dialog-overlay";
                    overlay.innerHTML = `<iframe class="wte-dialog-frame" src="${url}"></iframe>`;
                    document.body.appendChild(overlay);
                    const frame = overlay.querySelector("iframe");

                    openDialogs[id] = { overlay, frame, handlers: {} };

                    frame.addEventListener("load", () => {
                        try { frame.contentWindow.__wteDialogId = id; } catch (e) { /* cross-origin, no debería pasar (mismo origen) */ }
                    });

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
 * Superficie de la API de Excel usada por commands.js (medida con
 * grep antes de escribir este shim), para saber qué falta si algo
 * lanza "host-bridge: ... no implementado":
 *   getRange, worksheets.getItem/getItemOrNullObject/getActiveWorksheet/add,
 *   getRangeByIndexes, names.getItem/getItemOrNullObject/add,
 *   format.font (bold/italic/underline/color), format.indentLevel,
 *   format.fill, getUsedRangeOrNullObject, format.borders,
 *   format.autofitColumns, getCell, format.horizontalAlignment,
 *   format.columnWidth.
 */
