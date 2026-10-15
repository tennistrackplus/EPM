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
 * Adaptado para Tablas Draco (generado por tools/sync-addin.py desde esta plantilla).
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

    // El anfitrión (diseñador de Tablas Draco o motor sin interfaz) crea el iframe y le cuelga
    // el objeto de estado en iframe.__dracoHost: { state, selection, markDirty(), renderGrid() }.
    // Así cada iframe (el diseñador, el motor de cada tabla) tiene su propia "hoja" y no se pisan.
    function host() {
        const h = window.frameElement && window.frameElement.__dracoHost;
        if (!h) throw new Error("host-bridge: el iframe no tiene anfitrión (frameElement.__dracoHost).");
        return h;
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
    // FÓRMULAS (Tablas Draco: informes Estáticos / Híbridos)
    // ------------------------------------------------------------
    // Cada celda guarda su fórmula en `f` (además del valor mostrado en `v`):
    //   =EPM_VALUE("DIM","ATTR","VALOR","DISPLAY")  -> muestra DISPLAY
    //   =EPM_FORMULA("TEXTO", expr, "fmt", "barras") -> muestra TEXTO
    //   cualquier otra fórmula (=C6-C7, =SUMA(C6:C9)...) se CALCULA (recalc)
    // Operadores + - * / ^ & y comparaciones, paréntesis, referencias A1
    // ($ incluidos) y rangos; funciones en inglés o castellano: SUM/SUMA,
    // MIN, MAX, ABS, AVERAGE/PROMEDIO, ROUND/REDONDEAR, IF/SI,
    // IFERROR/SI.ERROR. Basta para lo que pinta el addin (EPM_FORMULA y los
    // límites de las barras de datos) y para lo que se teclea en el diseñador.
    const FX = (function () {
        const EPM_RE = /^\s*=\s*@?\s*EPM_(VALUE|FORMULA)\s*\(/i;
        const isEpm = (f) => typeof f === "string" && EPM_RE.test(f);

        function splitArgs(f) {
            const i0 = f.indexOf("(");
            const args = [];
            let cur = "", depth = 0, inStr = false;
            for (let i = i0 + 1; i < f.length; i++) {
                const ch = f[i];
                if (inStr) { cur += ch; if (ch === '"') { if (f[i + 1] === '"') { cur += '"'; i++; } else inStr = false; } continue; }
                if (ch === '"') { inStr = true; cur += ch; continue; }
                if (ch === "(" || ch === "{") { depth++; cur += ch; continue; }
                if (ch === ")" || ch === "}") { if (depth === 0) { args.push(cur.trim()); return args; } depth--; cur += ch; continue; }
                if (depth === 0 && (ch === "," || ch === ";")) { args.push(cur.trim()); cur = ""; continue; }
                cur += ch;
            }
            args.push(cur.trim());
            return args;
        }
        const unq = (a) => { const t = String(a || "").trim(); return (t.length >= 2 && t[0] === '"' && t[t.length - 1] === '"') ? t.slice(1, -1).replace(/""/g, '"') : t; };

        function epmDisplay(f) {
            const kind = EPM_RE.exec(f)[1].toUpperCase();
            const a = splitArgs(f);
            return kind === "FORMULA" ? unq(a[0]) : unq(a[3] !== undefined && a[3] !== "" ? a[3] : a[2]);
        }

        const colToIdx = (L) => { let n = 0; for (const ch of L.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };
        const idxToCol = (i) => { let n = i + 1, s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

        // Recorre la fórmula saltando literales de texto y aplica `fn` a cada trozo de código.
        function mapCode(f, fn) {
            let out = "", i = 0;
            while (i < f.length) {
                if (f[i] === '"') {
                    let j = i + 1;
                    while (j < f.length) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; }
                    out += f.slice(i, j + 1); i = j + 1;
                } else {
                    let j = i; while (j < f.length && f[j] !== '"') j++;
                    out += fn(f.slice(i, j)); i = j;
                }
            }
            return out;
        }
        const A1_RE = /(^|[^A-Za-z0-9_.$!\]])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])/g;
        const R1C1_RE = /(^|[^A-Za-z0-9_.!\]])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(])/g;

        // A1 -> R1C1 relativo a la celda (r, c) (0-based)
        function a1ToR1C1(f, r, c) {
            return mapCode(f, (code) => code.replace(A1_RE, (m, pre, dc, col, dr, row) => {
                const ci = colToIdx(col), ri = parseInt(row, 10) - 1;
                if (ci > 16383) return m;
                const R = dr ? "R" + (ri + 1) : (ri === r ? "R" : "R[" + (ri - r) + "]");
                const C = dc ? "C" + (ci + 1) : (ci === c ? "C" : "C[" + (ci - c) + "]");
                return pre + R + C;
            }));
        }
        // R1C1 -> A1 para la celda (r, c) (0-based)
        function r1c1ToA1(f, r, c) {
            return mapCode(f, (code) => code.replace(R1C1_RE, (m, pre, rr, cc) => {
                let ri = r, ci = c, absR = false, absC = false;
                if (rr) { if (rr[0] === "[") ri = r + parseInt(rr.slice(1, -1), 10); else { ri = parseInt(rr, 10) - 1; absR = true; } }
                if (cc) { if (cc[0] === "[") ci = c + parseInt(cc.slice(1, -1), 10); else { ci = parseInt(cc, 10) - 1; absC = true; } }
                if (ri < 0 || ci < 0) return pre + "#REF!";
                return pre + (absC ? "$" : "") + idxToCol(ci) + (absR ? "$" : "") + (ri + 1);
            }));
        }

        // ---- Evaluador
        const ERR = (code) => ({ err: code });
        const isErr = (x) => x && typeof x === "object" && x.err;
        function tokenize(src) {
            const t = []; let i = 0;
            while (i < src.length) {
                const ch = src[i];
                if (/\s/.test(ch)) { i++; continue; }
                if (ch === '"') { let j = i + 1, s = ""; while (j < src.length) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; } t.push({ k: "str", v: s }); i = j + 1; continue; }
                let m = /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?(?![A-Za-z0-9_(.])/.exec(src.slice(i));
                if (m) { t.push({ k: "ref", v: m[0].replace(/\$/g, "") }); i += m[0].length; continue; }
                m = /^\d+(\.\d+)?([eE][+-]?\d+)?%?/.exec(src.slice(i));
                if (m) { let n = parseFloat(m[0]); if (m[0].endsWith("%")) n /= 100; t.push({ k: "num", v: n }); i += m[0].length; continue; }
                m = /^[A-Za-z_][A-Za-z0-9_.]*(?=\s*\()/.exec(src.slice(i));
                if (m) { t.push({ k: "fn", v: m[0].toUpperCase() }); i += m[0].length; continue; }
                m = /^(TRUE|FALSE|VERDADERO|FALSO)\b/i.exec(src.slice(i));
                if (m) { t.push({ k: "num", v: /^(TRUE|VERDADERO)$/i.test(m[0]) ? 1 : 0 }); i += m[0].length; continue; }
                m = /^#(DIV\/0!|REF!|VALUE!|N\/A|NAME\?|¡VALOR!|¡REF!)/.exec(src.slice(i));
                if (m) { t.push({ k: "err", v: m[0] }); i += m[0].length; continue; }
                m = /^(<>|<=|>=|[+\-*/^&=<>(),;])/.exec(src.slice(i));
                if (m) { t.push({ k: "op", v: m[0] === ";" ? "," : m[0] }); i += m[0].length; continue; }
                throw new Error("token");
            }
            return t;
        }
        const ALIAS = { SUMA: "SUM", PROMEDIO: "AVERAGE", REDONDEAR: "ROUND", SI: "IF", "SI.ERROR": "IFERROR" };

        function evaluate(src, getCell) {
            let toks;
            try { toks = tokenize(src); } catch (e) { return ERR("#NAME?"); }
            let p = 0;
            const peek = () => toks[p], next = () => toks[p++];
            const num = (x) => {
                if (isErr(x)) return x;
                if (Array.isArray(x)) x = x[0];
                if (x === "" || x === null || x === undefined) return 0;
                if (typeof x === "number") return x;
                const n = Number(String(x).trim());
                return String(x).trim() !== "" && isFinite(n) ? n : ERR("#VALUE!");
            };
            const flat = (args) => args.flatMap(a => Array.isArray(a) ? a : [a]);
            const nums = (args, strict) => {
                const out = [];
                for (const a of args) {
                    if (Array.isArray(a)) { for (const x of a) { if (isErr(x)) return x; const n = Number(x); if (x !== "" && x !== null && isFinite(n)) out.push(n); } }
                    else { const n = num(a); if (isErr(n)) return n; out.push(n); }
                }
                return out;
            };
            const callFn = (name, args) => {
                name = ALIAS[name] || name;
                switch (name) {
                    case "SUM": { const n = nums(args); return isErr(n) ? n : n.reduce((a, b) => a + b, 0); }
                    case "AVERAGE": { const n = nums(args); return isErr(n) ? n : (n.length ? n.reduce((a, b) => a + b, 0) / n.length : ERR("#DIV/0!")); }
                    case "MIN": { const n = nums(args); return isErr(n) ? n : (n.length ? Math.min(...n) : 0); }
                    case "MAX": { const n = nums(args); return isErr(n) ? n : (n.length ? Math.max(...n) : 0); }
                    case "ABS": { const n = num(args[0]); return isErr(n) ? n : Math.abs(n); }
                    case "ROUND": { const n = num(args[0]), d = num(args[1] || 0); if (isErr(n)) return n; const k = Math.pow(10, d); return Math.round(n * k) / k; }
                    case "IF": { const c = num(args[0]); if (isErr(c)) return c; return c ? (args[1] === undefined ? 1 : args[1]) : (args[2] === undefined ? 0 : args[2]); }
                    case "IFERROR": { const v = Array.isArray(args[0]) ? args[0][0] : args[0]; return isErr(v) ? args[1] : v; }
                    default: return ERR("#NAME?");
                }
            };
            function primary() {
                const tk = next();
                if (!tk) return ERR("#VALUE!");
                if (tk.k === "num") return tk.v;
                if (tk.k === "str") return tk.v;
                if (tk.k === "err") return ERR(tk.v);
                if (tk.k === "ref") {
                    const [a, b] = tk.v.split(":");
                    const pa = /^([A-Za-z]+)(\d+)$/.exec(a);
                    const ra = parseInt(pa[2], 10) - 1, ca = colToIdx(pa[1]);
                    if (!b) return getCell(ra, ca);
                    const pb = /^([A-Za-z]+)(\d+)$/.exec(b);
                    const rb = parseInt(pb[2], 10) - 1, cb = colToIdx(pb[1]);
                    const out = [];
                    for (let r = Math.min(ra, rb); r <= Math.max(ra, rb); r++) for (let c = Math.min(ca, cb); c <= Math.max(ca, cb); c++) out.push(getCell(r, c));
                    return out;
                }
                if (tk.k === "fn") {
                    next(); // "("
                    const args = [];
                    if (!(peek() && peek().v === ")")) {
                        args.push(compare());
                        while (peek() && peek().v === ",") { next(); args.push(compare()); }
                    }
                    if (!peek() || peek().v !== ")") return ERR("#NAME?");
                    next();
                    return callFn(tk.v, args);
                }
                if (tk.v === "(") { const v = compare(); if (peek() && peek().v === ")") next(); return v; }
                if (tk.v === "-") { const v = num(unaryPow()); return isErr(v) ? v : -v; }
                if (tk.v === "+") return num(unaryPow());
                return ERR("#VALUE!");
            }
            function unaryPow() {
                let l = primary();
                while (peek() && peek().v === "^") { next(); const a = num(l), b = num(primary()); l = isErr(a) ? a : isErr(b) ? b : Math.pow(a, b); }
                return l;
            }
            function term() {
                let l = unaryPow();
                while (peek() && (peek().v === "*" || peek().v === "/")) {
                    const op = next().v, a = num(l), b = num(unaryPow());
                    if (isErr(a)) l = a; else if (isErr(b)) l = b;
                    else if (op === "/" && b === 0) l = ERR("#DIV/0!");
                    else l = op === "*" ? a * b : a / b;
                }
                return l;
            }
            function arith() {
                let l = term();
                while (peek() && (peek().v === "+" || peek().v === "-")) {
                    const op = next().v, a = num(l), b = num(term());
                    l = isErr(a) ? a : isErr(b) ? b : (op === "+" ? a + b : a - b);
                }
                return l;
            }
            function concat() {
                let l = arith();
                while (peek() && peek().v === "&") { next(); const r = arith(); l = (isErr(l) ? l : isErr(r) ? r : String(Array.isArray(l) ? l[0] : l) + String(Array.isArray(r) ? r[0] : r)); }
                return l;
            }
            function compare() {
                let l = concat();
                while (peek() && ["=", "<>", "<", ">", "<=", ">="].includes(peek().v)) {
                    const op = next().v, a = num(l), b = num(concat());
                    if (isErr(a)) { l = a; continue; } if (isErr(b)) { l = b; continue; }
                    l = (op === "=" ? a === b : op === "<>" ? a !== b : op === "<" ? a < b : op === ">" ? a > b : op === "<=" ? a <= b : a >= b) ? 1 : 0;
                }
                return l;
            }
            const res = compare();
            return Array.isArray(res) ? res[0] : res;
        }

        return { isEpm, epmDisplay, a1ToR1C1, r1c1ToA1, evaluate, splitArgs, unq, isErr };
    })();
    window.__dracoFormulaEngine = FX;   // lo usa también el diseñador (mismo motor para la vista previa)

    /**
     * Recalcula la hoja principal: fórmulas normales (no EPM) y barras de
     * datos (formato condicional). Se llama al terminar cada Excel.run.
     */
    function recalcMainSheet() {
        const editor = host();
        const cells = editor.state.cells || {};
        const memo = new Map(), busy = new Set();
        const val = (r, c) => {
            const k = r + "_" + c, o = cells[k];
            if (!o) return "";
            if (!o.f || FX.isEpm(o.f)) return o.v === undefined ? "" : o.v;
            if (memo.has(k)) return memo.get(k);
            if (busy.has(k)) return { err: "#REF!" };          // referencia circular
            busy.add(k);
            const v = FX.evaluate(String(o.f).replace(/^\s*=\s*@?/, ""), val);
            busy.delete(k);
            memo.set(k, v);
            return v;
        };
        Object.keys(cells).forEach(k => {
            const o = cells[k];
            if (!o || !o.f || FX.isEpm(o.f)) return;
            const [r, c] = k.split("_").map(Number);
            const v = val(r, c);
            o.v = FX.isErr(v) ? v.err : (typeof v === "number" ? String(Math.round(v * 1e10) / 1e10) : String(v));
        });

        // Barras de datos: posición del eje (a) y extremo de la barra (p) en
        // fracción del ancho de la celda, con los límites de la regla.
        Object.keys(cells).forEach(k => { if (cells[k] && cells[k].db) delete cells[k].db; });
        (editor.state.cf || []).forEach(rule => {
            const db = rule.db || {};
            const bound = (rb, fallback) => {
                if (!rb || !rb.formula) return fallback;
                const v = FX.evaluate(String(rb.formula).replace(/^\s*=\s*/, ""), val);
                return typeof v === "number" && isFinite(v) ? v : fallback;
            };
            const xs = [];
            for (let r = rule.r1; r <= rule.r2; r++) for (let c = rule.c1; c <= rule.c2; c++) {
                const o = cells[r + "_" + c]; const n = o ? Number(o.v) : NaN;
                if (o && o.v !== "" && isFinite(n)) xs.push(n);
            }
            if (!xs.length) return;
            const lo = bound(db.lowerBoundRule, Math.min(0, ...xs)), hi = bound(db.upperBoundRule, Math.max(0, ...xs));
            const span = hi - lo || 1;
            const frac = (x) => Math.max(0, Math.min(1, (x - lo) / span));
            for (let r = rule.r1; r <= rule.r2; r++) for (let c = rule.c1; c <= rule.c2; c++) {
                const o = cells[r + "_" + c]; const n = o ? Number(o.v) : NaN;
                if (!o || o.v === "" || !isFinite(n)) continue;
                const neg = n < 0;
                o.db = {
                    a: Math.round(frac(0) * 1000) / 1000,
                    p: Math.round(frac(n) * 1000) / 1000,
                    col: neg ? ((db.negativeFormat && db.negativeFormat.fillColor) || "#F0436A") : ((db.positiveFormat && db.positiveFormat.fillColor) || "#70D06A"),
                    only: db.showDataBarOnly ? 1 : 0
                };
            }
        });
    }

    function cfStore() {
        const editor = host();
        if (!editor.state.cf) editor.state.cf = [];
        return editor.state.cf;
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
        // Contenido de una celda: fórmula (empieza por "=") o valor.
        function setCellContent(r, c, val) {
            const key = `${r}_${c}`;
            if (!cells[key]) cells[key] = {};
            const o = cells[key];
            const text = val === null || val === undefined ? "" : (typeof val === "string" ? val : String(val));
            if (typeof val === "string" && text.trim().startsWith("=")) {
                o.f = text.trim();
                o.v = FX.isEpm(o.f) ? FX.epmDisplay(o.f) : (o.v === undefined ? "" : o.v); // el resto: recalcMainSheet
            } else {
                delete o.f;
                o.v = text;
            }
        }
        // Excel replica un valor único (escalar o matriz 1x1) a todo el rango.
        function setMatrix(v, fn) {
            if (!Array.isArray(v)) { forEachCell((r, c) => fn(r, c, v)); return; }
            if (v.length === 1 && Array.isArray(v[0]) && v[0].length === 1 && (r2 > r1 || c2 > c1)) { forEachCell((r, c) => fn(r, c, v[0][0])); return; }
            v.forEach((row, ri) => { if (Array.isArray(row)) row.forEach((val, ci) => fn(r1 + ri, c1 + ci, val)); });
        }
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
                // Como en Excel: un texto que empieza por "=" es una fórmula.
                setMatrix(v, (r, c, val) => setCellContent(r, c, val));
            },

            get numberFormat() {
                const out = [];
                for (let r = r1; r <= r2; r++) {
                    const row = [];
                    for (let c = c1; c <= c2; c++) row.push(getCellObj(r, c).nf || "General");
                    out.push(row);
                }
                return out;
            },
            set numberFormat(v) {
                // Excel REPLICA un formato único (cadena o matriz 1x1) a TODAS las celdas del rango: el addin
                // lo usa así (valuesRange.numberFormat = [["#,##0.00"]]). Antes solo se aplicaba a la primera.
                if (typeof v === "string") { forEachCell((r, c) => writeCellObj(r, c, undefined, { nf: v })); return; }
                if (!Array.isArray(v)) return;
                if (v.length === 1 && Array.isArray(v[0]) && v[0].length === 1 && (r2 > r1 || c2 > c1)) {
                    forEachCell((r, c) => writeCellObj(r, c, undefined, { nf: v[0][0] }));
                    return;
                }
                v.forEach((row, ri) => {
                    if (!Array.isArray(row)) return;
                    row.forEach((val, ci) => writeCellObj(r1 + ri, c1 + ci, undefined, { nf: val }));
                });
            },

            // Fórmulas: la celda guarda la fórmula en `f` y lo que muestra en
            // `v` (EPM_VALUE/EPM_FORMULA: su texto; el resto se calcula al
            // terminar el Excel.run, ver recalcMainSheet).
            get formulas() {
                const out = [];
                for (let r = r1; r <= r2; r++) {
                    const row = [];
                    for (let c = c1; c <= c2; c++) { const o = getCellObj(r, c); row.push(o.f !== undefined ? o.f : (o.v ?? "")); }
                    out.push(row);
                }
                return out;
            },
            set formulas(v) { setMatrix(v, (r, c, val) => setCellContent(r, c, val)); },
            get formulasR1C1() {
                const out = [];
                for (let r = r1; r <= r2; r++) {
                    const row = [];
                    for (let c = c1; c <= c2; c++) { const o = getCellObj(r, c); row.push(o.f !== undefined ? FX.a1ToR1C1(o.f, r, c) : (o.v ?? "")); }
                    out.push(row);
                }
                return out;
            },
            set formulasR1C1(v) {
                setMatrix(v, (r, c, val) => setCellContent(r, c, typeof val === "string" && val.trim().startsWith("=") ? FX.r1c1ToA1(val, r, c) : val));
            },
            get formulasLocal() { return range.formulas; },
            set formulasLocal(v) { range.formulas = v; },

            // Formato condicional: solo barras de datos, en la hoja principal.
            conditionalFormats: {
                add(type) {
                    const rule = { r1, c1, r2, c2, type, db: { positiveFormat: {}, negativeFormat: {} } };
                    if (isMainSheet(sheetName)) cfStore().push(rule);
                    return { dataBar: rule.db };
                },
                clearAll() {
                    if (!isMainSheet(sheetName)) return;
                    const list = cfStore();
                    const keep = list.filter(x => x.r2 < r1 || x.r1 > r2 || x.c2 < c1 || x.c1 > c2);
                    list.length = 0; list.push(...keep);
                },
                load() { return this; },
                items: []
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
                    get size() { return collectStyleFlags().fs || 11; },
                    set size(v) { forEachCell((r, c) => writeCellObj(r, c, undefined, { fs: v })); },
                    get name() { return collectStyleFlags().ff || "inherit"; },
                    set name(v) { forEachCell((r, c) => writeCellObj(r, c, undefined, { ff: v })); }
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
            // "Contents" (Sobrescribir formatos desactivado, ver
            // clearDracoNamedRanges en commands.js) solo borra el VALOR y
            // conserva color/fuente/bordes ya pintados; sin distinguir
            // esto de "All", se borraba siempre la celda entera pase lo
            // que pase esa opción.
            clear(applyTo) {
                const onlyContents = applyTo === "Contents";
                forEachCell((r, c) => {
                    if (onlyContents) {
                        const obj = getCellObj(r, c);
                        if (obj && typeof obj === "object") { delete obj.v; delete obj.f; delete obj.db; }
                    } else {
                        clearCellObj(r, c);
                    }
                });
                return range;
            },
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

    // Devuelve una promesa que se cumple cuando TODOS los manejadores (async) han terminado: así el anfitrión
    // sabe cuándo ha acabado de repintarse tras, p. ej., un clic en el +/- de una jerarquía.
    window.__fireExcelEvent = function (type, r, c) {
        const address = `Hoja1!${addressFor(r, c, r, c)}`;
        const args = { address, worksheetId: "Hoja1", source: "Draco" };
        // El clic en el +/- de una jerarquía se decide por la posición (offsetX, en puntos) dentro de la celda: aquí un
        // clic de usuario ya es "sobre la celda", así que se coloca en medio del icono con la propia estimación del addin.
        if (type === "onSingleClicked") {
            const o = backingCells("Hoja1")[`${r}_${c}`] || {};
            args.offsetY = 5;
            try {
                const b = window.estimateFirstLetterBoundsPt(o.ff || "Segoe UI", o.fs || 9, !!o.b, !!o.i, o.ind || 0, "Calibri", 11);
                args.offsetX = (b.start + b.end) / 2;
            } catch (e) { args.offsetX = 8; }
        }
        const pending = [];
        (eventHandlers[type] || []).slice().forEach(handler => {
            try {
                const result = handler(args);
                if (result && typeof result.then === "function") {
                    pending.push(result.catch(err => console.error(`host-bridge: error en handler de ${type}:`, err)));
                }
            } catch (err) {
                console.error(`host-bridge: error en handler de ${type}:`, err);
            }
        });
        return Promise.all(pending);
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
            // El addin actual sigue la selección a nivel de LIBRO (workbook.onSelectionChanged): mismo evento
            onSelectionChanged: makeEventApi("onSelectionChanged"),
            // Estilos del libro: el addin lee el "Normal" (fuente base) para medir el icono +/- de las jerarquías.
            styles: { getItem() { return { load() { return this; }, font: { name: "Calibri", size: 11 } }; } },
            // Modo de cálculo: el addin lo pone en Manual durante un refresco para que Excel no recalcule
            // a cada celda. Aquí no hay motor de cálculo: se acepta y se ignora.
            application: { calculationMode: "Automatic", load() { return this; }, calculate() {} },
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
        CalculationMode: { automatic: "Automatic", automaticExceptTables: "AutomaticExceptTables", manual: "Manual" },
        CalculationType: { recalculate: "Recalculate", full: "Full", fullRebuild: "FullRebuild" },
        ConditionalFormatType: { dataBar: "DataBar" },
        ConditionalDataBarDirection: { leftToRight: "LeftToRight", rightToLeft: "RightToLeft", context: "Context" },
        ConditionalDataBarAxisFormat: { automatic: "Automatic", none: "None", cellMidPoint: "CellMidPoint" },
        ConditionalFormatRuleType: { formula: "Formula", automatic: "Automatic", lowestValue: "LowestValue", highestValue: "HighestValue", number: "Number" },
        run(callback) {
            return Promise.resolve()
                .then(() => callback(fakeContext))
                .then(result => {
                    try { recalcMainSheet(); } catch (e) { console.warn("host-bridge: recálculo", e); }
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
                addHandlerAsync(eventType, handler, callback) {
                    window.addEventListener("message", (ev) => {
                        const m = ev.data;
                        if (!m || m.type !== "parentMessage") return;
                        handler({ message: m.payload });
                    });
                    // Como en Office: se avisa de que el manejador ya está puesto
                    // (calcDialog, entre otros, espera a esto para mandar "ready").
                    if (typeof callback === "function") Promise.resolve().then(() => callback({ status: "succeeded" }));
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
