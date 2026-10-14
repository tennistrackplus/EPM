/* ==========================================================
 * DracoCalc — cálculos entre miembros de un atributo
 * ----------------------------------------------------------
 * Un "cálculo" es un miembro más de un atributo (p.ej. "Variación"
 * en VERSION.version) cuyo valor sale de una fórmula sobre otros
 * miembros. Dos modos:
 *
 *   - "member" (por miembro / fijo): [Actual] - [AI Forecast]
 *       Se resuelve en el SQL (commands.js, buildCalcCtes). Los
 *       miembros referenciados se leen aunque NO estén en el filtro
 *       del informe (consulta auxiliar), así que funciona aunque solo
 *       se esté mostrando uno de ellos.
 *
 *   - "position" (por posición / variable): [#1] - [#2]
 *       Se pinta como fórmula de Excel (=C6-D6) que opera con los
 *       miembros que se ven, en el orden en que se ven.
 *
 * Solo + - * / y paréntesis, números y referencias. Se guardan por
 * informe dentro de design.fieldOptions con la clave
 * "__calc__|DIMENSION|ATRIBUTO" (así se persisten, copian y pegan
 * con el resto del diseño sin tocar ReportStore).
 * ========================================================== */
(function () {
    "use strict";

    const KEY_PREFIX = "__calc__";

    function key(dimension, attribute) {
        return `${KEY_PREFIX}|${String(dimension).trim().toUpperCase()}|${String(attribute).trim().toUpperCase()}`;
    }

    function isCalcKey(k) {
        return String(k).indexOf(KEY_PREFIX + "|") === 0;
    }

    function getCalcs(fieldOptions, dimension, attribute) {
        const list = (fieldOptions || {})[key(dimension, attribute)];
        return Array.isArray(list) ? list : [];
    }

    function defaults() {
        return {
            id: "c" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
            name: "Variación",
            mode: "member",
            expr: "",
            position: "end",
            numberFormat: "number",
            decimals: 2,
            display: "number",       // number | bar | barOnly
            colorNeg: "#F0436A",
            colorPos: "#70D06A",
            invertColors: false,
            labelLeft: 42,           // % del ancho reservado al nº negativo
            labelRight: 30           // % del ancho reservado al nº positivo
        };
    }

    /* ---------------- Tokenizador ---------------- */
    const OP_MAP = { "+": "+", "-": "-", "−": "-", "*": "*", "×": "*", "/": "/", "÷": "/" };

    function tokenize(text) {
        const src = String(text || "");
        const tokens = [];
        let i = 0;
        while (i < src.length) {
            const ch = src[i];
            if (/\s/.test(ch)) { i++; continue; }
            if (ch === "[") {
                const end = src.indexOf("]", i + 1);
                if (end === -1) throw new Error("Falta cerrar un corchete ]");
                const inner = src.slice(i + 1, end).trim();
                if (inner === "") throw new Error("Hay una referencia vacía []");
                tokens.push({ t: "ref", v: inner, at: i });
                i = end + 1;
                continue;
            }
            if (/[0-9]/.test(ch) || (ch === "." && /[0-9]/.test(src[i + 1] || ""))) {
                let j = i;
                while (j < src.length && /[0-9.,]/.test(src[j])) j++;
                const raw = src.slice(i, j).replace(",", ".");
                if (!/^\d*\.?\d+$|^\d+\.?\d*$/.test(raw)) throw new Error(`Número no válido: ${src.slice(i, j)}`);
                tokens.push({ t: "num", v: Number(raw), at: i });
                i = j;
                continue;
            }
            if (OP_MAP[ch]) { tokens.push({ t: "op", v: OP_MAP[ch], at: i }); i++; continue; }
            if (ch === "(" || ch === ")") { tokens.push({ t: ch, at: i }); i++; continue; }
            throw new Error(`Carácter no permitido: «${ch}». Usa [miembro], números, + - * / y paréntesis.`);
        }
        return tokens;
    }

    /* ---------------- Parser (descenso recursivo) ----------------
     * expr   := term (('+'|'-') term)*
     * term   := factor (('*'|'/') factor)*
     * factor := num | ref | '(' expr ')' | '-' factor
     */
    function parse(text, mode) {
        const tokens = tokenize(text);
        if (tokens.length === 0) throw new Error("La fórmula está vacía.");
        let p = 0;
        const peek = () => tokens[p];
        const next = () => tokens[p++];

        function factor() {
            const tk = next();
            if (!tk) throw new Error("La fórmula termina antes de tiempo.");
            if (tk.t === "num") return { k: "num", v: tk.v };
            if (tk.t === "ref") {
                if (mode === "position") {
                    const m = /^#?\s*(\d+)$/.exec(tk.v);
                    if (!m || Number(m[1]) < 1) throw new Error(`En modo posición las referencias son [#1], [#2]…: «[${tk.v}]» no vale.`);
                    return { k: "ref", v: Number(m[1]) };
                }
                return { k: "ref", v: tk.v };
            }
            if (tk.t === "(") {
                const e = expr();
                const close = next();
                if (!close || close.t !== ")") throw new Error("Falta cerrar un paréntesis.");
                return e;
            }
            if (tk.t === "op" && tk.v === "-") return { k: "neg", a: factor() };
            throw new Error("Se esperaba un miembro, un número o un paréntesis.");
        }
        function term() {
            let left = factor();
            while (peek() && peek().t === "op" && (peek().v === "*" || peek().v === "/")) {
                const op = next().v;
                left = { k: "bin", op, a: left, b: factor() };
            }
            return left;
        }
        function expr() {
            let left = term();
            while (peek() && peek().t === "op" && (peek().v === "+" || peek().v === "-")) {
                const op = next().v;
                left = { k: "bin", op, a: left, b: term() };
            }
            return left;
        }

        const ast = expr();
        if (p < tokens.length) {
            const tk = tokens[p];
            throw new Error(tk.t === ")" ? "Sobra un paréntesis de cierre." : "Falta un operador entre dos elementos.");
        }
        return ast;
    }

    function refs(ast, out) {
        out = out || [];
        if (!ast) return out;
        if (ast.k === "ref" && out.indexOf(ast.v) === -1) out.push(ast.v);
        if (ast.a) refs(ast.a, out);
        if (ast.b) refs(ast.b, out);
        return out;
    }

    /** Valida y devuelve { ok, ast, refs, error } sin lanzar. */
    function analyze(text, mode) {
        try {
            const ast = parse(text, mode);
            return { ok: true, ast, refs: refs(ast), error: null };
        } catch (e) {
            return { ok: false, ast: null, refs: [], error: e.message || String(e) };
        }
    }

    function fmtNum(v) {
        return String(v);
    }

    /** SQL: refSql(ref) -> expresión SQL; división protegida con NULLIF. */
    function toSql(ast, refSql) {
        switch (ast.k) {
            case "num": return fmtNum(ast.v);
            case "ref": return refSql(ast.v);
            case "neg": return "(-" + toSql(ast.a, refSql) + ")";
            case "bin": {
                const a = toSql(ast.a, refSql);
                const b = toSql(ast.b, refSql);
                if (ast.op === "/") return "(" + a + " / NULLIF(" + b + ", 0))";
                return "(" + a + " " + ast.op + " " + b + ")";
            }
        }
        return "NULL";
    }

    /** Excel (sintaxis inglesa, para range.formulas): refCell(ref) -> "C6" o null. */
    function toExcel(ast, refCell) {
        let missing = false;
        const walk = (n) => {
            switch (n.k) {
                case "num": return fmtNum(n.v);
                case "ref": {
                    const c = refCell(n.v);
                    if (!c) { missing = true; return "0"; }
                    return c;
                }
                case "neg": return "(-" + walk(n.a) + ")";
                case "bin": return "(" + walk(n.a) + n.op + walk(n.b) + ")";
            }
            return "0";
        };
        const body = walk(ast);
        if (missing) return null;
        return '=IFERROR(' + body + ',"")';
    }

    /** Formato de número de Excel para el cálculo (y variante con relleno para barras). */
    function numberFormat(calc) {
        const dec = Math.max(0, Math.min(10, Math.floor(Number(calc.decimals) || 0)));
        const d = dec > 0 ? "." + "0".repeat(dec) : "";
        const base = calc.numberFormat === "percent" ? "0" + d + "%" : "#,##0" + d;
        if (calc.display === "bar") {
            // Negativos pegados a la izquierda y positivos a la derecha
            // (relleno con espacios "* "), para que el número quede junto
            // al extremo hacia el que crece su barra.
            return "* " + base + "\\ ;\\ -" + base + "* ;* " + base + "\\ ";
        }
        return base;
    }

    window.DracoCalc = {
        KEY_PREFIX, key, isCalcKey, getCalcs, defaults,
        tokenize, parse, analyze, refs, toSql, toExcel, numberFormat
    };
})();
