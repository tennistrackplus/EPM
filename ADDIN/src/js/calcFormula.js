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

/* ==========================================================
 * DracoMeasureFormula — fórmulas de MEDIDAS CALCULADAS del modelo
 * semántico (p. ej. [Ingresos] - [Gastos], DIVIDIR([Margen]; [Ventas])).
 * ----------------------------------------------------------
 * Lista cerrada (nada de SQL libre): números, [Medida], + - * /,
 * paréntesis, comparaciones (= <> < > <= >=) dentro de SI, y las
 * funciones DIVIDIR(a; b), SI(cond; a; b), ABS(x), REDONDEAR(x; n)
 * (también en inglés: DIVIDE, IF, ROUND). Separador de argumentos «;»
 * (también «,» si no va entre dos cifras: 1,5 es un número).
 * toSql(formula, resolve): resolve(nombre) devuelve la expresión SQL ya
 * AGREGADA de cada medida referenciada, así la fórmula se calcula después
 * de agregar (el ratio de las sumas, también en subtotales). La misma
 * traducción vale para BigQuery y Snowflake.
 * ========================================================== */
(function () {
    "use strict";

    const FUNCS = {
        DIVIDIR: "DIV", DIVIDE: "DIV", SI: "IF", IF: "IF", ABS: "ABS", REDONDEAR: "ROUND", ROUND: "ROUND"
    };

    function tokenize(src) {
        const s = String(src || ""), out = [];
        let i = 0;
        while (i < s.length) {
            const c = s[i];
            if (/\s/.test(c)) { i++; continue; }
            if (c === "[") {
                const j = s.indexOf("]", i + 1);
                if (j === -1) throw new Error("Falta cerrar un corchete «]».");
                const name = s.slice(i + 1, j).trim();
                if (!name) throw new Error("Hay unos corchetes vacíos «[]».");
                out.push({ t: "ref", v: name }); i = j + 1; continue;
            }
            const num = /^\d+(?:[.,]\d+)?/.exec(s.slice(i));
            if (num) { out.push({ t: "num", v: num[0].replace(",", ".") }); i += num[0].length; continue; }
            const id = /^[A-Za-zÁÉÍÓÚÑáéíóúñ_]+/.exec(s.slice(i));
            if (id) {
                const f = FUNCS[id[0].toUpperCase()];
                if (!f) throw new Error(`«${id[0]}» no es una función permitida (DIVIDIR, SI, ABS, REDONDEAR). Las medidas van entre corchetes: [${id[0]}].`);
                out.push({ t: "fn", v: f }); i += id[0].length; continue;
            }
            const two = s.slice(i, i + 2);
            if (two === "<=" || two === ">=" || two === "<>") { out.push({ t: "cmp", v: two }); i += 2; continue; }
            if ("<>=".indexOf(c) !== -1) { out.push({ t: "cmp", v: c }); i++; continue; }
            if ("+-*/".indexOf(c) !== -1) { out.push({ t: "op", v: c }); i++; continue; }
            if (c === "(" || c === ")") { out.push({ t: c }); i++; continue; }
            if (c === ";" || c === ",") { out.push({ t: "sep" }); i++; continue; }
            throw new Error(`Carácter no permitido: «${c}».`);
        }
        return out;
    }

    /** Traduce a SQL. resolve(nombre) -> expresión SQL agregada de esa medida. */
    function toSql(formula, resolve) {
        const tk = tokenize(formula);
        if (!tk.length) throw new Error("La fórmula está vacía.");
        let p = 0;
        const peek = () => tk[p], next = () => tk[p++];
        const expect = (t, msg) => { const x = next(); if (!x || x.t !== t) throw new Error(msg); return x; };

        function cmp() {
            const a = add();
            const x = peek();
            if (x && x.t === "cmp") { next(); return "(" + a + " " + x.v + " " + add() + ")"; }
            return a;
        }
        function add() {
            let a = mul();
            while (peek() && peek().t === "op" && (peek().v === "+" || peek().v === "-")) { const o = next().v; a = a + " " + o + " " + mul(); }
            return a;
        }
        function mul() {
            let a = un();
            while (peek() && peek().t === "op" && (peek().v === "*" || peek().v === "/")) {
                const o = next().v; const b = un();
                // «/» también es división segura: sin error si el divisor es 0
                a = o === "/" ? "(" + a + ") / NULLIF(" + b + ", 0)" : a + " * " + b;
            }
            return a;
        }
        function un() {
            const x = peek();
            if (x && x.t === "op" && (x.v === "-" || x.v === "+")) { next(); return (x.v === "-" ? "-" : "") + un(); }
            return prim();
        }
        function args(n0, n1) {
            expect("(", "Falta «(» después de la función.");
            const list = [cmp()];
            while (peek() && peek().t === "sep") { next(); list.push(cmp()); }
            expect(")", "Falta cerrar un paréntesis «)».");
            if (list.length < n0 || list.length > n1) throw new Error("Número de argumentos incorrecto en una función.");
            return list;
        }
        function prim() {
            const x = next();
            if (!x) throw new Error("La fórmula está incompleta.");
            if (x.t === "num") return x.v;
            if (x.t === "ref") return "(" + resolve(x.v) + ")";
            if (x.t === "(") { const e = cmp(); expect(")", "Falta cerrar un paréntesis «)»."); return "(" + e + ")"; }
            if (x.t === "fn") {
                if (x.v === "DIV") { const a = args(2, 2); return "((" + a[0] + ") / NULLIF(" + a[1] + ", 0))"; }
                if (x.v === "IF") { const a = args(3, 3); return "(CASE WHEN " + a[0] + " THEN " + a[1] + " ELSE " + a[2] + " END)"; }
                if (x.v === "ABS") { const a = args(1, 1); return "ABS(" + a[0] + ")"; }
                if (x.v === "ROUND") { const a = args(1, 2); return "ROUND(" + a[0] + (a[1] !== undefined ? ", " + a[1] : "") + ")"; }
            }
            throw new Error("La fórmula tiene algo fuera de lugar cerca de «" + (x.v || x.t) + "».");
        }
        const sql = cmp();
        if (p < tk.length) throw new Error("La fórmula tiene algo de más cerca de «" + (tk[p].v || tk[p].t) + "».");
        return sql;
    }

    /** Medidas referenciadas ([Nombre]) en una fórmula. */
    function refs(formula) {
        try { return tokenize(formula).filter(t => t.t === "ref").map(t => t.v); } catch (e) { return []; }
    }

    window.DracoMeasureFormula = { tokenize, toSql, refs };
})();

/* ==========================================================
 * DracoAttrFormula — ATRIBUTOS CALCULADOS de una dimensión
 * (columna nueva calculada a partir de otras columnas de la MISMA tabla).
 * ----------------------------------------------------------
 * Lista cerrada de funciones (nada de SQL libre), traducidas a SQL que
 * vale para BigQuery y Snowflake (dialect = "bigquery" | "snowflake"):
 *   Texto:  CONCATENAR(a; b; ...) o a & b, IZQUIERDA(t; n), DERECHA(t; n),
 *           EXTRAE(t; inicio; n), MAYUS(t), MINUS(t), ESPACIOS(t),
 *           SUSTITUIR(t; buscar; poner), LARGO(t)
 *   Número: REDONDEAR(x; n), TRUNCAR(x; n), ABS(x)
 *   Fecha:  AÑO(f), MES(f), TRIMESTRE(f), DIA(f), INICIO_MES(f),
 *           INICIO_AÑO(f), TEXTO_FECHA(f; "AAAA-MM")
 *   Lógica: SI(condición; a; b)
 *   Grupos: TRAMOS(x; 1000; 5000; ...)  -> "< 1000", "1000 - 5000", ">= 5000"
 *           AGRUPAR(x; "60*,61*"; "Personal"; "62*"; "Servicios"; "Otros")
 *           (pares patrón -> etiqueta; * = comodín; el último sin pareja es
 *           el valor para el resto; sin él, el resto se queda como está)
 * Columnas entre corchetes: [CECO]. Textos entre comillas: "-".
 * ========================================================== */
(function () {
    "use strict";

    const F = {
        CONCATENAR: "CONCAT", CONCAT: "CONCAT", IZQUIERDA: "LEFT", LEFT: "LEFT", DERECHA: "RIGHT", RIGHT: "RIGHT",
        EXTRAE: "MID", MID: "MID", SUBSTR: "MID", MAYUS: "UPPER", "MAYÚS": "UPPER", UPPER: "UPPER", MINUS: "LOWER",
        "MINÚS": "LOWER", LOWER: "LOWER", ESPACIOS: "TRIM", TRIM: "TRIM", SUSTITUIR: "REPLACE", REPLACE: "REPLACE",
        LARGO: "LEN", LEN: "LEN", REDONDEAR: "ROUND", ROUND: "ROUND", TRUNCAR: "TRUNC", TRUNC: "TRUNC", ABS: "ABS",
        "AÑO": "YEAR", ANIO: "YEAR", YEAR: "YEAR", MES: "MONTH", MONTH: "MONTH", TRIMESTRE: "QUARTER", QUARTER: "QUARTER",
        DIA: "DAY", "DÍA": "DAY", DAY: "DAY", INICIO_MES: "SOM", INICIO_AÑO: "SOY", INICIO_ANIO: "SOY",
        TEXTO_FECHA: "FMTDATE", SI: "IF", IF: "IF", TRAMOS: "BANDS", AGRUPAR: "GROUP"
    };
    // Tipo del resultado según la función exterior (para comparar bien en filtros)
    const RESULT_TYPE = { LEN: "INTEGER", ROUND: "NUMERIC", TRUNC: "NUMERIC", ABS: "NUMERIC", YEAR: "INTEGER", MONTH: "INTEGER", QUARTER: "INTEGER", DAY: "INTEGER", SOM: "DATE", SOY: "DATE" };

    function tokenize(src) {
        const s = String(src || ""), out = [];
        let i = 0;
        while (i < s.length) {
            const c = s[i];
            if (/\s/.test(c)) { i++; continue; }
            if (c === "[") {
                const j = s.indexOf("]", i + 1);
                if (j === -1) throw new Error("Falta cerrar un corchete «]».");
                const name = s.slice(i + 1, j).trim();
                if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) throw new Error(`«[${name}]» no es un nombre de columna válido.`);
                out.push({ t: "col", v: name }); i = j + 1; continue;
            }
            if (c === '"' || c === "'") {
                let j = i + 1, str = "";
                while (j < s.length && s[j] !== c) str += s[j++];
                if (j >= s.length) throw new Error("Falta cerrar unas comillas.");
                out.push({ t: "str", v: str }); i = j + 1; continue;
            }
            const num = /^\d+(?:[.,]\d+)?/.exec(s.slice(i));
            if (num) { out.push({ t: "num", v: num[0].replace(",", ".") }); i += num[0].length; continue; }
            const id = /^[A-Za-zÁÉÍÓÚÑáéíóúñ_]+/.exec(s.slice(i));
            if (id) {
                const f = F[id[0].toUpperCase()];
                if (!f) throw new Error(`«${id[0]}» no es una función permitida. Las columnas van entre corchetes: [${id[0]}].`);
                out.push({ t: "fn", v: f }); i += id[0].length; continue;
            }
            const two = s.slice(i, i + 2);
            if (two === "<=" || two === ">=" || two === "<>") { out.push({ t: "cmp", v: two }); i += 2; continue; }
            if ("<>=".indexOf(c) !== -1) { out.push({ t: "cmp", v: c }); i++; continue; }
            if ("+-*/".indexOf(c) !== -1) { out.push({ t: "op", v: c }); i++; continue; }
            if (c === "&") { out.push({ t: "amp" }); i++; continue; }
            if (c === "(" || c === ")") { out.push({ t: c }); i++; continue; }
            if (c === ";" || c === ",") { out.push({ t: "sep" }); i++; continue; }
            throw new Error(`Carácter no permitido: «${c}».`);
        }
        return out;
    }

    const q = (v) => "'" + String(v).replace(/\\/g, "\\\\").replace(/'/g, "''") + "'";
    const str = (x) => "CAST(" + x + " AS STRING)";

    /** col(nombre) -> referencia SQL de la columna (con o sin alias de tabla). */
    function toSql(formula, col, dialect) {
        const sf = String(dialect || "bigquery").toLowerCase() === "snowflake";
        const tk = tokenize(formula);
        if (!tk.length) throw new Error("La fórmula está vacía.");
        let p = 0;
        const peek = () => tk[p], next = () => tk[p++];
        const expect = (t, msg) => { const x = next(); if (!x || x.t !== t) throw new Error(msg); return x; };

        function concat() {
            let a = cmp();
            const parts = [a];
            while (peek() && peek().t === "amp") { next(); parts.push(cmp()); }
            return parts.length > 1 ? "CONCAT(" + parts.map(str).join(", ") + ")" : a;
        }
        function cmp() {
            const a = add(); const x = peek();
            if (x && x.t === "cmp") { next(); return "(" + a + " " + x.v + " " + add() + ")"; }
            return a;
        }
        function add() {
            let a = mul();
            while (peek() && peek().t === "op" && (peek().v === "+" || peek().v === "-")) { const o = next().v; a = a + " " + o + " " + mul(); }
            return a;
        }
        function mul() {
            let a = un();
            while (peek() && peek().t === "op" && (peek().v === "*" || peek().v === "/")) {
                const o = next().v; const b = un();
                a = o === "/" ? "(" + a + ") / NULLIF(" + b + ", 0)" : a + " * " + b;
            }
            return a;
        }
        function un() {
            const x = peek();
            if (x && x.t === "op" && (x.v === "-" || x.v === "+")) { next(); return (x.v === "-" ? "-" : "") + un(); }
            return prim();
        }
        function args() {
            expect("(", "Falta «(» después de la función.");
            const list = [];
            if (peek() && peek().t === ")") { next(); return list; }
            list.push(concat());
            while (peek() && peek().t === "sep") { next(); list.push(concat()); }
            expect(")", "Falta cerrar un paréntesis «)».");
            return list;
        }
        const n = (a, min, max, name) => { if (a.length < min || a.length > max) throw new Error(`Número de argumentos incorrecto en ${name}.`); };
        function prim() {
            const x = next();
            if (!x) throw new Error("La fórmula está incompleta.");
            if (x.t === "num") return x.v;
            if (x.t === "str") return q(x.v);
            if (x.t === "col") return col(x.v);
            if (x.t === "(") { const e = concat(); expect(")", "Falta cerrar un paréntesis «)»."); return "(" + e + ")"; }
            if (x.t !== "fn") throw new Error("La fórmula tiene algo fuera de lugar cerca de «" + (x.v || x.t) + "».");
            const a = args();
            switch (x.v) {
                case "CONCAT": n(a, 1, 30, "CONCATENAR"); return "CONCAT(" + a.map(str).join(", ") + ")";
                case "LEFT": n(a, 2, 2, "IZQUIERDA"); return "SUBSTR(" + str(a[0]) + ", 1, " + a[1] + ")";
                case "RIGHT": n(a, 2, 2, "DERECHA"); return "RIGHT(" + str(a[0]) + ", " + a[1] + ")";
                case "MID": n(a, 3, 3, "EXTRAE"); return "SUBSTR(" + str(a[0]) + ", " + a[1] + ", " + a[2] + ")";
                case "UPPER": n(a, 1, 1, "MAYUS"); return "UPPER(" + str(a[0]) + ")";
                case "LOWER": n(a, 1, 1, "MINUS"); return "LOWER(" + str(a[0]) + ")";
                case "TRIM": n(a, 1, 1, "ESPACIOS"); return "TRIM(" + str(a[0]) + ")";
                case "REPLACE": n(a, 3, 3, "SUSTITUIR"); return "REPLACE(" + str(a[0]) + ", " + a[1] + ", " + a[2] + ")";
                case "LEN": n(a, 1, 1, "LARGO"); return "LENGTH(" + str(a[0]) + ")";
                case "ROUND": n(a, 1, 2, "REDONDEAR"); return "ROUND(" + a[0] + (a[1] !== undefined ? ", " + a[1] : "") + ")";
                case "TRUNC": n(a, 1, 2, "TRUNCAR"); return "TRUNC(" + a[0] + (a[1] !== undefined ? ", " + a[1] : "") + ")";
                case "ABS": n(a, 1, 1, "ABS"); return "ABS(" + a[0] + ")";
                case "YEAR": n(a, 1, 1, "AÑO"); return "EXTRACT(YEAR FROM " + a[0] + ")";
                case "MONTH": n(a, 1, 1, "MES"); return "EXTRACT(MONTH FROM " + a[0] + ")";
                case "QUARTER": n(a, 1, 1, "TRIMESTRE"); return "EXTRACT(QUARTER FROM " + a[0] + ")";
                case "DAY": n(a, 1, 1, "DIA"); return "EXTRACT(DAY FROM " + a[0] + ")";
                case "SOM": n(a, 1, 1, "INICIO_MES"); return sf ? "DATE_TRUNC('MONTH', " + a[0] + ")" : "DATE_TRUNC(DATE(" + a[0] + "), MONTH)";
                case "SOY": n(a, 1, 1, "INICIO_AÑO"); return sf ? "DATE_TRUNC('YEAR', " + a[0] + ")" : "DATE_TRUNC(DATE(" + a[0] + "), YEAR)";
                case "FMTDATE": {
                    n(a, 2, 2, "TEXTO_FECHA");
                    const m = /^'(.*)'$/.exec(a[1]);
                    if (!m) throw new Error("TEXTO_FECHA: el formato va entre comillas, p. ej. \"AAAA-MM\".");
                    const fmt = m[1];
                    if (!/^[AMD\-\/ .]+$/i.test(fmt)) throw new Error("TEXTO_FECHA: usa solo AAAA, MM, DD y separadores (- / . espacio).");
                    if (sf) return "TO_CHAR(" + a[0] + ", " + q(fmt.replace(/AAAA/gi, "YYYY").replace(/DD/gi, "DD").replace(/MM/gi, "MM")) + ")";
                    return "FORMAT_DATE(" + q(fmt.replace(/AAAA/gi, "%Y").replace(/MM/gi, "%m").replace(/DD/gi, "%d")) + ", DATE(" + a[0] + "))";
                }
                case "IF": n(a, 3, 3, "SI"); return "(CASE WHEN " + a[0] + " THEN " + a[1] + " ELSE " + a[2] + " END)";
                case "BANDS": {
                    if (a.length < 2) throw new Error("TRAMOS necesita el valor y al menos un corte.");
                    const x0 = a[0], cuts = a.slice(1);
                    if (!cuts.every(c => /^-?\d+(\.\d+)?$/.test(c))) throw new Error("TRAMOS: los cortes tienen que ser números.");
                    const nums = cuts.map(Number);
                    for (let i = 1; i < nums.length; i++) if (!(nums[i] > nums[i - 1])) throw new Error("TRAMOS: los cortes van de menor a mayor.");
                    let sql = "(CASE WHEN " + x0 + " IS NULL THEN NULL";
                    sql += " WHEN " + x0 + " < " + cuts[0] + " THEN " + q("< " + cuts[0]);
                    for (let i = 1; i < cuts.length; i++) sql += " WHEN " + x0 + " < " + cuts[i] + " THEN " + q(cuts[i - 1] + " - " + cuts[i]);
                    return sql + " ELSE " + q(">= " + cuts[cuts.length - 1]) + " END)";
                }
                case "GROUP": {
                    if (a.length < 3) throw new Error("AGRUPAR necesita el valor y al menos un par patrón; etiqueta.");
                    const x0 = str(a[0]), rest = a.slice(1);
                    const pairs = [], dflt = rest.length % 2 === 1 ? rest[rest.length - 1] : null;
                    for (let i = 0; i + 1 < rest.length; i += 2) pairs.push([rest[i], rest[i + 1]]);
                    let sql = "(CASE";
                    pairs.forEach(([pat, label]) => {
                        const m = /^'(.*)'$/.exec(pat);
                        if (!m) throw new Error("AGRUPAR: los patrones van entre comillas, p. ej. \"60*,61*\".");
                        const conds = m[1].split(",").map(s => s.trim()).filter(Boolean).map(pp => pp.indexOf("*") !== -1
                            ? x0 + " LIKE " + q(pp.replace(/%/g, "").replace(/_/g, "\\_").replace(/\*/g, "%"))
                            : x0 + " = " + q(pp));
                        if (!conds.length) throw new Error("AGRUPAR: hay un patrón vacío.");
                        sql += " WHEN " + conds.join(" OR ") + " THEN " + label;
                    });
                    return sql + " ELSE " + (dflt !== null ? dflt : x0) + " END)";
                }
            }
            throw new Error("Función no soportada.");
        }
        const sql = concat();
        if (p < tk.length) throw new Error("La fórmula tiene algo de más cerca de «" + (tk[p].v || tk[p].t) + "».");
        return sql;
    }

    function columns(formula) {
        try { return tokenize(formula).filter(t => t.t === "col").map(t => t.v); } catch (e) { return []; }
    }

    /** Tipo del resultado (para comparar con comillas o sin ellas en los filtros). */
    function resultType(formula) {
        try {
            const tk = tokenize(formula);
            if (tk.length && tk[0].t === "fn" && tk[tk.length - 1].t === ")") {
                // solo si la función exterior envuelve toda la fórmula
                let depth = 0, wraps = true;
                for (let i = 1; i < tk.length; i++) {
                    if (tk[i].t === "(") depth++;
                    if (tk[i].t === ")") { depth--; if (depth === 0 && i < tk.length - 1) { wraps = false; break; } }
                }
                if (wraps && RESULT_TYPE[tk[0].v]) return RESULT_TYPE[tk[0].v];
            }
        } catch (e) { /* nada */ }
        return "STRING";
    }

    window.DracoAttrFormula = { tokenize, toSql, columns, resultType };
})();
