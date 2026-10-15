/* ==========================================================
 * DracoVariables — variables del libro / de la tabla Draco
 * ----------------------------------------------------------
 * Una variable asocia un NOMBRE a un valor de un campo del modelo:
 *   { name: "PERIODO", dim: "PERIOD", attr: "mes", value: "2026-09", display: "2026-09" }
 *
 * Se usa en las líneas de un eje Estático:
 *   =EPM_VALUE("PERIOD","mes","$PERIODO$","$PERIODO$")
 * El SQL se genera CON EL MARCADOR ($PERIODO$) y se sustituye por el
 * valor justo antes de ejecutarse (resolveSql, desde executeSQL en
 * commands.js). Por eso el SQL guardado de una Tabla Draco vale para
 * cualquier valor de la variable: en tiempo de ejecución basta con
 * cambiar el valor (ver DracoTables.run(..., { variables }) en PLANNING).
 *
 * Se guardan en los ajustes del documento (Office.context.document.settings,
 * clave "epm_variables", como los informes): en el addin, en el libro;
 * en Tablas Draco (host-bridge), en la definición de la tabla.
 * Se puede sustituir temporalmente el valor de las variables con
 * setOverrides({ NOMBRE: valor | {value, display} }), sin tocar lo guardado.
 * ========================================================== */
(function () {
    "use strict";

    const KEY = "epm_variables";
    const NAME_RE = /^[A-Z][A-Z0-9_]{0,39}$/;
    const TOKEN_RE = /\$([A-Za-z][A-Za-z0-9_]{0,39})\$/g;
    let overrides = {};

    function settings() {
        return (typeof Office !== "undefined" && Office.context && Office.context.document && Office.context.document.settings) || null;
    }

    function readAll() {
        const st = settings();
        if (!st) return {};
        try {
            const raw = st.get(KEY);
            const obj = typeof raw === "string" ? JSON.parse(raw || "{}") : (raw || {});
            return (obj && typeof obj === "object") ? obj : {};
        } catch (e) {
            return {};
        }
    }

    function writeAll(obj) {
        const st = settings();
        if (!st) return Promise.resolve();
        st.set(KEY, JSON.stringify(obj));
        return new Promise((resolve) => st.saveAsync(() => resolve()));
    }

    function normName(n) {
        return String(n || "").trim().toUpperCase().replace(/^\$|\$$/g, "");
    }

    const API = {
        KEY,
        NAME_RE,

        /** Lista ordenada por nombre (con los valores sustituidos si hay overrides). */
        list() {
            const all = readAll();
            return Object.keys(all).sort().map(k => API.get(k)).filter(Boolean);
        },

        get(name) {
            const k = normName(name);
            const v = readAll()[k];
            if (!v) return null;
            const out = Object.assign({}, v, { name: k });
            if (Object.prototype.hasOwnProperty.call(overrides, k)) {
                const o = overrides[k];
                out.value = (o && typeof o === "object") ? String(o.value) : String(o);
                out.display = (o && typeof o === "object" && o.display !== undefined) ? String(o.display) : out.value;
                out.overridden = true;
            }
            return out;
        },

        /** Variables de un campo concreto (para el selector del eje Estático). */
        forField(dim, attr) {
            const D = String(dim || "").toUpperCase(), A = String(attr || "").toUpperCase();
            return API.list().filter(v => String(v.dim).toUpperCase() === D && String(v.attr).toUpperCase() === A);
        },

        validateName(name) {
            const k = normName(name);
            if (!k) return "Ponle un nombre a la variable.";
            if (!NAME_RE.test(k)) return "El nombre solo puede tener letras sin acentos, números y _ (y empezar por letra).";
            return "";
        },

        /** Sustituye la lista completa (lo que devuelve el diálogo de variables). */
        async saveAll(list) {
            const obj = {};
            (list || []).forEach(v => {
                const k = normName(v.name);
                if (!k || API.validateName(k)) return;
                obj[k] = { dim: String(v.dim || ""), attr: String(v.attr || ""), value: String(v.value ?? ""), display: String(v.display ?? v.value ?? "") };
            });
            await writeAll(obj);
        },

        async remove(name) {
            const all = readAll();
            delete all[normName(name)];
            await writeAll(all);
        },

        /** Valores temporales (tiempo de ejecución): no se guardan. {} para quitarlos. */
        setOverrides(map) { overrides = {}; Object.keys(map || {}).forEach(k => { overrides[normName(k)] = map[k]; }); },
        getOverrides() { return Object.assign({}, overrides); },

        /** "$PERIODO$" -> nombre, o null si el texto no es (solo) una variable. */
        tokenName(text) {
            const m = /^\s*\$([A-Za-z][A-Za-z0-9_]{0,39})\$\s*$/.exec(String(text || ""));
            return m ? m[1].toUpperCase() : null;
        },

        /** Valor de un texto que sea una variable ($X$ -> su valor); si no, el mismo texto. */
        resolveValue(text) {
            const n = API.tokenName(text);
            if (!n) return text;
            const v = API.get(n);
            return v ? v.value : text;
        },
        resolveDisplay(text) {
            const n = API.tokenName(text);
            if (!n) return text;
            const v = API.get(n);
            return v ? (v.display || v.value) : text;
        },

        /**
         * Sustituye en un SQL los marcadores $NOMBRE$ por el valor de su variable:
         *   '...$X$...'  (dentro de un literal de texto) -> el valor escapado para ese literal
         *   $X$          (fuera de comillas, p. ej. tipo numérico) -> el número, o un literal de texto
         * Lanza un error claro si una variable usada no existe.
         * provider: "bigquery" | "snowflake" | "fabric" (cómo se escapa la comilla).
         */
        resolveSql(sql, provider) {
            if (typeof sql !== "string" || sql.indexOf("$") === -1) return sql;
            const bq = provider === "bigquery";
            const escIn = (s) => bq ? String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'") : String(s).replace(/'/g, "''");
            const missing = new Set();
            const valueOf = (name) => {
                const v = API.get(name);
                if (!v) { missing.add(name.toUpperCase()); return null; }
                return v.value;
            };
            let out = "", i = 0, inStr = false;
            while (i < sql.length) {
                const ch = sql[i];
                if (ch === "'") {
                    // comilla de cierre/apertura (respeta '' y \' dentro de literales)
                    if (inStr && !bq && sql[i + 1] === "'") { out += "''"; i += 2; continue; }
                    inStr = !inStr; out += ch; i++; continue;
                }
                if (inStr && bq && ch === "\\") { out += sql.substr(i, 2); i += 2; continue; }
                if (ch === "$") {
                    const m = /^\$([A-Za-z][A-Za-z0-9_]{0,39})\$/.exec(sql.slice(i));
                    if (m) {
                        const val = valueOf(m[1]);
                        if (val === null) { out += m[0]; i += m[0].length; continue; }
                        if (inStr) out += escIn(val);
                        else out += (/^-?\d+(\.\d+)?$/.test(String(val).trim()) ? String(val).trim() : "'" + escIn(val) + "'");
                        i += m[0].length;
                        continue;
                    }
                }
                out += ch; i++;
            }
            if (missing.size) {
                throw new Error("El informe usa variables que no existen: " + Array.from(missing).map(n => "$" + n + "$").join(", ") + ". Créalas en «Variables».");
            }
            return out;
        },

        /** Nombres de variables usados en un texto (SQL o fórmula). */
        usedIn(text) {
            const out = new Set();
            String(text || "").replace(TOKEN_RE, (m, n) => { out.add(n.toUpperCase()); return m; });
            return Array.from(out);
        }
    };

    window.DracoVariables = API;
})();
