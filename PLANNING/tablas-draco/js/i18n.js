/**
 * ============================================================
 * DRACO — TRADUCCIÓN DE LA INTERFAZ (i18n)
 * ============================================================
 * Idioma original: español. La CLAVE de cada traducción es el propio
 * texto en español, así que:
 *   - Lo que no tenga traducción se sigue viendo en español (nunca se
 *     rompe nada por una clave que falte).
 *   - Las páginas no hay que reescribirlas: al cargar, y cada vez que el
 *     código pinta algo nuevo (innerHTML, toasts, diálogos), se traducen
 *     los textos y los atributos placeholder/title/aria-label/alt cuyo
 *     texto en español esté en el diccionario.
 *   - Los textos con partes variables se registran como plantillas:
 *     "Conexión guardada para {nombre}." traduce cualquier nombre.
 *   - En el código se puede usar también DracoI18n.t("texto", {param}).
 *
 * NO se traduce nunca:
 *   - Lo que está dentro de un elemento con el atributo data-no-i18n
 *     (listas de campos, miembros, modelos… es decir, DATOS).
 *   - <script>, <style>, <textarea>, <input> (su valor), <code>, <pre>.
 *   - Nada de lo que se escribe en las celdas de Excel ni en BigQuery.
 *
 * Idioma: preferencia guardada (DracoI18n.setLanguage) o, si no hay, el
 * de la interfaz de Office / del navegador. Español si empieza por "es";
 * inglés en cualquier otro caso.
 *
 * Diccionarios: js/i18n/<idioma>.js, que llaman a DracoI18n.register().
 * Este fichero es idéntico en el addin y en Planning.
 */
(function () {
    "use strict";

    const SOURCE_LANG = "es";
    const SUPPORTED = ["es", "en"];
    const PREF_KEY = "draco_lang";
    const SKIP_TAGS = { SCRIPT: 1, STYLE: 1, TEXTAREA: 1, CODE: 1, PRE: 1, NOSCRIPT: 1, TEMPLATE: 1 };
    const ATTRS = ["placeholder", "title", "aria-label", "alt"];

    const dicts = {};        // idioma -> { textoES: traducción }
    const patterns = {};     // idioma -> [{ re, names, tpl }]
    let lang = null;
    let observer = null;
    let applying = false;
    let skipSelector = "";   // selectores CSS de zonas de DATOS (no se traducen)

    function norm(s) {
        return String(s).replace(/\s+/g, " ").trim();
    }

    function escapeRe(s) {
        return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    function fill(tpl, params) {
        if (!params) return tpl;
        return tpl.replace(/\{(\w+)\}/g, (m, k) => (params[k] !== undefined && params[k] !== null ? String(params[k]) : m));
    }

    function detect() {
        let pref = null;
        try { pref = localStorage.getItem(PREF_KEY); } catch (e) { /* sin acceso */ }
        if (pref && SUPPORTED.includes(pref)) return pref;
        let ui = "";
        try {
            if (typeof Office !== "undefined" && Office.context && Office.context.displayLanguage) ui = Office.context.displayLanguage;
        } catch (e) { /* fuera de Office */ }
        if (!ui) ui = (navigator.languages && navigator.languages[0]) || navigator.language || SOURCE_LANG;
        return String(ui).toLowerCase().startsWith("es") ? "es" : "en";
    }

    /** Traduce un texto completo (exacto o por plantilla). null si no hay traducción. */
    function lookup(text) {
        if (lang === SOURCE_LANG) return null;
        const d = dicts[lang];
        if (!d) return null;
        const key = norm(text);
        if (!key) return null;
        if (Object.prototype.hasOwnProperty.call(d, key)) return d[key];
        const list = patterns[lang] || [];
        for (const p of list) {
            const m = p.re.exec(key);
            if (m) {
                const params = {};
                p.names.forEach((n, i) => { params[n] = m[i + 1]; });
                return fill(p.tpl, params);
            }
        }
        return null;
    }

    function skipped(el) {
        for (let n = el; n && n.nodeType === 1; n = n.parentElement) {
            if (SKIP_TAGS[n.tagName]) return true;
            if (n.hasAttribute && n.hasAttribute("data-no-i18n")) return true;
            if (skipSelector && n.matches && n.matches(skipSelector)) return true;
        }
        return false;
    }

    function translateTextNode(node) {
        const raw = node.nodeValue;
        if (!raw || !/[A-Za-zÁÉÍÓÚáéíóúÑñ]/.test(raw)) return;
        const tr = lookup(raw);
        if (tr === null || tr === norm(raw)) return;
        const lead = raw.match(/^\s*/)[0];
        const trail = raw.match(/\s*$/)[0];
        node.nodeValue = lead + tr + trail;
    }

    function translateAttrs(el) {
        for (const a of ATTRS) {
            if (!el.hasAttribute(a)) continue;
            const tr = lookup(el.getAttribute(a));
            if (tr !== null) el.setAttribute(a, tr);
        }
        // Botones <input type="button|submit"> muestran su value
        if (el.tagName === "INPUT" && /^(button|submit|reset)$/i.test(el.type) && el.value) {
            const tr = lookup(el.value);
            if (tr !== null) el.value = tr;
        }
    }

    function translateTree(root) {
        if (lang === SOURCE_LANG || !root) return;
        applying = true;
        try {
            if (root.nodeType === 3) {
                if (!root.parentElement || !skipped(root.parentElement)) translateTextNode(root);
                return;
            }
            if (root.nodeType !== 1 && root.nodeType !== 9 && root.nodeType !== 11) return;
            if (root.nodeType === 1 && skipped(root)) return;
            if (root.nodeType === 1) translateAttrs(root);
            const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
                acceptNode(n) {
                    if (n.nodeType === 1) {
                        if (SKIP_TAGS[n.tagName] || n.hasAttribute("data-no-i18n")) return NodeFilter.FILTER_REJECT;
                        if (skipSelector && n.matches(skipSelector)) return NodeFilter.FILTER_REJECT;
                        return NodeFilter.FILTER_ACCEPT;
                    }
                    return NodeFilter.FILTER_ACCEPT;
                }
            });
            let n;
            while ((n = walker.nextNode())) {
                if (n.nodeType === 1) translateAttrs(n);
                else translateTextNode(n);
            }
        } finally {
            applying = false;
        }
    }

    function startObserver() {
        if (observer || lang === SOURCE_LANG || typeof MutationObserver === "undefined") return;
        observer = new MutationObserver((mutations) => {
            if (applying) return;
            for (const m of mutations) {
                if (m.type === "childList") {
                    m.addedNodes.forEach(n => translateTree(n));
                } else if (m.type === "characterData") {
                    translateTree(m.target);
                } else if (m.type === "attributes") {
                    const el = m.target;
                    if (!skipped(el)) {
                        applying = true;
                        try { translateAttrs(el); } finally { applying = false; }
                    }
                }
            }
        });
        observer.observe(document.documentElement, {
            childList: true, subtree: true, characterData: true,
            attributes: true, attributeFilter: ATTRS.concat(["value"])
        });
    }

    /** alert/confirm/prompt con texto traducido. */
    function wrapDialogs() {
        ["alert", "confirm", "prompt"].forEach(fn => {
            const orig = window[fn];
            if (typeof orig !== "function" || orig.__dracoI18n) return;
            const wrapped = function (msg, ...rest) {
                return orig.call(window, DracoI18n.translateString(msg), ...rest);
            };
            wrapped.__dracoI18n = true;
            window[fn] = wrapped;
        });
    }

    const DracoI18n = {
        SUPPORTED,

        /** Registra un diccionario. Claves con {param} se tratan como plantillas. */
        register(language, entries) {
            const d = dicts[language] || (dicts[language] = {});
            const p = patterns[language] || (patterns[language] = []);
            Object.keys(entries || {}).forEach(k => {
                const key = norm(k);
                if (/\{\w+\}/.test(key)) {
                    const names = [];
                    const re = new RegExp("^" + escapeRe(key).replace(/\\\{(\w+)\\\}/g, (m, n) => { names.push(n); return "(.+?)"; }) + "$");
                    p.push({ re, names, tpl: entries[k] });
                } else {
                    d[key] = entries[k];
                }
            });
            // Plantillas más específicas (más texto fijo) primero
            p.sort((a, b) => b.re.source.length - a.re.source.length);
        },

        /** Añade selectores CSS de zonas que contienen DATOS y no deben traducirse. */
        skip(selectors) {
            const list = (Array.isArray(selectors) ? selectors : [selectors]).filter(Boolean);
            skipSelector = [skipSelector].concat(list).filter(Boolean).join(", ");
        },

        get language() {
            return lang || (lang = detect());
        },

        /** Cambia el idioma (se guarda como preferencia). Requiere recargar la página. */
        setLanguage(language) {
            try {
                if (language && SUPPORTED.includes(language)) localStorage.setItem(PREF_KEY, language);
                else localStorage.removeItem(PREF_KEY);
            } catch (e) { /* sin acceso */ }
        },

        /** Traducción para usar desde el código: t("Texto con {n}", { n: 3 }). */
        t(text, params) {
            if (!lang) lang = detect();
            if (lang === SOURCE_LANG) return fill(text, params);
            const d = dicts[lang] || {};
            const key = norm(text);
            const tpl = Object.prototype.hasOwnProperty.call(d, key) ? d[key] : null;
            if (tpl !== null) return fill(tpl, params);
            const viaPattern = lookup(fill(text, params));
            return viaPattern !== null ? viaPattern : fill(text, params);
        },

        /** Traduce un texto ya montado (exacto o por plantilla); si no hay traducción, lo devuelve igual. */
        translateString(text) {
            if (text === null || text === undefined) return text;
            if (!lang) lang = detect();
            const tr = lookup(String(text));
            return tr === null ? text : tr;
        },

        translateTree,

        /** Arranca: traduce la página y vigila lo que se pinte después. */
        start() {
            lang = detect();
            document.documentElement.setAttribute("lang", lang);
            if (lang === SOURCE_LANG) return;
            wrapDialogs();
            translateTree(document.body || document.documentElement);
            startObserver();
        }
    };

    window.DracoI18n = DracoI18n;

    // Arranque automático. Si Office está presente, se espera a Office.onReady
    // para conocer el idioma de su interfaz.
    function boot() {
        if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
            Office.onReady(() => DracoI18n.start()).catch(() => DracoI18n.start());
        } else {
            DracoI18n.start();
        }
    }
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
    else boot();
})();
