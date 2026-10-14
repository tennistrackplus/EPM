/**
 * calcDialog.js — ventana "Cálculos" (calcDialog.html)
 * ------------------------------------------------------------------------
 * Ventana de Office independiente, igual que filterDialog.html /
 * memberPicker.html: no tiene acceso a Excel ni a ExcelService, así que
 * todo pasa por mensajes con el taskpane (ver js/calcEditor.js):
 *
 *   taskpane -> diálogo
 *     { type:"init", dim, attr|null, attributes:[...], calcs:{ATTR:[...]} }
 *     { type:"members", attr, members:[...], visible:[...]|null, error }
 *     { type:"saved", attr, calcs:[...] }
 *   diálogo -> taskpane
 *     { type:"ready" }
 *     { type:"requestMembers", attr }
 *     { type:"save", attr, calcs:[...] }   (cada alta/edición/borrado)
 *     { type:"close" }
 *
 * Dos pantallas:
 *   1. Listado de los cálculos del atributo (o de la dimensión, si se
 *      entró desde la cabecera de la dimensión): Editar / Borrar / Nuevo.
 *   2. Editor de UN cálculo: nombre, tipo (Miembro fijo / Posición),
 *      fórmula, formato y visualización, con un selector lateral de
 *      miembros o posiciones con buscador, como el del picker.
 */
(function () {
    "use strict";

    const S = {
        dim: "", fixedAttr: null, attributes: [], calcs: {},  // calcs: ATTR(upper) -> [calc]
        members: {},          // ATTR(upper) -> { loading, members, visible, error }
        view: "list",
        edit: null,           // { attr, calc (copia), isNew }
        confirmDeleteId: null,
        search: ""
    };

    const send = (payload) => {
        try { Office.context.ui.messageParent(JSON.stringify(payload)); } catch (e) { console.error(e); }
    };
    const U = (s) => String(s || "").toUpperCase();

    const el = (tag, attrs, children) => {
        const n = document.createElement(tag);
        Object.entries(attrs || {}).forEach(([k, v]) => {
            if (v === undefined || v === null || v === false) return;
            if (k === "class") n.className = v;
            else if (k === "text") n.textContent = v;
            else if (k.startsWith("on")) n.addEventListener(k.slice(2), v);
            else n.setAttribute(k, v === true ? "" : v);
        });
        (children || []).forEach(c => c !== null && c !== undefined && n.appendChild(typeof c === "string" ? document.createTextNode(c) : c));
        return n;
    };

    const body = () => document.getElementById("dialogBody");
    const footer = () => document.getElementById("dialogFooter");

    function setTitle(parts, withBack) {
        const h = document.getElementById("dialogTitle");
        h.innerHTML = "";
        if (withBack) {
            h.appendChild(el("button", { class: "dialog-back-btn", title: "Volver al listado", "aria-label": "Volver al listado", onclick: () => backToList() }, ["←"]));
        }
        h.appendChild(el("span", { class: "fx", text: "ƒx" }));
        parts.forEach(p => h.appendChild(typeof p === "string" ? document.createTextNode(p) : p));
    }

    function calcsOf(attr) { return S.calcs[U(attr)] || []; }

    function ensureMembers(attr) {
        const k = U(attr);
        if (S.members[k]) return;
        S.members[k] = { loading: true, members: [], visible: null, error: null };
        send({ type: "requestMembers", attr });
    }

    /* ======================= LISTADO ======================= */
    function renderList() {
        S.view = "list";
        const title = S.fixedAttr
            ? [" Cálculos de ", el("b", { text: S.fixedAttr.toLowerCase() })]
            : [" Cálculos de la dimensión ", el("b", { text: S.dim.toLowerCase() })];
        setTitle(title, false);

        const rows = [];
        const attrs = S.fixedAttr ? [S.fixedAttr] : S.attributes;
        attrs.forEach(a => calcsOf(a).forEach(c => rows.push({ attr: a, calc: c })));

        const main = el("div", { class: "dialog-main" });
        main.appendChild(el("div", { class: "list-toolbar" }, [
            el("span", { class: "hint", text: rows.length
                ? "Los cálculos se añaden como una fila o columna más, detrás de los miembros del campo."
                : "" }),
            el("button", { class: "btn btn-primary", onclick: () => openEditor(null, null) }, ["+ Nuevo cálculo"])
        ]));

        if (rows.length === 0) {
            main.appendChild(el("div", { class: "empty", text: S.fixedAttr
                ? "Este campo todavía no tiene cálculos. Pulsa «+ Nuevo cálculo» para crear el primero (por ejemplo, Actual − Budget)."
                : "Ningún atributo de esta dimensión tiene cálculos todavía." }));
        } else {
            const withAttr = !S.fixedAttr;
            const table = el("div", { class: "calc-table" + (withAttr ? " with-attr" : ""), role: "table" });
            table.appendChild(el("div", { class: "calc-table-head", role: "row" }, [
                withAttr ? el("span", { text: "Atributo" }) : null,
                el("span", { text: "Nombre" }), el("span", { text: "Tipo" }), el("span", { text: "Fórmula" }), el("span", { text: "" })
            ]));
            rows.forEach(({ attr, calc }) => {
                const ok = window.DracoCalc.analyze(calc.expr, calc.mode).ok;
                const confirming = S.confirmDeleteId === calc.id;
                const actions = confirming
                    ? [el("span", { class: "hint", text: "¿Borrar?" }),
                       el("button", { class: "link-btn danger", onclick: (e) => { e.stopPropagation(); deleteCalc(attr, calc.id); } }, ["Sí"]),
                       el("button", { class: "link-btn", onclick: (e) => { e.stopPropagation(); S.confirmDeleteId = null; renderList(); } }, ["No"])]
                    : [el("button", { class: "link-btn", onclick: (e) => { e.stopPropagation(); openEditor(attr, calc); } }, ["Editar"]),
                       el("button", { class: "link-btn danger", onclick: (e) => { e.stopPropagation(); S.confirmDeleteId = calc.id; renderList(); } }, ["Borrar"])];
                table.appendChild(el("div", { class: "calc-row", role: "row", title: "Doble clic para editar", ondblclick: () => openEditor(attr, calc) }, [
                    withAttr ? el("span", { class: "calc-type", text: attr.toLowerCase() }) : null,
                    el("span", { class: "calc-name" }, [ok ? null : el("span", { class: "bad-badge", title: "La fórmula tiene errores", text: "!" }), el("span", { text: calc.name })]),
                    el("span", { class: "calc-type", text: calc.mode === "position" ? "Posición" : "Miembro fijo" }),
                    el("span", { class: "calc-formula-text", title: calc.expr, text: calc.expr }),
                    el("span", { class: "calc-actions" }, actions)
                ]));
            });
            main.appendChild(table);
        }

        body().innerHTML = "";
        body().appendChild(main);
        footer().innerHTML = "";
        footer().appendChild(el("span", { class: "footer-msg", id: "footerMsg" }));
        footer().appendChild(el("button", { class: "btn btn-secondary", onclick: () => send({ type: "close" }) }, ["Cerrar"]));
    }

    function deleteCalc(attr, id) {
        const list = calcsOf(attr).filter(c => c.id !== id);
        S.calcs[U(attr)] = list;
        S.confirmDeleteId = null;
        send({ type: "save", attr, calcs: list });
        renderList();
    }

    /* ======================= EDITOR ======================= */
    function openEditor(attr, calc) {
        S.confirmDeleteId = null;
        const isNew = !calc;
        const targetAttr = attr || S.fixedAttr || S.attributes[0];
        let c;
        if (isNew) {
            c = window.DracoCalc.defaults();
            let n = 1, name = "Variación";
            while (calcsOf(targetAttr).some(x => x.name === name)) name = "Variación " + (++n);
            c.name = name;
        } else {
            c = Object.assign(window.DracoCalc.defaults(), JSON.parse(JSON.stringify(calc)));
        }
        S.edit = { attr: targetAttr, calc: c, isNew };
        S.search = "";
        ensureMembers(targetAttr);
        renderEditor();
    }

    function backToList() {
        S.edit = null;
        renderList();
    }

    let formulaBox = null, statusEl = null, sideBox = null;

    function renderEditor() {
        S.view = "editor";
        const { calc: c, isNew } = S.edit;
        setTitle([isNew ? " Nuevo cálculo" : " Editar cálculo", el("span", { class: "hint", text: " · " + S.edit.attr.toLowerCase() })], true);

        const main = el("div", { class: "dialog-main with-side" });

        // Atributo: SOLO al crear desde la cabecera de la dimensión (desde
        // un atributo ya se sabe a cuál pertenece).
        if (isNew && !S.fixedAttr && S.attributes.length > 1) {
            const sel = el("select", { id: "fAttr" });
            S.attributes.forEach(a => sel.appendChild(el("option", { value: a, selected: a === S.edit.attr ? true : null, text: a.toLowerCase() })));
            sel.addEventListener("change", () => { S.edit.attr = sel.value; ensureMembers(sel.value); renderEditor(); });
            main.appendChild(el("div", { class: "form-row" }, [el("label", { for: "fAttr", text: "Campo" }), sel]));
        }

        const name = el("input", { type: "text", id: "fName", value: c.name, autocomplete: "off" });
        name.addEventListener("input", () => { c.name = name.value; });
        main.appendChild(el("div", { class: "form-row" }, [el("label", { for: "fName", text: "Nombre" }), name]));

        // Tipo
        const seg = el("div", { class: "segmented-control type-switch", role: "tablist" });
        [["member", "Miembro fijo"], ["position", "Posición"]].forEach(([v, label]) => {
            seg.appendChild(el("button", {
                type: "button", class: "segmented-option" + (c.mode === v ? " active" : ""), role: "tab",
                onclick: () => {
                    if (c.mode === v) return;
                    c.mode = v;
                    c.expr = v === "position" ? "[#1] - [#2]" : "";
                    S.search = "";
                    renderEditor();
                }
            }, [label]));
        });
        main.appendChild(el("div", { class: "form-row" }, [el("span", { class: "form-label", text: "Tipo" }), seg,
            el("span", { class: "hint", text: c.mode === "member"
                ? "Opera con miembros concretos (p. ej. [Actual] − [Budget]). Se calcula en la consulta, aunque esos miembros no estén en el filtro del informe."
                : "Opera con los miembros que se ven, por su orden: [#1] es el primero, [#2] el segundo… Se escribe como fórmula de Excel." })]));

        // Fórmula
        formulaBox = el("textarea", { class: "formula-box", id: "fExpr", spellcheck: "false", rows: "3" });
        formulaBox.value = c.expr || "";
        formulaBox.addEventListener("input", () => { c.expr = formulaBox.value; updateStatus(); });
        statusEl = el("div", { class: "status" });
        const ops = el("div", { class: "ops" });
        [["+", "+ "], ["−", "- "], ["×", "* "], ["÷", "/ "], ["(", "("], [")", ")"]].forEach(([lbl, ins]) =>
            ops.appendChild(el("button", { type: "button", class: "op", title: lbl, onclick: () => insert(ins) }, [lbl])));
        main.appendChild(el("div", { class: "form-row" }, [
            el("label", { for: "fExpr", text: "Fórmula" }), formulaBox, ops, statusEl,
            el("span", { class: "hint", text: c.mode === "member"
                ? "Haz clic en un miembro de la lista de la derecha para insertarlo."
                : "Haz clic en una posición de la lista de la derecha para insertarla." })
        ]));

        // Formato
        const nf = el("select", { id: "fNf" });
        [["number", "Número (1.234)"], ["percent", "Porcentaje (12%)"]].forEach(([v, l]) => nf.appendChild(el("option", { value: v, selected: c.numberFormat === v ? true : null, text: l })));
        nf.addEventListener("change", () => { c.numberFormat = nf.value; });
        const dec = el("input", { type: "number", id: "fDec", min: "0", max: "10", value: String(c.decimals) });
        dec.addEventListener("input", () => { c.decimals = Math.max(0, Math.min(10, Math.floor(Number(dec.value) || 0))); });
        main.appendChild(el("p", { class: "section-title", text: "Formato" }));
        main.appendChild(el("div", { class: "grid-2" }, [
            el("div", { class: "form-row" }, [el("label", { for: "fNf", text: "Formato de número" }), nf]),
            el("div", { class: "form-row" }, [el("label", { for: "fDec", text: "Decimales" }), dec])
        ]));

        // Visualización
        const disp = el("select", { id: "fDisp" });
        [["number", "Número"], ["bar", "Número y barra"], ["barOnly", "Solo barra"]].forEach(([v, l]) => disp.appendChild(el("option", { value: v, selected: c.display === v ? true : null, text: l })));
        const barBox = el("div", { style: "display:flex; flex-direction:column; gap:10px;" });
        const renderBar = () => {
            barBox.innerHTML = "";
            if (c.display === "number") return;
            const cn = el("input", { type: "color", class: "color-input", id: "fCn", value: c.colorNeg });
            cn.addEventListener("input", () => { c.colorNeg = cn.value; });
            const cp = el("input", { type: "color", class: "color-input", id: "fCp", value: c.colorPos });
            cp.addEventListener("input", () => { c.colorPos = cp.value; });
            barBox.appendChild(el("div", { class: "grid-2" }, [
                el("div", { class: "form-row" }, [el("label", { for: "fCn", text: "Color negativo" }), cn]),
                el("div", { class: "form-row" }, [el("label", { for: "fCp", text: "Color positivo" }), cp])
            ]));
            const inv = el("input", { type: "checkbox", id: "fInv", checked: c.invertColors ? true : null });
            inv.addEventListener("change", () => { c.invertColors = inv.checked; });
            barBox.appendChild(el("label", { class: "check-row", for: "fInv" }, [inv, "Invertir colores (negativo es bueno, p. ej. en costes)"]));
            if (c.display === "bar") {
                const ll = el("input", { type: "number", id: "fLl", min: "0", max: "90", value: String(c.labelLeft) });
                ll.addEventListener("input", () => { c.labelLeft = Math.max(0, Math.min(90, Number(ll.value) || 0)); });
                const lr = el("input", { type: "number", id: "fLr", min: "0", max: "90", value: String(c.labelRight) });
                lr.addEventListener("input", () => { c.labelRight = Math.max(0, Math.min(90, Number(lr.value) || 0)); });
                barBox.appendChild(el("div", { class: "grid-2" }, [
                    el("div", { class: "form-row" }, [el("label", { for: "fLl", text: "Espacio nº negativo (%)" }), ll]),
                    el("div", { class: "form-row" }, [el("label", { for: "fLr", text: "Espacio nº positivo (%)" }), lr])
                ]));
            }
        };
        disp.addEventListener("change", () => { c.display = disp.value; renderBar(); });
        main.appendChild(el("p", { class: "section-title", text: "Visualización" }));
        main.appendChild(el("div", { class: "form-row" }, [el("label", { for: "fDisp", text: "Mostrar como" }), disp]));
        renderBar();
        main.appendChild(barBox);
        main.appendChild(el("span", { class: "hint", text: "Formato y barras se aplican al actualizar si «Sobrescribir formatos» está activado en las propiedades del informe." }));

        // Selector lateral
        sideBox = el("div", { class: "dialog-side" });

        body().innerHTML = "";
        body().appendChild(main);
        body().appendChild(sideBox);
        renderSide();
        updateStatus();

        footer().innerHTML = "";
        footer().appendChild(el("span", { class: "footer-msg", id: "footerMsg" }));
        footer().appendChild(el("button", { class: "btn btn-secondary", onclick: () => backToList() }, ["Cancelar"]));
        footer().appendChild(el("button", { class: "btn btn-primary", onclick: () => saveEditor() }, [isNew ? "Crear cálculo" : "Guardar cambios"]));
        name.focus();
    }

    function renderSide() {
        if (!sideBox || !S.edit) return;
        const c = S.edit.calc;
        sideBox.innerHTML = "";

        if (c.mode === "position") {
            sideBox.appendChild(el("p", { class: "section-title", text: "Posiciones" }));
            sideBox.appendChild(el("span", { class: "hint", text: "En el orden en que se ven los miembros del campo en el informe." }));
            const list = el("div", { class: "modal-items-list picker-list" });
            for (let i = 1; i <= 12; i++) {
                list.appendChild(el("div", { class: "modal-item-row picker-row", role: "button", tabindex: "0",
                    onclick: () => insert(`[#${i}]`),
                    onkeydown: (e) => { if (e.key === "Enter") insert(`[#${i}]`); }
                }, [el("span", { text: `${i}º miembro visible` }), el("span", { class: "code", text: `[#${i}]` })]));
            }
            sideBox.appendChild(list);
            return;
        }

        sideBox.appendChild(el("p", { class: "section-title", text: "Miembros" }));
        const search = el("input", { type: "text", placeholder: "Buscar miembro…", autocomplete: "off", value: S.search });
        sideBox.appendChild(el("div", { class: "modal-toolbar" }, [search]));
        const list = el("div", { class: "modal-items-list picker-list" });
        sideBox.appendChild(list);

        const info = S.members[U(S.edit.attr)];
        const fill = () => {
            list.innerHTML = "";
            if (!info || info.loading) { list.appendChild(el("div", { class: "loading", text: "Cargando miembros…" })); return; }
            if (info.error) {
                list.appendChild(el("div", { class: "loading", text: "No se pudieron cargar los miembros. Puedes escribirlos entre corchetes, p. ej. [Actual]." }));
                return;
            }
            const q = S.search.trim().toLowerCase();
            const items = info.members.filter(m => !q || m.toLowerCase().indexOf(q) !== -1);
            if (items.length === 0) { list.appendChild(el("div", { class: "loading", text: "Sin resultados." })); return; }
            items.slice(0, 1000).forEach(m => {
                const hidden = info.visible && info.visible.indexOf(m) === -1;
                list.appendChild(el("div", {
                    class: "modal-item-row picker-row", role: "button", tabindex: "0",
                    title: hidden ? "No está en el filtro del informe: se usará solo para el cálculo" : "Insertar en la fórmula",
                    onclick: () => insert(`[${m}]`),
                    onkeydown: (e) => { if (e.key === "Enter") insert(`[${m}]`); }
                }, [el("span", { text: m }), hidden ? el("span", { class: "tag", text: "no visible" }) : null]));
            });
        };
        search.addEventListener("input", () => { S.search = search.value; fill(); });
        fill();
        if (info && info.visible) {
            sideBox.appendChild(el("span", { class: "hint", text: "«no visible»: no está en el filtro del informe; se lee solo para el cálculo y no se pinta." }));
        }
    }

    function insert(text) {
        if (!formulaBox) return;
        const s = formulaBox.selectionStart ?? formulaBox.value.length;
        const e = formulaBox.selectionEnd ?? formulaBox.value.length;
        const before = formulaBox.value.slice(0, s), after = formulaBox.value.slice(e);
        const pad = (before && !/[\s(]$/.test(before)) ? " " : "";
        formulaBox.value = before + pad + text + after;
        const pos = (before + pad + text).length;
        formulaBox.focus();
        formulaBox.setSelectionRange(pos, pos);
        S.edit.calc.expr = formulaBox.value;
        updateStatus();
    }

    function updateStatus() {
        if (!statusEl || !S.edit) return;
        const c = S.edit.calc;
        if (!String(c.expr || "").trim()) { statusEl.className = "status"; statusEl.textContent = ""; return; }
        const a = window.DracoCalc.analyze(c.expr, c.mode);
        statusEl.className = "status " + (a.ok ? "ok" : "bad");
        if (!a.ok) { statusEl.textContent = a.error; return; }
        if (c.mode === "member") {
            statusEl.textContent = "Fórmula válida · se calcula en la consulta";
        } else {
            const cols = "CDEFGHIJKLMNOP";
            const f = window.DracoCalc.toExcel(a.ast, (p) => (cols[p - 1] || "C") + "6");
            statusEl.textContent = "Fórmula válida · en Excel, p. ej.: " + (f ? f.replace(/^=IFERROR\((.*),""\)$/, "=$1") : "");
        }
    }

    function saveEditor() {
        const msg = document.getElementById("footerMsg");
        const { attr, calc: c, isNew } = S.edit;
        const nm = String(c.name || "").trim();
        if (!nm) { msg.textContent = "Ponle un nombre al cálculo."; return; }
        const others = calcsOf(attr).filter(x => x.id !== c.id);
        if (others.some(x => x.name.toLowerCase() === nm.toLowerCase())) { msg.textContent = `Ya hay un cálculo llamado «${nm}».`; return; }
        const info = S.members[U(attr)];
        if (info && info.members.some(m => m.toLowerCase() === nm.toLowerCase())) {
            msg.textContent = `«${nm}» ya es un miembro de ${attr.toLowerCase()}: usa otro nombre.`; return;
        }
        const a = window.DracoCalc.analyze(c.expr, c.mode);
        if (!a.ok) { msg.textContent = a.error; return; }

        c.name = nm;
        c.position = "end";
        const clean = {
            id: c.id, name: c.name, mode: c.mode, expr: c.expr, position: "end",
            numberFormat: c.numberFormat, decimals: c.decimals, display: c.display,
            colorNeg: c.colorNeg, colorPos: c.colorPos, invertColors: !!c.invertColors,
            labelLeft: c.labelLeft, labelRight: c.labelRight
        };
        const list = isNew ? [...calcsOf(attr), clean] : calcsOf(attr).map(x => x.id === c.id ? clean : x);
        S.calcs[U(attr)] = list;
        send({ type: "save", attr, calcs: list });
        backToList();
    }

    /* ======================= MENSAJES ======================= */
    function onParentMessage(arg) {
        let p;
        try { p = JSON.parse(arg.message); } catch (e) { return; }
        if (p.type === "init") {
            S.dim = p.dim;
            S.fixedAttr = p.attr || null;
            S.attributes = p.attributes || (p.attr ? [p.attr] : []);
            S.calcs = {};
            Object.keys(p.calcs || {}).forEach(k => { S.calcs[U(k)] = p.calcs[k]; });
            renderList();
            return;
        }
        if (p.type === "members") {
            S.members[U(p.attr)] = { loading: false, members: p.members || [], visible: p.visible || null, error: p.error || null };
            if (S.view === "editor" && S.edit && U(S.edit.attr) === U(p.attr)) renderSide();
        }
    }

    document.getElementById("btnCloseDialog").addEventListener("click", () => send({ type: "close" }));
    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
            if (S.view === "editor") backToList();
            else send({ type: "close" });
        }
    });

    Office.onReady(() => {
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParentMessage, () => {
            send({ type: "ready" });
        });
    });
})();
