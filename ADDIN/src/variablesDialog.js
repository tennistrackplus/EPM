/**
 * variablesDialog.js — ventana "Variables" (variablesDialog.html)
 * ------------------------------------------------------------------------
 * Ventana de Office independiente (como filterDialog.html): todo pasa por
 * mensajes con quien la abre (abrirVariables en commands.js):
 *
 *   padre -> diálogo
 *     { type:"init", variables:[...], fields:[{dim, attr, label}] }
 *     { type:"members", dim, attr, members:[...], error }
 *   diálogo -> padre
 *     { type:"ready" }
 *     { type:"requestMembers", dim, attr }
 *     { type:"save", variables:[...] }      (cada alta/edición/borrado)
 *     { type:"close" }
 *
 * Pantallas: listado (Editar / Borrar / Nueva variable) y editor (nombre,
 * campo y, a la derecha, selector de valor con buscador como el picker).
 */
(function () {
    "use strict";

    const S = { vars: [], fields: [], members: {}, view: "list", edit: null, confirm: null, search: "" };
    const send = (p) => { try { Office.context.ui.messageParent(JSON.stringify(p)); } catch (e) { console.error(e); } };
    const fkey = (dim, attr) => (String(dim) + "|" + String(attr)).toUpperCase();

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
    const $ = (id) => document.getElementById(id);
    const fieldLabel = (dim, attr) => {
        const f = S.fields.find(x => fkey(x.dim, x.attr) === fkey(dim, attr));
        return f ? f.label : (String(dim).toLowerCase() + " · " + String(attr).toLowerCase());
    };

    function title(text, back) {
        const h = $("title");
        h.innerHTML = "";
        if (back) h.appendChild(el("button", { class: "back-btn", title: "Volver", "aria-label": "Volver", onclick: () => renderList() }, ["←"]));
        h.appendChild(el("span", { class: "var-icon", text: "𝑥" }));
        h.appendChild(el("span", { text: text }));
    }

    function ensureMembers(dim, attr) {
        const k = fkey(dim, attr);
        if (S.members[k]) return;
        S.members[k] = { loading: true, list: [], error: null };
        send({ type: "requestMembers", dim, attr });
    }

    /* ---------------- Listado ---------------- */
    function renderList() {
        S.view = "list"; S.edit = null;
        title("Variables", false);
        const main = el("div", { class: "main" });
        main.appendChild(el("div", { style: "display:flex;justify-content:space-between;align-items:center;gap:8px" }, [
            el("span", { class: "hint", text: "Úsalas en las líneas de un eje Estático (doble clic → Variables). En la consulta, $NOMBRE$ se sustituye por su valor." }),
            el("button", { class: "btn btn-primary", onclick: () => openEditor(null) }, ["+ Nueva variable"])
        ]));
        if (!S.vars.length) {
            main.appendChild(el("div", { class: "empty", text: "Todavía no hay variables." }));
        } else {
            const t = el("div", { class: "table", role: "table" });
            t.appendChild(el("div", { class: "thead", role: "row" }, [el("span", { text: "Variable" }), el("span", { text: "Campo" }), el("span", { text: "Valor" }), el("span", { text: "" })]));
            S.vars.forEach(v => {
                const confirming = S.confirm === v.name;
                t.appendChild(el("div", { class: "trow", role: "row", ondblclick: () => openEditor(v) }, [
                    el("span", { class: "mono", text: "$" + v.name + "$" }),
                    el("span", { text: fieldLabel(v.dim, v.attr) }),
                    el("span", { title: v.value, text: v.display || v.value }),
                    el("span", { class: "actions" }, confirming
                        ? [el("span", { class: "hint", text: "¿Borrar?" }),
                           el("button", { class: "link-btn danger", onclick: () => { S.vars = S.vars.filter(x => x.name !== v.name); S.confirm = null; send({ type: "save", variables: S.vars }); renderList(); } }, ["Sí"]),
                           el("button", { class: "link-btn", onclick: () => { S.confirm = null; renderList(); } }, ["No"])]
                        : [el("button", { class: "link-btn", onclick: () => openEditor(v) }, ["Editar"]),
                           el("button", { class: "link-btn danger", onclick: () => { S.confirm = v.name; renderList(); } }, ["Borrar"])])
                ]));
            });
            main.appendChild(t);
        }
        $("body").innerHTML = ""; $("body").appendChild(main);
        $("footer").innerHTML = "";
        $("footer").appendChild(el("span", { class: "msg", id: "msg" }));
        $("footer").appendChild(el("button", { class: "btn btn-secondary", onclick: () => send({ type: "close" }) }, ["Cerrar"]));
    }

    /* ---------------- Editor ---------------- */
    function openEditor(v) {
        S.edit = v
            ? { isNew: false, orig: v.name, name: v.name, dim: v.dim, attr: v.attr, value: v.value, display: v.display || v.value }
            : { isNew: true, orig: null, name: "", dim: "", attr: "", value: "", display: "" };
        S.search = "";
        if (S.edit.dim) ensureMembers(S.edit.dim, S.edit.attr);
        renderEditor();
    }

    let sideBox = null;
    function renderEditor() {
        S.view = "editor";
        const e = S.edit;
        title(e.isNew ? "Nueva variable" : "Editar variable", true);
        const main = el("div", { class: "main with-side" });

        const name = el("input", { type: "text", id: "vName", class: "mono", value: e.name, placeholder: "PERIODO", autocomplete: "off", maxlength: "40" });
        name.addEventListener("input", () => { name.value = name.value.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "_").replace(/[^A-Z0-9_]/g, ""); e.name = name.value; });
        main.appendChild(el("div", { class: "row" }, [el("label", { for: "vName", text: "Nombre" }), name,
            el("span", { class: "hint", text: "Se usa como $NOMBRE$. Letras sin acentos, números y _." })]));

        const sel = el("select", { id: "vField" });
        sel.appendChild(el("option", { value: "", text: "— Elige un campo —" }));
        S.fields.forEach((f, i) => sel.appendChild(el("option", { value: String(i), selected: fkey(f.dim, f.attr) === fkey(e.dim, e.attr) ? true : null, text: f.label })));
        sel.addEventListener("change", () => {
            const f = S.fields[+sel.value];
            e.dim = f ? f.dim : ""; e.attr = f ? f.attr : ""; e.value = ""; e.display = ""; S.search = "";
            if (f) ensureMembers(f.dim, f.attr);
            renderSide();
            updateChosen();
        });
        main.appendChild(el("div", { class: "row" }, [el("label", { for: "vField", text: "Campo (dimensión · atributo)" }), sel]));

        main.appendChild(el("p", { class: "section-title", text: "Valor" }));
        main.appendChild(el("div", { id: "vChosen", class: "hint" }));

        sideBox = el("div", { class: "side" });
        $("body").innerHTML = "";
        $("body").appendChild(main);
        $("body").appendChild(sideBox);
        renderSide();
        updateChosen();

        $("footer").innerHTML = "";
        $("footer").appendChild(el("span", { class: "msg", id: "msg" }));
        $("footer").appendChild(el("button", { class: "btn btn-secondary", onclick: () => renderList() }, ["Cancelar"]));
        $("footer").appendChild(el("button", { class: "btn btn-primary", onclick: () => save() }, [e.isNew ? "Crear variable" : "Guardar cambios"]));
        name.focus();
    }

    function updateChosen() {
        const c = $("vChosen");
        if (!c) return;
        c.textContent = S.edit.value ? "Seleccionado: " + (S.edit.display || S.edit.value) : "Elige un valor en la lista de la derecha.";
    }

    function renderSide() {
        if (!sideBox || !S.edit) return;
        const e = S.edit;
        sideBox.innerHTML = "";
        sideBox.appendChild(el("p", { class: "section-title", text: "Valores" }));
        if (!e.dim) { sideBox.appendChild(el("div", { class: "empty", text: "Elige primero un campo." })); return; }
        const search = el("input", { type: "text", placeholder: "Buscar valor…", autocomplete: "off", value: S.search });
        sideBox.appendChild(el("div", { class: "modal-toolbar" }, [search]));
        const list = el("div", { class: "modal-items-list list", role: "listbox" });
        sideBox.appendChild(list);
        const info = S.members[fkey(e.dim, e.attr)];
        const fill = () => {
            list.innerHTML = "";
            if (!info || info.loading) { list.appendChild(el("div", { class: "loading", text: "Cargando valores…" })); return; }
            if (info.error) { list.appendChild(el("div", { class: "loading", text: "No se pudieron cargar los valores: " + info.error })); return; }
            const q = S.search.trim().toLowerCase();
            const items = info.list.filter(m => !q || m.toLowerCase().indexOf(q) !== -1).slice(0, 2000);
            if (!items.length) { list.appendChild(el("div", { class: "loading", text: "Sin resultados." })); return; }
            items.forEach(m => list.appendChild(el("div", {
                class: "modal-item-row pick" + (m === e.value ? " selected" : ""), role: "option", tabindex: "0",
                "aria-selected": m === e.value ? "true" : "false",
                onclick: () => { e.value = m; e.display = m; fill(); updateChosen(); },
                onkeydown: (ev) => { if (ev.key === "Enter") { e.value = m; e.display = m; fill(); updateChosen(); } },
                ondblclick: () => { e.value = m; e.display = m; save(); }
            }, [el("span", { text: m })])));
        };
        search.addEventListener("input", () => { S.search = search.value; fill(); });
        fill();
    }

    function save() {
        const e = S.edit, msg = $("msg");
        const name = String(e.name || "").trim().toUpperCase();
        if (!name) { msg.textContent = "Ponle un nombre a la variable."; return; }
        if (!/^[A-Z][A-Z0-9_]{0,39}$/.test(name)) { msg.textContent = "El nombre debe empezar por letra (sin acentos), con números o _."; return; }
        if (S.vars.some(v => v.name === name && v.name !== e.orig)) { msg.textContent = `Ya hay una variable $${name}$.`; return; }
        if (!e.dim) { msg.textContent = "Elige el campo."; return; }
        if (!e.value) { msg.textContent = "Elige un valor."; return; }
        const v = { name, dim: e.dim, attr: e.attr, value: e.value, display: e.display || e.value };
        S.vars = S.vars.filter(x => x.name !== e.orig && x.name !== name).concat([v]).sort((a, b) => a.name.localeCompare(b.name));
        send({ type: "save", variables: S.vars });
        renderList();
    }

    /* ---------------- Mensajes ---------------- */
    function onParent(arg) {
        let p;
        try { p = JSON.parse(arg.message); } catch (e) { return; }
        if (p.type === "init") {
            S.vars = (p.variables || []).slice().sort((a, b) => a.name.localeCompare(b.name));
            S.fields = p.fields || [];
            renderList();
        } else if (p.type === "members") {
            S.members[fkey(p.dim, p.attr)] = { loading: false, list: p.members || [], error: p.error || null };
            if (S.view === "editor" && S.edit && fkey(S.edit.dim, S.edit.attr) === fkey(p.dim, p.attr)) renderSide();
        }
    }

    $("btnClose").addEventListener("click", () => send({ type: "close" }));
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") { if (S.view === "editor") renderList(); else send({ type: "close" }); } });

    Office.onReady(() => {
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParent, () => send({ type: "ready" }));
    });
})();
