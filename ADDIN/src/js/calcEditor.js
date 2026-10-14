/* ==========================================================
 * Editor de cálculos entre miembros (botón ƒx del árbol de campos)
 * ----------------------------------------------------------
 * Pantalla modal (mismas clases que el resto de modales del
 * taskpane: modal-backdrop / modal-dialog / modal-footer, botones
 * btn-primary / btn-secondary, segmented-control) para crear,
 * editar y borrar los cálculos de UN atributo en el informe actual.
 *
 * Los cálculos se guardan en TaskPaneApp.state.fieldOptions bajo
 * DracoCalc.key(dim, attr), así que se persisten, copian y pegan
 * con el diseño del informe; al guardar se lanza el autoguardado
 * + refresco habitual (scheduleAutoUpdate).
 * ========================================================== */
(function () {
    "use strict";

    const el = (tag, attrs, children) => {
        const n = document.createElement(tag);
        Object.entries(attrs || {}).forEach(([k, v]) => {
            if (k === "class") n.className = v;
            else if (k === "text") n.textContent = v;
            else if (k === "style") n.setAttribute("style", v);
            else if (k.startsWith("on")) n.addEventListener(k.slice(2), v);
            else if (v !== undefined && v !== null && v !== false) n.setAttribute(k, v === true ? "" : v);
        });
        (children || []).forEach(c => c && n.appendChild(typeof c === "string" ? document.createTextNode(c) : c));
        return n;
    };

    const field = (label, control, hint) => el("label", {}, [label, control, hint ? el("span", { class: "calc-hint", text: hint }) : null]);

    const CalcEditor = {
        app: null,
        dim: null,
        attr: null,
        attributes: [],
        calcs: [],
        selectedId: null,
        members: null,       // null = sin cargar, [] = cargados
        membersError: null,
        root: null,

        /**
         * Abre el editor.
         *   app        : TaskPaneApp
         *   dimension  : nombre de la dimensión
         *   attribute  : atributo inicial (o null -> el primero)
         *   attributes : atributos de la dimensión (para el selector)
         */
        open(app, dimension, attribute, attributes) {
            if (!window.DracoCalc) return;
            if (!app.currentReportId) {
                alert("Crea o selecciona un informe antes de definir cálculos.");
                return;
            }
            this.app = app;
            this.dim = dimension;
            this.attributes = (attributes && attributes.length) ? attributes.slice() : [attribute];
            this.loadAttribute(attribute || this.attributes[0]);
            this.render();
        },

        loadAttribute(attribute) {
            this.attr = attribute;
            const stored = window.DracoCalc.getCalcs(this.app.state.fieldOptions, this.dim, attribute);
            this.calcs = stored.map(c => Object.assign(window.DracoCalc.defaults(), c));
            this.selectedId = this.calcs.length ? this.calcs[0].id : null;
            this.members = null;
            this.membersError = null;
            this.fetchMembers();
        },

        async fetchMembers() {
            const dim = this.dim, attr = this.attr;
            try {
                const sql = await window.ExcelService.buildFilterValuesSQL(dim, attr);
                if (!sql) throw new Error("No se encuentra el atributo en el modelo.");
                const json = await window.ExcelService.executeSQL(sql);
                if (dim !== this.dim || attr !== this.attr) return; // se cambió de atributo mientras tanto
                const items = (typeof loadJsonTree === "function") ? loadJsonTree(json) : [];
                this.members = Array.from(new Set(items.map(i => String(i.value)))).slice(0, 500);
            } catch (e) {
                this.members = [];
                this.membersError = (e && e.message) || String(e);
            }
            if (this.root) this.renderEditor();
        },

        // Miembros del atributo que el informe muestra según su filtro (null = sin filtro: todos)
        visibleMembers() {
            const f = (this.app.state.filters || []).find(x =>
                String(x.dimension).toUpperCase() === String(this.dim).toUpperCase()
                && String(x.realAttribute || x.name).toUpperCase() === String(this.attr).toUpperCase());
            if (!f || !f.filter) return null;
            const vals = [];
            (f.filter.values || []).forEach(v => vals.push(String(v)));
            (f.filter.items || []).forEach(it => vals.push(String(it && typeof it === "object" ? (it.value !== undefined ? it.value : it.display) : it)));
            return vals.length ? vals : null;
        },

        selected() {
            return this.calcs.find(c => c.id === this.selectedId) || null;
        },

        close() {
            if (this.root) this.root.remove();
            this.root = null;
        },

        /* ---------------- Render ---------------- */
        render() {
            this.close();
            const header = el("div", { class: "modal-header" }, [
                el("h3", {}, [el("span", { class: "calc-fx", text: "ƒx" }), " Cálculos de ", el("span", { text: this.dim.toLowerCase() })]),
                el("button", { class: "modal-close-btn", title: "Cerrar", "aria-label": "Cerrar", onclick: () => this.close() }, ["×"])
            ]);

            const attrSelect = el("select", { id: "calcAttr" });
            this.attributes.forEach(a => attrSelect.appendChild(el("option", { value: a, selected: a === this.attr, text: a.toLowerCase() })));
            attrSelect.addEventListener("change", () => {
                if (this.isDirty && !confirm("Hay cambios sin guardar en este atributo. ¿Descartarlos?")) {
                    attrSelect.value = this.attr;
                    return;
                }
                this.isDirty = false;
                this.loadAttribute(attrSelect.value);
                this.renderList();
                this.renderEditor();
            });

            this.listBox = el("div", { class: "calc-list" });
            this.editorBox = el("div", { class: "calc-editor field-options-group" });

            const body = el("div", { class: "modal-body" }, [
                el("div", { class: "field-options-group" }, [field("Atributo", attrSelect)]),
                this.listBox,
                this.editorBox
            ]);

            this.errorEl = el("span", { class: "calc-footer-msg" });
            const footer = el("div", { class: "modal-footer" }, [
                this.errorEl,
                el("button", { class: "btn-secondary", onclick: () => this.close() }, ["Cancelar"]),
                el("button", { class: "btn-primary", onclick: () => this.save() }, ["Guardar"])
            ]);

            const dialog = el("div", { class: "modal-dialog calc-dialog", role: "dialog", "aria-modal": "true", "aria-label": "Cálculos" }, [header, body, footer]);
            this.root = el("div", { class: "modal-backdrop" }, [dialog]);
            this.root.addEventListener("mousedown", (e) => { if (e.target === this.root) this.close(); });
            document.body.appendChild(this.root);
            this.isDirty = false;
            this.renderList();
            this.renderEditor();
        },

        renderList() {
            const box = this.listBox;
            box.innerHTML = "";
            this.calcs.forEach(c => {
                const ok = window.DracoCalc.analyze(c.expr, c.mode).ok;
                const row = el("button", {
                    type: "button",
                    class: "calc-list-item" + (c.id === this.selectedId ? " active" : ""),
                    onclick: () => { this.selectedId = c.id; this.renderList(); this.renderEditor(); }
                }, [
                    el("span", { class: "calc-fx", text: "ƒx" }),
                    el("span", { class: "calc-list-name", text: c.name || "(sin nombre)" }),
                    el("span", { class: "calc-list-mode", text: c.mode === "position" ? "posición" : "miembro" }),
                    ok ? null : el("span", { class: "calc-list-bad", title: "Fórmula con errores", text: "!" })
                ]);
                box.appendChild(row);
            });
            box.appendChild(el("button", {
                type: "button", class: "calc-add-btn",
                onclick: () => {
                    const c = window.DracoCalc.defaults();
                    let n = 1, base = "Variación", name = base;
                    while (this.calcs.some(x => x.name === name)) name = base + " " + (++n);
                    c.name = name;
                    this.calcs.push(c);
                    this.selectedId = c.id;
                    this.isDirty = true;
                    this.renderList();
                    this.renderEditor();
                }
            }, ["+ Nuevo cálculo"]));
        },

        renderEditor() {
            const box = this.editorBox;
            if (!box) return;
            box.innerHTML = "";
            const c = this.selected();
            if (!c) {
                box.appendChild(el("div", { class: "field-options-empty", text: "Este atributo no tiene cálculos. Pulsa «+ Nuevo cálculo» para crear uno." }));
                return;
            }
            const touch = () => { this.isDirty = true; };
            const refreshList = () => this.renderList();

            // --- Nombre
            const name = el("input", { type: "text", value: c.name });
            name.addEventListener("input", () => { c.name = name.value; touch(); refreshList(); });

            // --- Tipo (segmentado)
            const seg = el("div", { class: "segmented-control" });
            [["member", "Por miembro (fijo)"], ["position", "Por posición (variable)"]].forEach(([v, label]) => {
                seg.appendChild(el("button", {
                    type: "button", class: "segmented-option" + (c.mode === v ? " active" : ""),
                    onclick: () => {
                        if (c.mode === v) return;
                        c.mode = v;
                        c.expr = v === "position" ? "[#1] - [#2]" : "";
                        touch(); refreshList(); this.renderEditor();
                    }
                }, [label]));
            });
            const modeHint = c.mode === "member"
                ? "Opera con miembros concretos. Se calcula en la consulta, aunque esos miembros no estén en el filtro del informe."
                : "Opera con los miembros que se ven, por su orden (1º, 2º…). Se escribe como fórmula de Excel.";

            // --- Fórmula
            const ta = el("textarea", { class: "calc-formula", rows: "3", spellcheck: "false" });
            ta.value = c.expr || "";
            const status = el("div", { class: "calc-status" });
            const updateStatus = () => {
                const a = window.DracoCalc.analyze(ta.value, c.mode);
                status.className = "calc-status " + (a.ok ? "ok" : "bad");
                status.textContent = a.ok
                    ? (c.mode === "member" ? "Fórmula válida · se calcula en SQL" : "Fórmula válida · ej. " + this.excelExample(a.ast))
                    : a.error;
            };
            ta.addEventListener("input", () => { c.expr = ta.value; touch(); updateStatus(); refreshList(); });
            const insert = (text) => {
                const s = ta.selectionStart ?? ta.value.length, e = ta.selectionEnd ?? ta.value.length;
                const before = ta.value.slice(0, s), after = ta.value.slice(e);
                const pad = (before && !/[\s(]$/.test(before)) ? " " : "";
                ta.value = before + pad + text + after;
                const pos = (before + pad + text).length;
                ta.focus(); ta.setSelectionRange(pos, pos);
                c.expr = ta.value; touch(); updateStatus(); refreshList();
            };

            const ops = el("div", { class: "calc-ops" });
            [["+", "+ "], ["−", "- "], ["×", "* "], ["÷", "/ "], ["(", "("], [")", ")"]].forEach(([lbl, ins]) =>
                ops.appendChild(el("button", { type: "button", class: "calc-op", title: lbl, onclick: () => insert(ins) }, [lbl])));

            const chips = el("div", { class: "calc-chips" });
            if (c.mode === "position") {
                for (let i = 1; i <= 6; i++) {
                    chips.appendChild(el("button", { type: "button", class: "calc-chip", title: `Miembro visible nº ${i}`, onclick: () => insert(`[#${i}]`) }, [`${i}º`]));
                }
            } else if (this.members === null) {
                chips.appendChild(el("span", { class: "calc-hint", text: "Cargando miembros…" }));
            } else if (this.members.length === 0) {
                chips.appendChild(el("span", { class: "calc-hint", text: this.membersError ? "No se pudieron cargar los miembros: escríbelos entre corchetes, p. ej. [Actual]." : "Sin miembros." }));
            } else {
                const visible = this.visibleMembers();
                this.members.forEach(m => {
                    const hidden = visible && visible.indexOf(m) === -1;
                    chips.appendChild(el("button", {
                        type: "button", class: "calc-chip" + (hidden ? " hidden-member" : ""),
                        title: hidden ? "No está en el filtro del informe: se usa solo para el cálculo" : m,
                        onclick: () => insert(`[${m}]`)
                    }, [m, hidden ? el("span", { class: "calc-chip-note", text: " · no visible" }) : null]));
                });
            }

            // --- Posición
            const posSel = el("select");
            [["end", "Al final"], ["start", "Al principio"]].forEach(([v, l]) => posSel.appendChild(el("option", { value: v, selected: c.position === v, text: l })));
            posSel.addEventListener("change", () => { c.position = posSel.value; touch(); });

            // --- Formato
            const nf = el("select");
            [["number", "Número (1.234)"], ["percent", "Porcentaje (12%)"]].forEach(([v, l]) => nf.appendChild(el("option", { value: v, selected: c.numberFormat === v, text: l })));
            nf.addEventListener("change", () => { c.numberFormat = nf.value; touch(); });
            const dec = el("input", { type: "number", min: "0", max: "10", value: String(c.decimals) });
            dec.addEventListener("input", () => { c.decimals = Math.max(0, Math.min(10, Math.floor(Number(dec.value) || 0))); touch(); });

            // --- Visualización
            const disp = el("select");
            [["number", "Número"], ["bar", "Número y barra"], ["barOnly", "Solo barra"]].forEach(([v, l]) => disp.appendChild(el("option", { value: v, selected: c.display === v, text: l })));
            const barBox = el("div", { class: "calc-bar-options" });
            const renderBar = () => {
                barBox.innerHTML = "";
                if (c.display === "number") return;
                const cn = el("input", { type: "color", value: c.colorNeg });
                cn.addEventListener("input", () => { c.colorNeg = cn.value; touch(); });
                const cp = el("input", { type: "color", value: c.colorPos });
                cp.addEventListener("input", () => { c.colorPos = cp.value; touch(); });
                const inv = el("input", { type: "checkbox", checked: !!c.invertColors });
                inv.addEventListener("change", () => { c.invertColors = inv.checked; touch(); });
                barBox.appendChild(el("div", { class: "calc-grid-2" }, [field("Color negativo", cn), field("Color positivo", cp)]));
                barBox.appendChild(el("label", { class: "field-options-checkbox-row" }, [inv, el("span", { text: "Invertir colores (negativo es bueno, p. ej. costes)" })]));
                if (c.display === "bar") {
                    const ll = el("input", { type: "number", min: "0", max: "90", value: String(c.labelLeft) });
                    ll.addEventListener("input", () => { c.labelLeft = Math.max(0, Math.min(90, Number(ll.value) || 0)); touch(); });
                    const lr = el("input", { type: "number", min: "0", max: "90", value: String(c.labelRight) });
                    lr.addEventListener("input", () => { c.labelRight = Math.max(0, Math.min(90, Number(lr.value) || 0)); touch(); });
                    barBox.appendChild(el("div", { class: "calc-grid-2" }, [field("Espacio nº negativo (%)", ll), field("Espacio nº positivo (%)", lr)]));
                }
            };
            disp.addEventListener("change", () => { c.display = disp.value; touch(); renderBar(); });
            renderBar();

            const del = el("button", {
                type: "button", class: "calc-delete-btn",
                onclick: () => {
                    if (!confirm(`¿Quitar el cálculo «${c.name}»?`)) return;
                    this.calcs = this.calcs.filter(x => x.id !== c.id);
                    this.selectedId = this.calcs.length ? this.calcs[0].id : null;
                    touch(); this.renderList(); this.renderEditor();
                }
            }, ["Quitar cálculo"]);

            [
                el("div", { class: "field-options-section-title", text: "General" }),
                field("Nombre", name),
                field("Posición", posSel),
                el("div", { class: "field-options-section-title", text: "Fórmula" }),
                seg,
                el("span", { class: "calc-hint", text: modeHint }),
                ta,
                ops,
                el("div", { class: "calc-chips-title", text: c.mode === "position" ? "Posiciones" : "Miembros (pulsa para insertar)" }),
                chips,
                status,
                el("div", { class: "field-options-section-title", text: "Formato de número" }),
                el("div", { class: "calc-grid-2" }, [field("Formato", nf), field("Decimales", dec)]),
                el("div", { class: "field-options-section-title", text: "Visualización" }),
                field("Mostrar como", disp),
                barBox,
                el("span", { class: "calc-hint", text: "Formato y barras se aplican al actualizar si «Sobrescribir formatos» está activado en las propiedades del informe." }),
                del
            ].forEach(n => box.appendChild(n));
            updateStatus();
        },

        excelExample(ast) {
            const cols = ["C", "D", "E", "F", "G", "H"];
            const f = window.DracoCalc.toExcel(ast, (pos) => (cols[pos - 1] || "C") + "6");
            return f ? f.replace(/^=IFERROR\((.*),""\)$/, "=$1") : "";
        },

        save() {
            const names = new Set();
            for (const c of this.calcs) {
                const nm = String(c.name || "").trim();
                if (!nm) { this.showError("Todos los cálculos necesitan un nombre."); this.selectedId = c.id; this.renderList(); this.renderEditor(); return; }
                if (names.has(nm.toLowerCase())) { this.showError(`Hay dos cálculos llamados «${nm}».`); return; }
                names.add(nm.toLowerCase());
                if (this.members && this.members.some(m => m.toLowerCase() === nm.toLowerCase())) {
                    this.showError(`«${nm}» ya es un miembro de ${this.attr.toLowerCase()}: usa otro nombre.`); return;
                }
                const a = window.DracoCalc.analyze(c.expr, c.mode);
                if (!a.ok) { this.showError(`«${nm}»: ${a.error}`); this.selectedId = c.id; this.renderList(); this.renderEditor(); return; }
                c.name = nm;
            }

            const key = window.DracoCalc.key(this.dim, this.attr);
            if (this.calcs.length) {
                this.app.state.fieldOptions[key] = this.calcs.map(c => ({
                    id: c.id, name: c.name, mode: c.mode, expr: c.expr, position: c.position,
                    numberFormat: c.numberFormat, decimals: c.decimals, display: c.display,
                    colorNeg: c.colorNeg, colorPos: c.colorPos, invertColors: !!c.invertColors,
                    labelLeft: c.labelLeft, labelRight: c.labelRight
                }));
            } else {
                delete this.app.state.fieldOptions[key];
            }
            this.isDirty = false;
            this.close();
            if (typeof this.app.refreshCalcBadges === "function") this.app.refreshCalcBadges();
            this.app.scheduleAutoUpdate();
        },

        showError(msg) {
            if (this.errorEl) this.errorEl.textContent = msg;
        }
    };

    window.DracoCalcEditor = CalcEditor;
})();
