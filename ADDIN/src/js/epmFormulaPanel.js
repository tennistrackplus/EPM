/**
 * «Añadir fórmula» (cinta y clic derecho): línea calculada EPM_FORMULA en la
 * celda activa, con el mismo editor que «Fórmula» de Tablas Draco en
 * Planning (PLANNING/js/draco-table-designer.js, openFormulaPanel):
 * texto, fórmula, formato, barras e iconos.
 *
 * Diferencia con Planning: allí se hace clic en la rejilla para insertar
 * una referencia; aquí la rejilla es la propia hoja de Excel, así que se
 * selecciona en Excel y se pulsa «Insertar la selección» (el panel no
 * bloquea la hoja). La fórmula se escribe como si estuviera en la celda de
 * la línea: el add-in la copia a cada celda igual que copiar y pegar.
 *
 * Escribe =EPM_FORMULA(...) o =DRACO.EPM_FORMULA(...) según haya XLAM o no
 * (window.DracoFn, ver commands.js).
 */
(function () {
    "use strict";

    const FN_ES_EN = { "SI.ERROR": "IFERROR", "SI": "IF", "SUMA": "SUM", "PROMEDIO": "AVERAGE", "REDONDEAR": "ROUND" };

    function unq(s) {
        s = String(s == null ? "" : s).trim();
        if (s.length >= 2 && s.charAt(0) === '"' && s.charAt(s.length - 1) === '"') return s.slice(1, -1).replace(/""/g, '"');
        return s;
    }

    function el(tag, attrs, text) {
        const e = document.createElement(tag);
        Object.keys(attrs || {}).forEach(k => {
            if (k === "style") e.style.cssText = attrs[k];
            else if (k === "class") e.className = attrs[k];
            else e.setAttribute(k, attrs[k]);
        });
        if (text != null) e.textContent = text;
        return e;
    }

    function close() {
        const p = document.getElementById("epmFxPanel");
        if (p) p.remove();
    }

    async function writeFormula(target, f) {
        await Excel.run(async (context) => {
            const sh = context.workbook.worksheets.getItem(target.sheetName);
            if (window.DracoFn) window.DracoFn.suppress(true);
            try {
                sh.getRange(target.address).formulas = [[f]];
                await context.sync();
            } finally {
                if (window.DracoFn) window.DracoFn.suppress(false);
            }
        });
    }

    async function selectionRef(target) {
        let ref = "";
        await Excel.run(async (context) => {
            const r = context.workbook.getSelectedRange();
            r.load("address");
            await context.sync();
            const full = r.address, i = full.lastIndexOf("!");
            const sheet = i >= 0 ? full.slice(0, i).replace(/^'|'$/g, "").replace(/''/g, "'") : target.sheetName;
            const a = (i >= 0 ? full.slice(i + 1) : full).replace(/\$/g, "");
            ref = sheet === target.sheetName ? a : ("'" + sheet.replace(/'/g, "''") + "'!" + a);
        });
        return ref;
    }

    function open(target) {
        close();
        const fnName = window.DracoFn ? window.DracoFn.name("EPM_FORMULA") : "EPM_FORMULA";

        // Fórmula EPM_FORMULA que ya hubiera en la celda: se edita
        let a = ["", "", "", ""];
        const cur = typeof target.formula === "string" ? target.formula : "";
        const raw = (typeof splitDracoFunctionArgs === "function") ? splitDracoFunctionArgs(cur, "EPM_FORMULA") : null;
        if (raw) a = [unq(raw[0]), raw[1] || "", unq(raw[2]), unq(raw[3])];
        const editing = !!raw;

        const bars = String(a[3] || "").toLowerCase().split("|").map(x => x.trim());
        const barKind = bars.indexOf("solo barra") !== -1 ? "barOnly" : (bars.indexOf("barra") !== -1 || bars.indexOf("barras") !== -1 ? "bar" : "");
        const kv = (k, d) => { const x = bars.find(b => b.indexOf(k + "=") === 0); return x ? x.split("=")[1] : d; };
        const noAcc = t => t.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        const iconKind = bars.some(b => ["semaforo", "semaforos", "pelotitas"].indexOf(noAcc(b)) !== -1) ? "traffic"
            : (bars.some(b => ["senales", "senal"].indexOf(noAcc(b)) !== -1) ? "signs" : "");
        const cutsTxt = kv("cortes", "");
        const cutsV = cutsTxt ? cutsTxt.split(/[,;]/).map(x => x.trim()) : ["", ""];

        const panel = el("div", { id: "epmFxPanel", role: "dialog", "aria-label": "Añadir fórmula",
            style: "position:fixed;inset:0;z-index:5000;background:var(--color-surface,#fff);overflow:auto;padding:12px 14px;font-size:12px;display:flex;flex-direction:column;gap:8px" });

        const head = el("div", { style: "display:flex;align-items:center;gap:8px" });
        head.appendChild(el("strong", { style: "flex:1;font-size:14px" }, "ƒx " + (editing ? "Editar fórmula" : "Añadir fórmula")));
        const xBtn = el("button", { type: "button", class: "modal-close-btn", "aria-label": "Cerrar" }, "×");
        head.appendChild(xBtn);
        panel.appendChild(head);
        panel.appendChild(el("div", { style: "color:#64748b" }, "Celda: " + target.sheetName + "!" + target.address));

        const lbl = (text, input) => { const l = el("label", { style: "display:flex;flex-direction:column;gap:3px;font-weight:600" }); l.appendChild(el("span", null, text)); l.appendChild(input); return l; };
        const inp = (id, val, ph, extra) => el("input", Object.assign({ id, type: "text", value: val || "", placeholder: ph || "", style: "font-weight:400;padding:5px 6px" }, extra || {}));

        const txt = inp("epmFxText", a[0], "p. ej. Margen");
        const expr = inp("epmFxExpr", a[1], "p. ej. B7-B8", { autocomplete: "off", spellcheck: "false", style: "font-weight:400;padding:5px 6px;font-family:Consolas,monospace" });
        panel.appendChild(lbl("Texto", txt));
        panel.appendChild(lbl("Fórmula", expr));

        const insertAt = (text) => {
            const s = expr.selectionStart != null ? expr.selectionStart : expr.value.length;
            const e = expr.selectionEnd != null ? expr.selectionEnd : expr.value.length;
            expr.value = expr.value.slice(0, s) + text + expr.value.slice(e);
            const pos = s + text.length;
            expr.focus(); expr.setSelectionRange(pos, pos);
            validate();
        };

        const selRow = el("div", { style: "display:flex;gap:6px;align-items:center;flex-wrap:wrap" });
        const selBtn = el("button", { type: "button", class: "btn-secondary btn-sm" }, "⊕ Insertar la selección de Excel");
        selBtn.addEventListener("click", async () => {
            try { const ref = await selectionRef(target); if (ref) insertAt(ref); }
            catch (e) { err.textContent = "No se pudo leer la selección: " + (e.message || e); }
        });
        selRow.appendChild(selBtn);
        panel.appendChild(selRow);

        const ops = (list) => {
            const row = el("div", { style: "display:flex;gap:4px;flex-wrap:wrap" });
            list.forEach(([op, label, title]) => {
                const b = el("button", { type: "button", class: "btn-secondary btn-sm", style: "min-width:30px", title: title || "" }, label);
                b.addEventListener("click", () => insertAt(op));
                row.appendChild(b);
            });
            panel.appendChild(row);
        };
        ops([["+", "+"], ["-", "−"], ["*", "×"], ["/", "÷"], ["(", "("], [")", ")"]]);
        ops([["SI(", "SI(", "SI(condición; si se cumple; si no)"], [";", ";"], ["=", "="], ["<>", "≠"], ["<", "<"], [">", ">"], ["<=", "≤"], [">=", "≥"], ['""', '"texto"', "Texto entre comillas"]]);

        const hint = (t) => panel.appendChild(el("div", { style: "color:#64748b;line-height:1.35" }, t));
        hint("Selecciona una celda o un rango en Excel y pulsa «Insertar la selección». Escribe la fórmula como si estuviera en " + target.address + ": se copia a cada celda de la línea igual que copiar y pegar.");
        hint('Condiciones: SI(A6="HOLA";5;A6+10). Compara números o textos con = <> < > <= >=.');

        // Formato y barras
        const row2 = () => { const r = el("div", { style: "display:flex;gap:8px;flex-wrap:wrap;align-items:flex-end" }); panel.appendChild(r); return r; };
        const r1 = row2();
        const fmt = inp("epmFxFmt", a[2], "(opcional)", { list: "epmFxFmts" });
        const dl = el("datalist", { id: "epmFxFmts" });
        ["#,##0", "#,##0.00", "0%", "0.0%", "0.00%"].forEach(v => dl.appendChild(el("option", { value: v })));
        r1.appendChild(lbl("Formato", fmt)); r1.appendChild(dl);
        const sel = (id, opts, val) => { const s = el("select", { id, style: "padding:4px" }); opts.forEach(([v, t]) => { const o = el("option", { value: v }, t); if (v === val) o.selected = true; s.appendChild(o); }); return s; };
        const bar = sel("epmFxBar", [["", "Sin barras"], ["bar", "Número y barra"], ["barOnly", "Solo barra"]], barKind);
        r1.appendChild(lbl("Barras", bar));

        const barOpts = row2();
        const chk = (id, text, on) => { const l = el("label", { style: "display:flex;gap:4px;align-items:center" }); const c = el("input", { id, type: "checkbox" }); c.checked = !!on; l.appendChild(c); l.appendChild(el("span", null, text)); return { l, c }; };
        const inv = chk("epmFxInv", "Invertir colores", bars.indexOf("invertir") !== -1);
        const izq = el("input", { type: "number", min: "0", max: "90", value: kv("izq", "40"), style: "width:64px" });
        const dcha = el("input", { type: "number", min: "0", max: "90", value: kv("dcha", "25"), style: "width:64px" });
        barOpts.appendChild(inv.l); barOpts.appendChild(lbl("Esp. izq. %", izq)); barOpts.appendChild(lbl("Esp. dcha. %", dcha));

        const r3 = row2();
        const icon = sel("epmFxIcon", [["", "Sin iconos"], ["traffic", "Semáforo (pelotitas)"], ["signs", "Señales"]], iconKind);
        const cuts = sel("epmFxCuts", [["", "Automáticos (33 % / 67 %)"], ["fixed", "Fijos"]], cutsTxt ? "fixed" : "");
        r3.appendChild(lbl("Iconos", icon)); r3.appendChild(lbl("Cortes", cuts));
        const iconOpts = row2();
        const c1 = el("input", { type: "number", step: "any", value: cutsV[0] || "", placeholder: "p. ej. 0", style: "width:80px" });
        const c2 = el("input", { type: "number", step: "any", value: cutsV[1] || "", placeholder: "p. ej. 10", style: "width:80px" });
        const c1L = lbl("Ámbar desde", c1), c2L = lbl("Verde desde", c2);
        iconOpts.appendChild(c1L); iconOpts.appendChild(c2L);
        const iconOpts2 = row2();
        const iInv = chk("epmFxIInv", "Invertir (verde para valores bajos)", bars.indexOf("invertir-iconos") !== -1);
        const iOnly = chk("epmFxIOnly", "Solo el icono", bars.indexOf("solo-icono") !== -1);
        iconOpts2.appendChild(iInv.l); iconOpts2.appendChild(iOnly.l);

        const prev = el("code", { style: "display:block;background:#f1f5f9;padding:6px;border-radius:4px;word-break:break-all;font-size:11px" });
        panel.appendChild(prev);
        const foot = el("div", { style: "display:flex;gap:6px;align-items:center;margin-top:4px" });
        const err = el("span", { style: "flex:1;color:#dc2626" });
        const cancel = el("button", { type: "button", class: "btn-secondary btn-sm" }, "Cancelar");
        const ok = el("button", { type: "button", class: "btn-primary btn-sm" }, editing ? "Guardar" : "Insertar");
        foot.appendChild(err); foot.appendChild(cancel); foot.appendChild(ok);
        panel.appendChild(foot);

        const syncBar = () => {
            barOpts.hidden = !bar.value; izq.disabled = dcha.disabled = bar.value !== "bar";
            const ic = icon.value;
            cuts.disabled = !ic; iconOpts2.hidden = !ic;
            iconOpts.hidden = !ic || cuts.value !== "fixed";
            const sig = ic === "signs";
            c1L.firstChild.textContent = sig ? "Triángulo desde" : "Ámbar desde";
            c2L.firstChild.textContent = sig ? "Círculo verde desde" : "Verde desde";
        };

        const build = () => {
            const q = t => String(t).replace(/"/g, '""');
            const text = txt.value.trim(), ex = expr.value.trim(), fm = fmt.value.trim();
            let barSpec = "";
            if (bar.value) {
                const parts = [bar.value === "barOnly" ? "solo barra" : "barra"];
                if (inv.c.checked) parts.push("invertir");
                if (bar.value === "bar") { parts.push("izq=" + (Number(izq.value) || 0)); parts.push("dcha=" + (Number(dcha.value) || 0)); }
                barSpec = parts.join("|");
            }
            if (icon.value) {
                const ip = [icon.value === "signs" ? "señales" : "semaforo"];
                if (cuts.value === "fixed" && c1.value !== "" && c2.value !== "") ip.push("cortes=" + Number(c1.value) + "," + Number(c2.value));
                if (iInv.c.checked) ip.push("invertir-iconos");
                if (iOnly.c.checked) ip.push("solo-icono");
                barSpec = barSpec ? barSpec + "|" + ip.join("|") : ip.join("|");
            }
            // ";" fuera de comillas -> "," y SI/SUMA... -> IF/SUM... (Excel guarda las fórmulas en inglés)
            const exEn = ex.replace(/("(?:[^"]|"")*")|;|\b(SI\.ERROR|SI|SUMA|PROMEDIO|REDONDEAR)\s*\(/gi,
                (m, qs, fn) => qs ? qs : (fn ? FN_ES_EN[fn.toUpperCase()] + "(" : ","));
            const args = ['"' + q(text) + '"', exEn];
            if (fm || barSpec) args.push('"' + q(fm) + '"');
            if (barSpec) args.push('"' + barSpec + '"');
            return { text, expr: exEn, f: "=" + fnName + "(" + args.join(",") + ")" };
        };

        const validate = () => {
            const b = build();
            prev.textContent = b.f;
            let e = "";
            if (!b.text) e = "Pon un texto.";
            else if (!b.expr) e = "Escribe o compón la fórmula.";
            else {
                let d = 0;
                for (const ch of b.expr.replace(/"[^"]*"/g, "")) { if (ch === "(") d++; if (ch === ")") d--; if (d < 0) break; }
                if (d !== 0) e = "Los paréntesis no cuadran.";
                else if ((b.expr.match(/"/g) || []).length % 2) e = "Hay unas comillas sin cerrar.";
            }
            err.textContent = e;
            ok.disabled = !!e;
            return e ? null : b;
        };

        panel.addEventListener("input", validate);
        [bar, icon, cuts].forEach(s => s.addEventListener("change", () => { syncBar(); validate(); }));
        xBtn.addEventListener("click", close);
        cancel.addEventListener("click", close);
        ok.addEventListener("click", async () => {
            const b = validate();
            if (!b) return;
            ok.disabled = true;
            try {
                await writeFormula(target, b.f);
                close();
                if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.setAutoStatus) {
                    TaskPaneApp.setAutoStatus("Fórmula escrita en " + target.address + ". Actualiza el informe para calcular la línea.");
                }
            } catch (e) {
                // Excel rechaza la fórmula (referencia o función no válida)
                err.textContent = "Excel no acepta la fórmula: " + (e.message || e);
                ok.disabled = false;
            }
        });

        document.body.appendChild(panel);
        syncBar(); validate();
        (a[0] ? expr : txt).focus();
    }

    function attach() {
        if (typeof TaskPaneApp === "undefined") return false;
        TaskPaneApp.openEpmFormulaPanel = open;
        return true;
    }
    if (!attach()) document.addEventListener("DOMContentLoaded", attach);
})();
