/* ==========================================================
 * Cálculos entre miembros — lado TASKPANE de la ventana
 * "Cálculos" (calcDialog.html + js/calcDialog.js).
 * ----------------------------------------------------------
 * Igual que FilterModal (filtros): abre una ventana de Office
 * independiente con displayDialogAsync y habla con ella por
 * mensajes. La ventana no tiene acceso a Excel, así que aquí se
 * resuelven los miembros de cada atributo (misma consulta que el
 * filtro) y se guarda cada alta/edición/borrado en
 * TaskPaneApp.state.fieldOptions bajo DracoCalc.key(dim, attr),
 * lanzando el autoguardado + refresco habitual.
 * ========================================================== */
(function () {
    "use strict";

    const CalcEditor = {
        dialog: null,

        /**
         * app        : TaskPaneApp
         * dimension  : dimensión
         * attribute  : atributo (desde el ƒx de un campo) o null (desde la dimensión)
         * attributes : atributos de la dimensión
         */
        open(app, dimension, attribute, attributes) {
            if (!window.DracoCalc) return;
            if (!app.currentReportId) {
                app.setAutoStatus && app.setAutoStatus("Crea o selecciona un informe antes de definir cálculos.");
                return;
            }
            if (this.dialog) { try { this.dialog.close(); } catch (e) { /* ya cerrado */ } this.dialog = null; }

            const attrs = attribute ? [attribute] : (attributes || []).slice();
            const calcs = {};
            attrs.forEach(a => { calcs[a] = window.DracoCalc.getCalcs(app.state.fieldOptions, dimension, a); });

            const url = new URL("calcDialog.html", window.location.href).href;
            Office.context.ui.displayDialogAsync(url, { height: 72, width: 60, displayInIframe: false }, (res) => {
                if (res.status === Office.AsyncResultStatus.Failed) {
                    console.error("Cálculos: displayDialogAsync ha fallado:", res.error && res.error.message);
                    return;
                }
                const dialog = res.value;
                this.dialog = dialog;
                const reply = (obj) => { try { dialog.messageChild(JSON.stringify(obj)); } catch (e) { console.error(e); } };

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                    let p;
                    try { p = JSON.parse(arg.message); } catch (e) { return; }

                    if (p.type === "ready") {
                        reply({ type: "init", dim: dimension, attr: attribute || null, attributes: attrs, calcs });
                        return;
                    }
                    if (p.type === "requestMembers") {
                        reply(Object.assign({ type: "members", attr: p.attr }, await this.loadMembers(app, dimension, p.attr)));
                        return;
                    }
                    if (p.type === "save") {
                        const key = window.DracoCalc.key(dimension, p.attr);
                        if (Array.isArray(p.calcs) && p.calcs.length) app.state.fieldOptions[key] = p.calcs;
                        else delete app.state.fieldOptions[key];
                        if (typeof app.refreshCalcBadges === "function") app.refreshCalcBadges();
                        app.scheduleAutoUpdate();
                        return;
                    }
                    if (p.type === "close") {
                        try { dialog.close(); } catch (e) { /* ya cerrado */ }
                        this.dialog = null;
                    }
                });
                dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { this.dialog = null; });
            });
        },

        async loadMembers(app, dimension, attribute) {
            let visible = null;
            const f = (app.state.filters || []).find(x =>
                String(x.dimension).toUpperCase() === String(dimension).toUpperCase()
                && String(x.realAttribute || x.name).toUpperCase() === String(attribute).toUpperCase());
            if (f && f.filter) {
                const vals = [];
                (f.filter.values || []).forEach(v => vals.push(String(v)));
                (f.filter.items || []).forEach(it => vals.push(String(it && typeof it === "object" ? (it.value !== undefined ? it.value : it.display) : it)));
                if (vals.length) visible = vals;
            }
            try {
                const sql = await window.ExcelService.buildFilterValuesSQL(dimension, attribute);
                if (!sql) throw new Error("No se encuentra el atributo en el modelo.");
                const json = await window.ExcelService.executeSQL(sql);
                const items = (typeof loadJsonTree === "function") ? loadJsonTree(json) : [];
                const members = Array.from(new Set(items.map(i => String(i.value)))).slice(0, 5000);
                return { members, visible, error: null };
            } catch (e) {
                return { members: [], visible, error: (e && e.message) || String(e) };
            }
        }
    };

    window.DracoCalcEditor = CalcEditor;
})();
