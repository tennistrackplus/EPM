/**
 * DracoTableDesigner — diseñador de una "Tabla Draco" (informe tipo tabla dinámica del addin de Excel).
 *
 * Es un pseudo-Excel SOLO PARA DISEÑAR UN INFORME: a la izquierda, una rejilla de solo lectura que muestra lo que
 * pinta el informe (sin herramientas de editor de Excel: nada de fuentes, colores, bordes, combinar ni insertar
 * filas); a la derecha, el taskpane del addin (PLANNING/tablas-draco/, generado desde ADDIN/src por
 * tools/sync-addin.py) con TODAS sus opciones de diseño: modelo semántico, campos en Filas/Columnas/Filtros,
 * medidas, jerarquías, opciones de campo y propiedades del informe. Arriba, los botones de la cinta del addin que
 * tienen sentido aquí: Actualizar, Añadir filtro y Reconocimiento de miembros.
 *
 * Al guardar se genera el SQL del informe y se guarda junto a su definición en TABLAS_DRACO (ver draco-tables.js).
 */
const DracoTableDesigner = {
    TASKPANE_URL: "tablas-draco/taskpane.html?v=20261022w",
    MAX_ROWS: 600,
    MAX_COLS: 80,
    ctx: null,

    /** Abre el diseñador (ref = tabla existente o null para una nueva). Devuelve { id, nombre } al guardar, o null. */
    /**
     * opts.kpi: diseñador de KPI (bloque «KPI Draco»): misma pantalla, pero el informe debe tener UNA medida y solo
     * filtros; al guardar se guardan además sus piezas (def.kpi) para calcular todos los KPIs en una consulta.
     */
    async open(project, ref, opts) {
        if (this.ctx) throw Object.assign(new Error("El diseñador ya está abierto."), { code: "BUSY" });
        let cubes = [], row = null, def = null;
        // KPI del informe: vive DENTRO del informe (no en la base de datos). Se edita su definición en memoria y, al
        // guardar, se devuelve { kpi: { nombre, cubo, def, parts, vars } } para que el generador lo guarde en el informe.
        const inlineKpi = !!(opts && opts.kpi && !ref);
        try {
            cubes = await DracoTables.cubes(project.PROYECTO_ID);
            if (inlineKpi) { def = opts.def || null; row = { NOMBRE: opts.nombre || "", DESCRIPCION: "" }; }
            else if (ref) { row = await DracoTables.get(project.PROYECTO_ID, ref); def = JSON.parse(row.DEFINICION_JSON || "{}"); }
        } catch (err) {
            UI.toast("No se pudo abrir el diseñador: " + err.message, "error");
            throw err;
        }
        return new Promise((resolve) => {
            this.ctx = { project, ref, row, cubes, resolve, saved: null, initialSig: null, ready: false, raf: 0, sel: { r: 0, c: 0 },
                         kpi: !!((opts && opts.kpi) || (def && def.kpi)), inlineKpi, hasDef: !!def,
                         owner: (opts && opts.owner) || null, autoName: (opts && opts.autoName) || "" };   // tabla de un informe: nombre opcional
            this.buildModal();
            if (this.ctx.kpi) this.banner("KPI: arrastra UNA medida a «Filtros», junto con los filtros que quieras. El valor aparece a la derecha.", "info");
            if (!cubes.length) this.banner("Este proyecto no tiene cubos, y el informe se diseña sobre uno. Crea un cubo en «Cubos» y vuelve a abrir el diseñador.", "warn");
            this.startTaskpane(def);
        });
    },

    // ---------------------------------------------------------------------
    // Ventana
    // ---------------------------------------------------------------------
    buildModal() {
        const c = this.ctx, E = UI.escapeHtml;
        const overlay = document.createElement("div");
        overlay.className = "modal-overlay visible dtd-overlay" + (c.kpi ? " dtd-kpi" : "");
        overlay.id = "dtdOverlay";
        overlay.innerHTML = `
            <div class="dtd-modal" role="dialog" aria-label="Diseñador de tabla Draco">
                <div class="dtd-head">
                    <strong class="dtd-title">${c.kpi ? "KPI Draco · diseño" : "Tabla Draco · diseño del informe"}</strong>
                    <input id="dtdName" class="dtd-name" type="text" maxlength="${DracoTables.MAX_NAME}" placeholder="${c.kpi ? "Nombre del KPI" : (c.owner && !c.ref ? "Nombre (opcional)" : "Nombre de la tabla")}" value="${E(c.row ? c.row.NOMBRE : "")}">
                    <input id="dtdDesc" class="dtd-desc" type="text" maxlength="${DracoTables.MAX_DESC}" placeholder="Descripción (opcional)" value="${E(c.row ? (c.row.DESCRIPCION || "") : "")}">
                    <span class="dtd-spacer"></span>
                    <span id="dtdStatus" class="dtd-status" role="status"></span>
                    <button type="button" class="btn btn-secondary" id="dtdClose">Cerrar</button>
                    <button type="button" class="btn btn-primary" id="dtdSave" disabled>Guardar</button>
                </div>
                <div class="dtd-ribbon">
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdRefresh" title="Refrescar el informe: vuelve a leer el diseño y las líneas del eje Estático, genera el SQL y lo ejecuta">⟳ Refrescar</button>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdFilter" title="Añadir un filtro al informe">▽ Filtro</button>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdRecog" title="Reconocimiento de miembros">🔎 Reconocimiento</button>
                    ${c.kpi ? "" : `<span class="dtd-sep" aria-hidden="true"></span>
                    <span class="dtd-linkbar" id="dtdLinkBar" hidden><span>Diseñando</span><select id="dtdLinkSel" title="Informe del grupo que estás diseñando: el principal o uno de sus vinculados"></select><button type="button" class="btn btn-secondary btn-sm" id="dtdLinkDel" title="Quitar este informe vinculado" hidden>Quitar</button></span>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdLink" title="Añadir un informe vinculado: comparte las filas (se pinta a la derecha) o las columnas (se pinta debajo) con este informe, con sus propias medidas y filtros">🔗 Vincular informe</button>`}
                    <span class="dtd-sep" aria-hidden="true"></span>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdMember" title="Poner miembros (EPM_VALUE) en la celda seleccionada. También con doble clic.">＋ Miembro</button>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdFormula" title="Línea calculada (EPM_FORMULA) en la celda seleccionada"><i class="dtd-fx">ƒx</i> Fórmula</button>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdClear" title="Borrar el contenido de las celdas seleccionadas (Supr)">✕ Borrar</button>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdVars" title="Variables: un nombre asociado a un valor de un campo, para usar en las líneas del eje Estático ($NOMBRE$)"><i class="dtd-fx">𝑥</i> Variables</button>
                    <span class="dtd-hint">Ejes Estáticos: selecciona celdas y usa «Miembro», «Fórmula» o Supr; después «Refrescar».</span>
                </div>
                <div class="dtd-fbar">
                    <span class="dtd-fbar-addr" id="dtdAddr">A1</span>
                    <i class="dtd-fx">ƒx</i>
                    <input id="dtdFbar" class="dtd-fbar-input" type="text" readonly aria-label="Contenido de la celda seleccionada">
                </div>
                <div class="dtd-banner" id="dtdBanner" role="alert" hidden></div>
                <div class="dtd-body">
                    <div class="dtd-grid-wrap" id="dtdGridWrap"><div class="dtd-loading"><span class="spinner"></span> Preparando el diseñador…</div></div>
                    <div class="dtd-resizer" id="dtdResizer" title="Arrastra para cambiar el ancho del panel"></div>
                    <div class="dtd-panel" id="dtdPanel"></div>
                    <aside class="dtd-kpival" id="dtdKpiVal" aria-live="polite">
                        <div class="dtd-kpival-l">Valor del KPI</div>
                        <div class="dtd-kpival-v" id="dtdKpiV">—</div>
                        <div class="dtd-kpival-m" id="dtdKpiM"></div>
                        <p class="dtd-kpival-h">Pon una sola medida en «Filtros» y, si quieres, filtros por dimensión. Pulsa «Refrescar» para recalcular.</p>
                    </aside>
                </div>
            </div>`;
        document.body.appendChild(overlay);
        c.overlay = overlay;
        overlay.querySelector("#dtdClose").addEventListener("click", () => this.requestClose());
        overlay.querySelector("#dtdSave").addEventListener("click", () => this.save());
        overlay.querySelector("#dtdRefresh").addEventListener("click", () => this.refreshReport());
        overlay.querySelector("#dtdVars").addEventListener("click", () => this.callCommand("abrirVariables"));
        overlay.querySelector("#dtdFilter").addEventListener("click", () => this.callCommand("abrirAnadirFiltro"));
        if (!c.kpi) {
            overlay.querySelector("#dtdLink").addEventListener("click", () => this.openLinkDialog());
            overlay.querySelector("#dtdLinkSel").addEventListener("change", (e) => this.switchGroupReport(Number(e.target.value)));
            overlay.querySelector("#dtdLinkDel").addEventListener("click", () => this.removeLinked());
        }
        overlay.querySelector("#dtdRecog").addEventListener("click", () => this.callCommand("toggleMemberRecognition"));
        overlay.querySelector("#dtdMember").addEventListener("click", () => { const s = this.selTopLeft(); this.openMemberDialog(s.r, s.c); });
        overlay.querySelector("#dtdFormula").addEventListener("click", () => { const s = this.selTopLeft(); this.openFormulaPanel(s.r, s.c); });
        overlay.querySelector("#dtdClear").addEventListener("click", () => this.clearSelection());
        this.wireResizer();
        c.keyHandler = (e) => this.onKey(e);
        document.addEventListener("keydown", c.keyHandler);
        this.wireGridSelection(overlay.querySelector("#dtdGridWrap"));
    },

    status(text, kind) {
        const el = this.ctx && this.ctx.overlay.querySelector("#dtdStatus");
        if (!el) return;
        el.textContent = text || "";
        el.className = "dtd-status" + (kind ? " " + kind : "");
    },

    /** Aviso persistente (no desaparece solo como un toast): el motivo por el que algo no funciona. */
    banner(text, kind) {
        const el = this.ctx && this.ctx.overlay.querySelector("#dtdBanner");
        if (!el) return;
        el.hidden = !text;
        el.textContent = text || "";
        el.className = "dtd-banner" + (kind ? " " + kind : "");
    },

    wireResizer() {
        const c = this.ctx, rz = c.overlay.querySelector("#dtdResizer"), panel = c.overlay.querySelector("#dtdPanel");
        rz.addEventListener("mousedown", (e) => {
            e.preventDefault();
            const startX = e.clientX, startW = panel.getBoundingClientRect().width;
            rz.classList.add("dragging"); document.body.style.userSelect = "none";
            const move = (ev) => { panel.style.width = Math.max(300, Math.min(window.innerWidth - 360, startW - (ev.clientX - startX))) + "px"; };
            const up = () => { rz.classList.remove("dragging"); document.body.style.userSelect = ""; document.removeEventListener("mousemove", move); document.removeEventListener("mouseup", up); };
            document.addEventListener("mousemove", move); document.addEventListener("mouseup", up);
        });
    },

    // ---------------------------------------------------------------------
    // Taskpane del addin (en un iframe con su "Excel falso")
    // ---------------------------------------------------------------------
    startTaskpane(def) {
        const c = this.ctx;
        const host = {
            state: DracoTables.stateFromDefinition(def || {}),
            selection: { r1: 0, c1: 0, r2: 0, c2: 0 },
            cubes: c.cubes.map(x => ({ name: x.name })),
            resolveModel: async (name) => {
                const cube = c.cubes.find(x => x.name === name);
                if (!cube) return null;
                this.status("Cargando el modelo «" + name + "»…");
                try { const m = await DracoTables.resolveModel(c.project, cube); this.banner(""); return m; }
                catch (err) { this.banner(err.message, "bad"); UI.toast("No se pudo cargar el modelo del cubo «" + name + "».", "error"); return null; }
                finally { this.status(""); }
            },
            markDirty() {},
            renderGrid: () => this.scheduleRender()
        };
        c.host = host;
        const frame = document.createElement("iframe");
        frame.className = "dtd-taskpane";
        frame.title = "Diseño del informe";
        frame.__dracoHost = host;
        c.frame = frame;
        frame.addEventListener("load", () => this.onTaskpaneLoaded());
        frame.src = this.TASKPANE_URL;
        c.overlay.querySelector("#dtdPanel").appendChild(frame);
    },

    win() { return this.ctx && this.ctx.frame && this.ctx.frame.contentWindow; },

    async onTaskpaneLoaded() {
        const c = this.ctx;
        if (!c || c.ready) return;
        try {
            const w = this.win();
            // espera a que el taskpane termine de arrancar (modelos registrados + init)
            await Promise.race([w.__wteInitDonePromise || Promise.resolve(), new Promise(r => setTimeout(r, 15000))]);
            // Una tabla Draco = UN informe: si no hay ninguno (tabla nueva), se crea
            if (!w.ReportStore.getActiveReport()) await w.TaskPaneApp.addReport();
            if (c.kpi) {
                // KPI: solo campos y «Filtros» (la medida va en Filtros); sin Filas, Columnas ni opciones de campo
                const st = w.document.createElement("style");
                st.textContent = '[data-zone="rows"],[data-zone="columns"]{display:none!important}';
                w.document.head.appendChild(st);
            }
            // Una tabla ya guardada guarda su DISEÑO, no su resultado: se actualiza al abrir para que se vea
            if (c.ref || (c.inlineKpi && c.hasDef)) {
                this.status("Cargando datos…");
                try { await w.DracoEngineApi.refresh(undefined, null); }
                catch (err) { this.status("No se pudo actualizar: " + err.message, "bad"); }
            }
            c.ready = true;
            this.refreshLinkBar();
            c.initialSig = this.signature();
            c.overlay.querySelector("#dtdSave").disabled = false;
            c.overlay.querySelector("#dtdGridWrap").querySelector(".dtd-loading") && this.scheduleRender();
            if (!/^No se pudo actualizar/.test(c.overlay.querySelector("#dtdStatus").textContent)) this.status("");
        } catch (err) {
            console.error("DracoTableDesigner: error al preparar el taskpane", err);
            this.status("No se pudo preparar el diseñador: " + err.message, "bad");
        }
    },

    /** Huella del diseño (para saber si hay cambios sin guardar). */
    signature() {
        const s = (this.ctx && this.ctx.host && this.ctx.host.state.taskpaneSettings) || {};
        const st = (this.ctx && this.ctx.host) ? JSON.stringify(DracoTables.staticCells(this.ctx.host)) : "";
        return String(s.epm_reports || "") + "|" + String(s.epm_activeSemanticModel || "") + "|" + st;
    },

    /**
     * Refresca el informe por la misma vía que al abrir el diseñador (DracoEngineApi.refresh): relee el diseño y
     * las líneas del eje Estático (miembros, fórmulas, variables), genera el SQL y lo ejecuta. A diferencia del
     * comando "actualizar" del addin (que escribe los errores en una hoja que aquí no se ve), los errores se
     * muestran en el aviso del diseñador.
     */
    async refreshReport() {
        const c = this.ctx, w = this.win();
        if (!c || !c.ready || !w || !w.DracoEngineApi) { UI.toast("El diseñador todavía se está cargando.", "info"); return; }
        const btn = c.overlay.querySelector("#dtdRefresh");
        btn.disabled = true;
        this.status("Refrescando…");
        try {
            await w.DracoEngineApi.refresh(undefined, null);
            this.banner(c.kpi ? "KPI: arrastra UNA medida a «Filtros», junto con los filtros que quieras. El valor aparece a la derecha." : "", c.kpi ? "info" : undefined);
            this.status("");
        } catch (err) {
            console.error("DracoTableDesigner: refresco", err);
            this.banner("No se pudo refrescar: " + (err.message || err), "bad");
            this.status("");
        } finally {
            btn.disabled = false;
            this.scheduleRender();
        }
    },

    /** Llama a una función de comando del addin (commands.js) con el "event" con .completed() que espera. */
    async callCommand(fnName) {
        const w = this.win();
        if (!w || typeof w[fnName] !== "function") { UI.toast("El diseñador todavía se está cargando.", "info"); return; }
        try { await w[fnName]({ completed() {} }); } catch (err) { UI.toast("Error: " + err.message, "error"); }
        this.scheduleRender();
    },

    // ---------------------------------------------------------------------
    // Rejilla de solo lectura: lo que pinta el informe
    // ---------------------------------------------------------------------
    scheduleRender() {
        const c = this.ctx;
        if (!c || c.raf) return;
        c.raf = requestAnimationFrame(() => { c.raf = 0; this.renderGrid(); });
    },

    colLabel(i) { let s = "", n = i + 1; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; },

    renderGrid() {
        const c = this.ctx;
        if (!c || !c.host) return;
        const wrap = c.overlay.querySelector("#dtdGridWrap");
        const cells = c.host.state.cells || {};
        let maxR = 0, maxC = 0;
        Object.keys(cells).forEach(k => {
            const o = cells[k]; if (!o || ((o.v === undefined || o.v === "") && !o.bg)) return;
            const [r, cc] = k.split("_").map(Number);
            if (r > maxR) maxR = r; if (cc > maxC) maxC = cc;
        });
        const rows = Math.min(this.MAX_ROWS, Math.max(30, maxR + 6)), cols = Math.min(this.MAX_COLS, Math.max(14, maxC + 4));
        const E = UI.escapeHtml;
        const covered = new Set(), anchors = new Map();
        (c.host.state.merges || []).forEach(m => {
            anchors.set(m.r + "_" + m.c, m);
            for (let r = m.r; r < m.r + (m.rowSpan || 1); r++) for (let cc = m.c; cc < m.c + (m.colSpan || 1); cc++) if (r !== m.r || cc !== m.c) covered.add(r + "_" + cc);
        });
        let h = `<table class="dtd-grid"><thead><tr><th class="dtd-corner"></th>`;
        for (let cc = 0; cc < cols; cc++) h += `<th class="dtd-colh">${this.colLabel(cc)}</th>`;
        h += `</tr></thead><tbody>`;
        for (let r = 0; r < rows; r++) {
            h += `<tr><th class="dtd-rowh">${r + 1}</th>`;
            for (let cc = 0; cc < cols; cc++) {
                const k = r + "_" + cc;
                if (covered.has(k)) continue;
                const o = cells[k] || {}, m = anchors.get(k);
                const t = DracoTables.formatValue(o.v, o.nf);
                const text = t !== null ? t : (o.v === undefined ? "" : String(o.v));
                const num = t !== null || (o.v !== undefined && o.v !== "" && /^-?\d+(\.\d+)?$/.test(String(o.v)));
                const st = [];
                if (o.bg && o.bg !== "#FFFFFF") st.push("background:" + E(o.bg));
                if (o.col) st.push("color:" + E(o.col));
                if (o.b) st.push("font-weight:700");
                if (o.i) st.push("font-style:italic");
                if (o.u) st.push("text-decoration:underline");
                const al = o.al || (num ? "right" : "left");
                st.push("text-align:" + al);
                if (o.ind) st.push("padding-left:" + (6 + o.ind * 14) + "px");
                ["bt:top", "br:right", "bb:bottom", "bl:left"].forEach(p => { const [key, side] = p.split(":"); const b = o[key]; if (b && b.style !== "none") st.push("border-" + side + ":1px solid " + E(b.color || "#94A3B8")); });
                if (o.db) {                                  // barra de datos (EPM_FORMULA con "barra")
                    st.push(DracoTables.barCss(o.db));
                    if (!o.al) st.push("text-align:" + (Number(o.v) < 0 ? "left" : "right"));
                    if (o.db.only) st.push("color:transparent");
                }
                const cls = ["dtd-cell"];
                if (c.sel.r === r && c.sel.c === cc) cls.push("dtd-sel");
                if (this.inSel(r, cc)) cls.push("dtd-insel");
                const isEpm = typeof o.f === "string" && /EPM_(VALUE|FORMULA)\s*\(/i.test(o.f);
                if (isEpm) cls.push(/EPM_FORMULA/i.test(o.f) ? "dtd-epmf" : "dtd-epm");
                const tip = o.f ? ` title="${E(o.f)}"` : "";
                h += `<td class="${cls.join(" ")}" data-r="${r}" data-c="${cc}"${m ? ` rowspan="${m.rowSpan || 1}" colspan="${m.colSpan || 1}"` : ""}${tip} style="${st.join(";")}">${o.ic ? DracoTables.iconHtml(o.ic) : ""}${o.ic && o.ic.only ? "" : E(text)}</td>`;
            }
            h += `</tr>`;
        }
        h += `</tbody></table>`;
        const top = wrap.scrollTop, left = wrap.scrollLeft;
        wrap.innerHTML = h;
        wrap.scrollTop = top; wrap.scrollLeft = left;
        this.updateFormulaBar();
        if (c.kpi) this.updateKpiValue();
    },

    /** Un clic en una celda hace lo que haría en Excel: selecciona y, si es un +/- de jerarquía, la expande o contrae. */
    async onGridClick(e) {
        const td = e.target.closest && e.target.closest("td.dtd-cell");
        const c = this.ctx;
        if (!td || !c || !c.ready) return;
        const r = parseInt(td.dataset.r, 10), cc = parseInt(td.dataset.c, 10);
        c.sel = { r, c: cc };
        c.host.selection = { r1: r, c1: cc, r2: r, c2: cc };
        const w = this.win();
        try {
            if (w && typeof w.__fireExcelEvent === "function") {
                await w.__fireExcelEvent("onSelectionChanged", r, cc);
                await w.__fireExcelEvent("onSingleClicked", r, cc);
            }
        } catch (err) { console.error(err); }
        this.scheduleRender();
    },

    // ---------------------------------------------------------------------
    // Edición del eje Estático: selección, Supr, «Miembro» (EPM_VALUE) y
    // «Fórmula» (EPM_FORMULA). Las celdas se escriben con el mismo "Excel
    // falso" del taskpane (host-bridge), así que el motor del addin las lee
    // exactamente igual que en Excel.
    // ---------------------------------------------------------------------
    addr(r, c) { return this.colLabel(c) + (r + 1); },

    /** KPI: muestra el valor (la primera celda numérica que no es cabecera de lo que pinta el informe). */
    updateKpiValue() {
        const c = this.ctx, el = c && c.overlay.querySelector("#dtdKpiV");
        if (!el) return;
        const cells = (c.host && c.host.state.cells) || {};
        const keys = Object.keys(cells).map(k => k.split("_").map(Number)).sort((a, b) => a[0] - b[0] || a[1] - b[1]);
        let val = null, label = "";
        for (const [r, cc] of keys) {
            const o = cells[r + "_" + cc];
            if (o && !o.bg && o.v !== undefined && String(o.v).trim() !== "" && isFinite(Number(o.v)) && !(typeof o.f === "string" && /EPM_/i.test(o.f))) {
                val = DracoTables.formatValue(o.v, o.nf) || Number(o.v).toLocaleString("es-ES", { maximumFractionDigits: 2 });
                const above = cells[(r - 1) + "_" + cc]; label = above ? String(above.v || "") : "";
                break;
            }
        }
        el.textContent = val === null ? "—" : val;
        c.overlay.querySelector("#dtdKpiM").textContent = val === null ? "Todavía no hay valor." : (label ? "Medida: " + label : "");
    },

    selRect() {
        const s = this.ctx.selRange || { r1: this.ctx.sel.r, c1: this.ctx.sel.c, r2: this.ctx.sel.r, c2: this.ctx.sel.c };
        return { r1: Math.min(s.r1, s.r2), c1: Math.min(s.c1, s.c2), r2: Math.max(s.r1, s.r2), c2: Math.max(s.c1, s.c2) };
    },
    selTopLeft() { const s = this.selRect(); return { r: s.r1, c: s.c1 }; },
    inSel(r, c) {
        if (!this.ctx || !this.ctx.selRange) return false;
        const s = this.selRect();
        return (s.r2 > s.r1 || s.c2 > s.c1) && r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2;
    },

    /** Pinta la selección sin reconstruir la rejilla (para arrastrar con fluidez). */
    paintSelection() {
        const wrap = this.ctx.overlay.querySelector("#dtdGridWrap");
        wrap.querySelectorAll("td.dtd-sel, td.dtd-insel").forEach(td => td.classList.remove("dtd-sel", "dtd-insel"));
        const s = this.selRect(), multi = s.r2 > s.r1 || s.c2 > s.c1;
        wrap.querySelectorAll("td.dtd-cell").forEach(td => {
            const r = +td.dataset.r, c = +td.dataset.c;
            if (r === this.ctx.sel.r && c === this.ctx.sel.c) td.classList.add("dtd-sel");
            else if (multi && r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2) td.classList.add("dtd-insel");
        });
        this.updateFormulaBar();
    },

    updateFormulaBar() {
        const c = this.ctx;
        if (!c || !c.overlay) return;
        const s = this.selRect();
        const a = c.overlay.querySelector("#dtdAddr"), f = c.overlay.querySelector("#dtdFbar");
        if (!a || !f) return;
        a.textContent = (s.r2 > s.r1 || s.c2 > s.c1) ? this.addr(s.r1, s.c1) + ":" + this.addr(s.r2, s.c2) : this.addr(s.r1, s.c1);
        const o = (c.host && c.host.state.cells[c.sel.r + "_" + c.sel.c]) || {};
        f.value = o.f !== undefined ? o.f : (o.v === undefined ? "" : String(o.v));
    },

    wireGridSelection(wrap) {
        const c = this.ctx;
        const cellOf = (e) => { const td = e.target.closest && e.target.closest("td.dtd-cell"); return td ? { r: +td.dataset.r, c: +td.dataset.c } : null; };
        wrap.addEventListener("mousedown", (e) => {
            const p = cellOf(e);
            if (!p || e.button !== 0 || !c.ready) return;
            e.preventDefault();
            if (c.fx) { c.fxPick = { start: p, end: p }; return; }      // panel de fórmula: elegir referencias
            if (e.shiftKey && c.selRange) { c.selRange.r2 = p.r; c.selRange.c2 = p.c; }
            else { c.sel = { r: p.r, c: p.c }; c.selRange = { r1: p.r, c1: p.c, r2: p.r, c2: p.c }; }
            c.dragging = { moved: false, shift: e.shiftKey };
            this.paintSelection();
        });
        wrap.addEventListener("mouseover", (e) => {
            const p = cellOf(e);
            if (!p) return;
            if (c.fx && c.fxPick) { c.fxPick.end = p; this.highlightPick(); return; }
            if (!c.dragging) return;
            if (p.r !== c.selRange.r2 || p.c !== c.selRange.c2) { c.selRange.r2 = p.r; c.selRange.c2 = p.c; c.dragging.moved = true; this.paintSelection(); }
        });
        wrap.addEventListener("dblclick", (e) => {
            const p = cellOf(e);
            if (!p || !c.ready || c.fx) return;
            const o = c.host.state.cells[p.r + "_" + p.c] || {};
            if (typeof o.f === "string" && /EPM_FORMULA\s*\(/i.test(o.f)) this.openFormulaPanel(p.r, p.c);
            else this.openMemberDialog(p.r, p.c);
        });
        c.mouseUpHandler = (e) => {
            if (c.fx && c.fxPick) {
                const { start, end } = c.fxPick;
                c.fxPick = null;
                this.highlightPick(true);
                const r1 = Math.min(start.r, end.r), c1 = Math.min(start.c, end.c), r2 = Math.max(start.r, end.r), c2 = Math.max(start.c, end.c);
                this.fxInsert(r1 === r2 && c1 === c2 ? this.addr(r1, c1) : this.addr(r1, c1) + ":" + this.addr(r2, c2));
                return;
            }
            if (!c.dragging) return;
            const d = c.dragging; c.dragging = null;
            const s = this.selRect();
            // un clic simple (sin arrastrar) hace lo que en Excel: eventos de selección / +/- de jerarquía
            if (!d.moved && !d.shift && s.r1 === s.r2 && s.c1 === s.c2) this.fireCellClick(s.r1, s.c1);
        };
        document.addEventListener("mouseup", c.mouseUpHandler);
    },

    highlightPick(clear) {
        const c = this.ctx, wrap = c.overlay.querySelector("#dtdGridWrap");
        wrap.querySelectorAll("td.dtd-pick").forEach(td => td.classList.remove("dtd-pick"));
        if (clear || !c.fxPick) return;
        const { start, end } = c.fxPick;
        const r1 = Math.min(start.r, end.r), c1 = Math.min(start.c, end.c), r2 = Math.max(start.r, end.r), c2 = Math.max(start.c, end.c);
        wrap.querySelectorAll("td.dtd-cell").forEach(td => { const r = +td.dataset.r, cc = +td.dataset.c; if (r >= r1 && r <= r2 && cc >= c1 && cc <= c2) td.classList.add("dtd-pick"); });
    },

    async fireCellClick(r, cc) {
        const c = this.ctx;
        c.host.selection = { r1: r, c1: cc, r2: r, c2: cc };
        const w = this.win();
        try {
            if (w && typeof w.__fireExcelEvent === "function") {
                await w.__fireExcelEvent("onSelectionChanged", r, cc);
                await w.__fireExcelEvent("onSingleClicked", r, cc);
            }
        } catch (err) { console.error(err); }
        this.scheduleRender();
    },

    onKey(e) {
        const c = this.ctx;
        if (!c) return;
        if (e.key === "Escape") {
            if (document.querySelector(".wte-dialog-overlay")) return;
            if (c.overlay.querySelector(".dtd-dlg-back")) { this.closeMemberDialog(); return; }
            if (c.fx) { this.closeFormulaPanel(); return; }
            this.requestClose();
            return;
        }
        const t = e.target;
        const typing = t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.isContentEditable);
        if (typing || c.fx || c.overlay.querySelector(".dtd-dlg-back") || document.querySelector(".wte-dialog-overlay")) return;
        if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); this.clearSelection(); }
    },

    /** Escribe con el "Excel falso" del taskpane (recalcula y repinta al terminar). */
    async excelWrite(fn) {
        const w = this.win();
        if (!w || !w.Excel) { UI.toast("El diseñador todavía se está cargando.", "info"); return; }
        await w.Excel.run(async (context) => { fn(context.workbook.worksheets.getItem("Hoja1")); await context.sync(); });
        this.scheduleRender();
    },

    async clearSelection() {
        const c = this.ctx;
        if (!c || !c.ready) return;
        const s = this.selRect();
        await this.excelWrite(sh => sh.getRangeByIndexes(s.r1, s.c1, s.r2 - s.r1 + 1, s.c2 - s.c1 + 1).clear("Contents"));
    },

    // ---- «Miembro»: EPM_VALUE -------------------------------------------
    async modelFields() {
        const w = this.win();
        const out = [];
        const ag = await w.SemanticModelStore.getModelGrid("MODEL_ATRIBUTES");
        for (let R = 2; String(w.cellValue(ag, R, 2)).trim() !== ""; R++) {
            const dim = String(w.cellValue(ag, R, 2)).trim(), attr = String(w.cellValue(ag, R, 3)).trim();
            if (attr) out.push({ dim, attr, label: dim.toLowerCase() + " · " + attr.toLowerCase() });
        }
        out.sort((a, b) => a.label.localeCompare(b.label));
        const mg = await w.SemanticModelStore.getModelGrid("MODEL_MEASURES");
        if (String(w.cellValue(mg, 2, 2)).trim() !== "") out.push({ dim: "MEASURE", attr: "", label: "Σ Medidas" });
        return out;
    },

    async loadMembers(field) {
        const w = this.win();
        if (field.dim === "MEASURE") {
            const mg = await w.SemanticModelStore.getModelGrid("MODEL_MEASURES");
            const list = [];
            for (let R = 2; String(w.cellValue(mg, R, 2)).trim() !== ""; R++) list.push(String(w.cellValue(mg, R, 2)).trim());
            return list;
        }
        const sql = await w.ExcelService.buildFilterValuesSQL(field.dim, field.attr);
        if (!sql) throw new Error("No se encuentra el atributo en el modelo.");
        const json = await w.ExcelService.executeSQL(sql);
        return Array.from(new Set(w.loadJsonTree(json).map(i => String(i.value)))).slice(0, 5000);
    },

    /**
     * Campo (dimensión/atributo) de una celda del eje Estático, para abrir el selector ya en él:
     *   1) por su posición en el eje según el diseño (EDIT_REPORT: ancla H10/N10 y campos de Filas/Columnas);
     *      en una jerarquía, el nivel de las celdas vecinas de la misma columna/fila (si no, el primero);
     *   2) si no cae en ningún eje, por la celda EPM_VALUE más cercana de su columna (o de su fila).
     * Devuelve { dim, attr } (dim "MEASURE" para medidas) o null.
     */
    async inferField(r, cc) {
        const c = this.ctx, w = this.win(), cells = c.host.state.cells || {};
        const FX = w.__dracoFormulaEngine;
        const epmArgs = (o) => (o && typeof o.f === "string" && /EPM_VALUE\s*\(/i.test(o.f) && FX) ? FX.splitArgs(o.f).map(FX.unq) : null;
        let info = null;
        try {
            info = await w.Excel.run(async (ctx) => {
                const g = await w.getEditReportGrid(ctx);
                const pa = (x) => { const t = String(x || "").trim(); if (!t) return null; try { return w.parseAddress(t); } catch (e) { return null; } };
                return { rr: pa(w.cellValue(g, 10, 8)), rc: pa(w.cellValue(g, 10, 14)), rows: w.readDracoAxisFields(g, false), cols: w.readDracoAxisFields(g, true) };
            });
        } catch (e) { info = null; }
        const nearest = (dr, dc) => {
            for (let k = 1; k <= 300; k++) for (const sgn of [-1, 1]) {
                const a = epmArgs(cells[(r + sgn * k * dr) + "_" + (cc + sgn * k * dc)]);
                if (a) return a;
            }
            return null;
        };
        let field = null, axis = null;
        if (info && info.rr && info.rows.length && r >= info.rr.row - 1 && cc >= info.rr.col - 1 && cc < info.rr.col - 1 + info.rows.length) {
            axis = "rows"; field = info.rows[cc - (info.rr.col - 1)];
        } else if (info && info.rc && info.cols.length && cc >= info.rc.col - 1 && r >= info.rc.row - 1 && r < info.rc.row - 1 + info.cols.length) {
            axis = "cols"; field = info.cols[r - (info.rc.row - 1)];
        }
        const near = axis === "cols" ? nearest(0, 1) : axis === "rows" ? nearest(1, 0) : (nearest(1, 0) || nearest(0, 1));
        if (field) {
            if (field.type === "measure") return { dim: "MEASURE", attr: "" };
            if (field.type === "attr") return { dim: field.dim, attr: field.attr };
            if (field.type === "hier") {
                const lv = near && String(near[0]).toUpperCase() === field.dim.toUpperCase()
                    ? field.levels.find(l => l.attr.toUpperCase() === String(near[1]).toUpperCase()) : null;
                return { dim: field.dim, attr: (lv || field.levels[0]).attr };
            }
        }
        if (near) return String(near[0]).toUpperCase() === "MEASURE" ? { dim: "MEASURE", attr: "" } : { dim: near[0], attr: near[1] };
        return null;
    },

    async openMemberDialog(r, cc) {
        const c = this.ctx;
        if (!c || !c.ready) return;
        if (c.fx) this.closeFormulaPanel();
        const E = UI.escapeHtml;
        const cur = c.host.state.cells[r + "_" + cc] || {};
        const w = this.win();
        let curArgs = null;
        if (typeof cur.f === "string" && /EPM_VALUE\s*\(/i.test(cur.f) && w.__dracoFormulaEngine) curArgs = w.__dracoFormulaEngine.splitArgs(cur.f).map(w.__dracoFormulaEngine.unq);
        // Sin miembro en la celda: el campo se deduce de su fila/columna
        const guess = curArgs ? { dim: curArgs[0], attr: curArgs[1] } : await this.inferField(r, cc);
        const curVar = curArgs && w.DracoVariables ? w.DracoVariables.tokenName(curArgs[2]) : null;

        const back = document.createElement("div");
        back.className = "dtd-dlg-back";
        back.innerHTML = `
            <div class="dtd-dlg" role="dialog" aria-label="Poner miembros">
                <div class="dtd-dlg-head"><strong>Miembro en ${E(this.addr(r, cc))}</strong><button type="button" class="dtd-x" aria-label="Cerrar">×</button></div>
                <div class="dtd-dlg-body">
                    <label class="dtd-lbl">Campo<select id="dtdMField"><option value="">Cargando…</option></select></label>
                    <div class="dtd-tabs" role="tablist">
                        <button type="button" class="dtd-tab active" role="tab" data-tab="values" aria-selected="true">Valores</button>
                        <button type="button" class="dtd-tab" role="tab" data-tab="vars" aria-selected="false">Variables <span id="dtdMVarCount" class="dtd-badge">0</span></button>
                    </div>
                    <input id="dtdMSearch" class="dtd-in" type="text" placeholder="Buscar miembro…" autocomplete="off">
                    <div id="dtdMList" class="dtd-mlist"><div class="dtd-empty">Elige un campo.</div></div>
                    <div id="dtdMVars" class="dtd-mlist" hidden></div>
                    <div class="dtd-mfoot">
                        <span id="dtdMCount" class="dtd-muted">0 seleccionados</span>
                        <span class="dtd-spacer"></span>
                        <label class="dtd-radio"><input type="radio" name="dtdMDir" value="down" checked> Hacia abajo</label>
                        <label class="dtd-radio"><input type="radio" name="dtdMDir" value="right"> Hacia la derecha</label>
                    </div>
                </div>
                <div class="dtd-dlg-foot">
                    <span class="dtd-muted" id="dtdMFootHint">Varios miembros: uno por celda, desde ${E(this.addr(r, cc))}.</span>
                    <span class="dtd-spacer"></span>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdMCancel">Cancelar</button>
                    <button type="button" class="btn btn-primary btn-sm" id="dtdMOk" disabled>Insertar</button>
                </div>
            </div>`;
        c.overlay.appendChild(back);
        const $ = (sel) => back.querySelector(sel);
        const close = () => this.closeMemberDialog();
        $(".dtd-x").addEventListener("click", close);
        $("#dtdMCancel").addEventListener("click", close);
        back.addEventListener("mousedown", (e) => { if (e.target === back) close(); });

        let fields = [], members = [], chosen = new Set(), tab = "values", chosenVar = curVar || null;
        const sel = $("#dtdMField"), list = $("#dtdMList"), search = $("#dtdMSearch"), varsBox = $("#dtdMVars");
        const varsOf = () => { const f = fields[+sel.value]; return (f && f.dim !== "MEASURE" && w.DracoVariables) ? w.DracoVariables.forField(f.dim, f.attr) : []; };
        const fillVars = () => {
            const vs = varsOf();
            $("#dtdMVarCount").textContent = String(vs.length);
            varsBox.innerHTML = vs.length
                ? vs.map(v => `<label class="dtd-mrow"><input type="radio" name="dtdMVar" value="${E(v.name)}"${chosenVar === v.name ? " checked" : ""}><span class="dtd-mono">$${E(v.name)}$</span><span class="dtd-muted">= ${E(v.display || v.value)}</span></label>`).join("")
                : `<div class="dtd-empty">${fields[+sel.value] && fields[+sel.value].dim === "MEASURE" ? "Las medidas no usan variables." : "No hay variables de este campo. Créalas con «Variables» en la cinta."}</div>`;
        };
        const setTab = (t) => {
            tab = t;
            back.querySelectorAll(".dtd-tab").forEach(b => { const on = b.dataset.tab === t; b.classList.toggle("active", on); b.setAttribute("aria-selected", on ? "true" : "false"); });
            list.hidden = search.hidden = t !== "values";
            varsBox.hidden = t !== "vars";
            $(".dtd-mfoot").hidden = t !== "values";
            $("#dtdMFootHint").textContent = t === "vars"
                ? "La celda usará el valor de la variable; si cambia, cambia el informe."
                : "Varios miembros: uno por celda, desde " + this.addr(r, cc) + ".";
            refreshCount();
        };
        back.querySelectorAll(".dtd-tab").forEach(b => b.addEventListener("click", () => setTab(b.dataset.tab)));
        varsBox.addEventListener("change", (e) => { if (e.target.name === "dtdMVar") { chosenVar = e.target.value; refreshCount(); } });
        varsBox.addEventListener("dblclick", (e) => { const row = e.target.closest(".dtd-mrow"); if (!row) return; chosenVar = row.querySelector("input").value; refreshCount(); $("#dtdMOk").click(); });
        const refreshCount = () => {
            $("#dtdMCount").textContent = chosen.size + (chosen.size === 1 ? " seleccionado" : " seleccionados");
            $("#dtdMOk").disabled = tab === "vars" ? !chosenVar : chosen.size === 0;
        };
        const fill = () => {
            const q = search.value.trim().toLowerCase();
            const items = members.filter(m => !q || m.toLowerCase().indexOf(q) !== -1).slice(0, 1500);
            list.innerHTML = items.length
                ? items.map(m => `<label class="dtd-mrow"><input type="checkbox" value="${E(m)}"${chosen.has(m) ? " checked" : ""}><span>${E(m)}</span></label>`).join("")
                : `<div class="dtd-empty">${members.length ? "Sin resultados." : "Sin miembros."}</div>`;
        };
        list.addEventListener("change", (e) => { const cb = e.target; if (cb.type !== "checkbox") return; if (cb.checked) chosen.add(cb.value); else chosen.delete(cb.value); refreshCount(); });
        list.addEventListener("dblclick", (e) => { const row = e.target.closest(".dtd-mrow"); if (!row) return; chosen = new Set([row.querySelector("input").value]); refreshCount(); $("#dtdMOk").click(); });
        search.addEventListener("input", fill);
        const load = async () => {
            const f = fields[+sel.value];
            members = []; chosen = new Set(); refreshCount();
            fillVars();
            if (!f) { list.innerHTML = `<div class="dtd-empty">Elige un campo.</div>`; return; }
            list.innerHTML = `<div class="dtd-empty"><span class="spinner"></span> Cargando miembros…</div>`;
            try { members = await this.loadMembers(f); }
            catch (err) { list.innerHTML = `<div class="dtd-empty bad">${E(err.message || String(err))}</div>`; return; }
            if (curArgs && !curVar && f.dim.toUpperCase() === String(curArgs[0]).toUpperCase()) {
                const v = f.dim === "MEASURE" ? curArgs[1] : curArgs[2];
                if (members.indexOf(v) !== -1) chosen.add(v);
                refreshCount();
            }
            fill();
        };
        sel.addEventListener("change", load);
        $("#dtdMOk").addEventListener("click", async () => {
            const f = fields[+sel.value];
            if (f && tab === "vars") {
                if (!chosenVar) return;
                const token = "$" + chosenVar + "$";
                close();
                await this.excelWrite(sh => { sh.getRangeByIndexes(r, cc, 1, 1).formulas = [[w.buildEpmValueFormula(f.dim, f.attr, token, token)]]; });
                return;
            }
            if (!f || !chosen.size) return;
            const ordered = members.filter(m => chosen.has(m));
            const down = back.querySelector('input[name="dtdMDir"]:checked').value === "down";
            const formulas = ordered.map(m => f.dim === "MEASURE" ? w.buildEpmValueFormula("MEASURE", m, "", m) : w.buildEpmValueFormula(f.dim, f.attr, m, m));
            close();
            await this.excelWrite(sh => formulas.forEach((fx, i) => {
                sh.getRangeByIndexes(down ? r + i : r, down ? cc : cc + i, 1, 1).formulas = [[fx]];
            }));
            c.sel = { r: down ? r + formulas.length - 1 : r, c: down ? cc : cc + formulas.length - 1 };
            c.selRange = { r1: r, c1: cc, r2: c.sel.r, c2: c.sel.c };
            this.scheduleRender();
        });

        try {
            fields = await this.modelFields();
            sel.innerHTML = `<option value="">— Elige un campo —</option>` + fields.map((f, i) => `<option value="${i}">${E(f.label)}</option>`).join("");
            if (guess) {
                const i = fields.findIndex(f => f.dim.toUpperCase() === String(guess.dim).toUpperCase() && (f.dim === "MEASURE" || f.attr.toUpperCase() === String(guess.attr).toUpperCase()));
                if (i >= 0) { sel.value = String(i); await load(); }
            }
            if (curVar) setTab("vars");
            (sel.value ? search : sel).focus();
        } catch (err) {
            sel.innerHTML = `<option value="">No se pudo leer el modelo</option>`;
            list.innerHTML = `<div class="dtd-empty bad">${E(err.message || String(err))}. ¿Has elegido el modelo semántico (cubo) en el panel?</div>`;
        }
    },

    closeMemberDialog() {
        const b = this.ctx && this.ctx.overlay.querySelector(".dtd-dlg-back");
        if (b) b.remove();
    },

    // ---- «Fórmula»: EPM_FORMULA -----------------------------------------
    openFormulaPanel(r, cc) {
        const c = this.ctx;
        if (!c || !c.ready) return;
        this.closeMemberDialog();
        this.closeFormulaPanel();
        const E = UI.escapeHtml, w = this.win();
        const cur = c.host.state.cells[r + "_" + cc] || {};
        let a = ["", "", "", ""];
        if (typeof cur.f === "string" && /EPM_FORMULA\s*\(/i.test(cur.f) && w.__dracoFormulaEngine) {
            const raw = w.__dracoFormulaEngine.splitArgs(cur.f);
            a = [w.__dracoFormulaEngine.unq(raw[0]), raw[1] || "", w.__dracoFormulaEngine.unq(raw[2]), w.__dracoFormulaEngine.unq(raw[3])];
        }
        const bars = String(a[3] || "").toLowerCase().split("|").map(x => x.trim());
        const barKind = bars.indexOf("solo barra") !== -1 ? "barOnly" : (bars.indexOf("barra") !== -1 || bars.indexOf("barras") !== -1 ? "bar" : "");
        const kv = (k, d) => { const x = bars.find(b => b.indexOf(k + "=") === 0); return x ? x.split("=")[1] : d; };
        const noAcc = (t) => t.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        const iconKind = bars.some(b => ["semaforo", "semaforos", "pelotitas"].indexOf(noAcc(b)) !== -1) ? "traffic"
            : (bars.some(b => ["senales", "senal"].indexOf(noAcc(b)) !== -1) ? "signs" : "");
        const cutsTxt = kv("cortes", "");
        const cutsV = cutsTxt ? cutsTxt.split(/[,;]/).map(x => x.trim()) : ["", ""];

        const panel = document.createElement("div");
        panel.className = "dtd-fxpanel";
        panel.setAttribute("role", "dialog");
        panel.innerHTML = `
            <div class="dtd-dlg-head"><strong><i class="dtd-fx">ƒx</i> Fórmula en ${E(this.addr(r, cc))}</strong><button type="button" class="dtd-x" aria-label="Cerrar">×</button></div>
            <div class="dtd-dlg-body">
                <label class="dtd-lbl">Texto<input id="dtdFText" class="dtd-in" type="text" value="${E(a[0] || "")}" placeholder="p. ej. Margen"></label>
                <label class="dtd-lbl">Fórmula<input id="dtdFExpr" class="dtd-in dtd-mono" type="text" value="${E(a[1] || "")}" placeholder="Haz clic en las celdas de la rejilla" autocomplete="off" spellcheck="false"></label>
                <div class="dtd-ops">${["+", "-", "*", "/", "(", ")"].map(o => `<button type="button" class="dtd-op" data-op="${o}">${o === "*" ? "×" : o === "/" ? "÷" : o === "-" ? "−" : o}</button>`).join("")}</div>
                <div class="dtd-ops">${[["SI(", "SI("], [";", ";"], ["=", "="], ["<>", "≠"], ["<", "<"], [">", ">"], ["<=", "≤"], [">=", "≥"], ['""', '"texto"']].map(([o, l]) => `<button type="button" class="dtd-op${o.length > 2 ? " dtd-op-w" : ""}" data-op="${E(o)}" title="${o === '""' ? "Texto entre comillas" : o === "SI(" ? "SI(condición; si se cumple; si no)" : "Comparar"}">${E(l)}</button>`).join("")}</div>
                <div class="dtd-muted">Condiciones: <code>SI(A6="HOLA";5;A6+10)</code>. Compara números o textos (sin distinguir mayúsculas) con = &lt;&gt; &lt; &gt; &lt;= &gt;=.</div>
                <div class="dtd-muted">Haz clic en una celda (o arrastra un rango) de la rejilla para insertarla. Escribe la fórmula como si estuviera en ${E(this.addr(r, cc))}: se copia a cada celda de la línea igual que copiar y pegar (p. ej. las cabeceras de las líneas a restar).</div>
                <div class="dtd-row2">
                    <label class="dtd-lbl">Formato
                        <input id="dtdFFmt" class="dtd-in" type="text" list="dtdFFmts" value="${E(a[2] || "")}" placeholder="(opcional)">
                        <datalist id="dtdFFmts"><option value="#,##0"><option value="#,##0.00"><option value="0%"><option value="0.0%"><option value="0.00%"></datalist>
                    </label>
                    <label class="dtd-lbl">Barras
                        <select id="dtdFBar"><option value="">Sin barras</option><option value="bar"${barKind === "bar" ? " selected" : ""}>Número y barra</option><option value="barOnly"${barKind === "barOnly" ? " selected" : ""}>Solo barra</option></select>
                    </label>
                </div>
                <div class="dtd-row2" id="dtdFBarOpts">
                    <label class="dtd-check"><input id="dtdFInv" type="checkbox"${bars.indexOf("invertir") !== -1 ? " checked" : ""}> Invertir colores</label>
                    <label class="dtd-lbl">Esp. izq. %<input id="dtdFIzq" class="dtd-in" type="number" min="0" max="90" value="${E(kv("izq", "40"))}"></label>
                    <label class="dtd-lbl">Esp. dcha. %<input id="dtdFDcha" class="dtd-in" type="number" min="0" max="90" value="${E(kv("dcha", "25"))}"></label>
                </div>
                <div class="dtd-row2">
                    <label class="dtd-lbl">Iconos
                        <select id="dtdFIcon"><option value="">Sin iconos</option><option value="traffic"${iconKind === "traffic" ? " selected" : ""}>Semáforo (pelotitas)</option><option value="signs"${iconKind === "signs" ? " selected" : ""}>Señales</option></select>
                    </label>
                    <label class="dtd-lbl">Cortes
                        <select id="dtdFCuts"><option value="">Automáticos (33 % / 67 %)</option><option value="fixed"${cutsTxt ? " selected" : ""}>Fijos</option></select>
                    </label>
                </div>
                <div class="dtd-row2" id="dtdFIconOpts">
                    <label class="dtd-lbl" id="dtdFC1L">Ámbar desde<input id="dtdFC1" class="dtd-in" type="number" step="any" value="${E(cutsV[0] || "")}" placeholder="p. ej. 0"></label>
                    <label class="dtd-lbl" id="dtdFC2L">Verde desde<input id="dtdFC2" class="dtd-in" type="number" step="any" value="${E(cutsV[1] || "")}" placeholder="p. ej. 10"></label>
                </div>
                <div class="dtd-row2" id="dtdFIconOpts2">
                    <label class="dtd-check"><input id="dtdFIInv" type="checkbox"${bars.indexOf("invertir-iconos") !== -1 ? " checked" : ""}> Invertir (verde para valores bajos)</label>
                    <label class="dtd-check"><input id="dtdFIOnly" type="checkbox"${bars.indexOf("solo-icono") !== -1 ? " checked" : ""}> Solo el icono</label>
                </div>
                <code id="dtdFPrev" class="dtd-prev"></code>
            </div>
            <div class="dtd-dlg-foot">
                <span id="dtdFErr" class="dtd-err"></span>
                <span class="dtd-spacer"></span>
                <button type="button" class="btn btn-secondary btn-sm" id="dtdFCancel">Cancelar</button>
                <button type="button" class="btn btn-primary btn-sm" id="dtdFOk">${cur.f ? "Guardar" : "Insertar"}</button>
            </div>`;
        c.overlay.appendChild(panel);
        const $ = (s) => panel.querySelector(s);
        c.fx = { r, c: cc, panel, expr: $("#dtdFExpr") };
        // Junto al borde derecho de la rejilla (las cabeceras suelen estar a la izquierda) y arrastrable por el título.
        const wrapBox = c.overlay.querySelector("#dtdGridWrap").getBoundingClientRect();
        const modalBox = c.overlay.querySelector(".dtd-modal").getBoundingClientRect();
        panel.style.left = Math.max(10, wrapBox.right - modalBox.left - panel.offsetWidth - 24) + "px";
        panel.style.top = Math.max(10, wrapBox.top - modalBox.top + 16) + "px";
        panel.style.bottom = "auto";
        const head = panel.querySelector(".dtd-dlg-head");
        head.style.cursor = "move";
        head.addEventListener("mousedown", (e) => {
            if (e.target.closest(".dtd-x")) return;
            e.preventDefault();
            const sx = e.clientX, sy = e.clientY, sl = panel.offsetLeft, st = panel.offsetTop;
            const move = (ev) => { panel.style.left = (sl + ev.clientX - sx) + "px"; panel.style.top = Math.max(0, st + ev.clientY - sy) + "px"; };
            const up = () => { document.removeEventListener("mousemove", move); document.removeEventListener("mouseup", up); };
            document.addEventListener("mousemove", move); document.addEventListener("mouseup", up);
        });
        $(".dtd-x").addEventListener("click", () => this.closeFormulaPanel());
        $("#dtdFCancel").addEventListener("click", () => this.closeFormulaPanel());
        panel.querySelectorAll(".dtd-op").forEach(b => b.addEventListener("click", () => this.fxInsert(b.dataset.op)));
        const syncBar = () => {
            $("#dtdFBarOpts").hidden = !$("#dtdFBar").value; $("#dtdFIzq").disabled = $("#dtdFDcha").disabled = $("#dtdFBar").value !== "bar";
            const ic = $("#dtdFIcon").value;
            $("#dtdFCuts").disabled = !ic; $("#dtdFIconOpts2").hidden = !ic;
            $("#dtdFIconOpts").hidden = !ic || $("#dtdFCuts").value !== "fixed";
            const sig = ic === "signs";
            $("#dtdFC1L").firstChild.textContent = sig ? "Triángulo desde" : "Ámbar desde";
            $("#dtdFC2L").firstChild.textContent = sig ? "Círculo verde desde" : "Verde desde";
        };
        const build = () => {
            const esc = (t) => String(t).replace(/"/g, '""');
            const text = $("#dtdFText").value.trim(), expr = $("#dtdFExpr").value.trim(), fmt = $("#dtdFFmt").value.trim();
            let barSpec = "";
            if ($("#dtdFBar").value) {
                const parts = [$("#dtdFBar").value === "barOnly" ? "solo barra" : "barra"];
                if ($("#dtdFInv").checked) parts.push("invertir");
                if ($("#dtdFBar").value === "bar") { parts.push("izq=" + (Number($("#dtdFIzq").value) || 0)); parts.push("dcha=" + (Number($("#dtdFDcha").value) || 0)); }
                barSpec = parts.join("|");
            }
            const icon = $("#dtdFIcon").value;
            if (icon) {
                const ip = [icon === "signs" ? "señales" : "semaforo"];
                if ($("#dtdFCuts").value === "fixed" && $("#dtdFC1").value !== "" && $("#dtdFC2").value !== "") ip.push("cortes=" + Number($("#dtdFC1").value) + "," + Number($("#dtdFC2").value));
                if ($("#dtdFIInv").checked) ip.push("invertir-iconos");
                if ($("#dtdFIOnly").checked) ip.push("solo-icono");
                barSpec = barSpec ? barSpec + "|" + ip.join("|") : ip.join("|");
            }
            // ";" fuera de comillas -> "," (sintaxis inglesa de Excel, la que usa el addin al copiar la fórmula)
            // y nombres de función en castellano -> inglés (SI -> IF...), fuera de comillas
            const FN = { "SI.ERROR": "IFERROR", "SI": "IF", "SUMA": "SUM", "PROMEDIO": "AVERAGE", "REDONDEAR": "ROUND", "ABS": "ABS" };
            const exprEn = expr.replace(/("(?:[^"]|"")*")|;|\b(SI\.ERROR|SI|SUMA|PROMEDIO|REDONDEAR)\s*\(/gi,
                (m, q, fn) => q ? q : (fn ? FN[fn.toUpperCase()] + "(" : ","));
            const args = ['"' + esc(text) + '"', exprEn];
            if (fmt || barSpec) args.push('"' + esc(fmt) + '"');
            if (barSpec) args.push('"' + barSpec + '"');
            return { text, expr: exprEn, f: "=EPM_FORMULA(" + args.join(",") + ")" };
        };
        const validate = () => {
            const b = build();
            $("#dtdFPrev").textContent = b.f;
            let err = "";
            if (!b.text) err = "Pon un texto.";
            else if (!b.expr) err = "Escribe o compón la fórmula.";
            else {
                let d = 0;
                for (const ch of b.expr.replace(/"[^"]*"/g, "")) { if (ch === "(") d++; if (ch === ")") d--; if (d < 0) break; }
                if (d !== 0) err = "Los paréntesis no cuadran.";
                else if ((b.expr.match(/"/g) || []).length % 2) err = "Hay unas comillas sin cerrar.";
                else if (w.__dracoFormulaEngine) {
                    const v = w.__dracoFormulaEngine.evaluate(b.expr, () => 1);
                    if (v && v.err === "#NAME?") err = "Hay una función o un nombre que no se reconoce.";
                }
            }
            $("#dtdFErr").textContent = err;
            $("#dtdFOk").disabled = !!err;
            return err ? null : b;
        };
        panel.addEventListener("input", validate);
        $("#dtdFBar").addEventListener("change", () => { syncBar(); validate(); });
        $("#dtdFIcon").addEventListener("change", () => { syncBar(); validate(); });
        $("#dtdFCuts").addEventListener("change", () => { syncBar(); validate(); });
        $("#dtdFOk").addEventListener("click", async () => {
            const b = validate();
            if (!b) return;
            this.closeFormulaPanel();
            await this.excelWrite(sh => { sh.getRangeByIndexes(r, cc, 1, 1).formulas = [[b.f]]; });
        });
        syncBar(); validate();
        (a[0] ? $("#dtdFExpr") : $("#dtdFText")).focus();
    },

    /** Inserta texto (referencia u operador) en la fórmula, donde esté el cursor. */
    fxInsert(text) {
        const c = this.ctx;
        if (!c || !c.fx) return;
        const inp = c.fx.expr;
        const s = inp.selectionStart ?? inp.value.length, e = inp.selectionEnd ?? inp.value.length;
        inp.value = inp.value.slice(0, s) + text + inp.value.slice(e);
        const pos = s + text.length;
        inp.focus(); inp.setSelectionRange(pos, pos);
        inp.dispatchEvent(new Event("input", { bubbles: true }));
    },

    closeFormulaPanel() {
        const c = this.ctx;
        if (!c || !c.fx) return;
        try { c.fx.panel.remove(); } catch (e) { /* nada */ }
        c.fx = null; c.fxPick = null;
        this.highlightPick(true);
    },

    // ---------------------------------------------------------------------
    // Guardar / cerrar
    // ---------------------------------------------------------------------
    // ---------------------------------------------------------------------
    // Informes vinculados (ver INFORMES VINCULADOS en ADDIN/src/js/commands.js)
    // Una tabla Draco puede ser un PRINCIPAL con informes vinculados que comparten sus filas (a la derecha) o
    // sus columnas (debajo). Se diseñan de uno en uno (selector «Diseñando») y se pintan y guardan juntos.
    // ---------------------------------------------------------------------
    groupOf() {
        const w = this.win();
        if (!w || !w.ReportStore || typeof w.dracoLinkGroupMain !== "function") return null;
        const act = Number(w.ReportStore.getActiveReportId());
        const m = w.dracoLinkGroupMain(act);
        const mainId = (m === null || m === undefined) ? act : Number(m);
        return { act, mainId, ids: [mainId, ...w.dracoLinkedOf(mainId)] };
    },

    refreshLinkBar() {
        const c = this.ctx, w = this.win();
        if (!c || c.kpi || !c.overlay.querySelector("#dtdLinkBar")) return;
        const g = this.groupOf();
        const bar = c.overlay.querySelector("#dtdLinkBar"), sel = c.overlay.querySelector("#dtdLinkSel");
        if (!g || g.ids.length < 2) { bar.hidden = true; return; }
        sel.innerHTML = g.ids.map((id, i) => {
            const l = i ? w.dracoLinkOf(id) : null;
            const label = i ? `Vinculado ${i} · ${l && l.axis === "cols" ? "columnas" : "filas"}` : "Principal";
            return `<option value="${id}"${id === g.act ? " selected" : ""}>${label}</option>`;
        }).join("");
        bar.hidden = false;
        c.overlay.querySelector("#dtdLinkDel").hidden = g.act === g.mainId;
    },

    async switchGroupReport(id) {
        const w = this.win();
        if (!w || !w.TaskPaneApp) return;
        try {
            if (typeof w.TaskPaneApp.saveDesignOnly === "function") await w.TaskPaneApp.saveDesignOnly();
            await w.TaskPaneApp.onReportSelectorChange(String(id));
        } catch (err) { UI.toast("No se pudo cambiar de informe: " + (err.message || err), "error"); }
        this.refreshLinkBar();
        this.scheduleRender();
    },

    openLinkDialog() {
        const c = this.ctx, w = this.win();
        if (!c || !c.ready || !w || typeof w.dracoCreateLinkedReport !== "function") { UI.toast("El diseñador todavía se está cargando.", "info"); return; }
        const g = this.groupOf();
        const lockAxis = g && g.ids.length > 1 ? w.dracoLinkOf(g.ids[1]).axis : null;
        const old = c.overlay.querySelector(".dtd-linkdlg");
        if (old) old.remove();
        const box = document.createElement("div");
        box.className = "dtd-linkdlg";
        box.setAttribute("role", "dialog");
        box.setAttribute("aria-label", "Vincular informe");
        const opt = (v, t, d) => `<label class="${lockAxis && lockAxis !== v ? "off" : ""}"><input type="radio" name="dtdLinkAxis" value="${v}"${(lockAxis || "rows") === v ? " checked" : ""}${lockAxis && lockAxis !== v ? " disabled" : ""}><span><b>${t}</b><small>${d}</small></span></label>`;
        box.innerHTML = `<h4>Vincular un informe</h4>
            <p class="dtd-linkhint">Se crea un informe nuevo que comparte un eje con el ${g && g.ids.length > 1 ? "principal" : "de esta tabla"}. Tendrá sus propias medidas, el otro eje y sus filtros (por ejemplo, otro año); los filtros de la barra y las variables son los del principal.</p>
            ${opt("rows", "Filas", "Mismas filas. Se pinta a la derecha, sin repetir las cabeceras de fila.")}
            ${opt("cols", "Columnas", "Mismas columnas. Se pinta debajo, sin repetir las cabeceras de columna.")}
            ${lockAxis ? `<p class="dtd-linkhint">Este informe ya tiene vinculados por ${lockAxis === "cols" ? "columnas" : "filas"}: el nuevo se vincula por el mismo eje.</p>` : ""}
            <div class="dtd-linkact"><button type="button" class="btn btn-secondary btn-sm" data-act="cancel">Cancelar</button><button type="button" class="btn btn-primary btn-sm" data-act="ok">Crear vinculado</button></div>`;
        c.overlay.querySelector(".dtd-modal").appendChild(box);
        box.addEventListener("click", async (e) => {
            const b = e.target.closest("[data-act]");
            if (!b) return;
            if (b.dataset.act === "cancel") { box.remove(); return; }
            const axis = (box.querySelector("input[name=dtdLinkAxis]:checked") || {}).value || "rows";
            b.disabled = true;
            try {
                if (typeof w.TaskPaneApp.saveDesignOnly === "function") await w.TaskPaneApp.saveDesignOnly();
                const gg = this.groupOf();
                const newId = await w.dracoCreateLinkedReport(gg.mainId, axis);
                if (typeof w.TaskPaneApp.populateReportSelector === "function") await w.TaskPaneApp.populateReportSelector();
                await w.TaskPaneApp.onReportSelectorChange(String(newId));
                box.remove();
                this.refreshLinkBar();
                this.banner(`Informe vinculado creado. Añádele sus medidas y ${axis === "cols" ? "filas" : "columnas"} (y, si quieres, sus filtros) y pulsa «Refrescar»: se pinta ${axis === "cols" ? "debajo" : "a la derecha"} del principal. Cambia entre informes con «Diseñando».`, "info");
                this.scheduleRender();
            } catch (err) {
                b.disabled = false;
                UI.toast("No se pudo crear el informe vinculado: " + (err.message || err), "error");
            }
        });
    },

    async removeLinked() {
        const c = this.ctx, w = this.win();
        const g = this.groupOf();
        if (!g || g.act === g.mainId) return;
        if (!window.confirm("¿Quitar este informe vinculado? Se borra su diseño y deja de pintarse junto al principal.")) return;
        try {
            const gone = g.act;
            await w.TaskPaneApp.onReportSelectorChange(String(g.mainId));
            if (typeof w.dracoClearLinkedAreas === "function") {
                await w.Excel.run(async (ctx) => { await w.dracoClearLinkedAreas(ctx, [gone]); });
            }
            await w.ReportStore.deleteReport(gone);
            await w.ReportStore.setActiveReportId(g.mainId);
            if (typeof w.TaskPaneApp.populateReportSelector === "function") await w.TaskPaneApp.populateReportSelector();
            this.refreshLinkBar();
            await this.refreshReport();
        } catch (err) { UI.toast("No se pudo quitar el informe vinculado: " + (err.message || err), "error"); }
    },

    async save() {
        const c = this.ctx;
        if (!c || !c.ready) return;
        const nameEl = c.overlay.querySelector("#dtdName");
        const nombre = nameEl.value.trim();
        if (!nombre && !(c.owner && !c.ref)) { UI.toast(c.kpi ? "Ponle un nombre al KPI." : "Ponle un nombre a la tabla.", "error"); nameEl.focus(); return; }
        const btn = c.overlay.querySelector("#dtdSave");
        btn.disabled = true; btn.textContent = "Guardando…";
        try {
            const w = this.win();
            if (w.TaskPaneApp && typeof w.TaskPaneApp.saveDesignOnly === "function") await w.TaskPaneApp.saveDesignOnly();
            const report = w.ReportStore.getActiveReport();
            if (!report || !report.semanticModelName) throw new Error("Elige un modelo semántico (cubo) para el informe.");
            if (c.inlineKpi) {
                this.status("Comprobando el KPI…");
                const parts = await w.DracoEngineApi.kpiParts();
                const def = await DracoTables.exportDefinition(w, c.host);
                const vars = {};
                try { const d = JSON.parse(((def.settings || {}).epm_variables) || "{}"); Object.keys(d).forEach(k => { vars[k.toUpperCase()] = d[k].value; }); } catch (e) { /* sin variables */ }
                c.saved = { kpi: { nombre, cubo: report.semanticModelName, def, parts, vars } };
                UI.toast(`KPI «${nombre}» listo. Guarda el informe para conservarlo.`, "success");
                DracoTables._kpiCache = {};
                this.close();
                return;
            }
            this.status("Generando el SQL…");
            const gen = await w.DracoEngineApi.generateSql();
            if (!gen.sql) throw new Error(gen.mode && gen.mode !== "dynamic"
                ? "El eje Estático no tiene ninguna línea: selecciona celdas y añade miembros con «Miembro» (o doble clic)."
                : "El informe no tiene campos en Filas o Columnas: no hay nada que consultar.");
            const definition = await DracoTables.exportDefinition(w, c.host);
            // KPI: piezas para calcularlo junto a los demás KPIs del informe en una sola consulta (valida que sea un KPI)
            if (c.kpi) {
                this.status("Comprobando el KPI…");
                definition.kpi = await w.DracoEngineApi.kpiParts();
            }
            const cube = c.cubes.find(x => x.name === report.semanticModelName);
            const res = await DracoTables.save({
                id: c.ref, projectId: c.project.PROYECTO_ID, nombre,
                descripcion: c.overlay.querySelector("#dtdDesc").value, cubo: report.semanticModelName,
                definition, sql: gen.sql, modeloFecha: cube ? cube.fecha : "",
                owner: c.ref ? null : c.owner, autoName: c.autoName
            });
            c.saved = { id: res.id, nombre: res.nombre };
            UI.toast(c.kpi ? `KPI «${res.nombre}» guardado.` : `Tabla Draco «${res.nombre}» guardada.`, "success");
            // los valores de KPI en caché ya no valen
            DracoTables._kpiCache = {};
            this.close();
        } catch (err) {
            UI.toast("No se pudo guardar: " + (err.message || err), "error");
            this.status("");
            btn.disabled = false; btn.textContent = "Guardar";
        }
    },

    async requestClose() {
        const c = this.ctx;
        if (!c) return;
        if (c.ready && this.signature() !== c.initialSig) {
            const ok = await UI.confirm("Cerrar sin guardar", "Hay cambios sin guardar en el informe. ¿Cerrar y descartarlos?");
            if (!ok) return;
        }
        this.close();
    },

    close() {
        const c = this.ctx;
        if (!c) return;
        document.removeEventListener("keydown", c.keyHandler);
        if (c.mouseUpHandler) document.removeEventListener("mouseup", c.mouseUpHandler);
        if (c.raf) cancelAnimationFrame(c.raf);
        try { c.overlay.remove(); } catch (e) { /* nada */ }
        this.ctx = null;
        c.resolve(c.saved);
    }
};

window.DracoTableDesigner = DracoTableDesigner;
