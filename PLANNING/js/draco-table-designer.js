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
    TASKPANE_URL: "tablas-draco/taskpane.html?v=20261011a",
    MAX_ROWS: 600,
    MAX_COLS: 80,
    ctx: null,

    /** Abre el diseñador (ref = tabla existente o null para una nueva). Devuelve { id, nombre } al guardar, o null. */
    async open(project, ref) {
        if (this.ctx) throw Object.assign(new Error("El diseñador ya está abierto."), { code: "BUSY" });
        let cubes = [], row = null, def = null;
        try {
            cubes = await DracoTables.cubes(project.PROYECTO_ID);
            if (ref) { row = await DracoTables.get(project.PROYECTO_ID, ref); def = JSON.parse(row.DEFINICION_JSON || "{}"); }
        } catch (err) {
            UI.toast("No se pudo abrir el diseñador: " + err.message, "error");
            throw err;
        }
        return new Promise((resolve) => {
            this.ctx = { project, ref, row, cubes, resolve, saved: null, initialSig: null, ready: false, raf: 0, sel: { r: 0, c: 0 } };
            this.buildModal();
            this.startTaskpane(def);
        });
    },

    // ---------------------------------------------------------------------
    // Ventana
    // ---------------------------------------------------------------------
    buildModal() {
        const c = this.ctx, E = UI.escapeHtml;
        const overlay = document.createElement("div");
        overlay.className = "modal-overlay visible dtd-overlay";
        overlay.id = "dtdOverlay";
        overlay.innerHTML = `
            <div class="dtd-modal" role="dialog" aria-label="Diseñador de tabla Draco">
                <div class="dtd-head">
                    <strong class="dtd-title">Tabla Draco · diseño del informe</strong>
                    <input id="dtdName" class="dtd-name" type="text" maxlength="${DracoTables.MAX_NAME}" placeholder="Nombre de la tabla" value="${E(c.row ? c.row.NOMBRE : "")}">
                    <input id="dtdDesc" class="dtd-desc" type="text" maxlength="${DracoTables.MAX_DESC}" placeholder="Descripción (opcional)" value="${E(c.row ? (c.row.DESCRIPCION || "") : "")}">
                    <span class="dtd-spacer"></span>
                    <span id="dtdStatus" class="dtd-status" role="status"></span>
                    <button type="button" class="btn btn-secondary" id="dtdClose">Cerrar</button>
                    <button type="button" class="btn btn-primary" id="dtdSave" disabled>Guardar</button>
                </div>
                <div class="dtd-ribbon">
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdRefresh" title="Actualizar el informe (genera el SQL y lo ejecuta)">⟳ Actualizar</button>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdFilter" title="Añadir un filtro al informe">▽ Filtro</button>
                    <button type="button" class="btn btn-secondary btn-sm" id="dtdRecog" title="Reconocimiento de miembros">🔎 Reconocimiento</button>
                    <span class="dtd-hint">Diseña el informe en el panel de la derecha. Al guardar se genera su SQL y se guarda con la definición.</span>
                </div>
                <div class="dtd-body">
                    <div class="dtd-grid-wrap" id="dtdGridWrap"><div class="dtd-loading"><span class="spinner"></span> Preparando el diseñador…</div></div>
                    <div class="dtd-resizer" id="dtdResizer" title="Arrastra para cambiar el ancho del panel"></div>
                    <div class="dtd-panel" id="dtdPanel"></div>
                </div>
            </div>`;
        document.body.appendChild(overlay);
        c.overlay = overlay;
        overlay.querySelector("#dtdClose").addEventListener("click", () => this.requestClose());
        overlay.querySelector("#dtdSave").addEventListener("click", () => this.save());
        overlay.querySelector("#dtdRefresh").addEventListener("click", () => this.callCommand("actualizar"));
        overlay.querySelector("#dtdFilter").addEventListener("click", () => this.callCommand("abrirAnadirFiltro"));
        overlay.querySelector("#dtdRecog").addEventListener("click", () => this.callCommand("toggleMemberRecognition"));
        this.wireResizer();
        c.keyHandler = (e) => { if (e.key === "Escape" && !document.querySelector(".wte-dialog-overlay")) this.requestClose(); };
        document.addEventListener("keydown", c.keyHandler);
        overlay.querySelector("#dtdGridWrap").addEventListener("click", (e) => this.onGridClick(e));
    },

    status(text, kind) {
        const el = this.ctx && this.ctx.overlay.querySelector("#dtdStatus");
        if (!el) return;
        el.textContent = text || "";
        el.className = "dtd-status" + (kind ? " " + kind : "");
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
                try { return await DracoTables.resolveModel(c.project, cube); }
                catch (err) { UI.toast(err.message, "error"); return null; }
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
            // Una tabla ya guardada guarda su DISEÑO, no su resultado: se actualiza al abrir para que se vea
            if (c.ref) {
                this.status("Cargando datos…");
                try { await w.DracoEngineApi.refresh(undefined, null); }
                catch (err) { this.status("No se pudo actualizar: " + err.message, "bad"); }
            }
            c.ready = true;
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
        return String(s.epm_reports || "") + "|" + String(s.epm_activeSemanticModel || "");
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
                const sel = (c.sel.r === r && c.sel.c === cc) ? " dtd-sel" : "";
                h += `<td class="dtd-cell${sel}" data-r="${r}" data-c="${cc}"${m ? ` rowspan="${m.rowSpan || 1}" colspan="${m.colSpan || 1}"` : ""} style="${st.join(";")}">${E(text)}</td>`;
            }
            h += `</tr>`;
        }
        h += `</tbody></table>`;
        const top = wrap.scrollTop, left = wrap.scrollLeft;
        wrap.innerHTML = h;
        wrap.scrollTop = top; wrap.scrollLeft = left;
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
    // Guardar / cerrar
    // ---------------------------------------------------------------------
    async save() {
        const c = this.ctx;
        if (!c || !c.ready) return;
        const nameEl = c.overlay.querySelector("#dtdName");
        const nombre = nameEl.value.trim();
        if (!nombre) { UI.toast("Ponle un nombre a la tabla.", "error"); nameEl.focus(); return; }
        const btn = c.overlay.querySelector("#dtdSave");
        btn.disabled = true; btn.textContent = "Guardando…";
        try {
            const w = this.win();
            if (w.TaskPaneApp && typeof w.TaskPaneApp.saveDesignOnly === "function") await w.TaskPaneApp.saveDesignOnly();
            const report = w.ReportStore.getActiveReport();
            const st = w.TaskPaneApp && w.TaskPaneApp.state;
            if (st && (st.rowsStatic || st.colsStatic)) throw new Error("Los informes estáticos (ejes asimétricos) no se pueden guardar como tabla Draco: necesitan teclear miembros en una hoja de Excel. Usa ejes dinámicos.");
            if (!report || !report.semanticModelName) throw new Error("Elige un modelo semántico (cubo) para el informe.");
            this.status("Generando el SQL…");
            const gen = await w.DracoEngineApi.generateSql();
            if (!gen.sql) throw new Error("El informe no tiene campos en Filas o Columnas: no hay nada que consultar.");
            const definition = await DracoTables.exportDefinition(w, c.host);
            const cube = c.cubes.find(x => x.name === report.semanticModelName);
            const res = await DracoTables.save({
                id: c.ref, projectId: c.project.PROYECTO_ID, nombre,
                descripcion: c.overlay.querySelector("#dtdDesc").value, cubo: report.semanticModelName,
                definition, sql: gen.sql, modeloFecha: cube ? cube.fecha : ""
            });
            c.saved = { id: res.id, nombre: res.nombre };
            UI.toast(`Tabla Draco «${res.nombre}» guardada.`, "success");
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
        if (c.raf) cancelAnimationFrame(c.raf);
        try { c.overlay.remove(); } catch (e) { /* nada */ }
        this.ctx = null;
        c.resolve(c.saved);
    }
};

window.DracoTableDesigner = DracoTableDesigner;
