/**
 * ============================================================
 * DRACO PLANNING — EJECUCIÓN / MONITOR DE UN FLUJO (flow_run.html)
 * ============================================================
 * Pantalla independiente que recibe `?flujo_id=...` por querystring y:
 *
 *   1) Pinta la pantalla de variables definida en el editor de flujos
 *      (FLUJOS_SCREEN_BLOCKS / FLUJOS_SCREEN_VARIABLES) con inputs reales.
 *   2) Botón "Ejecutar":
 *        a) sube al storage cada fichero local seleccionado en una
 *           variable de tipo FILE (js/storage.js),
 *        b) llama al orquestador Python (python/flow_runner.py), vía
 *           stored procedure (Snowflake) o endpoint HTTP (BigQuery /
 *           otros), pasándole TODAS las variables de pantalla ya
 *           resueltas (las de fichero, como ruta de storage).
 *   3) Botón "Monitor": pinta la cadena de interfaces del flujo (igual
 *      que en el editor) coloreada según FLUJOS_RUN_STEPS de la última
 *      ejecución, y sondea esa tabla cada pocos segundos mientras el
 *      run esté en curso.
 */
const FlowRun = {
    flujoId: null,
    flow: null,          // { id, name, type, chain:[{id,interfaceId}], screen:{title, blocks} }
    interfaces: [],       // [{id, name, originType}] — solo lo necesario para pintar la cadena
    view: "screen",       // 'screen' | 'monitor'
    lastRunId: null,
    pollHandle: null,

    async init() {
        const params = new URLSearchParams(window.location.search);
        this.flujoId = params.get("flujo_id") || params.get("flowId") || params.get("id");

        if (!this.flujoId) {
            this.renderFatal("Falta el parámetro <code>flujo_id</code> en la URL. Ejemplo: <code>flow_run.html?flujo_id=...</code>");
            return;
        }
        if (!Provider.isConnected() || !Provider.isReady()) {
            window.location.href = "index.html";
            return;
        }

        document.getElementById("connStatus").innerHTML = `<span class="dot"></span>${Provider.label()}`;

        try {
            await this.load();
        } catch (err) {
            this.renderFatal(`Error al cargar el flujo: ${UI.escapeHtml(err.message)}`);
            return;
        }

        this.renderHeader();
        this.renderScreenView();
        this.bindTabs();
    },

    renderFatal(html) {
        document.getElementById("flowRunBody").innerHTML = `<div class="module-empty">${html}</div>`;
    },

    // ------------------------------------------------------------
    // Carga de la definición del flujo (equivalente a Flows.loadDetail,
    // pero autocontenida para no depender de todo js/flows.js)
    // ------------------------------------------------------------
    async load() {
        const rows = await Provider.runQuery(`SELECT * FROM ${Provider.qualifyControl("FLUJOS")} WHERE FLUJO_ID = '${Provider.esc(this.flujoId)}'`);
        const row = rows[0];
        if (!row) throw new Error(`No existe el flujo ${this.flujoId}`);

        const chainRows = await Provider.runQuery(`
            SELECT PASO_ID, INTERFAZ_ID, ORDEN FROM ${Provider.qualifyControl("FLUJOS_INTERFACES")}
            WHERE FLUJO_ID = '${Provider.esc(this.flujoId)}' ORDER BY ORDEN`);
        const chain = chainRows.map(c => ({ id: c.PASO_ID, interfaceId: c.INTERFAZ_ID, orden: c.ORDEN }));

        const blockRows = await Provider.runQuery(`
            SELECT BLOQUE_ID, TIPO, ORDEN, TITULO, CONTENIDO FROM ${Provider.qualifyControl("FLUJOS_SCREEN_BLOCKS")}
            WHERE FLUJO_ID = '${Provider.esc(this.flujoId)}' ORDER BY ORDEN`);
        const varRows = await Provider.runQuery(`
            SELECT VARIABLE_ID, BLOQUE_ID, NOMBRE, ETIQUETA, TIPO, ORDEN FROM ${Provider.qualifyControl("FLUJOS_SCREEN_VARIABLES")}
            WHERE FLUJO_ID = '${Provider.esc(this.flujoId)}' ORDER BY ORDEN`);
        const varsByBloque = {};
        varRows.forEach(v => {
            (varsByBloque[v.BLOQUE_ID] = varsByBloque[v.BLOQUE_ID] || [])
                .push({ id: v.VARIABLE_ID, name: v.NOMBRE, label: v.ETIQUETA || v.NOMBRE, type: v.TIPO || "STRING" });
        });
        const blocks = blockRows.map(b => {
            if (b.TIPO === "VARIABLE") {
                const v = (varsByBloque[b.BLOQUE_ID] || [])[0] || { id: Provider.newId(), name: "", label: "", type: "STRING" };
                return { id: b.BLOQUE_ID, kind: "variable", variable: v };
            }
            if (b.TIPO === "TEXTO") return { id: b.BLOQUE_ID, kind: "text", text: b.CONTENIDO || "" };
            return { id: b.BLOQUE_ID, kind: "frame", title: b.TITULO || "Frame", variables: varsByBloque[b.BLOQUE_ID] || [] };
        });

        this.flow = {
            id: row.FLUJO_ID,
            name: row.FLUJO,
            type: row.TIPO === "MANUAL" ? "manual" : "automatico",
            chain,
            screen: { title: row.SCREEN_TITLE || row.FLUJO, blocks }
        };

        if (chain.length) {
            const ids = chain.map(c => `'${Provider.esc(c.interfaceId)}'`).join(",");
            const ifaceRows = await Provider.runQuery(`
                SELECT INTERFAZ_ID, INTERFAZ, TIPO FROM ${Provider.qualifyControl("INTERFACES")}
                WHERE INTERFAZ_ID IN (${ids})`);
            this.interfaces = ifaceRows.map(r => ({ id: r.INTERFAZ_ID, name: r.INTERFAZ, originType: r.TIPO === "FICHERO" ? "fichero" : "tabla" }));
        } else {
            this.interfaces = [];
        }
    },

    interfaceById(id) {
        return this.interfaces.find(i => i.id === id);
    },

    // ------------------------------------------------------------
    // Cabecera + pestañas Pantalla / Monitor
    // ------------------------------------------------------------
    renderHeader() {
        document.getElementById("flowRunTitle").textContent = this.flow.screen.title || this.flow.name;
        document.getElementById("flowRunSubtitle").innerHTML =
            `Flujo: <strong>${UI.escapeHtml(this.flow.name)}</strong> · ${this.flow.chain.length} paso${this.flow.chain.length === 1 ? "" : "s"}`;
    },

    bindTabs() {
        document.getElementById("btnTabScreen").addEventListener("click", () => this.switchView("screen"));
        document.getElementById("btnTabMonitor").addEventListener("click", () => this.switchView("monitor"));
    },

    switchView(view) {
        this.view = view;
        document.getElementById("btnTabScreen").classList.toggle("active", view === "screen");
        document.getElementById("btnTabMonitor").classList.toggle("active", view === "monitor");
        if (view === "screen") {
            this.stopPolling();
            this.renderScreenView();
        } else {
            this.renderMonitorView();
        }
    },

    // ==============================================================
    // VISTA 1 — PANTALLA DE VARIABLES + EJECUTAR
    // ==============================================================
    renderScreenView() {
        const body = document.getElementById("flowRunBody");
        const blocksHtml = this.flow.screen.blocks.length
            ? this.flow.screen.blocks.map(b => this.blockHtml(b)).join("")
            : `<div class="module-empty">Este flujo no tiene variables de pantalla definidas.</div>`;

        body.innerHTML = `
            <div class="flow-run-screen">
                <div class="flow-screen-blocks flow-screen-blocks--run">${blocksHtml}</div>
                <div class="flow-run-actions">
                    <button class="btn btn-primary" id="btnExecuteFlow">▶ Ejecutar</button>
                    <span class="form-hint" id="flowRunStatusHint"></span>
                </div>
            </div>`;

        document.getElementById("btnExecuteFlow").addEventListener("click", () => this.execute());
    },

    blockHtml(b) {
        if (b.kind === "text") {
            return `<div class="flow-screen-block flow-screen-block--text flow-screen-block--static">${UI.renderFormattedText(b.text)}</div>`;
        }
        if (b.kind === "variable") {
            return `<div class="flow-screen-block flow-screen-block--var flow-screen-block--static">${this.inputHtml(b.variable)}</div>`;
        }
        // frame
        return `
            <div class="flow-screen-block flow-screen-block--frame flow-screen-block--static">
                <div class="flow-frame-header"><strong>${UI.escapeHtml(b.title || "Frame")}</strong></div>
                <div class="flow-frame-vars">
                    ${(b.variables || []).map(v => `<div class="flow-frame-var-row flow-frame-var-row--static">${this.inputHtml(v)}</div>`).join("")}
                </div>
            </div>`;
    },

    /** Input real (no deshabilitado) para una variable de pantalla, según su tipo. */
    inputHtml(v) {
        const id = `runvar_${v.id}`;
        const label = `<label for="${id}">${UI.escapeHtml(v.label || v.name)}</label>`;
        if (v.type === "FILE") {
            return `<div class="flow-field-preview">${label}<input type="file" id="${id}" data-var-name="${UI.escapeHtml(v.name)}" data-var-type="FILE"></div>`;
        }
        if (v.type === "BOOLEAN") {
            return `<div class="flow-field-preview flow-field-preview--checkbox"><input type="checkbox" id="${id}" data-var-name="${UI.escapeHtml(v.name)}" data-var-type="BOOLEAN">${label}</div>`;
        }
        const htmlType = { INTEGER: "number", FLOAT: "number", NUMERIC: "number", DATE: "date", DATETIME: "datetime-local", TIMESTAMP: "datetime-local" }[v.type] || "text";
        return `<div class="flow-field-preview">${label}<input type="${htmlType}" id="${id}" data-var-name="${UI.escapeHtml(v.name)}" data-var-type="${UI.escapeHtml(v.type || "STRING")}"></div>`;
    },

    /** Recoge del DOM {nombre: valor} + {nombre: File} para las de tipo FILE. */
    collectScreenValues() {
        const values = {};
        const files = {};
        document.querySelectorAll("#flowRunBody [data-var-name]").forEach(el => {
            const name = el.dataset.varName;
            const type = el.dataset.varType;
            if (type === "FILE") {
                if (el.files && el.files[0]) files[name] = el.files[0];
            } else if (type === "BOOLEAN") {
                values[name] = el.checked;
            } else {
                values[name] = el.value;
            }
        });
        return { values, files };
    },

    // ------------------------------------------------------------
    // EJECUTAR: sube ficheros locales al storage, luego llama al
    // orquestador (flow_runner.py) con todas las variables resueltas.
    // ------------------------------------------------------------
    async execute() {
        const btn = document.getElementById("btnExecuteFlow");
        const hint = document.getElementById("flowRunStatusHint");
        btn.disabled = true;

        try {
            const { values, files } = this.collectScreenValues();

            // 1) Mover los ficheros locales seleccionados al storage.
            const fileNames = Object.keys(files);
            for (let i = 0; i < fileNames.length; i++) {
                const varName = fileNames[i];
                const file = files[varName];
                const storagePath = Storage.buildPath(this.flujoId, varName, file.name);
                await Storage.upload(file, storagePath, (chunkIdx, totalChunks) => {
                    const progress = totalChunks > 1 ? ` [trozo ${chunkIdx}/${totalChunks}]` : "";
                    hint.textContent = `Subiendo "${file.name}" al storage (${i + 1}/${fileNames.length})${progress}...`;
                });
                // El valor de la variable pasa a ser la ruta de storage: es lo
                // que resuelve 'ruta_storage' en los pasos de tipo FICHERO.
                values[varName] = storagePath;
            }

            // 2) Lanzar el orquestador con las variables ya resueltas.
            hint.textContent = "Lanzando la ejecución...";
            const runId = Provider.newId();
            this.lastRunId = runId;
            await this.callOrchestrator(this.flujoId, values, runId);

            UI.toast(`"${this.flow.name}" lanzado.`, "success");
            hint.textContent = "";
            this.switchView("monitor");
        } catch (err) {
            UI.toast("Error al ejecutar el flujo: " + err.message, "error");
            hint.textContent = "";
        } finally {
            btn.disabled = false;
        }
    },

    /** Invoca python/flow_runner.py: stored procedure (Snowflake) o endpoint HTTP propio (BigQuery / otros). */
    async callOrchestrator(flujoId, variables, runId) {
        const variablesJson = JSON.stringify(variables);

        if (Provider.key() === "snowflake" && DracoConfig.flowRunnerProcedure) {
            const sql = `CALL ${DracoConfig.flowRunnerProcedure}('${Provider.esc(flujoId)}', '${Provider.esc(variablesJson)}', '${Provider.esc(runId)}')`;
            await Provider.runQuery(sql);
            return;
        }

        if (DracoConfig.flowRunnerHttpEndpoint) {
            const resp = await fetch(DracoConfig.flowRunnerHttpEndpoint, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ flujo_id: flujoId, variables, run_id: runId })
            });
            if (!resp.ok) throw new Error(`El orquestador respondió con error (HTTP ${resp.status})`);
            return;
        }

        throw new Error(
            "No hay ningún orquestador configurado: define DracoConfig.flowRunnerProcedure " +
            "(Snowflake) o DracoConfig.flowRunnerHttpEndpoint (BigQuery / otros) en js/config.js."
        );
    },

    // ==============================================================
    // VISTA 2 — MONITOR: cadena de interfaces coloreada por estado
    // ==============================================================
    async renderMonitorView() {
        const body = document.getElementById("flowRunBody");
        body.innerHTML = `
            <div class="flow-run-monitor">
                <div class="flow-chain-wrap" id="monitorChainWrap"><span class="spinner"></span></div>
                <div class="flow-run-monitor-footer" id="monitorFooter"></div>
            </div>`;
        await this.refreshMonitor();
        this.pollHandle = setInterval(() => this.refreshMonitor(), 3000);
    },

    stopPolling() {
        if (this.pollHandle) { clearInterval(this.pollHandle); this.pollHandle = null; }
    },

    async refreshMonitor() {
        try {
            const runRows = await Provider.runQuery(`
                SELECT RUN_ID, ESTADO, MENSAJE, FECHA_INICIO, FECHA_FIN FROM ${Provider.qualifyControl("FLUJOS_RUNS")}
                WHERE FLUJO_ID = '${Provider.esc(this.flujoId)}' ORDER BY FECHA_INICIO DESC LIMIT 1`);
            const run = runRows[0];

            if (!run) {
                document.getElementById("monitorChainWrap").innerHTML = this.chainHtml({});
                document.getElementById("monitorFooter").innerHTML = `<span class="form-hint">Todavía no se ha ejecutado este flujo.</span>`;
                this.stopPolling();
                return;
            }

            const stepRows = await Provider.runQuery(`
                SELECT PASO_ID, ESTADO, FILAS, MENSAJE FROM ${Provider.qualifyControl("FLUJOS_RUN_STEPS")}
                WHERE RUN_ID = '${Provider.esc(run.RUN_ID)}'`);
            const estadoByPaso = {};
            stepRows.forEach(s => { estadoByPaso[s.PASO_ID] = s; });

            document.getElementById("monitorChainWrap").innerHTML = this.chainHtml(estadoByPaso);
            document.getElementById("monitorFooter").innerHTML = this.footerHtml(run, stepRows);

            if (run.ESTADO !== "EN_CURSO" && run.ESTADO !== "PENDIENTE") this.stopPolling();
        } catch (err) {
            document.getElementById("monitorFooter").innerHTML = `<span class="form-hint">Error consultando el estado: ${UI.escapeHtml(err.message)}</span>`;
            this.stopPolling();
        }
    },

    chainHtml(estadoByPaso) {
        if (!this.flow.chain.length) return `<div class="module-empty module-empty--inline">Este flujo no tiene pasos.</div>`;
        return this.flow.chain.map((step, idx) => {
            const iface = this.interfaceById(step.interfaceId);
            const info = estadoByPaso[step.id];
            const estado = (info && info.ESTADO) || "PENDIENTE";
            const card = `
                <div class="flow-chain-card flow-run-step flow-run-step--${estado.toLowerCase()}" title="${info && info.MENSAJE ? UI.escapeHtml(info.MENSAJE) : ""}">
                    <div class="flow-chain-card-name">${iface ? UI.escapeHtml(iface.name) : "<em>Interfaz eliminada</em>"}</div>
                    <div class="flow-chain-card-meta">${this.stepStatusLabel(estado)}${info && typeof info.FILAS === "number" ? ` · ${info.FILAS} filas` : ""}</div>
                </div>`;
            const arrow = idx < this.flow.chain.length - 1 ? `<div class="flow-chain-arrow">→</div>` : "";
            return card + arrow;
        }).join("");
    },

    stepStatusLabel(estado) {
        return { PENDIENTE: "⏳ Pendiente", EN_CURSO: "▶ En ejecución", OK: "✔ Completado", ERROR: "✕ Error" }[estado] || estado;
    },

    footerHtml(run, stepRows) {
        const okCount = stepRows.filter(s => s.ESTADO === "OK").length;
        const badge = { PENDIENTE: "table-tag", EN_CURSO: "table-tag flow-status-scheduled", OK: "table-tag flow-status-ok", ERROR: "table-tag flow-status-error" }[run.ESTADO] || "table-tag";
        return `
            <span class="${badge}">${this.stepStatusLabel(run.ESTADO)}</span>
            <span class="form-hint">${okCount}/${stepRows.length} pasos completados · Run ${UI.escapeHtml(run.RUN_ID)}</span>
            ${run.ESTADO === "ERROR" && run.MENSAJE ? `<div class="form-hint flow-run-error-msg">${UI.escapeHtml(run.MENSAJE)}</div>` : ""}`;
    }
};

document.addEventListener("DOMContentLoaded", () => FlowRun.init());
