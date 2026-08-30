/**
 * ============================================================
 * DRACO PLANNING — WORKFLOWS: EJECUCIONES
 * ============================================================
 * Todo el flujo de ejecuciones vive en un único popup (igual que el
 * editor de un Workflow): listado de ejecuciones → detalle de una
 * ejecución, con "← Volver a ejecuciones" para navegar entre vistas
 * dentro del mismo popup y una X para cerrarlo del todo.
 *
 * Dentro del detalle de una ejecución, los pasos se gestionan **paso a
 * paso**: la cadena de arriba actúa como selector (con un badge de
 * color que indica si el paso está totalmente asignado), y debajo solo
 * se muestran las instancias del paso seleccionado — no todas las
 * instancias de todos los pasos apiladas.
 *
 *   - "Nueva ejecución" abre un popup más pequeño (nombre + variables
 *     del Paso 0) y, al crear, instancia todos los pasos ejecutables:
 *     una instancia por valor del driver (los elegidos en la
 *     definición, o todos los reales de la dimensión si no se eligió
 *     ninguno), o una única instancia si el paso no tiene driver.
 *   - Cada instancia arranca Bloqueada, Pendiente o Programada según el
 *     "Inicio" del paso. Se puede asignar (texto libre, todavía no hay
 *     módulo de Roles/usuarios), completar sus variables, y mover de
 *     estado: Pendiente → En curso → (si el paso requiere revisión) En
 *     revisión → Completado.
 *   - Al completarse TODAS las instancias de un paso se desbloquean
 *     automáticamente las del siguiente si su inicio es "al completar
 *     el paso anterior".
 *   - Las tareas del paso se muestran a título informativo; las de tipo
 *     "Flujo manual" enlazan directamente a flow_run.html.
 *
 * Persistencia en DRACO_CONTROL:
 *   - WORKFLOWS_RUNS              cabecera de la ejecución
 *   - WORKFLOWS_RUNS_INSTANCIAS   una fila por instancia de paso
 *   - WORKFLOWS_RUNS_VARIABLES    variables globales (Paso 0, INSTANCIA_ID = RUN_ID) y por instancia
 */
const WorkflowRuns = {
    ESTADOS: {
        BLOQUEADO: { label: "Bloqueado", cls: "flow-run-step--pendiente" },
        PENDIENTE: { label: "Pendiente", cls: "flow-run-step--pendiente" },
        EN_CURSO: { label: "En curso", cls: "flow-run-step--en_curso" },
        EN_REVISION: { label: "En revisión", cls: "flow-run-step--en_curso" },
        COMPLETADO: { label: "Completado", cls: "flow-run-step--ok" }
    },

    project: null,
    workflow: null,          // definición completa (Workflows.loadDetail)
    runs: [],
    currentRun: null,        // { id, name, estado, variables:{}, instancias:[...] }
    selectedRunStepId: null,
    view: "list",             // "list" | "detail"

    execSteps() {
        return this.workflow.steps.filter(s => !s.isPaso0);
    },

    paso0Step() {
        return this.workflow.steps.find(s => s.isPaso0) || null;
    },

    // ------------------------------------------------------------
    // Entrada: popup con el listado de ejecuciones del workflow
    // ------------------------------------------------------------
    async open(container, project, workflowId) {
        this.project = project;
        this.currentRun = null;
        this.view = "list";

        this.ensureOverlay();
        this.setModalTitle("Ejecuciones", "Cargando…");
        document.getElementById("wfRunsModalBody").innerHTML = `<span class="spinner"></span>`;
        this.overlay.classList.add("visible");

        this.workflow = await Workflows.loadDetail(workflowId);
        if (!this.workflow) { UI.toast("No se ha podido cargar el workflow.", "error"); this.closeModal(); return; }

        await this.loadRuns();
        this.showRunsList();
    },

    ensureOverlay() {
        let overlay = document.getElementById("wfRunsModal");
        if (overlay) { this.overlay = overlay; return; }

        overlay = document.createElement("div");
        overlay.className = "modal-overlay";
        overlay.id = "wfRunsModal";
        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div class="modal-title-edit-wrap">
                        <h3 id="wfRunsModalTitle"></h3>
                        <span class="modal-subtitle" id="wfRunsModalSubtitle"></span>
                    </div>
                    <div class="modal-header-right">
                        <button class="modal-close" id="wfRunsModalClose">&times;</button>
                    </div>
                </div>
                <div class="modal-body modal-body-flush" id="wfRunsModalBody" style="overflow-y:auto; padding:20px 24px;"></div>
            </div>`;
        document.body.appendChild(overlay);
        overlay.querySelector("#wfRunsModalClose").addEventListener("click", () => this.closeModal());
        this.overlay = overlay;
    },

    setModalTitle(title, subtitle) {
        document.getElementById("wfRunsModalTitle").textContent = title;
        document.getElementById("wfRunsModalSubtitle").textContent = subtitle || "";
    },

    closeModal() {
        if (this.overlay) this.overlay.classList.remove("visible");
    },

    async loadRuns() {
        try {
            this.runs = await Provider.runQuery(`
                SELECT RUN_ID, NOMBRE, ESTADO, FECHA_CREACION, FECHA_FIN
                FROM ${Provider.qualifyControl("WORKFLOWS_RUNS")}
                WHERE WORKFLOW_ID = '${Provider.esc(this.workflow.id)}'
                ORDER BY FECHA_CREACION DESC`);
        } catch (err) {
            this.runs = [];
            UI.toast("Error al cargar las ejecuciones: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Vista: listado de ejecuciones
    // ------------------------------------------------------------
    showRunsList() {
        this.view = "list";
        const steps = this.execSteps();
        this.setModalTitle(`Ejecuciones de "${this.workflow.name}"`, `${steps.length} paso${steps.length === 1 ? "" : "s"} definido${steps.length === 1 ? "" : "s"}`);

        const body = document.getElementById("wfRunsModalBody");
        body.innerHTML = `
            <div class="module-header" style="margin-bottom:16px;">
                <div></div>
                <button class="btn btn-primary btn-sm" id="btnNewRun" ${steps.length ? "" : "disabled"}>+ Nueva ejecución</button>
            </div>
            <div id="wfRunsListWrap">
                ${!steps.length ? `<div class="module-empty">Este workflow todavía no tiene pasos definidos.</div>` : ""}
                ${steps.length && !this.runs.length ? `<div class="module-empty">Todavía no hay ejecuciones. Crea la primera con "Nueva ejecución".</div>` : ""}
                ${this.runs.length ? `
                    <div class="data-list">
                        <table>
                            <thead><tr><th>Ejecución</th><th>Estado</th><th>Creada</th><th></th></tr></thead>
                            <tbody>
                                ${this.runs.map(r => `
                                    <tr>
                                        <td><strong>${UI.escapeHtml(r.NOMBRE)}</strong></td>
                                        <td><span class="table-tag ${r.ESTADO === 'COMPLETADO' ? 'flow-status-ok' : ''}">${r.ESTADO === "COMPLETADO" ? "Completada" : "En curso"}</span></td>
                                        <td>${UI.escapeHtml(String(r.FECHA_CREACION || "").substring(0, 19).replace("T", " "))}</td>
                                        <td>
                                            <div class="row-actions">
                                                <button data-open-run="${r.RUN_ID}" title="Abrir">▶</button>
                                                <button data-del-run="${r.RUN_ID}" class="danger" title="Eliminar">🗑</button>
                                            </div>
                                        </td>
                                    </tr>`).join("")}
                            </tbody>
                        </table>
                    </div>` : ""}
            </div>`;

        const newBtn = document.getElementById("btnNewRun");
        if (newBtn) newBtn.addEventListener("click", () => this.createRun());
        body.querySelectorAll("[data-open-run]").forEach(btn =>
            btn.addEventListener("click", () => this.openRun(btn.dataset.openRun)));
        body.querySelectorAll("[data-del-run]").forEach(btn =>
            btn.addEventListener("click", () => this.removeRun(btn.dataset.delRun)));
    },

    async removeRun(runId) {
        const run = this.runs.find(r => r.RUN_ID === runId);
        if (!run) return;
        const ok = await UI.confirm("Eliminar ejecución", `Se eliminará la ejecución <strong>${UI.escapeHtml(run.NOMBRE)}</strong> y todo su progreso.`);
        if (!ok) return;
        try {
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")} WHERE RUN_ID = '${Provider.esc(runId)}'`);
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")} WHERE RUN_ID = '${Provider.esc(runId)}'`);
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_RUNS")} WHERE RUN_ID = '${Provider.esc(runId)}'`);
            await this.loadRuns();
            this.showRunsList();
            UI.toast(`Ejecución "${run.NOMBRE}" eliminada.`, "success");
        } catch (err) {
            UI.toast("Error al eliminar la ejecución: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Popup compacto: nombre + variables del Paso 0
    // ------------------------------------------------------------
    async createRun() {
        const result = await this.openNewRunModal();
        if (!result) return;

        UI.toast("Creando ejecución…", "info");
        try {
            const runId = Provider.newId();
            const pid = this.project.PROYECTO_ID;
            const steps = this.execSteps();

            const instancias = [];
            for (let idx = 0; idx < steps.length; idx++) {
                const step = steps[idx];
                let valores = [null];
                if (step.driver.dimensionId) {
                    if (step.driver.valores.length) {
                        valores = step.driver.valores;
                    } else {
                        const dim = Workflows.dimensionById(step.driver.dimensionId);
                        if (dim) {
                            const keyCol = Provider.toIdentifier(dim.DIMENSION);
                            const fullTable = Provider.qualify(this.project.DATASET, dim.TABLA);
                            const rows = await Provider.runQuery(`SELECT DISTINCT ${keyCol} AS V FROM ${fullTable} ORDER BY ${keyCol} LIMIT 5000`);
                            valores = rows.length ? rows.map(r => String(r.V)) : [null];
                        }
                    }
                }

                let estadoInicial = "PENDIENTE";
                let fechaProgramada = "";
                if (step.inicio.tipo === "PASO_ANTERIOR" && idx > 0) estadoInicial = "BLOQUEADO";
                if (step.inicio.tipo === "FECHA") {
                    fechaProgramada = step.inicio.fecha || "";
                    estadoInicial = "PENDIENTE"; // sin scheduler: queda disponible, solo se informa la fecha prevista
                }

                valores.forEach(v => {
                    instancias.push({
                        id: Provider.newId(), pasoId: step.id, orden: idx,
                        driverValor: v, asignado: "", estado: estadoInicial,
                        fechaProgramada
                    });
                });
            }

            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_RUNS")}
                (RUN_ID, WORKFLOW_ID, PROYECTO_ID, NOMBRE, ESTADO, USUARIO, FECHA_CREACION, FECHA_FIN)
                VALUES ('${Provider.esc(runId)}', '${Provider.esc(this.workflow.id)}', '${Provider.esc(pid)}', '${Provider.esc(result.name)}', 'EN_CURSO',
                        ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), NULL)`);

            if (instancias.length) {
                const vals = instancias.map(i => `(
                    '${Provider.esc(runId)}', '${Provider.esc(i.pasoId)}', '${Provider.esc(i.id)}', ${i.orden},
                    ${i.driverValor !== null ? `'${Provider.esc(i.driverValor)}'` : "NULL"},
                    '${Provider.esc(i.asignado)}', '${Provider.esc(i.estado)}', '${Provider.esc(i.fechaProgramada)}', NULL, NULL
                )`).join(",\n");
                await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                    (RUN_ID, PASO_ID, INSTANCIA_ID, ORDEN, DRIVER_VALOR, ASIGNADO, ESTADO, FECHA_PROGRAMADA, FECHA_INICIO, FECHA_FIN)
                    VALUES ${vals}`);
            }

            const globalVarEntries = Object.entries(result.variables || {});
            if (globalVarEntries.length) {
                const varVals = globalVarEntries.map(([name, value]) =>
                    `('${Provider.esc(runId)}', '${Provider.esc(runId)}', '${Provider.esc(name)}', '${Provider.esc(value)}')`);
                await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")}
                    (RUN_ID, INSTANCIA_ID, NOMBRE, VALOR) VALUES ${varVals.join(",\n")}`);
            }

            await this.loadRuns();
            await this.openRun(runId);
            UI.toast(`Ejecución "${result.name}" creada con ${instancias.length} instancia(s).`, "success");
        } catch (err) {
            UI.toast("Error al crear la ejecución: " + err.message, "error");
        }
    },

    openNewRunModal() {
        return new Promise((resolve) => {
            const paso0 = this.paso0Step();
            const variables = paso0 ? paso0.variables : [];

            let overlay = document.getElementById("wfNewRunModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wfNewRunModal";
                document.body.appendChild(overlay);
            }

            const inputType = (t) => ({ INTEGER: "number", FLOAT: "number", NUMERIC: "number", DATE: "date", DATETIME: "datetime-local", TIMESTAMP: "datetime-local" }[t] || "text");

            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <div>
                            <h3>Nueva ejecución</h3>
                            <span class="modal-subtitle">${UI.escapeHtml(this.workflow.name)}</span>
                        </div>
                        <button class="modal-close" id="wfNewRunClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Nombre de la ejecución</label>
                            <input type="text" id="wfNewRunName" placeholder="Ej. Cierre Enero 2026">
                        </div>
                        ${variables.length ? `
                            <div class="flow-step-group-title" style="margin-top:6px;">Variables del workflow</div>
                            <div class="wf-newrun-vars">
                                ${variables.map(v => `
                                    <div class="form-group">
                                        <label>${UI.escapeHtml(v.label)}</label>
                                        <input type="${inputType(v.type)}" data-run-var="${UI.escapeHtml(v.name)}" placeholder="${UI.escapeHtml(v.label)}">
                                    </div>`).join("")}
                            </div>` : `<p class="form-hint">Este workflow no tiene variables definidas en el Paso 0.</p>`}
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="wfNewRunCancel">Cancelar</button>
                        <button class="btn btn-primary" id="wfNewRunCreate">Crear ejecución</button>
                    </div>
                </div>`;

            overlay.classList.add("visible");
            const nameInput = overlay.querySelector("#wfNewRunName");
            setTimeout(() => nameInput.focus(), 50);

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#wfNewRunClose").onclick = () => cleanup(null);
            overlay.querySelector("#wfNewRunCancel").onclick = () => cleanup(null);
            overlay.querySelector("#wfNewRunCreate").onclick = () => {
                const name = nameInput.value.trim();
                if (!name) { UI.toast("Indica un nombre para la ejecución.", "error"); return; }
                const values = {};
                overlay.querySelectorAll("[data-run-var]").forEach(input => { values[input.dataset.runVar] = input.value; });
                cleanup({ name, variables: values });
            };
            nameInput.onkeydown = (e) => { if (e.key === "Enter") overlay.querySelector("#wfNewRunCreate").click(); };
        });
    },

    // ------------------------------------------------------------
    // Vista: detalle de una ejecución (paso a paso)
    // ------------------------------------------------------------
    async openRun(runId) {
        const body = document.getElementById("wfRunsModalBody");
        body.innerHTML = `<span class="spinner"></span>`;
        try {
            const headRows = await Provider.runQuery(`SELECT * FROM ${Provider.qualifyControl("WORKFLOWS_RUNS")} WHERE RUN_ID = '${Provider.esc(runId)}'`);
            const head = headRows[0];
            if (!head) { UI.toast("No se ha encontrado la ejecución.", "error"); this.showRunsList(); return; }

            const instRows = await Provider.runQuery(`
                SELECT * FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                WHERE RUN_ID = '${Provider.esc(runId)}' ORDER BY ORDEN`);
            const varRows = await Provider.runQuery(`
                SELECT INSTANCIA_ID, NOMBRE, VALOR FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")}
                WHERE RUN_ID = '${Provider.esc(runId)}'`);

            const varsByInst = {};
            const globalVars = {};
            varRows.forEach(v => {
                if (v.INSTANCIA_ID === runId) { globalVars[v.NOMBRE] = v.VALOR; return; }
                (varsByInst[v.INSTANCIA_ID] = varsByInst[v.INSTANCIA_ID] || {})[v.NOMBRE] = v.VALOR;
            });

            this.currentRun = {
                id: head.RUN_ID, name: head.NOMBRE, estado: head.ESTADO,
                variables: globalVars,
                instancias: instRows.map(i => ({
                    id: i.INSTANCIA_ID, pasoId: i.PASO_ID, orden: parseInt(i.ORDEN || "0", 10),
                    driverValor: i.DRIVER_VALOR || null, asignado: i.ASIGNADO || "",
                    estado: i.ESTADO, fechaProgramada: i.FECHA_PROGRAMADA || "",
                    variables: varsByInst[i.INSTANCIA_ID] || {}
                }))
            };
            const steps = this.execSteps();
            this.selectedRunStepId = steps.length ? steps[0].id : null;
            this.view = "detail";
            this.renderRunDetail();
        } catch (err) {
            UI.toast("Error al abrir la ejecución: " + err.message, "error");
            this.showRunsList();
        }
    },

    instancesForStep(pasoId) {
        return this.currentRun.instancias.filter(i => i.pasoId === pasoId);
    },

    currentRunStep() {
        return this.execSteps().find(s => s.id === this.selectedRunStepId) || null;
    },

    renderRunDetail() {
        const run = this.currentRun;
        const steps = this.execSteps();
        const total = run.instancias.length;
        const done = run.instancias.filter(i => i.estado === "COMPLETADO").length;
        const pct = total ? Math.round((done / total) * 100) : 0;
        const globalVarEntries = Object.entries(run.variables || {});

        this.setModalTitle(run.name, this.workflow.name);

        const body = document.getElementById("wfRunsModalBody");
        body.innerHTML = `
            <button class="link-btn" id="btnBackToRuns" style="margin-bottom:12px;">← Volver a ejecuciones</button>
            <div class="wf-run-summary">
                <div class="wf-run-summary-progress">
                    <div class="wf-run-summary-progress-top">
                        <span>${done}/${total} instancias completadas</span>
                        <span class="table-tag ${run.estado === 'COMPLETADO' ? 'flow-status-ok' : ''}">${run.estado === "COMPLETADO" ? "Completada" : "En curso"}</span>
                    </div>
                    <div class="wf-run-progress-bar"><div class="wf-run-progress-fill" style="width:${pct}%;"></div></div>
                </div>
                ${globalVarEntries.length ? `
                    <div class="chip-row">
                        ${globalVarEntries.map(([k, v]) => `<span class="hier-chip">${UI.escapeHtml(k)} = ${UI.escapeHtml(v || "—")}</span>`).join("")}
                    </div>` : ""}
            </div>

            <div class="flow-chain-wrap wf-run-tabs" id="wfRunChainWrap"></div>
            <div id="wfRunStepBody"></div>`;

        document.getElementById("btnBackToRuns").addEventListener("click", () => this.showRunsList());

        this.renderRunChain();
        this.renderRunStepBody();
    },

    renderRunChain() {
        const steps = this.execSteps();
        const chainWrap = document.getElementById("wfRunChainWrap");
        if (!chainWrap) return;

        chainWrap.innerHTML = steps.map((step, idx) => {
            const instances = this.instancesForStep(step.id);
            const assignedCount = instances.filter(i => i.asignado && i.asignado.trim()).length;
            const fullyAssigned = instances.length > 0 && assignedCount === instances.length;
            const selected = step.id === this.selectedRunStepId;
            const card = `
                <div class="flow-chain-card wf-run-tab ${selected ? "is-selected" : ""}" data-run-step="${step.id}">
                    <div class="flow-chain-card-name">${idx + 1}. ${UI.escapeHtml(step.name)}</div>
                    <span class="wf-run-chain-badge ${fullyAssigned ? "wf-run-chain-badge--assigned" : "wf-run-chain-badge--unassigned"}">
                        ${fullyAssigned ? "✓" : "⚠"} ${assignedCount}/${instances.length} asignada${instances.length === 1 ? "" : "s"}
                    </span>
                </div>`;
            const arrow = idx < steps.length - 1 ? `<div class="flow-chain-arrow">→</div>` : "";
            return card + arrow;
        }).join("");

        chainWrap.querySelectorAll("[data-run-step]").forEach(card => {
            card.addEventListener("click", () => {
                this.selectedRunStepId = card.dataset.runStep;
                this.renderRunChain();
                this.renderRunStepBody();
            });
        });
    },

    renderRunStepBody() {
        const wrap = document.getElementById("wfRunStepBody");
        const step = this.currentRunStep();
        if (!step) { wrap.innerHTML = `<div class="module-empty">Este workflow no tiene pasos ejecutables.</div>`; return; }

        const instances = this.instancesForStep(step.id);
        const dim = step.driver.dimensionId ? Workflows.dimensionById(step.driver.dimensionId) : null;
        const taskCount = step.bloques.reduce((n, b) => n + b.tareas.length, 0);

        wrap.innerHTML = `
            <div class="flow-screen-block flow-screen-block--frame" style="margin-bottom:0;">
                <div class="flow-frame-header">
                    <span><strong>${UI.escapeHtml(step.name)}</strong></span>
                    <span class="load-fn-toolbar-spacer"></span>
                    ${step.revision ? `<span class="table-tag">Con revisión</span>` : ""}
                    ${dim ? `<span class="table-tag">Driver: ${UI.escapeHtml(dim.DIMENSION)}</span>` : ""}
                    ${taskCount ? `<button type="button" class="btn btn-secondary btn-sm" id="btnToggleTasks">Tareas (${taskCount})</button>` : ""}
                </div>
                <div class="flow-frame-vars" id="wfTasksPanel" style="display:none;">
                    ${step.bloques.map(b => `
                        <p class="form-hint"><strong>${UI.escapeHtml(b.titulo)}</strong></p>
                        ${b.tareas.map(t => this.taskRowHtml(t)).join("")}`).join("") || `<p class="form-hint">Sin tareas.</p>`}
                </div>
                <div class="wf-instance-grid">
                    ${instances.map(i => this.instanceCardHtml(step, i)).join("")}
                </div>
            </div>`;

        const toggleBtn = document.getElementById("btnToggleTasks");
        if (toggleBtn) toggleBtn.addEventListener("click", () => {
            const panel = document.getElementById("wfTasksPanel");
            panel.style.display = panel.style.display === "none" ? "block" : "none";
        });

        this.bindRunEvents(wrap);
    },

    taskRowHtml(t) {
        const typeInfo = Workflows.TASK_TYPES[t.tipo] || { label: t.tipo, icon: "•" };
        const runLink = t.tipo === "FLUJO_MANUAL" && t.refId
            ? `<button type="button" class="btn btn-secondary btn-sm" data-open-flow="${t.refId}">▶ Ejecutar</button>`
            : "";
        const showRef = t.refNombre && t.refNombre !== t.nombre;
        return `<div class="flow-target-row" title="${UI.escapeHtml(t.descripcion || "")}"><span class="flow-target-label">${typeInfo.icon} ${UI.escapeHtml(t.nombre || typeInfo.label)}${showRef ? " — " + UI.escapeHtml(t.refNombre) : ""}</span>${runLink}</div>`;
    },

    instanceCardHtml(step, inst) {
        const estadoInfo = this.ESTADOS[inst.estado] || this.ESTADOS.PENDIENTE;
        const visibleVars = step.variables || [];

        let actions = "";
        if (inst.estado === "BLOQUEADO") {
            actions = `<span class="form-hint">Se desbloqueará al completarse el paso anterior.</span>`;
        } else if (inst.estado === "PENDIENTE") {
            actions = `<button class="btn btn-secondary btn-sm" data-inst-action="start:${inst.id}">Iniciar</button>`;
        } else if (inst.estado === "EN_CURSO") {
            actions = step.revision
                ? `<button class="btn btn-primary btn-sm" data-inst-action="review:${inst.id}">Enviar a revisión</button>`
                : `<button class="btn btn-primary btn-sm" data-inst-action="complete:${inst.id}">Completar</button>`;
        } else if (inst.estado === "EN_REVISION") {
            actions = `<button class="btn btn-primary btn-sm" data-inst-action="approve:${inst.id}">Aprobar</button>
                       <button class="btn btn-secondary btn-sm" data-inst-action="reject:${inst.id}">Rechazar</button>`;
        } else if (inst.estado === "COMPLETADO") {
            actions = `<span class="form-hint">Completado.</span>`;
        }

        return `
            <div class="wf-instance-card ${estadoInfo.cls}" data-instance-card="${inst.id}">
                <div class="wf-instance-card-top">
                    <span class="wf-instance-card-title">${inst.driverValor !== null ? UI.escapeHtml(inst.driverValor) : "Instancia única"}</span>
                    <span class="table-tag">${estadoInfo.label}</span>
                </div>
                ${inst.fechaProgramada ? `<span class="table-tag" style="margin-bottom:8px;">Programado: ${UI.escapeHtml(inst.fechaProgramada)}</span>` : ""}
                <div class="form-group">
                    <label>Asignado a</label>
                    <input type="text" placeholder="Persona o grupo..." value="${UI.escapeHtml(inst.asignado)}" data-inst-assignee="${inst.id}" ${inst.estado === "COMPLETADO" ? "disabled" : ""}>
                </div>
                ${visibleVars.length ? `
                    <div class="form-group">
                        ${visibleVars.map(v => `
                            <label>${UI.escapeHtml(v.label)}</label>
                            <input type="text" value="${UI.escapeHtml(inst.variables[v.name] || "")}" data-inst-var="${inst.id}:${UI.escapeHtml(v.name)}" ${inst.estado === "COMPLETADO" ? "disabled" : ""} style="margin-bottom:6px;">
                        `).join("")}
                    </div>` : ""}
                <div class="row-actions" style="justify-content:flex-start; gap:8px;">${actions}</div>
            </div>`;
    },

    bindRunEvents(scope) {
        scope.querySelectorAll("[data-open-flow]").forEach(btn => {
            btn.addEventListener("click", () => window.open(`flow_run.html?flujo_id=${encodeURIComponent(btn.dataset.openFlow)}`, "_blank"));
        });
        scope.querySelectorAll("[data-inst-assignee]").forEach(input => {
            input.addEventListener("change", () => this.updateAssignee(input.dataset.instAssignee, input.value));
        });
        scope.querySelectorAll("[data-inst-var]").forEach(input => {
            input.addEventListener("change", () => {
                const [instId, varName] = input.dataset.instVar.split(":");
                this.updateVariable(instId, varName, input.value);
            });
        });
        scope.querySelectorAll("[data-inst-action]").forEach(btn => {
            btn.addEventListener("click", () => {
                const [action, instId] = btn.dataset.instAction.split(":");
                this.transition(instId, action);
            });
        });
    },

    async updateAssignee(instId, value) {
        try {
            await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                SET ASIGNADO = '${Provider.esc(value)}' WHERE INSTANCIA_ID = '${Provider.esc(instId)}'`);
            const inst = this.currentRun.instancias.find(i => i.id === instId);
            if (inst) inst.asignado = value;
            this.renderRunChain();
        } catch (err) {
            UI.toast("Error al guardar el responsable: " + err.message, "error");
        }
    },

    async updateVariable(instId, name, value) {
        try {
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")}
                WHERE RUN_ID = '${Provider.esc(this.currentRun.id)}' AND INSTANCIA_ID = '${Provider.esc(instId)}' AND NOMBRE = '${Provider.esc(name)}'`);
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")}
                (RUN_ID, INSTANCIA_ID, NOMBRE, VALOR) VALUES
                ('${Provider.esc(this.currentRun.id)}', '${Provider.esc(instId)}', '${Provider.esc(name)}', '${Provider.esc(value)}')`);
            const inst = this.currentRun.instancias.find(i => i.id === instId);
            if (inst) inst.variables[name] = value;
        } catch (err) {
            UI.toast("Error al guardar la variable: " + err.message, "error");
        }
    },

    async transition(instId, action) {
        const inst = this.currentRun.instancias.find(i => i.id === instId);
        if (!inst) return;

        const map = { start: "EN_CURSO", review: "EN_REVISION", approve: "COMPLETADO", reject: "EN_CURSO", complete: "COMPLETADO" };
        const nuevoEstado = map[action];
        if (!nuevoEstado) return;

        const isFinal = nuevoEstado === "COMPLETADO";
        try {
            await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                SET ESTADO = '${Provider.esc(nuevoEstado)}'${isFinal ? ", FECHA_FIN = CURRENT_TIMESTAMP()" : ""}${action === "start" ? ", FECHA_INICIO = CURRENT_TIMESTAMP()" : ""}
                WHERE INSTANCIA_ID = '${Provider.esc(instId)}'`);
            inst.estado = nuevoEstado;

            if (isFinal) await this.cascadeUnblock(inst.pasoId);
            await this.maybeCompleteRun();
            this.renderRunDetail();
        } catch (err) {
            UI.toast("Error al actualizar el estado: " + err.message, "error");
        }
    },

    async cascadeUnblock(pasoId) {
        const steps = this.execSteps();
        const stepIdx = steps.findIndex(s => s.id === pasoId);
        if (stepIdx === -1 || stepIdx === steps.length - 1) return;

        const stepInstances = this.instancesForStep(pasoId);
        if (!stepInstances.every(i => i.estado === "COMPLETADO")) return;

        const nextStep = steps[stepIdx + 1];
        if (nextStep.inicio.tipo !== "PASO_ANTERIOR") return;

        const toUnblock = this.instancesForStep(nextStep.id).filter(i => i.estado === "BLOQUEADO");
        if (!toUnblock.length) return;

        const ids = toUnblock.map(i => `'${Provider.esc(i.id)}'`).join(",");
        await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")} SET ESTADO = 'PENDIENTE' WHERE INSTANCIA_ID IN (${ids})`);
        toUnblock.forEach(i => { i.estado = "PENDIENTE"; });
    },

    async maybeCompleteRun() {
        const allDone = this.currentRun.instancias.every(i => i.estado === "COMPLETADO");
        if (allDone && this.currentRun.estado !== "COMPLETADO") {
            await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS")}
                SET ESTADO = 'COMPLETADO', FECHA_FIN = CURRENT_TIMESTAMP() WHERE RUN_ID = '${Provider.esc(this.currentRun.id)}'`);
            this.currentRun.estado = "COMPLETADO";
            UI.toast("¡Ejecución completada!", "success");
        }
    }
};
