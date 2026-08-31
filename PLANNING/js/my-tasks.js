/**
 * ============================================================
 * DRACO PLANNING — MIS TAREAS (task.html)
 * ============================================================
 * Pantalla de EJECUCIÓN para la persona asignada, no de administración:
 * no crea ni elimina workflows/ejecuciones/instancias, solo navega por
 * lo que ya existe y trabaja sobre ello.
 *
 * Navegación: Workflow -> Ejecución -> Paso -> Instancia (si el paso
 * tiene varias, p.ej. una por valor de driver) -> Tareas del paso
 * (bloques y tareas), igual que se definen en el editor de Workflows.
 *
 * "Asignada a mí" = mi email aparece en ASIGNADO (responsable) o en
 * REVISOR de la instancia (WORKFLOWS_RUNS_INSTANCIAS). Ambas son listas
 * separadas por comas; no se resuelven grupos todavía (no existe un
 * directorio de miembros — mismo límite que ya tiene workflow-runs.js).
 *
 * Ejecutar una tarea:
 *   - PARAMETRIZACION (Actualización de tablas) -> se busca el registro
 *     real en ACTUALIZACIONES y se abre exactamente el mismo modal que
 *     el botón ▶ de Administración: TableUpdates.startRun(record).
 *   - FLUJO_MANUAL (Flujo de carga) -> se abre flow_run.html en una
 *     pestaña nueva (mismo patrón ya usado en el resto de la app).
 *   - El resto de tipos (Plantilla Excel/Web, Función, HTML) todavía no
 *     tienen ejecución automática en Draco Planning: se muestra un
 *     popup informativo con la referencia y se completan a mano.
 *
 * Dependencias de módulos ya cargados en task.html: Provider, UI,
 * Workflows (definición + TASK_TYPES) y TableUpdates (para el play).
 */
const MyTasks = {
    project: null,
    myEmail: "",

    overviewRows: [],          // TODAS mis instancias del proyecto (todas las ejecuciones/workflows)
    workflowsOverview: [],     // agregado por workflow: [{id,name,total,pending}]
    currentWorkflowId: null,
    workflowDetail: null,      // Workflows.loadDetail() del workflow seleccionado

    runsForWorkflow: [],       // agregado por ejecución dentro del workflow seleccionado
    currentRunId: null,
    runInstances: [],          // mis instancias de la ejecución seleccionada

    currentStepId: null,
    currentInstanceId: null,

    // Estado de una INSTANCIA — mismas 6 fases que en Administración
    // (workflow-runs.js: WorkflowRuns.ESTADOS), duplicado aquí para que
    // esta pantalla no dependa de cargar ese módulo entero.
    ESTADOS: {
        BLOQUEADO: { label: "Bloqueado", cls: "wf-inst-status--bloqueado" },
        PENDIENTE: { label: "Pendiente", cls: "wf-inst-status--pendiente" },
        EN_CURSO: { label: "En curso", cls: "wf-inst-status--en_curso" },
        SUSPENDIDO: { label: "Suspendido", cls: "wf-inst-status--suspendido" },
        EN_REVISION: { label: "En revisión", cls: "wf-inst-status--en_revision" },
        COMPLETADO: { label: "Completado", cls: "wf-inst-status--completado" }
    },

    RUN_ESTADOS: {
        PENDIENTE: "Pendiente",
        EN_CURSO: "En curso",
        SUSPENDIDO: "Suspendida",
        COMPLETADO: "Completada"
    },

    renderEmpty(container, message) {
        container.innerHTML = `<div class="module-empty">${UI.escapeHtml(message)}</div>`;
    },

    // ------------------------------------------------------------
    // Entrada: se llama cada vez que TaskApp cambia de proyecto
    // ------------------------------------------------------------
    async render(container, project) {
        this.project = project;
        this.currentWorkflowId = null;
        this.currentRunId = null;
        this.currentStepId = null;
        this.currentInstanceId = null;
        container.innerHTML = `<div class="module-empty"><span class="spinner"></span> Buscando tus tareas...</div>`;

        try {
            await this.loadMyEmail();
            await this.loadOverview();
        } catch (err) {
            container.innerHTML = `<div class="module-empty">Error al buscar tus tareas: ${UI.escapeHtml(err.message)}</div>`;
            return;
        }

        this.workflowsOverview = this.groupWorkflows();

        if (!this.workflowsOverview.length) {
            container.innerHTML = `
                <div class="module-empty">
                    No tienes ninguna tarea asignada en el proyecto <strong>${UI.escapeHtml(project.PROYECTO)}</strong>.
                </div>`;
            return;
        }

        container.innerHTML = `
            <nav class="admin-menu mt-menu" id="mtMenu"></nav>
            <div class="admin-content mt-content">
                <div class="flow-part-header"><strong>Pasos del workflow</strong></div>
                <div class="flow-chain-wrap" id="mtChainWrap"></div>
                <div id="mtStepBody"></div>
            </div>`;

        const firstWithPending = this.workflowsOverview.find(w => w.pending > 0) || this.workflowsOverview[0];
        await this.selectWorkflow(firstWithPending.id);
    },

    // ------------------------------------------------------------
    // Menú lateral (mismo look que el menú de Administración): lista de
    // mis workflows y, bajo el que está seleccionado, sus ejecuciones
    // indentadas. Sin estado "en curso"/"pendiente" en texto — solo un
    // circulito con el nº de tareas pendientes cuando hay alguna.
    // ------------------------------------------------------------
    renderMenu() {
        const menu = document.getElementById("mtMenu");
        if (!menu) return;

        menu.innerHTML = this.workflowsOverview.map(w => {
            const active = w.id === this.currentWorkflowId;
            const runsHtml = (active && this.runsForWorkflow && this.runsForWorkflow.length)
                ? `<div class="mt-run-list">${this.runsForWorkflow.map(r => {
                    const runActive = r.id === this.currentRunId;
                    return `
                        <button type="button" class="mt-run-item ${runActive ? "is-active" : ""}" data-mt-run="${r.id}">
                            <span class="mt-run-name">${UI.escapeHtml(r.name)}</span>
                            ${r.pending ? `<span class="mt-run-badge">${r.pending}</span>` : ""}
                        </button>`;
                }).join("")}</div>`
                : "";
            return `
                <button type="button" class="admin-menu-item ${active ? "active" : ""}" data-mt-workflow="${w.id}">
                    <span class="admin-menu-icon">⛓</span>${UI.escapeHtml(w.name)}
                </button>
                ${runsHtml}`;
        }).join("");

        menu.querySelectorAll("[data-mt-workflow]").forEach(btn => {
            btn.addEventListener("click", () => {
                if (btn.dataset.mtWorkflow !== this.currentWorkflowId) this.selectWorkflow(btn.dataset.mtWorkflow);
            });
        });
        menu.querySelectorAll("[data-mt-run]").forEach(btn => {
            btn.addEventListener("click", () => {
                if (btn.dataset.mtRun !== this.currentRunId) this.selectRun(btn.dataset.mtRun);
            });
        });
    },

    async loadMyEmail() {
        // Se cachea en memoria: mientras la pestaña siga abierta no hace
        // falta volver a preguntarle a BigQuery/Snowflake quién soy.
        if (this.myEmail) return;
        try {
            const rows = await Provider.runQuery(`SELECT ${Provider.currentUserExpr()} AS EMAIL`);
            this.myEmail = ((rows[0] && rows[0].EMAIL) || "").trim();
        } catch (err) {
            this.myEmail = "";
        }
        if (!this.myEmail) UI.toast("No se ha podido identificar tu usuario para filtrar tus tareas.", "error");
    },

    // ASIGNADO/REVISOR son listas separadas por comas (emails o grupos).
    // Coincidencia EXACTA de un elemento de la lista, no un simple LIKE
    // (para que "ana@x.com" no case con "susana@x.com").
    emailInList(csv, email) {
        if (!csv || !email) return false;
        return csv.split(",").map(v => v.trim().toLowerCase()).includes(email.toLowerCase());
    },

    async loadOverview() {
        if (!this.myEmail) { this.overviewRows = []; return; }
        const email = Provider.esc(this.myEmail);

        // El LIKE de la consulta es solo un primer filtro en servidor (evita
        // traer todas las instancias del proyecto); la comprobación exacta
        // por elemento de la lista se hace después, en JS (ver emailInList).
        const rows = await Provider.runQuery(`
            SELECT i.INSTANCIA_ID, i.RUN_ID, i.PASO_ID, i.DRIVER_VALOR, i.ASIGNADO, i.REVISOR, i.ESTADO,
                   i.FECHA_PROGRAMADA, i.FECHA_PROGRAMADA_FIN,
                   r.WORKFLOW_ID, r.NOMBRE AS RUN_NOMBRE, r.ESTADO AS RUN_ESTADO,
                   w.WORKFLOW AS WORKFLOW_NOMBRE
            FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")} i
            JOIN ${Provider.qualifyControl("WORKFLOWS_RUNS")} r ON r.RUN_ID = i.RUN_ID
            JOIN ${Provider.qualifyControl("WORKFLOWS")} w ON w.WORKFLOW_ID = r.WORKFLOW_ID
            WHERE r.PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
              AND (i.ASIGNADO LIKE '%${email}%' OR i.REVISOR LIKE '%${email}%')
            ORDER BY r.FECHA_CREACION DESC, i.ORDEN`);

        this.overviewRows = rows
            .filter(r => this.emailInList(r.ASIGNADO, this.myEmail) || this.emailInList(r.REVISOR, this.myEmail))
            .map(r => ({
                id: r.INSTANCIA_ID, runId: r.RUN_ID, pasoId: r.PASO_ID, driverValor: r.DRIVER_VALOR,
                asignado: r.ASIGNADO || "", revisor: r.REVISOR || "", estado: r.ESTADO,
                fechaProgramada: r.FECHA_PROGRAMADA || "", fechaProgramadaFin: r.FECHA_PROGRAMADA_FIN || "",
                workflowId: r.WORKFLOW_ID, runNombre: r.RUN_NOMBRE, runEstado: r.RUN_ESTADO,
                workflowNombre: r.WORKFLOW_NOMBRE
            }));
    },

    groupWorkflows() {
        const map = new Map();
        this.overviewRows.forEach(r => {
            if (!map.has(r.workflowId)) map.set(r.workflowId, { id: r.workflowId, name: r.workflowNombre, total: 0, pending: 0 });
            const g = map.get(r.workflowId);
            g.total++;
            if (r.estado !== "COMPLETADO") g.pending++;
        });
        return Array.from(map.values()).sort((a, b) => (b.pending - a.pending) || a.name.localeCompare(b.name));
    },

    groupRuns(workflowId) {
        const map = new Map();
        this.overviewRows.filter(r => r.workflowId === workflowId).forEach(r => {
            if (!map.has(r.runId)) map.set(r.runId, { id: r.runId, name: r.runNombre, estado: r.runEstado, total: 0, pending: 0 });
            const g = map.get(r.runId);
            g.total++;
            if (r.estado !== "COMPLETADO") g.pending++;
        });
        return Array.from(map.values());
    },

    // ------------------------------------------------------------
    // Workflow -> carga definición (pasos/bloques/tareas) + lista de MIS
    // ejecuciones de ese workflow
    // ------------------------------------------------------------
    async selectWorkflow(id) {
        this.currentWorkflowId = id;
        this.currentRunId = null;
        this.currentStepId = null;
        this.currentInstanceId = null;
        this.runsForWorkflow = [];

        this.renderMenu();
        document.getElementById("mtChainWrap").innerHTML = "";
        const stepBody = document.getElementById("mtStepBody");
        stepBody.innerHTML = `<div class="module-empty"><span class="spinner"></span> Cargando workflow...</div>`;

        try {
            this.workflowDetail = await Workflows.loadDetail(id);
        } catch (err) {
            stepBody.innerHTML = `<div class="module-empty">Error al cargar el workflow: ${UI.escapeHtml(err.message)}</div>`;
            return;
        }
        if (!this.workflowDetail) {
            stepBody.innerHTML = `<div class="module-empty">No se ha podido cargar este workflow.</div>`;
            return;
        }

        this.runsForWorkflow = this.groupRuns(id);
        this.renderMenu();

        if (!this.runsForWorkflow.length) {
            stepBody.innerHTML = `<div class="module-empty">No tienes tareas asignadas en ninguna ejecución de este workflow.</div>`;
            return;
        }

        const pick = this.runsForWorkflow.find(r => r.estado === "EN_CURSO" && r.pending > 0)
            || this.runsForWorkflow.find(r => r.pending > 0)
            || this.runsForWorkflow[0];
        this.selectRun(pick.id);
    },

    // ------------------------------------------------------------
    // Ejecución -> filtra mis instancias de esa ejecución y pinta la
    // cadena de pasos
    // ------------------------------------------------------------
    selectRun(id) {
        this.currentRunId = id;
        this.currentStepId = null;
        this.currentInstanceId = null;
        this.runInstances = this.overviewRows.filter(r => r.runId === id);
        this.renderMenu();
        this.renderChain();
    },

    // ------------------------------------------------------------
    // Cadena de pasos (mismas tarjetas .flow-chain-card que el editor de
    // Workflows, solo con el nombre — sin meta ni badges) pero limitada a
    // los pasos donde tengo al menos una instancia asignada, numerados
    // según su posición REAL en el workflow para no perder el contexto.
    // ------------------------------------------------------------
    renderChain() {
        const wrap = document.getElementById("mtChainWrap");
        const steps = this.workflowDetail.steps.filter(s => !s.isPaso0);
        const relevant = steps
            .map((step, idx) => ({ step, idx }))
            .filter(x => this.runInstances.some(i => i.pasoId === x.step.id));

        if (!relevant.length) {
            wrap.innerHTML = "";
            document.getElementById("mtStepBody").innerHTML = `<div class="module-empty">No tienes tareas asignadas en esta ejecución.</div>`;
            return;
        }

        if (!this.currentStepId || !relevant.some(x => x.step.id === this.currentStepId)) {
            const withPending = relevant.find(x => this.runInstances.some(i => i.pasoId === x.step.id && i.estado !== "COMPLETADO"));
            this.currentStepId = (withPending || relevant[0]).step.id;
        }

        wrap.innerHTML = relevant.map(({ step, idx }, i) => {
            const mine = this.runInstances.filter(inst => inst.pasoId === step.id);
            const allDone = mine.length > 0 && mine.every(inst => inst.estado === "COMPLETADO");
            const selected = step.id === this.currentStepId;
            const card = `
                <div class="flow-chain-card mt-step-card ${selected ? "is-selected" : ""} ${allDone ? "is-done" : ""}" data-mt-step="${step.id}" title="${UI.escapeHtml(step.name)}">
                    <div class="flow-chain-card-name">${allDone ? "✓" : (idx + 1) + "."} ${UI.escapeHtml(step.name)}</div>
                </div>`;
            const arrow = i < relevant.length - 1 ? `<div class="flow-chain-arrow">→</div>` : "";
            return card + arrow;
        }).join("");

        wrap.querySelectorAll("[data-mt-step]").forEach(card => {
            card.addEventListener("click", () => {
                this.currentStepId = card.dataset.mtStep;
                this.currentInstanceId = null;
                this.renderChain();
            });
        });

        this.renderStepBody();
    },

    // ------------------------------------------------------------
    // Cuerpo del paso: selector de instancia (si hay más de una) +
    // panel de estado/acciones + tareas del paso
    // ------------------------------------------------------------
    renderStepBody() {
        const wrap = document.getElementById("mtStepBody");
        const step = this.workflowDetail.steps.find(s => s.id === this.currentStepId);
        if (!step) { wrap.innerHTML = ""; return; }

        const instances = this.runInstances.filter(i => i.pasoId === step.id);
        if (!instances.length) { wrap.innerHTML = `<div class="module-empty">No tienes instancias asignadas en este paso.</div>`; return; }

        if (!this.currentInstanceId || !instances.some(i => i.id === this.currentInstanceId)) {
            const pending = instances.find(i => i.estado !== "COMPLETADO");
            this.currentInstanceId = (pending || instances[0]).id;
        }

        wrap.innerHTML = `
            <div class="flow-part-header mt-step-header">
                <strong>Paso: ${UI.escapeHtml(step.name)}</strong>
                ${step.revision ? `<span class="table-tag">Con revisión</span>` : ""}
            </div>
            ${instances.length > 1 ? this.instanceSelectorHtml(instances) : ""}
            <div id="mtInstancePanel"></div>
            <div id="mtTasksPanel"></div>`;

        if (instances.length > 1) {
            wrap.querySelectorAll("#mtInstanceTabs [data-mt-inst]").forEach(btn => {
                btn.addEventListener("click", () => {
                    this.currentInstanceId = btn.dataset.mtInst;
                    this.renderStepBody();
                });
            });
        }

        const inst = instances.find(i => i.id === this.currentInstanceId);
        this.renderInstancePanel(step, inst);
        this.renderTasks(step, inst);
    },

    instanceSelectorHtml(instances) {
        return `<div class="mt-instance-tabs" id="mtInstanceTabs">
            ${instances.map(i => {
                const label = (i.driverValor !== null && i.driverValor !== undefined && i.driverValor !== "") ? i.driverValor : "Instancia única";
                const selected = i.id === this.currentInstanceId;
                return `<button type="button" class="mt-instance-tab mt-instance-tab--${i.estado.toLowerCase()} ${selected ? "is-selected" : ""}" data-mt-inst="${i.id}">
                            <span class="mt-instance-tab-dot"></span>${UI.escapeHtml(String(label))}
                        </button>`;
            }).join("")}
        </div>`;
    },

    // ------------------------------------------------------------
    // Panel de la instancia seleccionada: estado + controles según mi
    // rol en ella (responsable de ASIGNADO / revisor de REVISOR)
    // ------------------------------------------------------------
    renderInstancePanel(step, inst) {
        const panel = document.getElementById("mtInstancePanel");
        if (!inst) { panel.innerHTML = ""; return; }

        const estadoInfo = this.ESTADOS[inst.estado] || this.ESTADOS.PENDIENTE;
        const isResponsible = this.emailInList(inst.asignado, this.myEmail);
        const isReviewer = this.emailInList(inst.revisor, this.myEmail);

        let controls;
        if (inst.estado === "BLOQUEADO") {
            controls = `<span class="form-hint">Se desbloqueará al completarse el paso anterior.</span>`;
        } else if (inst.estado === "COMPLETADO") {
            controls = `<span class="form-hint">Instancia completada.</span>`;
        } else if (inst.estado === "EN_REVISION") {
            controls = isReviewer
                ? `<button type="button" class="btn btn-primary btn-sm" data-mt-action="approve">Aprobar</button>
                   <button type="button" class="btn btn-secondary btn-sm" data-mt-action="reject">Rechazar</button>`
                : `<span class="form-hint">Pendiente de revisión.</span>`;
        } else if (isResponsible) {
            // Un único botón de cierre — sin play/pausa/restablecer: esta
            // pantalla es de ejecución para la persona asignada, no de
            // administración del workflow.
            controls = step.revision
                ? `<button type="button" class="btn btn-primary btn-sm" data-mt-action="review">Enviar a revisión</button>`
                : `<button type="button" class="btn btn-primary btn-sm" data-mt-action="complete">✓ Completar</button>`;
        } else {
            controls = `<span class="form-hint">Sin acciones disponibles: no eres el responsable de esta instancia.</span>`;
        }

        panel.innerHTML = `
            <div class="mt-instance-panel">
                <div class="mt-instance-panel-top">
                    <span class="table-tag ${estadoInfo.cls}">${estadoInfo.label}</span>
                    ${inst.asignado ? `<span class="mt-instance-meta">Responsable: ${UI.escapeHtml(inst.asignado)}</span>` : ""}
                    ${inst.revisor ? `<span class="mt-instance-meta">Revisor: ${UI.escapeHtml(inst.revisor)}</span>` : ""}
                    ${inst.fechaProgramada ? `<span class="mt-instance-meta">Programado: ${UI.escapeHtml(String(inst.fechaProgramada))}</span>` : ""}
                </div>
                <div class="mt-instance-panel-controls">${controls}</div>
            </div>`;

        panel.querySelectorAll("[data-mt-action]").forEach(btn => {
            btn.addEventListener("click", () => this.transition(inst, step, btn.dataset.mtAction));
        });
    },

    async transition(inst, step, action) {
        const map = { review: "EN_REVISION", approve: "COMPLETADO", reject: "EN_CURSO", complete: "COMPLETADO" };
        const nuevoEstado = map[action];
        if (!nuevoEstado) return;
        const isFinal = nuevoEstado === "COMPLETADO";
        // La instancia puede venir de PENDIENTE directamente (ya no hay
        // paso intermedio "en curso" en esta pantalla), así que si todavía
        // no tenía FECHA_INICIO se registra en el mismo movimiento.
        const setsInicio = inst.estado === "PENDIENTE";
        try {
            await Provider.runQuery(`
                UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                SET ESTADO = '${Provider.esc(nuevoEstado)}'${setsInicio ? ", FECHA_INICIO = CURRENT_TIMESTAMP()" : ""}${isFinal ? ", FECHA_FIN = CURRENT_TIMESTAMP()" : ""}
                WHERE INSTANCIA_ID = '${Provider.esc(inst.id)}'`);
            inst.estado = nuevoEstado;
            if (isFinal) {
                await this.cascadeUnblock(inst.runId, step.id);
                await this.maybeCompleteRun(inst.runId);
            }
            this.workflowsOverview = this.groupWorkflows();
            this.runsForWorkflow = this.groupRuns(this.currentWorkflowId);
            this.renderMenu();
            this.renderChain();
        } catch (err) {
            UI.toast("Error al actualizar el estado: " + err.message, "error");
        }
    },

    // Al completarse la última instancia pendiente de un paso, desbloquea
    // las instancias BLOQUEADO del siguiente paso si su Inicio es "al
    // completar el paso anterior" — mismo criterio que Administración,
    // pero consultando directamente en BD (esta pantalla no tiene cargadas
    // las instancias de OTRAS personas para ese paso).
    async cascadeUnblock(runId, pasoId) {
        const steps = this.workflowDetail.steps.filter(s => !s.isPaso0);
        const idx = steps.findIndex(s => s.id === pasoId);
        if (idx === -1 || idx === steps.length - 1) return;
        try {
            const pending = await Provider.runQuery(`
                SELECT COUNT(*) AS N FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                WHERE RUN_ID = '${Provider.esc(runId)}' AND PASO_ID = '${Provider.esc(pasoId)}' AND ESTADO != 'COMPLETADO'`);
            if (Number((pending[0] && pending[0].N) || 0) > 0) return;

            const nextStep = steps[idx + 1];
            if (nextStep.inicio.tipo !== "PASO_ANTERIOR") return;

            await Provider.runQuery(`
                UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")} SET ESTADO = 'PENDIENTE'
                WHERE RUN_ID = '${Provider.esc(runId)}' AND PASO_ID = '${Provider.esc(nextStep.id)}' AND ESTADO = 'BLOQUEADO'`);
        } catch (err) {
            console.error("No se pudo desbloquear el siguiente paso:", err);
        }
    },

    async maybeCompleteRun(runId) {
        try {
            const pending = await Provider.runQuery(`
                SELECT COUNT(*) AS N FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                WHERE RUN_ID = '${Provider.esc(runId)}' AND ESTADO != 'COMPLETADO'`);
            if (Number((pending[0] && pending[0].N) || 0) === 0) {
                await Provider.runQuery(`
                    UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS")} SET ESTADO = 'COMPLETADO', FECHA_FIN = CURRENT_TIMESTAMP()
                    WHERE RUN_ID = '${Provider.esc(runId)}' AND ESTADO != 'COMPLETADO'`);
                const g = this.runsForWorkflow.find(r => r.id === runId);
                if (g) g.estado = "COMPLETADO";
                UI.toast("¡Ejecución completada!", "success");
            }
        } catch (err) {
            console.error("No se pudo comprobar si la ejecución se completó:", err);
        }
    },

    // ------------------------------------------------------------
    // Tareas del paso (bloques + tareas), igual que se definen en el
    // editor de Workflows
    // ------------------------------------------------------------
    renderTasks(step) {
        const wrap = document.getElementById("mtTasksPanel");
        const bloques = step.bloques.filter(b => b.tareas.length);
        if (!bloques.length) {
            wrap.innerHTML = `<div class="module-empty module-empty--inline">Este paso no tiene tareas definidas.</div>`;
            return;
        }

        wrap.innerHTML = `
            <div class="mt-tasks-title">Tareas del paso</div>
            ${bloques.map(b => `
                <div class="mt-task-block">
                    <div class="mt-task-block-title">${UI.escapeHtml(b.titulo)}</div>
                    ${b.tareas.map(t => this.taskCardHtml(t)).join("")}
                </div>`).join("")}`;

        wrap.querySelectorAll("[data-mt-task]").forEach(card => {
            card.addEventListener("click", () => {
                const tarea = this.findTask(step, card.dataset.mtTask);
                if (tarea) this.launchTask(tarea);
            });
        });
    },

    findTask(step, taskId) {
        for (const b of step.bloques) {
            const t = b.tareas.find(t => t.id === taskId);
            if (t) return t;
        }
        return null;
    },

    taskCardHtml(task) {
        const typeInfo = Workflows.TASK_TYPES[task.tipo] || { label: task.tipo, icon: "•" };
        const executable = task.tipo === "PARAMETRIZACION" || task.tipo === "FLUJO_MANUAL";
        return `
            <div class="wf-task-card mt-task-card" data-mt-task="${task.id}">
                <div class="wf-task-card-header">
                    <span class="wf-task-icon">${typeInfo.icon}</span>
                    <div class="wf-task-title-wrap">
                        <div class="wf-task-name">${UI.escapeHtml(task.nombre || typeInfo.label)}</div>
                        <div class="wf-task-meta">
                            <span class="table-tag">${UI.escapeHtml(typeInfo.label)}</span>
                            ${task.refNombre && task.refNombre !== task.nombre ? `<span class="wf-task-meta-ref">· ${UI.escapeHtml(task.refNombre)}</span>` : ""}
                        </div>
                    </div>
                    <span class="mt-task-play" title="${executable ? "Ejecutar" : "Ver referencia"}">${executable ? "▶" : "ℹ"}</span>
                </div>
                ${task.descripcion ? `<p class="wf-task-desc">${UI.escapeHtml(task.descripcion)}</p>` : ""}
            </div>`;
    },

    // ------------------------------------------------------------
    // Ejecutar una tarea: reutiliza el motor real de cada catálogo en
    // vez de reinventar la ejecución
    // ------------------------------------------------------------
    async launchTask(task) {
        if (task.tipo === "PARAMETRIZACION") {
            if (!task.refId) { UI.toast("Esta tarea no tiene una actualización de tablas asignada.", "error"); return; }
            try {
                const rows = await Provider.runQuery(`
                    SELECT ACTUALIZACION_ID, NOMBRE, DESCRIPCION, TABLA, VARIABLES_JSON, CAMPOS_JSON
                    FROM ${Provider.qualifyControl("ACTUALIZACIONES")}
                    WHERE ACTUALIZACION_ID = '${Provider.esc(task.refId)}'`);
                if (!rows.length) { UI.toast("No se ha encontrado la actualización de tablas de referencia.", "error"); return; }
                // Exactamente el mismo modal que abre el botón ▶ en
                // Administración > Actualización de tablas.
                TableUpdates.project = this.project;
                await TableUpdates.startRun(rows[0]);
            } catch (err) {
                UI.toast("Error al abrir la actualización de tablas: " + err.message, "error");
            }
            return;
        }

        if (task.tipo === "FLUJO_MANUAL") {
            if (!task.refId) { UI.toast("Esta tarea no tiene un flujo de carga asignado.", "error"); return; }
            window.open(`flow_run.html?flujo_id=${encodeURIComponent(task.refId)}`, "_blank");
            return;
        }

        this.openInfoModal(task);
    },

    // Tipos sin ejecución automática todavía (Plantilla Excel/Web,
    // Función, HTML): solo se muestra su referencia a modo de recordatorio.
    openInfoModal(task) {
        const typeInfo = Workflows.TASK_TYPES[task.tipo] || { label: task.tipo, icon: "•" };
        let overlay = document.getElementById("mtTaskInfoModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "mtTaskInfoModal";
            document.body.appendChild(overlay);
        }
        overlay.innerHTML = `
            <div class="modal-box">
                <div class="modal-header">
                    <h3>${typeInfo.icon} ${UI.escapeHtml(task.nombre || typeInfo.label)}</h3>
                    <button class="modal-close" id="mtTaskInfoClose">&times;</button>
                </div>
                <div class="modal-body">
                    <p class="form-hint">${UI.escapeHtml(typeInfo.label)}${task.refNombre ? " · " + UI.escapeHtml(task.refNombre) : ""}</p>
                    ${task.descripcion ? `<p>${UI.escapeHtml(task.descripcion)}</p>` : ""}
                    <p class="form-hint">Esta tarea todavía no tiene una ejecución automática configurada en Draco Planning. Complétala manualmente y márcala cuando termines.</p>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-primary" id="mtTaskInfoOk">Entendido</button>
                </div>
            </div>`;
        overlay.classList.add("visible");
        const close = () => overlay.classList.remove("visible");
        document.getElementById("mtTaskInfoClose").addEventListener("click", close);
        document.getElementById("mtTaskInfoOk").addEventListener("click", close);
    }
};
