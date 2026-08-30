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
        BLOQUEADO: { label: "Bloqueado", cls: "wf-inst-status--bloqueado" },
        PENDIENTE: { label: "Pendiente", cls: "wf-inst-status--pendiente" },
        EN_CURSO: { label: "En curso", cls: "wf-inst-status--en_curso" },
        SUSPENDIDO: { label: "Suspendido", cls: "wf-inst-status--suspendido" },
        EN_REVISION: { label: "En revisión", cls: "wf-inst-status--en_revision" },
        COMPLETADO: { label: "Completado", cls: "wf-inst-status--completado" }
    },

    // Pestaña de asignación (Responsable/Revisor) seleccionada por
    // instancia, solo en memoria (no persiste): por defecto Responsable.
    _assigneeTab: {},
    // Caché de "código -> texto descriptivo" del driver por dimensión,
    // para no repetir la consulta al cambiar de paso. Ver fetchDriverLabels.
    _driverLabelCache: {},

    // Estados de la CABECERA de una ejecución (WORKFLOWS_RUNS.ESTADO), usados
    // en el listado. Independientes de this.ESTADOS, que son los de cada
    // instancia de paso dentro del detalle.
    RUN_ESTADOS: {
        PENDIENTE: { label: "Pendiente", cls: "run-status-pendiente" },
        EN_CURSO: { label: "En curso", cls: "run-status-en_curso" },
        SUSPENDIDO: { label: "Suspendido", cls: "run-status-suspendido" },
        COMPLETADO: { label: "Completada", cls: "run-status-completado" }
    },

    // Estado resumido de un PASO (no de cada instancia) que se muestra en
    // la cadena de arriba: empieza Pendiente; en cuanto se pulsa play en
    // alguna instancia pasa a Activo; si se pulsa stop pasa a Suspendido;
    // si la propia ejecución está suspendida, el paso se ve Suspendido
    // pase lo que pase con sus instancias; el resto de casos (bloqueado,
    // en revisión, completado, mezcla de estados...) se agrupan como "En
    // construcción". Ver stepRunEstado().
    STEP_ESTADOS: {
        PENDIENTE: { label: "Pendiente", cls: "wf-step-status--pendiente" },
        ACTIVO: { label: "Activo", cls: "wf-step-status--activo" },
        SUSPENDIDO: { label: "Suspendido", cls: "wf-step-status--suspendido" },
        EN_CONSTRUCCION: { label: "En construcción", cls: "wf-step-status--construccion" }
    },

    stepRunEstado(instances) {
        if (this.currentRun && this.currentRun.estado === "SUSPENDIDO") return "SUSPENDIDO";
        if (!instances.length) return "PENDIENTE";
        if (instances.some(i => i.estado === "SUSPENDIDO")) return "SUSPENDIDO";
        if (instances.some(i => i.estado === "EN_CURSO")) return "ACTIVO";
        if (instances.every(i => i.estado === "PENDIENTE")) return "PENDIENTE";
        return "EN_CONSTRUCCION";
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
        overlay.querySelector("#wfRunsModalClose").addEventListener("click", () => this.handleCloseClick());
        this.overlay = overlay;
    },

    setModalTitle(title, subtitle) {
        document.getElementById("wfRunsModalTitle").textContent = title;
        document.getElementById("wfRunsModalSubtitle").textContent = subtitle || "";
    },

    // La X del popup no cierra siempre del todo: desde el detalle de una
    // ejecución, primero vuelve al listado (como el link que había antes);
    // solo cierra el popup si ya estábamos en el listado.
    handleCloseClick() {
        if (this.view === "detail") {
            this.showRunsList();
        } else {
            this.closeModal();
        }
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

            // Variables globales (Paso 0) de cada ejecución, para mostrarlas
            // en el listado en lugar de la fecha de creación. Se cargan en
            // una segunda consulta y se agrupan aquí en JS (evita depender
            // de funciones de agregación de strings distintas entre
            // BigQuery/Snowflake).
            if (this.runs.length) {
                const varRows = await Provider.runQuery(`
                    SELECT RUN_ID, NOMBRE, VALOR
                    FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")}
                    WHERE RUN_ID = INSTANCIA_ID
                      AND RUN_ID IN (${this.runs.map(r => `'${Provider.esc(r.RUN_ID)}'`).join(",")})`);
                const varsByRun = {};
                varRows.forEach(v => {
                    (varsByRun[v.RUN_ID] = varsByRun[v.RUN_ID] || []).push(`${v.NOMBRE}=${v.VALOR}`);
                });
                this.runs.forEach(r => { r.VARIABLES_TXT = (varsByRun[r.RUN_ID] || []).join(", "); });
            }
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
                            <thead><tr><th>Ejecución</th><th>Estado</th><th>Variables</th><th></th></tr></thead>
                            <tbody>
                                ${this.runs.map(r => {
                                    const estadoInfo = this.RUN_ESTADOS[r.ESTADO] || { label: r.ESTADO, cls: "" };
                                    const isCompletado = r.ESTADO === "COMPLETADO";
                                    return `
                                    <tr>
                                        <td><strong>${UI.escapeHtml(r.NOMBRE)}</strong></td>
                                        <td><span class="table-tag ${estadoInfo.cls}">${UI.escapeHtml(estadoInfo.label)}</span></td>
                                        <td>${r.VARIABLES_TXT
                                            ? `<span class="run-vars-summary" title="${UI.escapeHtml(r.VARIABLES_TXT)}">${UI.escapeHtml(r.VARIABLES_TXT)}</span>`
                                            : `<span class="run-vars-empty">—</span>`}</td>
                                        <td>
                                            <div class="row-actions">
                                                <button data-mod-run="${r.RUN_ID}" title="Modificar">✎</button>
                                                <button data-exec-run="${r.RUN_ID}" title="Ejecutar" ${r.ESTADO === "EN_CURSO" || isCompletado ? "disabled" : ""}>▶</button>
                                                <button data-susp-run="${r.RUN_ID}" title="Suspender" ${r.ESTADO === "SUSPENDIDO" || isCompletado ? "disabled" : ""}>⏸</button>
                                                <button data-del-run="${r.RUN_ID}" class="danger" title="Eliminar">🗑</button>
                                            </div>
                                        </td>
                                    </tr>`;
                                }).join("")}
                            </tbody>
                        </table>
                    </div>` : ""}
            </div>`;

        const newBtn = document.getElementById("btnNewRun");
        if (newBtn) newBtn.addEventListener("click", () => this.createRun());
        body.querySelectorAll("[data-mod-run]").forEach(btn =>
            btn.addEventListener("click", () => this.openRun(btn.dataset.modRun)));
        body.querySelectorAll("[data-exec-run]").forEach(btn =>
            btn.addEventListener("click", () => this.setRunEstado(btn.dataset.execRun, "EN_CURSO")));
        body.querySelectorAll("[data-susp-run]").forEach(btn =>
            btn.addEventListener("click", () => this.setRunEstado(btn.dataset.suspRun, "SUSPENDIDO")));
        body.querySelectorAll("[data-del-run]").forEach(btn =>
            btn.addEventListener("click", () => this.removeRun(btn.dataset.delRun)));
    },

    // ------------------------------------------------------------
    // Ejecutar / Suspender desde el listado: cambia el estado de la
    // CABECERA de la ejecución y lo persiste, para que se recupere tal
    // cual la próxima vez que se abra el listado.
    // ------------------------------------------------------------
    async setRunEstado(runId, estado) {
        const run = this.runs.find(r => r.RUN_ID === runId);
        if (!run) return;
        const estadoAnterior = run.ESTADO;
        run.ESTADO = estado; // optimista, para que la tabla reaccione al instante
        this.showRunsList();
        try {
            await Provider.runQuery(`
                UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS")}
                SET ESTADO = '${Provider.esc(estado)}'
                WHERE RUN_ID = '${Provider.esc(runId)}'`);
            const label = (this.RUN_ESTADOS[estado] || {}).label || estado;
            UI.toast(`Ejecución "${run.NOMBRE}" → ${label}.`, "success");
        } catch (err) {
            run.ESTADO = estadoAnterior;
            this.showRunsList();
            UI.toast("Error al actualizar el estado de la ejecución: " + err.message, "error");
        }
    },

    async removeRun(runId) {
        const run = this.runs.find(r => r.RUN_ID === runId);
        if (!run) return;
        const ok = await UI.confirm("Eliminar ejecución", `Se eliminará la ejecución <strong>${UI.escapeHtml(run.NOMBRE)}</strong> y todo su progreso.`);
        if (!ok) return;
        try {
            await Promise.all([
                Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")} WHERE RUN_ID = '${Provider.esc(runId)}'`),
                Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")} WHERE RUN_ID = '${Provider.esc(runId)}'`),
                Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_RUNS")} WHERE RUN_ID = '${Provider.esc(runId)}'`)
            ]);
            await this.loadRuns();
            this.showRunsList();
            UI.toast(`Ejecución "${run.NOMBRE}" eliminada.`, "success");
        } catch (err) {
            UI.toast("Error al eliminar la ejecución: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Popup compacto: nombre + descripción + variables del Paso 0
    // ------------------------------------------------------------
    async createRun() {
        const result = await this.openRunFormModal("create");
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

                let fechaProgramadaFin = "";
                if (step.fin.tipo === "FECHA") fechaProgramadaFin = step.fin.fecha || "";

                valores.forEach(v => {
                    instancias.push({
                        id: Provider.newId(), pasoId: step.id, orden: idx,
                        driverValor: v, asignado: "", estado: estadoInicial,
                        fechaProgramada, fechaProgramadaFin
                    });
                });
            }

            // Las 3 tablas (cabecera, instancias, variables globales) son
            // inserts independientes con los VALUES ya construidos en
            // memoria: se lanzan a la vez en vez de uno detrás de otro.
            const runInserts = [
                Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_RUNS")}
                    (RUN_ID, WORKFLOW_ID, PROYECTO_ID, NOMBRE, DESCRIPCION, ESTADO, USUARIO, FECHA_CREACION, FECHA_FIN)
                    VALUES ('${Provider.esc(runId)}', '${Provider.esc(this.workflow.id)}', '${Provider.esc(pid)}', '${Provider.esc(result.name)}', '${Provider.esc(result.description)}', 'EN_CURSO',
                            ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), NULL)`)
            ];

            if (instancias.length) {
                const vals = instancias.map(i => `(
                    '${Provider.esc(runId)}', '${Provider.esc(i.pasoId)}', '${Provider.esc(i.id)}', ${i.orden},
                    ${i.driverValor !== null ? `'${Provider.esc(i.driverValor)}'` : "NULL"},
                    '${Provider.esc(i.asignado)}', '${Provider.esc(i.estado)}', '${Provider.esc(i.fechaProgramada)}', '${Provider.esc(i.fechaProgramadaFin)}', NULL, NULL
                )`).join(",\n");
                runInserts.push(Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                    (RUN_ID, PASO_ID, INSTANCIA_ID, ORDEN, DRIVER_VALOR, ASIGNADO, ESTADO, FECHA_PROGRAMADA, FECHA_PROGRAMADA_FIN, FECHA_INICIO, FECHA_FIN)
                    VALUES ${vals}`));
            }

            const globalVarEntries = Object.entries(result.variables || {});
            if (globalVarEntries.length) {
                const varVals = globalVarEntries.map(([name, value]) =>
                    `('${Provider.esc(runId)}', '${Provider.esc(runId)}', '${Provider.esc(name)}', '${Provider.esc(value)}')`);
                runInserts.push(Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")}
                    (RUN_ID, INSTANCIA_ID, NOMBRE, VALOR) VALUES ${varVals.join(",\n")}`));
            }

            await Promise.all(runInserts);

            await this.loadRuns();
            await this.openRun(runId);
            UI.toast(`Ejecución "${result.name}" creada con ${instancias.length} instancia(s).`, "success");
        } catch (err) {
            UI.toast("Error al crear la ejecución: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Popup compacto: nombre + descripción + variables del Paso 0.
    // Se usa tanto para "Nueva ejecución" (mode = "create") como para
    // modificar una ya existente (mode = "edit", con `existingRun`
    // precargando nombre/descripción/variables).
    // ------------------------------------------------------------
    openRunFormModal(mode, existingRun = null) {
        return new Promise((resolve) => {
            const paso0 = this.paso0Step();
            const variables = paso0 ? paso0.variables : [];
            const isEdit = mode === "edit";

            let overlay = document.getElementById("wfNewRunModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wfNewRunModal";
                document.body.appendChild(overlay);
            }

            const inputType = (t) => ({ INTEGER: "number", FLOAT: "number", NUMERIC: "number", DATE: "date", DATETIME: "datetime-local", TIMESTAMP: "datetime-local" }[t] || "text");
            const existingVars = (isEdit && existingRun && existingRun.variables) || {};

            overlay.innerHTML = `
                <div class="modal-box modal-wide">
                    <div class="modal-header">
                        <div>
                            <h3>${isEdit ? "Modificar ejecución" : "Nueva ejecución"}</h3>
                            <span class="modal-subtitle">${UI.escapeHtml(this.workflow.name)}</span>
                        </div>
                        <button class="modal-close" id="wfNewRunClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Nombre de la ejecución</label>
                            <input type="text" id="wfNewRunName" placeholder="Ej. Cierre Enero 2026" value="${isEdit ? UI.escapeHtml(existingRun.name) : ""}">
                        </div>
                        <div class="form-group">
                            <label>Descripción</label>
                            <textarea id="wfNewRunDescription" rows="2" placeholder="Notas u observaciones sobre esta ejecución (opcional)">${isEdit ? UI.escapeHtml(existingRun.description || "") : ""}</textarea>
                        </div>
                        ${variables.length ? `
                            <div class="flow-step-group-title" style="margin-top:6px;">Variables del workflow</div>
                            <div class="wf-newrun-vars">
                                ${variables.map(v => `
                                    <div class="form-group">
                                        <label>${UI.escapeHtml(v.label)}</label>
                                        <input type="${inputType(v.type)}" data-run-var="${UI.escapeHtml(v.name)}" placeholder="${UI.escapeHtml(v.label)}" value="${UI.escapeHtml(existingVars[v.name] || "")}">
                                    </div>`).join("")}
                            </div>` : `<p class="form-hint">Este workflow no tiene variables definidas en el Paso 0.</p>`}
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="wfNewRunCancel">Cancelar</button>
                        <button class="btn btn-primary" id="wfNewRunCreate">${isEdit ? "Guardar cambios" : "Crear ejecución"}</button>
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
                const description = overlay.querySelector("#wfNewRunDescription").value.trim();
                const values = {};
                overlay.querySelectorAll("[data-run-var]").forEach(input => { values[input.dataset.runVar] = input.value; });
                cleanup({ name, description, variables: values });
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
            const [headRows, instRows, varRows] = await Promise.all([
                Provider.runQuery(`SELECT * FROM ${Provider.qualifyControl("WORKFLOWS_RUNS")} WHERE RUN_ID = '${Provider.esc(runId)}'`),
                Provider.runQuery(`
                    SELECT * FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                    WHERE RUN_ID = '${Provider.esc(runId)}' ORDER BY ORDEN`),
                Provider.runQuery(`
                    SELECT INSTANCIA_ID, NOMBRE, VALOR FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")}
                    WHERE RUN_ID = '${Provider.esc(runId)}'`)
            ]);
            const head = headRows[0];
            if (!head) { UI.toast("No se ha encontrado la ejecución.", "error"); this.showRunsList(); return; }

            const varsByInst = {};
            const globalVars = {};
            varRows.forEach(v => {
                if (v.INSTANCIA_ID === runId) { globalVars[v.NOMBRE] = v.VALOR; return; }
                (varsByInst[v.INSTANCIA_ID] = varsByInst[v.INSTANCIA_ID] || {})[v.NOMBRE] = v.VALOR;
            });

            this.currentRun = {
                id: head.RUN_ID, name: head.NOMBRE, description: head.DESCRIPCION || "", estado: head.ESTADO,
                variables: globalVars,
                instancias: instRows.map(i => ({
                    id: i.INSTANCIA_ID, pasoId: i.PASO_ID, orden: parseInt(i.ORDEN || "0", 10),
                    driverValor: i.DRIVER_VALOR || null, asignado: i.ASIGNADO || "", revisor: i.REVISOR || "",
                    estado: i.ESTADO, fechaProgramada: i.FECHA_PROGRAMADA || "", fechaProgramadaFin: i.FECHA_PROGRAMADA_FIN || "",
                    variables: varsByInst[i.INSTANCIA_ID] || {}
                }))
            };

            // Autocuración: si el workflow ganó pasos nuevos DESPUÉS de que
            // se creara esta ejecución, esos pasos no tienen instancias en
            // WORKFLOWS_RUNS_INSTANCIAS y antes se quedaban vacíos sin más
            // (por eso "el paso 2 y el 3 no salen nada"). Se generan aquí
            // igual que en createRun() y se insertan, para que la
            // ejecución quede al día con la definición actual del workflow.
            await this.ensureStepInstances();

            const steps = this.execSteps();
            this.selectedRunStepId = steps.length ? steps[0].id : null;
            this.view = "detail";
            this.renderRunDetail();
        } catch (err) {
            UI.toast("Error al abrir la ejecución: " + err.message, "error");
            this.showRunsList();
        }
    },

    // Genera e inserta instancias para los pasos ejecutables del workflow
    // que todavía no tengan ninguna en esta ejecución (ver comentario en
    // openRun). Usa la misma lógica que createRun() para decidir los
    // valores del driver y el estado inicial de cada instancia.
    async ensureStepInstances() {
        const steps = this.execSteps();
        const missingSteps = steps.filter(s => this.instancesForStep(s.id).length === 0);
        if (!missingSteps.length) return;

        const runId = this.currentRun.id;
        const newInstances = [];

        for (const step of missingSteps) {
            const idx = steps.indexOf(step);
            let valores = [null];
            if (step.driver.dimensionId) {
                if (step.driver.valores.length) {
                    valores = step.driver.valores;
                } else {
                    const dim = Workflows.dimensionById(step.driver.dimensionId);
                    if (dim) {
                        try {
                            const keyCol = Provider.toIdentifier(dim.DIMENSION);
                            const fullTable = Provider.qualify(this.project.DATASET, dim.TABLA);
                            const rows = await Provider.runQuery(`SELECT DISTINCT ${keyCol} AS V FROM ${fullTable} ORDER BY ${keyCol} LIMIT 5000`);
                            valores = rows.length ? rows.map(r => String(r.V)) : [null];
                        } catch (e) { valores = [null]; }
                    }
                }
            }

            let estadoInicial = "PENDIENTE";
            let fechaProgramada = "";
            if (step.inicio.tipo === "PASO_ANTERIOR" && idx > 0) {
                const prevInstances = this.instancesForStep(steps[idx - 1].id);
                const prevDone = prevInstances.length > 0 && prevInstances.every(i => i.estado === "COMPLETADO");
                estadoInicial = prevDone ? "PENDIENTE" : "BLOQUEADO";
            }
            if (step.inicio.tipo === "FECHA") {
                fechaProgramada = step.inicio.fecha || "";
                estadoInicial = "PENDIENTE";
            }
            let fechaProgramadaFin = "";
            if (step.fin.tipo === "FECHA") fechaProgramadaFin = step.fin.fecha || "";

            valores.forEach(v => {
                newInstances.push({
                    id: Provider.newId(), pasoId: step.id, orden: idx,
                    driverValor: v, asignado: "", estado: estadoInicial, fechaProgramada, fechaProgramadaFin
                });
            });
        }

        if (!newInstances.length) return;

        try {
            const vals = newInstances.map(i => `(
                '${Provider.esc(runId)}', '${Provider.esc(i.pasoId)}', '${Provider.esc(i.id)}', ${i.orden},
                ${i.driverValor !== null ? `'${Provider.esc(i.driverValor)}'` : "NULL"},
                '${Provider.esc(i.asignado)}', '${Provider.esc(i.estado)}', '${Provider.esc(i.fechaProgramada)}', '${Provider.esc(i.fechaProgramadaFin)}', NULL, NULL
            )`).join(",\n");
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                (RUN_ID, PASO_ID, INSTANCIA_ID, ORDEN, DRIVER_VALOR, ASIGNADO, ESTADO, FECHA_PROGRAMADA, FECHA_PROGRAMADA_FIN, FECHA_INICIO, FECHA_FIN)
                VALUES ${vals}`);
            newInstances.forEach(i => this.currentRun.instancias.push({
                id: i.id, pasoId: i.pasoId, orden: i.orden, driverValor: i.driverValor,
                asignado: "", revisor: "", estado: i.estado, fechaProgramada: i.fechaProgramada, fechaProgramadaFin: i.fechaProgramadaFin, variables: {}
            }));
        } catch (err) {
            UI.toast("No se han podido preparar los pasos nuevos de esta ejecución: " + err.message, "error");
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
        const globalVarEntries = Object.entries(run.variables || {});

        this.setModalTitle(run.name, this.workflow.name);

        const body = document.getElementById("wfRunsModalBody");
        body.innerHTML = `
            <div class="wf-run-summary wf-run-summary--compact">
                ${globalVarEntries.length ? `
                    <div class="chip-row" style="margin:0;">
                        ${globalVarEntries.map(([k, v]) => `<span class="hier-chip">${UI.escapeHtml(k)} = ${UI.escapeHtml(v || "—")}</span>`).join("")}
                    </div>` : "<span></span>"}
                <button type="button" class="btn-icon" id="btnEditRunInfo" title="Modificar nombre, descripción y variables">✎</button>
            </div>

            <div class="wf-run-steps" id="wfRunChainWrap"></div>
            <div id="wfRunStepBody"></div>`;

        document.getElementById("btnEditRunInfo").addEventListener("click", () => this.editRunInfo());

        this.renderRunChain();
        this.renderRunStepBody();
    },

    // ------------------------------------------------------------
    // Modificar nombre, descripción y variables globales (Paso 0) de una
    // ejecución ya creada, reutilizando el mismo popup de "Nueva
    // ejecución". No toca las instancias ni sus estados.
    // ------------------------------------------------------------
    async editRunInfo() {
        const run = this.currentRun;
        const result = await this.openRunFormModal("edit", run);
        if (!result) return;

        try {
            await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS")}
                SET NOMBRE = '${Provider.esc(result.name)}', DESCRIPCION = '${Provider.esc(result.description)}'
                WHERE RUN_ID = '${Provider.esc(run.id)}'`);

            // Variables globales: se sustituyen todas (borrar + insertar)
            // en vez de hacer un UPDATE fila a fila, igual que se crean en
            // createRun() — más simple y cubre añadir/quitar variables si
            // el Paso 0 ha cambiado entretanto.
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")}
                WHERE RUN_ID = '${Provider.esc(run.id)}' AND INSTANCIA_ID = '${Provider.esc(run.id)}'`);
            const globalVarEntries = Object.entries(result.variables || {});
            if (globalVarEntries.length) {
                const varVals = globalVarEntries.map(([name, value]) =>
                    `('${Provider.esc(run.id)}', '${Provider.esc(run.id)}', '${Provider.esc(name)}', '${Provider.esc(value)}')`);
                await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_RUNS_VARIABLES")}
                    (RUN_ID, INSTANCIA_ID, NOMBRE, VALOR) VALUES ${varVals.join(",\n")}`);
            }

            run.name = result.name;
            run.description = result.description;
            run.variables = result.variables || {};
            const listEntry = this.runs.find(r => r.RUN_ID === run.id);
            if (listEntry) listEntry.NOMBRE = result.name;

            this.renderRunDetail();
            UI.toast("Ejecución actualizada.", "success");
        } catch (err) {
            UI.toast("Error al modificar la ejecución: " + err.message, "error");
        }
    },

    renderRunChain() {
        const steps = this.execSteps();
        const chainWrap = document.getElementById("wfRunChainWrap");
        if (!chainWrap) return;

        chainWrap.innerHTML = steps.map((step, idx) => {
            const instances = this.instancesForStep(step.id);
            const assignedCount = instances.filter(i => i.asignado && i.asignado.trim()).length;
            const fullyAssigned = instances.length > 0 && assignedCount === instances.length;
            const stepDone = instances.length > 0 && instances.every(i => i.estado === "COMPLETADO");
            const selected = step.id === this.selectedRunStepId;
            const stepEstadoInfo = this.STEP_ESTADOS[this.stepRunEstado(instances)];
            const card = `
                <div class="wf-run-step-tab ${selected ? "is-selected" : ""} ${stepDone ? "is-done" : ""}" data-run-step="${step.id}">
                    <span class="wf-run-step-num">${stepDone ? "✓" : (idx + 1) + "."}</span>
                    <div class="wf-run-step-info">
                        <span class="wf-run-step-name">${UI.escapeHtml(step.name)}</span>
                        <span class="table-tag ${stepEstadoInfo.cls}">${stepEstadoInfo.label}</span>
                        <span class="wf-run-chain-badge ${fullyAssigned ? "wf-run-chain-badge--assigned" : "wf-run-chain-badge--unassigned"}">
                            <span class="wf-run-chain-badge-dot">${fullyAssigned ? "✓" : "⚠"}</span>${assignedCount}/${instances.length} asignada${instances.length === 1 ? "" : "s"}
                        </span>
                    </div>
                </div>`;
            const connector = idx < steps.length - 1 ? `<div class="wf-run-step-connector"></div>` : "";
            return card + connector;
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
        const bloques = step.bloques || [];
        const taskCount = bloques.reduce((n, b) => n + (b.tareas || []).length, 0);

        wrap.innerHTML = `
            <div class="flow-screen-block flow-screen-block--frame" style="margin-bottom:0;">
                <div class="flow-frame-header">
                    <span><strong>${UI.escapeHtml(step.name)}</strong></span>
                    <span class="load-fn-toolbar-spacer"></span>
                    ${step.revision ? `<span class="table-tag">Con revisión</span>` : ""}
                    ${dim ? `<span class="table-tag">Driver: ${UI.escapeHtml(dim.DIMENSION)}</span>` : ""}
                    ${taskCount ? `<button type="button" class="btn btn-secondary btn-sm" id="btnToggleTasks">Tareas (${taskCount})</button>` : ""}
                </div>
                ${this.stepDatesBlockHtml(step, instances)}
                <div class="flow-frame-vars" id="wfTasksPanel" style="display:none;">
                    ${bloques.map(b => `
                        <p class="form-hint"><strong>${UI.escapeHtml(b.titulo)}</strong></p>
                        ${(b.tareas || []).map(t => this.taskRowHtml(t)).join("")}`).join("") || `<p class="form-hint">Sin tareas.</p>`}
                </div>
                <div class="wf-instance-grid">
                    ${instances.length ? instances.map(i => this.instanceCardHtml(step, i)).join("") : `<div class="module-empty">Este paso todavía no tiene instancias.</div>`}
                </div>
            </div>`;

        const toggleBtn = document.getElementById("btnToggleTasks");
        if (toggleBtn) toggleBtn.addEventListener("click", () => {
            const panel = document.getElementById("wfTasksPanel");
            panel.style.display = panel.style.display === "none" ? "block" : "none";
        });

        this.bindRunEvents(wrap);
        this.bindStepDatesEvents(step, instances);
        this.fetchDriverLabels(step, instances);
    },

    // Cuando el Inicio y/o la Finalización del paso están definidos como
    // "Fecha concreta", la fecha ya no se fija en la definición del
    // workflow (ver Workflows.renderPropiedadesTab): se pide aquí, una
    // única vez por paso, y se aplica a todas sus instancias. Se muestra
    // un bloque con los selectores que correspondan, antes del bloque de
    // tareas.
    stepDatesBlockHtml(step, instances) {
        const needsInicio = step.inicio.tipo === "FECHA";
        const needsFin = step.fin.tipo === "FECHA";
        if (!needsInicio && !needsFin) return "";

        const fechaInicio = instances.length ? (instances[0].fechaProgramada || "") : "";
        const fechaFin = instances.length ? (instances[0].fechaProgramadaFin || "") : "";

        return `
            <div class="flow-frame-vars" id="wfStepDatesPanel">
                <p class="form-hint"><strong>Fechas del paso</strong></p>
                <div class="wf-step-dates-row">
                    ${needsInicio ? `
                        <div class="form-group">
                            <label>Fecha de inicio</label>
                            <input type="date" id="stepRunFechaInicio" value="${UI.escapeHtml(fechaInicio)}">
                        </div>` : ""}
                    ${needsFin ? `
                        <div class="form-group">
                            <label>Fecha de fin</label>
                            <input type="date" id="stepRunFechaFin" value="${UI.escapeHtml(fechaFin)}">
                        </div>` : ""}
                </div>
            </div>`;
    },

    bindStepDatesEvents(step, instances) {
        const inicioInput = document.getElementById("stepRunFechaInicio");
        if (inicioInput) inicioInput.addEventListener("change", (e) => this.updateStepFecha(step, instances, "fechaProgramada", "FECHA_PROGRAMADA", e.target.value));
        const finInput = document.getElementById("stepRunFechaFin");
        if (finInput) finInput.addEventListener("change", (e) => this.updateStepFecha(step, instances, "fechaProgramadaFin", "FECHA_PROGRAMADA_FIN", e.target.value));
    },

    // Aplica la misma fecha a TODAS las instancias del paso (no hay una
    // fecha distinta por instancia): actualiza en memoria y persiste con
    // un único UPDATE por PASO_ID + RUN_ID.
    async updateStepFecha(step, instances, field, column, value) {
        const anteriores = instances.map(i => i[field]);
        instances.forEach(i => { i[field] = value; });
        try {
            await Provider.runQuery(`
                UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                SET ${column} = '${Provider.esc(value)}'
                WHERE RUN_ID = '${Provider.esc(this.currentRun.id)}' AND PASO_ID = '${Provider.esc(step.id)}'`);
            this.renderRunStepBody();
        } catch (err) {
            instances.forEach((i, idx) => { i[field] = anteriores[idx]; });
            UI.toast("Error al guardar la fecha: " + err.message, "error");
        }
    },

    driverLabelCacheKey(step, driverValor) {
        return `${step.driver.dimensionId || ""}::${driverValor}`;
    },

    // Además del código del driver (p.ej. CC-1001), busca el texto
    // descriptivo del miembro en la propia tabla de la dimensión, para
    // mostrar "CC-1001 — Marketing" en la tarjeta. Como el esquema de una
    // dimensión es libre (no hay una columna "descripción" fija), se toma
    // la primera columna de texto que no sea la propia clave. No bloquea
    // el pintado inicial de las tarjetas: se resuelve en segundo plano y
    // luego se aplica solo al título (patchDriverLabels).
    async fetchDriverLabels(step, instances) {
        if (!step.driver.dimensionId) return;
        const dim = Workflows.dimensionById(step.driver.dimensionId);
        if (!dim) return;

        const codes = [...new Set(instances.map(i => i.driverValor).filter(v => v !== null))]
            .filter(code => this._driverLabelCache[this.driverLabelCacheKey(step, code)] === undefined);
        if (!codes.length) return;

        try {
            const keyCol = Provider.toIdentifier(dim.DIMENSION);
            const fullTable = Provider.qualify(this.project.DATASET, dim.TABLA);
            const inList = codes.map(c => `'${Provider.esc(c)}'`).join(",");
            const rows = await Provider.runQuery(`SELECT * FROM ${fullTable} WHERE ${keyCol} IN (${inList})`);
            rows.forEach(row => {
                const otherKeys = Object.keys(row).filter(k => k !== keyCol);
                const labelKey = otherKeys.find(k => typeof row[k] === "string" && row[k]);
                const code = String(row[keyCol]);
                this._driverLabelCache[this.driverLabelCacheKey(step, code)] = labelKey ? row[labelKey] : "";
            });
            codes.forEach(code => {
                const key = this.driverLabelCacheKey(step, code);
                if (this._driverLabelCache[key] === undefined) this._driverLabelCache[key] = "";
            });
            this.patchDriverLabels(step, instances);
        } catch (e) {
            // Sin descripción disponible: se sigue mostrando solo el código.
        }
    },

    patchDriverLabels(step, instances) {
        instances.forEach(inst => {
            const label = this._driverLabelCache[this.driverLabelCacheKey(step, inst.driverValor)];
            if (!label) return;
            const el = document.querySelector(`[data-inst-title="${inst.id}"]`);
            if (el && !el.querySelector(".wf-instance-card-title-sub")) {
                el.insertAdjacentHTML("beforeend", ` <span class="wf-instance-card-title-sub">— ${UI.escapeHtml(label)}</span>`);
            }
        });
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
        const controlsDisabled = inst.estado === "BLOQUEADO" || inst.estado === "COMPLETADO";

        // Los 3 botones (play/stop/restablecer) cubren Pendiente, En curso
        // y Suspendido. Bloqueado/En revisión/Completado son estados que no
        // se fijan a mano con estos botones, así que ahí se muestra la
        // etiqueta en su lugar.
        const showStatusLabel = ["BLOQUEADO", "EN_REVISION", "COMPLETADO"].includes(inst.estado);

        let actions = "";
        if (inst.estado === "BLOQUEADO") {
            actions = `<span class="form-hint">Se desbloqueará al completarse el paso anterior.</span>`;
        } else if (inst.estado === "EN_CURSO") {
            actions = step.revision
                ? `<button class="btn btn-primary btn-sm" data-inst-action="review:${inst.id}">Enviar a revisión</button>`
                : `<button class="btn btn-primary btn-sm" data-inst-action="complete:${inst.id}">Completar</button>`;
        } else if (inst.estado === "EN_REVISION") {
            actions = `<button class="btn btn-primary btn-sm" data-inst-action="approve:${inst.id}">Aprobar</button>
                       <button class="btn btn-secondary btn-sm" data-inst-action="reject:${inst.id}">Rechazar</button>`;
        }

        const title = inst.driverValor !== null ? UI.escapeHtml(inst.driverValor) : "Instancia única";
        const label = this._driverLabelCache[this.driverLabelCacheKey(step, inst.driverValor)];

        return `
            <div class="wf-instance-card ${estadoInfo.cls}" data-instance-card="${inst.id}">
                <div class="wf-instance-card-top">
                    <span class="wf-instance-card-title" data-inst-title="${inst.id}">${title}${label ? ` <span class="wf-instance-card-title-sub">— ${UI.escapeHtml(label)}</span>` : ""}</span>
                    ${showStatusLabel
                        ? `<span class="table-tag ${estadoInfo.cls}">${estadoInfo.label}</span>`
                        : this.instControlsHtml(inst, controlsDisabled)}
                </div>
                ${inst.fechaProgramada ? `<span class="table-tag" style="margin-bottom:8px;">Programado: ${UI.escapeHtml(inst.fechaProgramada)}</span>` : ""}
                ${this.assigneeBlockHtml(step, inst)}
                ${visibleVars.length ? `
                    <div class="form-group">
                        ${visibleVars.map(v => `
                            <label>${UI.escapeHtml(v.label)}</label>
                            <input type="text" value="${UI.escapeHtml(inst.variables[v.name] || "")}" data-inst-var="${inst.id}:${UI.escapeHtml(v.name)}" ${inst.estado === "COMPLETADO" ? "disabled" : ""} style="margin-bottom:6px;">
                        `).join("")}
                    </div>` : ""}
                ${actions ? `<div class="row-actions" style="justify-content:flex-start; gap:8px;">${actions}</div>` : ""}
            </div>`;
    },

    // Botones Play / Stop / Restablecer: fijan directamente el estado de
    // la instancia a En curso / Suspendido / Pendiente. El botón activo
    // (el que coincide con el estado actual) queda resaltado.
    instControlsHtml(inst, disabled) {
        const mk = (estado, icon, title) => `
            <button type="button" class="wf-inst-ctrl-btn ${inst.estado === estado ? "is-active is-active--" + estado.toLowerCase() : ""}"
                    data-inst-set="${estado}:${inst.id}" title="${title}" ${disabled ? "disabled" : ""}>${icon}</button>`;
        return `<div class="wf-inst-controls">
            ${mk("EN_CURSO", "▶", "Poner en curso")}
            ${mk("SUSPENDIDO", "⏸", "Suspender")}
            ${mk("PENDIENTE", "↺", "Restablecer a pendiente")}
        </div>`;
    },

    // Bloque "Asignado a": si el paso tiene revisión, se muestran dos
    // pestañas (Responsable / Revisor), cada una con su propio selector
    // múltiple de personas/grupos; si no tiene revisión, solo Responsable
    // sin pestañas.
    assigneeBlockHtml(step, inst) {
        const tab = step.revision ? (this._assigneeTab[inst.id] || "asignado") : "asignado";
        const tabs = step.revision ? `
            <div class="wf-assignee-tabs" data-assignee-tabs="${inst.id}">
                <button type="button" class="wf-assignee-tab ${tab === "asignado" ? "is-active" : ""}" data-assignee-tab="asignado:${inst.id}">Responsable</button>
                <button type="button" class="wf-assignee-tab ${tab === "revisor" ? "is-active" : ""}" data-assignee-tab="revisor:${inst.id}">Revisor</button>
            </div>` : "";

        return `
            <div class="form-group">
                ${step.revision ? "" : "<label>Asignado a</label>"}
                ${tabs}
                <div data-assignee-panel="${inst.id}">${this.multiAssigneeFieldHtml(inst, tab)}</div>
            </div>`;
    },

    // Selector múltiple de personas/grupos para el campo indicado
    // ("asignado" o "revisor"). Los valores se guardan en la misma
    // columna (ASIGNADO/REVISOR) como texto separado por comas. Si el
    // proveedor soporta buscador de usuarios se ofrece autocompletado;
    // si no (o para grupos, que no están en el directorio de usuarios),
    // se puede escribir libremente y pulsar Enter para añadir el chip.
    multiAssigneeFieldHtml(inst, field) {
        const disabled = inst.estado === "COMPLETADO";
        const raw = field === "revisor" ? (inst.revisor || "") : (inst.asignado || "");
        const values = raw.split(",").map(v => v.trim()).filter(Boolean);
        const chips = values.map(v => {
            const isGroup = !v.includes("@");
            return `<span class="wf-assignee-chip ${isGroup ? "wf-assignee-chip--group" : ""}">
                        <span class="wf-assignee-chip-icon">${isGroup ? "👥" : "👤"}</span>${UI.escapeHtml(v)}
                        ${disabled ? "" : `<button type="button" class="wf-assignee-chip-remove" data-assignee-remove="${field}:${inst.id}:${UI.escapeHtml(v)}">×</button>`}
                    </span>`;
        }).join("");

        const canSearch = Provider.canSearchUsers();
        return `
            <div class="wf-assignee-multi" data-assignee-multi="${field}:${inst.id}">
                <div class="wf-assignee-chip-row">${chips}</div>
                ${disabled ? "" : `
                    <div class="user-search" data-user-search="${field}:${inst.id}">
                        <input type="text" class="user-search-input" placeholder="${canSearch ? "Buscar por email o escribe un grupo…" : "Persona o grupo…"}" autocomplete="off"
                               data-assignee-input="${field}:${inst.id}" data-user-search-input>
                        <div class="user-search-results" data-user-search-results style="display:none;"></div>
                    </div>`}
            </div>`;
    },

    bindRunEvents(scope) {
        scope.querySelectorAll("[data-open-flow]").forEach(btn => {
            btn.addEventListener("click", () => window.open(`flow_run.html?flujo_id=${encodeURIComponent(btn.dataset.openFlow)}`, "_blank"));
        });
        scope.querySelectorAll("[data-user-search]").forEach(wrap => this.bindUserSearch(wrap));
        scope.querySelectorAll("[data-assignee-remove]").forEach(btn => {
            btn.addEventListener("click", () => {
                const [field, instId, value] = btn.dataset.assigneeRemove.split(":");
                this.removeAssignee(instId, field, value);
            });
        });
        scope.querySelectorAll("[data-assignee-tab]").forEach(btn => {
            btn.addEventListener("click", () => {
                const [field, instId] = btn.dataset.assigneeTab.split(":");
                this._assigneeTab[instId] = field;
                this.renderRunStepBody();
            });
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
        scope.querySelectorAll("[data-inst-set]").forEach(btn => {
            btn.addEventListener("click", () => {
                const [estado, instId] = btn.dataset.instSet.split(":");
                this.setInstanceEstado(instId, estado);
            });
        });
    },

    // Buscador de usuarios del campo de asignación (Responsable/Revisor):
    // consulta Provider.searchUsers() con un pequeño debounce y muestra un
    // desplegable con los emails que coinciden. Como es un selector
    // MÚLTIPLE, elegir un resultado (o pulsar Enter con lo escrito) AÑADE
    // un chip nuevo en vez de sustituir el valor. Lo escrito que no
    // coincide con ningún email del directorio se puede añadir igualmente
    // como grupo (el directorio de usuarios de Provider.searchUsers no
    // conoce grupos, así que un grupo siempre se añade escribiendo su
    // nombre y pulsando Enter o el botón "Añadir").
    bindUserSearch(wrap) {
        const input = wrap.querySelector("[data-user-search-input]");
        const results = wrap.querySelector("[data-user-search-results]");
        if (!input || !results) return;
        const [field, instId] = input.dataset.assigneeInput.split(":");

        const addValue = (value) => {
            value = value.trim();
            if (!value) return;
            results.style.display = "none";
            input.value = "";
            this.addAssignee(instId, field, value);
        };

        let debounceTimer = null;
        const runSearch = async () => {
            let users = [];
            try {
                users = await Provider.searchUsers(input.value);
            } catch (err) {
                results.style.display = "none";
                return;
            }
            if (!users.length) { results.style.display = "none"; results.innerHTML = ""; return; }
            results.innerHTML = users.map(email =>
                `<div class="user-search-item" data-user-email="${UI.escapeHtml(email)}">👤 ${UI.escapeHtml(email)}</div>`).join("");
            results.style.display = "block";
            results.querySelectorAll("[data-user-email]").forEach(item => {
                // mousedown (no click) para que se dispare ANTES del blur del input y no se pierda la selección
                item.addEventListener("mousedown", (e) => {
                    e.preventDefault();
                    addValue(item.dataset.userEmail);
                });
            });
        };

        input.addEventListener("input", () => {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(runSearch, 250);
        });
        input.addEventListener("focus", () => { if (input.value) runSearch(); });
        input.addEventListener("keydown", (e) => {
            if (e.key === "Enter") { e.preventDefault(); addValue(input.value); }
        });
        input.addEventListener("blur", () => { setTimeout(() => { results.style.display = "none"; }, 150); });
    },

    async saveAssigneeField(instId, field, values) {
        const value = values.join(",");
        const column = field === "revisor" ? "REVISOR" : "ASIGNADO";
        try {
            await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                SET ${column} = '${Provider.esc(value)}' WHERE INSTANCIA_ID = '${Provider.esc(instId)}'`);
            const inst = this.currentRun.instancias.find(i => i.id === instId);
            if (inst) inst[field] = value;
            this.renderRunChain();
            this.renderRunStepBody();
        } catch (err) {
            UI.toast("Error al guardar el " + (field === "revisor" ? "revisor" : "responsable") + ": " + err.message, "error");
        }
    },

    addAssignee(instId, field, value) {
        const inst = this.currentRun.instancias.find(i => i.id === instId);
        if (!inst) return;
        const current = (inst[field] || "").split(",").map(v => v.trim()).filter(Boolean);
        if (current.includes(value)) return;
        current.push(value);
        this.saveAssigneeField(instId, field, current);
    },

    removeAssignee(instId, field, value) {
        const inst = this.currentRun.instancias.find(i => i.id === instId);
        if (!inst) return;
        const current = (inst[field] || "").split(",").map(v => v.trim()).filter(Boolean).filter(v => v !== value);
        this.saveAssigneeField(instId, field, current);
    },

    async setInstanceEstado(instId, nuevoEstado) {
        const inst = this.currentRun.instancias.find(i => i.id === instId);
        if (!inst || inst.estado === nuevoEstado) return;
        try {
            await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                SET ESTADO = '${Provider.esc(nuevoEstado)}'${nuevoEstado === "EN_CURSO" ? ", FECHA_INICIO = CURRENT_TIMESTAMP()" : ""}
                WHERE INSTANCIA_ID = '${Provider.esc(instId)}'`);
            inst.estado = nuevoEstado;
            this.renderRunChain();
            this.renderRunStepBody();
        } catch (err) {
            UI.toast("Error al actualizar el estado: " + err.message, "error");
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

        const map = { review: "EN_REVISION", approve: "COMPLETADO", reject: "EN_CURSO", complete: "COMPLETADO" };
        const nuevoEstado = map[action];
        if (!nuevoEstado) return;

        const isFinal = nuevoEstado === "COMPLETADO";
        try {
            await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS_RUNS_INSTANCIAS")}
                SET ESTADO = '${Provider.esc(nuevoEstado)}'${isFinal ? ", FECHA_FIN = CURRENT_TIMESTAMP()" : ""}
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
