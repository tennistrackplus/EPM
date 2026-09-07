/**
 * ============================================================
 * DRACO PLANNING — WORKFLOWS (definición)
 * ============================================================
 * Un Workflow es una secuencia de Pasos, empezando siempre por un
 * **Paso 0** especial (no eliminable) donde solo se definen las
 * **variables del workflow** — de valor único, se piden al crear cada
 * ejecución y están disponibles para completar tareas en cualquier paso.
 *
 * El resto de pasos tienen 2 pestañas:
 *   - Propiedades: nombre, Inicio (izquierda) / Finalización (derecha)
 *     y, debajo, el Driver — todo junto en la misma pantalla.
 *       · Inicio: al iniciar el workflow / al completar el paso
 *         anterior / fecha concreta.
 *       · Revisión (sí/no).
 *       · Finalización: N/A, al enviarse a revisión (si revisión=true),
 *         al completarse (si revisión=false), o fecha concreta — se
 *         recalcula sola al cambiar Revisión salvo que el usuario haya
 *         elegido N/A o Fecha. Es lo que dispara el "al completar el
 *         paso anterior" del siguiente paso.
 *       · Driver: dimensión opcional por la que se reparte la ejecución
 *         del paso (ej. nodo superior de la jerarquía de CECOs) +
 *         valores concretos opcionales (si no se eligen, se reparte
 *         entre TODOS los valores de la dimensión). Sin driver, el paso
 *         se asigna en bloque a una persona o grupo.
 *   - Tareas del paso: organizadas en bloques — un menú lateral
 *     (igual que Dimensiones/Cubos/Interfaces) donde "+ Añadir bloque"
 *     crea uno nuevo y se reordenan arrastrando; a la derecha, las
 *     tareas del bloque seleccionado. Al añadir una tarea se pide
 *     nombre + descripción y se elige su tipo con tarjetas (igual que el
 *     selector de origen de datos en "Nueva interfaz"): Actualización de
 *     tablas, Flujos de carga, Mantenimiento de dimensiones, Plantillas
 *     Excel, Plantillas Web, Funciones o Páginas HTML. "Actualización de
 *     tablas", "Flujos de carga" y "Mantenimiento de dimensiones" se
 *     eligen de su catálogo real (ACTUALIZACIONES / FLUJOS tipo MANUAL /
 *     DIMENSIONES del proyecto); el resto son referencia libre por ahora.
 *     Cada tarea puede completar
 *     sus propias variables por constante o por una variable del workflow
 *     (Paso 0), con opción de ocultarlas en la pantalla de ejecución, y se
 *     puede editar (✎) o eliminar (✕) en cualquier momento.
 *
 * Persistencia en DRACO_CONTROL:
 *   - WORKFLOWS                        cabecera (nombre, descripción)
 *   - WORKFLOWS_PASOS                  pasos (propiedades + driver simple + ES_PASO0)
 *   - WORKFLOWS_PASOS_DRIVER_VALORES   valores concretos del driver
 *   - WORKFLOWS_PASOS_VARIABLES        variables (solo se usan en el Paso 0)
 *   - WORKFLOWS_PASOS_BLOQUES          bloques de tareas
 *   - WORKFLOWS_PASOS_TAREAS           tareas dentro de cada bloque
 *   - WORKFLOWS_PASOS_TAREAS_VALORES   valores/variables de cada tarea
 *
 * La ejecución (crear/gestionar instancias de un workflow ya definido)
 * vive en js/workflow-runs.js.
 */
const Workflows = {
    TABLE: "WORKFLOWS",
    ID_COL: "WORKFLOW_ID",
    NAME_COL: "WORKFLOW",

    TASK_TYPES: {
        // "catalog" marca los tipos cuya referencia se elige de un listado
        // real (en vez de texto libre): PARAMETRIZACION -> Actualizaciones,
        // FLUJO_MANUAL -> Flujos de carga de tipo manual.
        PARAMETRIZACION: { label: "Actualización de tablas", icon: "🗄", catalog: "ACTUALIZACION" },
        FLUJO_MANUAL: { label: "Flujos de carga", icon: "☺", catalog: "FLUJO" },
        MANTENIMIENTO_DIMENSION: { label: "Mantenimiento de dimensiones", icon: "🧩", catalog: "DIMENSION" },
        PLANTILLA_EXCEL: { label: "Plantillas Excel", icon: "📊", catalog: "BUCKET_XLSX" },
        PLANTILLA_WEB: { label: "Plantillas Web", icon: "🌐" },
        FUNCION: { label: "Funciones", icon: "ƒ" },
        HTML: { label: "Páginas HTML", icon: "⌗" },
        // Legado: tareas creadas antes de este cambio con el tipo genérico
        // "Plantilla" (sin distinguir Excel/Web). Ya no se ofrece en "+ Tarea"
        // pero se sigue mostrando correctamente si ya existía.
        PLANTILLA: { label: "Plantilla", icon: "▤", legacy: true }
    },

    // Orden y contenido del selector de tipo al crear una tarea nueva
    // (excluye los tipos "legacy").
    taskTypeChoices() {
        return Object.entries(this.TASK_TYPES)
            .filter(([, t]) => !t.legacy)
            .map(([key, t]) => ({ key, label: t.label, icon: t.icon }));
    },

    list: [],
    dimensions: [],
    manualFlows: [],
    actualizaciones: [],
    editing: null,
    editingIsNew: true,
    selectedStepId: null,
    selectedBlockId: null,
    activeTab: "propiedades",
    dragStepIdx: null,
    dragBlockIdx: null,

    async render(container, project) {
        this.container = container;
        this.project = project;

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <h3>Workflows</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET}</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewWorkflow">+ Nuevo workflow</button>
            </div>
            <div id="workflowsListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewWorkflow").addEventListener("click", () => this.openForm());

        await this.loadList();
        this.renderList();
    },

    // ------------------------------------------------------------
    // Listado
    // ------------------------------------------------------------
    async loadList() {
        try {
            // Las 2 consultas filtran por PROYECTO_ID directamente (la de
            // pasos no depende de los WORKFLOW_ID obtenidos en la primera),
            // así que se piden a la vez en vez de esperar a la primera para
            // lanzar la segunda.
            const [rows, steps] = await Promise.all([
                Provider.runQuery(`
                    SELECT ${this.ID_COL}, ${this.NAME_COL}, DESCRIPCION
                    FROM ${Provider.qualifyControl(this.TABLE)}
                    WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                    ORDER BY ${this.NAME_COL}`),
                Provider.runQuery(`
                    SELECT WORKFLOW_ID, COUNT(*) AS N
                    FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")}
                    WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}' AND (ES_PASO0 IS NULL OR ES_PASO0 = FALSE)
                    GROUP BY WORKFLOW_ID`)
            ]);

            const stepCounts = {};
            steps.forEach(s => { stepCounts[s.WORKFLOW_ID] = parseInt(s.N || "0", 10); });

            this.list = rows.map(r => ({
                id: r[this.ID_COL],
                name: r[this.NAME_COL],
                description: r.DESCRIPCION || "",
                stepCount: stepCounts[r[this.ID_COL]] || 0
            }));
        } catch (err) {
            this.list = [];
            UI.toast("Error al cargar los workflows: " + err.message, "error");
        }
    },

    renderList() {
        const wrap = document.getElementById("workflowsListWrap");
        if (!wrap) return;

        if (!this.list.length) {
            wrap.innerHTML = `<div class="module-empty">Todavía no hay workflows en este proyecto. Crea el primero con "Nuevo workflow".</div>`;
            return;
        }

        wrap.innerHTML = `
            <div class="data-list">
                <table>
                    <thead><tr><th>Workflow</th><th>Descripción</th><th>Pasos</th><th></th></tr></thead>
                    <tbody>
                        ${this.list.map(w => `
                            <tr>
                                <td><strong>${UI.escapeHtml(w.name)}</strong></td>
                                <td>${UI.escapeHtml(w.description || "—")}</td>
                                <td>${w.stepCount} paso${w.stepCount === 1 ? "" : "s"}</td>
                                <td>
                                    <div class="row-actions">
                                        <button data-runs-wf="${w.id}" title="Ejecuciones">▶</button>
                                        <button data-edit-wf="${w.id}" title="Editar">✎</button>
                                        <button data-del-wf="${w.id}" class="danger" title="Eliminar">🗑</button>
                                    </div>
                                </td>
                            </tr>`).join("")}
                    </tbody>
                </table>
            </div>`;

        wrap.querySelectorAll("[data-edit-wf]").forEach(btn =>
            btn.addEventListener("click", () => this.openForm(btn.dataset.editWf)));
        wrap.querySelectorAll("[data-del-wf]").forEach(btn =>
            btn.addEventListener("click", () => this.remove(btn.dataset.delWf)));
        wrap.querySelectorAll("[data-runs-wf]").forEach(btn =>
            btn.addEventListener("click", () => WorkflowRuns.open(this.container, this.project, btn.dataset.runsWf)));
    },

    async remove(id) {
        const wf = this.list.find(w => w.id === id);
        if (!wf) return;
        const ok = await UI.confirm("Eliminar workflow", `Se eliminará el workflow <strong>${UI.escapeHtml(wf.name)}</strong> y todos sus pasos, variables y tareas.`);
        if (!ok) return;
        try {
            // Mismo criterio que en persist(): TAREAS_VALORES depende de que
            // TAREAS todavía exista (subconsulta anidada, sin ir primero a
            // buscar PASO_ID/TAREA_ID por separado); el resto de tablas que
            // cuelgan de PASO_ID son independientes entre sí y se borran en
            // paralelo; WORKFLOWS_PASOS y la cabecera van al final.
            await Provider.runQuery(`
                DELETE FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS_VALORES")}
                WHERE TAREA_ID IN (
                    SELECT TAREA_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")}
                    WHERE PASO_ID IN (SELECT PASO_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}')
                )`);

            await Promise.all(["WORKFLOWS_PASOS_TAREAS", "WORKFLOWS_PASOS_DRIVER_VALORES", "WORKFLOWS_PASOS_VARIABLES", "WORKFLOWS_PASOS_BLOQUES"].map(table =>
                Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(table)} WHERE PASO_ID IN (SELECT PASO_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}')`)
            ));

            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}'`);
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);

            await this.loadList();
            this.renderList();
            UI.toast(`Workflow "${wf.name}" eliminado.`, "success");
        } catch (err) {
            UI.toast("Error al eliminar el workflow: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Datos auxiliares (dimensiones del proyecto, flujos manuales)
    // ------------------------------------------------------------
    async loadDimensions() {
        try {
            this.dimensions = await Provider.runQuery(`
                SELECT DIMENSION_ID, DIMENSION, DESCRIPCION, TABLA
                FROM ${Provider.qualifyControl("DIMENSIONES")}
                WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                ORDER BY DIMENSION`);
        } catch (e) { this.dimensions = []; }
    },

    async loadManualFlows() {
        try {
            const rows = await Provider.runQuery(`
                SELECT FLUJO_ID, FLUJO
                FROM ${Provider.qualifyControl("FLUJOS")}
                WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}' AND TIPO = 'MANUAL'
                ORDER BY FLUJO`);
            this.manualFlows = rows.map(r => ({ id: r.FLUJO_ID, name: r.FLUJO }));
        } catch (e) { this.manualFlows = []; }
    },

    // Catálogo de "Actualización de tablas" (ver js/table-updates.js) para
    // usarlas como referencia de una tarea de tipo PARAMETRIZACION.
    async loadActualizaciones() {
        try {
            const rows = await Provider.runQuery(`
                SELECT ACTUALIZACION_ID, NOMBRE, TABLA
                FROM ${Provider.qualifyControl("ACTUALIZACIONES")}
                WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                ORDER BY NOMBRE`);
            this.actualizaciones = rows.map(r => ({ id: r.ACTUALIZACION_ID, name: r.NOMBRE, tabla: r.TABLA }));
        } catch (e) { this.actualizaciones = []; }
    },

    async flowScreenVariables(flujoId) {
        try {
            const rows = await Provider.runQuery(`
                SELECT NOMBRE, ETIQUETA, TIPO
                FROM ${Provider.qualifyControl("FLUJOS_SCREEN_VARIABLES")}
                WHERE FLUJO_ID = '${Provider.esc(flujoId)}' ORDER BY ORDEN`);
            return rows;
        } catch (e) { return []; }
    },

    // Auto-reparación: añade la columna DESCRIPCION a WORKFLOWS_PASOS_TAREAS si
    // la tabla se creó antes de que existiera este campo (ver schema.js).
    async ensureTareaDescripcionColumn() {
        try {
            await Provider.runQuery(`ALTER TABLE ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")} ADD COLUMN IF NOT EXISTS DESCRIPCION STRING`);
        } catch (err) {
            console.error("No se pudo comprobar/añadir la columna DESCRIPCION en WORKFLOWS_PASOS_TAREAS:", err);
        }
    },

    dimensionById(id) {
        return this.dimensions.find(d => d.DIMENSION_ID === id) || null;
    },

    // ------------------------------------------------------------
    // Blancos
    // ------------------------------------------------------------
    blankWorkflow() {
        return { id: Provider.newId(), name: "", description: "", steps: [this.blankPaso0()] };
    },

    blankPaso0() {
        return {
            id: Provider.newId(),
            isPaso0: true,
            name: "Variables del workflow",
            inicio: { tipo: "INICIO_WORKFLOW", fecha: "" },
            revision: false,
            noBloqueaRevision: false,
            fin: { tipo: "COMPLETAR", fecha: "" },
            driver: { dimensionId: null, valores: [] },
            variables: [],
            bloques: []
        };
    },

    blankStep(name) {
        return {
            id: Provider.newId(),
            isPaso0: false,
            name: name || "",
            inicio: { tipo: "INICIO_WORKFLOW", fecha: "" },
            revision: false,
            noBloqueaRevision: false,
            fin: { tipo: "COMPLETAR", fecha: "" },
            driver: { dimensionId: null, valores: [] },
            variables: [],
            bloques: []
        };
    },

    // ------------------------------------------------------------
    // Carga completa de un workflow existente
    // ------------------------------------------------------------
    async loadDetail(id) {
        // Igual que en persist(): antes esto encadenaba 7 idas y vueltas
        // secuenciales (cabecera → pasos → [driver, variables, bloques,
        // tareas] → valores de tareas), porque cada una esperaba a la
        // anterior para sacar en JS la lista de PASO_ID/TAREA_ID por la que
        // filtrar. En realidad todas dependen solo del WORKFLOW_ID, no del
        // resultado de las demás — basta con expresar el filtro como una
        // subconsulta anidada (WORKFLOW_ID → PASO_ID → TAREA_ID) en vez de
        // un "IN (lista calculada en JS)", y así las 7 se pueden lanzar
        // todas a la vez con Promise.all en lugar de una detrás de otra.
        const pasosOfWorkflow = `SELECT PASO_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}'`;
        const tareasOfWorkflow = `SELECT TAREA_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")} WHERE PASO_ID IN (${pasosOfWorkflow})`;

        const [headerRows, pasoRows, driverValRows, varRows, bloqueRows, tareaRows, tareaValRows] = await Promise.all([
            Provider.runQuery(`SELECT * FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`),
            Provider.runQuery(`SELECT * FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}' ORDER BY ORDEN`),
            Provider.runQuery(`SELECT PASO_ID, VALOR FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_DRIVER_VALORES")} WHERE PASO_ID IN (${pasosOfWorkflow})`),
            Provider.runQuery(`SELECT PASO_ID, VARIABLE_ID, NOMBRE, ETIQUETA, TIPO, ORDEN FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_VARIABLES")} WHERE PASO_ID IN (${pasosOfWorkflow}) ORDER BY ORDEN`),
            Provider.runQuery(`SELECT PASO_ID, BLOQUE_ID, TITULO, ORDEN FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_BLOQUES")} WHERE PASO_ID IN (${pasosOfWorkflow}) ORDER BY ORDEN`),
            Provider.runQuery(`SELECT PASO_ID, BLOQUE_ID, TAREA_ID, TIPO, NOMBRE, DESCRIPCION, REF_ID, REF_NOMBRE, ORDEN FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")} WHERE PASO_ID IN (${pasosOfWorkflow}) ORDER BY ORDEN`),
            Provider.runQuery(`SELECT TAREA_ID, CLAVE, ETIQUETA, TIPO, VALOR, OCULTAR FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS_VALORES")} WHERE TAREA_ID IN (${tareasOfWorkflow})`)
        ]);

        const row = headerRows[0];
        if (!row) return null;

        const driverValsByPaso = {};
        driverValRows.forEach(d => { (driverValsByPaso[d.PASO_ID] = driverValsByPaso[d.PASO_ID] || []).push(d.VALOR); });

        const varsByPaso = {};
        varRows.forEach(v => { (varsByPaso[v.PASO_ID] = varsByPaso[v.PASO_ID] || []).push({ id: v.VARIABLE_ID, name: v.NOMBRE, label: v.ETIQUETA || v.NOMBRE, type: v.TIPO || "STRING" }); });

        const valsByTarea = {};
        tareaValRows.forEach(v => {
            (valsByTarea[v.TAREA_ID] = valsByTarea[v.TAREA_ID] || []).push({
                clave: v.CLAVE, etiqueta: v.ETIQUETA || v.CLAVE,
                tipo: (v.TIPO || "CONSTANTE").toLowerCase() === "variable" ? "variable" : "constante",
                valor: v.VALOR || "", ocultar: !!(v.OCULTAR === true || v.OCULTAR === "true" || v.OCULTAR === 1)
            });
        });

        const tareasByBloque = {};
        tareaRows.forEach(t => {
            (tareasByBloque[t.BLOQUE_ID] = tareasByBloque[t.BLOQUE_ID] || []).push({
                id: t.TAREA_ID, tipo: t.TIPO, nombre: t.NOMBRE || "", descripcion: t.DESCRIPCION || "",
                refId: t.REF_ID || null, refNombre: t.REF_NOMBRE || "",
                valores: valsByTarea[t.TAREA_ID] || []
            });
        });

        const bloquesByPaso = {};
        bloqueRows.forEach(b => {
            (bloquesByPaso[b.PASO_ID] = bloquesByPaso[b.PASO_ID] || []).push({
                id: b.BLOQUE_ID, titulo: b.TITULO || "Bloque",
                tareas: tareasByBloque[b.BLOQUE_ID] || []
            });
        });

        let steps = pasoRows.map(p => ({
            id: p.PASO_ID,
            isPaso0: !!(p.ES_PASO0 === true || p.ES_PASO0 === "true" || p.ES_PASO0 === 1),
            name: p.PASO,
            inicio: { tipo: p.INICIO_TIPO || "INICIO_WORKFLOW", fecha: p.INICIO_FECHA || "" },
            revision: !!(p.REVISION === true || p.REVISION === "true" || p.REVISION === 1),
            noBloqueaRevision: !!(p.NO_BLOQUEA_REVISION === true || p.NO_BLOQUEA_REVISION === "true" || p.NO_BLOQUEA_REVISION === 1),
            fin: { tipo: p.FIN_TIPO || "COMPLETAR", fecha: p.FIN_FECHA || "" },
            driver: { dimensionId: p.DRIVER_DIMENSION_ID || null, valores: driverValsByPaso[p.PASO_ID] || [] },
            variables: varsByPaso[p.PASO_ID] || [],
            bloques: bloquesByPaso[p.PASO_ID] || []
        }));

        // Compatibilidad con workflows creados antes de que existiera el Paso 0.
        if (!steps.some(s => s.isPaso0)) steps.unshift(this.blankPaso0());

        return { id: row[this.ID_COL], name: row[this.NAME_COL], description: row.DESCRIPCION || "", steps };
    },

    // ------------------------------------------------------------
    // Alta: nombre + descripción, luego editor completo
    // ------------------------------------------------------------
    async openForm(editId = null) {
        this.editingIsNew = !editId;
        this.selectedStepId = null;
        this.selectedBlockId = null;
        this.activeTab = "propiedades";

        // Los 4 catálogos auxiliares (dimensiones, flujos manuales,
        // actualizaciones, auto-reparación de columna) no dependen entre sí
        // ni de loadDetail — se piden todos a la vez, junto con el propio
        // workflow si se está editando uno existente, en vez de uno detrás
        // de otro.
        const [, , , , draft] = await Promise.all([
            this.loadDimensions(),
            this.loadManualFlows(),
            this.loadActualizaciones(),
            this.ensureTareaDescripcionColumn(),
            editId ? this.loadDetail(editId) : Promise.resolve(null)
        ]);

        if (editId) {
            if (!draft) { UI.toast("No se ha podido cargar el workflow.", "error"); return; }
            this.editing = draft;
            this.selectedStepId = draft.steps[0].id;
            this.openMainModal();
            return;
        }

        const blank = this.blankWorkflow();
        const basics = await this.openBasicsModal(blank, true);
        if (!basics) return;
        Object.assign(blank, basics);
        this.editing = blank;
        this.selectedStepId = blank.steps[0].id;
        this.openMainModal();
    },

    openBasicsModal(initial, isNew) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("wfBasicsModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wfBasicsModal";
                document.body.appendChild(overlay);
            }
            overlay.innerHTML = `
                <div class="modal-box">
                    <div class="modal-header">
                        <h3>${isNew ? "Nuevo workflow" : "Datos básicos"}</h3>
                        <button class="modal-close" id="wfBasicsClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Nombre del workflow</label>
                            <input type="text" id="wfBasicsName" placeholder="Ej. Cierre mensual" value="${UI.escapeHtml(initial.name || "")}">
                        </div>
                        <div class="form-group">
                            <label>Descripción</label>
                            <textarea id="wfBasicsDesc" rows="3" placeholder="Describe brevemente el propósito de este workflow...">${UI.escapeHtml(initial.description || "")}</textarea>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="wfBasicsCancel">Cancelar</button>
                        <button class="btn btn-primary" id="wfBasicsNext">Continuar</button>
                    </div>
                </div>`;

            overlay.classList.add("visible");
            const nameInput = overlay.querySelector("#wfBasicsName");
            setTimeout(() => { nameInput.focus(); }, 50);

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#wfBasicsClose").onclick = () => cleanup(null);
            overlay.querySelector("#wfBasicsCancel").onclick = () => cleanup(null);
            overlay.querySelector("#wfBasicsNext").onclick = () => {
                const name = nameInput.value.trim();
                if (!name) { UI.toast("Indica un nombre para el workflow.", "error"); return; }
                cleanup({ name, description: overlay.querySelector("#wfBasicsDesc").value.trim() });
            };
            nameInput.onkeydown = (e) => { if (e.key === "Enter") overlay.querySelector("#wfBasicsNext").click(); };
        });
    },

    // ------------------------------------------------------------
    // Editor completo
    // ------------------------------------------------------------
    openMainModal() {
        let overlay = document.getElementById("wfFormModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "wfFormModal";
            document.body.appendChild(overlay);
        }
        this.overlay = overlay;

        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div class="modal-title-edit-wrap">
                        <h3 id="wfModalTitle" class="modal-title-editable" contenteditable="true" spellcheck="false" title="Clic para renombrar el workflow"></h3>
                        <span class="modal-subtitle" id="wfModalSubtitle" contenteditable="true" spellcheck="false" title="Clic para editar la descripción"></span>
                    </div>
                    <div class="modal-header-right">
                        <button class="modal-close" id="wfFormClose">&times;</button>
                    </div>
                </div>
                <div class="modal-body modal-body-flush">
                    <div id="wfStepsPart" class="flow-part flow-part--chain"></div>
                    <div id="wfStepDetailPart" class="flow-part flow-part--mapping"></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="wfFormCancel">Cancelar</button>
                    <button class="btn btn-primary" id="wfFormSave">Guardar workflow</button>
                </div>
            </div>`;

        document.getElementById("wfFormClose").addEventListener("click", () => this.closeForm());
        document.getElementById("wfFormCancel").addEventListener("click", () => this.closeForm());
        document.getElementById("wfFormSave").addEventListener("click", () => this.save());

        const titleEl = document.getElementById("wfModalTitle");
        titleEl.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); titleEl.blur(); } });
        titleEl.addEventListener("blur", () => {
            const newName = titleEl.textContent.trim();
            if (!newName) { UI.toast("El workflow necesita un nombre.", "error"); titleEl.textContent = this.editing.name; return; }
            this.editing.name = newName;
            titleEl.textContent = newName;
        });

        const subEl = document.getElementById("wfModalSubtitle");
        subEl.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); subEl.blur(); } });
        subEl.addEventListener("blur", () => { this.editing.description = subEl.textContent.trim(); });

        overlay.classList.add("visible");

        document.getElementById("wfModalTitle").textContent = this.editing.name;
        document.getElementById("wfModalSubtitle").textContent = this.editing.description || "Añade una descripción...";

        this.renderStepsPart();
        this.renderStepDetail();
    },

    closeForm() {
        if (this.overlay) this.overlay.classList.remove("visible");
        this.editing = null;
    },

    // ------------------------------------------------------------
    // Pasos (cadena, arrastrable — el Paso 0 siempre va primero y fijo)
    // ------------------------------------------------------------
    renderStepsPart() {
        const part = document.getElementById("wfStepsPart");
        const wf = this.editing;

        const cards = wf.steps.map((s, idx) => {
            const selected = s.id === this.selectedStepId;
            const card = s.isPaso0 ? `
                <div class="flow-chain-card wf-paso0-card ${selected ? "is-selected" : ""}" data-step-idx="${idx}" title="Variables del workflow">
                    <div class="flow-chain-card-name">🧩 Variables</div>
                    <div class="flow-chain-card-meta">${s.variables.length} variable${s.variables.length === 1 ? "" : "s"}</div>
                </div>` : `
                <div class="flow-chain-card ${selected ? "is-selected" : ""}" draggable="true" data-step-idx="${idx}" title="Clic para editar el paso">
                    <div class="flow-chain-card-name">${idx}. ${UI.escapeHtml(s.name || "(sin nombre)")}</div>
                    <div class="flow-chain-card-meta">${this.stepMetaLabel(s)}</div>
                    <button type="button" class="flow-chain-card-remove" data-remove-step="${idx}" title="Eliminar paso">✕</button>
                </div>`;
            const arrow = idx < wf.steps.length - 1 ? `<div class="flow-chain-arrow">→</div>` : "";
            return card + arrow;
        }).join("");

        part.innerHTML = `
            <div class="flow-part-header"><strong>Pasos del workflow</strong></div>
            <div class="flow-chain-wrap" id="wfStepsWrap">
                ${cards}
                <button type="button" class="flow-chain-add" id="btnAddStep">+ Añadir paso</button>
            </div>`;

        document.getElementById("btnAddStep").addEventListener("click", () => this.addStep());

        part.querySelectorAll("[data-step-idx]").forEach(card => {
            card.addEventListener("click", (e) => {
                if (e.target.closest("[data-remove-step]")) return;
                const idx = parseInt(card.dataset.stepIdx, 10);
                this.selectedStepId = wf.steps[idx].id;
                this.selectedBlockId = null;
                this.activeTab = "propiedades";
                this.renderStepsPart();
                this.renderStepDetail();
            });
            if (card.getAttribute("draggable") === "true") {
                card.addEventListener("dragstart", () => { this.dragStepIdx = parseInt(card.dataset.stepIdx, 10); });
                card.addEventListener("dragover", (e) => e.preventDefault());
                card.addEventListener("drop", (e) => {
                    e.preventDefault();
                    const toIdx = parseInt(card.dataset.stepIdx, 10);
                    if (toIdx === 0 || this.dragStepIdx === null || this.dragStepIdx === toIdx) return;
                    const [moved] = wf.steps.splice(this.dragStepIdx, 1);
                    wf.steps.splice(toIdx, 0, moved);
                    this.dragStepIdx = null;
                    this.renderStepsPart();
                });
            }
        });
        part.querySelectorAll("[data-remove-step]").forEach(btn => {
            btn.addEventListener("click", async (e) => {
                e.stopPropagation();
                const idx = parseInt(btn.dataset.removeStep, 10);
                const step = wf.steps[idx];
                const ok = await UI.confirm("Eliminar paso", `Se eliminará el paso <strong>${UI.escapeHtml(step.name || "(sin nombre)")}</strong> con toda su configuración.`);
                if (!ok) return;
                wf.steps.splice(idx, 1);
                if (this.selectedStepId === step.id) this.selectedStepId = wf.steps[0].id;
                this.renderStepsPart();
                this.renderStepDetail();
            });
        });
    },

    stepMetaLabel(s) {
        const bits = [];
        bits.push(s.revision ? "Con revisión" : "Sin revisión");
        if (s.driver.dimensionId) {
            const dim = this.dimensionById(s.driver.dimensionId);
            bits.push(`Driver: ${dim ? UI.escapeHtml(dim.DIMENSION) : "—"}${s.driver.valores.length ? ` (${s.driver.valores.length})` : " (todos)"}`);
        }
        const taskCount = s.bloques.reduce((n, b) => n + b.tareas.length, 0);
        bits.push(`${taskCount} tarea${taskCount === 1 ? "" : "s"}`);
        return bits.join(" · ");
    },

    async addStep() {
        const name = await UI.openTextPromptModal({ title: "Nuevo paso", label: "Nombre del paso", placeholder: "Ej. Carga de datos reales" });
        if (name === null) return;
        if (!name.trim()) { UI.toast("Indica un nombre para el paso.", "error"); return; }
        const step = this.blankStep(name.trim());
        this.editing.steps.push(step);
        this.selectedStepId = step.id;
        this.selectedBlockId = null;
        this.activeTab = "propiedades";
        this.renderStepsPart();
        this.renderStepDetail();
    },

    currentStep() {
        return this.editing.steps.find(s => s.id === this.selectedStepId) || null;
    },

    // ------------------------------------------------------------
    // Detalle del paso seleccionado
    // ------------------------------------------------------------
    renderStepDetail() {
        const part = document.getElementById("wfStepDetailPart");
        const step = this.currentStep();

        if (!step) {
            part.innerHTML = `<div class="module-empty module-empty--inline">Selecciona o añade un paso para configurarlo.</div>`;
            return;
        }

        // Paso 0: sin pestañas, solo variables del workflow.
        if (step.isPaso0) {
            part.innerHTML = `
                <div class="flow-part-header"><strong>Variables del workflow</strong></div>
                <div class="flow-step-detail" id="wfStepTabBody"></div>`;
            this.renderVariablesTab(document.getElementById("wfStepTabBody"), step);
            return;
        }

        const tabs = [
            ["propiedades", "Propiedades"],
            ["tareas", "Tareas del paso"]
        ];
        if (!["propiedades", "tareas"].includes(this.activeTab)) this.activeTab = "propiedades";

        part.innerHTML = `
            <div class="flow-part-header"><strong>Paso: ${UI.escapeHtml(step.name || "(sin nombre)")}</strong></div>
            <div class="segmented" id="wfStepTabs">
                ${tabs.map(([key, label]) => `<button type="button" class="segmented-btn ${this.activeTab === key ? "active" : ""}" data-tab="${key}">${label}</button>`).join("")}
            </div>
            <div class="flow-step-detail" id="wfStepTabBody"></div>`;

        part.querySelectorAll("[data-tab]").forEach(btn => {
            btn.addEventListener("click", () => {
                this.activeTab = btn.dataset.tab;
                this.renderStepDetail();
            });
        });

        const body = document.getElementById("wfStepTabBody");
        if (this.activeTab === "propiedades") this.renderPropiedadesTab(body, step);
        else this.renderTareasTab(body, step);
    },

    // -------------------- Propiedades (nombre + inicio/fin + driver) --------------------
    finOptions(step) {
        const opts = [["NA", "Al finalizar el workflow"]];
        opts.push(step.revision ? ["REVISION", "Al validar todos los procesos"] : ["COMPLETAR", "Al completarse el paso"]);
        opts.push(["FECHA", "Fecha concreta"]);
        return opts;
    },

    renderPropiedadesTab(body, step) {
        const finOpts = this.finOptions(step);
        if (!finOpts.some(([k]) => k === step.fin.tipo)) step.fin.tipo = finOpts[1][0];
        const dim = step.driver.dimensionId ? this.dimensionById(step.driver.dimensionId) : null;

        body.innerHTML = `
            <div class="flow-step-group">
                <div class="form-group">
                    <label>Nombre del paso</label>
                    <input type="text" id="stepName" value="${UI.escapeHtml(step.name)}">
                </div>

                <div class="wf-timing-grid">
                    <div class="wf-timing-card wf-timing-card--start">
                        <div class="wf-timing-label">▶ Inicio</div>
                        <select id="stepInicioTipo">
                            <option value="INICIO_WORKFLOW" ${step.inicio.tipo === "INICIO_WORKFLOW" ? "selected" : ""}>Al iniciar el workflow</option>
                            <option value="PASO_ANTERIOR" ${step.inicio.tipo === "PASO_ANTERIOR" ? "selected" : ""}>Al completar el paso anterior</option>
                            <option value="FECHA" ${step.inicio.tipo === "FECHA" ? "selected" : ""}>Fecha concreta</option>
                        </select>
                        ${step.inicio.tipo === "FECHA" ? `<p class="form-hint">La fecha se indicará al ejecutar el workflow.</p>` : ""}
                        <label class="wf-timing-check"><input type="checkbox" id="stepRevision" ${step.revision ? "checked" : ""}> Requiere revisión</label>
                        ${step.revision ? `<label class="wf-timing-check"><input type="checkbox" id="stepNoBloqueaRevision" ${step.noBloqueaRevision ? "checked" : ""}> No bloquear durante la revisión</label>` : ""}
                    </div>
                    <div class="wf-timing-arrow">→</div>
                    <div class="wf-timing-card wf-timing-card--end">
                        <div class="wf-timing-label">■ Finalización</div>
                        <select id="stepFinTipo">
                            ${finOpts.map(([k, l]) => `<option value="${k}" ${step.fin.tipo === k ? "selected" : ""}>${l}</option>`).join("")}
                        </select>
                        ${step.fin.tipo === "FECHA" ? `<p class="form-hint">La fecha se indicará al ejecutar el workflow.</p>` : ""}
                        <p class="form-hint">Dispara el "al completar el paso anterior" del siguiente paso.</p>
                    </div>
                </div>

                <div class="wf-driver-section">
                    <div class="flow-step-group-title">Driver de reparto</div>
                    <p class="form-hint">Reparte la ejecución de este paso por los valores de una dimensión (ej. nodo superior de la jerarquía de CECOs → una ejecución por CECO, cada una asignable a una persona distinta). Sin driver, el paso se asigna en bloque a una persona o grupo.</p>
                    ${dim ? `
                        <div class="hier-chip-card" style="margin-bottom:12px;">
                            <div>
                                <strong>${UI.escapeHtml(dim.DIMENSION)}</strong>
                                <span class="hier-chip-levels">${step.driver.valores.length ? `${step.driver.valores.length} valor(es) seleccionado(s)` : "Todos los valores de la dimensión"}</span>
                            </div>
                            <div class="row-actions">
                                <button id="btnDriverValues" title="Elegir valores concretos">▤</button>
                                <button id="btnDriverClear" class="danger" title="Quitar driver">🗑</button>
                            </div>
                        </div>
                        ${step.driver.valores.length ? `
                            <div class="chip-row">
                                ${step.driver.valores.map(v => `<span class="hier-chip">${UI.escapeHtml(v)} <a href="#" data-remove-driver-val="${UI.escapeHtml(v)}" style="margin-left:4px;">✕</a></span>`).join("")}
                            </div>` : ""}
                    ` : `<button class="btn btn-secondary btn-sm" id="btnDriverPick">+ Seleccionar dimensión driver</button>`}
                </div>
            </div>`;

        document.getElementById("stepName").addEventListener("input", (e) => {
            step.name = e.target.value;
            this.renderStepsPart();
            document.querySelector("#wfStepDetailPart .flow-part-header strong").textContent = `Paso: ${step.name || "(sin nombre)"}`;
        });
        document.getElementById("stepInicioTipo").addEventListener("change", (e) => {
            step.inicio.tipo = e.target.value;
            this.renderPropiedadesTab(body, step);
        });

        document.getElementById("stepRevision").addEventListener("change", (e) => {
            step.revision = e.target.checked;
            if (!step.revision) step.noBloqueaRevision = false;
            if (step.fin.tipo === "REVISION" || step.fin.tipo === "COMPLETAR") {
                step.fin.tipo = step.revision ? "REVISION" : "COMPLETAR";
            }
            this.renderPropiedadesTab(body, step);
            this.renderStepsPart();
        });
        const noBloqueaRevision = document.getElementById("stepNoBloqueaRevision");
        if (noBloqueaRevision) noBloqueaRevision.addEventListener("change", (e) => { step.noBloqueaRevision = e.target.checked; });

        document.getElementById("stepFinTipo").addEventListener("change", (e) => {
            step.fin.tipo = e.target.value;
            this.renderPropiedadesTab(body, step);
        });

        const pickBtn = document.getElementById("btnDriverPick");
        if (pickBtn) pickBtn.addEventListener("click", async () => {
            const dimId = await UI.openDimensionPickerModal({ dimensionsList: this.dimensions });
            if (!dimId) return;
            step.driver.dimensionId = dimId;
            step.driver.valores = [];
            this.renderPropiedadesTab(body, step);
            this.renderStepsPart();
        });
        const clearBtn = document.getElementById("btnDriverClear");
        if (clearBtn) clearBtn.addEventListener("click", () => {
            step.driver = { dimensionId: null, valores: [] };
            this.renderPropiedadesTab(body, step);
            this.renderStepsPart();
        });
        const valuesBtn = document.getElementById("btnDriverValues");
        if (valuesBtn) valuesBtn.addEventListener("click", async () => {
            const keyCol = Provider.toIdentifier(dim.DIMENSION);
            const result = await UI.openDimensionValuesPickerModal({ project: this.project, dim, keyCol, selected: step.driver.valores });
            if (result === null) return;
            step.driver.valores = result;
            this.renderPropiedadesTab(body, step);
            this.renderStepsPart();
        });
        body.querySelectorAll("[data-remove-driver-val]").forEach(a => {
            a.addEventListener("click", (e) => {
                e.preventDefault();
                step.driver.valores = step.driver.valores.filter(v => v !== a.dataset.removeDriverVal);
                this.renderPropiedadesTab(body, step);
                this.renderStepsPart();
            });
        });
    },

    // -------------------- Variables (solo Paso 0) --------------------
    renderVariablesTab(body, step) {
        body.innerHTML = `
            <div class="flow-step-group">
                <p class="form-hint">Variables de valor único del workflow: se piden al crear cada ejecución y están disponibles para completar tareas en cualquier paso.</p>
                <div class="flow-step-group-title">
                    Variables
                    <button class="btn btn-secondary btn-sm" id="btnAddStepVar" style="float:right;">+ Nueva variable</button>
                </div>
                ${step.variables.length ? `
                    <div class="data-list">
                        <table>
                            <thead><tr><th>Nombre</th><th>Etiqueta</th><th>Tipo</th><th></th></tr></thead>
                            <tbody>
                                ${step.variables.map((v, idx) => `
                                    <tr>
                                        <td><strong>${UI.escapeHtml(v.name)}</strong></td>
                                        <td>${UI.escapeHtml(v.label)}</td>
                                        <td><span class="table-tag">${UI.escapeHtml(v.type)}</span></td>
                                        <td>
                                            <div class="row-actions">
                                                <button data-edit-var="${idx}" title="Editar">✎</button>
                                                <button data-del-var="${idx}" class="danger" title="Eliminar">🗑</button>
                                            </div>
                                        </td>
                                    </tr>`).join("")}
                            </tbody>
                        </table>
                    </div>` : `<div class="module-empty module-empty--inline">Este workflow todavía no tiene variables.</div>`}
            </div>`;

        document.getElementById("btnAddStepVar").addEventListener("click", async () => {
            const v = await UI.openWorkflowVariableModal({});
            if (!v) return;
            step.variables.push({ id: Provider.newId(), ...v });
            this.renderVariablesTab(body, step);
            this.renderStepsPart();
        });
        body.querySelectorAll("[data-edit-var]").forEach(btn => {
            btn.addEventListener("click", async () => {
                const idx = parseInt(btn.dataset.editVar, 10);
                const v = await UI.openWorkflowVariableModal({ current: step.variables[idx] });
                if (!v) return;
                step.variables[idx] = { ...step.variables[idx], ...v };
                this.renderVariablesTab(body, step);
            });
        });
        body.querySelectorAll("[data-del-var]").forEach(btn => {
            btn.addEventListener("click", () => {
                const idx = parseInt(btn.dataset.delVar, 10);
                step.variables.splice(idx, 1);
                this.renderVariablesTab(body, step);
                this.renderStepsPart();
            });
        });
    },

    // Todas las variables del workflow disponibles para mapear tareas (viven en el Paso 0).
    workflowVariables() {
        const paso0 = this.editing.steps.find(s => s.isPaso0);
        return paso0 ? paso0.variables : [];
    },

    // -------------------- Tareas del paso (bloques en menú lateral) --------------------
    renderTareasTab(body, step) {
        if (step.bloques.length && !step.bloques.some(b => b.id === this.selectedBlockId)) {
            this.selectedBlockId = step.bloques[0].id;
        }
        if (!step.bloques.length) this.selectedBlockId = null;

        body.innerHTML = `
            <div class="wf-blocks-layout">
                <div class="wf-block-menu" id="wfBlockMenu">
                    ${step.bloques.map((b, bIdx) => `
                        <div class="wf-block-menu-item ${b.id === this.selectedBlockId ? "active" : ""}" draggable="true" data-block-idx="${bIdx}">
                            <span class="wf-block-menu-item-name">${UI.escapeHtml(b.titulo)}</span>
                            <span class="wf-block-menu-item-count">${b.tareas.length}</span>
                            <button type="button" class="wf-block-menu-item-remove" data-remove-block="${bIdx}" title="Eliminar bloque">✕</button>
                        </div>`).join("")}
                    <button type="button" class="btn btn-secondary btn-sm wf-block-menu-add" id="btnAddBlock">+ Añadir bloque</button>
                </div>
                <div class="wf-block-detail" id="wfBlockDetail"></div>
            </div>`;

        document.getElementById("btnAddBlock").addEventListener("click", () => {
            const block = { id: Provider.newId(), titulo: "Nuevo bloque", tareas: [] };
            step.bloques.push(block);
            this.selectedBlockId = block.id;
            this.renderTareasTab(body, step);
            this.renderStepsPart();
        });

        this.bindBlockMenuEvents(body, step);
        this.renderBlockDetail(body, step);
    },

    bindBlockMenuEvents(body, step) {
        const menu = document.getElementById("wfBlockMenu");

        menu.querySelectorAll("[data-block-idx]").forEach(item => {
            item.addEventListener("click", (e) => {
                if (e.target.closest("[data-remove-block]")) return;
                this.selectedBlockId = step.bloques[parseInt(item.dataset.blockIdx, 10)].id;
                this.renderTareasTab(body, step);
            });
            item.addEventListener("dragstart", () => { this.dragBlockIdx = parseInt(item.dataset.blockIdx, 10); });
            item.addEventListener("dragover", (e) => e.preventDefault());
            item.addEventListener("drop", (e) => {
                e.preventDefault();
                const toIdx = parseInt(item.dataset.blockIdx, 10);
                if (this.dragBlockIdx === null || this.dragBlockIdx === toIdx) return;
                const [moved] = step.bloques.splice(this.dragBlockIdx, 1);
                step.bloques.splice(toIdx, 0, moved);
                this.dragBlockIdx = null;
                this.renderTareasTab(body, step);
            });
        });

        menu.querySelectorAll("[data-remove-block]").forEach(btn => {
            btn.addEventListener("click", async (e) => {
                e.stopPropagation();
                const idx = parseInt(btn.dataset.removeBlock, 10);
                const ok = await UI.confirm("Eliminar bloque", `Se eliminará el bloque <strong>${UI.escapeHtml(step.bloques[idx].titulo)}</strong> y todas sus tareas.`);
                if (!ok) return;
                const removedId = step.bloques[idx].id;
                step.bloques.splice(idx, 1);
                if (this.selectedBlockId === removedId) this.selectedBlockId = step.bloques.length ? step.bloques[0].id : null;
                this.renderTareasTab(body, step);
                this.renderStepsPart();
            });
        });
    },

    renderBlockDetail(body, step) {
        const detail = document.getElementById("wfBlockDetail");
        const block = step.bloques.find(b => b.id === this.selectedBlockId);

        if (!block) {
            detail.innerHTML = `<div class="module-empty module-empty--inline">Añade un bloque para empezar a crear tareas.</div>`;
            return;
        }

        detail.innerHTML = `
            <div class="wf-block-detail-header">
                <span class="modal-title-editable" contenteditable="true" spellcheck="false" id="wfBlockTitleInput">${UI.escapeHtml(block.titulo)}</span>
                <span class="load-fn-toolbar-spacer"></span>
                <button type="button" class="btn btn-primary btn-sm" id="btnAddTask">+ Tarea</button>
            </div>
            <div class="flow-frame-vars">
                ${block.tareas.length ? block.tareas.map((t, tIdx) => this.taskHtml(block, t, tIdx)).join("") : `<p class="form-hint">Sin tareas todavía.</p>`}
            </div>`;

        const titleEl = document.getElementById("wfBlockTitleInput");
        titleEl.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); titleEl.blur(); } });
        titleEl.addEventListener("blur", () => {
            block.titulo = titleEl.textContent.trim() || "Bloque";
            this.renderTareasTab(document.getElementById("wfStepTabBody"), step);
        });

        document.getElementById("btnAddTask").addEventListener("click", async () => {
            await this.addTask(step, block);
            this.renderTareasTab(document.getElementById("wfStepTabBody"), step);
            this.renderStepsPart();
        });

        this.bindTaskEvents(detail, step, block);
    },

    taskHtml(block, task, tIdx) {
        const typeInfo = this.TASK_TYPES[task.tipo] || { label: task.tipo, icon: "•" };
        const canAddCustomVar = task.tipo === "PLANTILLA_EXCEL" || task.tipo === "PLANTILLA_WEB" || task.tipo === "PLANTILLA" || task.tipo === "FUNCION" || task.tipo === "HTML";
        const showRef = task.refNombre && task.refNombre !== task.nombre;
        const varsEmptyMsg = task.tipo === "PARAMETRIZACION" ? "Esta tarea no necesita variables adicionales." : "Sin variables asignadas.";
        return `
            <div class="wf-task-card" data-task-idx="${tIdx}">
                <div class="wf-task-card-header">
                    <span class="wf-task-icon">${typeInfo.icon}</span>
                    <div class="wf-task-title-wrap">
                        <div class="wf-task-name">${UI.escapeHtml(task.nombre || typeInfo.label)}</div>
                        <div class="wf-task-meta">
                            <span class="table-tag">${UI.escapeHtml(typeInfo.label)}</span>
                            ${showRef ? `<span class="wf-task-meta-ref">· ${UI.escapeHtml(task.refNombre)}</span>` : ""}
                        </div>
                    </div>
                    <div class="wf-task-actions">
                        ${canAddCustomVar ? `<button type="button" class="wf-task-action-btn" data-add-taskvar="${tIdx}" title="Añadir variable">+</button>` : ""}
                        <button type="button" class="wf-task-action-btn" data-edit-task="${tIdx}" title="Editar tarea">✎</button>
                        <button type="button" class="wf-task-action-btn wf-task-action-btn--danger" data-remove-task="${tIdx}" title="Eliminar tarea">✕</button>
                    </div>
                </div>
                ${task.descripcion ? `<p class="wf-task-desc">${UI.escapeHtml(task.descripcion)}</p>` : ""}
                ${task.valores.length ? `
                    <div class="wf-task-vars">
                        ${task.valores.map((v, vIdx) => `
                            <div class="flow-target-row" data-edit-taskval="${tIdx}:${vIdx}" style="cursor:pointer;">
                                <span class="flow-target-label">${UI.escapeHtml(v.etiqueta || v.clave)}</span>
                                <span class="table-tag">${v.tipo === "variable" ? "= " + UI.escapeHtml(v.valor || "—") : UI.escapeHtml(v.valor || "(vacío)")}</span>
                                ${v.ocultar ? `<span class="table-tag">Oculta</span>` : ""}
                            </div>`).join("")}
                    </div>` : `<p class="wf-task-vars-empty">${varsEmptyMsg}</p>`}
            </div>`;
    },

    bindTaskEvents(detail, step, block) {
        detail.querySelectorAll("[data-remove-task]").forEach(btn => {
            btn.addEventListener("click", () => {
                const tIdx = parseInt(btn.dataset.removeTask, 10);
                block.tareas.splice(tIdx, 1);
                this.renderTareasTab(document.getElementById("wfStepTabBody"), step);
                this.renderStepsPart();
            });
        });

        detail.querySelectorAll("[data-edit-task]").forEach(btn => {
            btn.addEventListener("click", () => this.editTask(step, block, parseInt(btn.dataset.editTask, 10)));
        });

        detail.querySelectorAll("[data-add-taskvar]").forEach(btn => {
            btn.addEventListener("click", async () => {
                const tIdx = parseInt(btn.dataset.addTaskvar, 10);
                const clave = await UI.openTextPromptModal({ title: "Nueva variable de la tarea", label: "Nombre de la variable", placeholder: "ej. ruta_fichero" });
                if (!clave || !clave.trim()) return;
                const task = block.tareas[tIdx];
                const result = await UI.openWorkflowValueModal({ title: "Asignar valor", targetLabel: clave.trim(), stepVariables: this.workflowVariables() });
                if (!result || result === "remove") return;
                task.valores.push({ clave: clave.trim(), etiqueta: clave.trim(), tipo: result.type, valor: result.value, ocultar: !!result.hidden });
                this.renderTareasTab(document.getElementById("wfStepTabBody"), step);
            });
        });

        detail.querySelectorAll("[data-edit-taskval]").forEach(row => {
            row.addEventListener("click", async () => {
                const [tIdx, vIdx] = row.dataset.editTaskval.split(":").map(n => parseInt(n, 10));
                const task = block.tareas[tIdx];
                const v = task.valores[vIdx];
                const result = await UI.openWorkflowValueModal({
                    title: "Asignar valor", targetLabel: v.etiqueta || v.clave,
                    stepVariables: this.workflowVariables(),
                    current: { type: v.tipo, value: v.valor, hidden: v.ocultar }
                });
                if (result === null) return;
                if (result === "remove") { task.valores.splice(vIdx, 1); }
                else { v.tipo = result.type; v.valor = result.value; v.ocultar = !!result.hidden; }
                this.renderTareasTab(document.getElementById("wfStepTabBody"), step);
            });
        });
    },

    async addTask(step, block) {
        const basics = await UI.openTaskFormModal({
            title: "Nueva tarea",
            types: this.taskTypeChoices()
        });
        if (!basics) return;
        const { name, description, tipo } = basics;
        const typeInfo = this.TASK_TYPES[tipo] || {};

        if (typeInfo.catalog === "FLUJO") {
            const flow = await UI.openFlowManualPickerModal({ flows: this.manualFlows });
            if (!flow) return;
            const screenVars = await this.flowScreenVariables(flow.id);
            block.tareas.push({
                id: Provider.newId(), tipo, nombre: name, descripcion: description, refId: flow.id, refNombre: flow.name,
                valores: screenVars.map(v => ({ clave: v.NOMBRE, etiqueta: v.ETIQUETA || v.NOMBRE, tipo: "constante", valor: "", ocultar: false }))
            });
            return;
        }

        if (typeInfo.catalog === "ACTUALIZACION") {
            await this.loadActualizaciones();
            const picked = await UI.openActualizacionPickerModal({ items: this.actualizaciones });
            if (!picked) return;
            block.tareas.push({ id: Provider.newId(), tipo, nombre: name, descripcion: description, refId: picked.id, refNombre: picked.name, valores: [] });
            return;
        }

        if (typeInfo.catalog === "DIMENSION") {
            await this.loadDimensions();
            const dimId = await UI.openDimensionPickerModal({ dimensionsList: this.dimensions });
            if (!dimId) return;
            const dim = this.dimensionById(dimId);
            block.tareas.push({ id: Provider.newId(), tipo, nombre: name, descripcion: description, refId: dimId, refNombre: dim ? dim.DIMENSION : "", valores: [] });
            return;
        }

        if (typeInfo.catalog === "BUCKET_XLSX") {
            const picked = await UI.openBucketExcelPickerModal();
            if (!picked) return;
            block.tareas.push({
                id: Provider.newId(), tipo, nombre: name, descripcion: description,
                refId: JSON.stringify({ bucket: picked.bucket, name: picked.name }),
                refNombre: picked.name, valores: []
            });
            return;
        }

        // Plantillas Web, Funciones, Páginas HTML: todavía sin catálogo
        // propio — la referencia es el nombre que se acaba de escribir arriba.
        block.tareas.push({ id: Provider.newId(), tipo, nombre: name, descripcion: description, refId: null, refNombre: name, valores: [] });
    },

    /** Edita nombre/descripción de una tarea ya creada y, si es de tipo con catálogo, permite cambiar su referencia. */
    async editTask(step, block, tIdx) {
        const task = block.tareas[tIdx];
        const typeInfo = this.TASK_TYPES[task.tipo] || {};

        const basics = await UI.openTaskFormModal({
            title: "Editar tarea",
            name: task.nombre,
            description: task.descripcion || "",
            tipo: task.tipo,
            types: [{ key: task.tipo, label: typeInfo.label || task.tipo, icon: typeInfo.icon }],
            locked: true
        });
        if (!basics) return;
        task.nombre = basics.name;
        task.descripcion = basics.description;

        if (typeInfo.catalog === "FLUJO") {
            const changeRef = await UI.confirm("Cambiar flujo de carga", `Referencia actual: <strong>${UI.escapeHtml(task.refNombre || "—")}</strong>.<br>¿Quieres elegir otro flujo manual?`);
            if (changeRef) {
                const flow = await UI.openFlowManualPickerModal({ flows: this.manualFlows });
                if (flow) { task.refId = flow.id; task.refNombre = flow.name; }
            }
        } else if (typeInfo.catalog === "ACTUALIZACION") {
            const changeRef = await UI.confirm("Cambiar actualización de tablas", `Referencia actual: <strong>${UI.escapeHtml(task.refNombre || "—")}</strong>.<br>¿Quieres elegir otra actualización?`);
            if (changeRef) {
                await this.loadActualizaciones();
                const picked = await UI.openActualizacionPickerModal({ items: this.actualizaciones });
                if (picked) { task.refId = picked.id; task.refNombre = picked.name; }
            }
        } else if (typeInfo.catalog === "DIMENSION") {
            const changeRef = await UI.confirm("Cambiar dimensión", `Referencia actual: <strong>${UI.escapeHtml(task.refNombre || "—")}</strong>.<br>¿Quieres elegir otra dimensión?`);
            if (changeRef) {
                await this.loadDimensions();
                const dimId = await UI.openDimensionPickerModal({ dimensionsList: this.dimensions });
                if (dimId) {
                    const dim = this.dimensionById(dimId);
                    task.refId = dimId; task.refNombre = dim ? dim.DIMENSION : "";
                }
            }
        } else if (typeInfo.catalog === "BUCKET_XLSX") {
            const changeRef = await UI.confirm("Cambiar plantilla Excel", `Referencia actual: <strong>${UI.escapeHtml(task.refNombre || "—")}</strong>.<br>¿Quieres elegir otra plantilla del bucket?`);
            if (changeRef) {
                const picked = await UI.openBucketExcelPickerModal();
                if (picked) {
                    task.refId = JSON.stringify({ bucket: picked.bucket, name: picked.name });
                    task.refNombre = picked.name;
                }
            }
        } else {
            // Tipos sin catálogo: la "referencia" es simplemente el nombre.
            task.refNombre = task.nombre;
        }

        this.renderTareasTab(document.getElementById("wfStepTabBody"), step);
        this.renderStepsPart();
    },

    // ------------------------------------------------------------
    // Guardado
    // ------------------------------------------------------------
    async save() {
        const wf = this.editing;
        if (!wf.name) { UI.toast("El workflow necesita un nombre.", "error"); return; }
        for (const s of wf.steps) {
            if (!s.name) { UI.toast("Todos los pasos necesitan un nombre.", "error"); return; }
        }

        const btn = document.getElementById("wfFormSave");
        if (btn) { btn.disabled = true; btn.textContent = "Guardando…"; }
        try {
            await this.persist();
            const name = wf.name;
            this.editingIsNew = false;
            this.closeForm();
            await this.loadList();
            this.renderList();
            if (window.Draco && Draco.renderProgress) Draco.renderProgress();
            UI.toast(`Workflow "${name}" guardado.`, "success");
        } catch (err) {
            UI.toast("Error al guardar el workflow: " + err.message, "error");
        } finally {
            if (btn) { btn.disabled = false; btn.textContent = "Guardar workflow"; }
        }
    },

    async persist() {
        const wf = this.editing;
        const id = wf.id;
        const pid = this.project.PROYECTO_ID;

        // 1) Cabecera --------------------------------------------------
        if (this.editingIsNew) {
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS")}
                (WORKFLOW_ID, PROYECTO_ID, WORKFLOW, DESCRIPCION, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                VALUES ('${Provider.esc(id)}', '${Provider.esc(pid)}', '${Provider.esc(wf.name)}', '${Provider.esc(wf.description || "")}',
                        ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`);
        } else {
            await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS")}
                SET WORKFLOW = '${Provider.esc(wf.name)}', DESCRIPCION = '${Provider.esc(wf.description || "")}', FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                WHERE WORKFLOW_ID = '${Provider.esc(id)}'`);
        }

        // 2) Limpieza total de la configuración anterior de pasos -------
        // Antes esto eran hasta 7-8 idas y vueltas SECUENCIALES (2 SELECT
        // para sacar ids + 4 DELETE en un bucle uno detrás de otro + el
        // DELETE final de PASOS). Cada ida y vuelta a BigQuery/Snowflake
        // tiene su propio overhead de arranque de job, así que esto es lo
        // que hacía lento el guardado. Ahora:
        //   a) TAREAS_VALORES se borra con una única subconsulta anidada
        //      (sin ir primero a buscar los PASO_ID y luego los TAREA_ID).
        //      Va sola porque su subconsulta necesita que WORKFLOWS_PASOS_TAREAS
        //      todavía exista (se borra en el paso siguiente).
        //   b) El resto de tablas que cuelgan de PASO_ID no dependen entre
        //      sí, así que se borran en paralelo con Promise.all.
        //   c) WORKFLOWS_PASOS se borra el último, porque las subconsultas
        //      de (a) y (b) necesitan que sus filas sigan existiendo.
        await Provider.runQuery(`
            DELETE FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS_VALORES")}
            WHERE TAREA_ID IN (
                SELECT TAREA_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")}
                WHERE PASO_ID IN (SELECT PASO_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}')
            )`);

        await Promise.all(["WORKFLOWS_PASOS_TAREAS", "WORKFLOWS_PASOS_BLOQUES", "WORKFLOWS_PASOS_VARIABLES", "WORKFLOWS_PASOS_DRIVER_VALORES"].map(table =>
            Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(table)} WHERE PASO_ID IN (SELECT PASO_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}')`)
        ));

        await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}'`);

        if (!wf.steps.length) return;

        // 3) Pasos -------------------------------------------------------
        const pasoVals = wf.steps.map((s, idx) => `(
            '${Provider.esc(s.id)}', '${Provider.esc(id)}', '${Provider.esc(pid)}', '${Provider.esc(s.name)}', ${idx},
            ${s.isPaso0 ? "TRUE" : "FALSE"},
            '${Provider.esc(s.inicio.tipo)}', '${Provider.esc(s.inicio.fecha || "")}', ${s.revision ? "TRUE" : "FALSE"}, ${s.noBloqueaRevision ? "TRUE" : "FALSE"},
            '${Provider.esc(s.fin.tipo)}', '${Provider.esc(s.fin.fecha || "")}',
            ${s.driver.dimensionId ? `'${Provider.esc(s.driver.dimensionId)}'` : "NULL"},
            '${s.driver.dimensionId ? (s.driver.valores.length ? "VALORES" : "TODOS") : "NINGUNO"}'
        )`).join(",\n");

        // 4) Valores del driver -------------------------------------------
        const driverVals = [];
        wf.steps.forEach(s => (s.driver.valores || []).forEach(v => driverVals.push(`('${Provider.esc(pid)}', '${Provider.esc(s.id)}', '${Provider.esc(v)}')`)));

        // 5) Variables (viven en el Paso 0) ---------------------------------------------
        const varVals = [];
        wf.steps.forEach(s => (s.variables || []).forEach((v, idx) =>
            varVals.push(`('${Provider.esc(pid)}', '${Provider.esc(s.id)}', '${Provider.esc(v.id || Provider.newId())}', '${Provider.esc(v.name)}', '${Provider.esc(v.label)}', '${Provider.esc(v.type)}', ${idx})`)));

        // 6) Bloques + tareas + valores de tareas ---------------------------
        const bloqueVals = [], tareaVals = [], tareaValVals = [];
        wf.steps.forEach(s => {
            (s.bloques || []).forEach((b, bIdx) => {
                bloqueVals.push(`('${Provider.esc(pid)}', '${Provider.esc(s.id)}', '${Provider.esc(b.id)}', '${Provider.esc(b.titulo)}', ${bIdx})`);
                (b.tareas || []).forEach((t, tIdx) => {
                    tareaVals.push(`('${Provider.esc(pid)}', '${Provider.esc(s.id)}', '${Provider.esc(b.id)}', '${Provider.esc(t.id)}', '${Provider.esc(t.tipo)}', '${Provider.esc(t.nombre || "")}', '${Provider.esc(t.descripcion || "")}', ${t.refId ? `'${Provider.esc(t.refId)}'` : "NULL"}, '${Provider.esc(t.refNombre || "")}', ${tIdx})`);
                    (t.valores || []).forEach(v => {
                        tareaValVals.push(`('${Provider.esc(pid)}', '${Provider.esc(t.id)}', '${Provider.esc(v.clave)}', '${Provider.esc(v.etiqueta || v.clave)}', '${Provider.esc(v.tipo === "variable" ? "VARIABLE" : "CONSTANTE")}', '${Provider.esc(v.valor || "")}', ${v.ocultar ? "TRUE" : "FALSE"})`);
                    });
                });
            });
        });

        // Los 6 INSERT de aquí abajo son independientes entre sí (ninguno
        // necesita leer lo que acaba de grabar otro: los VALUES ya vienen
        // completos desde el estado en memoria), así que se lanzan todos a
        // la vez con Promise.all en vez de uno detrás de otro.
        const inserts = [
            Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS")}
                (PASO_ID, WORKFLOW_ID, PROYECTO_ID, PASO, ORDEN, ES_PASO0, INICIO_TIPO, INICIO_FECHA, REVISION, NO_BLOQUEA_REVISION, FIN_TIPO, FIN_FECHA, DRIVER_DIMENSION_ID, DRIVER_MODO)
                VALUES ${pasoVals}`)
        ];
        if (driverVals.length) {
            inserts.push(Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_DRIVER_VALORES")} (PROYECTO_ID, PASO_ID, VALOR) VALUES ${driverVals.join(",\n")}`));
        }
        if (varVals.length) {
            inserts.push(Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_VARIABLES")} (PROYECTO_ID, PASO_ID, VARIABLE_ID, NOMBRE, ETIQUETA, TIPO, ORDEN) VALUES ${varVals.join(",\n")}`));
        }
        if (bloqueVals.length) {
            inserts.push(Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_BLOQUES")} (PROYECTO_ID, PASO_ID, BLOQUE_ID, TITULO, ORDEN) VALUES ${bloqueVals.join(",\n")}`));
        }
        if (tareaVals.length) {
            inserts.push(Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")} (PROYECTO_ID, PASO_ID, BLOQUE_ID, TAREA_ID, TIPO, NOMBRE, DESCRIPCION, REF_ID, REF_NOMBRE, ORDEN) VALUES ${tareaVals.join(",\n")}`));
        }
        if (tareaValVals.length) {
            inserts.push(Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS_VALORES")} (PROYECTO_ID, TAREA_ID, CLAVE, ETIQUETA, TIPO, VALOR, OCULTAR) VALUES ${tareaValVals.join(",\n")}`));
        }
        await Promise.all(inserts);
    }
};
