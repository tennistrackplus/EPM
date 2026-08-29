/**
 * Controlador principal de app.html: sesión, selector de proyecto,
 * enrutado del menú de Administración y panel de progreso.
 * Funciona igual sobre BigQuery o sobre Snowflake a través de Provider.
 */
const Draco = {
    currentProject: null,   // { PROYECTO_ID, PROYECTO, DESCRIPCION, DATASET }
    projects: [],
    currentModule: "dimensiones",

    async init() {
        // Si el access token de Snowflake caducó mientras la app estaba
        // abierta (o al recargar), intenta renovarlo en silencio con el
        // refresh token antes de expulsar al usuario a index.html.
        if (Provider.key() === "snowflake") {
            await SF.ensureConnected();
        }

        if (!Provider.isConnected() || !Provider.isReady()) {
            window.location.href = "index.html";
            return;
        }

        document.getElementById("gcpProjectLabel").textContent = Provider.homeLabel();
        document.getElementById("connStatus").innerHTML = `<span class="dot"></span>${Provider.label()}`;
        UI.initBlockControls();
        this.bindTopbar();
        this.bindProjectBar();
        this.bindAdminMenu();

        await this.loadProjects();
        this.renderModule(this.currentModule);
    },

    bindTopbar() {
        document.getElementById("btnLogout").addEventListener("click", async () => {
            const ok = await UI.confirm("Cerrar sesión", "¿Seguro que quieres cerrar la sesión de Draco Planning?");
            if (!ok) return;
            Provider.logout();
            window.location.href = "index.html";
        });
    },

    bindProjectBar() {
        document.getElementById("projectSelect").addEventListener("change", (e) => {
            this.selectProject(e.target.value);
        });

        document.getElementById("btnNewProject").addEventListener("click", () => this.openNewProjectModal());
        document.getElementById("btnDeleteProject").addEventListener("click", () => this.deleteCurrentProject());
        document.getElementById("closeProjectModal").addEventListener("click", () => UI.closeModal("projectModal"));
        document.getElementById("cancelNewProject").addEventListener("click", () => UI.closeModal("projectModal"));
        document.getElementById("saveNewProject").addEventListener("click", () => this.saveNewProject());

        document.getElementById("newProjectName").addEventListener("input", (e) => {
            const ident = Provider.toIdentifier(e.target.value);
            const preview = document.getElementById("newProjectDatasetPreview");
            preview.innerHTML = ident
                ? `Se creará el dataset/esquema <code>${DracoConfig.prefix}${ident}</code>`
                : "";
        });
    },

    bindAdminMenu() {
        document.querySelectorAll(".admin-menu-item").forEach(btn => {
            btn.addEventListener("click", () => {
                document.querySelectorAll(".admin-menu-item").forEach(b => b.classList.remove("active"));
                btn.classList.add("active");
                this.renderModule(btn.dataset.module);
            });
        });
    },

    async loadProjects() {
        const select = document.getElementById("projectSelect");
        select.innerHTML = `<option value="">Cargando proyectos...</option>`;
        try {
            const sql = `SELECT PROYECTO_ID, PROYECTO, DESCRIPCION, DATASET
                         FROM ${Provider.qualifyControl("PROYECTOS")}
                         ORDER BY FECHA_CREACION DESC`;
            this.projects = await Provider.runQuery(sql);

            if (!this.projects.length) {
                select.innerHTML = `<option value="">No hay proyectos todavía — crea uno</option>`;
                this.currentProject = null;
                this.updateProjectDescription();
                this.renderModule(this.currentModule);
                this.renderProgress();
                return;
            }

            select.innerHTML = this.projects.map(p =>
                `<option value="${p.PROYECTO_ID}">${UI.escapeHtml(p.PROYECTO)}</option>`
            ).join("");

            const savedId = localStorage.getItem("draco_current_project");
            const toSelect = this.projects.find(p => p.PROYECTO_ID === savedId) || this.projects[0];
            select.value = toSelect.PROYECTO_ID;
            this.selectProject(toSelect.PROYECTO_ID);
        } catch (err) {
            select.innerHTML = `<option value="">Error al cargar proyectos</option>`;
            UI.toast("Error al cargar proyectos: " + err.message, "error");
        }
    },

    selectProject(projectId) {
        this.currentProject = this.projects.find(p => p.PROYECTO_ID === projectId) || null;
        if (this.currentProject) {
            localStorage.setItem("draco_current_project", projectId);
        }
        this.updateProjectDescription();
        this.renderModule(this.currentModule);
        this.renderProgress();
    },

    updateProjectDescription() {
        const el = document.getElementById("projectDescription");
        el.textContent = this.currentProject
            ? (this.currentProject.DESCRIPCION || "Sin descripción") + `  ·  dataset: ${this.currentProject.DATASET}`
            : "";
    },

    openNewProjectModal() {
        document.getElementById("newProjectName").value = "";
        document.getElementById("newProjectDesc").value = "";
        document.getElementById("newProjectDatasetPreview").innerHTML = "";
        UI.openModal("projectModal");
        setTimeout(() => document.getElementById("newProjectName").focus(), 50);
    },

    async saveNewProject() {
        const name = document.getElementById("newProjectName").value.trim();
        const desc = document.getElementById("newProjectDesc").value.trim();
        const btn = document.getElementById("saveNewProject");

        if (!name) {
            UI.toast("Indica un nombre para el proyecto.", "error");
            return;
        }

        const ident = Provider.toIdentifier(name);
        if (!ident) {
            UI.toast("El nombre del proyecto debe contener letras o números.", "error");
            return;
        }

        const datasetId = `${DracoConfig.prefix}${ident}`;

        btn.disabled = true;
        btn.textContent = "Creando...";
        try {
            const exists = await Provider.containerExists(datasetId);
            if (exists) {
                UI.toast(`El dataset/esquema ${datasetId} ya existe. Elige otro nombre.`, "error");
                return;
            }

            await Provider.createContainer(datasetId, desc);

            const projectId = Provider.newId();
            const sql = `INSERT INTO ${Provider.qualifyControl("PROYECTOS")}
                (PROYECTO_ID, PROYECTO, DESCRIPCION, DATASET, USUARIO, FECHA_CREACION)
                VALUES ('${Provider.esc(projectId)}', '${Provider.esc(name)}', '${Provider.esc(desc)}', '${Provider.esc(datasetId)}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP())`;
            await Provider.runQuery(sql);

            UI.toast(`Proyecto "${name}" creado correctamente.`, "success");
            UI.closeModal("projectModal");
            localStorage.setItem("draco_current_project", projectId);
            await this.loadProjects();
        } catch (err) {
            UI.toast("Error al crear el proyecto: " + err.message, "error");
        } finally {
            btn.disabled = false;
            btn.textContent = "Crear proyecto";
        }
    },

    async deleteCurrentProject() {
        if (!this.currentProject) {
            UI.toast("Selecciona primero un proyecto.", "error");
            return;
        }
        const p = this.currentProject;
        const ok = await UI.confirm(
            "Eliminar proyecto",
            `Vas a eliminar el proyecto <strong>${UI.escapeHtml(p.PROYECTO)}</strong> y su dataset/esquema <strong>${p.DATASET}</strong> (con todas sus tablas y datos). Esta acción no se puede deshacer.`
        );
        if (!ok) return;

        try {
            await Provider.deleteContainer(p.DATASET);
            const del = (table, col) => Provider.runQuery(
                `DELETE FROM ${Provider.qualifyControl(table)} WHERE ${col} = '${Provider.esc(p.PROYECTO_ID)}'`);
            await del("DIMENSIONES", "PROYECTO_ID");
            await del("CUBOS", "PROYECTO_ID");
            await del("PARAMETRIZACIONES", "PROYECTO_ID");
            await del("PROYECTOS", "PROYECTO_ID");

            UI.toast(`Proyecto "${p.PROYECTO}" eliminado.`, "success");
            localStorage.removeItem("draco_current_project");
            await this.loadProjects();
        } catch (err) {
            UI.toast("Error al eliminar el proyecto: " + err.message, "error");
        }
    },

    renderModule(moduleKey) {
        this.currentModule = moduleKey;
        const container = document.getElementById("adminContent");

        if (!this.currentProject) {
            container.innerHTML = `
                <div class="module-empty">
                    Selecciona o crea un proyecto en la barra superior para empezar a trabajar.
                </div>`;
            return;
        }

        const modules = {
            dimensiones: () => Dimensions.render(container, this.currentProject),
            cubos: () => Cubes.render(container, this.currentProject),
            parametrizacion: () => Parametrizacion.render(container, this.currentProject),
            actualizaciones: () => TableUpdates.render(container, this.currentProject),
            cargas: () => Loads.render(container, this.currentProject),
            "flujos-carga": () => Flows.render(container, this.currentProject),
            workflows: () => Workflows.render(container, this.currentProject),
            widgets: () => Widgets.render(container, this.currentProject),
        };

        const labels = {
            funciones: ["Funciones", "ƒ"],
            "flujos-proceso": ["Flujos de proceso", "⟲"],
            roles: ["Roles", "☺"]
        };

        if (modules[moduleKey]) {
            modules[moduleKey]();
        } else if (labels[moduleKey]) {
            const [label, icon] = labels[moduleKey];
            container.innerHTML = `
                <div class="module-placeholder">
                    <div class="placeholder-icon">${icon}</div>
                    <h3>${label}</h3>
                    <p>Este módulo estará disponible próximamente en Draco Planning.</p>
                </div>`;
        }
    },

    async renderProgress() {
        const panel = document.getElementById("progressPanel");
        if (!this.currentProject) {
            panel.innerHTML = `<p class="progress-empty">Selecciona o crea un proyecto para ver tu progreso.</p>`;
            return;
        }

        const pid = this.currentProject.PROYECTO_ID;
        let dimCount = 0, cuboCount = 0;

        try {
            const r1 = await Provider.runQuery(`SELECT COUNT(*) AS N FROM ${Provider.qualifyControl("DIMENSIONES")} WHERE PROYECTO_ID='${Provider.esc(pid)}'`);
            dimCount = parseInt(r1[0]?.N || "0", 10);
            const r2 = await Provider.runQuery(`SELECT COUNT(*) AS N FROM ${Provider.qualifyControl("CUBOS")} WHERE PROYECTO_ID='${Provider.esc(pid)}'`);
            cuboCount = parseInt(r2[0]?.N || "0", 10);
        } catch (err) {
            panel.innerHTML = `<p class="progress-empty">No se pudo calcular el progreso: ${UI.escapeHtml(err.message)}</p>`;
            return;
        }

        let loadsCount = 0, flowsCount = 0;
        try {
            const r3 = await Provider.runQuery(`SELECT COUNT(*) AS N FROM ${Provider.qualifyControl("INTERFACES")} WHERE PROYECTO_ID='${Provider.esc(pid)}'`);
            loadsCount = parseInt(r3[0]?.N || "0", 10);
        } catch (e) { loadsCount = 0; }
        try {
            const r4 = await Provider.runQuery(`SELECT COUNT(*) AS N FROM ${Provider.qualifyControl("FLUJOS")} WHERE PROYECTO_ID='${Provider.esc(pid)}'`);
            flowsCount = parseInt(r4[0]?.N || "0", 10);
        } catch (e) { flowsCount = 0; }

        let workflowsCount = 0;
        try {
            const r5 = await Provider.runQuery(`SELECT COUNT(*) AS N FROM ${Provider.qualifyControl("WORKFLOWS")} WHERE PROYECTO_ID='${Provider.esc(pid)}'`);
            workflowsCount = parseInt(r5[0]?.N || "0", 10);
        } catch (e) { workflowsCount = 0; }

        const steps = [
            { done: true, title: "Crear proyecto", desc: `Proyecto "${this.currentProject.PROYECTO}" creado` },
            { done: dimCount > 0, title: "Definir dimensiones", desc: dimCount > 0 ? `${dimCount} dimensión(es) definida(s)` : "Aún no has creado ninguna dimensión" },
            { done: cuboCount > 0, title: "Definir cubos", desc: cuboCount > 0 ? `${cuboCount} cubo(s) definido(s)` : "Aún no has creado ningún cubo" },
            { done: loadsCount > 0, title: "Configurar interfaces", desc: loadsCount > 0 ? `${loadsCount} interfaz(ces) definida(s)` : "Aún no has creado ninguna interfaz" },
            { done: flowsCount > 0, title: "Diseñar flujos de carga", desc: flowsCount > 0 ? `${flowsCount} flujo(s) definido(s)` : "Aún no has creado ningún flujo" },
            { done: workflowsCount > 0, title: "Definir workflows", desc: workflowsCount > 0 ? `${workflowsCount} workflow(s) definido(s)` : "Aún no has creado ningún workflow" }
        ];

        const doneCount = steps.filter(s => s.done).length;
        const pct = Math.round((doneCount / steps.length) * 100);

        panel.innerHTML = `
            <div class="progress-summary">
                ${this.progressRingSvg(pct)}
                <div class="progress-summary-text">
                    <strong>${pct}% completado</strong>
                    <span>${doneCount} de ${steps.length} pasos listos</span>
                </div>
            </div>
            <div class="progress-steps">
                ${steps.map((s, i) => `
                    <div class="progress-step ${s.done ? "done" : ""}">
                        <div class="progress-step-check">${s.done ? "✓" : i + 1}</div>
                        <div class="progress-step-body">
                            <strong>${UI.escapeHtml(s.title)}</strong>
                            <span>${UI.escapeHtml(s.desc)}</span>
                        </div>
                    </div>
                `).join("")}
            </div>
        `;
    },

    progressRingSvg(pct) {
        const r = 24, c = 2 * Math.PI * r;
        const offset = c - (pct / 100) * c;
        return `
            <svg class="progress-ring" viewBox="0 0 56 56">
                <circle cx="28" cy="28" r="${r}" fill="none" stroke="var(--gray-200)" stroke-width="6"/>
                <circle cx="28" cy="28" r="${r}" fill="none" stroke="var(--brand-accent)" stroke-width="6"
                        stroke-dasharray="${c}" stroke-dashoffset="${offset}" stroke-linecap="round"
                        transform="rotate(-90 28 28)"/>
            </svg>`;
    }
};

document.addEventListener("DOMContentLoaded", () => Draco.init());
