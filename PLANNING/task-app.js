/**
 * ============================================================
 * DRACO PLANNING — TAREAS (task.html)
 * ============================================================
 * Controlador principal de task.html: mismo patrón de sesión y
 * selector de proyecto que app.js (Draco), pero:
 *   - Si no hay sesión válida, redirige a index.html (igual que app.js).
 *   - El selector de proyecto NO permite crear ni eliminar proyectos
 *     (esos botones no existen en esta pantalla): solo elegir uno de
 *     los ya creados en Administración.
 *   - En vez de renderizar el menú de Administración, delega en
 *     MyTasks (js/my-tasks.js) el panel de "Mis tareas".
 */
const TaskApp = {
    currentProject: null,   // { PROYECTO_ID, PROYECTO, DESCRIPCION, DATASET }
    projects: [],

    async init() {
        // Igual que Draco.init(): si el access token de Snowflake caducó
        // mientras la pantalla estaba cerrada, intenta renovarlo en
        // silencio con el refresh token antes de expulsar a index.html.
        if (Provider.key() === "snowflake") {
            await SF.ensureConnected();
        }

        if (!Provider.isConnected() || !Provider.isReady()) {
            window.location.href = "index.html";
            return;
        }

        document.getElementById("gcpProjectLabel").textContent = Provider.homeLabel();
        document.getElementById("connStatus").innerHTML = `<span class="dot"></span>${Provider.label()}`;
        this.bindTopbar();
        this.bindProjectBar();

        await this.loadProjects();
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
    },

    // Mismo criterio de "proyecto recordado" que Draco.loadProjects(),
    // así que si el usuario ya trabajó en un proyecto desde app.html,
    // esta pantalla arranca directamente sobre el mismo.
    async loadProjects() {
        const select = document.getElementById("projectSelect");
        select.innerHTML = `<option value="">Cargando proyectos...</option>`;
        try {
            const sql = `SELECT PROYECTO_ID, PROYECTO, DESCRIPCION, DATASET
                         FROM ${Provider.qualifyControl("PROYECTOS")}
                         ORDER BY FECHA_CREACION DESC`;
            this.projects = await Provider.runQuery(sql);

            if (!this.projects.length) {
                select.innerHTML = `<option value="">No hay proyectos todavía</option>`;
                this.currentProject = null;
                this.updateProjectDescription();
                MyTasks.renderEmpty(document.getElementById("myTasksRoot"),
                    "No hay proyectos creados todavía. Pide a un administrador que cree uno en Administración.");
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
            // Comparte la misma clave que app.html: cambiar de proyecto en
            // una pantalla lo deja seleccionado también en la otra.
            localStorage.setItem("draco_current_project", projectId);
        }
        this.updateProjectDescription();

        const root = document.getElementById("myTasksRoot");
        if (!this.currentProject) {
            MyTasks.renderEmpty(root, "Selecciona un proyecto en la barra superior para ver tus tareas.");
            return;
        }
        MyTasks.render(root, this.currentProject);
    },

    updateProjectDescription() {
        const el = document.getElementById("projectDescription");
        el.textContent = this.currentProject
            ? (this.currentProject.DESCRIPCION || "Sin descripción") + `  ·  dataset: ${this.currentProject.DATASET}`
            : "";
    }
};

document.addEventListener("DOMContentLoaded", () => TaskApp.init());
