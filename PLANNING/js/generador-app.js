/**
 * ============================================================
 * DRACO PLANNING — INFORMES · GENERADOR (generador.html)
 * ============================================================
 * Controlador de generador.html. Mismo patrón de sesión y de selector de
 * proyecto que task-app.js (task.html):
 *   - Si no hay sesión válida, redirige a index.html.
 *   - El selector solo ELIGE proyecto (no crea ni elimina).
 *
 * La herramienta de informes (informes-herramienta.html) corre en un
 * iframe del mismo origen con ?host=draco. Ella NO toca la base de datos:
 * pide las operaciones por postMessage y este controlador las ejecuta con
 * la sesión del usuario (Informes -> Provider), siempre dentro del
 * proyecto seleccionado:
 *
 *   herramienta -> aquí   { dpTool:1, id, op, payload }   op: list | get | save | delete
 *                         { dpTool:1, op:'dirty', value } (hay cambios sin guardar)
 *                         { dpTool:1, op:'page', payload } (puente futuro: la página
 *                              pide tablas/SQL; todavía no disponible)
 *   aquí -> herramienta   { dpHost:1, id, ok, data | error:{code,message} }
 *
 * Si la sesión caduca a mitad de trabajo, la operación devuelve el error
 * NO_AUTH y la herramienta propone "inicia sesión en otra pestaña y
 * reintenta": como el token se lee de localStorage en cada petición, basta
 * con volver a entrar en Draco en otra pestaña y no se pierde nada.
 */
const GeneradorApp = {
    currentProject: null,   // { PROYECTO_ID, PROYECTO, DESCRIPCION, DATASET }
    projects: [],
    dirty: false,           // la herramienta tiene cambios sin guardar
    pendingInformeId: null, // ?informe_id=... para abrir un informe al cargar

    async init() {
        if (Provider.key() === "snowflake") {
            await SF.ensureConnected();
        }
        if (!Provider.isConnected() || !Provider.isReady()) {
            window.location.href = "index.html";
            return;
        }

        document.getElementById("gcpProjectLabel").textContent = Provider.homeLabel();
        document.getElementById("connStatus").innerHTML = `<span class="dot"></span>${Provider.label()}`;
        this.pendingInformeId = new URLSearchParams(window.location.search).get("informe_id");

        this.bindTopbar();
        this.bindProjectBar();
        window.addEventListener("message", (e) => this.onMessage(e));
        window.addEventListener("beforeunload", (e) => {
            if (this.dirty) { e.preventDefault(); e.returnValue = ""; }
        });

        await this.loadProjects();
    },

    bindTopbar() {
        document.getElementById("btnLogout").addEventListener("click", async () => {
            const extra = this.dirty ? " Hay cambios sin guardar en el informe abierto y se perderán." : "";
            const ok = await UI.confirm("Cerrar sesión", "¿Seguro que quieres cerrar la sesión de Draco Planning?" + extra);
            if (!ok) return;
            this.dirty = false;
            Provider.logout();
            window.location.href = "index.html";
        });
    },

    bindProjectBar() {
        document.getElementById("projectSelect").addEventListener("change", (e) => {
            this.selectProject(e.target.value);
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
                select.innerHTML = `<option value="">No hay proyectos todavía</option>`;
                this.currentProject = null;
                this.showEmpty("No hay proyectos creados todavía. Pide a un administrador que cree uno en Administración.");
                return;
            }

            select.innerHTML = this.projects.map(p =>
                `<option value="${UI.escapeHtml(p.PROYECTO_ID)}">${UI.escapeHtml(p.PROYECTO)}</option>`
            ).join("");

            // Si se abre un informe concreto (?informe_id=), arranca en SU proyecto
            let wantedId = localStorage.getItem("draco_current_project");
            if (this.pendingInformeId) {
                try {
                    const meta = await Informes.getMeta(this.pendingInformeId);
                    if (this.projects.some(p => p.PROYECTO_ID === meta.proyectoId)) wantedId = meta.proyectoId;
                } catch (e) { /* si no se encuentra, se abre el proyecto recordado */ }
            }
            const toSelect = this.projects.find(p => p.PROYECTO_ID === wantedId) || this.projects[0];
            select.value = toSelect.PROYECTO_ID;
            this.selectProject(toSelect.PROYECTO_ID, true);
        } catch (err) {
            select.innerHTML = `<option value="">Error al cargar proyectos</option>`;
            UI.toast("Error al cargar proyectos: " + err.message, "error");
        }
    },

    async selectProject(projectId, first = false) {
        const next = this.projects.find(p => p.PROYECTO_ID === projectId) || null;
        const select = document.getElementById("projectSelect");

        if (!first && this.dirty && this.currentProject && next && next.PROYECTO_ID !== this.currentProject.PROYECTO_ID) {
            const ok = await UI.confirm("Cambiar de proyecto",
                "Hay cambios sin guardar en el informe abierto. Si cambias de proyecto se perderán.");
            if (!ok) { select.value = this.currentProject.PROYECTO_ID; return; }
        }

        this.currentProject = next;
        this.dirty = false;
        if (next) localStorage.setItem("draco_current_project", next.PROYECTO_ID);
        document.getElementById("projectDescription").textContent = next ? (next.DESCRIPCION || "") : "";
        this.mountTool();
    },

    showEmpty(message) {
        const empty = document.getElementById("toolEmpty");
        empty.textContent = message;
        empty.hidden = false;
        document.getElementById("toolFrame").hidden = true;
    },

    /** (Re)carga la herramienta en el iframe, vacía o con el informe pedido en la URL. */
    mountTool() {
        const frame = document.getElementById("toolFrame");
        if (!this.currentProject) { this.showEmpty("Selecciona un proyecto en la barra superior para diseñar sus informes."); return; }
        const qs = new URLSearchParams({ host: "draco" });
        if (this.pendingInformeId) { qs.set("informe_id", this.pendingInformeId); this.pendingInformeId = null; }
        document.getElementById("toolEmpty").hidden = true;
        frame.hidden = false;
        qs.set("v", "20261012c");   // cambia en cada versión para que el navegador no use una herramienta antigua en caché
        frame.src = "informes-herramienta.html?" + qs.toString();
    },

    // ---------------------------------------------------------
    // Puente con la herramienta
    // ---------------------------------------------------------
    async onMessage(e) {
        if (e.origin !== window.location.origin) return;
        const frame = document.getElementById("toolFrame");
        if (!frame || e.source !== frame.contentWindow) return;
        const d = e.data;
        if (!d || !d.dpTool) return;

        if (d.op === "dirty") { this.dirty = !!d.value; return; }
        if (d.op === "ready") return;
        if (d.op === "page") { this.handlePage(d.payload, frame); return; }
        if (!d.id) return;

        const reply = (ok, data, error) => {
            frame.contentWindow.postMessage({ dpHost: 1, id: d.id, ok, data, error }, window.location.origin);
        };
        try {
            reply(true, await this.handle(d.op, d.payload || {}));
        } catch (err) {
            reply(false, null, { code: err.code || "ERR", message: err.message || "Error desconocido" });
        }
    },

    /**
     * Peticiones de la PÁGINA del informe (bloque "Actualización de tablas"):
     * llegan relayadas por la herramienta como { id, op, payload } y se
     * ejecutan aquí con la sesión del usuario (ver js/informes-datos.js).
     */
    async handlePage(pl, frame) {
        const rid = pl && pl.id;
        const reply = (ok, data, error) => frame.contentWindow.postMessage(
            { dpHost: 1, op: "page-reply", id: rid, ok, data, error }, window.location.origin);
        try {
            const op = pl && pl.op;
            reply(true, (typeof op === "string" && op.indexOf("dt.") === 0)
                ? await DracoTables.handle(this.currentProject, op, pl.payload)               // Tablas Draco (vista previa)
                : await InformesDatos.handle(this.currentProject, op, pl && pl.payload));
        } catch (err) {
            reply(false, null, { code: err.code || "ERR", message: err.message || "Error desconocido", details: err.details });
        }
    },

    async handle(op, p) {
        if (!this.currentProject) throw Object.assign(new Error("No hay ningún proyecto seleccionado."), { code: "BAD_INPUT" });
        const pid = this.currentProject.PROYECTO_ID;

        switch (op) {
            case "list":
                return Informes.list(pid);
            case "get": {
                const r = await Informes.get(p.id, pid);
                return { id: r.id, nombre: r.nombre, descripcion: r.descripcion, spec: r.spec };
            }
            case "save":
                return Informes.save({ id: p.id || null, projectId: pid, nombre: p.nombre, descripcion: p.descripcion, spec: p.spec });
            case "delete":
                await Informes.remove(p.id, pid);
                return { id: p.id };
            case "tu.list":
            case "tu.describe":
                return InformesDatos.handle(this.currentProject, op, p);      // el editor elige la vista de una actualización de tablas
            case "dt.list":
            case "dt.design":
                // el editor elige una Tabla Draco o abre su diseñador (pseudo-Excel con el taskpane del addin)
                return DracoTables.handle(this.currentProject, op, p, { openDesigner: (project, ref) => DracoTableDesigner.open(project, ref) });
            default:
                throw Object.assign(new Error("Operación desconocida: " + op), { code: "BAD_INPUT" });
        }
    }
};
