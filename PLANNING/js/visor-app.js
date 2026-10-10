/**
 * ============================================================
 * DRACO PLANNING — INFORMES · VISOR (visor.html)
 * ============================================================
 * Lista los informes del proyecto (con buscador y favoritos del usuario)
 * y abre el elegido. Mismo patrón de sesión y de proyecto que task-app.js.
 *
 * El informe NO se guarda como HTML: aquí se lee su definición (JSON) de
 * DRACO_CONTROL.INFORMES y se le entrega a la herramienta de informes
 * (informes-herramienta.html?mode=viewer, mismo origen), que genera el
 * HTML en ese momento. La página generada corre dentro de otro iframe
 * aislado (sandbox, sin acceso a la sesión de Draco): cuando necesite
 * datos o abrir tablas/SQL (a futuro), lo pedirá por mensajes y será este
 * controlador quien lo ejecute con la sesión del usuario.
 *
 *   aquí -> visor    { dpHost:1, op:'load', spec, nombre }
 *   visor -> aquí    { dpTool:1, op:'ready' | 'loaded' | 'error' | 'page' }
 *
 * Enlace directo a un informe: visor.html?informe_id=<id>
 */
const VisorApp = {
    currentProject: null,   // { PROYECTO_ID, PROYECTO, DESCRIPCION, DATASET }
    projects: [],
    list: [],               // [{ id, nombre, descripcion, usuario, modificado, fav }]
    currentId: null,
    pendingInformeId: null,
    viewerReady: null,      // Promise que se resuelve cuando la herramienta avisa "ready"
    _resolveReady: null,
    sideHidden: false,

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
        this.bindViewerBar();
        window.addEventListener("message", (e) => this.onMessage(e));

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

    bindViewerBar() {
        document.getElementById("vwSearch").addEventListener("input", () => this.renderList());
        document.getElementById("vwList").addEventListener("click", (e) => this.onListClick(e));
        document.getElementById("btnToggleSide").addEventListener("click", () => this.toggleSide());
        document.getElementById("btnFav").addEventListener("click", () => { if (this.currentId) this.toggleFavorite(this.currentId); });
        document.getElementById("btnReload").addEventListener("click", () => { if (this.currentId) this.openInforme(this.currentId); });
        document.getElementById("btnEditar").addEventListener("click", () => {
            if (this.currentId) window.open("generador.html?informe_id=" + encodeURIComponent(this.currentId), "_blank");
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
                this.renderListMessage("No hay proyectos creados todavía. Pide a un administrador que cree uno en Administración.");
                return;
            }

            select.innerHTML = this.projects.map(p =>
                `<option value="${UI.escapeHtml(p.PROYECTO_ID)}">${UI.escapeHtml(p.PROYECTO)}</option>`
            ).join("");

            // Enlace directo: arranca en el proyecto del informe pedido
            let wantedId = localStorage.getItem("draco_current_project");
            if (this.pendingInformeId) {
                try {
                    const meta = await Informes.getMeta(this.pendingInformeId);
                    if (this.projects.some(p => p.PROYECTO_ID === meta.proyectoId)) wantedId = meta.proyectoId;
                } catch (e) { this.showViewerMessage("No se ha encontrado el informe pedido en el enlace."); this.pendingInformeId = null; }
            }
            const toSelect = this.projects.find(p => p.PROYECTO_ID === wantedId) || this.projects[0];
            select.value = toSelect.PROYECTO_ID;
            this.selectProject(toSelect.PROYECTO_ID);
        } catch (err) {
            select.innerHTML = `<option value="">Error al cargar proyectos</option>`;
            UI.toast("Error al cargar proyectos: " + err.message, "error");
        }
    },

    async selectProject(projectId) {
        this.currentProject = this.projects.find(p => p.PROYECTO_ID === projectId) || null;
        document.getElementById("projectDescription").textContent = this.currentProject ? (this.currentProject.DESCRIPCION || "") : "";
        if (this.currentProject) localStorage.setItem("draco_current_project", this.currentProject.PROYECTO_ID);
        this.currentId = null;
        this.clearViewer();
        await this.loadList();
        if (this.pendingInformeId) {
            const id = this.pendingInformeId;
            this.pendingInformeId = null;
            if (this.list.some(r => r.id === id)) this.openInforme(id);
        }
    },

    // ---------------------------------------------------------
    // Lista
    // ---------------------------------------------------------
    async loadList() {
        if (!this.currentProject) { this.renderListMessage("Selecciona un proyecto."); return; }
        this.renderListMessage("Cargando informes...", true);
        try {
            const pid = this.currentProject.PROYECTO_ID;
            const [rows, favs] = await Promise.all([Informes.list(pid), Informes.favoriteIds(pid)]);
            this.list = rows.map(r => ({ ...r, fav: favs.has(r.id) }));
            this.renderList();
        } catch (err) {
            this.renderListError(err, () => this.loadList());
        }
    },

    renderListMessage(text, spinner = false) {
        document.getElementById("vwList").innerHTML =
            `<div class="module-empty">${spinner ? '<span class="spinner"></span> ' : ""}${UI.escapeHtml(text)}</div>`;
    },

    renderListError(err, retry) {
        const wrap = document.getElementById("vwList");
        const noAuth = err && err.code === "NO_AUTH";
        wrap.innerHTML = `<div class="module-empty">
            ${UI.escapeHtml(noAuth ? err.message + " Inicia sesión en Draco en otra pestaña y pulsa Reintentar." : "Error al cargar los informes: " + err.message)}
            <div style="margin-top:10px;"><button class="btn btn-secondary btn-sm" id="vwRetry">Reintentar</button></div></div>`;
        document.getElementById("vwRetry").addEventListener("click", retry);
    },

    renderList() {
        const q = (document.getElementById("vwSearch").value || "").trim().toLowerCase();
        const match = (r) => !q || (`${r.nombre} ${r.descripcion}`).toLowerCase().includes(q);
        const filtered = this.list.filter(match);
        const wrap = document.getElementById("vwList");

        if (!this.list.length) {
            wrap.innerHTML = `<div class="module-empty">Todavía no hay informes en este proyecto. Crea el primero en el <a href="generador.html">generador</a>.</div>`;
            return;
        }
        if (!filtered.length) {
            wrap.innerHTML = `<div class="module-empty">Ningún informe coincide con la búsqueda.</div>`;
            return;
        }

        const favs = filtered.filter(r => r.fav);
        const item = (r) => `
            <div class="vw-item${r.id === this.currentId ? " is-active" : ""}" data-id="${UI.escapeHtml(r.id)}" role="button" tabindex="0">
                <button class="vw-star${r.fav ? " is-on" : ""}" data-star="${UI.escapeHtml(r.id)}" title="${r.fav ? "Quitar de favoritos" : "Añadir a favoritos"}">${r.fav ? "★" : "☆"}</button>
                <div class="vw-item-main">
                    <div class="vw-item-name">${UI.escapeHtml(r.nombre)}</div>
                    <div class="vw-item-meta">${UI.escapeHtml((r.descripcion ? r.descripcion + " · " : "") + (r.modificado || ""))}</div>
                </div>
            </div>`;

        wrap.innerHTML =
            (favs.length ? `<div class="vw-group">Favoritos</div>${favs.map(item).join("")}` : "") +
            `<div class="vw-group">${favs.length ? "Todos los informes" : "Informes"}</div>${filtered.map(item).join("")}`;
    },

    onListClick(e) {
        const star = e.target.closest("[data-star]");
        if (star) { e.stopPropagation(); this.toggleFavorite(star.getAttribute("data-star")); return; }
        const it = e.target.closest("[data-id]");
        if (it) this.openInforme(it.getAttribute("data-id"));
    },

    async toggleFavorite(id) {
        const r = this.list.find(x => x.id === id);
        if (!r || !this.currentProject) return;
        try {
            r.fav = await Informes.setFavorite(this.currentProject.PROYECTO_ID, id, !r.fav);
            this.renderList();
            this.updateToolbar();
        } catch (err) {
            UI.toast(err.code === "NO_AUTH" ? err.message + " Inicia sesión en Draco en otra pestaña y vuelve a intentarlo." : "No se pudo actualizar el favorito: " + err.message, "error");
        }
    },

    toggleSide() {
        this.sideHidden = !this.sideHidden;
        document.getElementById("vwLayout").classList.toggle("side-hidden", this.sideHidden);
    },

    // ---------------------------------------------------------
    // Abrir un informe
    // ---------------------------------------------------------
    clearViewer() {
        document.getElementById("viewerFrame").hidden = true;
        document.getElementById("vwToolbar").hidden = true;
        this.showViewerMessage("Selecciona un informe de la lista para abrirlo.");
    },

    showViewerMessage(text, retry) {
        const el = document.getElementById("vwEmpty");
        el.hidden = false;
        el.innerHTML = UI.escapeHtml(text) + (retry ? `<div style="margin-top:12px;"><button class="btn btn-secondary btn-sm" id="vwViewerRetry">Reintentar</button></div>` : "");
        if (retry) document.getElementById("vwViewerRetry").addEventListener("click", retry);
    },

    updateToolbar() {
        const r = this.list.find(x => x.id === this.currentId);
        const bar = document.getElementById("vwToolbar");
        bar.hidden = !r;
        if (!r) return;
        document.getElementById("vwTitle").textContent = r.nombre;
        const fav = document.getElementById("btnFav");
        fav.textContent = r.fav ? "★ Favorito" : "☆ Favorito";
        fav.classList.toggle("is-on", !!r.fav);
    },

    /** Crea el iframe de la herramienta en modo visor (una sola vez) y espera a que esté listo. */
    ensureViewer() {
        const frame = document.getElementById("viewerFrame");
        if (this.viewerReady) return this.viewerReady;
        this.viewerReady = new Promise((resolve) => { this._resolveReady = resolve; });
        frame.src = "informes-herramienta.html?mode=viewer";
        return this.viewerReady;
    },

    async openInforme(id) {
        if (!this.currentProject) return;
        const retry = () => this.openInforme(id);
        this.currentId = id;
        this.renderList();
        this.updateToolbar();
        this.showViewerMessage("Abriendo informe…");
        try {
            const r = await Informes.get(id, this.currentProject.PROYECTO_ID);
            await this.ensureViewer();
            const frame = document.getElementById("viewerFrame");
            frame.hidden = false;
            frame.contentWindow.postMessage({ dpHost: 1, op: "load", spec: r.spec, nombre: r.nombre }, window.location.origin);
            document.getElementById("vwEmpty").hidden = true;
            try { history.replaceState(null, "", "?informe_id=" + encodeURIComponent(id)); } catch (e) { /* sin historial */ }
        } catch (err) {
            document.getElementById("viewerFrame").hidden = true;
            if (err.code === "NO_AUTH") {
                this.showViewerMessage(err.message + " Inicia sesión en Draco en otra pestaña y pulsa Reintentar.", retry);
            } else if (err.code === "NOT_FOUND") {
                this.showViewerMessage("Este informe ya no existe (puede que lo hayan eliminado).");
                this.loadList();
            } else {
                this.showViewerMessage("No se pudo abrir el informe: " + err.message, retry);
            }
        }
    },

    // ---------------------------------------------------------
    // Mensajes de la herramienta (modo visor)
    // ---------------------------------------------------------
    onMessage(e) {
        if (e.origin !== window.location.origin) return;
        const frame = document.getElementById("viewerFrame");
        if (!frame || e.source !== frame.contentWindow) return;
        const d = e.data;
        if (!d || !d.dpTool) return;

        if (d.op === "ready" && this._resolveReady) { this._resolveReady(); this._resolveReady = null; }
        else if (d.op === "error") { UI.toast("No se pudo mostrar el informe: " + (d.message || "definición no válida"), "error"); }
        else if (d.op === "page") { this.handlePage(d.payload, frame); }
    },

    /** Peticiones de la página del informe (bloque "Actualización de tablas"): se ejecutan aquí con la sesión del usuario. */
    async handlePage(pl, frame) {
        const rid = pl && pl.id;
        const reply = (ok, data, error) => frame.contentWindow.postMessage(
            { dpHost: 1, op: "page-reply", id: rid, ok, data, error }, window.location.origin);
        try {
            reply(true, await InformesDatos.handle(this.currentProject, pl && pl.op, pl && pl.payload));
        } catch (err) {
            reply(false, null, { code: err.code || "ERR", message: err.message || "Error desconocido", details: err.details });
        }
    }
};
