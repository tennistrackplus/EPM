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
 *   - Sin más parámetros se abre en modo "solo el informe": sin barra de Draco, selector
 *     de proyecto, lista ni barra de herramientas (clase html.vw-full, puesta ya en el <head>
 *     para que no parpadee). Un botón discreto abajo a la derecha permite pasar a pantalla
 *     completa del navegador (Esc para salir) o volver al visor con la lista.
 *   - Con &lista=1 se abre el visor normal. Al abrir un informe desde la lista, la dirección
 *     se escribe con &lista=1 para que recargar mantenga la vista; "Enlace" copia la limpia.
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
    fullMode: false,        // true = solo el informe (enlace directo limpio)
    notFound: false,        // el informe pedido en el enlace no existe
    listLoaded: false,      // la lista de informes ya se ha leído (en modo "solo el informe" se omite hasta que se pide)
    preloaded: null,        // informe ya leído en paralelo con la lista de proyectos

    async init() {
        // Modo backend: renueva el token de Google si hace falta (ver js/google-auth.js).
        if (window.GoogleAuth) await GoogleAuth.resume();
        if (!Provider.isConnected() || !Provider.isReady()) {
            window.location.href = "index.html";
            return;
        }

        document.getElementById("gcpProjectLabel").textContent = Provider.homeLabel();
        document.getElementById("connStatus").innerHTML = `<span class="dot"></span>${Provider.label()}`;
        this.pendingInformeId = new URLSearchParams(window.location.search).get("informe_id");
        this.fullMode = document.documentElement.classList.contains("vw-full");

        this.bindTopbar();
        this.bindProjectBar();
        this.bindViewerBar();
        window.addEventListener("message", (e) => this.onMessage(e));
        if (this.pendingInformeId) this.ensureViewer();     // la herramienta empieza a cargar ya, mientras se consulta la base de datos

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
        document.getElementById("btnFull").addEventListener("click", () => { if (this.currentId) this.setFull(true); });
        document.getElementById("btnLink").addEventListener("click", () => this.copyLink());
        document.getElementById("fabFs").addEventListener("click", () => this.toggleFullscreen());
        document.getElementById("fabExit").addEventListener("click", () => this.setFull(false));
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
            // Con enlace directo, el informe se lee a la vez que la lista de proyectos (antes, una consulta detrás de otra)
            const projP = Provider.runQuery(sql);
            const repP = this.pendingInformeId ? Informes.get(this.pendingInformeId).then(r => ({ r }), e => ({ e })) : null;
            this.projects = await projP;

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
                const res = await repP;
                if (res.r && this.projects.some(p => p.PROYECTO_ID === res.r.proyectoId)) { wantedId = res.r.proyectoId; this.preloaded = res.r; }
                else if (res.r || (res.e && res.e.code === "NOT_FOUND")) { this.notFound = true; this.pendingInformeId = null; }
                /* NO_AUTH u otro error: se deja pendingInformeId y openInforme lo gestiona con su "Reintentar" */
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
        if (this.pendingInformeId && this.fullMode) {
            // Modo "solo el informe": no hace falta la lista ni los favoritos (se cargan si se pasa al visor normal)
            const id = this.pendingInformeId;
            this.pendingInformeId = null;
            this.listLoaded = false;
            this.openInforme(id);
            return;
        }
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
            this.listLoaded = true;
            this.renderList();
            this.updateToolbar();
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

    // ---------------------------------------------------------
    // Modo "solo el informe"
    // ---------------------------------------------------------
    /** Dirección limpia (solo el informe) del informe abierto. */
    cleanLink() {
        return window.location.origin + window.location.pathname + "?informe_id=" + encodeURIComponent(this.currentId);
    },

    /** Mantiene la dirección de la barra coherente con el modo (en modo lista lleva &lista=1). */
    syncUrl() {
        if (!this.currentId) return;
        try { history.replaceState(null, "", "?informe_id=" + encodeURIComponent(this.currentId) + (this.fullMode ? "" : "&lista=1")); } catch (e) { /* sin historial */ }
    },

    setFull(on) {
        if (!on && !this.currentId) { window.location.href = "visor.html"; return; }
        this.fullMode = !!on;
        document.documentElement.classList.toggle("vw-full", this.fullMode);
        if (!on && !this.listLoaded) this.loadList();
        if (!on && document.fullscreenElement) document.exitFullscreen().catch(() => {});
        this.syncUrl();
    },

    toggleFullscreen() {
        if (document.fullscreenElement) { document.exitFullscreen().catch(() => {}); return; }
        document.documentElement.requestFullscreen().catch(() => {
            UI.toast("El navegador no ha permitido la pantalla completa. Prueba con F11.", "info");
        });
    },

    async copyLink() {
        if (!this.currentId) return;
        try {
            await navigator.clipboard.writeText(this.cleanLink());
            UI.toast("Enlace copiado: abre solo el informe.", "success");
        } catch (e) {
            window.prompt("Copia este enlace (abre solo el informe):", this.cleanLink());
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
        this.showViewerMessage(this.notFound
            ? "No se ha encontrado el informe de este enlace (puede que lo hayan eliminado)." + (this.fullMode ? " Pulsa ☰ (abajo a la derecha) para ver la lista de informes." : "")
            : "Selecciona un informe de la lista para abrirlo.");
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
        frame.src = "informes-herramienta.html?mode=viewer&v=20261027a";   // la versión evita una herramienta antigua en caché
        return this.viewerReady;
    },

    async openInforme(id) {
        if (!this.currentProject) return;
        const retry = () => this.openInforme(id);
        this.currentId = id;
        this.notFound = false;
        this.renderList();
        this.updateToolbar();
        this.showViewerMessage("Abriendo informe…");
        try {
            let r = (this.preloaded && this.preloaded.id === id) ? this.preloaded : null;
            this.preloaded = null;
            if (!r) r = await Informes.get(id, this.currentProject.PROYECTO_ID);
            // Mientras se prepara la herramienta y se dibuja la página, ya se leen sus tablas de actualización
            try { InformesDatos.prefetch(this.currentProject, r.spec); } catch (e) { /* sin precarga */ }
            try { DracoTables.prefetch(this.currentProject, r.spec); } catch (e) { /* sin precarga */ }
            await this.ensureViewer();
            const frame = document.getElementById("viewerFrame");
            frame.hidden = false;
            frame.contentWindow.postMessage({ dpHost: 1, op: "load", spec: r.spec, nombre: r.nombre }, window.location.origin);
            document.getElementById("vwEmpty").hidden = true;
            this.syncUrl();
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
            const op = pl && pl.op;
            reply(true, (typeof op === "string" && op.indexOf("fl.") === 0)
                ? await InformesFlujos.handle(this.currentProject, op, pl.payload)            // Flujos de carga
                : (typeof op === "string" && op.indexOf("dt.") === 0)
                ? await DracoTables.handle(this.currentProject, op, pl.payload)               // Tablas Draco
                : await InformesDatos.handle(this.currentProject, op, pl && pl.payload));
        } catch (err) {
            reply(false, null, { code: err.code || "ERR", message: err.message || "Error desconocido", details: err.details });
        }
    }
};
