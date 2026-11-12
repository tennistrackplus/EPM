/**
 * ============================================================
 * DRACO — TOKEN DE GIT SOLO EN LA SESIÓN (auditoría F-06)
 * ============================================================
 * El token de acceso al repositorio de modelos (PAT de GitHub/GitLab) ya
 * NO se guarda en la conexión (localStorage, en disco): se guarda en
 * sessionStorage, por conexión, y desaparece al cerrar Excel.
 *
 * Los diálogos "Abrir/Guardar modelo semántico" son ventanas aparte (no
 * comparten sessionStorage con el panel): piden el token al panel que los
 * abrió por la mensajería del Dialog API, igual que el token de Google.
 *
 * Los tokens que versiones anteriores guardaban en la conexión se mueven a
 * la sesión y se borran del disco la primera vez que carga el panel.
 */
const GitSession = {
    PREFIX: "draco_git_pat:",
    REQUEST: "epm-git-pat-request",
    RESPONSE: "epm-git-pat-response",

    get(connId) {
        try { return (connId && sessionStorage.getItem(this.PREFIX + connId)) || ""; } catch (e) { return ""; }
    },

    set(connId, pat) {
        if (!connId) return;
        const v = String(pat || "").trim();
        try {
            if (v) sessionStorage.setItem(this.PREFIX + connId, v);
            else sessionStorage.removeItem(this.PREFIX + connId);
        } catch (e) { /* sin sessionStorage */ }
    },

    /** semanticRepo de la conexión con el token de la sesión. */
    repoConfig(conn) {
        const repo = (conn && conn.config && conn.config.semanticRepo) || null;
        if (!repo) return null;
        return Object.assign({}, repo, { token: this.get(conn.id) });
    },

    /** Mueve a la sesión los tokens guardados en disco por versiones anteriores. */
    migrateLegacy() {
        try {
            if (typeof Connections !== "undefined" && typeof Connections.list === "function") {
                Connections.list().forEach(c => {
                    const repo = c.config && c.config.semanticRepo;
                    if (repo && repo.token) {
                        if (!this.get(c.id)) this.set(c.id, repo.token);
                        const config = Object.assign({}, c.config, { semanticRepo: Object.assign({}, repo, { token: "" }) });
                        Connections.update(c.id, { config });
                    }
                });
            }
            const raw = localStorage.getItem("bq_semantic_repo");
            if (raw) {
                const r = JSON.parse(raw);
                if (r && r.token) {
                    r.token = "";
                    localStorage.setItem("bq_semantic_repo", JSON.stringify(r));
                }
            }
        } catch (e) {
            console.warn("[Draco] No se pudieron migrar los tokens de Git:", e);
        }
    },

    /** En quien abre el diálogo: responde a la petición de tokens. */
    handleDialogMessage(dialog, rawMessage) {
        try {
            const msg = JSON.parse(rawMessage);
            if (!msg || msg.type !== this.REQUEST) return false;
            const pats = {};
            if (typeof Connections !== "undefined" && typeof Connections.list === "function") {
                Connections.list().forEach(c => { const t = this.get(c.id); if (t) pats[c.id] = t; });
            }
            dialog.messageChild(JSON.stringify({ type: this.RESPONSE, pats }));
            return true;
        } catch (e) {
            return false;
        }
    },

    /** En el diálogo: pide los tokens a quien lo abrió (espera un tiempo corto). */
    requestFromOpener(timeoutMs = 1200) {
        return new Promise((resolve) => {
            if (typeof Office === "undefined" || !Office.context || !Office.context.ui
                || typeof Office.context.ui.messageParent !== "function"
                || typeof Office.context.ui.addHandlerAsync !== "function") {
                resolve(false);
                return;
            }
            let done = false;
            const finish = (ok) => { if (!done) { done = true; resolve(ok); } };
            const timer = setTimeout(() => finish(false), timeoutMs);
            try {
                Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
                    try {
                        const msg = JSON.parse(arg.message);
                        if (msg && msg.type === this.RESPONSE) {
                            clearTimeout(timer);
                            Object.keys(msg.pats || {}).forEach(id => this.set(id, msg.pats[id]));
                            finish(true);
                        }
                    } catch (e) { /* otro mensaje */ }
                });
                Office.context.ui.messageParent(JSON.stringify({ type: this.REQUEST }));
            } catch (e) {
                clearTimeout(timer);
                finish(false);
            }
        });
    }
};
window.GitSession = GitSession;
