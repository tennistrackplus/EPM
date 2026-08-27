/**
 * ============================================================
 * GitRepo — cliente REST mínimo para explorar un repositorio Git
 * (GitHub / GitLab) desde el navegador, al estilo de cómo Looker
 * conecta un proyecto LookML a un repositorio:
 *   - URL del repositorio + rama + token de acceso (PAT), guardados
 *     en la conexión (ver BQ.getSemanticRepo() / login.html).
 *   - Aquí SOLO se listan carpetas y ficheros .lkml (metadatos): no
 *     se descarga ni se procesa el contenido de ningún fichero.
 * ============================================================
 */
const GitRepo = {

    /**
     * Analiza una URL de repositorio (HTTPS) y extrae { provider, owner, repo }.
     * Soporta:
     *   https://github.com/org/repo(.git)
     *   https://gitlab.com/group/subgroup/repo(.git)
     */
    parseUrl(url) {
        if (!url) return null;
        try {
            const clean = url.trim().replace(/\.git$/i, "").replace(/\/+$/, "");
            const u = new URL(clean);
            const parts = u.pathname.replace(/^\/+/, "").split("/").filter(Boolean);
            if (parts.length < 2) return null;

            if (u.hostname === "github.com") {
                return { provider: "github", owner: parts[0], repo: parts[1] };
            }
            if (u.hostname === "gitlab.com" || u.hostname.startsWith("gitlab.")) {
                // GitLab admite subgrupos: todo menos el último segmento es el "owner" (namespace)
                const repo = parts[parts.length - 1];
                const owner = parts.slice(0, -1).join("/");
                return { provider: "gitlab", owner, repo };
            }
            return null;
        } catch (e) {
            return null;
        }
    },

    /**
     * Lista el contenido de una carpeta del repositorio: [{name, path, type: "file"|"dir"}]
     * Solo se muestran las carpetas y los ficheros terminados en .lkml (el resto se filtra,
     * ya que este selector es exclusivamente para localizar modelos LookML).
     *
     * repoConfig: { type: "github"|"gitlab", url, branch, token }
     */
    async listContents(repoConfig, path = "") {
        const type = repoConfig && repoConfig.type;
        if (type === "github") return this._listGitHub(repoConfig, path);
        if (type === "gitlab") return this._listGitLab(repoConfig, path);
        throw new Error(`Tipo de repositorio no soportado todavía para listar: "${type || "(ninguno)"}"`);
    },

    async _listGitHub(repoConfig, path) {
        const ref = this.parseUrl(repoConfig.url);
        if (!ref) throw new Error("URL de repositorio de GitHub no válida.");

        const branch = (repoConfig.branch || "").trim();
        const cleanPath = (path || "").replace(/^\/+|\/+$/g, "");
        let apiUrl = `https://api.github.com/repos/${ref.owner}/${ref.repo}/contents/${encodeURI(cleanPath)}`;
        if (branch) apiUrl += `?ref=${encodeURIComponent(branch)}`;

        const headers = { Accept: "application/vnd.github+json" };
        if (repoConfig.token) headers.Authorization = `Bearer ${repoConfig.token}`;

        const res = await fetch(apiUrl, { headers });
        if (!res.ok) throw new Error(this._friendlyError(res.status, "GitHub"));

        const data = await res.json();
        const items = Array.isArray(data) ? data : [data];

        return items
            .filter(item => item.type === "dir" || (item.type === "file" && item.name.toLowerCase().endsWith(".lkml")))
            .map(item => ({ name: item.name, path: item.path, type: item.type === "dir" ? "dir" : "file" }))
            .sort(this._sortDirsFirst);
    },

    async _listGitLab(repoConfig, path) {
        const ref = this.parseUrl(repoConfig.url);
        if (!ref) throw new Error("URL de repositorio de GitLab no válida.");

        const projectId = encodeURIComponent(`${ref.owner}/${ref.repo}`);
        const branch = (repoConfig.branch || "").trim();
        const cleanPath = (path || "").replace(/^\/+|\/+$/g, "");

        let apiUrl = `https://gitlab.com/api/v4/projects/${projectId}/repository/tree?per_page=100`;
        if (cleanPath) apiUrl += `&path=${encodeURIComponent(cleanPath)}`;
        if (branch) apiUrl += `&ref=${encodeURIComponent(branch)}`;

        const headers = {};
        if (repoConfig.token) headers["PRIVATE-TOKEN"] = repoConfig.token;

        const res = await fetch(apiUrl, { headers });
        if (!res.ok) throw new Error(this._friendlyError(res.status, "GitLab"));

        const data = await res.json();

        return (Array.isArray(data) ? data : [])
            .filter(item => item.type === "tree" || (item.type === "blob" && item.name.toLowerCase().endsWith(".lkml")))
            .map(item => ({ name: item.name, path: item.path, type: item.type === "tree" ? "dir" : "file" }))
            .sort(this._sortDirsFirst);
    },

    _sortDirsFirst(a, b) {
        if (a.type !== b.type) return a.type === "dir" ? -1 : 1;
        return a.name.localeCompare(b.name);
    },

    _friendlyError(status, providerLabel) {
        if (status === 401 || status === 403) {
            return `${providerLabel}: acceso denegado (${status}). Revisa el token de acceso de la conexión.`;
        }
        if (status === 404) {
            return `${providerLabel}: repositorio, rama o ruta no encontrados (404).`;
        }
        return `${providerLabel}: no se pudo listar el repositorio (HTTP ${status}).`;
    }
};
