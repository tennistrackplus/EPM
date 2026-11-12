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

    /**
     * Lee el contenido de texto (UTF-8) de un fichero del repositorio, para
     * poder importarlo (p.ej. un .lkml elegido en "Abrir modelo semántico").
     *
     * repoConfig: { type: "github"|"gitlab", url, branch, token }
     * path: ruta completa del fichero dentro del repo
     */
    async getFileContent(repoConfig, path) {
        const type = repoConfig && repoConfig.type;
        if (type === "github") return this._getGitHubFileContent(repoConfig, path);
        if (type === "gitlab") return this._getGitLabFileContent(repoConfig, path);
        throw new Error(`Tipo de repositorio no soportado todavía para leer: "${type || "(ninguno)"}"`);
    },

    async _getGitHubFileContent(repoConfig, path) {
        const ref = this.parseUrl(repoConfig.url);
        if (!ref) throw new Error("URL de repositorio de GitHub no válida.");

        const branch = (repoConfig.branch || "").trim();
        const cleanPath = (path || "").replace(/^\/+|\/+$/g, "");
        if (!cleanPath) throw new Error("Falta la ruta del fichero.");

        let apiUrl = `https://api.github.com/repos/${ref.owner}/${ref.repo}/contents/${encodeURI(cleanPath)}`;
        if (branch) apiUrl += `?ref=${encodeURIComponent(branch)}`;

        const headers = { Accept: "application/vnd.github+json" };
        if (repoConfig.token) headers.Authorization = `Bearer ${repoConfig.token}`;

        const res = await fetch(apiUrl, { headers });
        if (!res.ok) throw new Error(this._friendlyError(res.status, "GitHub"));

        const data = await res.json();

        if (!data || data.type !== "file" || typeof data.content !== "string") {
            throw new Error("La ruta seleccionada no corresponde a un fichero válido.");
        }

        return this._fromBase64Utf8(data.content.replace(/\n/g, ""));
    },

    async _getGitLabFileContent(repoConfig, path) {
        const ref = this.parseUrl(repoConfig.url);
        if (!ref) throw new Error("URL de repositorio de GitLab no válida.");

        const projectId = encodeURIComponent(`${ref.owner}/${ref.repo}`);
        const branch = (repoConfig.branch || "").trim() || "main";
        const cleanPath = (path || "").replace(/^\/+|\/+$/g, "");
        if (!cleanPath) throw new Error("Falta la ruta del fichero.");

        const rawUrl = `https://gitlab.com/api/v4/projects/${projectId}/repository/files/${encodeURIComponent(cleanPath)}/raw?ref=${encodeURIComponent(branch)}`;
        const headers = repoConfig.token ? { "PRIVATE-TOKEN": repoConfig.token } : {};

        const res = await fetch(rawUrl, { headers });
        if (!res.ok) throw new Error(this._friendlyError(res.status, "GitLab"));

        return await res.text();
    },

    /**
     * Crea o actualiza un fichero en el repositorio (commit directo sobre la
     * rama indicada). Si el fichero ya existe se sobrescribe (actualización);
     * si no existe, se crea.
     *
     * repoConfig: { type: "github"|"gitlab", url, branch, token }
     * path: ruta completa del fichero dentro del repo (carpeta + nombre)
     * content: contenido de texto a guardar (UTF-8)
     * commitMessage: mensaje de commit
     */
    async putFile(repoConfig, path, content, commitMessage) {
        const type = repoConfig && repoConfig.type;
        if (type === "github") return this._putGitHub(repoConfig, path, content, commitMessage);
        if (type === "gitlab") return this._putGitLab(repoConfig, path, content, commitMessage);
        throw new Error(`Tipo de repositorio no soportado todavía para guardar: "${type || "(ninguno)"}"`);
    },

    // btoa solo admite Latin1: hay que pasar por encodeURIComponent/escape
    // para poder codificar en base64 un contenido UTF-8 sin perder acentos.
    _toBase64Utf8(str) {
        return btoa(unescape(encodeURIComponent(str)));
    },

    // Inversa de _toBase64Utf8: decodifica un base64 (posiblemente con
    // saltos de línea, como devuelve la API de GitHub) a texto UTF-8.
    _fromBase64Utf8(b64) {
        return decodeURIComponent(escape(atob(b64)));
    },

    async _putGitHub(repoConfig, path, content, commitMessage) {
        const ref = this.parseUrl(repoConfig.url);
        if (!ref) throw new Error("URL de repositorio de GitHub no válida.");

        const branch = (repoConfig.branch || "").trim();
        const cleanPath = (path || "").replace(/^\/+|\/+$/g, "");
        if (!cleanPath) throw new Error("Falta la ruta/nombre de fichero.");

        const headers = { Accept: "application/vnd.github+json" };
        if (repoConfig.token) headers.Authorization = `Bearer ${repoConfig.token}`;

        const apiUrl = `https://api.github.com/repos/${ref.owner}/${ref.repo}/contents/${encodeURI(cleanPath)}`;

        // Si el fichero ya existe hay que mandar su "sha" para actualizarlo
        // en vez de que GitHub rechace la petición por conflicto.
        let sha;
        const getUrl = branch ? `${apiUrl}?ref=${encodeURIComponent(branch)}` : apiUrl;
        const getRes = await fetch(getUrl, { headers });
        if (getRes.ok) {
            const existing = await getRes.json();
            sha = existing.sha;
        } else if (getRes.status !== 404) {
            throw new Error(this._friendlyError(getRes.status, "GitHub"));
        }

        const body = {
            message: commitMessage || `Actualiza ${cleanPath}`,
            content: this._toBase64Utf8(content)
        };
        if (branch) body.branch = branch;
        if (sha) body.sha = sha;

        const putRes = await fetch(apiUrl, {
            method: "PUT",
            headers: { ...headers, "Content-Type": "application/json" },
            body: JSON.stringify(body)
        });

        if (!putRes.ok) {
            const errBody = await putRes.json().catch(() => null);
            throw new Error((errBody && errBody.message) ? `GitHub: ${errBody.message}` : this._friendlyError(putRes.status, "GitHub"));
        }

        return await putRes.json();
    },

    async _putGitLab(repoConfig, path, content, commitMessage) {
        const ref = this.parseUrl(repoConfig.url);
        if (!ref) throw new Error("URL de repositorio de GitLab no válida.");

        const projectId = encodeURIComponent(`${ref.owner}/${ref.repo}`);
        const branch = (repoConfig.branch || "").trim() || "main";
        const cleanPath = (path || "").replace(/^\/+|\/+$/g, "");
        if (!cleanPath) throw new Error("Falta la ruta/nombre de fichero.");

        const fileUrl = `https://gitlab.com/api/v4/projects/${projectId}/repository/files/${encodeURIComponent(cleanPath)}`;
        const authHeaders = repoConfig.token ? { "PRIVATE-TOKEN": repoConfig.token } : {};

        // GitLab distingue crear (POST) de actualizar (PUT): hay que saber
        // primero si el fichero ya existe en esa rama.
        const getRes = await fetch(`${fileUrl}?ref=${encodeURIComponent(branch)}`, { headers: authHeaders });
        if (!getRes.ok && getRes.status !== 404) {
            throw new Error(this._friendlyError(getRes.status, "GitLab"));
        }
        const method = getRes.ok ? "PUT" : "POST";

        const res = await fetch(fileUrl, {
            method,
            headers: { ...authHeaders, "Content-Type": "application/json" },
            body: JSON.stringify({
                branch,
                content,
                encoding: "text",
                commit_message: commitMessage || `Actualiza ${cleanPath}`
            })
        });

        if (!res.ok) {
            const errBody = await res.json().catch(() => null);
            throw new Error((errBody && errBody.message) ? `GitLab: ${errBody.message}` : this._friendlyError(res.status, "GitLab"));
        }

        return await res.json();
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
