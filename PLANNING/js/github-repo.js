/**
 * ============================================================
 * DRACO PLANNING — CLIENTE GITHUB (para el .lkml del modelo semántico)
 * ============================================================
 * Equivalente reducido de ADDIN/src/js/gitRepo.js: aquí solo hace falta
 * "crear o actualizar un fichero" (commit directo sobre una rama) contra
 * la API REST de contenidos de GitHub — no hace falta listar carpetas
 * ni leer ficheros, porque el .lkml siempre se guarda en la misma ruta
 * predecible (ver SemanticModel.generateAndSave en js/semantic-model.js).
 *
 * Configuración en js/config.js -> DracoConfig.semanticModelGithub:
 *   { url: "https://github.com/<owner>/<repo>", branch: "main", token: "<PAT>" }
 */
const GithubRepo = {

    /** Extrae { owner, repo } de una URL https://github.com/<owner>/<repo>(.git) */
    parseUrl(url) {
        if (!url) return null;
        try {
            const clean = url.trim().replace(/\.git$/i, "").replace(/\/+$/, "");
            const u = new URL(clean);
            if (u.hostname !== "github.com") return null;
            const parts = u.pathname.replace(/^\/+/, "").split("/").filter(Boolean);
            if (parts.length < 2) return null;
            return { owner: parts[0], repo: parts[1] };
        } catch (e) {
            return null;
        }
    },

    // btoa solo admite Latin1: hay que pasar por encodeURIComponent/escape
    // para poder codificar en base64 un contenido UTF-8 sin perder acentos.
    _toBase64Utf8(str) {
        return btoa(unescape(encodeURIComponent(str)));
    },

    _friendlyError(status) {
        if (status === 401 || status === 403) {
            return `GitHub: acceso denegado (${status}). Revisa DracoConfig.semanticModelGithub.token en js/config.js.`;
        }
        if (status === 404) {
            return "GitHub: repositorio o rama no encontrados (404). Revisa DracoConfig.semanticModelGithub.url/branch.";
        }
        return `GitHub: no se pudo guardar el fichero (HTTP ${status}).`;
    },

    /**
     * Crea o actualiza `path` dentro del repositorio configurado, con
     * `content` (texto UTF-8) y el mensaje de commit indicado. Si el
     * fichero ya existe se sobrescribe (se manda su "sha" actual).
     *
     * repoConfig: { url, branch, token } (por defecto,
     *              DracoConfig.semanticModelGithub)
     */
    async putFile(path, content, commitMessage, repoConfig) {
        const cfg = repoConfig || (typeof DracoConfig !== "undefined" && DracoConfig.semanticModelGithub) || {};

        const ref = this.parseUrl(cfg.url);
        if (!ref) throw new Error("DracoConfig.semanticModelGithub.url no es una URL de GitHub válida (revisa js/config.js).");
        if (!cfg.token) throw new Error("Falta DracoConfig.semanticModelGithub.token en js/config.js (Personal Access Token con permiso de escritura sobre el repo).");

        const branch = (cfg.branch || "main").trim();
        const cleanPath = (path || "").replace(/^\/+|\/+$/g, "");
        if (!cleanPath) throw new Error("Falta la ruta/nombre de fichero a guardar en GitHub.");

        const headers = { Accept: "application/vnd.github+json", Authorization: `Bearer ${cfg.token}` };
        const apiUrl = `https://api.github.com/repos/${ref.owner}/${ref.repo}/contents/${encodeURI(cleanPath)}`;

        // Si el fichero ya existe hay que mandar su "sha" para actualizarlo
        // en vez de que GitHub rechace la petición por conflicto.
        let sha;
        const getRes = await fetch(`${apiUrl}?ref=${encodeURIComponent(branch)}`, { headers });
        if (getRes.ok) {
            const existing = await getRes.json();
            sha = existing.sha;
        } else if (getRes.status !== 404) {
            throw new Error(this._friendlyError(getRes.status));
        }

        const body = {
            message: commitMessage || `Actualiza ${cleanPath}`,
            content: this._toBase64Utf8(content),
            branch
        };
        if (sha) body.sha = sha;

        const putRes = await fetch(apiUrl, {
            method: "PUT",
            headers: { ...headers, "Content-Type": "application/json" },
            body: JSON.stringify(body)
        });

        if (!putRes.ok) {
            const errBody = await putRes.json().catch(() => null);
            throw new Error((errBody && errBody.message) ? `GitHub: ${errBody.message}` : this._friendlyError(putRes.status));
        }

        return await putRes.json();
    },

    /**
     * Lee el contenido de `path` dentro del repositorio configurado (texto
     * UTF-8 decodificado). Devuelve null si el fichero no existe (404).
     * Se usa para volver a abrir el .lkml que generó SemanticModel al
     * guardar un cubo (ver js/widget-pivot.js).
     */
    async getFile(path, repoConfig) {
        const cfg = repoConfig || (typeof DracoConfig !== "undefined" && DracoConfig.semanticModelGithub) || {};

        const ref = this.parseUrl(cfg.url);
        if (!ref) throw new Error("DracoConfig.semanticModelGithub.url no es una URL de GitHub válida (revisa js/config.js).");
        if (!cfg.token) throw new Error("Falta DracoConfig.semanticModelGithub.token en js/config.js (Personal Access Token con permiso de lectura sobre el repo).");

        const branch = (cfg.branch || "main").trim();
        const cleanPath = (path || "").replace(/^\/+|\/+$/g, "");
        if (!cleanPath) throw new Error("Falta la ruta del fichero a leer en GitHub.");

        const headers = { Accept: "application/vnd.github+json", Authorization: `Bearer ${cfg.token}` };
        const apiUrl = `https://api.github.com/repos/${ref.owner}/${ref.repo}/contents/${encodeURI(cleanPath)}?ref=${encodeURIComponent(branch)}`;

        const res = await fetch(apiUrl, { headers });
        if (res.status === 404) return null;
        if (!res.ok) throw new Error(this._friendlyError(res.status));

        const data = await res.json();
        if (!data || typeof data.content !== "string") throw new Error("GitHub: respuesta sin contenido de fichero.");
        // decodeURIComponent/escape es lo simétrico de _toBase64Utf8 (btoa
        // solo admite Latin1) para volver a obtener el texto UTF-8 original.
        return decodeURIComponent(escape(atob(data.content.replace(/\n/g, ""))));
    }
};
