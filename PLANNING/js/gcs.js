/**
 * GCS — acceso de solo lectura a Google Cloud Storage para PLANNING.
 *
 * Mismo enfoque que ADDIN/src/js/gcsExport.js (llamadas REST directas a
 * storage.googleapis.com con el token OAuth ya obtenido por BQ), pero
 * adaptado a este contexto:
 *   - El proyecto de Google Cloud ya se conoce (BQ.getGcpProject()), no
 *     hace falta elegirlo como en el selector del add-in.
 *   - No existe un "bucket configurado" único: el bucket es un parámetro
 *     explícito en cada llamada (se elige en el propio selector de la
 *     plantilla, ver UI.openBucketExcelPickerModal en ui.js).
 *
 * El scope OAuth "cloud-platform" (ya en DracoConfig.googleScopes) incluye
 * acceso de lectura/escritura a Cloud Storage, así que no hace falta pedir
 * un scope adicional ni reconectar.
 */
const GCS = {
    API_BASE: "https://storage.googleapis.com/storage/v1/b",

    _requireToken() {
        const token = BQ.getToken();
        if (!token) {
            const err = new Error("Sesión de Google/BigQuery no válida o expirada. Vuelve a conectarte.");
            err.code = "NO_AUTH";
            throw err;
        }
        return token;
    },

    /** Lista los buckets de Cloud Storage visibles en el proyecto GCP activo. */
    async listBuckets() {
        const token = this._requireToken();
        const projectId = BQ.getGcpProject();
        if (!projectId) throw new Error("No hay un proyecto de Google Cloud activo.");

        const names = [];
        let pageToken = "";
        do {
            let url = `${this.API_BASE}?project=${encodeURIComponent(projectId)}` +
                `&fields=items(name),nextPageToken&maxResults=200`;
            if (pageToken) url += `&pageToken=${encodeURIComponent(pageToken)}`;

            const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
            const data = await response.json().catch(() => ({}));
            if (!response.ok || data.error) {
                throw new Error((data.error && data.error.message) || `Error HTTP ${response.status}`);
            }
            (data.items || []).forEach(b => names.push(b.name));
            pageToken = data.nextPageToken || "";
        } while (pageToken);

        names.sort((a, b) => a.localeCompare(b));
        return names;
    },

    /**
     * Lista los objetos .xlsx/.xlsm de un bucket (opcionalmente bajo un
     * prefijo/carpeta). Devuelve [{ name, size, updated }], más recientes
     * primero.
     */
    async listXlsxObjects(bucket, prefix) {
        if (!bucket) throw new Error("Falta indicar el bucket.");
        const token = this._requireToken();

        let url = `${this.API_BASE}/${encodeURIComponent(bucket)}/o` +
            `?fields=items(name,size,updated,contentType),nextPageToken`;
        if (prefix) url += `&prefix=${encodeURIComponent(prefix)}`;

        const items = [];
        let pageToken = "";
        do {
            const pageUrl = url + (pageToken ? `&pageToken=${encodeURIComponent(pageToken)}` : "");
            const response = await fetch(pageUrl, { headers: { Authorization: `Bearer ${token}` } });
            const data = await response.json().catch(() => ({}));
            if (!response.ok || data.error) {
                throw new Error((data.error && data.error.message) || `Error HTTP ${response.status}`);
            }
            (data.items || []).forEach(o => {
                if (/\.xlsm?$|\.xlsx$/i.test(o.name)) items.push(o);
            });
            pageToken = data.nextPageToken || "";
        } while (pageToken);

        items.sort((a, b) => new Date(b.updated) - new Date(a.updated));
        return items;
    },

    /** Descarga un objeto de un bucket (Blob), usando el token OAuth. */
    async downloadObject(bucket, objectName) {
        const token = this._requireToken();
        const url = `${this.API_BASE}/${encodeURIComponent(bucket)}/o/${encodeURIComponent(objectName)}?alt=media`;
        const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
        if (!response.ok) {
            let msg = `Error HTTP ${response.status}`;
            try {
                const data = await response.json();
                if (data.error && data.error.message) msg = data.error.message;
            } catch (e) { /* respuesta binaria, sin JSON de error */ }
            throw new Error(msg);
        }
        return await response.blob();
    },

    /** Descarga un objeto y dispara la descarga en el navegador. */
    async downloadObjectToBrowser(bucket, objectName) {
        const blob = await this.downloadObject(bucket, objectName);
        const blobUrl = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = blobUrl;
        a.download = objectName.split("/").pop();
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(blobUrl), 4000);
    }
};

window.GCS = GCS;
