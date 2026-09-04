/**
 * ============================================================
 * EPM ADD-IN — EXPORTAR EXCEL A GOOGLE CLOUD STORAGE
 * ============================================================
 * Sube el .xlsx activo (tal cual está guardado en disco/OneDrive) a un
 * bucket de Google Cloud Storage, reutilizando el token OAuth de Google
 * ya guardado por login.js/BQ ("bigquery_access_token"). Para que la
 * subida funcione, ese token debe incluir el scope
 * "https://www.googleapis.com/auth/devstorage.read_write"
 * (añadido en config.js); si el usuario conectó BigQuery ANTES de este
 * cambio, tendrá que reconectar (cerrar sesión y volver a conectar) para
 * obtener un token con permisos de Cloud Storage.
 */
const GCS = {
    UPLOAD_BASE: "https://storage.googleapis.com/upload/storage/v1/b",
    API_BASE: "https://storage.googleapis.com/storage/v1/b",

    /** Comprueba sesión y bucket configurado; devuelve el nombre del bucket o lanza error. */
    _requireBucket() {
        if (!BQ.isConnected()) {
            const err = new Error("No hay una sesión de Google/BigQuery activa. Conéctate primero desde \"Conexiones\".");
            err.code = "NO_AUTH";
            throw err;
        }
        const bucket = BQ.getExportBucket();
        if (!bucket) {
            const err = new Error("No hay ningún bucket de Cloud Storage configurado. Indícalo en la conexión de BigQuery, en el campo \"Bucket de exportación\".");
            err.code = "NO_BUCKET";
            throw err;
        }
        return bucket;
    },

    /**
     * Lista los objetos .xlsx/.xlsm del bucket configurado (opcionalmente bajo
     * un prefijo/carpeta). Devuelve un array de { name, size, updated } ordenado
     * por fecha de modificación descendente (más recientes primero).
     */
    async listXlsxObjects(prefix) {
        const bucket = this._requireBucket();
        const token = BQ.getToken();
        if (!token) {
            const err = new Error("Sesión de Google/BigQuery no válida o expirada. Inicia sesión de nuevo.");
            err.code = "NO_AUTH";
            throw err;
        }

        let url = `${this.API_BASE}/${encodeURIComponent(bucket)}/o` +
            `?fields=items(name,size,updated,contentType),nextPageToken`;
        if (prefix) url += `&prefix=${encodeURIComponent(prefix)}`;

        const items = [];
        let pageToken = "";
        do {
            const pageUrl = url + (pageToken ? `&pageToken=${encodeURIComponent(pageToken)}` : "");
            const response = await fetch(pageUrl, {
                headers: { "Authorization": `Bearer ${token}` }
            });
            const data = await response.json().catch(() => ({}));
            if (!response.ok || data.error) {
                const msg = (data.error && data.error.message) || `Error HTTP ${response.status}`;
                throw new Error(msg);
            }
            (data.items || []).forEach(o => {
                if (/\.xlsm?$|\.xlsx$/i.test(o.name)) items.push(o);
            });
            pageToken = data.nextPageToken || "";
        } while (pageToken);

        items.sort((a, b) => new Date(b.updated) - new Date(a.updated));
        return items;
    },

    /**
     * Descarga un objeto del bucket configurado (Blob) usando el token OAuth
     * (funciona igual con buckets privados o públicos).
     */
    async downloadObject(objectName) {
        const bucket = this._requireBucket();
        const token = BQ.getToken();
        if (!token) {
            const err = new Error("Sesión de Google/BigQuery no válida o expirada. Inicia sesión de nuevo.");
            err.code = "NO_AUTH";
            throw err;
        }

        const url = `${this.API_BASE}/${encodeURIComponent(bucket)}/o/${encodeURIComponent(objectName)}?alt=media`;
        const response = await fetch(url, {
            headers: { "Authorization": `Bearer ${token}` }
        });
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

    /** Descarga un objeto del bucket y dispara la descarga en el navegador
     *  (o el propio Excel/Office la abre, según el sistema del usuario). */
    async downloadObjectToBrowser(objectName) {
        const blob = await this.downloadObject(objectName);
        const blobUrl = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = blobUrl;
        a.download = objectName.split("/").pop();
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(blobUrl), 30000);
    },

    /** Lee el archivo Excel activo (comprimido, tal cual .xlsx) como Uint8Array,
     *  uniendo todas las porciones ("slices") que entrega Office.js. */
    getWorkbookBytes() {
        return new Promise((resolve, reject) => {
            if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
                reject(new Error("Esta función solo está disponible dentro de Excel (Office.js no disponible)."));
                return;
            }

            Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (fileResult) => {
                if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error("No se pudo leer el archivo: " + (fileResult.error && fileResult.error.message)));
                    return;
                }

                const file = fileResult.value;
                const sliceCount = file.sliceCount;
                const slices = new Array(sliceCount);
                let received = 0;
                let failed = false;

                const finish = () => {
                    file.closeAsync();
                    if (failed) return;
                    let total = 0;
                    slices.forEach(s => { total += s.length; });
                    const bytes = new Uint8Array(total);
                    let offset = 0;
                    slices.forEach(s => { bytes.set(s, offset); offset += s.length; });
                    resolve(bytes);
                };

                if (sliceCount === 0) {
                    finish();
                    return;
                }

                for (let i = 0; i < sliceCount; i++) {
                    file.getSliceAsync(i, (sliceResult) => {
                        if (failed) return;
                        if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
                            failed = true;
                            file.closeAsync();
                            reject(new Error("No se pudo leer un fragmento del archivo: " + (sliceResult.error && sliceResult.error.message)));
                            return;
                        }
                        slices[sliceResult.value.index] = new Uint8Array(sliceResult.value.data);
                        received++;
                        if (received === sliceCount) finish();
                    });
                }
            });
        });
    },

    /** Nombre sugerido para el objeto: el del archivo actual si termina en
     *  .xlsx/.xlsm, o "Informe_EPM_<fecha-hora>.xlsx" si no se puede saber
     *  (archivo nunca guardado, entorno sin Office.context.document.url, etc.) */
    getSuggestedFileName() {
        const raw = (typeof Office !== "undefined" && Office.context && Office.context.document && Office.context.document.url) || "";
        let name = "";
        try {
            name = decodeURIComponent(raw.split(/[\\/]/).pop() || "");
        } catch (e) {
            name = raw.split(/[\\/]/).pop() || "";
        }
        if (!name || !/\.xlsm?$|\.xlsx$/i.test(name)) {
            const ts = new Date().toISOString().replace(/[:.]/g, "-");
            name = `Informe_EPM_${ts}.xlsx`;
        }
        return name;
    },

    /** Sube unos bytes concretos a "bucket/objectName" (subida simple, uploadType=media) */
    async uploadBytes(bucket, objectName, bytes, contentType) {
        const token = BQ.getToken();
        if (!token) {
            const err = new Error("Sesión de Google/BigQuery no válida o expirada. Inicia sesión de nuevo.");
            err.code = "NO_AUTH";
            throw err;
        }

        const url = `${this.UPLOAD_BASE}/${encodeURIComponent(bucket)}/o` +
            `?uploadType=media&name=${encodeURIComponent(objectName)}`;

        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": contentType || "application/octet-stream"
            },
            body: bytes
        });

        const data = await response.json().catch(() => ({}));
        if (!response.ok || data.error) {
            const msg = (data.error && data.error.message) || `Error HTTP ${response.status}`;
            const err = new Error(msg);
            err.details = data.error;
            throw err;
        }
        return data;
    },

    /**
     * Flujo completo del botón "Guardar en bucket": valida sesión y bucket
     * configurados, lee el libro activo tal cual está en disco y lo sube.
     * Devuelve la respuesta de la API de Cloud Storage (incluye "bucket",
     * "name", "mediaLink", etc.).
     */
    async saveActiveWorkbookToBucket(objectName) {
        const bucket = this._requireBucket();

        const bytes = await this.getWorkbookBytes();
        const name = objectName || this.getSuggestedFileName();
        const contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        return await this.uploadBytes(bucket, name, bytes, contentType);
    }
};

window.GCS = GCS;
