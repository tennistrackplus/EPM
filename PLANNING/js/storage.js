/**
 * ============================================================
 * DRACO PLANNING — STORAGE (subida de ficheros del navegador)
 * ============================================================
 * Capa de abstracción equivalente a `Provider` pero para el storage de
 * ficheros: el resto de la app (flow-run.js) llama a `Storage.upload`,
 * sin conocer si por debajo hay un bucket GCS o un backend propio.
 *
 * - BigQuery: si defines `DracoConfig.bigqueryUploadBucket`, el fichero se
 *   sube DIRECTAMENTE a ese bucket de Google Cloud Storage reutilizando el
 *   token OAuth de la sesión de BigQuery (sin backend intermedio). Si
 *   prefieres tu propio esquema (URL firmada, proxy propio...), usa
 *   `DracoConfig.storageUploadUrlBuilder` en su lugar.
 */
const Storage = {
    /**
     * Sube `file` (objeto File del navegador) a `storagePath` y devuelve
     * la propia ruta de storage (para usarla como valor de la variable
     * 'ruta_storage' del paso de fichero del flujo).
     */
    async upload(file, storagePath, onProgress) {
        return this.uploadViaPut(file, storagePath);
    },

    /** Subida genérica: bucket de GCS si está configurado, si no PUT contra una URL firmada (backend propio). */
    async uploadViaPut(file, storagePath) {
        if (DracoConfig.bigqueryUploadBucket) {
            return this.uploadToGcsBucket(file, storagePath);
        }

        if (typeof DracoConfig.storageUploadUrlBuilder !== "function") {
            throw new Error(
                "Configura DracoConfig.bigqueryUploadBucket o DracoConfig.storageUploadUrlBuilder " +
                "en js/config.js para poder subir ficheros al storage (ver comentarios junto a esas constantes)."
            );
        }
        const built = await DracoConfig.storageUploadUrlBuilder(storagePath, file);
        const url = typeof built === "string" ? built : built.url;
        const extraHeaders = (built && typeof built === "object" && built.headers) || {};
        const resp = await fetch(url, {
            method: "PUT",
            body: file,
            headers: { "Content-Type": file.type || "application/octet-stream", ...extraHeaders }
        });
        if (!resp.ok) {
            throw new Error(`Error subiendo "${file.name}" al storage (HTTP ${resp.status})`);
        }
        return storagePath;
    },

    /**
     * Subida directa a un bucket de Google Cloud Storage (DracoConfig.
     * bigqueryUploadBucket), reutilizando el token OAuth de la sesión de
     * BigQuery activa (BQ.getToken(), scope "devstorage.read_write" - ver
     * js/config.js). No hace falta URL firmada ni backend intermedio: el
     * bucket solo necesita el permiso "Storage Object Creator" para el
     * usuario/grupo que usa Planning (ver README.md, sección "Interfaces
     * de fichero en BigQuery").
     */
    async uploadToGcsBucket(file, storagePath) {
        const token = BQ.getToken();
        if (!token) {
            throw new Error("Sesión de BigQuery no válida o expirada: vuelve a iniciar sesión antes de subir el fichero.");
        }
        const bucket = DracoConfig.bigqueryUploadBucket;
        const objectName = String(storagePath).replace(/^\/+/, "");
        const url = `https://storage.googleapis.com/upload/storage/v1/b/${encodeURIComponent(bucket)}/o` +
            `?uploadType=media&name=${encodeURIComponent(objectName)}`;
        const resp = await fetch(url, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": file.type || "application/octet-stream"
            },
            body: file
        });
        if (!resp.ok) {
            const detail = await resp.text().catch(() => "");
            throw new Error(`Error subiendo "${file.name}" al bucket "${bucket}" (HTTP ${resp.status}). ${detail}`);
        }
        return storagePath;
    },

    /**
     * Construye una ruta de storage por defecto para un fichero de un flujo.
     * Nombre plano, sin subcarpetas, para que el flujo lo localice sin
     * depender de la estructura de carpetas del bucket.
     */
    buildPath(flujoId, varName, fileName) {
        const safeName = String(fileName || "fichero").replace(/[^\w.\-]+/g, "_");
        const safeFlujo = String(flujoId || "flujo").replace(/[^\w.\-]+/g, "_");
        const safeVar = String(varName || "var").replace(/[^\w.\-]+/g, "_");
        return `${safeFlujo}__${safeVar}__${Date.now()}_${safeName}`;
    }
};
