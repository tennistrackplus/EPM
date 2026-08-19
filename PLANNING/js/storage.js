/**
 * ============================================================
 * DRACO PLANNING — STORAGE (subida de ficheros del navegador)
 * ============================================================
 * Capa de abstracción equivalente a `Provider` pero para el storage de
 * ficheros: el resto de la app (flow-run.js) llama a `Storage.upload`,
 * sin conocer si por debajo hay un stage de Snowflake, un bucket GCS/S3
 * o un backend propio.
 *
 * Por defecto necesita `DracoConfig.storageUploadUrlBuilder`, una función
 * (storagePath, file) => URL a la que hacer PUT con el contenido del
 * fichero (típicamente una URL firmada de tu bucket, o un endpoint propio
 * — igual que el proxy CORS opcional de Snowflake en proxy/cloudflare-worker.js).
 * Si tu backend necesita otra cosa (POST multipart, cabeceras extra...),
 * sustituye el cuerpo de `upload` por tu lógica.
 */
const Storage = {
    /**
     * Sube `file` (objeto File del navegador) a `storagePath` y devuelve
     * la propia ruta de storage (para usarla como valor de la variable
     * 'ruta_storage' del paso de fichero del flujo).
     */
    async upload(file, storagePath) {
        if (typeof DracoConfig.storageUploadUrlBuilder !== "function") {
            throw new Error(
                "Configura DracoConfig.storageUploadUrlBuilder en js/config.js para poder " +
                "subir ficheros al storage (ver comentario junto a esa constante)."
            );
        }
        const url = await DracoConfig.storageUploadUrlBuilder(storagePath, file);
        const resp = await fetch(url, {
            method: "PUT",
            body: file,
            headers: { "Content-Type": file.type || "application/octet-stream" }
        });
        if (!resp.ok) {
            throw new Error(`Error subiendo "${file.name}" al storage (HTTP ${resp.status})`);
        }
        return storagePath;
    },

    /** Construye una ruta de storage por defecto para un fichero de un flujo. */
    buildPath(flujoId, varName, fileName) {
        const safeName = String(fileName || "fichero").replace(/[^\w.\-]+/g, "_");
        return `flows/${flujoId}/${varName}/${Date.now()}_${safeName}`;
    }
};
