/**
 * ============================================================
 * DRACO PLANNING — STORAGE (subida de ficheros del navegador)
 * ============================================================
 * Capa de abstracción equivalente a `Provider` pero para el storage de
 * ficheros: el resto de la app (flow-run.js) llama a `Storage.upload`,
 * sin conocer si por debajo hay un stage de Snowflake, un bucket GCS/S3
 * o un backend propio.
 *
 * - Snowflake: no hay forma de hacer un PUT HTTP directo a un stage
 *   interno desde el navegador (el comando PUT lo implementan los
 *   drivers/SnowSQL, no la SQL API). En su lugar, el fichero se trocea,
 *   se manda por SQL normal (INSERT ... vía Provider.runQuery) a una
 *   tabla buffer, y un stored procedure (SP_FINALIZE_FILE_UPLOAD, ver
 *   sql/02_snowflake_file_upload.sql) lo reensambla y lo escribe en el
 *   stage con session.file.put_stream().
 * - BigQuery / otros: usa `DracoConfig.storageUploadUrlBuilder`, una
 *   función (storagePath, file) => URL a la que hacer PUT con el
 *   contenido del fichero (típicamente una URL firmada de tu bucket).
 */
const Storage = {
    /**
     * Sube `file` (objeto File del navegador) a `storagePath` y devuelve
     * la propia ruta de storage (para usarla como valor de la variable
     * 'ruta_storage' del paso de fichero del flujo).
     */
    async upload(file, storagePath, onProgress) {
        if (Provider.key() === "snowflake") {
            return this.uploadToSnowflakeStage(file, storagePath, onProgress);
        }
        return this.uploadViaPut(file, storagePath);
    },

    /** Subida genérica: PUT contra una URL firmada (GCS/S3/backend propio). */
    async uploadViaPut(file, storagePath) {
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

    /**
     * Subida a un stage de Snowflake: trocea el fichero, sube cada trozo en
     * base64 por SQL (INSERT) y pide al procedure que lo reensamble en el
     * stage. `onProgress(chunkIndex, totalChunks)` es opcional, para pintar
     * un hint de progreso en flow-run.js.
     */
    async uploadToSnowflakeStage(file, storagePath, onProgress) {
        const chunkBytes = DracoConfig.snowflakeUploadChunkBytes || 4 * 1024 * 1024;
        const totalChunks = Math.max(1, Math.ceil(file.size / chunkBytes));
        const uploadId = Provider.newId();

        for (let i = 0; i < totalChunks; i++) {
            const start = i * chunkBytes;
            const end = Math.min(start + chunkBytes, file.size);
            const chunkB64 = await this.blobToBase64(file.slice(start, end));
            await Provider.runQuery(`
                INSERT INTO ${Provider.qualifyControl("FILE_UPLOAD_CHUNKS")}
                    (UPLOAD_ID, CHUNK_INDEX, CHUNK_B64, FECHA_CARGA)
                VALUES ('${Provider.esc(uploadId)}', ${i}, '${Provider.esc(chunkB64)}', CURRENT_TIMESTAMP())`);
            if (onProgress) onProgress(i + 1, totalChunks);
        }

        const stage = DracoConfig.snowflakeUploadStage || "@DRACO_LANDING";
        const proc = DracoConfig.snowflakeFinalizeUploadProcedure || "DRACO_CONTROL.SP_FINALIZE_FILE_UPLOAD";
        const rows = await Provider.runQuery(
            `CALL ${proc}('${Provider.esc(uploadId)}', ${totalChunks}, '${Provider.esc(storagePath)}', '${Provider.esc(stage)}')`
        );
        const resultCol = rows[0] && Object.values(rows[0])[0];
        if (resultCol && String(resultCol).startsWith("ERROR")) {
            throw new Error(String(resultCol));
        }
        return storagePath;
    },

    /** Lee un Blob y devuelve su contenido en base64 (sin el prefijo data:...;base64,). */
    blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => {
                const result = String(reader.result || "");
                const comma = result.indexOf(",");
                resolve(comma >= 0 ? result.slice(comma + 1) : result);
            };
            reader.onerror = () => reject(reader.error || new Error("Error leyendo el fichero en el navegador."));
            reader.readAsDataURL(blob);
        });
    },

    /**
     * Construye una ruta de storage por defecto para un fichero de un flujo.
     * Sin subcarpetas a propósito: Snowflake, al descargar (GET) un fichero
     * que vive en una subcarpeta del stage, intenta recrear esa misma
     * subcarpeta en el disco local del stored procedure, y ahí solo /tmp es
     * escribible — con nombre plano ese problema no existe.
     */
    buildPath(flujoId, varName, fileName) {
        const safeName = String(fileName || "fichero").replace(/[^\w.\-]+/g, "_");
        const safeFlujo = String(flujoId || "flujo").replace(/[^\w.\-]+/g, "_");
        const safeVar = String(varName || "var").replace(/[^\w.\-]+/g, "_");
        return `${safeFlujo}__${safeVar}__${Date.now()}_${safeName}`;
    }
};
