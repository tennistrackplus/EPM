-- ============================================================
-- DRACO PLANNING · Subida de ficheros a un stage de Snowflake
-- ============================================================
-- Solo aplica al proveedor Snowflake (para BigQuery se usa
-- DracoConfig.storageUploadUrlBuilder con una URL firmada de GCS/S3,
-- ver js/storage.js).
--
-- Por qué existe esto: el navegador NO puede ejecutar el comando PUT
-- nativo de Snowflake (lo implementan los drivers/SnowSQL, no hay
-- equivalente HTTP público para subir a un stage interno). La app
-- rodea esto así:
--   1) js/storage.js trocea el fichero en Blobs de
--      DracoConfig.snowflakeUploadChunkBytes, los pasa a base64 y los
--      inserta como filas en DRACO_CONTROL.FILE_UPLOAD_CHUNKS via SQL
--      normal (INSERT), que sí es accesible por HTTP (SQL API v2).
--   2) Al terminar, llama a este procedure, que lee todos los trozos
--      de esa subida, los reensambla en memoria y los escribe en el
--      stage con session.file.put_stream() (Snowpark) — esto sí puede
--      escribir en un stage interno desde dentro de Snowflake, sin
--      pasar por un cliente local.
--   3) Borra los trozos ya consumidos del buffer.
--
-- Requisito previo: haber ejecutado sql/00_control_schema.sql (crea
-- DRACO_CONTROL.FILE_UPLOAD_CHUNKS).
--
-- El stage de destino (@DRACO_LANDING en tu caso) ya lo tienes creado.
-- No necesita SERVER-SIDE ENCRYPTION especial para este flujo (eso solo
-- hace falta si además vas a generar URLs prefirmadas de descarga con
-- GET_PRESIGNED_URL). Grants mínimos necesarios para el rol que ejecuta
-- este procedure (normalmente basta con el rol OWNER del propio SP):
--
--   GRANT READ, WRITE ON STAGE DRACO_LANDING TO ROLE <rol_del_procedure>;
--   GRANT USAGE ON DATABASE DRACO TO ROLE <rol_del_procedure>;
--   GRANT USAGE ON SCHEMA DRACO.DRACO_CONTROL TO ROLE <rol_del_procedure>;
--   GRANT SELECT, INSERT, DELETE ON TABLE DRACO.DRACO_CONTROL.FILE_UPLOAD_CHUNKS TO ROLE <rol_del_procedure>;
--
-- (el rol que llama a `CALL` desde el navegador solo necesita poder
-- ejecutar el procedure + INSERT en FILE_UPLOAD_CHUNKS; el procedure
-- corre con los privilegios de su OWNER salvo que lo crees con
-- EXECUTE AS CALLER).

CREATE OR REPLACE PROCEDURE DRACO_CONTROL.SP_FINALIZE_FILE_UPLOAD(
    UPLOAD_ID     STRING,
    TOTAL_CHUNKS  INTEGER,
    STORAGE_PATH  STRING,
    STAGE_NAME    STRING
)
RETURNS STRING
LANGUAGE PYTHON
RUNTIME_VERSION = '3.11'
PACKAGES = ('snowflake-snowpark-python')
HANDLER = 'run'
AS
$$
import base64
import io


def run(session, upload_id, total_chunks, storage_path, stage_name):
    rows = session.sql(
        "SELECT CHUNK_INDEX, CHUNK_B64 FROM DRACO_CONTROL.FILE_UPLOAD_CHUNKS "
        "WHERE UPLOAD_ID = ? ORDER BY CHUNK_INDEX",
        params=[upload_id],
    ).collect()

    if len(rows) != total_chunks:
        raise Exception(
            f"Subida incompleta para UPLOAD_ID={upload_id}: se esperaban "
            f"{total_chunks} trozos y hay {len(rows)} en el buffer. "
            "Reintenta la subida del fichero."
        )

    buffer = io.BytesIO()
    for r in rows:
        buffer.write(base64.b64decode(r["CHUNK_B64"]))
    buffer.seek(0)
    total_bytes = buffer.getbuffer().nbytes

    stage_ref = stage_name if stage_name.startswith("@") else f"@{stage_name}"
    full_stage_path = f"{stage_ref.rstrip('/')}/{storage_path.lstrip('/')}"

    session.file.put_stream(
        buffer,
        full_stage_path,
        auto_compress=False,
        overwrite=True,
    )

    session.sql(
        "DELETE FROM DRACO_CONTROL.FILE_UPLOAD_CHUNKS WHERE UPLOAD_ID = ?",
        params=[upload_id],
    ).collect()

    return f"OK: {total_bytes} bytes escritos en {full_stage_path}"
$$;
