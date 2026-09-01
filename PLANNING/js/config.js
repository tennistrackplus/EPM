/**
 * ============================================================
 * DRACO PLANNING — CONFIGURACIÓN
 * ============================================================
 * Rellena aquí tus credenciales antes de desplegar la app.
 */
const DracoConfig = {
    // ---------------------------------------------------------
    // BigQuery (Google OAuth 2.0 — flujo implícito)
    // ---------------------------------------------------------
    // Mismo Client ID de Google que ya usa ADDIN/src/js/config.js: es un
    // único OAuth Client ID de tipo "Aplicación web" que puede tener
    // VARIAS "URI de redireccionamiento autorizadas" a la vez, así que no
    // hace falta crear uno nuevo para Planning — solo añadir su
    // auth-callback.html (tu-dominio-planning/auth-callback.html) a la
    // lista de redirect URIs de este Client ID en Google Cloud Console
    // (Credenciales → este Client ID → "URIs de redireccionamiento
    // autorizados" → Añadir URI). Si el add-in y Planning van en el mismo
    // dominio (recomendado, ver sección 1bis del README), puede que la
    // URI ya esté cubierta si usan la misma ruta base.
    googleClientId: "316357511817-lck6pdotv8mrb7n72pahuukt2e0fvsrt.apps.googleusercontent.com",
    googleScopes: [
        "https://www.googleapis.com/auth/bigquery",
        "https://www.googleapis.com/auth/cloud-platform",
        "https://www.googleapis.com/auth/userinfo.email"
    ].join(" "),

    // ---------------------------------------------------------
    // Snowflake (OAuth 2.0 — Authorization Code + PKCE, cliente público)
    // ---------------------------------------------------------
    // Client ID que Snowflake genera al crear la SECURITY INTEGRATION
    // (ver sql/01_snowflake_oauth_integration.sql). No hace falta client
    // secret: al ser un cliente público, PKCE sustituye al secreto.
    snowflakeClientId: "m6fB0orhRaDGx6UBtpOKQWxXlKQ=",
    snowflakeScopes: "refresh_token",

    // Nombre del dataset/base de datos de control donde viven las tablas
    // maestras de Draco (PROYECTOS, DIMENSIONES, CUBOS).
    controlDataset: "DRACO_CONTROL",

    // Base de datos Snowflake donde vivirán todos los esquemas de Draco
    // (DRACO_CONTROL y DRACO_<PROYECTO> son esquemas dentro de ella).
    snowflakeDatabase: "DRACO",

    // Prefijo aplicado a los datasets/esquemas de cada proyecto Draco y a
    // las tablas físicas de dimensiones/cubos.
    prefix: "DRACO_",

    // ---------------------------------------------------------
    // Ejecución de flujos (flow_run.html) — ver js/storage.js y
    // python/flow_runner.py.
    // ---------------------------------------------------------
    // Procedimiento almacenado (Snowflake) que lanza flow_runner.main().
    // Se invoca como: CALL <flowRunnerProcedure>(flujo_id, variables_json, run_id)
    flowRunnerProcedure: "DRACO_CONTROL.SP_RUN_FLUJO",

    // Alternativa para BigQuery (sin stored procs Python nativos): endpoint
    // HTTP propio (Cloud Run / Cloud Function) que envuelva flow_runner.py.
    // Recibe POST { flujo_id, variables, run_id } y devuelve el mismo JSON
    // que flow_runner.run_flow(). Ver python/cloud_run_main.py (ya viene
    // listo para desplegar) y el paso a paso en README.md, sección
    // "Interfaces de fichero en BigQuery". Déjalo vacío si usas Snowflake.
    // Ejemplo tras desplegar en Cloud Run:
    //   flowRunnerHttpEndpoint: "https://draco-flow-runner-xxxxx-ew.a.run.app",
    flowRunnerHttpEndpoint: "https://draco-flow-runner-569601069650.europe-west1.run.app",

    // ---------------------------------------------------------
    // Subida de ficheros a BigQuery (bucket de Google Cloud Storage)
    // ---------------------------------------------------------
    // Nombre del bucket de GCS donde js/storage.js sube el fichero
    // DIRECTAMENTE desde el navegador (ver Storage.uploadToGcsBucket en
    // js/storage.js), reutilizando el mismo token OAuth ya obtenido para
    // BigQuery (scope "cloud-platform"): no hace falta URL firmada ni
    // backend intermedio. Solo hay que: 1) crear el bucket, 2) darle al
    // usuario (o grupo) que usa Planning el rol "Storage Object Creator"
    // (o superior) sobre ese bucket, y 3) que python/cloud_run_main.py
    // tenga la variable de entorno UPLOAD_BUCKET con el MISMO nombre, para
    // que el flujo pueda luego leer el fichero desde ahí. Paso a paso
    // completo en README.md, sección "Interfaces de fichero en BigQuery".
    // Déjalo vacío si usas Snowflake, o si prefieres tu propio esquema de
    // subida (ver storageUploadUrlBuilder más abajo).
    bigqueryUploadBucket: "draco-landing", // p.ej. "draco-landing-mi-proyecto"

    // Alternativa a bigqueryUploadBucket: función que TÚ implementas para
    // subir el fichero a tu propio backend (URL firmada de GCS/S3, un
    // proxy propio como proxy/cloudflare-worker.js, etc.). SOLO se usa
    // cuando Provider.key() !== "snowflake" y bigqueryUploadBucket está
    // vacío. Recibe (storagePath, file) y debe devolver, o bien un string
    // con la URL a la que hacer PUT del contenido del fichero, o bien un
    // objeto { url, headers } si necesitas cabeceras extra (p.ej.
    // Authorization). Para Snowflake NO hace falta configurar nada de
    // esto: ver el bloque de abajo.
    storageUploadUrlBuilder: null, // (storagePath, file) => "https://..." | { url, headers }

    // ---------------------------------------------------------
    // Subida de ficheros a Snowflake (js/storage.js + SP_FINALIZE_FILE_UPLOAD)
    // ---------------------------------------------------------
    // Snowflake no permite un PUT HTTP directo a un stage desde el navegador,
    // así que el fichero se trocea en base64 y se reensambla en un stored
    // procedure (ver sql/02_snowflake_file_upload.sql). Debe coincidir con
    // el `stage_name` que uses al invocar python/flow_runner.py::main() para
    // que el flujo pueda encontrar luego el fichero.
    snowflakeUploadStage: "@DRACO_LANDING",
    snowflakeFinalizeUploadProcedure: "DRACO_CONTROL.SP_FINALIZE_FILE_UPLOAD",
    // Tamaño de cada trozo (antes de base64) enviado por INSERT. 4MB es un
    // valor conservador para no acercarse al límite de tamaño de sentencia
    // SQL de la SQL API v2; auméntalo si tus ficheros son grandes y quieres
    // menos peticiones, o bájalo si ves errores de "statement too large".
    snowflakeUploadChunkBytes: 4 * 1024 * 1024,

    // ---------------------------------------------------------
    // Repositorio GitHub del modelo semántico LookML (.lkml) — BigQuery
    // ---------------------------------------------------------
    // Cuando el proveedor activo es BigQuery, al guardar un cubo (además
    // del YAML de arriba) se genera su equivalente .lkml (mismo formato
    // que ya exporta el add-in de Excel, ver ADDIN/src/js/lkmlExport.js
    // y js/lkml-export.js) y se hace commit directo en este repositorio,
    // vía la API REST de contenidos de GitHub (ver js/github-repo.js).
    // Ruta dentro del repo: semantic_models/<proyecto>/<cubo>.lkml.
    semanticModelGithub: {
        url: "https://github.com/tennistrackplus/SEMANTIC_MODEL",
        branch: "main",
        // El PAT ya NO se guarda aquí (GitHub lo bloqueaba al detectarlo
        // en el código). Se lee en el momento de hacer el commit desde el
        // secreto "github-pat-draco" de Google Secret Manager — ver
        // BQ.getGithubPatFromSecretManager() en js/bigquery.js y su uso en
        // SemanticModel.generateAndPushLkml() (js/semantic-model.js).
        token: ""
    }
};
