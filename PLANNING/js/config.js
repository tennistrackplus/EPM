/**
 * ============================================================
 * DRACO PLANNING — CONFIGURACIÓN
 * ============================================================
 * Rellena aquí tus credenciales antes de desplegar la app.
 */
const DracoConfig = {
    // ---------------------------------------------------------
    // Modo de inicio de sesión "OAuth del cliente" (flujo implícito).
    // Auditoría F-10: en PRODUCCIÓN debe quedar en false, de modo que
    // solo se admita el servicio de autenticación (AUTH-BROKER).
    // Ponerlo a true solo en entornos de prueba.
    // ---------------------------------------------------------
    allowSimpleOAuth: true,
    
    // ---------------------------------------------------------
    // BigQuery (Google OAuth 2.0)
    // ---------------------------------------------------------
    // El OAuth Client ID NO va aquí: lo configura cada cliente en su
    // conexión (o usa su servicio de autenticación). Ver js/google-auth.js.
    // Sin cloud-platform (auditoría F-04): solo BigQuery, Cloud Storage
    // (subida de ficheros) y el email del usuario.
    // Cloud Storage (abrir/guardar en bucket, subida de ficheros). Desactivado:
    // la pantalla de consentimiento de Google solo pide BigQuery y el email.
    // Ponerlo a true si se usan buckets.
    useCloudStorage: false,

    get googleScopes() {
        return [
            "https://www.googleapis.com/auth/bigquery",
            ...(this.useCloudStorage ? ["https://www.googleapis.com/auth/devstorage.read_write"] : []),
            "https://www.googleapis.com/auth/userinfo.email"
        ].join(" ");
    },

    // Nombre del dataset/base de datos de control donde viven las tablas
    // maestras de Draco (PROYECTOS, DIMENSIONES, CUBOS).
    controlDataset: "DRACO_CONTROL",

    // Prefijo aplicado a los datasets/esquemas de cada proyecto Draco y a
    // las tablas físicas de dimensiones/cubos.
    prefix: "DRACO_",

    // ---------------------------------------------------------
    // Ejecución de flujos (flow_run.html) — ver js/storage.js y
    // python/flow_runner.py.
    // ---------------------------------------------------------
    // Endpoint HTTP propio (Cloud Run / Cloud Function) que envuelva flow_runner.py.
    // Recibe POST { flujo_id, variables, run_id } y devuelve el mismo JSON
    // que flow_runner.run_flow(). Ver python/cloud_run_main.py (ya viene
    // listo para desplegar) y el paso a paso en README.md, sección
    // "Interfaces de fichero en BigQuery".
    // Ejemplo tras desplegar en Cloud Run:
    //   flowRunnerHttpEndpoint: "https://draco-flow-runner-xxxxx-ew.a.run.app",
    flowRunnerHttpEndpoint: "https://draco-flow-runner-569601069650.europe-west1.run.app",

    // ---------------------------------------------------------
    // Subida de ficheros a BigQuery (bucket de Google Cloud Storage)
    // ---------------------------------------------------------
    // Nombre del bucket de GCS donde js/storage.js sube el fichero
    // DIRECTAMENTE desde el navegador (ver Storage.uploadToGcsBucket en
    // js/storage.js), reutilizando el mismo token OAuth ya obtenido para
    // BigQuery (scope "devstorage.read_write"): no hace falta URL firmada ni
    // backend intermedio. Solo hay que: 1) crear el bucket, 2) darle al
    // usuario (o grupo) que usa Planning el rol "Storage Object Creator"
    // (o superior) sobre ese bucket, y 3) que python/cloud_run_main.py
    // tenga la variable de entorno UPLOAD_BUCKET con el MISMO nombre, para
    // que el flujo pueda luego leer el fichero desde ahí. Paso a paso
    // completo en README.md, sección "Interfaces de fichero en BigQuery".
    // Déjalo vacío si prefieres tu propio esquema de subida (ver
    // storageUploadUrlBuilder más abajo).
    bigqueryUploadBucket: "draco-landing", // p.ej. "draco-landing-mi-proyecto"

    // Alternativa a bigqueryUploadBucket: función que TÚ implementas para
    // subir el fichero a tu propio backend (URL firmada de GCS, etc.).
    // SOLO se usa cuando bigqueryUploadBucket está vacío. Recibe (storagePath, file) y debe devolver, o bien un string
    // con la URL a la que hacer PUT del contenido del fichero, o bien un
    // objeto { url, headers } si necesitas cabeceras extra (p.ej.
    // Authorization).
    storageUploadUrlBuilder: null, // (storagePath, file) => "https://..." | { url, headers }

    // ---------------------------------------------------------
    // Repositorio GitHub del modelo semántico LookML (.lkml) — BigQuery
    // ---------------------------------------------------------
    // Al guardar un cubo (además
    // del YAML de arriba) se genera su equivalente .lkml (mismo formato
    // que ya exporta el add-in de Excel, ver ADDIN/src/js/lkmlExport.js
    // y js/lkml-export.js) y se hace commit directo en este repositorio,
    // vía la API REST de contenidos de GitHub (ver js/github-repo.js).
    // Ruta dentro del repo: semantic_models/<proyecto>/<cubo>.lkml.
    // Valores POR DEFECTO: lo normal es indicar URL y rama en la conexión.
    // El PAT nunca va aquí: lo introduce el usuario y se guarda solo en la
    // sesión del navegador (ver BQ.getGitRepoConfig() en js/bigquery.js).
    semanticModelGithub: {
        url: "",
        branch: "main"
    }
};

window.DracoConfig = DracoConfig;
