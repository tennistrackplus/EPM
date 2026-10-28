/**
 * ============================================================
 * EPM ADD-IN — CONFIGURACIÓN DE PROVEEDORES DE DATOS
 * ============================================================
 * Parámetros de conexión a BigQuery.
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
    // Auditoría F-15: celdas de depuración de EDIT_REPORT (SQL, JSON de
    // resultados y trazas de planificación). En false no se guardan en el
    // libro; ponerlo a true solo para diagnosticar un problema.
    // ---------------------------------------------------------
    debugWorkbookCells: false,
    // ---------------------------------------------------------
    // BigQuery (Google OAuth 2.0)
    // ---------------------------------------------------------
    // El OAuth Client ID NO va aquí: lo configura cada cliente en su
    // conexión (o usa su servicio de autenticación). Ver js/google-auth.js.
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
    }
};
