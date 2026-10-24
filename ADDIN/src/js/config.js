/**
 * ============================================================
 * EPM ADD-IN — CONFIGURACIÓN DE PROVEEDORES DE DATOS
 * ============================================================
 * Parámetros de conexión a BigQuery.
 */
const DracoConfig = {
    // ---------------------------------------------------------
    // BigQuery (Google OAuth 2.0)
    // ---------------------------------------------------------
    // El OAuth Client ID NO va aquí: lo configura cada cliente en su
    // conexión (o usa su servicio de autenticación). Ver js/google-auth.js.
    googleScopes: [
        "https://www.googleapis.com/auth/bigquery",
        "https://www.googleapis.com/auth/devstorage.read_write",
        "https://www.googleapis.com/auth/userinfo.email"
    ].join(" ")
};
