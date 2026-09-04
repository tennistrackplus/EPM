/**
 * ============================================================
 * EPM ADD-IN — CONFIGURACIÓN DE PROVEEDORES DE DATOS
 * ============================================================
 * Client IDs y parámetros de conexión de cada motor. El Client ID
 * de Google ya estaba embebido en login.js; se centraliza aquí
 * junto con la configuración de Snowflake para que ambos motores
 * se configuren en un único sitio.
 */
const DracoConfig = {
    // ---------------------------------------------------------
    // BigQuery (Google OAuth 2.0 — flujo implícito)
    // ---------------------------------------------------------
    googleClientId: "316357511817-lck6pdotv8mrb7n72pahuukt2e0fvsrt.apps.googleusercontent.com",
    googleScopes: [
        "https://www.googleapis.com/auth/bigquery.readonly",
        "https://www.googleapis.com/auth/devstorage.read_write",
        "https://www.googleapis.com/auth/userinfo.email"
    ].join(" "),

    // ---------------------------------------------------------
    // Snowflake (OAuth 2.0 — Authorization Code + PKCE, cliente público)
    // ---------------------------------------------------------
    // ⚠️ Client ID de la SECURITY INTEGRATION de Snowflake usada por
    // Draco Planning. Si el add-in se sirve desde un dominio/origen
    // distinto al de Planning, añade también su auth-callback-snowflake.html
    // a OAUTH_REDIRECT_URI en esa integración (puede llevar varias URIs),
    // o crea una integración propia para el add-in.
    snowflakeClientId: "m6fB0orhRaDGx6UBtpOKQWxXlKQ=",
    snowflakeScopes: "refresh_token",

    // Base de datos Snowflake por defecto si el usuario no indica otra
    // al conectar (puede sobreescribirse en el panel de login).
    snowflakeDatabase: "DRACO"
};
