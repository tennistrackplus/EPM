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
        "https://www.googleapis.com/auth/bigquery",
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
    snowflakeDatabase: "DRACO",

    // ---------------------------------------------------------
    // Microsoft Fabric (Entra ID OAuth 2.0 — Authorization Code + PKCE, SPA)
    // ---------------------------------------------------------
    // Id. de aplicación (cliente) del registro de app en Entra ID.
    // La app debe ser MULTIINQUILINO y tener como plataforma
    // "Aplicación de página única (SPA)" con redirect URI:
    //   https://tennistrackplus.github.io/EPM/ADDIN/src/auth-callback-fabric.html
    fabricClientId: "8603265d-d2de-4ba6-90a3-46136b8abdaa",
    // Tenant por defecto si la conexión no indica uno.
    // "organizations" = cualquier cuenta de trabajo/escuela de cualquier tenant.
    fabricDefaultTenant: "organizations",
    // Token para el SQL endpoint de Fabric (audiencia de Azure SQL) +
    // refresh token (offline_access) + nombre de usuario (openid profile).
    fabricScopes: [
        "https://database.windows.net//user_impersonation",
        "offline_access",
        "openid",
        "profile"
    ].join(" "),
    // Permisos para la API REST de Fabric (listar workspaces y leer la
    // definición de los modelos semánticos al "Abrir modelo semántico").
    // Se piden en el MISMO login que el SQL (consentimiento único); el
    // token para esta API se obtiene después con el refresh token.
    fabricApiScopes: [
        "https://api.fabric.microsoft.com/Workspace.Read.All",
        "https://api.fabric.microsoft.com/SemanticModel.ReadWrite.All"
    ].join(" "),
    // URL del gateway TDS (carpeta gateway/ de este repo), sin barra final.
    fabricGatewayUrl: "TU_FABRIC_GATEWAY_URL"
};
