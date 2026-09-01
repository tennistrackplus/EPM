# Draco Planning

App web funcional (HTML/CSS/JS puro, sin build ni frameworks) para modelar
proyectos, dimensiones y cubos sobre BigQuery **o Snowflake**, con el mismo
lenguaje visual que vuestra app anterior (EPM Data Studio).

## 1. Qué hay ya funcional

- **Login / landing (`index.html`)**: selector de conectores (BigQuery,
  Snowflake, Microsoft Fabric, Amazon, SAP Datasphere). **BigQuery y
  Snowflake están realmente conectados**; el resto son placeholders con
  "estará disponible próximamente".
- **Capa de proveedor (`js/provider.js`)**: todo el resto de la app
  (proyectos, dimensiones, cubos, progreso) habla con `Provider`, no
  directamente con BigQuery o Snowflake — por eso las mismas pantallas
  funcionan igual sobre cualquiera de los dos motores.
- **Conexión real a BigQuery**: OAuth 2.0 implícito (popup), token en
  `localStorage`, llamadas REST directas (`js/bigquery.js`).
- **Conexión real a Snowflake**: OAuth 2.0 *Authorization Code + PKCE*
  (cliente público, sin client secret) y SQL API v2 (`js/snowflake.js`).
  Requiere que un ACCOUNTADMIN ejecute una vez
  `sql/01_snowflake_oauth_integration.sql` para dar de alta el cliente
  OAuth de Draco Planning en tu cuenta de Snowflake.
- **Bootstrap del esquema de control**: al conectar y elegir tu "hogar"
  de trabajo (proyecto GCP, o cuenta+warehouse de Snowflake):
  - **Snowflake**: se crea automáticamente y en silencio si no existe
    (comportamiento sin cambios).
  - **BigQuery (nuevo)**: en vez de crearlo en silencio, la app
    comprueba primero si ya existe la primera tabla
    (`DRACO_CONTROL.PROYECTOS`, `DracoSchema.controlSchemaExists()`
    en `js/schema.js`). Si no existe, muestra un aviso con un botón
    **"Instalar esquema de control"** (y, si la conexión todavía no
    trae un proyecto GCP elegido, primero pide seleccionarlo). Solo al
    pulsar el botón se crean el dataset y las tablas — ver
    `js/auth.js::proceedBigQuery()` / `installControlSchema()`.
  El script SQL equivalente está en `sql/00_control_schema.sql`.
- **Panel principal (`app.html`)**: barra de proyecto (crear/eliminar),
  bloques **Administración** y **Mi Progreso** (maximizables/contraíbles),
  y dentro de Administración:
  - **Dimensiones**: la clave principal de la tabla es siempre el propio
    nombre de la dimensión (ej. dimensión `CUENTA` → columna `CUENTA`,
    clave). El resto de columnas son "Atributos", y cualquiera de ellos
    también se puede marcar como clave para soportar claves compuestas
    (ej. Clase de coste + Sociedad). Además, un atributo (como mucho uno,
    es opcional) se puede marcar con el icono 🏷 **"Desc."** para indicar
    que es el **atributo descriptivo** de la dimensión (ej. `CUENTA` =
    "4300001", atributo descriptivo = "Proveedores nacionales"). Es un
    estado tipo radio: marcar uno desmarca el anterior, y se puede
    desmarcar volviendo a pulsarlo (no es obligatorio tener uno). Se
    guarda como `isDescription: true` dentro de `CAMPOS_JSON` y viaja
    después al modelo semántico (ver más abajo).
    - **Valores** (icono ▤): ahora se abre en un **popup casi a pantalla
      completa**. Rejilla editable de los datos físicos — añadir filas,
      pegar bloques copiados de Excel, exportar a CSV/Excel, importar
      desde archivo (sustituyendo todo o de forma incremental por clave).
      "Guardar" vuelca la rejilla completa a la tabla. **Valida que la
      clave sea única antes de guardar** (ni BigQuery ni Snowflake la
      fuerzan a nivel de motor): si hay valores de clave repetidos, las
      filas se marcan en rojo y el guardado se bloquea hasta corregirlas.
    - **Jerarquías** (icono ⛓): arrastra atributos a una lista de niveles
      (superior → inferior) con vista previa en vivo sobre los datos
      reales de la dimensión.
  - **Cubos**: para añadir dimensiones se pulsa "+ Añadir dimensión", lo
    que abre un **popup con buscador** y una tabla (Dimensión /
    Descripción) — clic en una fila la añade y cierra el popup. Pensado
    para proyectos con muchas dimensiones (probado con 200 simuladas: se
    filtra al instante). Cada dimensión añadida se puede quitar con la ✕.
    Las medidas se definen libremente (nombre + tipo), igual que antes.
  - **Workflows** (nuevo, solo definición — la ejecución se aborda más
    adelante): un workflow es una secuencia de **pasos**, arrastrable para
    reordenar, que empieza siempre por un **Paso 0** fijo (no se puede
    eliminar) donde solo se definen las **variables del workflow** —de
    valor único, se piden al crear cada ejecución y quedan disponibles
    para completar tareas en cualquier paso. El resto de pasos tienen 2
    pestañas:
    - **Propiedades**: nombre; Inicio a la izquierda / Finalización a la
      derecha (tarjetas con conector visual entre ambas) y, debajo, el
      Driver — todo en una sola pantalla.
      - Inicio: al iniciar el workflow / al completar el paso anterior /
        fecha concreta.
      - Revisión (sí/no).
      - Finalización: N/A, al enviarse a revisión o al completarse
        —según el valor de Revisión— o fecha concreta. Es lo que dispara
        el "al completar el paso anterior" del siguiente paso.
      - Driver: dimensión opcional por la que se reparte la ejecución
        del paso (ej. CECOs) + valores concretos opcionales (buscador
        sobre los datos reales de la dimensión); sin selección se
        reparte entre todos los valores. Sin dimensión, el paso se
        asigna en bloque.
    - **Tareas del paso**: los bloques son un menú lateral (igual que
      Dimensiones/Cubos/Interfaces) — "+ Añadir bloque" crea uno nuevo y
      se reordenan arrastrando; a la derecha, las tareas del bloque
      seleccionado. Cada tarea es de tipo *flujo manual* (con buscador
      sobre los Flujos de carga manuales del proyecto, cargando
      automáticamente sus variables de pantalla), *actualización de
      tabla de parametrización* (buscador de Dimensiones), o *plantilla*
      / *función* / *página HTML* (estos tres últimos, sin catálogo
      propio todavía, se referencian por nombre libre). Cada variable
      que expone la tarea se completa por constante o por una variable
      del workflow (Paso 0), y se puede **ocultar** de la pantalla de
      ejecución.
  - **Ejecuciones de un Workflow** (nuevo, `js/workflow-runs.js`): desde el
    listado de Workflows, el botón ▶ abre un **popup** (90% de pantalla,
    igual que el editor de un workflow) con el listado de ejecuciones de
    ese workflow; "← Volver a ejecuciones" navega entre listado y
    detalle dentro del mismo popup, y la X lo cierra. "Nueva ejecución"
    abre un segundo popup, de tamaño normal, pidiendo el nombre y el
    valor de cada variable del Paso 0; al confirmar, instancia todos los
    pasos ejecutables: si un paso tiene driver, una instancia por cada
    valor (los elegidos en la definición, o todos los reales de la
    dimensión si no se eligió ninguno); si no, una única instancia. La
    asignación se hace **paso a paso**: arriba se repite la cadena de
    pasos a modo de pestañas, cada una con un badge de color (verde ✓ /
    rojo ⚠) según si todas sus instancias están asignadas, y debajo solo
    se muestran —en una rejilla de tarjetas— las instancias del paso
    seleccionado, no las de todos los pasos apiladas. Cada instancia se
    puede asignar (texto libre, todavía no hay módulo de Roles/usuarios),
    completar sus variables, y mover de estado: Pendiente → En curso →
    (si el paso requiere revisión) En revisión → Completado. Al
    completarse todas las instancias de un paso se desbloquean
    automáticamente las del siguiente si su inicio es "al completar el
    paso anterior". Las tareas se muestran a título informativo; las de
    tipo *flujo manual*
    enlazan a `flow_run.html` para ejecutarse de verdad. Los pasos con
    inicio "fecha concreta" no tienen scheduler en servidor: quedan
    disponibles igualmente y solo se informa de la fecha prevista.
  - El resto de módulos del menú (Cargas de datos, Funciones, Flujos de
    proceso, Roles) son pantallas "próximamente".
  - **Modelo semántico (YAML) de un Cubo** (nuevo, `js/semantic-model.js`):
    cada vez que se guarda un cubo (crear o editar) Draco genera
    automáticamente un YAML con la tabla de hechos, sus dimensiones
    (con atributos y jerarquías) y lo sube al mismo sitio donde ya se
    guardan otros ficheros del proyecto — el stage de Snowflake
    (`@DRACO_LANDING` por defecto), reutilizando el mecanismo de
    `js/storage.js` (trocear + `INSERT` + `SP_FINALIZE_FILE_UPLOAD`,
    ver sección 7bis y `sql/02_snowflake_file_upload.sql`). No requiere
    infraestructura nueva. Si el proveedor activo no es Snowflake, o la
    subida falla, el modelo se genera igualmente y se ofrece como
    descarga en el navegador — nunca bloquea el guardado del cubo. La
    ruta resultante y la fecha de generación quedan en
    `CUBOS.MODELO_YAML_PATH` / `MODELO_YAML_FECHA`. Ver la sección
    "Modelo semántico (YAML)" más abajo para la estructura completa,
    pensada para que el add-in de Excel (`addin/src`) la lea y rellene
    sus pestañas `model_dimension` / `model_Fact` / `model_hier` — **de
    momento este cambio solo genera y guarda el YAML; el add-in no se
    ha tocado**.

El grid de valores usa **SheetJS** (cargado por CDN en `app.html`) para
leer/escribir CSV y Excel sin necesitar backend.

⚠️ **Importante**: al editar una dimensión o cubo, la tabla física se recrea
con `CREATE OR REPLACE TABLE`, lo que **borra los datos existentes** si
cambias las columnas. Es el comportamiento esperado en fase de modelado.

## 1bis. Conexiones compartidas con el add-in de Excel (nuevo)

`js/connections.js` es una copia literal de `ADDIN/src/js/connections.js`
(mismas claves de `localStorage`: `epm_connections` / `epm_active_connection_id`,
mismo formato de objeto `{id, name, provider, config, ...}`). La landing
(`index.html` + `js/auth.js`) ahora:

- Al arrancar, si hay conexiones guardadas las lista arriba (mismo
  componente visual `.conn-card` que el add-in); un clic reutiliza sus
  parámetros y conecta sin volver a preguntarlos (si el token OAuth
  sigue vivo, ni siquiera reabre el popup de login). "+ Nueva conexión"
  despliega el selector de conectores de siempre.
- Cada vez que conectas (o cambias el proyecto GCP / los campos de
  Snowflake) se crea o actualiza una entrada en `Connections` con
  `Connections.create()` / `Connections.update()`.
- Para BigQuery se guardan dos campos de proyecto en `config`:
  `homeProjectId` (el que usa Planning para `DRACO_CONTROL` y
  `DRACO_<proyecto>`) y `billingProjectId` (el que ya usaba el add-in
  para pagar las consultas). Por defecto se rellenan los dos con el
  mismo valor, así una conexión creada en cualquiera de las dos apps
  sirve tal cual en la otra.

⚠️ **Importante — mismo origen**: esto solo funciona "de verdad" (lo
creado en un sitio aparece en el otro sin tocar nada más) si Planning y
el add-in se sirven desde el **mismo origen** (mismo esquema+dominio+
puerto, aunque sea en rutas distintas: `https://tuapp.com/planning/` y
`https://tuapp.com/addin/`), porque `localStorage` es por origen. Si
alguna vez se despliegan en dominios distintos, dejan de compartir el
almacén automáticamente y haría falta un backend común (o guardar las
conexiones como filas en `DRACO_CONTROL` en vez de en `localStorage`).

## 2. Puesta en marcha — BigQuery

1. Habilita la **BigQuery API** en tu proyecto de GCP.
2. Crea credenciales **OAuth 2.0 Client ID** (tipo *Aplicación web*):
   - **Orígenes de JavaScript autorizados**: el dominio donde publiques esta app.
   - **URI de redireccionamiento autorizados**: esa URL + `/auth-callback.html`.
3. Copia el Client ID en `js/config.js` (`googleClientId`).

## 3. Puesta en marcha — Snowflake

1. Como `ACCOUNTADMIN` (o rol con `CREATE INTEGRATION`), ejecuta
   `sql/01_snowflake_oauth_integration.sql` sustituyendo `{REDIRECT_URI}`
   por tu dominio + `/auth-callback-snowflake.html`. Esto crea un cliente
   OAuth **público** (protegido con PKCE, sin client secret — apto para
   una app de navegador).
2. Copia el `OAUTH_CLIENT_ID` que te devuelve `DESCRIBE SECURITY
   INTEGRATION DRACO_PLANNING_OAUTH;` en `js/config.js`
   (`snowflakeClientId`).
3. Al conectar desde la app, te pedirá: identificador de cuenta (ej.
   `xy12345.eu-west-1`), warehouse, base de datos (por defecto `DRACO`) y,
   opcionalmente, un rol.

### ⚠️ Snowflake · si te bloquea CORS

El dominio `*.snowflakecomputing.com` no siempre envía cabeceras
`Access-Control-Allow-Origin` para peticiones hechas directamente desde el
navegador (a diferencia de `googleapis.com`, pensado para clientes JS). Si
al conectar ves en la consola un error de tipo **"CORS" / "Failed to
fetch"**, necesitas un proxy ligero que reenvíe las peticiones añadiendo
esas cabeceras:

1. Despliega `proxy/cloudflare-worker.js` (instrucciones dentro del
   archivo — es gratis y tarda unos minutos).
2. En `js/snowflake.js`, cambia `apiBase: ""` por la URL de tu proxy.

No hace falta tocar nada más: el resto del código ya usa `this.base()` en
vez de la URL de Snowflake directamente, así que el cambio es de una sola
línea.

## 4. Puesta en marcha — común a ambos

Sirve la carpeta con cualquier servidor estático (no vale `file://`, hace
falta `http(s)://` para los popups de login y `postMessage`). Por ejemplo:
`npx serve .` o `python3 -m http.server`. Abre `index.html`, elige tu
conector, conecta y entra.

## 5. Modelo de datos de control

Dataset/esquema `DRACO_CONTROL` (uno por proyecto GCP o por base de datos
Snowflake):

| Tabla         | Columnas clave                                                                 |
|---------------|---------------------------------------------------------------------------------|
| `PROYECTOS`   | PROYECTO_ID, PROYECTO, DESCRIPCION, DATASET, USUARIO, FECHA_CREACION            |
| `DIMENSIONES` | DIMENSION_ID, PROYECTO_ID, DIMENSION, DESCRIPCION, TABLA, CAMPOS_JSON, ...      |
| `CUBOS`       | CUBO_ID, PROYECTO_ID, CUBOS, DESCRIPCION, TABLA, CAMPOS_JSON, ...               |
| `JERARQUIAS`  | JERARQUIA_ID, DIMENSION_ID, PROYECTO_ID, JERARQUIA, NIVELES_JSON, ...           |
| `WORKFLOWS`   | WORKFLOW_ID, PROYECTO_ID, WORKFLOW, DESCRIPCION, ...                            |
| `WORKFLOWS_PASOS` | PASO_ID, WORKFLOW_ID, PASO, ORDEN, ES_PASO0, INICIO_TIPO, REVISION, FIN_TIPO, DRIVER_DIMENSION_ID, DRIVER_MODO, ... |
| `WORKFLOWS_PASOS_DRIVER_VALORES` | PASO_ID, VALOR (valores concretos del driver)                  |
| `WORKFLOWS_PASOS_VARIABLES` | PASO_ID, VARIABLE_ID, NOMBRE, ETIQUETA, TIPO                        |
| `WORKFLOWS_PASOS_BLOQUES` | PASO_ID, BLOQUE_ID, TITULO, ORDEN                                      |
| `WORKFLOWS_PASOS_TAREAS` | BLOQUE_ID, TAREA_ID, TIPO, NOMBRE, REF_ID, REF_NOMBRE, ORDEN            |
| `WORKFLOWS_PASOS_TAREAS_VALORES` | TAREA_ID, CLAVE, ETIQUETA, TIPO, VALOR, OCULTAR                |
| `WORKFLOWS_RUNS` | RUN_ID, WORKFLOW_ID, PROYECTO_ID, NOMBRE, ESTADO, ...                                |
| `WORKFLOWS_RUNS_INSTANCIAS` | RUN_ID, PASO_ID, INSTANCIA_ID, DRIVER_VALOR, ASIGNADO, ESTADO, ...       |
| `WORKFLOWS_RUNS_VARIABLES` | RUN_ID, INSTANCIA_ID, NOMBRE, VALOR                                       |

`DIMENSIONES.CAMPOS_JSON` guarda un array; el primer elemento es siempre la
clave principal (`__isPrimaryName: true`, nombre = identificador de la
dimensión). `CUBOS.CAMPOS_JSON` guarda un objeto `{ dimensions: [...],
measures: [...] }`. `JERARQUIAS.NIVELES_JSON` guarda el array ordenado de
identificadores de columna, nivel superior primero.

Añadí `CAMPOS_JSON` (definición de columnas en JSON) sobre las 3 columnas
que pediste, para poder recargar el diseñador de campos al editar.

En BigQuery, cada proyecto Draco es un **dataset** `DRACO_<PROYECTO>`. En
Snowflake, cada proyecto Draco es un **esquema** `DRACO_<PROYECTO>` dentro
de la base de datos que elijas (por defecto `DRACO`). En ambos casos viven
ahí las tablas físicas `DRACO_<DIMENSION>` y `DRACO_<CUBO>`.

Los tipos de campo del diseñador (STRING, INTEGER, FLOAT, NUMERIC,
BOOLEAN, DATE, DATETIME, TIMESTAMP) se traducen automáticamente al tipo
físico de cada motor (`Provider.mapFieldType`) — por ejemplo FLOAT se
convierte en `FLOAT64` en BigQuery y se queda en `FLOAT` en Snowflake.

## 6. Estructura de archivos

```
index.html                    Landing / conectores
auth-callback.html            Callback OAuth de BigQuery
auth-callback-snowflake.html  Callback OAuth de Snowflake
app.html                      Panel principal
flow_run.html                 Ejecutar / monitorizar un flujo (?flujo_id=...)
css/theme.css                 Design system (reutilizado de la app anterior)
css/modal.css                 Modales y formularios
css/app.css                   Layout de la app
css/loads.css                 Estilos del módulo Interfaces (cargas de datos)
css/flows.css                 Estilos del módulo Flujos de carga
css/flow-run.css              Estilos de flow_run.html
js/config.js                  Client IDs y constantes editables
js/bigquery.js                Cliente REST de BigQuery
js/snowflake.js                Cliente OAuth+PKCE / SQL API v2 de Snowflake
js/provider.js                 Capa de abstracción BigQuery ↔ Snowflake
js/schema.js                    Bootstrap del esquema de control
js/auth.js                       Lógica de login (solo en index.html)
js/ui.js                          Toasts, modales, confirmaciones, maximizar/contraer
js/storage.js                      Abstracción de storage (subida de ficheros); usado por flow_run.html y por semantic-model.js
js/dimensions.js                   Módulo Dimensiones
js/semantic-model.js                Genera y guarda el modelo semántico YAML de un cubo
js/cubes.js                         Módulo Cubos
js/loads.js                          Módulo Interfaces (cargas de datos)
js/flows.js                           Módulo Flujos de carga
js/flow-run.js                         Lógica de flow_run.html
js/workflows.js                        Módulo Workflows (definición)
js/workflow-runs.js                    Ejecuciones (runs) de un Workflow
css/workflows.css                       Estilos específicos de Workflows
js/app.js                               Controlador principal de app.html
python/interface_loader.py               Motor de mapeo de una interfaz (origen -> cubo)
python/storage_io.py                      Storage + lectura de ficheros a DataFrame
python/flow_runner.py                      Orquestador de un flujo (cadena de interfaces)
sql/00_control_schema.sql             Script SQL de referencia (ambos motores)
sql/01_snowflake_oauth_integration.sql Alta del cliente OAuth en Snowflake
proxy/cloudflare-worker.js             Proxy CORS opcional para Snowflake
```

## 7bis. Ejecución de flujos (nuevo)

Los **Flujos de carga** ya se pueden ejecutar de verdad, no solo modelar:

- **`python/flow_runner.py`** — orquestador: recibe el `FLUJO_ID` y las
  variables de la pantalla como parámetro de entrada, lee la cadena de
  interfaces (`FLUJOS_INTERFACES` / `FLUJOS_INTERFACES_TARGETS`) y va
  llamando, paso a paso, al motor de `python/interface_loader.py`,
  sustituyendo cada `target` (constante o variable de pantalla) por su
  valor real. Para los pasos de tipo **FICHERO** resuelve dos variables,
  `ruta_local` y `ruta_storage`, y antes de mapear descarga el fichero
  del storage con `python/storage_io.py` (que también sabe leerlo a un
  DataFrame según sus separadores/codificación). Cada paso queda
  registrado en las nuevas tablas de control `FLUJOS_RUNS` /
  `FLUJOS_RUN_STEPS` (estado, filas, error). Pensado para desplegarse
  como Snowflake Python Stored Procedure (ver docstring de `main()`).
- **`flow_run.html`** — pantalla standalone: `flow_run.html?flujo_id=...`
  pinta la pantalla de variables del flujo con inputs reales (incluido
  un nuevo tipo de variable, **FILE**, que muestra un selector de
  fichero) y dos botones:
  - **▶ Ejecutar**: sube cada fichero local seleccionado al storage
    (`js/storage.js` — en BigQuery basta con `DracoConfig.
    bigqueryUploadBucket`, ver sección 7ter) y a continuación lanza el
    orquestador (`DracoConfig.flowRunnerProcedure` vía `CALL` en
    Snowflake, o `DracoConfig.flowRunnerHttpEndpoint` en BigQuery — ver
    sección 7ter).
  - **🖥 Monitor**: repinta la cadena de interfaces del flujo (misma
    tarjeta que en el editor) coloreada según `FLUJOS_RUN_STEPS` de la
    última ejecución — gris = pendiente, azul pulsando = en ejecución,
    verde = completado, rojo = error — y la sondea cada 3s mientras
    siga en curso.
  - Desde el editor de flujos (`app.html`), el botón "▶ Ejecutar /
    Monitor" abre directamente esta pantalla para el flujo guardado.

## 7ter. Interfaces de fichero en BigQuery — despliegue paso a paso

Esta sección explica, para quien no ha tocado nunca Google Cloud, en qué
se convierte todo lo de Snowflake (stage + stored procedure Python) para
BigQuery, dónde queda guardado el fichero, cómo ver los procesos
ejecutados y qué hay que configurar en BigQuery/GCS. Son 3 piezas:

| Pieza                      | Snowflake                                   | BigQuery                                                        |
|----------------------------|----------------------------------------------|-------------------------------------------------------------------|
| Guardado del fichero       | Stage interno (`@DRACO_LANDING`)              | Bucket de **Google Cloud Storage** (GCS)                          |
| El "Python que se programa"| Stored Procedure Python (`SP_RUN_FLUJO`)      | **Cloud Run** (o Cloud Functions Gen2) ejecutando `python/cloud_run_main.py` |
| Cómo se lanza              | `CALL SP_RUN_FLUJO(...)` desde el navegador   | `POST` a la URL del servicio (`DracoConfig.flowRunnerHttpEndpoint`) |
| Motor de lectura/escritura | `SnowflakeEngine` (`python/interface_loader.py`) | `BigQueryEngine` (mismo fichero, misma lógica de mapeo)         |
| Monitorización             | Tablas `FLUJOS_RUNS`/`FLUJOS_RUN_STEPS` (igual en ambos) + logs de Snowflake | Las mismas tablas (ahora en BigQuery) + **Cloud Logging / Cloud Run** |

En BigQuery no existen los "Python Stored Procedures" nativos de
Snowflake, así que el mismo `python/flow_runner.py` se despliega como un
pequeño servicio HTTP (Cloud Run) al que la app llama por `fetch()`.

### Paso 1 — Crear el bucket donde se guardan los ficheros subidos

1. Google Cloud Console → **Cloud Storage → Buckets → Crear**.
2. Nombre único global, p.ej. `draco-landing-mi-proyecto`. Región: la
   misma que uses para BigQuery/Cloud Run (p.ej. `europe-west1`), para
   evitar costes de tráfico entre regiones. El resto de opciones
   (clase de almacenamiento Standard, control de acceso uniforme) se
   puede dejar por defecto.
3. **Permisos**: pestaña "Permisos" del bucket → "Otorgar acceso" →
   añade el usuario (o el grupo de Google) que va a usar Planning, rol
   **"Storage Object Creator"** (`roles/storage.objectCreator`). Esto es
   lo único que hace falta para que el navegador pueda subir el fichero
   directamente al bucket usando su propia sesión de Google (el mismo
   login que ya usa para BigQuery).
4. En `js/config.js`, rellena `bigqueryUploadBucket` con ese nombre:
   `bigqueryUploadBucket: "draco-landing-mi-proyecto"`.

Con esto, cuando subes un fichero desde `flow_run.html`, el navegador lo
manda con un `POST` directo a
`https://storage.googleapis.com/upload/storage/v1/b/<bucket>/o` usando
el mismo token OAuth de tu sesión de BigQuery — no hace falta URL
firmada ni backend intermedio (ver `Storage.uploadToGcsBucket` en
`js/storage.js`). El objeto queda guardado en el bucket con un nombre
plano tipo `flujo123__fichero__1700000000000_datos.csv` (sin
subcarpetas, ver `Storage.buildPath`).

### Paso 2 — En qué se convierte el Python (Cloud Run)

`python/interface_loader.py` y `python/storage_io.py` ya incluyen los
adaptadores para BigQuery (`BigQueryEngine` y `GCSStorage`, hermanos de
`SnowflakeEngine`/`SnowflakeStageStorage`): la lógica de mapeo
(`apply_mapping`) es exactamente la misma, solo cambia de dónde se lee y
dónde se escribe. El nuevo fichero `python/cloud_run_main.py` envuelve
`flow_runner.run_flow()` en un endpoint HTTP.

Desplegarlo (con `gcloud` ya instalado y autenticado — `gcloud init`):

1. Habilita las APIs necesarias (una vez por proyecto):
   ```
   gcloud services enable run.googleapis.com bigquery.googleapis.com storage.googleapis.com
   ```
2. Desde la carpeta `python/` de este proyecto:
   ```
   gcloud run deploy draco-flow-runner \
       --source . \
       --region europe-west1 \
       --allow-unauthenticated \
       --set-env-vars GCP_PROJECT_ID=<tu-proyecto>,CONTROL_DATASET=DRACO_CONTROL,UPLOAD_BUCKET=draco-landing-mi-proyecto
   ```
   (usa la misma región que el bucket). El flag `--allow-unauthenticated`
   hace pública la URL del servicio; si quieres una protección mínima
   extra sin montar IAM, añade también
   `--set-env-vars DRACO_SHARED_KEY=<una-clave-larga-aleatoria>` y
   guarda esa misma clave en el `fetch` que hace `js/flow-run.js` (puedes
   envolver `DracoConfig.flowRunnerHttpEndpoint` con tu propio wrapper si
   necesitas mandar esa cabecera; por defecto la app no la manda).
3. Al terminar, `gcloud` imprime una URL tipo
   `https://draco-flow-runner-xxxxx-ew.a.run.app`. Cópiala en
   `js/config.js` → `flowRunnerHttpEndpoint`.
4. **Permisos que necesita el propio servicio** (su cuenta de servicio,
   por defecto la "Compute Engine default service account" del
   proyecto): en el dataset `DRACO_CONTROL` (y en cada dataset
   `DRACO_<PROYECTO>` que uses como destino) rol **"BigQuery Data
   Editor"**, y a nivel de proyecto **"BigQuery Job User"** (para poder
   lanzar las queries/load jobs); en el bucket, rol **"Storage Object
   Viewer"** como mínimo (para poder descargar los ficheros subidos).

Alternativa: Cloud Functions Gen2 (mismo código, `requirements.txt`
incluido) si prefieres no usar `gcloud run deploy --source` — en la
consola: **Cloud Functions → Crear función → 2ª generación**, runtime
Python 3.11, punto de entrada `main`, sube los ficheros de `python/`, y
configura las mismas variables de entorno.

### Paso 3 — Cómo se monitoriza desde Google Cloud

Hay dos niveles de monitorización, iguales que en Snowflake:

- **Dentro de la app**: `flow_run.html` (pestaña "🖥 Monitor") sondea las
  tablas `FLUJOS_RUNS` / `FLUJOS_RUN_STEPS` cada 3s — esto funciona igual
  en BigQuery que en Snowflake, sin nada que configurar.
- **En Google Cloud Console** (para ver ejecuciones, errores de
  arranque, tiempos, etc.):
  - **Cloud Run → tu servicio (`draco-flow-runner`) → pestaña
    "Registros" (Logs)**: aquí aparece cada petición HTTP recibida, con
    su código de estado y cualquier `print`/excepción de
    `flow_runner.py`. También puedes filtrar por severidad (ERROR) desde
    el mismo panel.
  - **Cloud Run → pestaña "Métricas"**: nº de peticiones, latencia,
    instancias activas, uso de memoria/CPU — útil para ver si un flujo
    se queda colgado o si necesitas subir la memoria asignada al
    servicio (por defecto 512 MiB; para ficheros grandes puede hacer
    falta más).
  - **Cloud Logging** (`Logging → Explorador de registros`), filtrando
    por `resource.type="cloud_run_revision"` y
    `resource.labels.service_name="draco-flow-runner"`: el mismo log
    que en la pestaña de Cloud Run, pero con más opciones de búsqueda y
    creación de alertas (p.ej. avisarte por email si aparece un `ERROR`).
  - **BigQuery → Historial de trabajos ("Job history")** o pestaña
    "Personal history"/"Project history" del editor de BigQuery: cada
    `load_table_from_dataframe` de `BigQueryEngine.write_table` aparece
    aquí como un "load job", con filas cargadas, bytes procesados y
    errores si los hay.

### Resumen de lo que hay que configurar

- ✅ Un bucket de GCS (Paso 1) con permiso "Storage Object Creator" para
  quien use Planning.
- ✅ El dataset `DRACO_CONTROL` en BigQuery (ya se crea solo la primera
  vez que conectas, ver `js/schema.js` — no hace falta nada especial
  para que funcione con ficheros).
- ✅ Un servicio Cloud Run (o Cloud Function) con permisos de "BigQuery
  Data Editor" + "BigQuery Job User" + "Storage Object Viewer" (Paso 2).
- ✅ `js/config.js`: `bigqueryUploadBucket` y `flowRunnerHttpEndpoint`
  rellenos.
- ❌ No hace falta ninguna `SECURITY INTEGRATION`, `STAGE` ni
  `FILE_UPLOAD_CHUNKS` — eso es solo la parte de Snowflake
  (`sql/02_snowflake_file_upload.sql` no aplica aquí).

## 8. Modelo semántico (YAML)

Un fichero por cubo, generado y sobrescrito automáticamente cada vez
que se guarda ese cubo. Ruta dentro del stage:

```
semantic_models/<PROYECTO_identificador>/<CUBO_identificador>.yaml
```

(`<..._identificador>` = `Provider.toIdentifier(nombre)`, el mismo
identificador que se usa para nombrar las tablas físicas). Al guardar
el mismo cubo dos veces se **sobrescribe** el mismo fichero
(`overwrite=True` en `session.file.put_stream()`), no se acumulan
versiones.

### 8bis. LookML (.lkml) en GitHub — solo BigQuery

Cuando el proveedor activo es **BigQuery**, al guardar un cubo (además
del YAML de arriba) Draco genera también su equivalente **.lkml**
(`js/lkml-export.js`, mismo formato que ya exporta el add-in de Excel
en `ADDIN/src/js/lkmlExport.js`: un `view` con la tabla de hechos y sus
measures, un `view` por dimensión con sus atributos/PK y jerarquías
documentadas como comentario, y un `explore` con los joins) y hace
commit directo de ese fichero en el repositorio de GitHub configurado
en `js/config.js` → `DracoConfig.semanticModelGithub` (por defecto,
`https://github.com/tennistrackplus/SEMANTIC_MODEL`), vía la API REST
de contenidos (`js/github-repo.js`, mismo mecanismo que
`ADDIN/src/js/gitRepo.js`). Ruta dentro del repo (misma carpeta y
nombre que el YAML, solo cambia la extensión):

```
semantic_models/<PROYECTO_identificador>/<CUBO_identificador>.lkml
```

Hace falta rellenar `DracoConfig.semanticModelGithub.token` con un
Personal Access Token de GitHub con permiso de escritura sobre el
repo. Igual que el YAML, esto **nunca bloquea el guardado del cubo**:
si falla la generación o la subida (falta el token, sin conexión,
repo/rama no encontrados...), solo se avisa por toast.

Estructura completa (los tipos entre paréntesis son orientativos):

```yaml
format_version: 1                 # versión del formato del propio YAML

model:                            # metadatos de la generación
  name: <string>                    # nombre del cubo
  cube_id: <string>                 # CUBO_ID en DRACO_CONTROL.CUBOS
  project: <string>                 # nombre del proyecto Draco
  project_id: <string>              # PROYECTO_ID
  engine: snowflake|bigquery        # motor activo al generar el YAML
  database: <string>                # base de datos (Snowflake) o proyecto GCP (BigQuery)
  schema: <string>                  # esquema/dataset del proyecto, DRACO_<PROYECTO>
  generated_at: <ISO8601 timestamp> # momento de generación (UTC)

fact:                             # LA TABLA DE HECHOS -> pestaña model_Fact
  name: <string>                    # nombre del cubo
  table: <string>                   # tabla física, ej. DRACO_VENTAS (sin qualificar; usar model.database + model.schema)
  description: <string|null>
  measures:                         # medidas definidas libremente en el cubo
    - name: <string>
      column: <string>                # nombre de columna físico (identificador)
      type: STRING|INTEGER|FLOAT|NUMERIC|BOOLEAN|DATE|DATETIME|TIMESTAMP
  foreign_keys:                     # una por cada dimensión añadida al cubo
    - dimension: <string>             # nombre de la dimensión (== dimensions[].name)
      column: <string>                # columna FK en la tabla de hechos (mismo nombre que la clave de la dimensión)
      references_dimension: <string>  # redundante con "dimension", pensado para lectura directa
      references_column: <string>     # columna PK en la tabla de la dimensión (== dimensions[].key_column)

dimensions:                       # LAS DIMENSIONES -> pestaña model_dimension
  - name: <string>                  # nombre de la dimensión
    dimension_id: <string>          # DIMENSION_ID en DRACO_CONTROL.DIMENSIONES
    table: <string>                 # tabla física, ej. DRACO_CUENTA
    description: <string|null>
    key_attribute: <string>         # nombre del atributo clave principal (== nombre de la dimensión)
    key_column: <string>            # columna física de esa clave
    description_attribute: <string|null>  # atributo marcado con 🏷 en el diseñador; null si no hay ninguno
    description_column: <string|null>     # columna física de ese atributo; null si no hay ninguno
    attributes:                     # TODOS los atributos, incluida la propia clave principal
      - name: <string>
        column: <string>              # nombre de columna físico (identificador)
        type: STRING|INTEGER|FLOAT|NUMERIC|BOOLEAN|DATE|DATETIME|TIMESTAMP
        is_key: <bool>                 # true en la clave principal y en claves compuestas
        is_description: <bool>         # true solo en el atributo marcado con 🏷 (como mucho uno por dimensión)
    hierarchies:                    # LAS JERARQUÍAS DE ESTA DIMENSIÓN -> pestaña model_hier
      - name: <string>                # nombre de la jerarquía
        levels:                       # nivel superior primero
          - level: <int>                # 1-based
            attribute: <string>          # nombre del atributo en ese nivel
            column: <string>             # columna física de ese atributo
```

Notas para quien vaya a leer esto desde el add-in (`addin/src`):

- **`model_Fact`**: una fila de cabecera con `fact.name` / `fact.table`
  / `fact.description`, más una fila por `fact.measures[]` y una fila
  por `fact.foreign_keys[]` (o todo en la misma tabla si la pestaña
  mezcla medidas y FKs — a definir cuando se aborde el add-in).
- **`model_dimension`**: una fila por cada `dimensions[].attributes[]`,
  repitiendo `dimensions[].name` / `table` como columnas de contexto,
  y usando `is_key` / `is_description` para pintar los mismos dos
  ticks que hay en el diseñador de Draco.
- **`model_hier`**: una fila por cada `dimensions[].hierarchies[].levels[]`,
  repitiendo el nombre de la dimensión y de la jerarquía como columnas
  de contexto.
- Todas las claves foráneas y de jerarquía referencian atributos **por
  nombre y por columna física** a la vez, para que el add-in pueda
  usar la que le resulte más cómoda sin tener que recalcular
  identificadores.
- `description_attribute` / `description_column` pueden venir a
  `null`: no es obligatorio que una dimensión tenga uno marcado.

Por ahora **el add-in no lee este YAML todavía** — este cambio se
limita a generarlo y guardarlo de forma fiable en cada guardado de
cubo, dejando la estructura cerrada y documentada para abordar la
lectura desde `addin/src` como siguiente paso.

## 9. Siguientes pasos sugeridos

Dime por cuál seguimos: Funciones, Flujos de proceso, Roles (para
asignar ejecuciones a personas/grupos reales en vez de texto libre),
disparar de verdad las tareas de tipo plantilla/función/parametrización/
HTML durante una ejecución, o cualquier ajuste sobre lo ya construido, y
lo abordamos con la misma mecánica (listado por proyecto + alta/edición +
tabla de metadatos en `DRACO_CONTROL`, funcionando sobre ambos motores
vía `Provider`).

