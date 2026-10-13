# Informes (generador y visor) en Draco Planning — cambios

Descomprime este zip **encima de la raíz del repositorio** (las rutas ya empiezan por `PLANNING/`).
No hay que tocar `config.js` ni instalar nada: es JavaScript puro como el resto.

## Marca Orange
Nuevo perfil de marca **Orange** (corporativo.orange.es): menú negro con enlaces blancos en negrita, naranja #FF7900 como único color de marca, fondo blanco, textos en negro, Helvetica y esquinas rectas. Para el menú negro se añade la opción de estilo `nav:"black"`. Cambia solo `informes-herramienta.html`.

## Rediseño del editor: estructura y propiedades (solo estético)
- **Estructura (paleta):** Menú, Sección, Fila, Pestañas y Pie. Cabecera, Preguntas y Llamada a la acción pasan a ser **tipos de sección**; Menú sale desactivado (ya existe) y Fila no hace nada por ahora.
- **Página:** título, tipografía y «aprovechar la pantalla»; sin botón del menú, exportación ni «añadir sección».
- **Menú:** marca, icono (sin efecto), opciones de menú (añadir, quitar, nombre; no enlazadas a las secciones) y botones (acción + Exportar).
- **Sección:** selector de tipo (Cabecera, Contenido, Canvas —sin efecto—, Preguntas frecuentes, Llamada a la acción), Fondo, Distribución, Reparto, Espaciado y Notas del orador. Al cambiar de tipo se conservan etiqueta, título y subtítulo, se pide confirmación si dejan de verse elementos y se guardan en `_prev` (no viaja en el HTML exportado) para poder volver.
- **Etiqueta, Título, Subtítulo y Botón:** cada uno con su panel. Un clic en el título, la etiqueta o el subtítulo dentro de la página abre su panel; en la barra superior, el del Menú.
- **Pestañas:** título y lista de pestañas (añadir, quitar, nombre); cada pestaña con Nombre, Distribución, Reparto, Espaciado (se aplica en la página) y Notas del orador (se usan en las diapositivas de esa pestaña).
- Cambia solo `informes-herramienta.html` (más versiones en `generador.html`, `visor.html`, `generador-app.js`, `visor-app.js`).

## Corrección: al guardar una tabla Draco daba «Unclosed string literal»
El SQL generado por el addin lleva saltos de línea y BigQuery no admite saltos de línea reales dentro de un texto entre comillas simples. Ahora los textos del `INSERT`/`UPDATE` de `TABLAS_DRACO` se escriben con los saltos escapados (`\n`), que entienden BigQuery y Snowflake, y vuelven a ser saltos al leerlos. El mismo fallo afectaba al guardado de **informes** cuando su descripción tenía varias líneas: también corregido en `js/informes.js`. Cambia `js/draco-tables.js`, `js/informes.js` (y la versión en `generador.html` / `visor.html`).

## Ajuste: espaciado de la tabla multidimensional
Ahora tiene el mismo espaciado que la tabla normal, en Planning y en el informe: misma altura de fila (25 px en Planning, 45,4 px en el informe), mismo relleno de celda y mismo tamaño de texto; las etiquetas de fila van centradas y con el tamaño del texto normal (antes 30 px / 43,4 px y etiquetas de 14 px). Cambia `css/app.css` e `informes-herramienta.html` (más los números de versión de los html y de `generador-app.js` / `visor-app.js`).

## Corrección: Tablas Draco no dejaba elegir modelo semántico
El modelo de un cubo solo se leía del `.lkml` de GitHub. Ahora, si no existe o falla (y siempre con Snowflake, que no genera `.lkml`), se reconstruye desde el propio cubo en la base de datos. El diseñador muestra un aviso persistente con el motivo cuando un modelo no carga (y reintenta al volver a elegirlo) y avisa si el proyecto no tiene cubos. Además, las páginas de `tablas-draco/` llevan ahora número de versión en sus scripts (caché). Cambia `js/draco-tables.js`, `js/draco-table-designer.js`, `css/draco-tables.css`, `generador.html`, `visor.html` y `tablas-draco/` (regenerado con `python3 tools/sync-addin.py`).

## Ajuste: tabla multidimensional
Sin cabeceras de nombre de campo (esquina, fila de campos de Filas y título del Valor): solo los valores de Filas y Columnas; el nombre del campo queda como tooltip. Las celdas siguen el tipo de la columna del bloque **Valor** (si su ayuda de búsqueda es Checkbox, todas son checkbox; las cabeceras nunca llevan campos de entrada). En la validación de un campo ya no se muestra el texto de ayuda de «Completar con todos los valores». Cambia `js/table-updates.js`, `css/app.css` e `informes-herramienta.html` (más los números de versión de los html y de `generador-app.js` / `visor-app.js`).

## Novedad: Tablas Draco (informes tipo tabla dinámica del addin)
Nuevo componente **Tabla Draco** en el generador de informes. En *Propiedades*, el botón **Diseñador del informe** abre un pseudo-Excel solo para diseñar (rejilla de solo lectura + el taskpane del addin con todas sus opciones). Se guarda la definición y el **SQL ya generado** en la tabla nueva `TABLAS_DRACO`; el visor ejecuta ese SQL sin regenerarlo y dibuja el resultado con los estilos del libro (con +/- de jerarquías). Si el cubo cambia después, avisa de «SQL desfasado».
- **Archivos nuevos:** `PLANNING/js/draco-tables.js`, `PLANNING/js/draco-table-designer.js`, `PLANNING/css/draco-tables.css`, `PLANNING/tablas-draco/` (GENERADO: motor y diseñador del addin; no editar), `tools/sync-addin.py` y `tools/templates/` (generan la carpeta anterior desde `ADDIN/src`: `python3 tools/sync-addin.py`).
- **Archivos tuyos modificados:** `js/schema.js` (tabla `TABLAS_DRACO`), `js/informes.js` (la crea con las demás), `generador.html`, `visor.html`, `js/generador-app.js`, `js/visor-app.js`, `informes-herramienta.html`, `README.md`.
- **Pendiente / no probado:** BigQuery y Snowflake reales (solo mi simulador); escritura de planificación (oculta); los informes **estáticos/asimétricos NO están soportados** (el diseñador oculta su interruptor); modelos de Power BI.

## Novedad de esta versión: editor con barras laterales
`PLANNING/informes-herramienta.html` incluye ahora una barra izquierda (**Componentes** en iconos + **Estructura** en árbol con filas y columnas) y la barra derecha de **Propiedades**; las dos se contraen y expanden. Se arrastra y suelta en el árbol (reordenar, mover bloques entre secciones, soltar componentes de la paleta). Dentro de Draco se abre directamente en modo diseño. Ver README §7quater.

## Novedad: enlace directo «solo el informe»
`visor.html?informe_id=<id>` abre ahora el informe SIN barras (ni la de Draco, ni el selector de proyecto, ni la lista), con un botón discreto abajo a la derecha: ⛶ pantalla completa del navegador (Esc para salir) y ☰ volver al visor con lista. `&lista=1` abre el visor normal; en el visor, *⛶ Completo* y *🔗 Enlace* (copia la dirección limpia). Cambia solo `visor.html`, `js/visor-app.js` y `css/informes.css`.

## Novedad: «Actualizar tabla» dentro de un informe
Nuevo componente junto a *Tabla* (solo dentro de Draco). En *Propiedades* eliges una **vista de Actualización de tablas** del proyecto y rellenas **sus variables**; en el informe se muestra una rejilla editable con los estilos del libro y el mismo comportamiento que la pantalla de Actualización de tablas (filtro/orden, añadir/eliminar, pegar de Excel, desplegables y búsqueda con texto, importar/exportar, claves y obligatorios, grabar con resumen).

- **Archivo nuevo:** `PLANNING/js/informes-datos.js` (atiende las peticiones del informe con la sesión del usuario, reutilizando `TableUpdates`).
- **Cambios en generador.html / visor.html / sus controladores:** cargan `table-updates.js`, la librería XLSX y `informes-datos.js`, y enrutan las peticiones.
- **Seguridad:** el informe nunca envía SQL, tabla ni filtro; Draco los recalcula desde la definición guardada, dentro del proyecto activo.
- **Diferencias con la pantalla original:** variables fijadas al diseñar; sin variables de tipo archivo; sin redimensionar columnas; CSV importados como texto (la original convierte `2026-01` en fecha); importación emparejada por la clave completa; si el INSERT falla al grabar se intenta restaurar los datos; tope de 50.000 filas. Grabar sigue siendo «borrar el subconjunto filtrado y reinsertar», sin transacción, como la original.
- **En diapositivas, PDF y PowerPoint** el bloque se sustituye por una nota.

## Archivos nuevos (8)
| Archivo | Para qué |
|---|---|
| `PLANNING/generador.html` + `js/generador-app.js` | Diseñar informes. Sesión y selector de proyecto como `task.html`; ejecuta las operaciones de la herramienta con la sesión del usuario |
| `PLANNING/visor.html` + `js/visor-app.js` | Lista del proyecto (buscador, favoritos) y apertura de informes; enlace directo `visor.html?informe_id=<id>` |
| `PLANNING/js/informes-datos.js` | Actualización de tablas dentro de un informe |
| `PLANNING/js/informes.js` | Acceso a datos (`INFORMES`, `INFORMES_FAVORITOS`) vía `Provider` (BigQuery y Snowflake) |
| `PLANNING/css/informes.css` | Estilos con los tokens de `theme.css` |
| `PLANNING/informes-herramienta.html` | La herramienta de informes (un solo archivo autocontenido) |

## Archivos modificados (4) — ver `CAMBIOS.diff`
- `js/schema.js`: +2 DDL (`INFORMES`, `INFORMES_FAVORITOS`) y los dos nombres en `TABLES`.
- `app.html`: dos iconos en la barra superior (Generador, Visor) y `?v=` de `schema.js`.
- `task.html`: solo el `?v=` de `schema.js` (para que se carguen las tablas nuevas).
- `README.md`: sección 7quater y lista de archivos.

## Tablas nuevas en DRACO_CONTROL (clave por proyecto)
- `INFORMES`: INFORME_ID, **PROYECTO_ID**, NOMBRE, DESCRIPCION, SPEC_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION
- `INFORMES_FAVORITOS`: **PROYECTO_ID**, INFORME_ID, USUARIO, FECHA_ALTA

Se crean con el bootstrap del siguiente inicio de sesión, y además `Informes` las garantiza la primera
vez que se usan en cada pestaña (por si alguien tiene una sesión viva al desplegar).

## Qué se ha comprobado y qué NO
Comprobado (navegador real, 244 comprobaciones automáticas nuevas (más la regresión anterior de la herramienta) contra un BigQuery SIMULADO sobre SQLite que
interpreta el SQL con el escapado de BigQuery): guardar/sobrescribir/guardar como/abrir/eliminar, nombre
único por proyecto, aislamiento entre proyectos, favoritos por usuario, límite de 500 KB, sesión caducada
con reintento, enlace directo entre proyectos, caracteres difíciles (comillas, barras, saltos de línea, emoji),
contenido hostil, y que la página generada no puede leer el token de `localStorage`.

**NO comprobado**: contra un BigQuery real, ni contra Snowflake (el SQL es el mismo para ambos salvo la
conversión de fechas y `SESSION_USER()`/`CURRENT_USER()`), ni el despliegue en vuestro hosting.
Primera prueba recomendada: iniciar sesión, abrir `generador.html`, cargar el ejemplo, Guardar, y mirar la fila
en `DRACO_CONTROL.INFORMES`.

## Límites conocidos
- Sin roles: quien tenga acceso a `DRACO_CONTROL` ve y edita los informes del proyecto.
- Un solo diseñador por informe: guardar sobrescribe (no hay versiones ni control de edición simultánea).
- El puente para abrir tablas/SQL desde un informe existe pero responde "no disponible" (siguiente fase).
- Los `.js` de la raíz de `PLANNING` son copias antiguas de los de `js/` (varios difieren); la app carga los de `js/`.
