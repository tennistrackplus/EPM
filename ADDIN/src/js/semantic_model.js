/**
 * Modelo Semántico BigQuery Connector - Lógica de Interfaz y Generador de Hojas de Excel.
 * Conserva la integración completa con BigQuery, modales de dimensiones/medidas y árbol de selección.
 */

// Token de acceso OAuth2 para la API de BigQuery (almacenado en el navegador o sesión)
let accessToken = localStorage.getItem('google_access_token') || '';

// Estado global de la aplicación
let fieldsList = [];             // Lista de campos de la tabla de hechos actual
let currentFactProject = '';    // Proyecto BigQuery seleccionado
let currentFactDataset = '';    // Dataset BigQuery seleccionado
let currentFactTable = '';      // Tabla de Hechos seleccionada

let currentEditingIndex = -1;   // Índice del campo actualmente editado en los modales
let treeTarget = 'FACT';         // Identificador del origen de búsqueda árbol: 'FACT' o 'DIM'
let dimAttributesTemp = [];     // Atributos temporales para el modal de dimensiones

// Inicialización cuando la API de Office JS está lista
Office.onReady(() => {
  initSemanticModelApp();
});

/**
 * Registra todos los escuchadores de eventos para los botones de la interfaz y modales.
 */
function initSemanticModelApp() {
  // 1. Ayuda de búsqueda (Matchcode estilo SAP) para la Tabla de Hechos
  const sapSearchBtn = document.getElementById('sapSearchBtn');
  const factFullConcat = document.getElementById('factFullConcat');
  if (sapSearchBtn) sapSearchBtn.addEventListener('click', () => openTreeModal('FACT'));
  if (factFullConcat) factFullConcat.addEventListener('click', () => openTreeModal('FACT'));

  // 2. Cierre del modal del Árbol de Tablas
  const btnCloseTreeModal = document.getElementById('btnCloseTreeModal');
  if (btnCloseTreeModal) btnCloseTreeModal.addEventListener('click', closeTreeModal);

  // 3. Modal de Configuración de Medidas
  const btnCloseMeasureModal = document.getElementById('btnCloseMeasureModal');
  const btnSaveMeasureModal = document.getElementById('btnSaveMeasureModal');
  if (btnCloseMeasureModal) btnCloseMeasureModal.addEventListener('click', closeMeasureModal);
  if (btnSaveMeasureModal) btnSaveMeasureModal.addEventListener('click', saveMeasureConfig);

  // 4. Modal de Configuración de Dimensiones y Búsqueda de Tablas de Atributos
  const btnCloseDimModal = document.getElementById('btnCloseDimModal');
  const btnSaveDimModal = document.getElementById('btnSaveDimModal');
  const dimSapSearchBtn = document.getElementById('dimSapSearchBtn');
  const dimRelFullConcat = document.getElementById('dimRelFullConcat');

  if (btnCloseDimModal) btnCloseDimModal.addEventListener('click', closeDimModal);
  if (btnSaveDimModal) btnSaveDimModal.addEventListener('click', saveDimConfig);
  if (dimSapSearchBtn) dimSapSearchBtn.addEventListener('click', () => openTreeModal('DIM'));
  if (dimRelFullConcat) dimRelFullConcat.addEventListener('click', () => openTreeModal('DIM'));

  // 5. Botón Principal: Generar Modelo Semántico en Excel
  const btnGenerateModel = document.getElementById('btnGenerateModel');
  if (btnGenerateModel) {
    btnGenerateModel.addEventListener('click', handleGenerateModelClick);
  }
}

/**
 * Retorna las cabeceras estándar con el Bearer Token para consumir la API REST de BigQuery.
 */
function getAuthHeaders() {
  return {
    'Authorization': `Bearer ${accessToken}`,
    'Content-Type': 'application/json'
  };
}

/**
 * Obtiene la lista de proyectos en BigQuery para el usuario autenticado.
 */
async function fetchBigQueryProjects() {
  const response = await fetch('https://bigquery.googleapis.com/bigquery/v2/projects', {
    headers: getAuthHeaders()
  });
  if (!response.ok) throw new Error('Error al obtener la lista de proyectos de BigQuery.');
  const data = await response.json();
  return data.projects || [];
}

/**
 * Obtiene la lista de datasets para un proyecto específico.
 */
async function fetchBigQueryDatasets(projectId) {
  const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets`, {
    headers: getAuthHeaders()
  });
  if (!response.ok) throw new Error(`Error al obtener datasets del proyecto ${projectId}.`);
  const data = await response.json();
  return data.datasets || [];
}

/**
 * Obtiene la lista de tablas dentro de un dataset.
 */
async function fetchBigQueryTables(projectId, datasetId) {
  const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables`, {
    headers: getAuthHeaders()
  });
  if (!response.ok) throw new Error(`Error al obtener tablas del dataset ${datasetId}.`);
  const data = await response.json();
  return data.tables || [];
}

/**
 * Obtiene el esquema completo de columnas de una tabla seleccionada.
 */
async function fetchBigQuerySchema(projectId, datasetId, tableId) {
  const response = await fetch(`https://bigquery.googleapis.com/bigquery/v2/projects/${projectId}/datasets/${datasetId}/tables/${tableId}`, {
    headers: getAuthHeaders()
  });
  if (!response.ok) throw new Error(`Error al obtener esquema de la tabla ${tableId}.`);
  const data = await response.json();
  return data.schema || { fields: [] };
}

/**
 * Abre la ventana modal con la vista jerárquica de Proyectos -> Datasets -> Tablas.
 * @param {'FACT' | 'DIM'} target - Indica si se busca la tabla de hechos o de dimensión.
 */
async function openTreeModal(target) {
  treeTarget = target;
  const modal = document.getElementById('treeModal');
  const container = document.getElementById('treeContainer');

  if (!modal || !container) return;

  container.innerHTML = '<p style="padding:10px; font-style:italic; color:#64748b;">Cargando proyectos desde BigQuery...</p>';
  modal.style.display = 'block';

  try {
    const projects = await fetchBigQueryProjects();
    if (projects.length === 0) {
      container.innerHTML = '<p style="padding:10px; color:#ef4444;">No se encontraron proyectos disponibles.</p>';
      return;
    }

    const rootUl = document.createElement('ul');
    rootUl.className = 'tree-list';

    for (const proj of projects) {
      const projId = proj.id;
      const projLi = document.createElement('li');
      projLi.className = 'tree-item';

      const projHeader = document.createElement('div');
      projHeader.className = 'tree-header';
      projHeader.innerHTML = `<span class="tree-toggle">▶</span> 📁 <strong>${projId}</strong>`;

      const projChildrenUl = document.createElement('ul');
      projChildrenUl.className = 'tree-list tree-children';

      let datasetsLoaded = false;
      projHeader.addEventListener('click', async () => {
        const toggle = projHeader.querySelector('.tree-toggle');
        const isOpen = projChildrenUl.classList.contains('open');

        if (isOpen) {
          projChildrenUl.classList.remove('open');
          toggle.textContent = '▶';
        } else {
          projChildrenUl.classList.add('open');
          toggle.textContent = '▼';

          if (!datasetsLoaded) {
            projChildrenUl.innerHTML = '<li style="padding:4px 12px; font-style:italic; color:#94a3b8;">Cargando datasets...</li>';
            try {
              const datasets = await fetchBigQueryDatasets(projId);
              projChildrenUl.innerHTML = '';
              
              for (const ds of datasets) {
                const dsId = ds.datasetReference.datasetId;
                const dsLi = document.createElement('li');
                dsLi.className = 'tree-item';

                const dsHeader = document.createElement('div');
                dsHeader.className = 'tree-header';
                dsHeader.innerHTML = `<span class="tree-toggle">▶</span> 📂 <span>${dsId}</span>`;

                const dsChildrenUl = document.createElement('ul');
                dsChildrenUl.className = 'tree-list tree-children';

                let tablesLoaded = false;
                dsHeader.addEventListener('click', async (e) => {
                  e.stopPropagation();
                  const dsToggle = dsHeader.querySelector('.tree-toggle');
                  const isDsOpen = dsChildrenUl.classList.contains('open');

                  if (isDsOpen) {
                    dsChildrenUl.classList.remove('open');
                    dsToggle.textContent = '▶';
                  } else {
                    dsChildrenUl.classList.add('open');
                    dsToggle.textContent = '▼';

                    if (!tablesLoaded) {
                      dsChildrenUl.innerHTML = '<li style="padding:4px 12px; font-style:italic; color:#94a3b8;">Cargando tablas...</li>';
                      try {
                        const tables = await fetchBigQueryTables(projId, dsId);
                        dsChildrenUl.innerHTML = '';

                        for (const tbl of tables) {
                          const tblId = tbl.tableReference.tableId;
                          const tblLi = document.createElement('li');
                          tblLi.className = 'tree-item';

                          const tblHeader = document.createElement('div');
                          tblHeader.className = 'tree-header table-item';
                          tblHeader.innerHTML = `📄 ${tblId}`;

                          tblHeader.addEventListener('click', (ev) => {
                            ev.stopPropagation();
                            onSelectTreeTable(projId, dsId, tblId);
                          });

                          tblLi.appendChild(tblHeader);
                          dsChildrenUl.appendChild(tblLi);
                        }
                        tablesLoaded = true;
                      } catch (errTbl) {
                        dsChildrenUl.innerHTML = '<li style="padding:4px 12px; color:#ef4444;">Error al cargar tablas.</li>';
                      }
                    }
                  }
                });

                dsLi.appendChild(dsHeader);
                dsLi.appendChild(dsChildrenUl);
                projChildrenUl.appendChild(dsLi);
              }
              datasetsLoaded = true;
            } catch (errDs) {
              projChildrenUl.innerHTML = '<li style="padding:4px 12px; color:#ef4444;">Error al cargar datasets.</li>';
            }
          }
        }
      });

      projLi.appendChild(projHeader);
      projLi.appendChild(projChildrenUl);
      rootUl.appendChild(projLi);
    }

    container.innerHTML = '';
    container.appendChild(rootUl);
  } catch (error) {
    console.error('Error al poblar el árbol de BigQuery:', error);
    container.innerHTML = '<p style="padding:10px; color:#ef4444;">Error al consultar la API de BigQuery.</p>';
  }
}

/**
 * Cierra el modal de selección del árbol.
 */
function closeTreeModal() {
  const modal = document.getElementById('treeModal');
  if (modal) modal.style.display = 'none';
}

/**
 * Evento desencadenado al seleccionar una tabla de BigQuery desde el árbol.
 */
async function onSelectTreeTable(projectId, datasetId, tableId) {
  closeTreeModal();

  if (treeTarget === 'FACT') {
    currentFactProject = projectId;
    currentFactDataset = datasetId;
    currentFactTable = tableId;

    const factFullConcat = document.getElementById('factFullConcat');
    const hiddenProj = document.getElementById('factProject');
    const hiddenData = document.getElementById('factDataset');
    const hiddenTbl = document.getElementById('factTable');

    if (factFullConcat) factFullConcat.value = `${projectId}.${datasetId}.${tableId}`;
    if (hiddenProj) hiddenProj.value = projectId;
    if (hiddenData) hiddenData.value = datasetId;
    if (hiddenTbl) hiddenTbl.value = tableId;

    // Obtener esquema de la tabla de hechos
    try {
      const schema = await fetchBigQuerySchema(projectId, datasetId, tableId);
      const fields = schema.fields || [];

      fieldsList = fields.map((field, idx) => {
        const isNumeric = ['INTEGER', 'FLOAT', 'NUMERIC', 'BIGNUMERIC', 'INT64', 'FLOAT64'].includes(field.type);
        return {
          position: idx + 1, // Posición ordinal 1-based index (1, 2, ... 16, 17, 18)
          name: field.name,
          type: field.type,
          role: isNumeric ? 'METRIC' : 'DIMENSION',
          alias: field.name,
          enabled: true,
          aggregation: isNumeric ? 'SUM' : '',
          format: 'Auto',
          dimProject: '',
          dimDataset: '',
          dimTable: '',
          dimField: field.name,
          dimAttributes: []
        };
      });

      renderFieldsTable();
      const fieldsCard = document.getElementById('fieldsCard');
      if (fieldsCard) fieldsCard.style.display = 'block';

    } catch (err) {
      console.error('Error al cargar el esquema de la tabla de hechos:', err);
    }

  } else if (treeTarget === 'DIM') {
    const dimRelFullConcat = document.getElementById('dimRelFullConcat');
    const dimRelProject = document.getElementById('dimRelProject');
    const dimRelDataset = document.getElementById('dimRelDataset');
    const dimRelTable = document.getElementById('dimRelTable');

    if (dimRelFullConcat) dimRelFullConcat.value = `${projectId}.${datasetId}.${tableId}`;
    if (dimRelProject) dimRelProject.value = projectId;
    if (dimRelDataset) dimRelDataset.value = datasetId;
    if (dimRelTable) dimRelTable.value = tableId;

    // Cargar atributos de la tabla de dimensión seleccionada
    try {
      const schema = await fetchBigQuerySchema(projectId, datasetId, tableId);
      const fields = schema.fields || [];

      dimAttributesTemp = fields.map((f) => ({
        name: f.name,
        alias: f.name,
        isKey: f.name === currentEditingField?.name,
        enabled: true,
        hier1: false,
        hier2: false
      }));

      renderDimAttributesTable();
      const attrContainer = document.getElementById('attributesContainer');
      if (attrContainer) attrContainer.style.display = 'block';

    } catch (err) {
      console.error('Error al cargar los atributos de la dimensión:', err);
    }
  }
}

/**
 * Renderiza la lista de campos de la tabla de hechos en la UI principal.
 */
function renderFieldsTable() {
  const tbody = document.getElementById('fieldsTbody');
  if (!tbody) return;

  tbody.innerHTML = '';

  fieldsList.forEach((field, index) => {
    const tr = document.createElement('tr');

    const tdName = document.createElement('td');
    tdName.innerHTML = `<strong>${field.name}</strong> <span style="font-size:9px; color:#94a3b8;">(#${field.position})</span>`;

    const tdAlias = document.createElement('td');
    tdAlias.textContent = field.alias || field.name;

    const tdType = document.createElement('td');
    const isNum = field.role === 'METRIC';
    tdType.innerHTML = `<span style="padding:2px 6px; border-radius:3px; font-size:10px; font-weight:600; background:${isNum ? '#e0f2fe' : '#dcfce7'}; color:${isNum ? '#0369a1' : '#15803d'};">${field.role}</span>`;

    const tdEnable = document.createElement('td');
    tdEnable.className = 'checkbox-cell';
    const chkEnable = document.createElement('input');
    chkEnable.type = 'checkbox';
    chkEnable.checked = field.enabled !== false;
    chkEnable.addEventListener('change', (e) => {
      field.enabled = e.target.checked;
    });
    tdEnable.appendChild(chkEnable);

    const tdActions = document.createElement('td');
    const btnConfig = document.createElement('button');
    btnConfig.className = 'btn-sm';
    btnConfig.textContent = field.role === 'METRIC' ? '📐 Medida' : '🏛️ Dimensión';
    btnConfig.addEventListener('click', () => {
      if (field.role === 'METRIC') {
        openMeasureModal(index);
      } else {
        openDimModal(index);
      }
    });
    tdActions.appendChild(btnConfig);

    tr.appendChild(tdName);
    tr.appendChild(tdAlias);
    tr.appendChild(tdType);
    tr.appendChild(tdEnable);
    tr.appendChild(tdActions);

    tbody.appendChild(tr);
  });
}

function openMeasureModal(index) {
  currentEditingIndex = index;
  const field = fieldsList[index];
  if (!field) return;

  document.getElementById('modalMeasureFieldName').textContent = field.name;
  document.getElementById('modalMeasureAlias').value = field.alias || field.name;
  document.getElementById('modalMeasureAgg').value = field.aggregation || 'SUM';
  document.getElementById('modalMeasureFormat').value = field.format || 'Auto';

  document.getElementById('measureModal').style.display = 'block';
}

function closeMeasureModal() {
  document.getElementById('measureModal').style.display = 'none';
}

function saveMeasureConfig() {
  if (currentEditingIndex < 0 || currentEditingIndex >= fieldsList.length) return;

  const field = fieldsList[currentEditingIndex];
  field.alias = document.getElementById('modalMeasureAlias').value.trim() || field.name;
  field.aggregation = document.getElementById('modalMeasureAgg').value;
  field.format = document.getElementById('modalMeasureFormat').value;

  closeMeasureModal();
  renderFieldsTable();
}

function openDimModal(index) {
  currentEditingIndex = index;
  const field = fieldsList[index];
  if (!field) return;

  currentEditingField = field;
  document.getElementById('modalDimFieldName').textContent = field.name;
  document.getElementById('modalDimAlias').value = field.alias || field.name;

  const fullConcat = field.dimProject && field.dimDataset && field.dimTable ? `${field.dimProject}.${field.dimDataset}.${field.dimTable}` : '';
  document.getElementById('dimRelFullConcat').value = fullConcat;
  document.getElementById('dimRelProject').value = field.dimProject || '';
  document.getElementById('dimRelDataset').value = field.dimDataset || '';
  document.getElementById('dimRelTable').value = field.dimTable || '';

  dimAttributesTemp = field.dimAttributes ? [...field.dimAttributes] : [];
  if (dimAttributesTemp.length > 0) {
    renderDimAttributesTable();
    document.getElementById('attributesContainer').style.display = 'block';
  } else {
    document.getElementById('attributesContainer').style.display = 'none';
  }

  document.getElementById('dimensionModal').style.display = 'block';
}

function closeDimModal() {
  document.getElementById('dimensionModal').style.display = 'none';
}

function renderDimAttributesTable() {
  const tbody = document.getElementById('dimAttributesTbody');
  if (!tbody) return;

  tbody.innerHTML = '';

  dimAttributesTemp.forEach((attr, idx) => {
    const tr = document.createElement('tr');

    const tdName = document.createElement('td');
    tdName.textContent = attr.name;

    const tdAlias = document.createElement('td');
    const inputAlias = document.createElement('input');
    inputAlias.type = 'text';
    inputAlias.value = attr.alias || attr.name;
    inputAlias.style.fontSize = '11px';
    inputAlias.addEventListener('input', (e) => {
      attr.alias = e.target.value;
    });
    tdAlias.appendChild(inputAlias);

    const tdKey = document.createElement('td');
    tdKey.className = 'checkbox-cell';
    const chkKey = document.createElement('input');
    chkKey.type = 'checkbox';
    chkKey.checked = !!attr.isKey;
    chkKey.addEventListener('change', (e) => {
      attr.isKey = e.target.checked;
    });
    tdKey.appendChild(chkKey);

    const tdEnable = document.createElement('td');
    tdEnable.className = 'checkbox-cell';
    const chkEnable = document.createElement('input');
    chkEnable.type = 'checkbox';
    chkEnable.checked = attr.enabled !== false;
    chkEnable.addEventListener('change', (e) => {
      attr.enabled = e.target.checked;
    });
    tdEnable.appendChild(chkEnable);

    const tdHier1 = document.createElement('td');
    tdHier1.className = 'checkbox-cell';
    const chkHier1 = document.createElement('input');
    chkHier1.type = 'checkbox';
    chkHier1.checked = !!attr.hier1;
    chkHier1.addEventListener('change', (e) => {
      attr.hier1 = e.target.checked;
    });
    tdHier1.appendChild(chkHier1);

    const tdHier2 = document.createElement('td');
    tdHier2.className = 'checkbox-cell';
    const chkHier2 = document.createElement('input');
    chkHier2.type = 'checkbox';
    chkHier2.checked = !!attr.hier2;
    chkHier2.addEventListener('change', (e) => {
      attr.hier2 = e.target.checked;
    });
    tdHier2.appendChild(chkHier2);

    tr.appendChild(tdName);
    tr.appendChild(tdAlias);
    tr.appendChild(tdKey);
    tr.appendChild(tdEnable);
    tr.appendChild(tdHier1);
    tr.appendChild(tdHier2);

    tbody.appendChild(tr);
  });
}

function saveDimConfig() {
  if (currentEditingIndex < 0 || currentEditingIndex >= fieldsList.length) return;

  const field = fieldsList[currentEditingIndex];
  field.alias = document.getElementById('modalDimAlias').value.trim() || field.name;
  field.dimProject = document.getElementById('dimRelProject').value;
  field.dimDataset = document.getElementById('dimRelDataset').value;
  field.dimTable = document.getElementById('dimRelTable').value;
  field.dimAttributes = [...dimAttributesTemp];

  closeDimModal();
  renderFieldsTable();
}

/**
 * Evento disparado al hacer clic en "Generar Modelo Semántico en Excel".
 */
async function handleGenerateModelClick() {
  const factProj = document.getElementById('factProject')?.value || currentFactProject;
  const factData = document.getElementById('factDataset')?.value || currentFactDataset;
  const factTbl = document.getElementById('factTable')?.value || currentFactTable;

  if (!factProj || !factData || !factTbl) {
    alert('Por favor, selecciona una Tabla de Hechos válida antes de continuar.');
    return;
  }

  try {
    await generateExcelSemanticModel(factProj, factData, factTbl, fieldsList, "Modelo_Semantico");
  } catch (error) {
    console.error('Error al generar el modelo semántico:', error);
  }
}

/**
 * Genera el modelo semántico en hojas de cálculo de Excel (MODEL_DIMENSION y MODEL_MEASURES).
 * Asigna matrices 2D explícitas para prevenir el pegado de texto concatenado en una única celda.
 *
 * @param {string} factProject - Nombre del proyecto BigQuery de origen para los hechos.
 * @param {string} factDataset - Nombre del dataset BigQuery de origen para los hechos.
 * @param {string} factTable - Nombre de la tabla de hechos en BigQuery.
 * @param {Array<Object>} fieldsList - Lista de objetos de campo con su configuración.
 * @param {string} [modelName] - Nombre del modelo semántico (opcional).
 */
async function generateExcelSemanticModel(factProject, factDataset, factTable, fieldsList, modelName = "Modelo_Semantico") {
  try {
    await Excel.run(async (context) => {

      // ==========================================
      // 1. GENERACIÓN PESTAÑA: MODEL_DIMENSION
      // ==========================================
      let dimSheet = context.workbook.worksheets.getItemOrNullObject("MODEL_DIMENSION");
      await context.sync();

      if (dimSheet.isNullObject) {
        dimSheet = context.workbook.worksheets.add("MODEL_DIMENSION");
      } else {
        dimSheet.getRange().clear();
      }

      // Encabezados requeridos para Dimensiones en la matriz bidimensional
      const dimHeaders = [
        "FILA", 
        "DIMENSION", 
        "FACT_PROJECT", 
        "FACT_DATASET", 
        "FACT_TABLE", 
        "FACT_FIELD", 
        "DIM_PROJECT", 
        "DIM_DATASET", 
        "DIM_TABLE", 
        "DIM_FIELD"
      ];

      const dimRows = [];
      fieldsList.forEach((field, index) => {
        const isEnabled = field.enabled !== false;
        const isDimension = field.role === 'DIMENSION' || !field.role;

        if (isEnabled && isDimension) {
          // Posición ordinal en la tabla de hechos (1-based index)
          const fila = field.position !== undefined ? field.position : (index + 1);
          const factProj = factProject || "";
          const factData = factDataset || "";
          const factTbl = factTable || "";
          const factFld = field.name || "";

          // Si el campo tiene una tabla de dimensión vinculada con atributos
          if (field.dimAttributes && field.dimAttributes.length > 0) {
            field.dimAttributes.forEach((attr) => {
              if (attr.enabled !== false) {
                dimRows.push([
                  fila,
                  attr.alias || attr.name,
                  factProj,
                  factData,
                  factTbl,
                  factFld,
                  field.dimProject || factProj,
                  field.dimDataset || factData,
                  field.dimTable || factTbl,
                  attr.name || factFld
                ]);
              }
            });
          } else {
            // Caso sin tabla de atributos ajena (apunta a los hechos o tabla ligada simple)
            const dimension = field.alias || field.name;
            const dimProj = field.dimProject || factProj;
            const dimData = field.dimDataset || factData;
            const dimTbl = field.dimTable || factTbl;
            const dimFld = field.dimField || factFld;

            dimRows.push([
              fila, 
              dimension, 
              factProj, 
              factData, 
              factTbl,
              factFld, 
              dimProj, 
              dimData, 
              dimTbl, 
              dimFld
            ]);
          }
        }
      });

      // Construir la matriz bidimensional completa (Cabecera + Filas)
      const dimMatrix = [dimHeaders, ...dimRows];
      if (dimRows.length > 0) {
        // Asignación explícita mediante rango por índices para garantizar columnas independientes
        const dimRange = dimSheet.getRangeByIndexes(0, 0, dimMatrix.length, dimHeaders.length);
        dimRange.values = dimMatrix;

        // Formatear como tabla nativa de Excel
        const dimTable = dimSheet.tables.add(dimRange, true /* hasHeaders */);
        dimTable.name = "Table_MODEL_DIMENSION_" + Date.now();
        dimTable.getHeaderRowRange().format.fill.color = "#0F6CBD";
        dimTable.getHeaderRowRange().format.font.color = "#FFFFFF";
        dimTable.getHeaderRowRange().format.font.bold = true;
        
        dimSheet.getUsedRange().format.autofitColumns();
      }

      // ==========================================
      // 2. GENERACIÓN PESTAÑA: MODEL_MEASURES
      // ==========================================
      let measSheet = context.workbook.worksheets.getItemOrNullObject("MODEL_MEASURES");
      await context.sync();

      if (measSheet.isNullObject) {
        measSheet = context.workbook.worksheets.add("MODEL_MEASURES");
      } else {
        measSheet.getRange().clear();
      }

      // Encabezados requeridos para Métricas/Medidas
      const measHeaders = [
        "FILA", 
        "MEASURE", 
        "FACT_PROJECT", 
        "FACT_DATASET", 
        "FACT_TABLE", 
        "FACT_FIELD", 
        "AGREGATION", 
        "FORMAT"
      ];

      const measRows = [];
      fieldsList.forEach((field, index) => {
        const isEnabled = field.enabled !== false;
        const isMeasure = field.role === 'METRIC' || field.role === 'MEASURE';

        if (isEnabled && isMeasure) {
          const fila = field.position !== undefined ? field.position : (index + 1);
          const measure = field.alias || field.name;
          const factProj = factProject || "";
          const factData = factDataset || "";
          const factTbl = factTable || "";
          const factFld = field.name || "";
          const aggregation = field.aggregation || "SUM";
          const format = field.format || "Auto";

          measRows.push([
            fila, 
            measure, 
            factProj, 
            factData, 
            factTbl,
            factFld, 
            aggregation, 
            format
          ]);
        }
      });

      // Construir la matriz bidimensional completa para medidas
      const measMatrix = [measHeaders, ...measRows];
      if (measRows.length > 0) {
        // Asignación de matriz 2D limpia por rango
        const measRange = measSheet.getRangeByIndexes(0, 0, measMatrix.length, measHeaders.length);
        measRange.values = measMatrix;

        // Formatear como tabla nativa de Excel
        const measTable = measSheet.tables.add(measRange, true /* hasHeaders */);
        measTable.name = "Table_MODEL_MEASURES_" + Date.now();
        measTable.getHeaderRowRange().format.fill.color = "#107C41";
        measTable.getHeaderRowRange().format.font.color = "#FFFFFF";
        measTable.getHeaderRowRange().format.font.bold = true;
        
        measSheet.getUsedRange().format.autofitColumns();
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Error al generar hojas de Excel del Modelo Semántico:", error);
    throw error;
  }
}