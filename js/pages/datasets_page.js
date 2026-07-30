/**
 * BigQuery Excel Connector - Datasets Explorer Page
 */

import { BigQueryService } from '../bigquery.js';
import { Store } from '../state.js';
import { Navigation } from '../navigation.js';
import { Utils } from '../utils.js';

export async function renderDatasetsPage() {
  const state = Store.getState();
  const projectId = state.selectedProject;

  if (!projectId) {
    return `
      <div class="page-container p-base fade-in">
        <div class="card p-base text-center">
          <p class="text-xs color-subtext mb-sm">Debes seleccionar un Proyecto antes de ver sus Datasets.</p>
          <button class="btn btn-primary text-xs" onclick="window.appNavigate('projects')">Ir a Proyectos</button>
        </div>
      </div>
    `;
  }

  const datasets = await BigQueryService.getDatasets(projectId);

  setTimeout(() => {
    Utils.$$('.dataset-item').forEach(item => {
      item.addEventListener('click', () => {
        const datasetId = item.getAttribute('data-dataset-id');
        Store.setState({ selectedDataset: datasetId, selectedTable: null });
        Utils.showToast(`Dataset seleccionado: ${datasetId}`, 'info');
        Navigation.navigateTo('tables');
      });
    });
  }, 0);

  const datasetsHtml = datasets.map(d => `
    <div class="card p-base mb-sm cursor-pointer dataset-item hover-card ${d.id === state.selectedDataset ? 'border-primary' : ''}" data-dataset-id="${d.id}">
      <div class="flex items-center justify-between">
        <div>
          <div class="text-sm font-bold">${d.id}</div>
          <div class="text-xs color-subtext">Ubicación: ${d.location}</div>
        </div>
        <svg class="fluent-icon" viewBox="0 0 20 20" width="16" height="16" fill="currentColor"><path d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z"/></svg>
      </div>
    </div>
  `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base">
        <div class="text-xs color-subtext mb-xs">Proyecto: <span class="font-semibold">${projectId}</span></div>
        <h1 class="text-xl font-bold">Datasets</h1>
        <p class="text-xs color-subtext">Conjuntos de datos estructurados disponibles</p>
      </div>

      <div class="datasets-list">
        ${datasetsHtml}
      </div>
    </div>
  `;
}