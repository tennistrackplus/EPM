/**
 * BigQuery Excel Connector - Tables Explorer Page
 */

import { BigQueryService } from '../bigquery.js';
import { Store } from '../state.js';
import { Navigation } from '../navigation.js';
import { Utils } from '../utils.js';

export async function renderTablesPage() {
  const state = Store.getState();
  const { selectedProject, selectedDataset } = state;

  if (!selectedProject || !selectedDataset) {
    return `
      <div class="page-container p-base fade-in">
        <div class="card p-base text-center">
          <p class="text-xs color-subtext mb-sm">Debes seleccionar Proyecto y Dataset para explorar las tablas.</p>
          <button class="btn btn-primary text-xs" onclick="window.appNavigate('datasets')">Volver a Datasets</button>
        </div>
      </div>
    `;
  }

  const tables = await BigQueryService.getTables(selectedProject, selectedDataset);

  setTimeout(() => {
    Utils.$$('.table-item').forEach(item => {
      item.addEventListener('click', () => {
        const tableId = item.getAttribute('data-table-id');
        Store.setState({ selectedTable: tableId });
        Utils.showToast(`Tabla seleccionada: ${tableId}`, 'info');
        Navigation.navigateTo('schema');
      });
    });
  }, 0);

  const tablesHtml = tables.map(t => `
    <div class="card p-base mb-sm cursor-pointer table-item hover-card ${t.id === state.selectedTable ? 'border-primary' : ''}" data-table-id="${t.id}">
      <div class="flex items-center justify-between mb-xs">
        <div class="text-sm font-bold">${t.id}</div>
        <span class="badge ${t.type === 'VIEW' ? 'badge-info' : ''}">${t.type}</span>
      </div>
      <div class="text-xs color-subtext flex gap-base">
        <span>Filas: ${t.numRows ? Utils.formatNumber(t.numRows) : 'N/A'}</span>
        <span>Tamaño: ${t.sizeBytes ? Utils.formatBytes(t.sizeBytes) : 'N/A'}</span>
      </div>
    </div>
  `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base">
        <div class="text-xs color-subtext mb-xs">${selectedProject} / <span class="font-semibold">${selectedDataset}</span></div>
        <h1 class="text-xl font-bold">Tablas y Vistas</h1>
        <p class="text-xs color-subtext">Selecciona una entidad para ver su esquema o consultar</p>
      </div>

      <div class="tables-list">
        ${tablesHtml}
      </div>
    </div>
  `;
}