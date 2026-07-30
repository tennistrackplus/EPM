/**
 * BigQuery Excel Connector - Linked Sheets Manager Page
 */

import { ExcelService } from '../excel.js';
import { Utils } from '../utils.js';

export async function renderSheetsPage() {
  setTimeout(() => {
    const btnSyncAll = Utils.$('#btn-sync-all-sheets');
    if (btnSyncAll) {
      btnSyncAll.addEventListener('click', async () => {
        Utils.showToast('Sincronizando todas las hojas vinculadas con BigQuery...', 'info');
        await new Promise(r => setTimeout(r, 1500));
        Utils.showToast('Todas las hojas actualizadas correctamente.', 'success');
      });
    }
  }, 0);

  const mockBoundSheets = [
    {
      sheetName: 'BQ_Resultados',
      querySummary: 'SELECT * FROM `bigquery-public-data.usa_names.usa_1910_2013`...',
      rowCount: 100,
      lastSync: 'Hace 10 min'
    },
    {
      sheetName: 'Reporte_Ventas',
      querySummary: 'SELECT region, SUM(total_amount) FROM `bq-analytics-prod.sales_marts`...',
      rowCount: 4,
      lastSync: 'Hace 3 horas'
    }
  ];

  const sheetsListHtml = mockBoundSheets.map(s => `
    <div class="card p-base mb-sm">
      <div class="flex items-center justify-between mb-xs">
        <div class="text-sm font-bold font-mono color-primary">${s.sheetName}</div>
        <span class="badge badge-info">${s.lastSync}</span>
      </div>
      <div class="text-xs font-mono color-subtext mb-xs overflow-hidden text-ellipsis" style="white-space: nowrap;">
        ${s.querySummary}
      </div>
      <div class="flex items-center justify-between text-xs color-subtext mt-xs">
        <span>Filas vinculadas: ${s.rowCount}</span>
        <button class="btn btn-secondary text-xs">Refrescar Hoja</button>
      </div>
    </div>
  `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base flex items-center justify-between">
        <div>
          <h1 class="text-xl font-bold">Hojas Vinculadas</h1>
          <p class="text-xs color-subtext">Vínculos de datos activos en este libro de Excel</p>
        </div>
        <button class="btn btn-primary text-xs" id="btn-sync-all-sheets">Refrescar Todo</button>
      </div>

      <div class="linked-sheets-list">
        ${sheetsListHtml}
      </div>
    </div>
  `;
}