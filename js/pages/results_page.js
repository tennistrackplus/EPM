/**
 * BigQuery Excel Connector - Query Results & Excel Export Page
 */

import { Store } from '../state.js';
import { ExcelService } from '../excel.js';
import { Utils } from '../utils.js';

export async function renderResultsPage() {
  const state = Store.getState();
  const result = state.activeQuery.lastResult;

  if (!result || !result.rows || result.rows.length === 0) {
    return `
      <div class="page-container p-base fade-in">
        <div class="card p-base text-center">
          <p class="text-xs color-subtext mb-sm">No hay resultados de consulta disponibles para mostrar.</p>
          <button class="btn btn-primary text-xs" onclick="window.appNavigate('editor')">Ir a Consola SQL</button>
        </div>
      </div>
    `;
  }

  setTimeout(() => {
    const btnExcel = Utils.$('#btn-export-to-excel');
    if (btnExcel) {
      btnExcel.addEventListener('click', async () => {
        await ExcelService.writeToSheet(result.columns, result.rows, 'BQ_Resultados');
      });
    }
  }, 0);

  const colsHtml = result.columns.map(c => `<th class="p-xs font-mono text-xs border-b bg-subtle">${c}</th>`).join('');
  const rowsHtml = result.rows.slice(0, 15).map(r => `
    <tr class="border-b">
      ${r.map(val => `<td class="p-xs text-xs font-mono">${val}</td>`).join('')}
    </tr>
  `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base flex items-center justify-between">
        <div>
          <h1 class="text-xl font-bold">Resultados</h1>
          <p class="text-xs color-subtext">${result.totalRows} filas | Tiempo: ${result.executionTimeMs}ms</p>
        </div>
        <button class="btn btn-primary text-xs" id="btn-export-to-excel">Insertar en Excel</button>
      </div>

      <div class="card p-none overflow-x-auto mb-base" style="max-height: 320px;">
        <table class="w-full text-left" style="border-collapse: collapse;">
          <thead style="position: sticky; top: 0; z-index: 10;">
            <tr>${colsHtml}</tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      </div>

      <p class="text-xs color-subtext text-center">Mostrando las primeras 15 filas en vista previa.</p>
    </div>
  `;
}