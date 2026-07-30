/**
 * BigQuery Excel Connector - Table Preview Page
 */

import { Store } from '../state.js';
import { Utils } from '../utils.js';

export async function renderPreviewPage() {
  const state = Store.getState();
  const { selectedTable } = state;

  if (!selectedTable) {
    return `
      <div class="page-container p-base fade-in">
        <div class="card p-base text-center">
          <p class="text-xs color-subtext mb-sm">No hay ninguna tabla seleccionada para previsualizar.</p>
          <button class="btn btn-primary text-xs" onclick="window.appNavigate('tables')">Volver a Tablas</button>
        </div>
      </div>
    `;
  }

  // Sample Mock Data
  const cols = ['order_id', 'customer_name', 'region', 'total_amount', 'transaction_date'];
  const sampleRows = [
    ['ORD-1001', 'Client Alpha 1', 'EMEA', '1250.50', '2026-01-02'],
    ['ORD-1002', 'Client Alpha 2', 'NORTH_AMERICA', '890.00', '2026-01-03'],
    ['ORD-1003', 'Client Alpha 3', 'LATAM', '3400.10', '2026-01-04'],
    ['ORD-1004', 'Client Alpha 4', 'APAC', '150.25', '2026-01-05']
  ];

  const headerHtml = cols.map(c => `<th class="p-xs font-mono text-xs border-b bg-subtle">${c}</th>`).join('');
  const bodyHtml = sampleRows.map(r => `
    <tr class="border-b">
      ${r.map(val => `<td class="p-xs text-xs font-mono">${val}</td>`).join('')}
    </tr>
  `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base">
        <div class="text-xs color-subtext mb-xs">Previsualización gratuita (Muestra de 4 filas)</div>
        <h1 class="text-xl font-bold">${selectedTable}</h1>
      </div>

      <div class="card p-none overflow-x-auto mb-base">
        <table class="w-full text-left" style="border-collapse: collapse;">
          <thead>
            <tr>${headerHtml}</tr>
          </thead>
          <tbody>
            ${bodyHtml}
          </tbody>
        </table>
      </div>

      <button class="btn btn-primary w-full text-xs" onclick="window.appNavigate('editor')">Abrir en Consola SQL</button>
    </div>
  `;
}