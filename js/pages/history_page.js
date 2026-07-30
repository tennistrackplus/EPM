/**
 * BigQuery Excel Connector - Query History Page
 */

import { Store } from '../state.js';
import { Navigation } from '../navigation.js';
import { Utils } from '../utils.js';

export async function renderHistoryPage() {
  const history = Store.getState().queryHistory;

  setTimeout(() => {
    Utils.$$('.history-item').forEach(item => {
      item.addEventListener('click', () => {
        const sql = decodeURIComponent(item.getAttribute('data-sql'));
        Store.setState({ activeQuery: { ...Store.getState().activeQuery, sql } });
        Utils.showToast('Consulta cargada desde el historial', 'info');
        Navigation.navigateTo('editor');
      });
    });

    const btnClear = Utils.$('#btn-clear-history');
    if (btnClear) {
      btnClear.addEventListener('click', () => {
        Store.setState({ queryHistory: [] });
        Utils.showToast('Historial borrado', 'info');
        Navigation.navigateTo('history');
      });
    }
  }, 0);

  const mockHistory = history.length > 0 ? history : [
    {
      id: 1,
      sql: 'SELECT * FROM `bigquery-public-data.usa_names.usa_1910_2013` LIMIT 100;',
      timestamp: '2026-07-30 14:22:10',
      duration: '1.14s',
      bytesProcessed: '50.00 MB',
      status: 'SUCCESS'
    },
    {
      id: 2,
      sql: 'SELECT region, SUM(total_amount) FROM `bq-analytics-prod.sales_marts.fact_orders_v2` GROUP BY region;',
      timestamp: '2026-07-30 11:05:43',
      duration: '2.80s',
      bytesProcessed: '500.00 MB',
      status: 'SUCCESS'
    },
    {
      id: 3,
      sql: 'SELECT * FROM `bq-finance-reporting.invalid_table` LIMIT 10;',
      timestamp: '2026-07-29 16:40:12',
      duration: '0.45s',
      bytesProcessed: '0 Bytes',
      status: 'ERROR'
    }
  ];

  const historyListHtml = mockHistory.map(item => `
    <div class="card p-base mb-sm cursor-pointer history-item hover-card" data-sql="${encodeURIComponent(item.sql)}">
      <div class="flex items-center justify-between mb-xs">
        <span class="badge ${item.status === 'SUCCESS' ? 'badge-success' : 'badge-error'}">${item.status}</span>
        <span class="text-xs color-subtext font-mono">${item.timestamp}</span>
      </div>
      <div class="text-xs font-mono mb-xs overflow-hidden text-ellipsis" style="white-space: nowrap; max-width: 100%; color: var(--color-text);">
        ${item.sql}
      </div>
      <div class="text-xs color-subtext flex justify-between">
        <span>Tiempo: ${item.duration}</span>
        <span>Procesado: ${item.bytesProcessed}</span>
      </div>
    </div>
  `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base flex items-center justify-between">
        <div>
          <h1 class="text-xl font-bold">Historial de Consultas</h1>
          <p class="text-xs color-subtext">Registro de peticiones ejecutadas recientemente</p>
        </div>
        <button class="btn btn-secondary text-xs" id="btn-clear-history">Borrar</button>
      </div>

      <div class="history-list">
        ${historyListHtml}
      </div>
    </div>
  `;
}