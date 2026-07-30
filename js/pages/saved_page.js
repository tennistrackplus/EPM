/**
 * BigQuery Excel Connector - Saved Queries Management Page
 */

import { Store } from '../state.js';
import { Navigation } from '../navigation.js';
import { Utils } from '../utils.js';

export async function renderSavedPage() {
  const savedQueries = Store.getState().savedQueries;

  setTimeout(() => {
    Utils.$$('.saved-query-item').forEach(item => {
      item.addEventListener('click', () => {
        const sql = item.getAttribute('data-sql');
        Store.setState({ activeQuery: { ...Store.getState().activeQuery, sql } });
        Utils.showToast('Consulta cargada en el editor', 'info');
        Navigation.navigateTo('editor');
      });
    });
  }, 0);

  const listHtml = savedQueries.length === 0 
    ? `<div class="card p-base text-center color-subtext text-xs">No tienes consultas guardadas.</div>`
    : savedQueries.map(q => `
      <div class="card p-base mb-sm cursor-pointer saved-query-item hover-card" data-sql="${encodeURIComponent(q.sql)}">
        <div class="flex items-center justify-between mb-xs">
          <div class="text-sm font-bold">${q.name}</div>
          <div class="text-xs color-subtext">${q.date}</div>
        </div>
        <div class="text-xs font-mono color-subtext overflow-hidden text-ellipsis" style="white-space: nowrap;">${q.sql}</div>
      </div>
    `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base">
        <h1 class="text-xl font-bold">Consultas Guardadas</h1>
        <p class="text-xs color-subtext">Biblioteca de consultas SQL reutilizables</p>
      </div>

      <div class="saved-queries-list">
        ${listHtml}
      </div>
    </div>
  `;
}