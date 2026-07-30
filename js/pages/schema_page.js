/**
 * BigQuery Excel Connector - Table Schema Inspector Page
 */

import { Store } from '../state.js';
import { Navigation } from '../navigation.js';
import { Utils } from '../utils.js';

export async function renderSchemaPage() {
  const state = Store.getState();
  const { selectedProject, selectedDataset, selectedTable } = state;

  if (!selectedTable) {
    return `
      <div class="page-container p-base fade-in">
        <div class="card p-base text-center">
          <p class="text-xs color-subtext mb-sm">Selecciona una tabla para explorar su esquema.</p>
          <button class="btn btn-primary text-xs" onclick="window.appNavigate('tables')">Ir a Tablas</button>
        </div>
      </div>
    `;
  }

  // Mock Schema Fields
  const fields = [
    { name: 'order_id', type: 'STRING', mode: 'REQUIRED', description: 'Identificador único de pedido' },
    { name: 'customer_name', type: 'STRING', mode: 'NULLABLE', description: 'Nombre o Razón Social del cliente' },
    { name: 'region', type: 'STRING', mode: 'NULLABLE', description: 'Región comercial (EMEA, LATAM, APAC...)' },
    { name: 'total_amount', type: 'NUMERIC', mode: 'NULLABLE', description: 'Importe total antes de impuestos' },
    { name: 'transaction_date', type: 'DATE', mode: 'REQUIRED', description: 'Fecha contable del registro' }
  ];

  setTimeout(() => {
    const btnQuery = Utils.$('#btn-query-schema-table');
    if (btnQuery) {
      btnQuery.addEventListener('click', () => {
        const query = `SELECT * FROM \`${selectedProject}.${selectedDataset}.${selectedTable}\` LIMIT 100;`;
        Store.setState({ activeQuery: { ...Store.getState().activeQuery, sql: query } });
        Navigation.navigateTo('editor');
      });
    }
  }, 0);

  const rowsHtml = fields.map(f => `
    <tr class="border-b">
      <td class="p-xs font-semibold font-mono text-xs">${f.name}</td>
      <td class="p-xs text-xs"><span class="badge badge-info">${f.type}</span></td>
      <td class="p-xs text-xs color-subtext">${f.mode}</td>
      <td class="p-xs text-xs color-subtext">${f.description}</td>
    </tr>
  `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base flex items-center justify-between">
        <div>
          <div class="text-xs color-subtext mb-xs">${selectedDataset}.${selectedTable}</div>
          <h1 class="text-xl font-bold">Esquema DDL</h1>
        </div>
        <button class="btn btn-primary text-xs" id="btn-query-schema-table">Consultar Tabla</button>
      </div>

      <div class="card p-none overflow-hidden mb-base">
        <table class="w-full text-left" style="border-collapse: collapse;">
          <thead>
            <tr class="bg-subtle border-b text-xs color-subtext">
              <th class="p-xs">Campo</th>
              <th class="p-xs">Tipo</th>
              <th class="p-xs">Modo</th>
              <th class="p-xs">Descripción</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      </div>

      <div class="flex gap-sm">
        <button class="btn btn-secondary text-xs w-full" onclick="window.appNavigate('preview')">Ver Previsualización</button>
      </div>
    </div>
  `;
}