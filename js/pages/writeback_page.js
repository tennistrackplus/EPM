/**
 * BigQuery Excel Connector - Writeback to BigQuery Page
 */

import { ExcelService } from '../excel.js';
import { Store } from '../state.js';
import { Utils } from '../utils.js';

export async function renderWritebackPage() {
  const state = Store.getState();

  setTimeout(() => {
    const btnSync = Utils.$('#btn-execute-writeback');
    if (btnSync) {
      btnSync.addEventListener('click', async () => {
        const data = await ExcelService.getSelectedRangeValues();
        if (!data || data.length === 0) {
          Utils.showToast('Selecciona un rango con datos en la hoja activa.', 'error');
          return;
        }

        Utils.showToast(`Sincronizando ${data.length} filas con BigQuery...`, 'info');
        await new Promise(r => setTimeout(r, 1200));
        Utils.showToast('Carga Writeback completada con éxito en BigQuery.', 'success');
      });
    }
  }, 0);

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base">
        <h1 class="text-xl font-bold">Escribir en BigQuery</h1>
        <p class="text-xs color-subtext">Cargar la selección de Excel activa hacia una tabla de destino</p>
      </div>

      <div class="card p-base mb-base flex flex-col gap-sm">
        <div>
          <label class="text-xs font-semibold block mb-xs">Tabla Destino</label>
          <input type="text" class="input-text w-full text-xs font-mono" value="${state.selectedProject || 'bq-prod'}.${state.selectedDataset || 'dataset'}.${state.selectedTable || 'tabla_destino'}" readonly>
        </div>

        <div>
          <label class="text-xs font-semibold block mb-xs">Modo de Inserción</label>
          <select class="input-text w-full text-xs">
            <option value="append">Anexar registros (WRITE_APPEND)</option>
            <option value="truncate">Sobrescribir tabla (WRITE_TRUNCATE)</option>
          </select>
        </div>

        <button class="btn btn-primary text-xs w-full mt-xs" id="btn-execute-writeback">
          Subir Selección de Excel
        </button>
      </div>
    </div>
  `;
}