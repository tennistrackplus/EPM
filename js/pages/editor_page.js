/**
 * BigQuery Excel Connector - SQL Query Editor Page
 */

import { Store } from '../state.js';
import { BigQueryService } from '../bigquery.js';
import { Navigation } from '../navigation.js';
import { Utils } from '../utils.js';

export async function renderEditorPage() {
  const state = Store.getState();
  const activeQuery = state.activeQuery.sql || '';

  setTimeout(() => {
    const editorInput = Utils.$('#sql-editor-textarea');
    const dryRunInfo = Utils.$('#dry-run-info');
    const btnRun = Utils.$('#btn-execute-query');
    const btnSave = Utils.$('#btn-save-query');

    // Trigger Dry-Run on input change
    let timer;
    if (editorInput) {
      editorInput.addEventListener('input', () => {
        clearTimeout(timer);
        timer = setTimeout(async () => {
          const sql = editorInput.value;
          if (sql.trim().length > 10) {
            const res = await BigQueryService.dryRunQuery(sql);
            if (dryRunInfo) {
              dryRunInfo.textContent = `Esta consulta procesará ${res.formattedBytes} (Coste est.: ${res.estimatedCostUsd})`;
            }
          }
        }, 400);
      });
    }

    if (btnRun) {
      btnRun.addEventListener('click', async () => {
        const sql = editorInput.value;
        if (!sql.trim()) {
          Utils.showToast('Por favor introduce una consulta SQL válida.', 'error');
          return;
        }

        btnRun.disabled = true;
        btnRun.textContent = 'Ejecutando...';

        try {
          await BigQueryService.executeQuery(sql);
          Utils.showToast('Consulta ejecutada con éxito', 'success');
          Navigation.navigateTo('results');
        } catch (e) {
          Utils.showToast(`Error de ejecución: ${e.message}`, 'error');
        } finally {
          btnRun.disabled = false;
          btnRun.textContent = 'Ejecutar Consulta';
        }
      });
    }

    if (btnSave) {
      btnSave.addEventListener('click', () => {
        const sql = editorInput.value;
        if (!sql.trim()) return;

        const currentSaved = Store.getState().savedQueries;
        const newSaved = [
          { id: Date.now(), name: `Consulta ${currentSaved.length + 1}`, sql, date: new Date().toLocaleDateString() },
          ...currentSaved
        ];
        Store.setState({ savedQueries: newSaved });
        Utils.showToast('Consulta guardada correctamente', 'success');
      });
    }
  }, 0);

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base flex items-center justify-between">
        <div>
          <h1 class="text-xl font-bold">Consola SQL</h1>
          <p class="text-xs color-subtext">Escribe consultas BigQuery SQL con estimación de costes</p>
        </div>
        <button class="btn btn-secondary text-xs" id="btn-save-query">Guardar SQL</button>
      </div>

      <div class="card p-xs mb-sm">
        <textarea id="sql-editor-textarea" class="input-text w-full font-mono text-xs" style="min-height: 180px; resize: vertical;" spellcheck="false" placeholder="SELECT * FROM \`project.dataset.table\` LIMIT 100;">${activeQuery}</textarea>
      </div>

      <div class="flex items-center justify-between mb-base">
        <span class="text-xs color-subtext font-mono" id="dry-run-info">Esperando consulta...</span>
        <button class="btn btn-primary text-xs" id="btn-execute-query">Ejecutar Consulta</button>
      </div>
    </div>
  `;
}