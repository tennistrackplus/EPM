/**
 * BigQuery Excel Connector - Scheduled Jobs Page
 */

import { Store } from '../state.js';
import { Utils } from '../utils.js';

export async function renderScheduledPage() {
  const jobs = Store.getState().scheduledJobs;

  setTimeout(() => {
    const btnNew = Utils.$('#btn-create-schedule');
    if (btnNew) {
      btnNew.addEventListener('click', () => {
        Utils.showModal({
          title: 'Programar Refresco Automático',
          bodyHtml: `
            <div class="flex flex-col gap-sm">
              <div>
                <label class="text-xs font-semibold block mb-xs">Nombre de la Tarea</label>
                <input type="text" id="sched-name" class="input-text w-full text-xs" placeholder="Ej: Refresco Diario Ventas">
              </div>
              <div>
                <label class="text-xs font-semibold block mb-xs">Frecuencia</label>
                <select id="sched-freq" class="input-text w-full text-xs">
                  <option value="hourly">Cada hora</option>
                  <option value="daily" selected>Diario (08:00 AM)</option>
                  <option value="weekly">Semanal (Lunes)</option>
                </select>
              </div>
              <div>
                <label class="text-xs font-semibold block mb-xs">Hoja Destino</label>
                <input type="text" id="sched-sheet" class="input-text w-full text-xs" value="BQ_Resultados">
              </div>
            </div>
          `,
          confirmText: 'Crear Programación',
          onConfirm: () => {
            const name = document.getElementById('sched-name').value || 'Tarea Programada';
            const freq = document.getElementById('sched-freq').value;
            const sheet = document.getElementById('sched-sheet').value;

            const newJobs = [
              ...jobs,
              { id: Date.now(), name, frequency: freq, targetSheet: sheet, active: true, lastRun: 'Nunca' }
            ];
            Store.setState({ scheduledJobs: newJobs });
            Utils.showToast('Automatización programada creada con éxito', 'success');
            window.appNavigate('scheduled');
          }
        });
      });
    }
  }, 0);

  const mockJobs = jobs.length > 0 ? jobs : [
    {
      id: 101,
      name: 'Actualización Diaria de Ventas',
      frequency: 'Diario (08:00 AM)',
      targetSheet: 'Reporte_Ventas',
      active: true,
      lastRun: '2026-07-30 08:00:00'
    },
    {
      id: 102,
      name: 'Consolidado Cierre Mensual',
      frequency: 'Mensual (Día 1)',
      targetSheet: 'Finanzas_BigQuery',
      active: false,
      lastRun: '2026-07-01 00:00:00'
    }
  ];

  const jobsListHtml = mockJobs.map(j => `
    <div class="card p-base mb-sm">
      <div class="flex items-center justify-between mb-xs">
        <div class="text-sm font-bold">${j.name}</div>
        <span class="badge ${j.active ? 'badge-success' : 'badge-info'}">${j.active ? 'Activa' : 'Pausada'}</span>
      </div>
      <div class="text-xs color-subtext flex flex-col gap-xs mb-sm">
        <div><strong>Frecuencia:</strong> ${j.frequency}</div>
        <div><strong>Hoja Destino:</strong> <span class="font-mono">${j.targetSheet}</span></div>
        <div><strong>Última Ejecución:</strong> ${j.lastRun}</div>
      </div>
      <div class="flex gap-xs">
        <button class="btn btn-secondary text-xs flex-1">Ejecutar Ahora</button>
        <button class="btn btn-secondary text-xs">${j.active ? 'Pausar' : 'Reanudar'}</button>
      </div>
    </div>
  `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base flex items-center justify-between">
        <div>
          <h1 class="text-xl font-bold">Automatizaciones</h1>
          <p class="text-xs color-subtext">Refrescos de datos programados en segundo plano</p>
        </div>
        <button class="btn btn-primary text-xs" id="btn-create-schedule">+ Nueva Tarea</button>
      </div>

      <div class="scheduled-jobs-list">
        ${jobsListHtml}
      </div>
    </div>
  `;
}