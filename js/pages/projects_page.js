/**
 * BigQuery Excel Connector - Projects Explorer Page
 */

import { BigQueryService } from '../bigquery.js';
import { Store } from '../state.js';
import { Navigation } from '../navigation.js';
import { Utils } from '../utils.js';

export async function renderProjectsPage() {
  const projects = await BigQueryService.getProjects();
  const selectedProject = Store.getState().selectedProject;

  setTimeout(() => {
    Utils.$$('.project-item').forEach(item => {
      item.addEventListener('click', () => {
        const projectId = item.getAttribute('data-project-id');
        Store.setState({ selectedProject: projectId, selectedDataset: null, selectedTable: null });
        Utils.showToast(`Proyecto seleccionado: ${projectId}`, 'info');
        Navigation.navigateTo('datasets');
      });
    });
  }, 0);

  const projectsHtml = projects.map(p => `
    <div class="card p-base mb-sm cursor-pointer project-item hover-card ${p.id === selectedProject ? 'border-primary' : ''}" data-project-id="${p.id}">
      <div class="flex items-center justify-between">
        <div>
          <div class="text-sm font-bold">${p.name}</div>
          <div class="text-xs color-subtext">ID: ${p.id}</div>
        </div>
        <span class="badge">${p.role}</span>
      </div>
    </div>
  `).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base">
        <h1 class="text-xl font-bold">Proyectos GCP</h1>
        <p class="text-xs color-subtext">Selecciona un proyecto de Google Cloud para inspeccionar sus recursos</p>
      </div>

      <div class="mb-base">
        <input type="text" class="input-text w-full text-xs" placeholder="Buscar proyecto por nombre o ID...">
      </div>

      <div class="projects-list">
        ${projectsHtml}
      </div>
    </div>
  `;
}