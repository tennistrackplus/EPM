/**
 * BigQuery Excel Connector - Home Dashboard Page
 */

import { Store } from '../state.js';
import { Navigation } from '../navigation.js';
import { Login } from '../login.js';

export async function renderHomePage() {
  const state = Store.getState();
  const user = Login.currentUser || { name: 'Invitado', tenant: 'Sin Licencia' };

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base">
        <h1 class="text-xl font-bold">Bienvenido, ${user.name}</h1>
        <p class="text-xs color-subtext">Panel de control y acceso rápido a conectores Enterprise</p>
      </div>

      <!-- General Status Card -->
      <div class="card p-base mb-base flex items-center justify-between" style="border-left: 4px solid var(--color-primary);">
        <div>
          <div class="text-xs font-semibold uppercase color-subtext mb-xs">Estado de Conexión</div>
          <div class="text-base font-bold flex items-center gap-xs">
            <span class="status-indicator ${Login.isAuthenticated() ? 'status-connected' : 'status-disconnected'}"></span>
            ${Login.isAuthenticated() ? 'Conectado a Google Cloud' : 'Sesión no iniciada'}
          </div>
          <div class="text-xs color-subtext mt-xs">Licencia: ${user.tenant}</div>
        </div>
        <button class="btn btn-secondary text-xs" id="btn-home-auth">
          ${Login.isAuthenticated() ? 'Gestionar Cuenta' : 'Iniciar Sesión'}
        </button>
      </div>

      <!-- Quick Actions Grid -->
      <div class="mb-base">
        <h2 class="text-sm font-semibold mb-xs">Acciones Rápidas</h2>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: var(--spacing-sm);">
          <div class="card p-sm cursor-pointer hover-card" data-page="editor">
            <div class="text-sm font-bold color-primary mb-xs">SQL Console</div>
            <p class="text-xs color-subtext">Escribir y ejecutar consultas BigQuery personalizadas.</p>
          </div>
          <div class="card p-sm cursor-pointer hover-card" data-page="projects">
            <div class="text-sm font-bold color-primary mb-xs">Explorador</div>
            <p class="text-xs color-subtext">Navegar por Proyectos, Datasets y Tablas.</p>
          </div>
          <div class="card p-sm cursor-pointer hover-card" data-page="scheduled">
            <div class="text-sm font-bold color-primary mb-xs">Automatizaciones</div>
            <p class="text-xs color-subtext">Ver refrescos programados en Excel.</p>
          </div>
          <div class="card p-sm cursor-pointer hover-card" data-page="providers">
            <div class="text-sm font-bold color-primary mb-xs">Proveedores</div>
            <p class="text-xs color-subtext">Cambiar entre BigQuery, Fabric, SAP, etc.</p>
          </div>
        </div>
      </div>

      <!-- Active Context Summary -->
      <div class="card p-base">
        <h2 class="text-sm font-semibold mb-xs">Contexto Activo</h2>
        <div class="text-xs flex flex-col gap-xs">
          <div class="flex justify-between border-b pb-xs">
            <span class="color-subtext">Proyecto:</span>
            <span class="font-semibold">${state.selectedProject || 'No seleccionado'}</span>
          </div>
          <div class="flex justify-between border-b pb-xs">
            <span class="color-subtext">Dataset:</span>
            <span class="font-semibold">${state.selectedDataset || 'No seleccionado'}</span>
          </div>
          <div class="flex justify-between">
            <span class="color-subtext">Tabla:</span>
            <span class="font-semibold">${state.selectedTable || 'No seleccionada'}</span>
          </div>
        </div>
      </div>
    </div>
  `;
}