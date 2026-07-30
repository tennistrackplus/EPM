/**
 * BigQuery Excel Connector - Enterprise Providers Selection Page
 */

import { Providers } from '../providers.js';
import { Store } from '../state.js';
import { Utils } from '../utils.js';

export async function renderProvidersPage() {
  const currentProvider = Store.getState().activeProvider;

  setTimeout(() => {
    Utils.$$('.provider-card').forEach(card => {
      card.addEventListener('click', () => {
        const providerId = card.getAttribute('data-provider-id');
        Store.setState({ activeProvider: providerId });
        Utils.showToast(`Proveedor cambiado a: ${Providers.getById(providerId).name}`, 'success');
        renderProvidersPage().then(html => {
          const mount = document.getElementById('page-mount');
          if (mount) mount.innerHTML = html;
        });
      });
    });
  }, 0);

  const providerListHtml = Providers.list.map(p => {
    const isSelected = p.id === currentProvider;
    return `
      <div class="card p-base mb-sm cursor-pointer provider-card ${isSelected ? 'border-primary' : ''}" data-provider-id="${p.id}">
        <div class="flex items-center gap-sm mb-xs">
          <div style="width: 32px; height: 32px; display: flex; align-items: center; justify-content: center;">
            ${p.icon}
          </div>
          <div class="flex-1">
            <div class="text-sm font-bold flex items-center justify-between">
              ${p.name}
              ${isSelected ? '<span class="badge badge-primary">Activo</span>' : ''}
            </div>
            <div class="text-xs color-subtext">${p.type}</div>
          </div>
        </div>
        <p class="text-xs color-subtext">${p.description}</p>
      </div>
    `;
  }).join('');

  return `
    <div class="page-container p-base fade-in">
      <div class="mb-base">
        <h1 class="text-xl font-bold">Proveedores Enterprise</h1>
        <p class="text-xs color-subtext">Selecciona la fuente de datos predeterminada para tu libro de Excel</p>
      </div>

      <div class="providers-container">
        ${providerListHtml}
      </div>
    </div>
  `;
}