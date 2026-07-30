/**
 * BigQuery Excel Connector - Main Application Engine
 */

import { Navigation } from './navigation.js';
import { Providers } from './providers.js';
import { Login } from './login.js';
import { Dashboard } from './dashboard.js';
import { Settings } from './settings.js';
import { Dialogs } from './dialogs.js';
import { Utils } from './utils.js';

// Estado global en memoria para la aplicación
export const AppState = {
  user: null,
  selectedProvider: 'bigquery',
  selectedProject: null,
  selectedDataset: null,
  selectedTable: null,
  activeQuery: {
    sql: '',
    lastResult: null
  },
  settings: {
    theme: 'light',
    maxRowsDefault: 1000,
    autoFormatTables: true,
    enableVbaBridge: false
  }
};

/**
 * Inicialización principal al cargar Office JS
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log('✅ Conectado nativamente a Microsoft Excel Host.');
  } else {
    console.log('⚠️ Modo Standalone Web (Fuera de Microsoft Excel).');
  }

  initApp();
});

/**
 * Arranca los servicios clave de la SPA
 */
function initApp() {
  try {
    // 1. Inicializar sistema de navegación entre páginas
    Navigation.init();

    // 2. Verificar estado de sesión guardado
    Login.checkSession();

    // 3. Registrar eventos globales en botones e interfaz
    bindGlobalEvents();

    console.log('🚀 Aplicación inicializada correctamente.');
  } catch (error) {
    console.error('❌ Error crítico al inicializar la app:', error);
    Utils.showToast('Error de inicialización de la aplicación', 'error');
  }
}

/**
 * Eventos globales de la interfaz (navegación y cierres de modal)
 */
function bindGlobalEvents() {
  // Navegación mediante atributos data-nav en botones del HTML
  document.addEventListener('click', (e) => {
    const navBtn = e.target.closest('[data-nav]');
    if (navBtn) {
      e.preventDefault();
      const page = navBtn.getAttribute('data-nav');
      Navigation.navigateTo(page);
    }
  });

  // Exponer función de navegación al objeto global window
  window.appNavigate = (page) => Navigation.navigateTo(page);
}