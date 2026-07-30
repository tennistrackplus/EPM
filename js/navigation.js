/**
 * BigQuery Excel Connector - Navigation & Router Engine
 * SPA Router for taskpane views and shell sidebar controller
 */

import { Utils } from './utils.js';

export const Navigation = {
  currentPage: null,
  pageRegistry: {},

  init: () => {
    Navigation.bindEvents();
  },

  registerPage: (pageId, renderFunction) => {
    Navigation.pageRegistry[pageId] = renderFunction;
  },

  bindEvents: () => {
    // Sidebar toggle handling
    const btnSidebarToggle = Utils.$('#btn-sidebar-toggle');
    const sidebar = Utils.$('#app-sidebar');

    if (btnSidebarToggle && sidebar) {
      btnSidebarToggle.addEventListener('click', () => {
        sidebar.classList.toggle('sidebar-expanded');
        sidebar.classList.toggle('sidebar-collapsed');
      });
    }

    // Sidebar navigation click delegation
    document.addEventListener('click', (e) => {
      const navBtn = e.target.closest('[data-page]');
      if (navBtn) {
        const pageId = navBtn.getAttribute('data-page');
        Navigation.navigateTo(pageId);
      }
    });

    // Profile shortcut click
    const btnProfile = Utils.$('#btn-profile-shortcut');
    if (btnProfile) {
      btnProfile.addEventListener('click', () => {
        Navigation.navigateTo('profile');
      });
    }
  },

  navigateTo: async (pageId, params = {}) => {
    const mountPoint = Utils.$('#page-mount');
    if (!mountPoint) return;

    if (!Navigation.pageRegistry[pageId]) {
      console.warn(`Page '${pageId}' not found in registry. Falling back to 'error404'.`);
      if (Navigation.pageRegistry['error404']) {
        pageId = 'error404';
      } else {
        mountPoint.innerHTML = `<div class="p-base">Error: Page not found (${pageId})</div>`;
        return;
      }
    }

    // Update Sidebar Active state
    Utils.$$('.nav-item').forEach(item => {
      if (item.getAttribute('data-page') === pageId) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });

    // Render Page HTML Content
    Navigation.currentPage = pageId;
    mountPoint.innerHTML = '';
    
    try {
      const pageContent = await Navigation.pageRegistry[pageId](params);
      if (typeof pageContent === 'string') {
        mountPoint.innerHTML = pageContent;
      } else if (pageContent instanceof HTMLElement) {
        mountPoint.appendChild(pageContent);
      }
      
      // Auto-scroll to top on view change
      const stage = Utils.$('.app-stage');
      if (stage) stage.scrollTop = 0;

    } catch (error) {
      console.error(`Failed rendering page ${pageId}:`, error);
      Utils.showToast(`Error loading view: ${error.message}`, 'error');
    }
  }
};