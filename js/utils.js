/**
 * BigQuery Excel Connector - Utility Functions & Helpers
 * Core ES6 JS Utilities for DOM, Storage, Formatting & UI Components
 */

export const Utils = {
  /**
   * Safe DOM selector
   */
  $: (selector, scope = document) => scope.querySelector(selector),
  $$: (selector, scope = document) => Array.from(scope.querySelectorAll(selector)),

  /**
   * Element creation helper
   */
  createElement: (tag, classes = [], attributes = {}, text = '') => {
    const el = document.createElement(tag);
    if (classes.length) el.classList.add(...classes.filter(Boolean));
    Object.entries(attributes).forEach(([key, val]) => el.setAttribute(key, val));
    if (text) el.textContent = text;
    return el;
  },

  /**
   * Local Storage Safe Wrapper with JSON Parsing
   */
  storage: {
    get: (key, defaultValue = null) => {
      try {
        const item = localStorage.getItem(`bq_ec_${key}`);
        return item ? JSON.parse(item) : defaultValue;
      } catch (e) {
        console.error('Storage Read Error:', e);
        return defaultValue;
      }
    },
    set: (key, value) => {
      try {
        localStorage.setItem(`bq_ec_${key}`, JSON.stringify(value));
      } catch (e) {
        console.error('Storage Write Error:', e);
      }
    },
    remove: (key) => {
      try {
        localStorage.removeItem(`bq_ec_${key}`);
      } catch (e) {
        console.error('Storage Remove Error:', e);
      }
    }
  },

  /**
   * UI Toast Notification Engine
   */
  showToast: (message, type = 'info', duration = 3500) => {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = Utils.createElement('div', ['toast', `toast-${type}`]);
    
    let iconSvg = '';
    if (type === 'success') {
      iconSvg = `<svg class="fluent-icon" viewBox="0 0 20 20" width="16" height="16" fill="var(--color-success)"><path d="M10 2a8 8 0 100 16 8 8 0 000-16zm3.707 6.707l-4 4a1 1 0 01-1.414 0l-2-2a1 1 0 111.414-1.414L9 10.586l3.293-3.293a1 1 0 011.414 1.414z"/></svg>`;
    } else if (type === 'error') {
      iconSvg = `<svg class="fluent-icon" viewBox="0 0 20 20" width="16" height="16" fill="var(--color-error)"><path d="M10 2a8 8 0 100 16 8 8 0 000-16zm3.707 5.293a1 1 0 00-1.414-1.414L10 8.586 7.707 6.293a1 1 0 00-1.414 1.414L8.586 10l-2.293 2.293a1 1 0 101.414 1.414L10 11.414l2.293 2.293a1 1 0 001.414-1.414L11.414 10l2.293-2.293z"/></svg>`;
    } else {
      iconSvg = `<svg class="fluent-icon" viewBox="0 0 20 20" width="16" height="16" fill="var(--color-info)"><path d="M10 2a8 8 0 100 16 8 8 0 000-16zm0 4a1 1 0 110 2 1 1 0 010-2zm1 4v4a1 1 0 11-2 0v-4a1 1 0 112 0z"/></svg>`;
    }

    toast.innerHTML = `
      ${iconSvg}
      <span class="text-sm font-medium" style="color: var(--color-text);">${message}</span>
    `;

    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transition = 'opacity 300ms ease';
      setTimeout(() => toast.remove(), 300);
    }, duration);
  },

  /**
   * Global Dynamic Modal Dialog Manager
   */
  showModal: ({ title, bodyHtml, confirmText = 'Confirm', cancelText = 'Cancel', onConfirm }) => {
    const container = document.getElementById('modal-container');
    if (!container) return;

    container.innerHTML = `
      <div class="modal-dialog">
        <div class="modal-header">
          <h2 class="modal-title">${title}</h2>
          <button class="btn-icon modal-close-btn" aria-label="Close">
            <svg class="fluent-icon" viewBox="0 0 20 20" width="16" height="16"><path d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" fill="currentColor"/></svg>
          </button>
        </div>
        <div class="modal-body">${bodyHtml}</div>
        <div class="modal-footer">
          <button class="btn btn-secondary modal-cancel-btn">${cancelText}</button>
          <button class="btn btn-primary modal-confirm-btn">${confirmText}</button>
        </div>
      </div>
    `;

    container.classList.add('active');
    container.setAttribute('aria-hidden', 'false');

    const closeModal = () => {
      container.classList.remove('active');
      container.setAttribute('aria-hidden', 'true');
      container.innerHTML = '';
    };

    container.querySelector('.modal-close-btn').onclick = closeModal;
    container.querySelector('.modal-cancel-btn').onclick = closeModal;
    container.querySelector('.modal-confirm-btn').onclick = async () => {
      if (onConfirm) await onConfirm();
      closeModal();
    };
  },

  /**
   * Data Formatter Utilities
   */
  formatBytes: (bytes, decimals = 2) => {
    if (!+bytes) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${parseFloat((bytes / Math.pow(k, i)).toFixed(dm))} ${sizes[i]}`;
  },

  formatNumber: (num) => {
    return new Intl.NumberFormat().format(num);
  }
};