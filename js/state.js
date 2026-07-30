/**
 * BigQuery Excel Connector - Reactive State Management Store
 * Holds global application state and notifies subscribers on mutation.
 */

import { Utils } from './utils.js';

class StateStore {
  constructor() {
    this.listeners = new Set();
    
    // Initial State Structure
    this.state = {
      activeProvider: 'bigquery',
      selectedProject: Utils.storage.get('selected_project', null),
      selectedDataset: Utils.storage.get('selected_dataset', null),
      selectedTable: Utils.storage.get('selected_table', null),
      savedQueries: Utils.storage.get('saved_queries', []),
      queryHistory: Utils.storage.get('query_history', []),
      scheduledJobs: Utils.storage.get('scheduled_jobs', []),
      activeQuery: {
        sql: 'SELECT * FROM `bigquery-public-data.usa_names.usa_1910_2013` LIMIT 100;',
        isExecuting: false,
        lastResult: null,
        bytesProcessed: 0,
        estimatedCost: 0
      },
      settings: Utils.storage.get('app_settings', {
        theme: 'light',
        autoFormatTables: true,
        maxRowsDefault: 1000,
        enableVbaBridge: true
      })
    };
  }

  getState() {
    return { ...this.state };
  }

  setState(partialState) {
    this.state = {
      ...this.state,
      ...partialState
    };
    
    // Persist key variables
    if (partialState.selectedProject !== undefined) Utils.storage.set('selected_project', this.state.selectedProject);
    if (partialState.selectedDataset !== undefined) Utils.storage.set('selected_dataset', this.state.selectedDataset);
    if (partialState.selectedTable !== undefined) Utils.storage.set('selected_table', this.state.selectedTable);
    if (partialState.savedQueries !== undefined) Utils.storage.set('saved_queries', this.state.savedQueries);
    if (partialState.queryHistory !== undefined) Utils.storage.set('query_history', this.state.queryHistory);
    if (partialState.scheduledJobs !== undefined) Utils.storage.set('scheduled_jobs', this.state.scheduledJobs);
    if (partialState.settings !== undefined) Utils.storage.set('app_settings', this.state.settings);

    this.notify();
  }

  subscribe(listener) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  notify() {
    this.listeners.forEach(listener => listener(this.state));
  }
}

export const Store = new StateStore();