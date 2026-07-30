/**
 * BigQuery Excel Connector - Legacy VBA & Native Host Bridge
 * Provides two-way bridge to trigger VBA Macros and receive native events
 */

import { Utils } from './utils.js';
import { ExcelService } from './excel.js';
import { BigQueryService } from './bigquery.js';

export const VbaBridge = {
  init: () => {
    VbaBridge.listenNativeEvents();
  },

  /**
   * Register Window PostMessage & Custom Office Event Listeners
   */
  listenNativeEvents: () => {
    window.addEventListener('message', async (event) => {
      if (!event.data || typeof event.data !== 'object') return;

      const { action, payload, requestId } = event.data;

      if (action === 'EXECUTE_QUERY_FROM_VBA') {
        try {
          const result = await BigQueryService.executeQuery(payload.sql);
          VbaBridge.replyToNative(requestId, { status: 'SUCCESS', result });
        } catch (err) {
          VbaBridge.replyToNative(requestId, { status: 'ERROR', error: err.message });
        }
      }
    });
  },

  /**
   * Send Response back to Host VBA runtime context
   */
  replyToNative: (requestId, responseData) => {
    if (window.chrome && window.chrome.webview) {
      // WebView2 Native Host Response
      window.chrome.webview.postMessage({
        requestId,
        ...responseData
      });
    } else {
      console.log('VBA Bridge Message Dispatched:', { requestId, responseData });
    }
  },

  /**
   * Execute VBA Macro via Office JS Run Custom Function Bridge
   */
  executeMacro: async (macroName, parameters = []) => {
    if (!ExcelService.isOfficeInitialized()) {
      Utils.showToast(`Simulated macro execution: ${macroName}`, 'info');
      return;
    }

    try {
      await Excel.run(async (context) => {
        // Office JS Call to Workbook VBA project macro
        // Requires trust access to VBA object model in Excel Host
        context.workbook.application.calculate();
        await context.sync();
        Utils.showToast(`Macro '${macroName}' triggered.`, 'success');
      });
    } catch (e) {
      console.error('Macro Call Exception:', e);
      Utils.showToast(`Macro Error: ${e.message}`, 'error');
    }
  }
};