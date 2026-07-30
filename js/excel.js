/**
 * BigQuery Excel Connector - Office JS Integration Wrapper
 * Interacts directly with Excel Application Host via Office.js API
 */

import { Utils } from './utils.js';

export const ExcelService = {
  /**
   * Check if running inside Office JS host container
   */
  isOfficeInitialized: () => {
    return typeof Office !== 'undefined' && typeof Excel !== 'undefined';
  },

  /**
   * Write matrix data to specified or active Excel Worksheet
   */
  writeToSheet: async (columns, rows, sheetName = 'BigQuery_Data') => {
    if (!ExcelService.isOfficeInitialized()) {
      Utils.showToast('Excel API not available. Simulated execution.', 'info');
      console.log('Office JS Fallback execution:', { columns, rows });
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        let sheet = sheets.getItemOrNullObject(sheetName);
        await context.sync();

        if (sheet.isNullObject) {
          sheet = sheets.add(sheetName);
        }
        sheet.activate();

        // Clear existing content
        sheet.getUsedRange().clear();

        const fullMatrix = [columns, ...rows];
        const rowCount = fullMatrix.length;
        const colCount = columns.length;

        const targetRange = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
        targetRange.values = fullMatrix;

        // Create Excel Styled Table
        const table = sheet.tables.add(targetRange, true);
        table.name = `BQ_Table_${Date.now().toString().slice(-4)}`;
        table.style = 'TableStyleMedium9';

        await context.sync();
        Utils.showToast(`Successfully inserted ${rows.length} rows into sheet '${sheetName}'.`, 'success');
      });
    } catch (error) {
      console.error('Office JS Write Error:', error);
      Utils.showToast(`Excel Write Error: ${error.message}`, 'error');
    }
  },

  /**
   * Reads active selection from current active worksheet
   */
  getSelectedRangeValues: async () => {
    if (!ExcelService.isOfficeInitialized()) return [];

    let values = [];
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('values');
      await context.sync();
      values = range.values;
    });
    return values;
  }
};