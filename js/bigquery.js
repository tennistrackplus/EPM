/**
 * BigQuery Excel Connector - BigQuery Service API Controller
 * Simulates REST API endpoints for Cloud Resource Manager & BigQuery Core API
 */

import { Store } from './state.js';
import { Utils } from './utils.js';

export const BigQueryService = {
  /**
   * Fetch GCP Projects available for active user
   */
  getProjects: async () => {
    await new Promise(r => setTimeout(r, 400));
    return [
      { id: 'bq-analytics-prod', name: 'Analytics Production (bq-analytics-prod)', role: 'Owner' },
      { id: 'bq-finance-reporting', name: 'Finance & ERP (bq-finance-reporting)', role: 'Viewer' },
      { id: 'bq-marketing-data', name: 'Marketing Insights (bq-marketing-data)', role: 'Editor' },
      { id: 'bigquery-public-data', name: 'Google Public Datasets', role: 'Viewer' }
    ];
  },

  /**
   * Fetch Datasets inside a Project
   */
  getDatasets: async (projectId) => {
    await new Promise(r => setTimeout(r, 350));
    if (projectId === 'bigquery-public-data') {
      return [
        { id: 'usa_names', location: 'US' },
        { id: 'covid19_open_data', location: 'US' },
        { id: 'ga4_obfuscated_sample_ecommerce', location: 'US' }
      ];
    }
    return [
      { id: 'sales_marts', location: 'EU' },
      { id: 'financial_ledger', location: 'EU' },
      { id: 'customer_360', location: 'EU' }
    ];
  },

  /**
   * Fetch Tables inside a Dataset
   */
  getTables: async (projectId, datasetId) => {
    await new Promise(r => setTimeout(r, 300));
    return [
      { id: 'fact_orders_v2', type: 'TABLE', numRows: 1450200, sizeBytes: 524288000 },
      { id: 'dim_customers', type: 'TABLE', numRows: 85000, sizeBytes: 12582912 },
      { id: 'v_monthly_revenue', type: 'VIEW', numRows: 0, sizeBytes: 0 },
      { id: 'historical_archive', type: 'EXTERNAL', numRows: 8900100, sizeBytes: 21474836480 }
    ];
  },

  /**
   * Dry Run Query to estimate cost and byte consumption
   */
  dryRunQuery: async (sql) => {
    await new Promise(r => setTimeout(r, 200));
    // Estimate bytes based on string length & simulated complexity
    const estimatedBytes = Math.max(1048576, sql.length * 1048576 * 12);
    const terabytes = estimatedBytes / (1024 * 1024 * 1024 * 1024);
    const estimatedCostUsd = (terabytes * 6.25).toFixed(4); // $6.25 per TB scan

    return {
      valid: true,
      bytesProcessed: estimatedBytes,
      formattedBytes: Utils.formatBytes(estimatedBytes),
      estimatedCostUsd: `$${estimatedCostUsd}`
    };
  },

  /**
   * Execute Query SQL and return Structured Records
   */
  executeQuery: async (sql) => {
    Store.setState({ activeQuery: { ...Store.getState().activeQuery, isExecuting: true } });

    await new Promise(r => setTimeout(r, 1200));

    // Dummy Mock Data Generation matching SQL results
    const columns = ['order_id', 'customer_name', 'region', 'total_amount', 'transaction_date'];
    const rows = [];
    const regions = ['EMEA', 'NORTH_AMERICA', 'LATAM', 'APAC'];

    for (let i = 1; i <= 50; i++) {
      rows.push([
        `ORD-${1000 + i}`,
        `Client Alpha ${i}`,
        regions[i % regions.length],
        (Math.random() * 2500 + 50).toFixed(2),
        new Date(2026, 0, (i % 28) + 1).toISOString().split('T')[0]
      ]);
    }

    const result = {
      columns,
      rows,
      totalRows: rows.length,
      executionTimeMs: 1140,
      bytesProcessed: 52428800
    };

    Store.setState({
      activeQuery: {
        sql,
        isExecuting: false,
        lastResult: result,
        bytesProcessed: result.bytesProcessed,
        estimatedCost: 0.0003
      }
    });

    return result;
  }
};