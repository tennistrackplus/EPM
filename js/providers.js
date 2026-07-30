/**
 * BigQuery Excel Connector - Enterprise Data Providers Registry
 * Specification and metadata for supported enterprise data sources
 */

export const Providers = {
  list: [
    {
      id: 'bigquery',
      name: 'Google BigQuery',
      type: 'Cloud Data Warehouse',
      accentColor: 'var(--color-provider-bigquery)',
      description: 'Serverless, highly scalable enterprise data warehouse by Google Cloud.',
      status: 'active',
      icon: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zM19 18H6c-2.21 0-4-1.79-4-4 0-2.05 1.53-3.76 3.56-3.97l1.07-.11.5-.95C8.08 7.14 9.94 6 12 6c2.62 0 4.88 1.86 5.39 4.43l.3 1.5 1.53.11c1.56.1 2.78 1.41 2.78 2.96 0 1.65-1.35 3-3 3z" fill="#4285F4"/></svg>`
    },
    {
      id: 'fabric',
      name: 'Microsoft Fabric',
      type: 'Unified Data Platform',
      accentColor: 'var(--color-provider-fabric)',
      description: 'All-in-one analytics solution for enterprises covering data engineering and lakehouses.',
      status: 'active',
      icon: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><rect x="3" y="3" width="8" height="8" rx="1" fill="#0078D4"/><rect x="13" y="3" width="8" height="8" rx="1" fill="#0078D4"/><rect x="3" y="13" width="8" height="8" rx="1" fill="#0078D4"/><rect x="13" y="13" width="8" height="8" rx="1" fill="#0078D4"/></svg>`
    },
    {
      id: 'snowflake',
      name: 'Snowflake',
      type: 'Data Cloud Platform',
      accentColor: 'var(--color-provider-snowflake)',
      description: 'Cloud Data Platform built for multi-cloud data warehousing and data lakes.',
      status: 'active',
      icon: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><path d="M12 2v20M2 12h20M4.93 4.93l14.14 14.14M4.93 19.07L19.07 4.93" stroke="#29B5E8" stroke-width="2" stroke-linecap="round"/></svg>`
    },
    {
      id: 'sap_datasphere',
      name: 'SAP Datasphere',
      type: 'Business Data Fabric',
      accentColor: 'var(--color-provider-sap)',
      description: 'Comprehensive data service enabling seamless access to SAP and non-SAP data.',
      status: 'active',
      icon: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><path d="M3 6h18v12H3z" stroke="#008FD3" stroke-width="2"/><path d="M7 10h10M7 14h6" stroke="#008FD3" stroke-width="2"/></svg>`
    },
    {
      id: 'sap_cds',
      name: 'SAP Core Data Services',
      type: 'Semantic Data Layer',
      accentColor: 'var(--color-provider-sap)',
      description: 'Domain-specific language and framework for defining rich semantic data models.',
      status: 'active',
      icon: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="9" stroke="#008FD3" stroke-width="2"/><path d="M12 7v10M7 12h10" stroke="#008FD3" stroke-width="2"/></svg>`
    },
    {
      id: 'redshift',
      name: 'Amazon Redshift',
      type: 'Cloud Data Warehouse',
      accentColor: 'var(--color-provider-redshift)',
      description: 'Fast, simple, cost-effective petabyte-scale data warehousing service by AWS.',
      status: 'active',
      icon: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5" stroke="#CC292B" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`
    }
  ],

  getById: (providerId) => {
    return Providers.list.find(p => p.id === providerId) || null;
  }
};