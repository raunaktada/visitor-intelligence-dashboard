import React, { useState, useEffect, useMemo, useCallback } from 'react';
import './Dashboard.css';

type Row = Record<string, unknown>;
type SummaryMatrix = unknown[][];

const SHEETS = [
  'Summary',
  'Customers & Partners',
  'Data Centers',
  'Defense',
  'Manufacturing',
  'Healthcare & MedTech',
  'CPG',
];

const API = '';
const POLL_INTERVAL_MS = 60 * 1000; // re-check every minute; server cache is 5 min

export default function Dashboard() {
  const [allData, setAllData]     = useState<Record<string, Row[] | SummaryMatrix>>({});
  const [loading, setLoading]     = useState(true);
  const [syncing, setSyncing]     = useState(false);
  const [activeTab, setActiveTab] = useState(SHEETS[0]);
  const [search, setSearch]       = useState('');
  const [toast, setToast]         = useState('');
  const [lastUpdated, setLastUpdated] = useState<Date | null>(null);

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  }, []);

  const fetchData = useCallback(async (silent = false) => {
    if (!silent) setLoading(true);
    try {
      const res = await fetch(`${API}/api/visitors`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setAllData(data);
      setLastUpdated(new Date());
    } catch {
      if (!silent) showToast('Failed to fetch data');
    } finally {
      if (!silent) setLoading(false);
    }
  }, [showToast]);

  // Initial load + background polling so SharePoint edits show up automatically
  useEffect(() => {
    fetchData();
    const interval = setInterval(() => fetchData(true), POLL_INTERVAL_MS);
    return () => clearInterval(interval);
  }, [fetchData]);

  const handleSync = async () => {
    setSyncing(true);
    try {
      const res = await fetch(`${API}/api/sync`, { method: 'POST' });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || err.error || `HTTP ${res.status}`);
      }
      const data = await res.json();
      showToast(`Synced ${data.recordsSynced} records across ${data.sheets?.length ?? 0} sheets`);
      await fetchData(true);
    } catch (err) {
      showToast(`Sync failed: ${err instanceof Error ? err.message : 'unknown error'}`);
    } finally {
      setSyncing(false);
    }
  };

  const isSummary = activeTab === 'Summary';
  const dataSheets = SHEETS.filter(s => s !== 'Summary');

  const rows: Row[] = isSummary ? [] : ((allData[activeTab] as Row[]) || []);
  const summaryRows: SummaryMatrix = isSummary ? ((allData['Summary'] as SummaryMatrix) || []) : [];

  const headers = useMemo(() => {
    if (!rows.length) return [];
    return Object.keys(rows[0]).filter(h => h.trim());
  }, [rows]);

  const filtered = useMemo(() => {
    if (!search) return rows;
    const q = search.toLowerCase();
    return rows.filter(row =>
      Object.values(row).some(v => String(v).toLowerCase().includes(q))
    );
  }, [rows, search]);

  const exportCSV = () => {
    if (!filtered.length) return;
    const csvRows = [
      headers.join(','),
      ...filtered.map(row =>
        headers.map(h => `"${String(row[h] ?? '').replace(/"/g, '""')}"`).join(',')
      ),
    ];
    const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `${activeTab.replace(/[^a-z0-9]/gi, '_')}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    showToast(`Exported ${filtered.length} rows`);
  };

  const exportAll = () => {
    const sections: string[] = [];
    let totalRowsExported = 0;

    for (const sheet of dataSheets) {
      const sheetRows = (allData[sheet] as Row[]) || [];
      if (!sheetRows.length) continue;

      const sheetHeaders = Object.keys(sheetRows[0]).filter(h => h.trim());
      const csvLines = [
        sheetHeaders.join(','),
        ...sheetRows.map(row =>
          sheetHeaders.map(h => `"${String(row[h] ?? '').replace(/"/g, '""')}"`).join(',')
        ),
      ];
      sections.push(`## ${sheet}`, ...csvLines, '');
      totalRowsExported += sheetRows.length;
    }

    if (!sections.length) return;

    const blob = new Blob([sections.join('\n')], { type: 'text/csv' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = 'website_visitors_all_sheets.csv';
    a.click();
    URL.revokeObjectURL(url);
    showToast(`Exported ${totalRowsExported} rows across ${dataSheets.length} sheets`);
  };

  const totalRows = rows.length;
  const linkedInKey = headers.find(h => h.toLowerCase().includes('linkedin'));

  // Map Excel tab labels (Summary sheet wording) to our actual tab keys
  const tabAliases: Record<string, string> = {
    'Current Customers or Partners': 'Customers & Partners',
    'Data Center BuildOut Sector':   'Data Centers',
    'Defense Sector':                'Defense',
    'Manufacturing':                 'Manufacturing',
    'Healthcare & MedTech':          'Healthcare & MedTech',
    'CPG':                           'CPG',
  };

  return (
    <div className="dashboard">
      {toast && <div className="toast">{toast}</div>}

      <header className="dashboard-header">
        <div className="header-left">
          <h1>Website Visitors</h1>
          <span className="brand-badge">TADA</span>
        </div>
        <div className="header-right">
          {lastUpdated && (
            <span className="last-updated">Updated {lastUpdated.toLocaleTimeString()} (your local time)</span>
          )}
          <button className="btn btn-primary" onClick={handleSync} disabled={syncing}>
            {syncing ? 'Syncing…' : '⟳ Sync'}
          </button>
          <button className="btn btn-secondary" onClick={exportCSV} disabled={!filtered.length || isSummary}>
            ↓ Export Tab
          </button>
          <button className="btn btn-secondary" onClick={exportAll}>
            ↓ Export All
          </button>
        </div>
      </header>

      {/* Sheet Tabs */}
      <div className="tabs">
        {SHEETS.map(sheet => (
          <button
            key={sheet}
            className={`tab ${activeTab === sheet ? 'tab-active' : ''}`}
            onClick={() => { setActiveTab(sheet); setSearch(''); }}
          >
            {sheet}
            {sheet !== 'Summary' && (
              <span className="tab-count">{(allData[sheet] as Row[])?.length ?? 0}</span>
            )}
          </button>
        ))}
      </div>

      {isSummary ? (
        loading ? (
          <div className="loading">Loading data from SharePoint…</div>
        ) : (
          <div className="excel-summary">
            {summaryRows.map((row, i) => {
              const colA = String(row[0] ?? '').trim();
              const colB = String(row[1] ?? '').trim();

              if (!colA && !colB) return <div key={i} className="excel-summary-spacer" />;

              if (colA && !colB) {
                // Title / info lines (rows 0-1) come from column A
                const isTitle = i === 0;
                return (
                  <div key={i} className={isTitle ? 'excel-summary-title' : 'excel-summary-info'}>
                    {colA}
                  </div>
                );
              }

              if (colB && colB !== 'TAB') {
                const mapped = tabAliases[colB];
                return (
                  <div
                    key={i}
                    className={`excel-summary-row ${mapped ? 'clickable' : ''}`}
                    onClick={() => mapped && setActiveTab(mapped)}
                  >
                    <span>{mapped || colB}</span>
                    {mapped && (
                      <span className="excel-summary-count">
                        {(allData[mapped] as Row[])?.length ?? 0}
                      </span>
                    )}
                  </div>
                );
              }

              if (colA === ' TAB' || colA === 'TAB') {
                return <div key={i} className="excel-summary-heading">{colA.trim()}</div>;
              }

              return null;
            })}
          </div>
        )
      ) : (
        <>
          {/* Search */}
          <div className="filters-bar">
            <input
              className="search-input"
              placeholder={`Search in ${activeTab}…`}
              value={search}
              onChange={e => setSearch(e.target.value)}
            />
            <span className="result-count">{filtered.length} of {totalRows} records</span>
          </div>

          {/* Table */}
          {loading ? (
            <div className="loading">Loading data from SharePoint…</div>
          ) : rows.length === 0 ? (
            <div className="loading">No data in this sheet.</div>
          ) : (
            <div className="table-wrapper">
              <table className="visitors-table">
                <thead>
                  <tr>
                    {headers.map(h => <th key={h}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {filtered.length === 0 ? (
                    <tr><td colSpan={headers.length} className="empty-state">No records match your search</td></tr>
                  ) : filtered.map((row, i) => (
                    <tr key={i}>
                      {headers.map(h => {
                        const value = row[h];
                        if (h === linkedInKey && typeof value === 'string' && value.trim()) {
                          const url = value.startsWith('http') ? value : `https://${value}`;
                          return (
                            <td key={h}>
                              <a href={url} target="_blank" rel="noopener noreferrer" className="linkedin-link">
                                View profile ↗
                              </a>
                            </td>
                          );
                        }
                        return <td key={h}>{String(value ?? '—')}</td>;
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </div>
  );
}
