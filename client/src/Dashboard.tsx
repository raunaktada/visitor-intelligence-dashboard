import React, { useState, useEffect, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import './Dashboard.css';
import { classifyCompany, TabId } from './classify';

interface Company {
  name: string;
  domain: string;
  users: number;
  sessions: number;
  views: number;
}

const MONTHS = [
  { key: 'feb', label: 'Feb 2026', file: '/data/feb-2026.csv' },
  { key: 'mar', label: 'Mar 2026', file: '/data/mar-2026.csv' },
  { key: 'apr', label: 'Apr 2026', file: '/data/apr-2026.csv' },
  { key: 'may', label: 'May 2026', file: '/data/may-2026.csv' },
  { key: 'jun', label: 'Jun 2026', file: '/data/jun-2026.csv' },
  { key: 'jul', label: 'Jul 2026', file: '/data/jul-2026.csv' },
];

const TABS: { id: TabId; label: string }[] = [
  { id: 'customers',  label: 'Customers' },
  { id: 'industrial', label: 'Industrial Manufacturers' },
  { id: 'healthcare', label: 'Healthcare / Medtech / Hospitals' },
  { id: 'cpg',        label: 'CPG' },
  { id: 'other',      label: 'Other Manufacturers' },
  { id: 'under1b',    label: 'Under $1B' },
];

function parseCSV(text: string): Company[] {
  return text
    .trim()
    .split('\n')
    .slice(1)
    .flatMap(line => {
      const m = line.match(/^"([^"]*)","([^"]*)","([^"]*)","([^"]*)","([^"]*)"$/);
      if (!m) return [];
      return [{ name: m[1], domain: m[2], users: +m[3], sessions: +m[4], views: +m[5] }];
    });
}

function deduplicateByDomain(companies: Company[]): Company[] {
  const map = new Map<string, Company>();
  for (const c of companies) {
    const key = c.domain.toLowerCase();
    const existing = map.get(key);
    if (existing) {
      map.set(key, {
        ...existing,
        users:    existing.users    + c.users,
        sessions: existing.sessions + c.sessions,
        views:    existing.views    + c.views,
      });
    } else {
      map.set(key, { ...c });
    }
  }
  return Array.from(map.values()).sort((a, b) => b.users - a.users);
}

export default function Dashboard() {
  const [monthData,     setMonthData]     = useState<Record<string, Company[]>>({});
  const [loading,       setLoading]       = useState(true);
  const [selectedMonth, setSelectedMonth] = useState<string>('all');
  const [activeTab,     setActiveTab]     = useState<TabId>('industrial');
  const [viewMode,      setViewMode]      = useState<'companies' | 'individuals'>('companies');
  const [search,        setSearch]        = useState('');
  const [toast,         setToast]         = useState('');

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  }, []);

  useEffect(() => {
    Promise.all(
      MONTHS.map(async m => {
        const res  = await fetch(m.file);
        const text = await res.text();
        return { key: m.key, companies: parseCSV(text) };
      })
    ).then(results => {
      const data: Record<string, Company[]> = {};
      for (const r of results) data[r.key] = r.companies;
      setMonthData(data);
      setLoading(false);
    });
  }, []);

  const activeCompanies: Company[] = useMemo(() => {
    if (selectedMonth === 'all') {
      return deduplicateByDomain(Object.values(monthData).flat());
    }
    return [...(monthData[selectedMonth] ?? [])].sort((a, b) => b.users - a.users);
  }, [selectedMonth, monthData]);

  const tabCounts = useMemo(() => {
    const counts: Record<TabId, number> = {
      customers: 0, industrial: 0, healthcare: 0, cpg: 0, other: 0, under1b: 0,
    };
    for (const c of activeCompanies) {
      counts[classifyCompany(c.name, c.domain)]++;
    }
    return counts;
  }, [activeCompanies]);

  const tabCompanies: Company[] = useMemo(() => {
    if (activeTab === 'under1b') return [];
    return activeCompanies.filter(c => classifyCompany(c.name, c.domain) === activeTab);
  }, [activeCompanies, activeTab]);

  const filtered: Company[] = useMemo(() => {
    if (!search) return tabCompanies;
    const q = search.toLowerCase();
    return tabCompanies.filter(
      c => c.name.toLowerCase().includes(q) || c.domain.toLowerCase().includes(q)
    );
  }, [tabCompanies, search]);

  const currentMonthLabel = selectedMonth === 'all'
    ? 'All_Time'
    : (MONTHS.find(m => m.key === selectedMonth)?.label ?? selectedMonth);

  const currentTabLabel = TABS.find(t => t.id === activeTab)?.label ?? activeTab;

  const exportTab = () => {
    if (!filtered.length) return;
    const rows = filtered.map(c => ({
      'Company Name': c.name,
      Domain:         c.domain,
      Users:          c.users,
      Sessions:       c.sessions,
      Views:          c.views,
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), currentTabLabel.substring(0, 31));
    XLSX.writeFile(wb, `TADA_Visitors_${currentMonthLabel}_${currentTabLabel}.xlsx`);
    showToast(`Exported ${rows.length} companies`);
  };

  const exportAll = () => {
    const wb = XLSX.utils.book_new();
    let total = 0;
    for (const tab of TABS) {
      if (tab.id === 'under1b') continue;
      const rows = activeCompanies
        .filter(c => classifyCompany(c.name, c.domain) === tab.id)
        .map(c => ({
          'Company Name': c.name,
          Domain:         c.domain,
          Users:          c.users,
          Sessions:       c.sessions,
          Views:          c.views,
        }));
      if (!rows.length) continue;
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), tab.label.substring(0, 31));
      total += rows.length;
    }
    if (!total) return;
    XLSX.writeFile(wb, `TADA_Visitors_${currentMonthLabel}_All.xlsx`);
    showToast(`Exported ${total} companies across all tabs`);
  };

  const isPlaceholder =
    activeTab === 'under1b' ||
    (activeTab === 'customers' && tabCompanies.length === 0) ||
    viewMode === 'individuals';

  return (
    <div className="dashboard">
      {toast && <div className="toast">{toast}</div>}

      <header className="dashboard-header">
        <div className="header-left">
          <h1>Website Visitors</h1>
          <span className="brand-badge">TADA</span>
        </div>
        <div className="header-right">
          <button
            className="btn btn-secondary"
            onClick={exportTab}
            disabled={!filtered.length || isPlaceholder}
          >
            ↓ Export Tab
          </button>
          <button
            className="btn btn-secondary"
            onClick={exportAll}
            disabled={loading}
          >
            ↓ Export All
          </button>
        </div>
      </header>

      {/* Month selector */}
      <div className="month-bar">
        <button
          className={`month-btn${selectedMonth === 'all' ? ' month-btn-active' : ''}`}
          onClick={() => setSelectedMonth('all')}
        >
          All Time
        </button>
        {MONTHS.map(m => (
          <button
            key={m.key}
            className={`month-btn${selectedMonth === m.key ? ' month-btn-active' : ''}`}
            onClick={() => setSelectedMonth(m.key)}
          >
            {m.label}
          </button>
        ))}
      </div>

      {/* Industry tabs */}
      <div className="tabs">
        {TABS.map(tab => (
          <button
            key={tab.id}
            className={`tab${activeTab === tab.id ? ' tab-active' : ''}`}
            onClick={() => { setActiveTab(tab.id); setSearch(''); setViewMode('companies'); }}
          >
            {tab.label}
            <span className="tab-count">{tabCounts[tab.id]}</span>
          </button>
        ))}
      </div>

      {/* Companies / Individuals toggle + search */}
      <div className="view-toggle-bar">
        <div className="view-toggle">
          <button
            className={`toggle-btn${viewMode === 'companies' ? ' toggle-active' : ''}`}
            onClick={() => setViewMode('companies')}
          >
            Companies
          </button>
          <button
            className={`toggle-btn${viewMode === 'individuals' ? ' toggle-active' : ''}`}
            onClick={() => setViewMode('individuals')}
          >
            Individuals
          </button>
        </div>
        <div className="filters-right">
          <input
            className="search-input"
            placeholder="Search companies…"
            value={search}
            onChange={e => setSearch(e.target.value)}
            disabled={isPlaceholder}
          />
          {!isPlaceholder && (
            <span className="result-count">{filtered.length} companies</span>
          )}
        </div>
      </div>

      {/* Content */}
      {loading ? (
        <div className="loading">Loading visitor data…</div>
      ) : activeTab === 'under1b' ? (
        <div className="placeholder-state">
          <div className="placeholder-icon">📊</div>
          <h3>Under $1B Revenue</h3>
          <p>This tab will show companies with under $1B annual revenue once revenue data is available from Clay.</p>
        </div>
      ) : activeTab === 'customers' && tabCompanies.length === 0 ? (
        <div className="placeholder-state">
          <div className="placeholder-icon">🤝</div>
          <h3>Customers</h3>
          <p>Add TADA customer domains to the <code>CUSTOMER_DOMAINS</code> set in <code>classify.ts</code> to populate this tab.</p>
        </div>
      ) : viewMode === 'individuals' ? (
        <div className="placeholder-state">
          <div className="placeholder-icon">👥</div>
          <h3>Individual Visitors</h3>
          <p>Individual visitor data from Clay will appear here once the Clay integration is connected.</p>
        </div>
      ) : (
        <div className="table-wrapper">
          <table className="visitors-table">
            <thead>
              <tr>
                <th>Company</th>
                <th>Domain</th>
                <th className="num-col">Users</th>
                <th className="num-col">Sessions</th>
                <th className="num-col">Views</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={5} className="empty-state">No companies match your search</td></tr>
              ) : filtered.map((c, i) => (
                <tr key={i}>
                  <td>{c.name}</td>
                  <td>
                    <a
                      href={`https://${c.domain}`}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="domain-link"
                    >
                      {c.domain}
                    </a>
                  </td>
                  <td className="num-col">{c.users}</td>
                  <td className="num-col">{c.sessions}</td>
                  <td className="num-col">{c.views}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
