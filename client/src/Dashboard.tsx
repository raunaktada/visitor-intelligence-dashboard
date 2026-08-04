import React, { useState, useEffect, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import './Dashboard.css';
import { classifyCompany, TabId, CUSTOMER_CONTACTS, CUSTOMER_DOMAINS } from './classify';

interface Company {
  name: string;
  domain: string;
  users: number;
  sessions: number;
  views: number;
}

interface Individual {
  firstName: string;
  lastName: string;
  email: string;
  company: string;
  title: string;
  website: string;
  pageViews: number;
  city: string;
  state: string;
  linkedin: string;
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

function parseIndividuals(text: string): Individual[] {
  const lines = text.trim().split('\n');
  // header: first_name,last_name,email,company_name,title,website,last_seen_at,page_views,city,state,linkedin_url
  return lines.slice(1).flatMap(line => {
    // Split on commas but respect quoted fields
    const fields: string[] = [];
    let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQ = !inQ; }
      else if (ch === ',' && !inQ) { fields.push(cur); cur = ''; }
      else { cur += ch; }
    }
    fields.push(cur);
    if (fields.length < 11) return [];
    return [{
      firstName: fields[0].trim(),
      lastName:  fields[1].trim(),
      email:     fields[2].trim(),
      company:   fields[3].trim(),
      title:     fields[4].trim(),
      website:   fields[5].trim(),
      pageViews: parseInt(fields[7]) || 0,
      city:      fields[8].trim(),
      state:     fields[9].trim(),
      linkedin:  fields[10].trim(),
    }];
  });
}

function domainFromUrl(url: string): string {
  try {
    return new URL(url.startsWith('http') ? url : `https://${url}`).hostname.replace(/^www\./, '');
  } catch {
    return url.toLowerCase().replace(/^www\./, '');
  }
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
  const [individuals,   setIndividuals]   = useState<Individual[]>([]);
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
    Promise.all([
      ...MONTHS.map(async m => {
        const res  = await fetch(m.file);
        const text = await res.text();
        return { key: m.key, companies: parseCSV(text) };
      }),
      fetch('/data/individuals-all.csv').then(r => r.text()).then(text => {
        setIndividuals(parseIndividuals(text));
        return null;
      }),
    ]).then(results => {
      const data: Record<string, Company[]> = {};
      for (const r of results) {
        if (r) data[r.key] = r.companies;
      }
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

  // Map domain → individuals for inline display in company rows
  const individualsByDomain = useMemo(() => {
    const map = new Map<string, Individual[]>();
    for (const p of individuals) {
      const d = domainFromUrl(p.website);
      if (!d) continue;
      if (!map.has(d)) map.set(d, []);
      map.get(d)!.push(p);
    }
    return map;
  }, [individuals]);

  // Individuals — always all-time (Artisan has no monthly breakdown)
  const tabIndividuals: Individual[] = useMemo(() => {
    if (activeTab === 'under1b') return [];
    return individuals.filter(p => {
      const domain = domainFromUrl(p.website);
      return classifyCompany(p.company, domain) === activeTab;
    });
  }, [individuals, activeTab]);

  const filteredIndividuals: Individual[] = useMemo(() => {
    if (!search) return tabIndividuals;
    const q = search.toLowerCase();
    return tabIndividuals.filter(p =>
      p.firstName.toLowerCase().includes(q) ||
      p.lastName.toLowerCase().includes(q) ||
      p.company.toLowerCase().includes(q) ||
      p.title.toLowerCase().includes(q)
    );
  }, [tabIndividuals, search]);

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
    (activeTab === 'customers' && tabCompanies.length === 0 && tabIndividuals.length === 0);

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
            placeholder={viewMode === 'individuals' ? 'Search individuals…' : 'Search companies…'}
            value={search}
            onChange={e => setSearch(e.target.value)}
            disabled={isPlaceholder}
          />
          {!isPlaceholder && (
            <span className="result-count">
              {viewMode === 'individuals'
                ? `${filteredIndividuals.length} individuals (all time)`
                : `${filtered.length} companies`}
            </span>
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
        <div className="table-wrapper">
          <table className="visitors-table">
            <thead>
              <tr>
                <th>Name</th>
                <th>Title</th>
                <th>Company</th>
                <th>Location</th>
                <th>Page Views</th>
                <th>LinkedIn</th>
              </tr>
            </thead>
            <tbody>
              {filteredIndividuals.length === 0 ? (
                <tr><td colSpan={6} className="empty-state">No individuals match your search</td></tr>
              ) : filteredIndividuals.map((p, i) => (
                <tr key={i}>
                  <td>{p.firstName} {p.lastName}</td>
                  <td className="secondary-text">{p.title}</td>
                  <td>{p.company}</td>
                  <td className="secondary-text">{[p.city, p.state].filter(Boolean).join(', ')}</td>
                  <td className="num-col">{p.pageViews}</td>
                  <td>
                    {p.linkedin && (
                      <a href={p.linkedin} target="_blank" rel="noopener noreferrer" className="linkedin-link">
                        View ↗
                      </a>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
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
              ) : filtered.map((c, i) => {
                // SharePoint contacts (Customers tab only)
                const spContacts = CUSTOMER_DOMAINS.has(c.domain) ? (CUSTOMER_CONTACTS[c.domain] ?? []) : [];
                // Artisan individuals matched by domain
                const artisanPeople = individualsByDomain.get(c.domain) ?? [];
                // Artisan names already in SP contacts — deduplicate
                const spNames = new Set(spContacts.map(p => p.name.toLowerCase()));
                const extraPeople = artisanPeople.filter(p =>
                  !spNames.has(`${p.firstName} ${p.lastName}`.toLowerCase())
                );
                return (
                  <tr key={i}>
                    <td>
                      <div>{c.name}</div>
                      {(spContacts.length > 0 || extraPeople.length > 0) && (
                        <div className="contact-list">
                          {spContacts.map((p, j) => (
                            <span key={`sp-${j}`} className="contact-chip">
                              {p.linkedin
                                ? <a href={p.linkedin} target="_blank" rel="noopener noreferrer" className="contact-link">{p.name}</a>
                                : p.name}
                              <span className="contact-title">{p.title}</span>
                            </span>
                          ))}
                          {extraPeople.map((p, j) => (
                            <span key={`art-${j}`} className="contact-chip">
                              {p.linkedin
                                ? <a href={p.linkedin} target="_blank" rel="noopener noreferrer" className="contact-link">{p.firstName} {p.lastName}</a>
                                : `${p.firstName} ${p.lastName}`}
                              <span className="contact-title">{p.title}</span>
                            </span>
                          ))}
                        </div>
                      )}
                    </td>
                    <td>
                      <a href={`https://${c.domain}`} target="_blank" rel="noopener noreferrer" className="domain-link">
                        {c.domain}
                      </a>
                    </td>
                    <td className="num-col">{c.users}</td>
                    <td className="num-col">{c.sessions}</td>
                    <td className="num-col">{c.views}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
