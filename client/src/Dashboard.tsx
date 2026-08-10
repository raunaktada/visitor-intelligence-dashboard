import React, { useState, useEffect, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import './Dashboard.css';

const MONTH_KEYS = ['Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'Jun 2026', 'Jul 2026'];

type TabId = 'customers' | 'defense' | 'datacenter' | 'manufacturing' | 'healthcare' | 'cpg' | 'under1b';

interface MonthData { users: number; sessions: number; views: number; }
interface Contact { name: string; title: string; email: string; linkedin: string; pageViewed: number; }
interface Company {
  name: string;
  revenue: string;
  category?: string;
  months: Record<string, MonthData>;
  contacts: Contact[];
}

const TABS: { id: TabId; label: string; sheet: string }[] = [
  { id: 'customers',     label: 'Customers & Partners',       sheet: 'Customers & Partners' },
  { id: 'defense',       label: 'Defense Manufacturers',      sheet: 'Defense Manufacturers - 1B+' },
  { id: 'datacenter',    label: 'Data Center Manufacturers',  sheet: 'Data Center Mfg - 1B+' },
  { id: 'manufacturing', label: 'Manufacturing',              sheet: 'Manufacturing - 1B+' },
  { id: 'healthcare',    label: 'Healthcare / MedTech',       sheet: 'Healthcare_MedTech - 1B+' },
  { id: 'cpg',           label: 'CPG',                        sheet: 'CPG - 1B+' },
  { id: 'under1b',       label: 'Under $1B',                  sheet: 'Under $1B (250M-1B)' },
];

const API_URL = import.meta.env.PROD ? '' : ((import.meta.env.VITE_API_URL as string) || 'http://localhost:5001');

export default function Dashboard() {
  const [allData,        setAllData]        = useState<Record<string, Company[]>>({});
  const [loading,        setLoading]        = useState(true);
  const [error,          setError]          = useState('');
  const [selectedMonth,  setSelectedMonth]  = useState<string>('all');
  const [activeTab,      setActiveTab]      = useState<TabId>('customers');
  const [viewMode,       setViewMode]       = useState<'companies' | 'contacts'>('companies');
  const [search,         setSearch]         = useState('');
  const [toast,          setToast]          = useState('');

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  }, []);

  useEffect(() => {
    fetch(`${API_URL}/api/visitors`)
      .then(r => r.json())
      .then(data => { setAllData(data); setLoading(false); })
      .catch(() => { setError('Failed to load data from server.'); setLoading(false); });
  }, []);

  const activeSheet = TABS.find(t => t.id === activeTab)!.sheet;
  const sheetCompanies: Company[] = useMemo(() => allData[activeSheet] ?? [], [allData, activeSheet]);

  // Compute users/sessions/views for selected month or all-time sum
  const companiesWithTotals = useMemo(() => sheetCompanies.map(c => {
    if (selectedMonth === 'all') {
      const totals = Object.values(c.months).reduce(
        (acc, m) => ({ users: acc.users + m.users, sessions: acc.sessions + m.sessions, views: acc.views + m.views }),
        { users: 0, sessions: 0, views: 0 }
      );
      return { ...c, ...totals };
    }
    const m = c.months[selectedMonth] ?? { users: 0, sessions: 0, views: 0 };
    return { ...c, ...m };
  }).filter(c => selectedMonth === 'all' || c.users > 0 || c.sessions > 0 || c.views > 0)
    .sort((a, b) => b.users - a.users),
  [sheetCompanies, selectedMonth]);

  const filtered = useMemo(() => {
    if (!search) return companiesWithTotals;
    const q = search.toLowerCase();
    return companiesWithTotals.filter(c =>
      c.name.toLowerCase().includes(q) ||
      (c.category ?? '').toLowerCase().includes(q)
    );
  }, [companiesWithTotals, search]);

  // Contacts view — flatten all contacts from filtered companies
  const filteredContacts = useMemo(() => {
    const q = search.toLowerCase();
    return sheetCompanies.flatMap(c =>
      c.contacts.map(p => ({ ...p, company: c.name, revenue: c.revenue }))
    ).filter(p =>
      !search ||
      p.name.toLowerCase().includes(q) ||
      p.company.toLowerCase().includes(q) ||
      p.title.toLowerCase().includes(q)
    );
  }, [sheetCompanies, search]);

  const tabCounts = useMemo(() => {
    const counts = {} as Record<TabId, number>;
    for (const tab of TABS) counts[tab.id] = (allData[tab.sheet] ?? []).length;
    return counts;
  }, [allData]);

  const currentMonthLabel = selectedMonth === 'all' ? 'All_Time' : selectedMonth.replace(' ', '_');
  const currentTabLabel   = TABS.find(t => t.id === activeTab)!.label;

  const exportTab = () => {
    if (!filtered.length) return;
    const rows = filtered.map(c => ({
      'Company Name': c.name, Revenue: c.revenue,
      Users: (c as any).users, Sessions: (c as any).sessions, Views: (c as any).views,
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
      const companies = allData[tab.sheet] ?? [];
      if (!companies.length) continue;
      const rows = companies.map(c => ({ 'Company Name': c.name, Revenue: c.revenue }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), tab.label.substring(0, 31));
      total += rows.length;
    }
    if (!total) return;
    XLSX.writeFile(wb, `TADA_Visitors_${currentMonthLabel}_All.xlsx`);
    showToast(`Exported ${total} companies across all tabs`);
  };

  return (
    <div className="dashboard">
      {toast && <div className="toast">{toast}</div>}

      <header className="dashboard-header">
        <div className="header-left"><h1>Website Visitors</h1></div>
        <div className="header-right">
          <button className="btn btn-secondary" onClick={exportTab} disabled={!filtered.length}>↓ Export Tab</button>
          <button className="btn btn-secondary" onClick={exportAll} disabled={loading}>↓ Export All</button>
        </div>
      </header>

      {/* Month selector */}
      <div className="month-bar">
        <button className={`month-btn${selectedMonth === 'all' ? ' month-btn-active' : ''}`} onClick={() => setSelectedMonth('all')}>All Time</button>
        {MONTH_KEYS.map(m => (
          <button key={m} className={`month-btn${selectedMonth === m ? ' month-btn-active' : ''}`} onClick={() => setSelectedMonth(m)}>{m}</button>
        ))}
      </div>

      {/* Industry tabs */}
      <div className="tabs">
        {TABS.map(tab => (
          <button key={tab.id} className={`tab${activeTab === tab.id ? ' tab-active' : ''}`}
            onClick={() => { setActiveTab(tab.id); setSearch(''); setViewMode('companies'); }}>
            {tab.label}
            <span className="tab-count">{tabCounts[tab.id] ?? 0}</span>
          </button>
        ))}
      </div>

      {/* Companies / Contacts toggle + search */}
      <div className="view-toggle-bar">
        <div className="view-toggle">
          <button className={`toggle-btn${viewMode === 'companies' ? ' toggle-active' : ''}`} onClick={() => setViewMode('companies')}>Companies</button>
          <button className={`toggle-btn${viewMode === 'contacts'  ? ' toggle-active' : ''}`} onClick={() => setViewMode('contacts')}>Contacts</button>
        </div>
        <div className="filters-right">
          <input className="search-input"
            placeholder={viewMode === 'contacts' ? 'Search contacts…' : 'Search companies…'}
            value={search} onChange={e => setSearch(e.target.value)} />
          <span className="result-count">
            {viewMode === 'contacts'
              ? `${filteredContacts.length} contacts`
              : `${filtered.length} companies`}
          </span>
        </div>
      </div>

      {/* Content */}
      {loading ? (
        <div className="loading">Loading visitor data…</div>
      ) : error ? (
        <div className="placeholder-state"><div className="placeholder-icon">⚠️</div><p>{error}</p></div>
      ) : viewMode === 'contacts' ? (
        <div className="table-wrapper">
          <table className="visitors-table">
            <thead>
              <tr>
                <th>Name</th><th>Title</th><th>Company</th><th>Email</th><th>LinkedIn</th>
              </tr>
            </thead>
            <tbody>
              {filteredContacts.length === 0 ? (
                <tr><td colSpan={5} className="empty-state">No contacts found</td></tr>
              ) : filteredContacts.map((p, i) => (
                <tr key={i}>
                  <td>{p.name}</td>
                  <td className="secondary-text">{p.title}</td>
                  <td>{p.company}</td>
                  <td className="secondary-text">{p.email}</td>
                  <td>
                    {p.linkedin && <a href={p.linkedin} target="_blank" rel="noopener noreferrer" className="linkedin-link">View ↗</a>}
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
                <th>Revenue</th>
                {activeTab === 'under1b' && <th>Category</th>}
                <th className="num-col">Users</th>
                <th className="num-col">Sessions</th>
                <th className="num-col">Views</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={activeTab === 'under1b' ? 6 : 5} className="empty-state">No companies match your search</td></tr>
              ) : filtered.map((c, i) => (
                <tr key={i}>
                  <td>
                    <div>{c.name}</div>
                    {c.contacts.length > 0 && (
                      <div className="contact-list">
                        {c.contacts.map((p, j) => (
                          <span key={j} className="contact-chip">
                            {p.linkedin
                              ? <a href={p.linkedin} target="_blank" rel="noopener noreferrer" className="contact-link">{p.name}</a>
                              : p.name}
                            {p.title && <span className="contact-title">{p.title}</span>}
                          </span>
                        ))}
                      </div>
                    )}
                  </td>
                  <td className="secondary-text">{c.revenue}</td>
                  {activeTab === 'under1b' && <td className="secondary-text">{c.category}</td>}
                  <td className="num-col">{(c as any).users || 0}</td>
                  <td className="num-col">{(c as any).sessions || 0}</td>
                  <td className="num-col">{(c as any).views || 0}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
