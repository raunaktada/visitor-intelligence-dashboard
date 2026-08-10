import React, { useState, useEffect, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import './Dashboard.css';

const MONTH_KEYS = ['Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'Jun 2026', 'Jul 2026'];

type TabId = 'customers' | 'defense' | 'datacenter' | 'manufacturing' | 'healthcare' | 'cpg' | 'under1b';

interface MonthData { users: number; sessions: number; views: number; }
interface Contact { name: string; title: string; email: string; linkedin: string; pageViewed: string; }
interface Company {
  name: string;
  revenue: string;
  category?: string;
  months: Record<string, MonthData>;
  contacts: Contact[];
}
interface Individual {
  firstName: string; lastName: string; email: string; company: string;
  title: string; website: string; pageViews: number; linkedin: string;
}

const TABS: { id: TabId; label: string; sheet: string }[] = [
  { id: 'customers',     label: 'Customers & Partners', sheet: 'Customers & Partners' },
  { id: 'defense',       label: 'Defense',              sheet: 'Defense Manufacturers - 1B+' },
  { id: 'datacenter',    label: 'Data Center',          sheet: 'Data Center Mfg - 1B+' },
  { id: 'manufacturing', label: 'Manufacturing',        sheet: 'Manufacturing - 1B+' },
  { id: 'healthcare',    label: 'Healthcare / MedTech', sheet: 'Healthcare_MedTech - 1B+' },
  { id: 'cpg',           label: 'CPG',                  sheet: 'CPG - 1B+' },
  { id: 'under1b',       label: 'Under $1B',            sheet: 'Under $1B (250M-1B)' },
];

const API_URL = import.meta.env.PROD ? '' : ((import.meta.env.VITE_API_URL as string) || 'http://localhost:5001');

// Clean revenue: strip ranges, ensure $X.XB format
function cleanRevenue(raw: string): string {
  if (!raw) return '';
  let s = raw.trim().replace(/^[~≈]/g, '').trim();
  // Take only first value if range (e.g. "$41.9B – $43.5B" → "$41.9B")
  s = s.split(/\s*[–\-]\s*\$[\d]/)[0].trim();
  // Ensure $ prefix
  if (s && !s.startsWith('$')) s = '$' + s;
  // Remove spaces: "$ 0.9 B" → "$0.9B"
  s = s.replace(/\$\s+/, '$').replace(/\s+B$/i, 'B');
  return s;
}

// Classify a company/domain into a tab
const CUSTOMER_DOMAINS = new Set(['caterpillar.com','cummins.com','gehealthcare.com','oshkoshcorp.com','defense.oshkoshcorp.com','osfhealthcare.org','steelscape.com','pointcore.com','thewheelgroup.com']);
const DEFENSE_DOMAINS = new Set(['boeing.com','gdmissionsystems.com','collinsaerospace.com','lockheedmartin.com','raytheon.com','northropgrumman.com','l3harris.com','baesystems.com','draper.com','zistos.com','tlgaerospace.com','gdls.com','textron.com','lmco.com']);
const DEFENSE_KW = ['defense','aerospace','military','naval','weapons','armament'];
const DATACENTER_DOMAINS = new Set(['nvidia.com','intel.com','amd.com','tsmc.com','st.com','teradyne.com','lamresearch.com','appliedmaterials.com','broadcom.com','qualcomm.com','micron.com']);
const DATACENTER_KW = ['semiconductor','microelectronics','data center','datacenter'];
const HEALTH_KW = ['health','medical','hospital','clinic','care','pharma','biotech','therapeutics','dental','vision','orthopedic','surgical','diagnostics','radiology','oncology','hospice','rehabilitation','nursing','physician','medicines','drug'];
const CPG_KW = ['food','beverage','grocery','retail','beauty','cosmetic','household','cleaning','snack','tobacco','apparel','clothing','fashion','shoe','footwear'];

function classifyDomain(name: string, domain: string): TabId {
  const d = domain.toLowerCase(), n = name.toLowerCase();
  if (CUSTOMER_DOMAINS.has(d)) return 'customers';
  if (DEFENSE_DOMAINS.has(d) || DEFENSE_KW.some(k => n.includes(k))) return 'defense';
  if (DATACENTER_DOMAINS.has(d) || DATACENTER_KW.some(k => n.includes(k))) return 'datacenter';
  if (HEALTH_KW.some(k => n.includes(k) || d.includes(k))) return 'healthcare';
  if (CPG_KW.some(k => n.includes(k))) return 'cpg';
  return 'manufacturing';
}

function domainFromUrl(url: string): string {
  try { return new URL(url.startsWith('http') ? url : `https://${url}`).hostname.replace(/^www\./, '').toLowerCase(); }
  catch { return url.toLowerCase().replace(/^www\./, ''); }
}

function parseIndividuals(text: string): Individual[] {
  const lines = text.trim().split('\n');
  return lines.slice(1).flatMap(line => {
    const fields: string[] = []; let cur = '', inQ = false;
    for (const ch of line) {
      if (ch === '"') inQ = !inQ;
      else if (ch === ',' && !inQ) { fields.push(cur); cur = ''; }
      else cur += ch;
    }
    fields.push(cur);
    if (fields.length < 11) return [];
    return [{ firstName: fields[0].trim(), lastName: fields[1].trim(), email: fields[2].trim(),
      company: fields[3].trim(), title: fields[4].trim(), website: fields[5].trim(),
      pageViews: parseInt(fields[7]) || 0, linkedin: fields[10].trim() }];
  });
}

export default function Dashboard() {
  const [allData,       setAllData]       = useState<Record<string, Company[]>>({});
  const [individuals,   setIndividuals]   = useState<Individual[]>([]);
  const [loading,       setLoading]       = useState(true);
  const [error,         setError]         = useState('');
  const [selectedMonth, setSelectedMonth] = useState<string>('all');
  const [activeTab,     setActiveTab]     = useState<TabId>('customers');
  const [viewMode,      setViewMode]      = useState<'companies' | 'contacts'>('companies');
  const [search,        setSearch]        = useState('');
  const [toast,         setToast]         = useState('');

  const showToast = useCallback((msg: string) => {
    setToast(msg); setTimeout(() => setToast(''), 3000);
  }, []);

  useEffect(() => {
    Promise.all([
      fetch(`${API_URL}/api/visitors`).then(r => r.json()),
      fetch('/data/individuals-all.csv').then(r => r.text()).then(parseIndividuals).catch(() => []),
    ]).then(([data, inds]) => {
      setAllData(data);
      setIndividuals(inds);
      setLoading(false);
    }).catch(() => { setError('Failed to load data.'); setLoading(false); });
  }, []);

  const activeSheet = TABS.find(t => t.id === activeTab)!.sheet;
  const sheetCompanies: Company[] = useMemo(() => allData[activeSheet] ?? [], [allData, activeSheet]);

  // Company view: Sales Intel data only, sorted by sessions desc
  const companiesWithTotals = useMemo(() => {
    return sheetCompanies.map(c => {
      if (selectedMonth === 'all') {
        const t = Object.values(c.months).reduce(
          (a, m) => ({ users: a.users + m.users, sessions: a.sessions + m.sessions, views: a.views + m.views }),
          { users: 0, sessions: 0, views: 0 }
        );
        return { ...c, revenue: cleanRevenue(c.revenue), ...t };
      }
      const m = c.months[selectedMonth] ?? { users: 0, sessions: 0, views: 0 };
      return { ...c, revenue: cleanRevenue(c.revenue), ...m };
    })
    .filter(c => selectedMonth === 'all' || (c as any).sessions > 0 || (c as any).users > 0)
    .sort((a, b) => {
      // Under $1B: sort by revenue; others: by sessions
      if (activeTab === 'under1b') {
        const ra = parseFloat((a.revenue || '0').replace(/[^0-9.]/g, '')) || 0;
        const rb = parseFloat((b.revenue || '0').replace(/[^0-9.]/g, '')) || 0;
        return rb - ra;
      }
      return (b as any).sessions - (a as any).sessions;
    });
  }, [sheetCompanies, selectedMonth, activeTab]);

  const filtered = useMemo(() => {
    if (!search) return companiesWithTotals;
    const q = search.toLowerCase();
    return companiesWithTotals.filter(c =>
      c.name.toLowerCase().includes(q) || (c.category ?? '').toLowerCase().includes(q)
    );
  }, [companiesWithTotals, search]);

  // Contacts view: SharePoint contacts + Artisan individuals, classified to active tab
  const tabContacts = useMemo(() => {
    const contacts: { name: string; title: string; company: string; email: string; linkedin: string; pageViewed: string; source: string }[] = [];

    // SharePoint contacts from current tab
    for (const c of sheetCompanies) {
      for (const p of c.contacts) {
        contacts.push({ ...p, pageViewed: String(p.pageViewed ?? ''), company: c.name, source: 'sharepoint' });
      }
    }

    // Artisan individuals classified to this tab
    const seen = new Set(contacts.map(p => p.name.toLowerCase()));
    for (const p of individuals) {
      const domain = domainFromUrl(p.website);
      const tab = classifyDomain(p.company, domain);
      if (tab !== activeTab) continue;
      const fullName = `${p.firstName} ${p.lastName}`.trim().toLowerCase();
      if (seen.has(fullName)) continue;
      seen.add(fullName);
      contacts.push({
        name: `${p.firstName} ${p.lastName}`.trim(),
        title: p.title, company: p.company, email: p.email,
        linkedin: p.linkedin, pageViewed: String(p.pageViews || ''), source: 'artisan',
      });
    }

    return contacts;
  }, [sheetCompanies, individuals, activeTab]);

  const filteredContacts = useMemo(() => {
    if (!search) return tabContacts;
    const q = search.toLowerCase();
    return tabContacts.filter(p =>
      p.name.toLowerCase().includes(q) ||
      p.company.toLowerCase().includes(q) ||
      p.title.toLowerCase().includes(q)
    );
  }, [tabContacts, search]);

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
      'Company Name': c.name, 'Revenue (Billions)': c.revenue,
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
      const rows = companies.map(c => ({ 'Company Name': c.name, 'Revenue (Billions)': cleanRevenue(c.revenue) }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), tab.label.substring(0, 31));
      total += rows.length;
    }
    if (!total) return;
    XLSX.writeFile(wb, `TADA_Visitors_${currentMonthLabel}_All.xlsx`);
    showToast(`Exported ${total} companies across all tabs`);
  };

  const colSpan = activeTab === 'under1b' ? 6 : 5;

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

      <div className="month-bar">
        <button className={`month-btn${selectedMonth === 'all' ? ' month-btn-active' : ''}`} onClick={() => setSelectedMonth('all')}>All Time</button>
        {MONTH_KEYS.map(m => (
          <button key={m} className={`month-btn${selectedMonth === m ? ' month-btn-active' : ''}`} onClick={() => setSelectedMonth(m)}>{m}</button>
        ))}
      </div>

      <div className="tabs">
        {TABS.map(tab => (
          <button key={tab.id} className={`tab${activeTab === tab.id ? ' tab-active' : ''}`}
            onClick={() => { setActiveTab(tab.id); setSearch(''); setViewMode('companies'); }}>
            {tab.label}
            <span className="tab-count">{tabCounts[tab.id] ?? 0}</span>
          </button>
        ))}
      </div>

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
            {viewMode === 'contacts' ? `${filteredContacts.length} contacts` : `${filtered.length} companies`}
          </span>
        </div>
      </div>

      {loading ? (
        <div className="loading">Loading visitor data…</div>
      ) : error ? (
        <div className="placeholder-state"><div className="placeholder-icon">⚠️</div><p>{error}</p></div>
      ) : viewMode === 'contacts' ? (
        <div className="table-wrapper">
          <table className="visitors-table">
            <thead>
              <tr>
                <th>Name</th><th>Title</th><th>Company</th><th>Email</th><th className="num-col">Page Views</th><th>LinkedIn</th>
              </tr>
            </thead>
            <tbody>
              {filteredContacts.length === 0 ? (
                <tr><td colSpan={6} className="empty-state">No contacts found</td></tr>
              ) : filteredContacts.map((p, i) => (
                <tr key={i}>
                  <td>{p.name}</td>
                  <td className="secondary-text">{p.title}</td>
                  <td>{p.company}</td>
                  <td className="secondary-text">{p.email}</td>
                  <td className="num-col">
                    {p.pageViewed?.startsWith('http')
                      ? <a href={p.pageViewed} target="_blank" rel="noopener noreferrer" className="linkedin-link">View ↗</a>
                      : (p.pageViewed || '')}
                  </td>
                  <td>{p.linkedin && <a href={p.linkedin} target="_blank" rel="noopener noreferrer" className="linkedin-link">View ↗</a>}</td>
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
                <th>Revenue (Billions)</th>
                {activeTab === 'under1b' && <th>Category</th>}
                <th className="num-col">Users</th>
                <th className="num-col">Sessions</th>
                <th className="num-col">Views</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={colSpan} className="empty-state">No companies match your search</td></tr>
              ) : filtered.map((c, i) => (
                <tr key={i}>
                  <td>{c.name}</td>
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
