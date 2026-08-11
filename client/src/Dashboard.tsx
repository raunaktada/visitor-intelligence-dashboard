import React, { useState, useEffect, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import './Dashboard.css';

const MONTH_KEYS = ['Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'Jun 2026', 'Jul 2026'];

type TabId = 'all' | 'customers' | 'defense' | 'datacenter' | 'manufacturing' | 'healthcare' | 'cpg' | 'under1b';

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
  { id: 'all',           label: 'All Companies',        sheet: '' },
  { id: 'customers',     label: 'Customers & Partners', sheet: 'Customers & Partners' },
  { id: 'defense',       label: 'Defense',              sheet: 'Defense Manufacturers - 1B+' },
  { id: 'datacenter',    label: 'Data Center',          sheet: 'Data Center Mfg - 1B+' },
  { id: 'manufacturing', label: 'Manufacturing',        sheet: 'Manufacturing - 1B+' },
  { id: 'healthcare',    label: 'Healthcare / MedTech', sheet: 'Healthcare_MedTech - 1B+' },
  { id: 'cpg',           label: 'CPG',                  sheet: 'CPG - 1B+' },
  { id: 'under1b',       label: 'Under $1B',            sheet: 'Under $1B (250M-1B)' },
];

// Strip legal suffixes for fuzzy dedup (mirrors sharepoint normalizeName)
function normalizeName(n: string): string {
  return n.toLowerCase()
    .replace(/[',]/g, '')
    .replace(/\b(inc|llc|corp|corporation|ltd|plc|co|company|systems?|group|holdings|international|global|enterprises|industries)\b/g, '')
    .replace(/[^a-z0-9]/g, '')
    .trim();
}

// Hardcoded categories for Under $1B companies missing Category in the sheet
const UNDER1B_CATEGORIES: Record<string, string> = {
  // Healthcare / MedTech
  'hnl lab medicine':                          'Healthcare / MedTech',
  'ocli vision':                               'Healthcare / MedTech',
  'blanchard valley health system':            'Healthcare / MedTech',
  'texoma medical center':                     'Healthcare / MedTech',
  'the stepping stones group, llc':            'Healthcare / MedTech',
  'hurley medical center':                     'Healthcare / MedTech',
  'palos health':                              'Healthcare / MedTech',
  "children's hospital of the king's daughters": 'Healthcare / MedTech',
  'crouse health':                             'Healthcare / MedTech',
  'springfield clinic':                        'Healthcare / MedTech',
  'senderra rx specialty pharmacy':            'Healthcare / MedTech',
  'brightstar care':                           'Healthcare / MedTech',
  'owensboro health':                          'Healthcare / MedTech',
  "nicklaus children's health system":         'Healthcare / MedTech',
  'starkey hearing':                           'Healthcare / MedTech',
  'lifehealthcare':                            'Healthcare / MedTech',
  'mary washington healthcare':                'Healthcare / MedTech',
  'hillcrest healthcare system':               'Healthcare / MedTech',
  'eyebuydirect':                              'Healthcare / MedTech',
  'piedmont':                                  'Healthcare / MedTech',
  // CPG
  "stella & chewy's":                         'CPG',
  'varsity spirit':                            'CPG',
  "papa murphy's international":               'CPG',
  'world market':                              'CPG',
  'rodan + fields':                            'CPG',
  // Defense
  'draper':                                    'Defense',
  'sos international llc':                     'Defense',
  'janicki industries':                        'Defense',
  // Manufacturing
  "d'addario & company, inc.":                 'Manufacturing',
  'gsfsgroup':                                 'Manufacturing',
  'the tile shop':                             'Manufacturing',
  'camco manufacturing inc.':                  'Manufacturing',
  'warn industries':                           'Manufacturing',
  'civic merchandising inc':                   'Manufacturing',
  'kamco supply corp.':                        'Manufacturing',
  'giesecke & devrient america, inc':          'Manufacturing',
  'sb energy':                                 'Manufacturing',
  'aaf international':                         'Manufacturing',
  'core molding technologies':                 'Manufacturing',
  'piedmont plastics':                         'Manufacturing',
  'heatcraft worldwide refrigeration':         'Manufacturing',
  'techniplas':                                'Manufacturing',
  'carolina cat':                              'Manufacturing',
  'douglas dynamics, inc.':                    'Manufacturing',
  'onesource distributors':                    'Manufacturing',
  'cognex corporation':                        'Manufacturing',
  'chemtreat, inc':                            'Manufacturing',
  'crane chempharma & energy':                 'Manufacturing',
  'mustang cat':                               'Manufacturing',
  'carter machinery':                          'Manufacturing',
  'kamax':                                     'Manufacturing',
  'oregon tool, inc':                          'Manufacturing',
  'infineum':                                  'Manufacturing',
  'ogcc behavioral service center inc':        'Healthcare / MedTech',
  // Revenue-overflow companies that land in Under $1B
  'peloton interactive':                       'Manufacturing',
  'lozier corporation':                        'Manufacturing',
  'trimedx':                                   'Healthcare / MedTech',
  'solar turbines':                            'Manufacturing',
  'milton cat':                                'Manufacturing',
  'milton catt':                               'Manufacturing',
  'parker lord':                               'Manufacturing',
  'adentra group':                             'Manufacturing',
};

// Sub-categories for Manufacturing tab companies
const MFG_CATEGORIES: Record<string, string> = {
  // Automotive & Mobility
  'general motors':                   'Automotive & Mobility',
  'tesla, inc.':                      'Automotive & Mobility',
  'nissan motor corporation':         'Automotive & Mobility',
  'daimler truck holding':            'Automotive & Mobility',
  'volkswagen of america, inc':       'Automotive & Mobility',
  'american honda motor co., inc.':   'Automotive & Mobility',
  'adient':                           'Automotive & Mobility',
  'dana limited':                     'Automotive & Mobility',
  'mahle gmbh':                       'Automotive & Mobility',
  'clarios':                          'Automotive & Mobility',
  'inteva products':                  'Automotive & Mobility',
  'maxxis international':             'Automotive & Mobility',
  'goodyear tire & rubber':           'Automotive & Mobility',
  'cnh industrial':                   'Automotive & Mobility',
  'embraco':                          'Automotive & Mobility',
  'hutchinson industries, inc':       'Automotive & Mobility',
  // Electronics, Electrical & Technology
  'apple inc.':                       'Electronics, Electrical & Technology',
  'foxconn':                          'Electronics, Electrical & Technology',
  'panasonic':                        'Electronics, Electrical & Technology',
  'honeywell':                        'Electronics, Electrical & Technology',
  'general electric':                 'Electronics, Electrical & Technology',
  'eaton':                            'Electronics, Electrical & Technology',
  'samsung':                          'Electronics, Electrical & Technology',
  'motorola mobility llc':            'Electronics, Electrical & Technology',
  'superior essex':                   'Electronics, Electrical & Technology',
  't3 automation':                    'Electronics, Electrical & Technology',
  'julian electric inc.':             'Electronics, Electrical & Technology',
  'acuity brands, inc':               'Electronics, Electrical & Technology',
  'signify':                          'Electronics, Electrical & Technology',
  'sensata technologies, inc':        'Electronics, Electrical & Technology',
  'nidec':                            'Electronics, Electrical & Technology',
  'kuka':                             'Electronics, Electrical & Technology',
  'wavetronix':                       'Electronics, Electrical & Technology',
  'coorsTek, inc.':                   'Electronics, Electrical & Technology',
  // Industrial Machinery & Equipment
  'johnson controls':                 'Industrial Machinery & Equipment',
  'heatcraft worldwide refrigeration':'Industrial Machinery & Equipment',
  'stanley black & decker':           'Industrial Machinery & Equipment',
  'kress corporation':                'Industrial Machinery & Equipment',
  'agrinautics, inc.':                'Industrial Machinery & Equipment',
  'metso':                            'Industrial Machinery & Equipment',
  'lennox international inc':         'Industrial Machinery & Equipment',
  'konecranes':                       'Industrial Machinery & Equipment',
  'donaldson company, inc':           'Industrial Machinery & Equipment',
  'solar turbines':                   'Industrial Machinery & Equipment',
  'nordson':                          'Industrial Machinery & Equipment',
  'regal rexnord corporation':        'Industrial Machinery & Equipment',
  'oregon tool, inc':                 'Industrial Machinery & Equipment',
  'thermax limited':                  'Industrial Machinery & Equipment',
  'hillenbrand, inc':                 'Industrial Machinery & Equipment',
  'applied industrial technologies, inc.': 'Industrial Machinery & Equipment',
  'micropulse':                       'Industrial Machinery & Equipment',
  // Chemical & Specialty Chemical
  'bp p.l.c.':                        'Chemical & Specialty Chemical',
  'akzo nobel':                       'Chemical & Specialty Chemical',
  'the lebermuth company':            'Chemical & Specialty Chemical',
  'chevron phillips chemical company':'Chemical & Specialty Chemical',
  'huntsman international llc':       'Chemical & Specialty Chemical',
  'axalta':                           'Chemical & Specialty Chemical',
  'dow':                              'Chemical & Specialty Chemical',
  'the lubrizol corporation':         'Chemical & Specialty Chemical',
  'stepan company':                   'Chemical & Specialty Chemical',
  'wacker chemical corporation usa':  'Chemical & Specialty Chemical',
  'element solutions inc':            'Chemical & Specialty Chemical',
  'sea-land chemical company':        'Chemical & Specialty Chemical',
  // Packaging, Paper & Printing
  'international paper':              'Packaging, Paper & Printing',
  'crown holdings, inc':              'Packaging, Paper & Printing',
  'amcor':                            'Packaging, Paper & Printing',
  'avery dennison corporation':       'Packaging, Paper & Printing',
  'tetra pak':                        'Packaging, Paper & Printing',
  'domtar':                           'Packaging, Paper & Printing',
  'sun chemical':                     'Packaging, Paper & Printing',
  'wendel printing':                  'Packaging, Paper & Printing',
  'coloring book solutions':          'Packaging, Paper & Printing',
  // Building Materials, Construction & Industrial Materials
  'the sherwin-williams company':     'Building Materials & Construction',
  'holcim':                           'Building Materials & Construction',
  'ufp industries, inc':              'Building Materials & Construction',
  'clayton homes':                    'Building Materials & Construction',
  's.h. chooi fasteners':             'Building Materials & Construction',
  'superior essex':                   'Building Materials & Construction',
  // Food & Beverage
  'ab inbev':                         'Food & Beverage',
  'mccormick fona':                   'Food & Beverage',
  // Consumer Products & Appliances
  'whirlpool':                        'Consumer Products & Appliances',
  'ge appliances':                    'Consumer Products & Appliances',
};

const HEALTHCARE_CATEGORIES: Record<string, string> = {
  // Hospitals & Health Systems
  'adventhealth':                                        'Hospitals & Health Systems',
  'piedmont':                                            'Hospitals & Health Systems',
  'ssm health':                                          'Hospitals & Health Systems',
  'commonspirit':                                        'Hospitals & Health Systems',
  'advocate health care':                                'Hospitals & Health Systems',
  'ochsner health':                                      'Hospitals & Health Systems',
  'memorial hermann':                                    'Hospitals & Health Systems',
  'bon secours mercy':                                   'Hospitals & Health Systems',
  'carilion clinic':                                     'Hospitals & Health Systems',
  'henry ford health':                                   'Hospitals & Health Systems',
  'tenet healthcare':                                    'Hospitals & Health Systems',
  'lifepoint health':                                    'Hospitals & Health Systems',
  'hca healthcare':                                      'Hospitals & Health Systems',
  'ascension':                                           'Hospitals & Health Systems',
  'providence':                                          'Hospitals & Health Systems',
  'sutter health':                                       'Hospitals & Health Systems',
  'banner health':                                       'Hospitals & Health Systems',
  'davita inc':                                          'Hospitals & Health Systems',
  'community health systems':                            'Hospitals & Health Systems',
  'johns hopkins medicine':                              'Hospitals & Health Systems',
  'newyork-presbyterian hospital':                       'Hospitals & Health Systems',
  'christus health':                                     'Hospitals & Health Systems',
  'hackensack meridian health':                          'Hospitals & Health Systems',
  'vanderbilt university medical center':                'Hospitals & Health Systems',
  'md anderson cancer center':                           'Hospitals & Health Systems',
  'northwestern medicine':                               'Hospitals & Health Systems',
  'encompass health corporation':                        'Hospitals & Health Systems',
  'endeavor health':                                     'Hospitals & Health Systems',
  'baptist health':                                      'Hospitals & Health Systems',
  'carle health':                                        'Hospitals & Health Systems',
  'texas children\'s hospital':                          'Hospitals & Health Systems',
  'yale new haven hospital':                             'Hospitals & Health Systems',
  'froedtert health':                                    'Hospitals & Health Systems',
  'oak street health':                                   'Hospitals & Health Systems',
  'brown university health':                             'Hospitals & Health Systems',
  'franciscan missionaries of our lady health system':   'Hospitals & Health Systems',
  'catholic health':                                     'Hospitals & Health Systems',
  'christianacare':                                      'Hospitals & Health Systems',
  'cone health':                                         'Hospitals & Health Systems',
  'jackson health system':                               'Hospitals & Health Systems',
  'sca health':                                          'Hospitals & Health Systems',
  'children\'s healthcare of atlanta':                   'Hospitals & Health Systems',
  'university health':                                   'Hospitals & Health Systems',
  'tower health':                                        'Hospitals & Health Systems',
  'advent health medical group':                         'Hospitals & Health Systems',
  'children\'s hospital colorado':                       'Hospitals & Health Systems',
  'hca midwest health':                                  'Hospitals & Health Systems',
  'moffitt cancer center':                               'Hospitals & Health Systems',
  'children\'s minnesota':                               'Hospitals & Health Systems',
  'atlanticare':                                         'Hospitals & Health Systems',
  'bronson healthcare':                                  'Hospitals & Health Systems',
  'sunnybrook':                                          'Hospitals & Health Systems',
  // Medical Devices & Equipment
  'bd (becton dickinson)':                               'Medical Devices & Equipment',
  'baxter':                                              'Medical Devices & Equipment',
  'zimmer biomet':                                       'Medical Devices & Equipment',
  'endologix llc':                                       'Medical Devices & Equipment',
  'sciex':                                               'Medical Devices & Equipment',
  'aed essentials':                                      'Medical Devices & Equipment',
  'trimedx':                                             'Medical Devices & Equipment',
  'medtronic':                                           'Medical Devices & Equipment',
  'stryker':                                             'Medical Devices & Equipment',
  'fresenius medical care':                              'Medical Devices & Equipment',
  'edwards lifesciences corporation':                    'Medical Devices & Equipment',
  'hologic':                                             'Medical Devices & Equipment',
  'icu medical':                                         'Medical Devices & Equipment',
  'bracco':                                              'Medical Devices & Equipment',
  'teleflex':                                            'Medical Devices & Equipment',
  'natera':                                              'Medical Devices & Equipment',
  'siemens healthineers':                                'Medical Devices & Equipment',
  // Pharma & Life Sciences
  'abbvie inc':                                          'Pharma & Life Sciences',
  'pfizer inc':                                          'Pharma & Life Sciences',
  'bristol myers squibb':                                'Pharma & Life Sciences',
  'sanofi':                                              'Pharma & Life Sciences',
  'abbott':                                              'Pharma & Life Sciences',
  'boehringer ingelheim international gmbh':             'Pharma & Life Sciences',
  'csl':                                                 'Pharma & Life Sciences',
  'icon plc':                                            'Pharma & Life Sciences',
  'dr. reddy\'s laboratories':                           'Pharma & Life Sciences',
  'parexel international':                               'Pharma & Life Sciences',
  'radiology partners':                                  'Pharma & Life Sciences',
  'endo international plc':                              'Pharma & Life Sciences',
  'jubilant':                                            'Pharma & Life Sciences',
  'biopharmaceutical':                                   'Pharma & Life Sciences',
  // Health Insurance & Benefits
  'unitedhealth group':                                  'Health Insurance & Benefits',
  'cvs health':                                          'Health Insurance & Benefits',
  'united healthcare services, inc.':                    'Health Insurance & Benefits',
  'optum':                                               'Health Insurance & Benefits',
  'elevance health':                                     'Health Insurance & Benefits',
  'walgreen co':                                         'Health Insurance & Benefits',
  // Healthcare IT & Services
  'onco360 oncology pharmacy':                           'Healthcare IT & Services',
  'biopharmaceutical':                                   'Healthcare IT & Services',
};

const CPG_CATEGORIES: Record<string, string> = {
  'ab inbev':       'Food & Beverage',
  'mccormick fona': 'Food & Beverage',
};

// Map industry tab id → human-readable label for "All Companies" category column
const TAB_LABELS: Record<string, string> = {
  customers: 'Customers & Partners',
  defense: 'Defense',
  datacenter: 'Data Center',
  manufacturing: 'Manufacturing',
  healthcare: 'Healthcare / MedTech',
  cpg: 'CPG',
  under1b: 'Under $1B',
};

const API_URL = import.meta.env.PROD ? '' : ((import.meta.env.VITE_API_URL as string) || 'http://localhost:5001');

function revenueInBillions(cleaned: string): number {
  if (!cleaned) return Infinity; // unknown — don't filter out
  const n = parseFloat(cleaned.replace(/[^0-9.]/g, ''));
  return isNaN(n) ? Infinity : n;
}

// Clean revenue: strip ranges, ensure $X.XB format
function cleanRevenue(raw: string): string {
  if (!raw) return '';
  let s = raw.trim().replace(/^[~≈]/g, '').trim();
  // Take only first value if range (e.g. "$41.9B – $43.5B" → "$41.9B")
  s = s.split(/\s*[–\-]\s*\$[\d]/)[0].trim();
  // Ensure $ prefix
  if (s && !s.startsWith('$')) s = '$' + s;
  // Remove spaces: "$ 0.9 B" → "$0.9B"
  s = s.replace(/\$\s+/, '$').replace(/\s*B$/i, '');
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

  // Deduplicate companies across sheets: if a company appears in multiple sheets,
  // keep it only in the highest-priority tab (CPG/Healthcare/Defense over Manufacturing)
  const SHEET_PRIORITY = [
    'Customers & Partners',
    'Defense Manufacturers - 1B+',
    'Data Center Mfg - 1B+',
    'Healthcare_MedTech - 1B+',
    'CPG - 1B+',
    'Manufacturing - 1B+',
    'Under $1B (250M-1B)',
  ];
  // Companies that appear in the wrong sheet in SharePoint — force them to the right tab
  const FORCE_TO_SHEET: Record<string, string> = {
    'endologix llc':                'Healthcare_MedTech - 1B+',
    'varsity spirit':               'CPG - 1B+',
    'stein mart, inc':              'CPG - 1B+',
    'no more tears inc':            'CPG - 1B+',
    'aed essentials':               'Healthcare_MedTech - 1B+',
    'piedmont':                     'Healthcare_MedTech - 1B+',
    'ab inbev':                     'CPG - 1B+',
    'mccormick fona':               'CPG - 1B+',
  };

  const dedupedData = useMemo(() => {
    const claimed = new Set<string>();
    const result: Record<string, Company[]> = {};
    for (const sheet of SHEET_PRIORITY) result[sheet] = [];

    // First pass: assign each company to its forced sheet (if any), else its natural sheet
    for (const sheet of SHEET_PRIORITY) {
      for (const c of (allData[sheet] ?? [])) {
        const key = normalizeName(c.name);
        if (claimed.has(key)) continue;
        claimed.add(key);
        const targetSheet = FORCE_TO_SHEET[c.name.toLowerCase()] ?? sheet;
        result[targetSheet].push(c);
      }
    }
    return result;
  }, [allData]);

  const activeSheet = TABS.find(t => t.id === activeTab)!.sheet;
  const sheetCompanies: Company[] = useMemo(() => {
    if (activeTab === 'all') {
      // Combine all sheets in priority order, tagging each company with its industry
      const seen = new Set<string>();
      const result: Company[] = [];
      for (const tab of TABS.filter(t => t.id !== 'all')) {
        for (const c of (dedupedData[tab.sheet] ?? [])) {
          const key = normalizeName(c.name);
          if (seen.has(key)) continue;
          seen.add(key);
          result.push({ ...c, category: TAB_LABELS[tab.id] ?? tab.label });
        }
      }
      return result;
    }
    return dedupedData[activeSheet] ?? [];
  }, [dedupedData, activeSheet, activeTab]);

  // Company view: Sales Intel data only, sorted by sessions desc
  const addTotals = useCallback((c: Company) => {
    const category = c.category || UNDER1B_CATEGORIES[c.name.toLowerCase()] || MFG_CATEGORIES[c.name.toLowerCase()] || CPG_CATEGORIES[c.name.toLowerCase()] || HEALTHCARE_CATEGORIES[c.name.toLowerCase()] || '';
    if (selectedMonth === 'all') {
      const t = Object.values(c.months).reduce(
        (a, m) => ({ users: a.users + m.users, sessions: a.sessions + m.sessions, views: a.views + m.views }),
        { users: 0, sessions: 0, views: 0 }
      );
      return { ...c, category, revenue: cleanRevenue(c.revenue), ...t };
    }
    const m = c.months[selectedMonth] ?? { users: 0, sessions: 0, views: 0 };
    return { ...c, category, revenue: cleanRevenue(c.revenue), ...m };
  }, [selectedMonth]);

  const companiesWithTotals = useMemo(() => {
    let base: Company[];
    if (activeTab === 'all') {
      base = sheetCompanies; // already merged + deduped in sheetCompanies memo
    } else if (activeTab === 'under1b') {
      // Include the under1b sheet + overflow from other sheets where revenue < $1B
      const under1bNames = new Set((dedupedData['Under $1B (250M-1B)'] ?? []).map(c => c.name.toLowerCase()));
      const overflow: Company[] = [];
      for (const tab of TABS) {
        if (tab.id === 'under1b' || tab.id === 'customers') continue;
        for (const c of (dedupedData[tab.sheet] ?? [])) {
          const rev = revenueInBillions(cleanRevenue(c.revenue));
          if (rev < 1 && !under1bNames.has(c.name.toLowerCase())) overflow.push(c);
        }
      }
      base = [...sheetCompanies, ...overflow];
    } else {
      // Exclude companies with known revenue < $1B (but never filter Customers & Partners)
      base = activeTab === 'customers'
        ? sheetCompanies
        : sheetCompanies.filter(c => revenueInBillions(cleanRevenue(c.revenue)) >= 1);
    }

    return base
      .filter(c => revenueInBillions(cleanRevenue(c.revenue)) >= 0.25)
      .map(addTotals)
      .filter(c => selectedMonth === 'all' || (c as any).sessions > 0 || (c as any).users > 0)
      .sort((a, b) => {
        const ra = parseFloat((a.revenue || '0').replace(/[^0-9.]/g, '')) || 0;
        const rb = parseFloat((b.revenue || '0').replace(/[^0-9.]/g, '')) || 0;
        return rb - ra;
      });
  }, [sheetCompanies, allData, selectedMonth, activeTab, addTotals]);

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
    for (const tab of TABS) {
      if (tab.id === 'all') continue;
      const companies = (dedupedData[tab.sheet] ?? []).filter(
        c => revenueInBillions(cleanRevenue(c.revenue)) >= 0.25
      );
      counts[tab.id] = selectedMonth === 'all'
        ? companies.length
        : companies.filter(c => {
            const m = c.months[selectedMonth];
            return m && (m.users > 0 || m.sessions > 0);
          }).length;
    }
    counts['all'] = (Object.keys(counts) as TabId[]).reduce((sum, id) => sum + (counts[id] ?? 0), 0);
    return counts;
  }, [dedupedData, selectedMonth]);

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

  const showCategory = activeTab === 'under1b' || activeTab === 'all' || activeTab === 'manufacturing' || activeTab === 'healthcare';
  const colSpan = showCategory ? 6 : 5;

  return (
    <div className="dashboard">
      {toast && <div className="toast">{toast}</div>}

      <header className="dashboard-header">
        <div className="header-left">
          <img src="/tada-logo.png" alt="TADA" className="tada-logo" />
          <h1>Website Visitors</h1>
        </div>
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
                {showCategory && <th>Category</th>}
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
                  {showCategory && <td className="secondary-text">{c.category}</td>}
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
