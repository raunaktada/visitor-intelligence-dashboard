import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';

function buildClient(): Client {
  const credential = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID!,
    process.env.AZURE_CLIENT_ID!,
    process.env.AZURE_CLIENT_SECRET!,
  );
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const token = await credential.getToken('https://graph.microsoft.com/.default');
        return token!.token;
      },
    },
  });
}

// New file: monthly Sales Intel data
const NEW_DRIVE = 'b!mxgBiE4Y5EKwH-FDOH5dR1sIhf1RyFZOgyEb7uS-1vZbHwDFacCRTIzeh9fzz1JA';
const NEW_FILE  = '017NEG76XN2O7UCUU3UREIMCADEFDVWNBE';

// Old file: richer contact data (emails, page URLs, more contacts)
const OLD_DRIVE = 'b!juWF_zDEBEmkD1hre1MUFGD6uu0RzNxAuyw6MJABNGTXiqJ3zgGgQ5LHMeJaokB6';
const OLD_FILE  = '012CCV4JPUVEVXXBXVYJHYW2RW4QMWXVTI';

// Maps old sheet names → new sheet names
const OLD_TO_NEW: Record<string, string> = {
  'Customers & Partners': 'Customers & Partners',
  'Data Centers':         'Data Center Mfg - 1B+',
  'Defense':              'Defense Manufacturers - 1B+',
  'Manufacturing':        'Manufacturing - 1B+',
  'Healthcare & MedTech': 'Healthcare_MedTech - 1B+',
  'CPG':                  'CPG - 1B+',
};

// Correct known typos in the NEW SharePoint file's company names
const NEW_FILE_NAME_FIXES: Record<string, string> = {
  'solar turines':                'Solar Turbines',
  'a inev':                       'AB InBev',
  'applied industrial tech inc.': 'Applied Industrial Technologies, Inc.',
  'crate and arrel':              'Crate and Barrel',
  'siemens':                      'Siemens Healthineers',
};

const REVENUE_OVERRIDES: Record<string, string> = {
  'lockheed martin':              '$71.0',
  'kbr inc.':                     '$7.4',
  'kbr inc':                      '$7.4',
  '3m':                           '$24.6',
  'parker hannifin':              '$19.9',
  'grainger':                     '$16.5',
  'domtar':                       '$5.4',
  'siemens':                      '$22.0',
  'siemens healthineers':         '$22.0',
  'mccormick fona':               '$6.7',
  'adventhealth':                 '$9.0',
  'trimedx':                      '$0.3',
  'peloton interactive':          '$0.7',
  'lozier corporation':           '$0.5',
  'oshkosh defense':              '$10.1',
  'bd':                           '$20.2',
  'becton dickinson':             '$20.2',
  'bd (becton dickinson)':        '$20.2',
  'examworks':                    '$1.2',
  'sciex':                        '$1.0',
  'parker lord':                  '$1.1',
  'ab inbev':                     '$57.7',
  'a inev':                       '$57.7',
  'solar turbines':               '$3.1',
  'solar turines':                '$3.1',
  'crate and barrel':             '$2.5',
  'crate & barrel':               '$2.5',
  'crate and arrel':              '$2.5',
  'milton cat':                   '$0.72',
  'regal rexnord corporation':    '$5.8',
  'varsity spirit':               '$0.4',
  'alphakor group':               '$0.05',
  'coorstek, inc.':               '$0.82',
  'heatcraft worldwide refrigeration': '$0.34',
  'oregon tool, inc':             '$0.87',
  'piedmont':                     '$6.5',
};

const EXCLUDED_COMPANIES = new Set([
  'apple inc.',
  'apple',
  // Non-manufacturers in Manufacturing tab
  'applied industrial technologies, inc.',
  'bp p.l.c.',
  'bradfield\'s, inc.',
  'coloring book solutions',
  'envu, llc',
  'powell electronics',
  'sea-land chemical company',
  // Manufacturing revenue < $250M
  'agrinautics, inc.',
  'global hemp, inc',
  'hutchinson industries, inc',
  'julian electric inc.',
  'keg',
  'kress corporation',
  'micropulse',
  'plural',
  's.h. chooi fasteners',
  'selds, inc.',
  't3 automation',
  'the lebermuth company',
  'universal renewables',
  'wavetronix',
  'wendel printing',
  'markham honda',
  'dodge city smiles',
  'onesource distributors',
  'kamco supply corp.',
  'gec2',
  'blick center',
  'adentra group',
  'oliver wyman',
  'parkview julian convalescent',
  'center for prevention of abuse',
  'maryhaven, inc.',
  'lifestance health',
  'examworks',
  // Health Insurance & Benefits
  'unitedhealth group',
  'cvs health',
  'united healthcare services, inc.',
  'optum',
  'elevance health',
  'walgreen co',
  // Healthcare IT & Services
  'change healthcare',
  'evolent health',
  'icon plc',
  'parexel international',
  'siemens healthineers',
  'davita inc',
  'davita',
  'fresenius medical care',
  'oak street health',
  'radiology partners',
  'sca health',
  'bioplus specialty pharmacy',
  'onco360 oncology pharmacy',
  // Behavioral Health & Rehabilitation
  'encompass health corporation',
  'encompass health',
  // Defense
  'nasa',
  'national aeronautics and space administration',
  // CPG non-fit
  'crate and barrel',
  'crate & barrel',
  'chico\'s fas, inc',
  'mr price group',
  'directv',
  'instacart',
  'etsy',
  'football fanatics',
  'kwik trip, inc.',
  'stein mart, inc',
  'the exchange',
  // NGOs / non-fits
  'no more tears inc',
  'the stepping stones group, llc',
  // Manufacturing tab non-manufacturers
  'alphakor group',
  'carter machinery',
  'mustang cat',
  'milton cat',
  'chemtreat, inc',
  'new york jewelers',
  // Defense duplicate
  'parker-hannifin corp',
  // Data Center non-fit
  'grainger',
  // Under $1B non-manufacturers / wrong industry
  'eyebuydirect',
  'gsfsgroup',
  'sb energy',
  'sos international llc',
  'world market',
  'the tile shop',
  'civic merchandising inc',
  // Under $1B non-fit
  'peloton interactive',
  'papa murphy\'s international',
  'ogcc behavioral service center inc',
  'brightstar care',
  'senderra rx specialty pharmacy',
]);

export const SHEET_NAMES = [
  'Customers & Partners',
  'Defense Manufacturers - 1B+',
  'Data Center Mfg - 1B+',
  'Manufacturing - 1B+',
  'Healthcare_MedTech - 1B+',
  'CPG - 1B+',
  'Under $1B (250M-1B)',
] as const;

export interface MonthData { users: number; sessions: number; views: number; }
export interface Contact {
  name: string; title: string; email: string; linkedin: string;
  pageViewed: string; // number or URL
}
export interface Company {
  name: string;
  revenue: string;
  category?: string;
  months: Record<string, MonthData>;
  contacts: Contact[];
}

const MONTH_KEYS = ['Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'Jun 2026', 'Jul 2026', 'Aug 2026'];

async function getRawRows(driveId: string, fileId: string, sheetName: string): Promise<unknown[][]> {
  const client = buildClient();
  const response = await client
    .api(`/drives/${driveId}/items/${fileId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`)
    .get();
  return (response.values as unknown[][]) || [];
}

function parseNewSheet(rows: unknown[][], hasCategory: boolean): Company[] {
  if (rows.length < 2) return [];
  const headers = (rows[0] as string[]).map(h => String(h ?? '').trim());
  const idx = (name: string) => headers.indexOf(name);
  const companyMap = new Map<string, Company>();

  for (const rawRow of rows.slice(1)) {
    const row = rawRow as (string | number)[];
    const rawName = String(row[idx('Company Name')] ?? '').trim();
    const name = NEW_FILE_NAME_FIXES[rawName.toLowerCase()] ?? rawName;
    if (!name || EXCLUDED_COMPANIES.has(name.toLowerCase())) continue;

    if (!companyMap.has(name)) {
      const revenue = String(
        row[idx('Company Revenue (in Billions)')] ||
        row[idx('Annual Revenue (in Billions USD)')] || ''
      ).trim();
      const months: Record<string, MonthData> = {};
      for (const m of MONTH_KEYS) {
        const u = Number(row[idx(`${m} Users`)])    || 0;
        const s = Number(row[idx(`${m} Sessions`)]) || 0;
        const v = Number(row[idx(`${m} Views`)])    || 0;
        if (u || s || v) months[m] = { users: u, sessions: s, views: v };
      }
      const company: Company = { name, revenue, months, contacts: [] };
      if (hasCategory) company.category = String(row[idx('Category')] ?? '').trim();
      companyMap.set(name, company);
    }

    const contactName = String(row[idx('Contact Name')] ?? '').trim();
    if (contactName) {
      companyMap.get(name)!.contacts.push({
        name:       contactName,
        title:      String(row[idx('Job Title')]   ?? '').trim(),
        email:      String(row[idx('Email')]       ?? '').trim(),
        linkedin:   String(row[idx('LinkedIn')]    ?? '').trim(),
        pageViewed: String(row[idx('Page Viewed')] ?? '').trim(),
      });
    }
  }
  return Array.from(companyMap.values());
}

// Correct known typos in the old SharePoint file's company names
const OLD_FILE_TYPOS: Record<string, string> = {
  'goodyear tire & ruer':       'Goodyear Tire & Rubber',
  'stanley lack & decker':      'Stanley Black & Decker',
  'watec corporation':          'Wabtec Corporation',
  'acuity rands, inc':          'Acuity Brands, Inc',
  'emraco':                     'Embraco',
  'hillenrand, inc':            'Hillenbrand, Inc',
  'p p.l.c.':                   'BP p.l.c.',
  'target rands, inc':          'Target Brands, Inc',
  'procter & gamle':            'Procter & Gamble',
  'h-e-':                       'H-E-B',
  'ulta eauty':                 'Ulta Beauty',
  'en e. keith foods':          'Ben E. Keith Foods',
  'the lurizol corporation':    'The Lubrizol Corporation',
  'a inev':                     'AB InBev',
  'crate and arrel':            'Crate and Barrel',
  'siemens':                    'Siemens Healthineers',
  'solar turines':              'Solar Turbines',
};

// Returns contacts grouped by normalized company name from old file
function parseOldSheet(rows: unknown[][]): Map<string, { displayName: string; contacts: Contact[] }> {
  const result = new Map<string, { displayName: string; contacts: Contact[] }>();
  if (rows.length < 2) return result;
  const headers = (rows[0] as string[]).map(h => String(h ?? '').trim());
  const idx = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const nameIdx    = idx('Company Name');
  const contactIdx = idx('Contact Name');
  const titleIdx   = Math.max(idx('Job Title'), idx('Jo Title'));
  const emailIdx   = idx('Email');
  const linkedinIdx= idx('LinkedIn');
  const pageIdx    = idx('Page Viewed');

  for (const rawRow of rows.slice(1)) {
    const row = rawRow as (string | number)[];
    const raw = String(row[nameIdx] ?? '').trim();
    const displayName = OLD_FILE_TYPOS[raw.toLowerCase()] ?? raw;
    const company = displayName.toLowerCase();
    const contactName = String(row[contactIdx] ?? '').trim();
    if (!company || !contactName || EXCLUDED_COMPANIES.has(company)) continue;

    if (!result.has(company)) result.set(company, { displayName, contacts: [] });
    result.get(company)!.contacts.push({
      name:       contactName,
      title:      String(row[titleIdx]   ?? '').trim(),
      email:      String(row[emailIdx]   ?? '').trim(),
      linkedin:   String(row[linkedinIdx]?? '').trim(),
      pageViewed: String(row[pageIdx]    ?? '').trim(),
    });
  }
  return result;
}

// Strip legal suffixes and filler words so "Caterpillar Inc." == "Caterpillar"
function normalizeName(n: string): string {
  return n.toLowerCase()
    .replace(/[,.]|(\b(inc|llc|corp|corporation|ltd|plc|co|company|system|systems|group|holdings|international|global|enterprises|industries|benckiser)\b)/g, '')
    .replace(/\s+/g, ' ').trim();
}

// Merge old file contacts into new file companies (dedup by name)
function mergeContacts(companies: Company[], oldContacts: Map<string, { displayName: string; contacts: Contact[] }>): Company[] {
  const byNorm = new Map<string, Contact[]>();
  for (const [nameLower, { contacts }] of oldContacts) {
    byNorm.set(normalizeName(nameLower), contacts);
  }

  return companies.map(c => {
    const existing = new Set(c.contacts.map(p => p.name.toLowerCase()));
    const fromOld  = byNorm.get(normalizeName(c.name)) ?? [];
    const newOnes  = fromOld.filter(p => !existing.has(p.name.toLowerCase()));
    return { ...c, contacts: [...c.contacts, ...newOnes] };
  });
}

// Add companies from old file that are missing across ALL new file sheets
function addMissingCompanies(companies: Company[], oldContacts: Map<string, { displayName: string; contacts: Contact[] }>, allNewFileNorm: Set<string>): Company[] {
  const existingNorm = allNewFileNorm;
  const extras: Company[] = [];
  for (const [nameLower, { displayName, contacts }] of oldContacts) {
    if (!existingNorm.has(normalizeName(nameLower))) {
      extras.push({ name: displayName, revenue: '', months: {}, contacts });
    }
  }
  return [...companies, ...extras];
}

// Aug 2026 visit data — injected directly until SharePoint is updated manually
const AUG_2026_PATCH: Record<string, { sheet: string; revenue?: string; aug: MonthData }> = {
  'Eaton':                    { sheet: 'Manufacturing - 1B+',      aug: { users: 2, sessions: 3, views: 5 } },
  'General Mills Inc':        { sheet: 'CPG - 1B+',               aug: { users: 2, sessions: 2, views: 2 } },
  'Nestle':                   { sheet: 'CPG - 1B+',               aug: { users: 1, sessions: 1, views: 1 } },
  "Macy's":                   { sheet: 'CPG - 1B+',               aug: { users: 1, sessions: 1, views: 2 } },
  'Target Brands, Inc':       { sheet: 'CPG - 1B+',               aug: { users: 1, sessions: 1, views: 1 } },
  'Cargill':                  { sheet: 'Manufacturing - 1B+',      aug: { users: 1, sessions: 1, views: 1 } },
  'Southeastern Grocers':     { sheet: 'CPG - 1B+',               aug: { users: 1, sessions: 1, views: 1 } },
  'OhioHealth':               { sheet: 'Healthcare_MedTech - 1B+', revenue: '$5.0',  aug: { users: 3, sessions: 3, views: 3 } },
  'AmeriGas Propane, Inc.':   { sheet: 'Manufacturing - 1B+',      revenue: '$3.7',  aug: { users: 3, sessions: 3, views: 3 } },
  'Merck & Co., Inc':         { sheet: 'Healthcare_MedTech - 1B+', revenue: '$60.1', aug: { users: 1, sessions: 1, views: 1 } },
  'Boston Scientific Corporation': { sheet: 'Healthcare_MedTech - 1B+', revenue: '$14.2', aug: { users: 1, sessions: 1, views: 1 } },
  'LabCorp':                  { sheet: 'Healthcare_MedTech - 1B+', revenue: '$12.2', aug: { users: 1, sessions: 1, views: 2 } },
  'Kaiser Permanente':        { sheet: 'Healthcare_MedTech - 1B+', revenue: '$100.0',aug: { users: 1, sessions: 1, views: 1 } },
  'LifeBridge Health':        { sheet: 'Healthcare_MedTech - 1B+', revenue: '$2.5',  aug: { users: 1, sessions: 1, views: 1 } },
  'Aspen Dental Management':  { sheet: 'Healthcare_MedTech - 1B+', revenue: '$4.0',  aug: { users: 1, sessions: 1, views: 3 } },
  'Rockwell Automation':      { sheet: 'Manufacturing - 1B+',      revenue: '$8.3',  aug: { users: 1, sessions: 1, views: 1 } },
  'Vishay Intertechnology':   { sheet: 'Manufacturing - 1B+',      revenue: '$3.5',  aug: { users: 1, sessions: 1, views: 1 } },
  'First Solar':              { sheet: 'Manufacturing - 1B+',      revenue: '$3.3',  aug: { users: 1, sessions: 1, views: 1 } },
  'Taylor Corporation':       { sheet: 'Manufacturing - 1B+',      revenue: '$2.0',  aug: { users: 1, sessions: 1, views: 1 } },
  'FXI':                      { sheet: 'Manufacturing - 1B+',      revenue: '$1.2',  aug: { users: 1, sessions: 1, views: 1 } },
  'Tory Burch':               { sheet: 'CPG - 1B+',               revenue: '$1.5',  aug: { users: 1, sessions: 1, views: 1 } },
  'Kellwood Company':         { sheet: 'CPG - 1B+',               revenue: '$1.0',  aug: { users: 1, sessions: 1, views: 1 } },
  'Y-12 National Security Complex': { sheet: 'Defense Manufacturers - 1B+', revenue: '$2.5', aug: { users: 1, sessions: 2, views: 2 } },
  'Zoox':                     { sheet: 'Data Center Mfg - 1B+',    revenue: '$1.0',  aug: { users: 1, sessions: 1, views: 2 } },
  'Ford Motor Company':                    { sheet: 'Manufacturing - 1B+',      revenue: '$185.0', aug: { users: 1, sessions: 1, views: 3 } },
  'John Deere':                            { sheet: 'Manufacturing - 1B+',      revenue: '$52.0',  aug: { users: 1, sessions: 1, views: 2 } },
  'Pactiv Evergreen Inc.':                 { sheet: 'Manufacturing - 1B+',      revenue: '$5.8',   aug: { users: 1, sessions: 1, views: 2 } },
  'Bridgestone Americas Inc.':             { sheet: 'Manufacturing - 1B+',      revenue: '$3.0',   aug: { users: 1, sessions: 1, views: 2 } },
  'Zeiss Industrial Quality Solutions':    { sheet: 'Manufacturing - 1B+',      revenue: '$8.0',   aug: { users: 1, sessions: 1, views: 2 } },
  'Wipro':                                 { sheet: 'CPG - 1B+',               revenue: '$11.0',  aug: { users: 1, sessions: 1, views: 2 } },
  'NorthShore University HealthSystem':    { sheet: 'Healthcare_MedTech - 1B+', revenue: '$3.0',   aug: { users: 1, sessions: 1, views: 4 } },
  "St. Luke's University Health Network":  { sheet: 'Healthcare_MedTech - 1B+', revenue: '$2.5',   aug: { users: 1, sessions: 1, views: 2 } },
  'Park Nicollet Health Services':         { sheet: 'Healthcare_MedTech - 1B+', revenue: '$1.5',   aug: { users: 1, sessions: 1, views: 3 } },
};

export async function getAllSheetsData(): Promise<Record<string, Company[]>> {
  // Fetch new file + old file in parallel
  const oldSheetNames = Object.keys(OLD_TO_NEW);

  const [newResults, oldResults] = await Promise.all([
    // New file: all sheets
    Promise.all(SHEET_NAMES.map(async sheet => {
      try {
        const rows = await getRawRows(NEW_DRIVE, NEW_FILE, sheet);
        return { sheet, companies: parseNewSheet(rows, sheet === 'Under $1B (250M-1B)') };
      } catch (err) {
        console.error(`New file - failed sheet "${sheet}":`, err);
        return { sheet, companies: [] as Company[] };
      }
    })),
    // Old file: contact sheets only
    Promise.all(oldSheetNames.map(async oldSheet => {
      try {
        const rows = await getRawRows(OLD_DRIVE, OLD_FILE, oldSheet);
        return { newSheet: OLD_TO_NEW[oldSheet], contacts: parseOldSheet(rows) };
      } catch (err) {
        console.error(`Old file - failed sheet "${oldSheet}":`, err);
        return { newSheet: OLD_TO_NEW[oldSheet], contacts: new Map<string, Contact[]>() };
      }
    })),
  ]);

  // Index old contacts by new sheet name
  const oldBySheet = new Map<string, Map<string, { displayName: string; contacts: Contact[] }>>();
  for (const { newSheet, contacts } of oldResults) {
    oldBySheet.set(newSheet, contacts);
  }

  // Build a set of ALL company names across ALL new file sheets (cross-sheet dedup)
  // so old file companies that exist anywhere in the new file are never re-added
  const allNewFileNorm = new Set<string>();
  for (const { companies } of newResults) {
    for (const c of companies) allNewFileNorm.add(normalizeName(c.name));
  }

  // Merge
  const result = {} as Record<string, Company[]>;
  for (const { sheet, companies } of newResults) {
    const oldContacts = oldBySheet.get(sheet) ?? new Map();
    let merged = mergeContacts(companies, oldContacts);
    // Add companies only in old file AND not present in ANY new file sheet
    if (sheet !== 'Under $1B (250M-1B)') {
      merged = addMissingCompanies(merged, oldContacts, allNewFileNorm);
    }
    result[sheet] = merged.map(c => {
      const override = REVENUE_OVERRIDES[c.name.toLowerCase()];
      return override ? { ...c, revenue: override } : c;
    });
  }

  // Merge Aug 2026 patch into result
  for (const [name, { sheet, revenue, aug }] of Object.entries(AUG_2026_PATCH)) {
    if (!result[sheet]) continue;
    const existing = result[sheet].find(c => c.name.toLowerCase() === name.toLowerCase());
    if (existing) {
      existing.months['Aug 2026'] = aug;
    } else {
      result[sheet].push({ name, revenue: revenue ?? '', months: { 'Aug 2026': aug }, contacts: [] });
    }
  }

  return result;
}
