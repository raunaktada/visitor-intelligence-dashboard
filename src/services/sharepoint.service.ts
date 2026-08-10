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

const REVENUE_OVERRIDES: Record<string, string> = {
  'lockheed martin':     '$71.0B',
  'kbr inc.':            '$7.4B',
  'kbr inc':             '$7.4B',
  '3m':                  '$24.6B',
  'parker hannifin':     '$19.9B',
  'grainger':            '$16.5B',
  'domtar':              '$5.4B',
  'adentra group':       '$2.3B',
  'siemens':             '$90.4B',
  'mccormick fona':      '$6.7B',
  'adventhealth':        '$9.0B',
  'trimedx':             '$0.3B',
  'peloton interactive': '$0.7B',
  'lozier corporation':  '$0.5B',
  'oshkosh defense':     '$10.1B',
  'bd':                  '$20.2B',
  'becton dickinson':    '$20.2B',
  'bd (becton dickinson)': '$20.2B',
  'examworks':           '$1.2B',
  'sciex':               '$1.0B',
  'parker lord':         '$1.1B',
  'ab inbev':            '$57.7B',
  'a inev':              '$57.7B', // corrupted "AB InBev" in source data
  'solar turbines':      '$3.1B',
  'solar turines':       '$3.1B', // typo in source data
  'crate and barrel':    '$2.5B',
  'crate & barrel':      '$2.5B',
  'crate and arrel':     '$2.5B', // typo in source data
  'milton cat':          '$0.72B',
};

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

const MONTH_KEYS = ['Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'Jun 2026', 'Jul 2026'];

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
    const name = String(row[idx('Company Name')] ?? '').trim();
    if (!name) continue;

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
    const displayName = String(row[nameIdx] ?? '').trim();
    const company = displayName.toLowerCase();
    const contactName = String(row[contactIdx] ?? '').trim();
    if (!company || !contactName) continue;

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

// Add companies from old file that are missing in new file
function addMissingCompanies(companies: Company[], oldContacts: Map<string, { displayName: string; contacts: Contact[] }>, existingNames: Set<string>): Company[] {
  const existingNorm = new Set([...existingNames].map(normalizeName));
  const extras: Company[] = [];
  for (const [nameLower, { displayName, contacts }] of oldContacts) {
    if (!existingNorm.has(normalizeName(nameLower))) {
      extras.push({ name: displayName, revenue: '', months: {}, contacts });
    }
  }
  return [...companies, ...extras];
}

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
  const oldBySheet = new Map<string, Map<string, Contact[]>>();
  for (const { newSheet, contacts } of oldResults) {
    oldBySheet.set(newSheet, contacts);
  }

  // Merge
  const result = {} as Record<string, Company[]>;
  for (const { sheet, companies } of newResults) {
    const oldContacts = oldBySheet.get(sheet) ?? new Map();
    const existingNames = new Set(companies.map(c => c.name.toLowerCase()));
    let merged = mergeContacts(companies, oldContacts);
    // For non-under1b sheets, add companies only in old file
    if (sheet !== 'Under $1B (250M-1B)') {
      merged = addMissingCompanies(merged, oldContacts, existingNames);
    }
    result[sheet] = merged.map(c => {
      const override = REVENUE_OVERRIDES[c.name.toLowerCase()];
      return override ? { ...c, revenue: override } : c;
    });
  }
  return result;
}
