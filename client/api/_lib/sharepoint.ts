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
function parseOldSheet(rows: unknown[][]): Map<string, Contact[]> {
  const result = new Map<string, Contact[]>();
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
    const company = String(row[nameIdx] ?? '').trim().toLowerCase();
    const contactName = String(row[contactIdx] ?? '').trim();
    if (!company || !contactName) continue;

    if (!result.has(company)) result.set(company, []);
    result.get(company)!.push({
      name:       contactName,
      title:      String(row[titleIdx]   ?? '').trim(),
      email:      String(row[emailIdx]   ?? '').trim(),
      linkedin:   String(row[linkedinIdx]?? '').trim(),
      pageViewed: String(row[pageIdx]    ?? '').trim(),
    });
  }
  return result;
}

// Merge old file contacts into new file companies (dedup by name)
function mergeContacts(companies: Company[], oldContacts: Map<string, Contact[]>): Company[] {
  return companies.map(c => {
    const existing = new Set(c.contacts.map(p => p.name.toLowerCase()));
    const fromOld  = oldContacts.get(c.name.toLowerCase()) ?? [];
    const newOnes  = fromOld.filter(p => !existing.has(p.name.toLowerCase()));
    return { ...c, contacts: [...c.contacts, ...newOnes] };
  });
}

// Add companies from old file that are missing in new file
function addMissingCompanies(companies: Company[], oldContacts: Map<string, Contact[]>, existingNames: Set<string>): Company[] {
  const extras: Company[] = [];
  for (const [nameLower, contacts] of oldContacts) {
    if (!existingNames.has(nameLower)) {
      // Find display name from first contact's row (we only have lowercase key)
      extras.push({ name: contacts[0]?.name ?? nameLower, revenue: '', months: {}, contacts });
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
    result[sheet] = merged;
  }
  return result;
}
