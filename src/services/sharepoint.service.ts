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

export const SHEET_NAMES = [
  'Customers & Partners',
  'Defense Manufacturers - 1B+',
  'Data Center Mfg - 1B+',
  'Manufacturing - 1B+',
  'Healthcare_MedTech - 1B+',
  'CPG - 1B+',
  'Under $1B (250M-1B)',
] as const;

export type SheetName = typeof SHEET_NAMES[number];

export interface MonthData { users: number; sessions: number; views: number; }
export interface Contact { name: string; title: string; email: string; linkedin: string; pageViewed: number; }
export interface Company {
  name: string;
  revenue: string;
  category?: string;
  months: Record<string, MonthData>;
  contacts: Contact[];
}

const MONTH_KEYS = ['Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'Jun 2026', 'Jul 2026'];

async function getRawRows(sheetName: string): Promise<unknown[][]> {
  const client = buildClient();
  const driveId = process.env.SHAREPOINT_DRIVE_ID!;
  const fileId  = process.env.SHAREPOINT_FILE_ID!;
  const response = await client
    .api(`/drives/${driveId}/items/${fileId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`)
    .get();
  return (response.values as unknown[][]) || [];
}

function parseSheet(rows: unknown[][], hasCategory: boolean): Company[] {
  if (rows.length < 2) return [];
  const headers = (rows[0] as string[]).map(h => String(h ?? '').trim());

  // Group rows by company name — multiple rows = multiple contacts
  const companyMap = new Map<string, Company>();

  for (const rawRow of rows.slice(1)) {
    const row = rawRow as (string | number)[];
    const name = String(row[headers.indexOf('Company Name')] ?? '').trim();
    if (!name) continue;

    if (!companyMap.has(name)) {
      const revenue = String(row[headers.indexOf('Company Revenue (in Billions)')] || row[headers.indexOf('Annual Revenue (in Billions USD)')] || '').trim();
      const months: Record<string, MonthData> = {};
      for (const m of MONTH_KEYS) {
        const u = Number(row[headers.indexOf(`${m} Users`)])    || 0;
        const s = Number(row[headers.indexOf(`${m} Sessions`)]) || 0;
        const v = Number(row[headers.indexOf(`${m} Views`)])    || 0;
        if (u || s || v) months[m] = { users: u, sessions: s, views: v };
      }
      const company: Company = { name, revenue, months, contacts: [] };
      if (hasCategory) company.category = String(row[headers.indexOf('Category')] ?? '').trim();
      companyMap.set(name, company);
    }

    const contactName = String(row[headers.indexOf('Contact Name')] ?? '').trim();
    if (contactName) {
      companyMap.get(name)!.contacts.push({
        name:       contactName,
        title:      String(row[headers.indexOf('Job Title')]   ?? '').trim(),
        email:      String(row[headers.indexOf('Email')]       ?? '').trim(),
        linkedin:   String(row[headers.indexOf('LinkedIn')]    ?? '').trim(),
        pageViewed: Number(row[headers.indexOf('Page Viewed')] ?? 0),
      });
    }
  }

  return Array.from(companyMap.values());
}

export async function getAllSheetsData(): Promise<Record<SheetName, Company[]>> {
  const result = {} as Record<SheetName, Company[]>;
  await Promise.all(
    SHEET_NAMES.map(async sheet => {
      try {
        const rows = await getRawRows(sheet);
        result[sheet] = parseSheet(rows, sheet === 'Under $1B (250M-1B)');
      } catch (err) {
        console.error(`Failed to fetch sheet "${sheet}":`, err);
        result[sheet] = [];
      }
    })
  );
  return result;
}
