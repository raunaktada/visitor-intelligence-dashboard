import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';

function getConfig() {
  return {
    tenantId:     process.env.AZURE_TENANT_ID!,
    clientId:     process.env.AZURE_CLIENT_ID!,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
    driveId:      process.env.SHAREPOINT_DRIVE_ID!,
    fileId:       process.env.SHAREPOINT_FILE_ID!,
  };
}

function buildClient(): Client {
  const config = getConfig();
  const credential = new ClientSecretCredential(
    config.tenantId,
    config.clientId,
    config.clientSecret,
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

export async function findSiteAndFileIds(siteName: string, fileName: string): Promise<void> {
  const client = buildClient();
  const sites = await client.api(`/sites?search=${encodeURIComponent(siteName)}`).get();
  const site = sites.value[0];
  if (!site) { console.error('No site found for:', siteName); return; }
  console.log('SHAREPOINT_SITE_ID=' + site.id);

  const drive = await client.api(`/sites/${site.id}/drive`).get();
  console.log('SHAREPOINT_DRIVE_ID=' + drive.id);

  const files = await client.api(`/sites/${site.id}/drive/items/root/children`).get();
  const file = files.value.find((f: { name: string }) => f.name === fileName);
  if (!file) { console.error('No file found for:', fileName); return; }
  console.log('SHAREPOINT_FILE_ID=' + file.id);
}

export const SHEETS = [
  'Summary',
  'Customers & Partners',
  'Data Centers',
  'Defense',
  'Manufacturing',
  'Healthcare & MedTech',
  'CPG',
];

async function getUsedRange(sheetName: string): Promise<unknown[][]> {
  const client = buildClient();
  const { driveId, fileId } = getConfig();

  const response = await client
    .api(`/drives/${driveId}/items/${fileId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`)
    .get();

  return (response.values as unknown[][]) || [];
}

export async function getSheetData(sheetName: string): Promise<Record<string, unknown>[]> {
  const rows = await getUsedRange(sheetName);
  if (rows.length < 2) return [];

  const headers = (rows[0] as string[]).map(h => String(h).trim()).filter(Boolean);

  return rows.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj: Record<string, unknown> = {};
      headers.forEach((h, i) => { if (h) obj[h] = row[i] ?? ''; });
      return obj;
    });
}

// Summary sheet is a free-form layout (title + info line + tab list), not a table.
// Return it as-is so the frontend can render it exactly like the Excel sheet.
export async function getSummaryRaw(): Promise<unknown[][]> {
  return getUsedRange('Summary');
}

export async function getAllSheetsData(): Promise<Record<string, unknown>> {
  const result: Record<string, unknown> = {};
  await Promise.all(
    SHEETS.map(async sheet => {
      try {
        result[sheet] = sheet === 'Summary' ? await getSummaryRaw() : await getSheetData(sheet);
      } catch (err) {
        console.error(`Failed to fetch sheet "${sheet}":`, err);
        result[sheet] = sheet === 'Summary' ? [] : [];
      }
    })
  );
  return result;
}
