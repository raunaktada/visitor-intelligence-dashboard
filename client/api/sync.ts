import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getAllSheetsData } from './_lib/sharepoint.js';

function countRecords(data: Record<string, unknown>): number {
  return Object.entries(data).reduce((sum, [sheet, rows]) => {
    if (sheet === 'Summary') return sum; // free-form layout, not records
    return sum + (Array.isArray(rows) ? rows.length : 0);
  }, 0);
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const data = await getAllSheetsData();
    res.status(200).json({ success: true, recordsSynced: countRecords(data), sheets: Object.keys(data) });
  } catch (err) {
    console.error('POST /api/sync error:', err);
    res.status(500).json({ error: 'Sync failed', detail: err instanceof Error ? err.message : String(err) });
  }
}
