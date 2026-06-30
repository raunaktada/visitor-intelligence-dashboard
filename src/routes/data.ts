import express from 'express';
import fs from 'fs/promises';
import path from 'path';
import { getAllSheetsData } from '../services/sharepoint.service';

const router = express.Router();
const CACHE_DIR  = path.join(process.cwd(), '.cache');
const CACHE_FILE = path.join(CACHE_DIR, 'sharepoint-data.json');

// Short TTL so edits made in the Excel sheet show up on the dashboard
// without anyone having to click Sync.
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes

interface CachedData {
  timestamp: number;
  data: Record<string, unknown>;
}

async function readCache(): Promise<CachedData | null> {
  try {
    const raw = await fs.readFile(CACHE_FILE, 'utf-8');
    const cached = JSON.parse(raw) as CachedData;
    return Date.now() - cached.timestamp < CACHE_TTL_MS ? cached : null;
  } catch {
    return null;
  }
}

async function writeCache(data: Record<string, unknown>): Promise<void> {
  await fs.mkdir(CACHE_DIR, { recursive: true });
  await fs.writeFile(CACHE_FILE, JSON.stringify({ timestamp: Date.now(), data }, null, 2));
}

function countRecords(data: Record<string, unknown>): number {
  return Object.entries(data).reduce((sum, [sheet, rows]) => {
    if (sheet === 'Summary') return sum; // free-form layout, not records
    return sum + (Array.isArray(rows) ? rows.length : 0);
  }, 0);
}

// GET /api/visitors — all sheets, cached for 5 minutes
router.get('/visitors', async (_req, res) => {
  try {
    const cached = await readCache();
    if (cached) return res.json(cached.data);

    const data = await getAllSheetsData();
    await writeCache(data);
    res.json(data);
  } catch (err) {
    console.error('GET /api/visitors error:', err);
    res.status(500).json({ error: 'Failed to fetch data' });
  }
});

// POST /api/sync — force an immediate refresh, bypassing cache
router.post('/sync', async (_req, res) => {
  try {
    const data = await getAllSheetsData();
    await writeCache(data);
    res.json({ success: true, recordsSynced: countRecords(data), sheets: Object.keys(data) });
  } catch (err) {
    console.error('POST /api/sync error:', err);
    res.status(500).json({ error: 'Sync failed', detail: err instanceof Error ? err.message : String(err) });
  }
});

export default router;
