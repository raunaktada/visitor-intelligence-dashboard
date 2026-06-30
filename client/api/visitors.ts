import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getAllSheetsData } from './_lib/sharepoint.js';

// In-memory cache — resets on cold start, which naturally keeps this fresh.
// Serverless functions have no writable filesystem, so this replaces the
// old file-based cache from the Express version.
let cache: { timestamp: number; data: Record<string, unknown> } | null = null;
const CACHE_TTL_MS = 2 * 60 * 1000; // 2 minutes

export default async function handler(_req: VercelRequest, res: VercelResponse) {
  try {
    if (cache && Date.now() - cache.timestamp < CACHE_TTL_MS) {
      return res.status(200).json(cache.data);
    }

    const data = await getAllSheetsData();
    cache = { timestamp: Date.now(), data };
    res.status(200).json(data);
  } catch (err) {
    console.error('GET /api/visitors error:', err);
    res.status(500).json({ error: 'Failed to fetch data', detail: err instanceof Error ? err.message : String(err) });
  }
}
