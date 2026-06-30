import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import path from 'path';
import cron from 'node-cron';
import dataRouter from './routes/data';

dotenv.config({ path: path.resolve(__dirname, '../.env') });

const app = express();
app.use(cors({ origin: process.env.CORS_ORIGIN || 'http://localhost:3000' }));
app.use(express.json());

app.use('/api', dataRouter);

app.get('/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

const PORT = process.env.PORT || 5000;

// Keep the cache warm every 5 minutes so SharePoint edits show up
// on the dashboard without anyone clicking Sync.
cron.schedule('*/5 * * * *', async () => {
  try {
    const res = await fetch(`http://localhost:${PORT}/api/sync`, { method: 'POST' });
    const data = await res.json();
    console.log('[cron] Background sync:', data);
  } catch (err) {
    console.error('[cron] Background sync failed:', err);
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log(`Try: http://localhost:${PORT}/api/visitors`);
});

export default app;
