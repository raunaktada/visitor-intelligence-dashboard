import { HttpFunction } from '@google-cloud/functions-framework';
import axios from 'axios';

const API_URL = process.env.BACKEND_URL || 'http://localhost:5000';

export const scheduledSync: HttpFunction = async (req, res) => {
  try {
    const result = await axios.post(`${API_URL}/api/sync`, {
      triggeredBy: 'cloud-scheduler',
    });
    console.log('Sync complete:', result.data);
    res.json({ success: true, ...result.data });
  } catch (err) {
    console.error('Scheduled sync failed:', err);
    res.status(500).json({ error: 'Sync failed' });
  }
};
