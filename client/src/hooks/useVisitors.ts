import { useState, useEffect } from 'react';

export interface VisitorRecord {
  id: string;
  company: string;
  contact: string;
  email: string;
  lastVisit: string;
  pagesViewed: string;
  visitCount: number;
  leadStatus: string;
  techStack: string;
  notes: string;
  visitingStatus: string;
  lastEdited: string;
  editedBy: string;
}

const API = import.meta.env.VITE_API_URL || 'http://localhost:5000';

export function useVisitors() {
  const [visitors, setVisitors] = useState<VisitorRecord[]>([]);
  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState<string | null>(null);

  const fetchVisitors = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`${API}/api/visitors`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setVisitors(await res.json());
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error');
    } finally {
      setLoading(false);
    }
  };

  const sync = async (): Promise<number> => {
    const res = await fetch(`${API}/api/sync`, { method: 'POST' });
    if (!res.ok) throw new Error('Sync failed');
    const data = await res.json();
    await fetchVisitors();
    return data.recordsSynced;
  };

  useEffect(() => { fetchVisitors(); }, []);

  return { visitors, loading, error, refetch: fetchVisitors, sync };
}
