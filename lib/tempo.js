export const TEMPO_URL = 'https://api.tempo.io/4';

export function tempoAuth() {
  const token = process.env.TEMPO_API_TOKEN;
  if (!token) throw new Error('TEMPO_API_TOKEN no configurado en .env.local');
  return token;
}

// Fetch all worklogs from Tempo for a date range, optionally filtered by accountIds
export async function fetchTempoWorklogs(dateFrom, dateTo, accountIds = null) {
  const token = tempoAuth();
  const all = [];
  let offset = 0;
  const limit = 1000;

  while (true) {
    const params = new URLSearchParams({ from: dateFrom, to: dateTo, limit, offset });
    const r = await fetch(`${TEMPO_URL}/worklogs?${params}`, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    });
    if (!r.ok) {
      const text = await r.text();
      throw new Error(`Tempo API ${r.status}: ${text}`);
    }
    const data = await r.json();
    const results = data.results || [];
    all.push(...results);

    const meta = data.metadata || {};
    if (!meta.next || results.length === 0) break;
    offset += results.length;
  }

  if (accountIds) {
    const ids = new Set(accountIds);
    return all.filter(wl => ids.has(wl.author?.accountId));
  }
  return all;
}
