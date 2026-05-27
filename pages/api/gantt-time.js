import { resolveDisplayNames } from '../../lib/jira';

async function fetchViaTempoAPI(dateFrom, dateTo, ids) {
  const token = process.env.TEMPO_API_TOKEN;
  if (!token) throw new Error('TEMPO_API_TOKEN no configurado');
  const teamSet = new Set(ids);
  const all = [];
  let offset = 0;
  while (true) {
    const params = new URLSearchParams({ from: dateFrom, to: dateTo, limit: 1000, offset });
    const r = await fetch(`https://api.tempo.io/4/worklogs?${params}`, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    });
    if (!r.ok) throw new Error(`Tempo API ${r.status}: ${await r.text()}`);
    const data = await r.json();
    const results = data.results || [];
    all.push(...results.filter(wl => teamSet.has(wl.author?.accountId)));
    if (!data.metadata?.next || !results.length) break;
    offset += results.length;
  }
  return all;
}

export default async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).end();
  const { dateFrom = '2026-01-01', dateTo = '2026-03-31', userIds } = req.query;
  const ids = userIds ? userIds.split(',').map(s => s.trim()).filter(Boolean) : [];
  if (!ids.length) return res.status(400).json({ error: 'Se requiere userIds' });

  try {
    const tempoWls = await fetchViaTempoAPI(dateFrom, dateTo, ids);
    const uniqueIds = [...new Set(tempoWls.map(wl => wl.author?.accountId).filter(Boolean))];
    const nameMap = await resolveDisplayNames(uniqueIds);
    const round = v => Math.round(v * 100) / 100;

    const personData = {};
    tempoWls.forEach(wl => {
      const accountId = wl.author?.accountId || '';
      if (!accountId) return;
      const name     = nameMap[accountId] || wl.author?.displayName || accountId;
      const date     = wl.startDate || '';
      const hrs      = (wl.timeSpentSeconds || 0) / 3600;
      const issueKey = wl.issue?.key || '';
      const issueTitle = wl.issue?.summary || '';

      if (!personData[accountId]) personData[accountId] = { accountId, name, byDay: {}, byTicket: {} };
      const pd = personData[accountId];

      pd.byDay[date] = (pd.byDay[date] || 0) + hrs;

      if (issueKey) {
        if (!pd.byTicket[issueKey]) pd.byTicket[issueKey] = { key: issueKey, title: issueTitle, byDay: {} };
        pd.byTicket[issueKey].byDay[date] = (pd.byTicket[issueKey].byDay[date] || 0) + hrs;
      }
    });

    Object.values(personData).forEach(pd => {
      Object.keys(pd.byDay).forEach(d => { pd.byDay[d] = round(pd.byDay[d]); });
      Object.values(pd.byTicket).forEach(tk => {
        Object.keys(tk.byDay).forEach(d => { tk.byDay[d] = round(tk.byDay[d]); });
      });
    });

    const persons = Object.values(personData).sort((a, b) => a.name.localeCompare(b.name));
    return res.status(200).json({ persons, dateFrom, dateTo });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
