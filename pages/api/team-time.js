import { jiraPost, jiraAuth, OMNI_TEAM_IDS, TICKET_FIELDS, mapTicket, resolveDisplayNames } from '../../lib/jira';
import { getCredentials } from '../../lib/credentials';

const TEMPO_ID = '557058:295406f3-a1fc-4733-b906-dd15d021bd79';

// Fetch ALL worklogs for a ticket (paginated), filtered to date range
async function fetchRangeWorklogs(key, dateFrom, dateTo, auth, jiraUrl) {
  const all = [];
  let startAt = 0;
  while (true) {
    const r = await fetch(
      `${jiraUrl}/rest/api/3/issue/${key}/worklog?maxResults=100&startAt=${startAt}`,
      { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } }
    );
    if (!r.ok) break;
    const data = await r.json();
    const wls = data.worklogs || [];
    all.push(...wls);
    if (!wls.length || all.length >= data.total) break;
    startAt += wls.length;
  }
  return all.filter(wl => {
    const d = (wl.started || '').split('T')[0];
    return d >= dateFrom && d <= dateTo;
  });
}

function aggregateWorklogs(rangeWls) {
  const round = h => Math.round(h * 100) / 100;
  const wlRangeSecs = rangeWls.reduce((s, wl) => s + (wl.timeSpentSeconds || 0), 0);
  const byMonth = {};
  rangeWls.forEach(wl => {
    const month = (wl.started || '').slice(0, 7);
    const hrs   = (wl.timeSpentSeconds || 0) / 3600;
    byMonth[month] = (byMonth[month] || 0) + hrs;
  });
  Object.keys(byMonth).forEach(m => { byMonth[m] = round(byMonth[m]); });
  return { wlRangeSecs, wlRangeHours: round(wlRangeSecs / 3600), wlByMonth: byMonth };
}

async function fetchViaTempoAPI(dateFrom, dateTo, ids, token) {
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

function groupTempoByIssue(tempoWls, nameMap = {}) {
  const byIssue = {};
  const round = h => Math.round(h * 100) / 100;
  tempoWls.forEach(wl => {
    const id   = String(wl.issue?.id || '');
    if (!id) return;
    const accountId = wl.author?.accountId || '';
    const name = nameMap[accountId] || wl.author?.displayName || accountId || 'Desconocido';
    const secs = wl.timeSpentSeconds || 0;
    const month = (wl.startDate || '').slice(0, 7);
    if (!byIssue[id]) byIssue[id] = { secs: 0, byAuthor: {}, byMonth: {}, byAuthorByMonth: {} };
    const rec = byIssue[id];
    rec.secs += secs;
    rec.byAuthor[name] = (rec.byAuthor[name] || 0) + secs / 3600;
    rec.byMonth[month] = (rec.byMonth[month] || 0) + secs / 3600;
    if (!rec.byAuthorByMonth[name]) rec.byAuthorByMonth[name] = {};
    rec.byAuthorByMonth[name][month] = (rec.byAuthorByMonth[name][month] || 0) + secs / 3600;
  });
  Object.values(byIssue).forEach(rec => {
    Object.keys(rec.byAuthor).forEach(a => { rec.byAuthor[a] = round(rec.byAuthor[a]); });
    Object.keys(rec.byMonth).forEach(m => { rec.byMonth[m] = round(rec.byMonth[m]); });
    Object.values(rec.byAuthorByMonth).forEach(mm =>
      Object.keys(mm).forEach(m => { mm[m] = round(mm[m]); })
    );
  });
  return byIssue;
}

export default async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).end();

  const {
    dateFrom = '2026-01-01',
    dateTo   = '2026-03-31',
    userIds,
    keys,
    projectKeys,
    projectExclude: projectExcludeParam,
    facturacion: facturacionParam,
    page     = '0',
    pageSize = '100',
    all,
  } = req.query;

  const creds        = getCredentials(req);
  const ids          = userIds ? userIds.split(',').map(s => s.trim()).filter(Boolean) : OMNI_TEAM_IDS;
  const tempoToken   = creds.tempoToken;
  const useTempoAPI  = !!tempoToken;

  const projInclude    = projectKeys         ? projectKeys.split(',').map(s => s.trim()).filter(Boolean) : [];
  const projExclude    = projectExcludeParam ? projectExcludeParam.split(',').map(s => s.trim()).filter(Boolean) : [];
  const facturacionVal = facturacionParam || '';
  const returnAll      = all === 'true';
  const pageNum        = Math.max(0, parseInt(page, 10) || 0);
  const pageSz         = Math.max(1, Math.min(500, parseInt(pageSize, 10) || 100));

  function projJqlClause() {
    const parts = [];
    if (projInclude.length) parts.push(`project in (${projInclude.map(k => `"${k}"`).join(',')})`);
    if (projExclude.length) parts.push(`project not in (${projExclude.map(k => `"${k}"`).join(',')})`);
    return parts.length ? ' AND ' + parts.join(' AND ') : '';
  }

  try {
    // ── Compliance tab ────────────────────────────────────────────────────────
    if (keys) {
      const data = await jiraPost('/rest/api/3/search/jql', {
        jql: `key in (${keys})`,
        fields: TICKET_FIELDS,
        maxResults: 500,
      }, creds);

      let byIssue = {};
      if (useTempoAPI) {
        const tempoWls = await fetchViaTempoAPI(dateFrom, dateTo, ids, tempoToken);
        const uniqueIds = [...new Set(tempoWls.map(wl => wl.author?.accountId).filter(Boolean))];
        const nameMap   = await resolveDisplayNames(uniqueIds, creds);
        byIssue = groupTempoByIssue(tempoWls, nameMap);
      }

      const round = h => Math.round(h * 100) / 100;
      const tickets = (data.issues || []).map(issue => {
        const t = mapTicket(issue, dateFrom, dateTo);
        if (useTempoAPI) {
          const rec = byIssue[issue.id];
          if (rec) {
            t.wlRangeSecs       = rec.secs;
            t.wlRangeHours      = round(rec.secs / 3600);
            t.wlAuthors         = Object.keys(rec.byAuthor);
            t.wlByAuthor        = rec.byAuthor;
            t.wlByMonth         = rec.byMonth;
            t.wlByAuthorByMonth = rec.byAuthorByMonth;
          }
        }
        return t;
      });
      return res.status(200).json({ tickets, dateFrom, dateTo, mode: useTempoAPI ? 'tempo' : 'jira' });
    }

    // ── Main view ─────────────────────────────────────────────────────────────
    if (useTempoAPI) {
      const tempoWls  = await fetchViaTempoAPI(dateFrom, dateTo, ids, tempoToken);
      const uniqueIds = [...new Set(tempoWls.map(wl => wl.author?.accountId).filter(Boolean))];
      const nameMap   = await resolveDisplayNames(uniqueIds, creds);
      const byIssue   = groupTempoByIssue(tempoWls, nameMap);

      if (!Object.keys(byIssue).length) {
        return res.status(200).json({
          tickets: [], total: 0, totalHours: 0, totalPages: 0,
          page: pageNum, pageSize: pageSz,
          allProjects: [], allUsers: [],
          dateFrom, dateTo, mode: 'tempo',
        });
      }

      // Fetch Jira issues (with optional project filter in JQL)
      const issueIds = Object.keys(byIssue);
      let allIssues  = [];
      const BATCH    = 100;
      const pjClause = projJqlClause();

      for (let i = 0; i < issueIds.length; i += BATCH) {
        const batch = issueIds.slice(i, i + BATCH);
        try {
          const jql  = `id in (${batch.join(',')})${pjClause}`;
          const data = await jiraPost('/rest/api/3/search/jql', {
            jql, fields: TICKET_FIELDS, maxResults: BATCH,
          }, creds);
          allIssues.push(...(data.issues || []));
          let nextToken = data.nextPageToken || null;
          while (nextToken) {
            const pg = await jiraPost('/rest/api/3/search/jql', {
              jql, fields: TICKET_FIELDS, maxResults: BATCH, nextPageToken: nextToken,
            }, creds);
            allIssues.push(...(pg.issues || []));
            nextToken = pg.nextPageToken || null;
          }
        } catch { /* best effort */ }
      }

      const round = h => Math.round(h * 100) / 100;
      const mapped = allIssues.map(issue => {
        const t   = mapTicket(issue, null, null);
        const rec = byIssue[issue.id];
        if (rec) {
          t.wlRangeSecs       = rec.secs;
          t.wlRangeHours      = round(rec.secs / 3600);
          t.wlAuthors         = Object.keys(rec.byAuthor);
          t.wlByAuthor        = rec.byAuthor;
          t.wlByMonth         = rec.byMonth;
          t.wlByAuthorByMonth = rec.byAuthorByMonth;
          t.wlNeedsFullFetch  = false;
        }
        return t;
      });

      // Fetch all ancestor levels recursively
      const allKeys       = new Set(mapped.map(t => t.key));
      const parentTickets = [];
      let toCheck         = mapped;
      let maxRounds       = 5;

      while (maxRounds-- > 0) {
        const missing = [...new Set(toCheck.map(t => t.parentKey).filter(k => k && !allKeys.has(k)))];
        if (!missing.length) break;
        const fetched = [];
        for (let i = 0; i < missing.length; i += 50) {
          const batch = missing.slice(i, i + 50);
          try {
            const d = await jiraPost('/rest/api/3/search/jql', {
              jql: `key in (${batch.map(k => `"${k}"`).join(',')})`,
              fields: TICKET_FIELDS,
              maxResults: 100,
            }, creds);
            (d.issues || []).forEach(issue => {
              const t = mapTicket(issue, null, null);
              t.wlRangeSecs = 0; t.wlRangeHours = 0; t.wlByMonth = {}; t.wlByAuthor = {}; t.wlAuthors = [];
              fetched.push(t);
              allKeys.add(t.key);
            });
          } catch (e) { console.error('[parent-fetch] error', e.message); }
        }
        if (!fetched.length) break;
        parentTickets.push(...fetched);
        toCheck = fetched;
      }

      // Build lookup for facturación propagation
      const allByKey = new Map([...mapped, ...parentTickets].map(t => [t.key, t]));

      function resolveFacturacion(key, visited = new Set()) {
        if (!key || visited.has(key)) return null;
        visited.add(key);
        const t = allByKey.get(key);
        if (!t) return null;
        if (t.facturacion) return t.facturacion;
        return resolveFacturacion(t.parentKey, visited);
      }

      // Propagate facturación from epics to tasks that don't have it
      mapped.forEach(t => {
        if (!t.facturacion && t.parentKey) {
          t.facturacion = resolveFacturacion(t.parentKey) || null;
        }
      });

      // Dropdown metadata (accurate only on unfiltered load; frontend caches it)
      const allUsers = [...new Set(mapped.flatMap(t => t.wlAuthors || []))].sort();
      const allProjects = [...new Map(
        [...mapped, ...parentTickets].filter(t => t.project).map(t => [t.project, t.projectName || t.project])
      ).entries()].map(([key, name]) => ({ key, name })).sort((a, b) => a.name.localeCompare(b.name));

      // Filter to tickets with time; apply facturación filter
      let withTime = mapped.filter(t => t.wlRangeHours > 0);
      if (facturacionVal) {
        withTime = withTime.filter(t => t.facturacion === facturacionVal);
      }

      // Totals (pre-pagination)
      const totalHours = Math.round(withTime.reduce((s, t) => s + t.wlRangeHours, 0) * 100) / 100;
      const total      = withTime.length;
      const totalPages = returnAll ? 1 : Math.max(1, Math.ceil(total / pageSz));

      // Paginate (or return all) leaf tickets
      const pageTickets = returnAll ? withTime : withTime.slice(pageNum * pageSz, (pageNum + 1) * pageSz);

      // Include full parent chain
      const result = new Map(pageTickets.map(t => [t.key, t]));
      pageTickets.forEach(t => {
        let pk = t.parentKey;
        while (pk && !result.has(pk)) {
          const parent = allByKey.get(pk);
          if (!parent) break;
          result.set(pk, parent);
          pk = parent.parentKey;
        }
      });

      return res.status(200).json({
        tickets: [...result.values()],
        total, totalHours, totalPages,
        page: returnAll ? 0 : pageNum,
        pageSize: returnAll ? total : pageSz,
        allProjects, allUsers,
        dateFrom, dateTo, mode: 'tempo',
      });

    } else {
      // ── Jira-only fallback ────────────────────────────────────────────────
      const pjClause   = projJqlClause();
      const jql        = `worklogDate >= "${dateFrom}" AND worklogDate <= "${dateTo}"${pjClause}`;
      let allIssues    = [];
      let nextToken    = null;
      do {
        const body = { jql, fields: TICKET_FIELDS, maxResults: 100 };
        if (nextToken) body.nextPageToken = nextToken;
        const data = await jiraPost('/rest/api/3/search/jql', body, creds);
        allIssues.push(...(data.issues || []));
        nextToken = data.nextPageToken || null;
      } while (nextToken);

      const auth      = jiraAuth(creds);
      const allTickets = allIssues.map(i => mapTicket(i, dateFrom, dateTo));
      const needsFull  = allTickets.filter(t => t.wlNeedsFullFetch);
      const BATCH = 8;
      for (let i = 0; i < needsFull.length; i += BATCH) {
        await Promise.all(needsFull.slice(i, i + BATCH).map(async t => {
          try {
            const wls = await fetchRangeWorklogs(t.key, dateFrom, dateTo, auth, creds.jiraUrl);
            Object.assign(t, aggregateWorklogs(wls));
          } catch { /* best effort */ }
        }));
      }

      let withTime = allTickets.filter(t => t.wlRangeHours > 0);

      const allProjects = [...new Map(
        allTickets.filter(t => t.project).map(t => [t.project, t.projectName || t.project])
      ).entries()].map(([key, name]) => ({ key, name })).sort((a, b) => a.name.localeCompare(b.name));
      const allUsers = [...new Set(allTickets.flatMap(t => t.wlAuthors || []))].sort();

      const resultKeys     = new Set(withTime.map(t => t.key));
      const missingParents = [...new Set(withTime.map(t => t.parentKey).filter(k => k && !resultKeys.has(k)))];
      let parentTickets = [];
      for (let i = 0; i < missingParents.length; i += 50) {
        const batch = missingParents.slice(i, i + 50);
        try {
          const d = await jiraPost('/rest/api/3/search/jql', {
            jql: `key in (${batch.map(k => `"${k}"`).join(',')})`,
            fields: TICKET_FIELDS, maxResults: 100,
          }, creds);
          parentTickets.push(...(d.issues || []).map(i => mapTicket(i, dateFrom, dateTo)));
        } catch { /* best effort */ }
      }

      const totalHours = Math.round(withTime.reduce((s, t) => s + t.wlRangeHours, 0) * 100) / 100;
      const total      = withTime.length;
      const totalPages = returnAll ? 1 : Math.max(1, Math.ceil(total / pageSz));
      const pageTickets = returnAll ? withTime : withTime.slice(pageNum * pageSz, (pageNum + 1) * pageSz);
      const pageKeys    = new Set(pageTickets.map(t => t.key));
      const pageParents = parentTickets.filter(p => !pageKeys.has(p.key));

      return res.status(200).json({
        tickets: [...pageTickets, ...pageParents],
        total, totalHours, totalPages,
        page: returnAll ? 0 : pageNum,
        pageSize: returnAll ? total : pageSz,
        allProjects, allUsers,
        dateFrom, dateTo,
        mode: 'jira',
        warning: 'Modo Jira: se muestran horas totales de TODOS los equipos. Configura TEMPO_API_TOKEN para filtrar por equipo Omni Solutions.',
      });
    }
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
