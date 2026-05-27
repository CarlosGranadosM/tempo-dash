export const JIRA_URL = process.env.JIRA_URL || 'https://omnipro.atlassian.net';

// Equipo Omni Solutions (11 activos)
export const OMNI_TEAM_IDS = [
  '5f11d276502ce1001d2d7c73',   // Massiel Delgado
  '616f3877bcb57400680ce32e',   // Alejandro Valencia
  '628545a1ba5c2500682e3252',   // Carlos Granados
  '712020:444b8233-fddf-4c3c-bbca-8e3948c56c02', // Danna Acero Bermudez
  '6189524fc23a4f0069747831',   // Edgar Elias Peña De La Torre
  '616715e82f6aed006826d8af',   // Juan Diego Vallejo
  '712020:19aba626-02d2-4224-bc79-3523fec2cd0c', // Juan Nicolas Lemus Castro
  '600f20253b1af00069823a4d',   // Julián Valdés
  '5ffe1779642089014144e01c',   // Saulo Castillo
  '712020:e1c2d045-6f5e-4f72-ad41-a8a4f87d5803', // Vanessa Remicio
  '61f18390e4a724006ae16e2a',   // Cristian Ospina Pelaez
  '712020:c86d9f7f-1d52-48e7-8025-ae4ec0cfc8cf', // Santiago Yepes
];

// Ex-miembros con tiempo registrado — se incluyen en consultas para preservar histórico
export const OMNI_EX_MEMBERS = [
  { accountId: '712020:ec4c47db-7b77-4aad-92e6-c035c8f17724', displayName: 'Edwar Ivan Alba Jerez' },
];

export const OMNI_SOLUTIONS_GROUP_ID = '5f943d89-83fc-4906-88a9-2c9c3cb6b5e4';

export const CLOSED_STATUSES = new Set([
  'finalizado','finalizada','cancelado','cancelada','rechazado','rechazada',
  'done','done infraestructura','closed','cancelled','rejected','descartado',"won't do",
]);

export function jiraAuth() {
  const email = process.env.JIRA_EMAIL;
  const token = process.env.JIRA_API_TOKEN;
  if (!email || !token) throw new Error('Credenciales Jira no configuradas');
  return Buffer.from(`${email}:${token}`).toString('base64');
}

export async function jiraPost(path, body) {
  const auth = jiraAuth();
  const r = await fetch(`${JIRA_URL}${path}`, {
    method: 'POST',
    headers: { Authorization: `Basic ${auth}`, 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify(body),
  });
  if (!r.ok) throw new Error(`Jira ${r.status}: ${await r.text()}`);
  return r.json();
}

export async function resolveDisplayNames(accountIds) {
  if (!accountIds?.length) return {};
  const auth = jiraAuth();
  const params = accountIds.map(id => `accountId=${encodeURIComponent(id)}`).join('&');
  const r = await fetch(`${JIRA_URL}/rest/api/3/user/bulk?${params}&maxResults=200`, {
    headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' },
  });
  if (!r.ok) return {};
  const data = await r.json();
  const map = {};
  (data.values || []).forEach(u => { map[u.accountId] = u.displayName || u.accountId; });
  return map;
}

export async function jiraGet(path) {
  const auth = jiraAuth();
  const r = await fetch(`${JIRA_URL}${path}`, {
    headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' },
  });
  if (!r.ok) throw new Error(`Jira ${r.status}: ${await r.text()}`);
  return r.json();
}

export const TICKET_FIELDS = [
  'summary','description','assignee','reporter','status','issuetype','created','updated',
  'duedate','priority','timetracking','subtasks','parent','labels','project',
  'customfield_10019','customfield_10010','customfield_10014', // storyPoints, sprint, epicLink
  'customfield_10342', // Facturación
  'worklog',
];

// Extract plain text from Jira Atlassian Document Format (ADF)
export function extractAdfText(adf) {
  if (!adf) return '';
  if (typeof adf === 'string') return adf;
  const parts = [];
  function walk(node) {
    if (!node) return;
    if (node.type === 'text' && node.text) { parts.push(node.text); return; }
    if (Array.isArray(node.content)) node.content.forEach(walk);
  }
  walk(adf);
  return parts.join(' ').replace(/\s+/g, ' ').trim().slice(0, 500);
}

export function mapTicket(issue, dateFrom, dateTo, teamIds = null) {
  const f = issue.fields;
  const descriptionText = extractAdfText(f.description);
  const estS = f.timetracking?.originalEstimateSeconds ?? 0;
  const logS = f.timetracking?.timeSpentSeconds ?? 0;
  const isCl = CLOSED_STATUSES.has((f.status?.name || '').toLowerCase());
  const remS = isCl ? 0 : (f.timetracking?.remainingEstimateSeconds ?? Math.max(0, estS - logS));

  const spArr     = Array.isArray(f.customfield_10010) ? f.customfield_10010 : [];
  const activeSpr = spArr.find(s => s?.state === 'active') || spArr[spArr.length - 1];

  const wlField = f.worklog || {};
  const allWls  = wlField.worklogs || [];
  const rangeWls = allWls.filter(wl => {
    const d = (wl.started || '').split('T')[0];
    const inRange  = (!dateFrom || d >= dateFrom) && (!dateTo || d <= dateTo);
    const inTeam   = !teamIds  || teamIds.has(wl.author?.accountId);
    return inRange && inTeam;
  });

  const wlRangeSecs  = rangeWls.reduce((s, wl) => s + (wl.timeSpentSeconds || 0), 0);
  const wlAuthors    = [...new Set(rangeWls.map(wl => wl.author?.displayName).filter(Boolean))];
  const wlTotal      = wlField.total || 0;

  const wlByMonth = {};
  const wlByAuthorByMonth = {};
  rangeWls.forEach(wl => {
    const month  = (wl.started || '').slice(0, 7); // YYYY-MM
    const author = wl.author?.displayName || 'Desconocido';
    const hrs    = (wl.timeSpentSeconds || 0) / 3600;
    wlByMonth[month] = (wlByMonth[month] || 0) + hrs;
    if (!wlByAuthorByMonth[author]) wlByAuthorByMonth[author] = {};
    wlByAuthorByMonth[author][month] = (wlByAuthorByMonth[author][month] || 0) + hrs;
  });
  Object.keys(wlByMonth).forEach(m => { wlByMonth[m] = Math.round(wlByMonth[m] * 100) / 100; });
  Object.values(wlByAuthorByMonth).forEach(mMap =>
    Object.keys(mMap).forEach(m => { mMap[m] = Math.round(mMap[m] * 100) / 100; })
  );

  return {
    id:             issue.id,
    key:            issue.key,
    title:          f.summary,
    project:        f.project?.key  || issue.key.split('-')[0],
    projectName:    f.project?.name || '',
    assignee:       f.assignee?.displayName || 'Sin asignar',
    assigneeId:     f.assignee?.accountId   || null,
    status:         f.status?.name          || 'Desconocido',
    type:           f.issuetype?.name       || 'Task',
    isEpic:         (f.issuetype?.name || '').toLowerCase() === 'epic',
    isSubtask:      !!(f.parent?.fields?.issuetype?.name === 'Subtask' || f.issuetype?.subtask),
    created:        f.created?.split('T')[0] || null,
    updated:        f.updated?.split('T')[0] || null,
    dueDate:        f.duedate || null,
    priority:       f.priority?.name        || 'Medium',
    estimatedHours: Math.round(estS / 3600 * 10) / 10,
    loggedHours:    Math.round(logS  / 3600 * 10) / 10,
    remainingHours: Math.round(remS  / 3600 * 10) / 10,
    wlRangeHours:   Math.round(wlRangeSecs / 3600 * 10) / 10,
    wlRangeSecs,
    wlAuthors,
    wlTotal,
    wlByMonth,
    wlByAuthorByMonth,
    wlNeedsFullFetch: wlTotal > allWls.length,
    labels:         Array.isArray(f.labels) ? f.labels : [],
    parentKey:      f.parent?.key || f.customfield_10014 || '',
    parentSummary:  f.parent?.fields?.summary || '',
    storyPoints:    f.customfield_10019 ?? null,
    sprint:         activeSpr?.name || '',
    subtasks:       (f.subtasks || []).map(s => s.key),
    facturacion:    f.customfield_10342?.value || null,
    description:    descriptionText,
  };
}
