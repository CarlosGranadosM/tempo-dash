import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';

// ─── Constants ────────────────────────────────────────────────────────────────
const JIRA_URL = 'https://omnipro.atlassian.net';

// ─── Credenciales (por usuario, guardadas en su navegador) ──────────────────────
const CREDS_KEY = 'tempo-creds';

function loadCreds() {
  if (typeof window === 'undefined') return {};
  try { return JSON.parse(localStorage.getItem(CREDS_KEY) || '{}'); }
  catch { return {}; }
}

function saveCreds(creds) {
  if (typeof window === 'undefined') return;
  localStorage.setItem(CREDS_KEY, JSON.stringify(creds));
}

function credHeaders() {
  const c = loadCreds();
  const h = {};
  if (c.JIRA_URL)       h['x-jira-url']    = c.JIRA_URL;
  if (c.JIRA_EMAIL)     h['x-jira-email']  = c.JIRA_EMAIL;
  if (c.JIRA_API_TOKEN) h['x-jira-token']  = c.JIRA_API_TOKEN;
  if (c.TEMPO_API_TOKEN) h['x-tempo-token'] = c.TEMPO_API_TOKEN;
  return h;
}

// Wrapper de fetch que adjunta las credenciales del usuario como headers.
function apiFetch(url, options = {}) {
  return fetch(url, { ...options, headers: { ...credHeaders(), ...(options.headers || {}) } });
}

const RESULTADO_OPTS = ['Se logró', 'Se logró parcialmente', 'No se logró', 'Se replanificó', 'Sin trabajo'];
const JUSTIF_OPTS = [
  '—', 'Cambio de prioridad', 'Atención de bugs no planificados', 'Dependencias con otros equipos',
  'Bloqueos técnicos', 'Mayor complejidad técnica', 'Tiempo en despliegues / estabilización',
  'Falta de registro de tiempo', 'Actividades transversales', 'Vacaciones / ausencias', 'Otro',
];
const RES_CLR = {
  'Se logró':               { bg: '#dcfce7', tx: '#15803d', bd: '#86efac' },
  'Se logró parcialmente':  { bg: '#fef9c3', tx: '#854d0e', bd: '#fde047' },
  'No se logró':            { bg: '#fee2e2', tx: '#b91c1c', bd: '#fca5a5' },
  'Se replanificó':         { bg: '#ede9fe', tx: '#5b21b6', bd: '#c4b5fd' },
  'Sin trabajo':            { bg: '#f1f5f9', tx: '#475569', bd: '#cbd5e1' },
};

function fmtH(h) { return h > 0 ? `${h}h` : '—'; }

const MONTH_ABBR = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
function getMonths(from, to) {
  const months = [];
  let [y, m] = from.split('-').map(Number);
  const [ty, tm] = to.split('-').map(Number);
  while (y < ty || (y === ty && m <= tm)) {
    months.push(`${y}-${String(m).padStart(2, '0')}`);
    if (++m > 12) { m = 1; y++; }
  }
  return months;
}
function fmtMonthLabel(ym) {
  const [y, mo] = ym.split('-');
  return `${MONTH_ABBR[parseInt(mo) - 1]} '${y.slice(2)}`;
}

// ─── Grouping ─────────────────────────────────────────────────────────────────
const GROUP_OPTIONS = [
  { value: 'proyecto',    label: 'Proyecto' },
  { value: 'epica',       label: 'Épica' },
  { value: 'tipo',        label: 'Tipo' },
  { value: 'estado',      label: 'Estado' },
  { value: 'quarter',     label: 'Quarter' },
  { value: 'facturacion', label: 'Facturación' },
  { value: 'persona',     label: 'Persona' },
];

const GROUP_LEVEL_LABELS = {
  proyecto:    'Proyecto',
  epica:       'Épica',
  tipo:        'Tipo',
  estado:      'Estado',
  quarter:     'Quarter',
  facturacion: 'Facturación',
  persona:     'Persona',
};

const GROUP_DEPTH_STYLES = [
  { bg: 'bg-indigo-50 hover:bg-indigo-100',   text: 'text-indigo-800',  mono: '#6366f1' },
  { bg: 'bg-blue-50 hover:bg-blue-100',       text: 'text-blue-800',    mono: '#3b82f6' },
  { bg: 'bg-violet-50 hover:bg-violet-100',   text: 'text-violet-800',  mono: '#8b5cf6' },
  { bg: 'bg-emerald-50 hover:bg-emerald-100', text: 'text-emerald-800', mono: '#059669' },
  { bg: 'bg-amber-50 hover:bg-amber-100',     text: 'text-amber-800',   mono: '#d97706' },
  { bg: 'bg-rose-50 hover:bg-rose-100',       text: 'text-rose-800',    mono: '#e11d48' },
];

// Columnas reordenables (arrastrar) + ordenables (click)
const REORDERABLE_COLS = [
  { id: 'proyecto',     label: 'Proyecto',            sortKey: 'project' },
  { id: 'ticket',       label: 'Ticket',              sortKey: 'key' },
  { id: 'tipo',         label: 'Tipo',                sortKey: 'type' },
  { id: 'titulo',       label: 'Título / Iniciativa', sortKey: 'title' },
  { id: 'descripcion',  label: 'Descripción',         sortKey: null },
  { id: 'epica',        label: 'Épica',               sortKey: null },
  { id: 'objetivo',     label: 'Objetivo',            sortKey: null },
  { id: 'personas',     label: 'Registró tiempo',     sortKey: null },
  { id: 'quarter',      label: 'Quarter',             sortKey: null },
  { id: 'estado',       label: 'Estado',              sortKey: 'status' },
  { id: 'cumplimiento', label: 'Cumpl.',              sortKey: 'status' },
  { id: 'facturacion',  label: 'Fact.',               sortKey: 'facturacion' },
];
// Columnas métricas fijas al final (solo ordenables, no arrastrables)
const FIXED_METRIC_COLS = [
  { id: 'hest',   label: 'H. Est.',  sortKey: 'totalEstH' },
  { id: 'htotal', label: 'H. Total', sortKey: 'totalHours' },
];

// Columnas para la vista de desglose (incluye Tipo de ticket)
const DETAIL_REORDERABLE_COLS = [
  { id: 'proyecto',    label: 'Proyecto',            sortKey: 'project' },
  { id: 'ticket',      label: 'Ticket',              sortKey: 'key' },
  { id: 'tipo',        label: 'Tipo',                sortKey: 'type' },
  { id: 'titulo',      label: 'Título',              sortKey: 'title' },
  { id: 'descripcion', label: 'Descripción',         sortKey: null },
  { id: 'epica',       label: 'Épica',               sortKey: null },
  { id: 'objetivo',    label: 'Objetivo',            sortKey: null },
  { id: 'personas',    label: 'Registró tiempo',     sortKey: null },
  { id: 'quarter',     label: 'Quarter',             sortKey: null },
  { id: 'estado',      label: 'Estado',              sortKey: 'status' },
  { id: 'facturacion', label: 'Fact.',               sortKey: 'facturacion' },
];

function getGroupKey(t, level) {
  if (level === 'proyecto')    return [t.project || 'Sin proyecto'];
  if (level === 'epica')       return [t._epicKey ? `${t._epicKey}\x00${t._epicTitle || t._epicKey}` : `__noepic__\x00Sin épica`];
  if (level === 'tipo')        return [t.type || 'Sin tipo'];
  if (level === 'estado')      return [t.status || 'Sin estado'];
  if (level === 'quarter') {
    const qs = [...new Set(
      Object.entries(t.wlByMonth || {})
        .filter(([, h]) => h > 0)
        .map(([ym]) => { const m = parseInt(ym.split('-')[1]); return m <= 3 ? 'Q1' : m <= 6 ? 'Q2' : m <= 9 ? 'Q3' : 'Q4'; })
    )].sort();
    return qs.length ? qs : ['Sin quarter'];
  }
  if (level === 'facturacion') return [t.facturacion || 'Sin clasificar'];
  if (level === 'persona')     return t.wlAuthors?.length ? t.wlAuthors : ['Sin asignar'];
  return ['—'];
}

function buildGroupTree(tickets, levels, depth = 0) {
  if (!levels.length) return null;
  const level = levels[0];
  const rest = levels.slice(1);
  const map = new Map();
  tickets.forEach(t => {
    getGroupKey(t, level).forEach(k => {
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(t);
    });
  });
  const groups = [];
  map.forEach((gTickets, rawKey) => {
    // For persona groups: re-tag each ticket so leaf rendering knows which author to show
    if (level === 'persona') {
      gTickets = gTickets.map(t => ({ ...t, _groupPersona: rawKey }));
    }
    let label = rawKey, linkKey = null;
    if (level === 'proyecto') {
      label = gTickets.find(t => t.projectName)?.projectName || rawKey;
    } else if (level === 'epica') {
      const [key, title] = rawKey.split('\x00');
      label = title || key;
      linkKey = key !== '__noepic__' && !key.startsWith('__') ? key : null;
    } else {
      label = rawKey;
    }
    const r = v => Math.round(v * 100) / 100;
    groups.push({
      id: `${depth}:${level}:${rawKey}`, level, depth, label, linkKey,
      totalHours: r(gTickets.reduce((s, t) => s + (t.totalHours || 0), 0)),
      totalEstH:  r(gTickets.reduce((s, t) => s + (t.totalEstH  || 0), 0)),
      tickets: gTickets,
      children: buildGroupTree(gTickets, rest, depth + 1),
    });
  });
  return groups.sort((a, b) => b.totalHours - a.totalHours);
}

function collectGroupIds(groups) {
  if (!groups) return [];
  return groups.flatMap(g => [g.id, ...collectGroupIds(g.children)]);
}

function renderLeafCell(colId, task, descriptions, descLoading) {
  const epicKey   = task._epicKey && task._epicKey !== task.key && !task._epicKey.startsWith('__') ? task._epicKey : null;
  const epicTitle = task._epicTitle && task._epicTitle !== task.title ? task._epicTitle : null;
  const desc = descriptions[task.key];
  switch (colId) {
    case 'tipo': {
      const typeColor =
        task.isEpic                     ? 'bg-purple-100 text-purple-700' :
        task.isSubtask                  ? 'bg-slate-100  text-slate-500'  :
        (task.type || '').toLowerCase() === 'bug'     ? 'bg-red-100    text-red-600'    :
        (task.type || '').toLowerCase() === 'story'   ? 'bg-blue-100   text-blue-700'   :
        (task.type || '').toLowerCase() === 'subtask' ? 'bg-slate-100  text-slate-500'  :
                                          'bg-gray-100   text-gray-500';
      return (
        <td key="tipo" className="px-3 py-1.5 whitespace-nowrap">
          <span className={`inline-flex px-1.5 py-0.5 rounded text-xs font-medium ${typeColor}`}>
            {task.type || 'Task'}
          </span>
        </td>
      );
    }
    case 'proyecto':
      return <td key="proyecto" className="px-3 py-1.5 max-w-[130px]"><div className="truncate text-gray-500 text-xs font-medium" title={task.projectName || task.project}>{task.projectName || task.project || '—'}</div></td>;
    case 'ticket':
      return <td key="ticket" className="px-3 py-1.5 whitespace-nowrap"><a href={`${JIRA_URL}/browse/${task.key}`} target="_blank" rel="noreferrer" className="text-blue-600 hover:underline font-semibold text-xs">{task.key}</a></td>;
    case 'titulo':
      return (
        <td key="titulo" className="px-3 py-1.5 max-w-[180px]">
          <div className="truncate text-gray-700 text-xs" title={task.title}>{task.title}</div>
          {task.subtaskKeys?.length > 0 && <div className="text-gray-400 text-xs">{task.subtaskKeys.length} subtareas</div>}
        </td>
      );
    case 'descripcion':
      return (
        <td key="descripcion" className="px-3 py-1.5 max-w-[200px]">
          {task.description
            ? <p className="text-gray-500 text-xs leading-snug line-clamp-2" title={task.description}>{task.description}</p>
            : <span className="text-gray-300 text-xs">—</span>}
        </td>
      );
    case 'epica':
      return (
        <td key="epica" className="px-3 py-1.5 max-w-[140px]">
          {epicKey
            ? <a href={`${JIRA_URL}/browse/${epicKey}`} target="_blank" rel="noreferrer" title={epicTitle || epicKey} className="text-violet-600 hover:underline text-xs font-medium block truncate">{epicKey}</a>
            : <span className="text-gray-300 text-xs">—</span>}
          {epicTitle && <div className="text-gray-400 text-xs truncate" title={epicTitle}>{epicTitle}</div>}
        </td>
      );
    case 'objetivo':
      return (
        <td key="objetivo" className="px-3 py-1.5 max-w-[220px]">
          {descLoading && !desc
            ? <span className="text-gray-300 text-xs animate-pulse">generando…</span>
            : desc
            ? <p className="text-gray-600 text-xs leading-snug line-clamp-3" title={desc}>{desc}</p>
            : <span className="text-gray-300 text-xs">—</span>}
        </td>
      );
    case 'personas':
      return <td key="personas" className="px-3 py-1.5 text-gray-600 text-xs leading-snug">{(task.wlAuthors || []).map(a => a.split(' ')[0]).join(', ') || '—'}</td>;
    case 'quarter': {
      const Q_STYLE = {
        Q1: 'bg-blue-100   text-blue-700',
        Q2: 'bg-green-100  text-green-700',
        Q3: 'bg-orange-100 text-orange-700',
        Q4: 'bg-purple-100 text-purple-700',
      };
      const quarters = [...new Set(
        Object.entries(task.wlByMonth || {})
          .filter(([, h]) => h > 0)
          .map(([ym]) => {
            const m = parseInt(ym.split('-')[1]);
            return m <= 3 ? 'Q1' : m <= 6 ? 'Q2' : m <= 9 ? 'Q3' : 'Q4';
          })
      )].sort();
      return (
        <td key="quarter" className="px-3 py-1.5 whitespace-nowrap">
          {quarters.length
            ? <div className="flex gap-0.5 flex-wrap">{quarters.map(q => (
                <span key={q} className={`px-1.5 py-0.5 rounded text-xs font-bold ${Q_STYLE[q]}`}>{q}</span>
              ))}</div>
            : <span className="text-gray-300 text-xs">—</span>}
        </td>
      );
    }
    case 'cumplimiento': {
      const s = (task.status || '').toLowerCase();
      const cumplido = ['done','done infraestructura','finalizado','finalizada',
                        'completado','completada','closed'].includes(s);
      return (
        <td key="cumplimiento" className="px-3 py-1.5 whitespace-nowrap">
          <span className={`inline-flex px-1.5 py-0.5 rounded text-xs font-semibold ${
            cumplido ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-600'
          }`}>
            {cumplido ? '✓ Cumplido' : '✗ No cumplido'}
          </span>
        </td>
      );
    }
    case 'estado':
      return <td key="estado" className="px-3 py-1.5"><StatusBadge s={task.status} /></td>;
    case 'facturacion':
      return <td key="facturacion" className="px-3 py-1.5"><FactBadge v={task.facturacion} /></td>;
    default:
      return <td key={colId} />;
  }
}

function renderGroupedTable(tree, leafTickets, expandedNodes, toggleNode, months, descriptions = {}, descLoading = false, columnOrder = REORDERABLE_COLS.map(c => c.id), showAllDataOnSplit = false, splitByAuthor = true) {
  const depth = leafTickets?.[0]?._depth ?? 0;

  if (!tree) {
    const indent = `${8 + depth * 20}px`;
    const rows = [];
    for (const task of (leafTickets || [])) {
      const allAuthorEntries = Object.entries(task.wlByAuthor || {}).filter(([, h]) => h > 0);
      // When inside a persona group, only show that person's row
      const authorEntries = task._groupPersona
        ? allAuthorEntries.filter(([a]) => a === task._groupPersona)
        : allAuthorEntries;
      // Split into per-author rows when: multiple authors exist, OR we're in a persona group
      const shouldSplit = splitByAuthor && (task._groupPersona ? authorEntries.length >= 1 : authorEntries.length > 1);

      if (!shouldSplit) {
        rows.push(
          <tr key={`t:${task.key}:${depth}`} className="border-b hover:bg-gray-50 bg-white">
            <td style={{ paddingLeft: indent }} className="py-1.5 w-6" />
            {columnOrder.map(colId => renderLeafCell(colId, task, descriptions, descLoading))}
            <td className="px-3 py-1.5 font-mono text-gray-400 text-xs">{fmtH(task.totalEstH)}</td>
            <td className="px-3 py-1.5 font-mono font-bold text-blue-700 text-xs">{fmtH(task.totalHours)}</td>
            {months.map(m => (
              <td key={m} className="px-3 py-1.5 font-mono text-right text-blue-600 text-xs">
                {task.wlByMonth?.[m] > 0 ? `${task.wlByMonth[m]}h` : '—'}
              </td>
            ))}
          </tr>
        );
      } else {
        authorEntries.forEach(([author, authorHours], i) => {
          const isFirst    = i === 0;
          const isLast     = i === authorEntries.length - 1;
          const authorMths = task.wlByAuthorByMonth?.[author] || {};
          rows.push(
            <tr key={`t:${task.key}:${author}:${depth}`}
              className={`hover:bg-gray-50 bg-white ${isLast ? 'border-b border-gray-200' : 'border-b border-dashed border-gray-100'}`}>
              <td style={{ paddingLeft: indent }} className="py-1 w-6 text-center text-gray-200 text-xs">{!isFirst ? '│' : ''}</td>
              {columnOrder.map(colId => {
                if (colId === 'personas') {
                  return (
                    <td key="personas" className="px-3 py-1 text-gray-700 text-xs font-medium">
                      {author.split(' ').slice(0, 2).join(' ')}
                    </td>
                  );
                }
                if (!isFirst && !showAllDataOnSplit) return <td key={colId} className="px-3 py-1" />;
                return renderLeafCell(colId, task, descriptions, descLoading);
              })}
              <td className="px-3 py-1 font-mono text-gray-400 text-xs">{isFirst ? fmtH(task.totalEstH) : ''}</td>
              <td className="px-3 py-1 font-mono font-bold text-blue-700 text-xs">{fmtH(Math.round(authorHours * 100) / 100)}</td>
              {months.map(m => {
                const mh = Math.round((authorMths[m] || 0) * 100) / 100;
                return <td key={m} className="px-3 py-1 font-mono text-right text-blue-600 text-xs">{mh > 0 ? `${mh}h` : '—'}</td>;
              })}
            </tr>
          );
        });
      }
    }
    return rows;
  }

  return tree.flatMap(group => {
    const isOpen   = expandedNodes.has(group.id);
    const groupPl  = `${8 + group.depth * 20}px`;
    const mhByMonth = months.map(m => Math.round(group.tickets.reduce((s, t) => s + (t.wlByMonth?.[m] || 0), 0) * 100) / 100);
    const authors   = [...new Set(group.tickets.flatMap(t => t.wlAuthors || []))].map(a => a.split(' ')[0]).slice(0, 5).join(', ');
    const dc = GROUP_DEPTH_STYLES[group.depth] || GROUP_DEPTH_STYLES[GROUP_DEPTH_STYLES.length - 1];
    const header = (
      <tr key={group.id} className={`border-b cursor-pointer ${dc.bg}`} onClick={() => toggleNode(group.id)}>
        <td style={{ paddingLeft: groupPl }} className={`py-2 w-6 text-xs ${dc.text}`}>{isOpen ? '▼' : '▶'}</td>
        <td colSpan={columnOrder.length} className={`px-3 py-2 ${dc.text}`}>
          <div className="flex items-center gap-2 min-w-0">
            {group.linkKey
              ? <a href={`${JIRA_URL}/browse/${group.linkKey}`} target="_blank" rel="noreferrer" onClick={e => e.stopPropagation()} className={`font-bold text-xs whitespace-nowrap hover:underline ${dc.text}`}>{group.linkKey}</a>
              : <span className="opacity-50 text-xs font-normal uppercase tracking-wider">{GROUP_LEVEL_LABELS[group.level] || group.level}</span>}
            <span className="font-semibold text-xs truncate">{group.label}</span>
            {authors && group.level !== 'proyecto' && group.level !== 'persona' && <span className="text-xs opacity-50 shrink-0">{authors}</span>}
            {group.level === 'facturacion' && <FactBadge v={group.label === 'Sin clasificar' ? null : group.label} />}
          </div>
        </td>
        <td className="px-3 py-2 font-mono text-gray-400 text-xs">{fmtH(group.totalEstH)}</td>
        <td className={`px-3 py-2 font-mono font-bold text-sm ${dc.text}`}>{fmtH(group.totalHours)}</td>
        {mhByMonth.map((mh, i) => (
          <td key={months[i]} className="px-3 py-2 font-mono text-right text-xs" style={{ color: dc.mono }}>
            {mh > 0 ? `${mh}h` : '—'}
          </td>
        ))}
      </tr>
    );
    if (!isOpen) return [header];
    const tagged = group.children === null
      ? group.tickets.map(t => ({ ...t, _depth: group.depth + 1 }))
      : null;
    return [header, ...renderGroupedTable(group.children, tagged, expandedNodes, toggleNode, months, descriptions, descLoading, columnOrder, showAllDataOnSplit, splitByAuthor)];
  });
}

function autoResultado(t) {
  const s = (t.status || '').toLowerCase();
  if (['cancelado','cancelada','rechazado','rechazada','cancelled','rejected'].some(x => s.includes(x))) return 'No se logró';
  if (['finalizado','finalizada','done','closed'].some(x => s.includes(x))) return t.wlRangeHours > 0 ? 'Se logró' : 'No se logró';
  if (t.wlRangeHours === 0) return 'Sin trabajo';
  if (t.remainingHours > 0) return 'Se logró parcialmente';
  return 'Se logró parcialmente';
}

// ─── Build hierarchy from flat ticket list (supports any depth) ───────────────
function buildHierarchy(tickets) {
  const byKey = {};
  tickets.forEach(t => { byKey[t.key] = { ...t, children: [], subtaskHours: 0, subtaskEstH: 0, subtaskKeys: [] }; });

  // Walk up the parent chain to find the nearest epic ancestor key
  function epicAncestorKey(key, depth = 0) {
    if (!key || depth > 6) return null;
    const t = byKey[key];
    if (!t) return null;
    if (t.isEpic) return key;
    return epicAncestorKey(t.parentKey, depth + 1);
  }

  const epics  = [];
  const orphans = [];

  // First pass: collect epics
  tickets.forEach(t => { if (t.isEpic) epics.push(byKey[t.key]); });

  // Second pass: place non-epics under their epic, aggregate subtasks into direct parent
  tickets.forEach(t => {
    if (t.isEpic) return;

    const directParent = t.parentKey ? byKey[t.parentKey] : null;

    if (directParent && !directParent.isEpic) {
      // Level 3+: aggregate into direct parent (story/task), which will be placed under its epic
      directParent.subtaskHours += t.wlRangeHours || 0;
      directParent.subtaskEstH  += t.estimatedHours || 0;
      directParent.subtaskKeys.push(t.key);
      // Merge author-level hour data so the parent can split by author correctly
      Object.entries(t.wlByMonth || {}).forEach(([m, h]) => {
        directParent.wlByMonth = directParent.wlByMonth || {};
        directParent.wlByMonth[m] = (directParent.wlByMonth[m] || 0) + h;
      });
      Object.entries(t.wlByAuthor || {}).forEach(([author, h]) => {
        if (!h) return;
        directParent.wlByAuthor = directParent.wlByAuthor || {};
        directParent.wlByAuthor[author] = (directParent.wlByAuthor[author] || 0) + h;
        directParent.wlByAuthorByMonth = directParent.wlByAuthorByMonth || {};
        if (!directParent.wlByAuthorByMonth[author]) directParent.wlByAuthorByMonth[author] = {};
        Object.entries(t.wlByAuthorByMonth?.[author] || {}).forEach(([m, mh]) => {
          directParent.wlByAuthorByMonth[author][m] = (directParent.wlByAuthorByMonth[author][m] || 0) + mh;
        });
        if (!(directParent.wlAuthors || []).includes(author)) {
          directParent.wlAuthors = [...(directParent.wlAuthors || []), author];
        }
      });
      return;
    }

    // Direct parent is an epic, OR we need to find the epic ancestor
    const epicKey = directParent?.isEpic
      ? t.parentKey
      : epicAncestorKey(t.parentKey || '');

    if (epicKey && byKey[epicKey]) {
      const epic = byKey[epicKey];
      if (!epic.children.some(c => c.key === t.key)) epic.children.push(byKey[t.key]);
    } else {
      orphans.push(byKey[t.key]);
    }
  });

  // Compute task totals (own hours + aggregated subtask hours)
  Object.values(byKey).forEach(t => {
    t.totalHours = (t.wlRangeHours || 0) + t.subtaskHours;
    t.totalEstH  = (t.estimatedHours || 0) + t.subtaskEstH;
  });

  // Compute epic totals + aggregate monthly breakdown from all children
  epics.forEach(e => {
    e.totalHours    = e.children.reduce((s, c) => s + (c.totalHours || 0), e.wlRangeHours || 0);
    e.totalEstH     = e.children.reduce((s, c) => s + (c.totalEstH  || 0), e.estimatedHours || 0);
    e.totalChildren = e.children.length;
    const merged = {};
    [...e.children, e].forEach(item => {
      Object.entries(item.wlByMonth || {}).forEach(([mo, h]) => {
        merged[mo] = (merged[mo] || 0) + h;
      });
    });
    e.wlByMonth = merged;

    // Propagate epic-level fields to children that don't have them set
    // (in Jira, Facturación is typically set on the epic, not on individual tasks)
    e.children.forEach(c => {
      if (!c.facturacion && e.facturacion) c.facturacion = e.facturacion;
      if (!c.project     && e.project)     c.project     = e.project;
      if (!c.projectName && e.projectName) c.projectName = e.projectName;
    });
  });

  // Group orphans under virtual "Sin Épica"
  const result = epics.filter(e => e.children.length > 0 || e.wlRangeHours > 0);
  if (orphans.filter(t => t.wlRangeHours > 0 || t.totalHours > 0).length > 0) {
    const orphanTasks = orphans.filter(t => t.wlRangeHours > 0 || t.totalHours > 0);
    result.push({
      key: '__orphan__', title: 'Sin Épica', isEpic: true, isVirtual: true,
      children: orphanTasks,
      totalHours: orphanTasks.reduce((s, t) => s + (t.totalHours || 0), 0),
      totalEstH:  orphanTasks.reduce((s, t) => s + (t.totalEstH  || 0), 0),
      wlByMonth:  {},
    });
  }
  return result;
}

// ─── Vista detalle: todos los tickets con tiempo propio (sin jerarquía) ──────
function buildDetailView(tickets) {
  const byKey = {};
  tickets.forEach(t => { byKey[t.key] = t; });

  function resolveEpic(t, depth = 0) {
    if (!t || depth > 6) return { key: null, title: null };
    if (t.isEpic) return { key: t.key, title: t.title };
    const parent = t.parentKey ? byKey[t.parentKey] : null;
    return resolveEpic(parent, depth + 1);
  }

  function resolveField(t, field, depth = 0) {
    if (!t || depth > 6) return null;
    if (t[field]) return t[field];
    const parent = t.parentKey ? byKey[t.parentKey] : null;
    return resolveField(parent, field, depth + 1);
  }

  return tickets
    .filter(t => (t.wlRangeHours || 0) > 0)
    .map(t => {
      const epic = resolveEpic(t);
      return {
        ...t,
        totalHours:  Math.round((t.wlRangeHours  || 0) * 100) / 100,
        totalEstH:   Math.round((t.estimatedHours || 0) * 100) / 100,
        _epicKey:    epic.key,
        _epicTitle:  epic.title,
        facturacion: t.facturacion || resolveField(t.parentKey ? byKey[t.parentKey] : null, 'facturacion') || null,
        project:     t.project     || resolveField(t.parentKey ? byKey[t.parentKey] : null, 'project')     || null,
        projectName: t.projectName || resolveField(t.parentKey ? byKey[t.parentKey] : null, 'projectName') || null,
      };
    });
}

// ─── App ──────────────────────────────────────────────────────────────────────
const CONFIG_FIELDS = [
  { key: 'JIRA_URL',          label: 'Jira URL',           type: 'text',     placeholder: 'https://tu-dominio.atlassian.net' },
  { key: 'JIRA_EMAIL',        label: 'Jira Email',         type: 'email',    placeholder: 'usuario@empresa.com' },
  { key: 'JIRA_API_TOKEN',    label: 'Jira API Token',     type: 'password', placeholder: 'ATATT3x…' },
  { key: 'TEMPO_API_TOKEN',   label: 'Tempo API Token',    type: 'password', placeholder: 'spSaK…' },
];

function SettingsModal({ onClose, onSaved }) {
  const [config, setConfig] = useState({});
  const [saved,  setSaved]  = useState(false);
  const [show,   setShow]   = useState({});

  useEffect(() => {
    setConfig(loadCreds());
  }, []);

  const save = () => {
    const cleaned = {};
    for (const f of CONFIG_FIELDS) {
      const v = (config[f.key] || '').trim();
      if (v) cleaned[f.key] = v;
    }
    saveCreds(cleaned);
    setSaved(true);
    setTimeout(() => setSaved(false), 2500);
    onSaved?.();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg mx-4">
        <div className="flex items-center justify-between px-6 py-4 border-b">
          <h2 className="text-base font-bold text-gray-900 flex items-center gap-2">
            <svg className="w-5 h-5 text-gray-500" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>
            </svg>
            Configuración
          </h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-xl leading-none">&times;</button>
        </div>
        <div className="px-6 py-4 space-y-3">
          <p className="text-xs text-gray-500">
            Tus credenciales se guardan <strong>solo en este navegador</strong> y se envían en cada consulta. No se almacenan en el servidor.
          </p>
          {CONFIG_FIELDS.map(f => (
            <div key={f.key}>
              <label className="block text-xs font-semibold text-gray-600 mb-1">{f.label}</label>
              <div className="relative">
                <input
                  type={f.type === 'password' && !show[f.key] ? 'password' : 'text'}
                  value={config[f.key] || ''}
                  onChange={e => setConfig(p => ({ ...p, [f.key]: e.target.value }))}
                  placeholder={f.placeholder}
                  autoComplete="off"
                  className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none pr-9"
                />
                {f.type === 'password' && (
                  <button type="button" onClick={() => setShow(p => ({ ...p, [f.key]: !p[f.key] }))}
                    className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 text-xs">
                    {show[f.key] ? '🙈' : '👁'}
                  </button>
                )}
              </div>
            </div>
          ))}

          {/* Callout: cómo obtener los tokens */}
          {(() => {
            const jiraBase = (config.JIRA_URL || 'https://id.atlassian.net').replace(/\/$/, '');
            const jiraTokenUrl  = 'https://id.atlassian.com/manage-profile/security/api-tokens';
            const tempoTokenUrl = `${jiraBase}/plugins/servlet/ac/io.tempo.jira/tempo-app#!/configuration/api-integration`;
            return (
              <div className="rounded-lg border border-blue-200 bg-blue-50 p-3 text-xs space-y-2.5">
                <p className="font-semibold text-blue-900 flex items-center gap-1.5">
                  <svg className="w-3.5 h-3.5 shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
                  ¿Cómo obtener los tokens?
                </p>
                <div>
                  <p className="font-medium text-blue-900 mb-0.5">Jira API Token</p>
                  <ol className="list-decimal list-inside space-y-0.5 text-blue-700">
                    <li>
                      Abre{' '}
                      <a href={jiraTokenUrl} target="_blank" rel="noopener noreferrer"
                        className="underline font-medium hover:text-blue-900 break-all">
                        id.atlassian.com → Security → API tokens
                      </a>
                    </li>
                    <li>Clic en <strong>Create API token</strong> → ponle un nombre → <strong>Copy</strong></li>
                  </ol>
                </div>
                <div>
                  <p className="font-medium text-blue-900 mb-0.5">Tempo API Token</p>
                  <ol className="list-decimal list-inside space-y-0.5 text-blue-700">
                    <li>
                      Abre{' '}
                      <a href={tempoTokenUrl} target="_blank" rel="noopener noreferrer"
                        className="underline font-medium hover:text-blue-900 break-all">
                        {jiraBase} → Tempo → Settings → API Integration
                      </a>
                    </li>
                    <li>Clic en <strong>New Token</strong> → ponle un nombre → <strong>Copy</strong></li>
                  </ol>
                </div>
              </div>
            );
          })()}
        </div>
        <div className="px-6 py-3 border-t flex justify-end gap-2">
          <button onClick={onClose} className="px-4 py-2 text-sm border border-gray-300 rounded-lg hover:bg-gray-50">Cerrar</button>
          <button onClick={save}
            className="px-4 py-2 text-sm bg-blue-600 text-white rounded-lg hover:bg-blue-700 flex items-center gap-1.5">
            {saved ? '✓ Guardado' : 'Guardar'}
          </button>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [tab,            setTab]            = useState('time');
  const [teams,          setTeams]          = useState([]);
  const [selectedTeamId, setSelectedTeamId] = useState('');
  const [teamIds,        setTeamIds]        = useState(new Set());
  const [teamMembers,    setTeamMembers]    = useState([]);
  const [showSettings,   setShowSettings]   = useState(false);

  const applyTeam = useCallback(async (groupId) => {
    setSelectedTeamId(groupId);
    try {
      const r = await apiFetch(`/api/teams?groupId=${encodeURIComponent(groupId)}`);
      const d = await r.json();
      const members = d.members || [];
      setTeamMembers(members);
      setTeamIds(new Set(members.map(m => m.accountId)));
    } catch {}
  }, []);

  useEffect(() => {
    apiFetch('/api/teams')
      .then(r => r.json())
      .then(d => {
        const grps = d.groups || [];
        setTeams(grps);
        // No auto-seleccionar equipo: el usuario elige primero
      })
      .catch(() => {});
  }, [applyTeam]);

  return (
    <div className="min-h-screen bg-gray-100">
      {showSettings && <SettingsModal onClose={() => setShowSettings(false)} onSaved={() => window.location.reload()} />}
      {/* Header */}
      <div className="bg-white border-b shadow-sm px-6 py-3 flex items-center gap-4">
        <div className="flex items-center gap-2">
          <span className="text-2xl">⏱</span>
          <div>
            <h1 className="font-bold text-gray-900 leading-tight">Omni Tempo</h1>
            <p className="text-xs text-gray-400">Análisis de tiempo · Equipo Omnisolutions</p>
          </div>
        </div>
        <div className="flex gap-1 ml-6 border-b-0">
          {[
            { id: 'time',    label: '📊 Tiempo Registrado' },
            { id: 'detail',  label: '🔍 Desglose de Registros' },
            { id: 'gantt',   label: '📅 Registro por Período' },
          ].map(t => (
            <button key={t.id} onClick={() => setTab(t.id)}
              className={`px-4 py-2 text-sm font-medium rounded-t transition-colors ${tab === t.id ? 'bg-blue-600 text-white' : 'text-gray-500 hover:text-gray-700 hover:bg-gray-100'}`}>
              {t.label}
            </button>
          ))}
        </div>
        <div className="ml-auto flex items-center gap-3">
          {teams.length > 0 && (
            <div className="flex items-center gap-2">
              <span className="text-xs text-gray-500 font-semibold">Equipo:</span>
              <select value={selectedTeamId} onChange={e => applyTeam(e.target.value)}
                className="border rounded px-2 py-1 text-sm">
                {teams.map(t => <option key={t.groupId} value={t.groupId}>{t.name}</option>)}
              </select>
            </div>
          )}
          <button onClick={() => setShowSettings(true)} title="Configuración"
            className="p-2 text-gray-400 hover:text-gray-700 hover:bg-gray-100 rounded-lg transition-colors">
            <svg className="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>
            </svg>
          </button>
        </div>
      </div>

      {/* Barra de equipo persistente entre pestañas */}
      {teamMembers.length > 0 && (
        <div className={`border-b px-6 py-2 flex items-center gap-2.5 flex-wrap ${teamIds.size < teamMembers.length ? 'bg-blue-50' : 'bg-white'}`}>
          <span className={`text-xs font-bold uppercase tracking-wide shrink-0 flex items-center gap-1.5 ${teamIds.size < teamMembers.length ? 'text-blue-600' : 'text-gray-400'}`}>
            Equipo
            {teamIds.size < teamMembers.length && <span className="w-1.5 h-1.5 rounded-full bg-blue-500 inline-block"/>}
          </span>
          <span className="text-xs text-gray-400">({teamIds.size}/{teamMembers.length})</span>
          <button onClick={() => setTeamIds(new Set(teamMembers.map(m => m.accountId)))} className="text-xs text-blue-600 hover:underline">Todos</button>
          <button onClick={() => setTeamIds(new Set())} className="text-xs text-gray-400 hover:underline">Ninguno</button>
          <span className="w-px h-4 bg-gray-200 mx-1"/>
          {teamMembers.map(m => {
            const active = teamIds.has(m.accountId);
            return (
              <button key={m.accountId}
                onClick={() => setTeamIds(prev => { const n = new Set(prev); active ? n.delete(m.accountId) : n.add(m.accountId); return n; })}
                title={m.displayName || m.name}
                className={`px-2 py-0.5 rounded text-xs border transition-colors ${
                  m.isExMember
                    ? active ? 'bg-amber-100 text-amber-700 border-amber-300' : 'bg-gray-100 text-gray-400 border-gray-200 line-through'
                    : active ? 'bg-blue-100 text-blue-700 border-blue-300' : 'bg-gray-100 text-gray-400 border-gray-200 line-through'
                }`}>
                {(m.displayName || m.name || '').split(' ').slice(0, 2).join(' ')}{m.isExMember ? ' ex' : ''}
              </button>
            );
          })}
        </div>
      )}

      <div className="max-w-screen-2xl mx-auto px-4 py-5">
        {tab === 'time'    && <TimeTab       teamIds={teamIds} setTeamIds={setTeamIds} teamMembers={teamMembers} />}
        {tab === 'detail'  && <DetailTimeTab teamIds={teamIds} setTeamIds={setTeamIds} teamMembers={teamMembers} />}
        {tab === 'gantt'   && <GanttTimeTab  teamIds={teamIds} teamMembers={teamMembers} />}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// TAB 1: TIEMPO REGISTRADO
// ══════════════════════════════════════════════════════════════════════════════
function TimeTab({ teamIds, setTeamIds, teamMembers }) {
  const [dateFrom,   setDateFrom]   = useState('2026-01-01');
  const [dateTo,     setDateTo]     = useState('2026-03-31');
  const [data,       setData]       = useState(null);
  const [hierarchy,  setHierarchy]  = useState([]);
  const [loadPct,    setLoadPct]    = useState(null);
  const [error,      setError]      = useState('');
  const [search,      setSearch]      = useState('');
  const [groupBy,     setGroupBy]     = useState(['proyecto']);
  const [columnOrder, setColumnOrder] = useState(REORDERABLE_COLS.map(c => c.id));
  const [sortCol,     setSortCol]     = useState('htotal');
  const [sortDir,     setSortDir]     = useState('desc');
  const dragCol = useRef(null);
  const [expandedNodes,      setExpandedNodes]      = useState(new Set());
  const [projectInclude,     setProjectInclude]     = useState([]);
  const [projectExclude,     setProjectExclude]     = useState([]);
  const [personInclude,      setPersonInclude]      = useState([]);
  const [personExclude,      setPersonExclude]      = useState([]);
  const [facturacionFilter,  setFacturacionFilter]  = useState('');
  const [colFilters,         setColFilters]         = useState({});
  const [showFilters,        setShowFilters]        = useState(false);
  const [showPersonCards,    setShowPersonCards]    = useState(true);
  const [showFacturacion,    setShowFacturacion]    = useState(true);
  const [cachedMeta,         setCachedMeta]         = useState(null);
  const [descriptions,       setDescriptions]       = useState({});
  const [descLoading,        setDescLoading]        = useState(false);
  const [exportPct,          setExportPct]          = useState(null);
  const hasLoadedRef  = useRef(false);
  const mountedRef    = useRef(false);
  const loadRef       = useRef(null);

  const resetView = () => { setData(null); setHierarchy([]); setError(''); };

  const teamIdsKey = [...teamIds].sort().join(',');
  useEffect(() => { if (hasLoadedRef.current) resetView(); }, [dateFrom, dateTo]);
  useEffect(() => {
    if (!mountedRef.current) { mountedRef.current = true; return; }
    if (teamIds.size > 0) loadRef.current?.();
  }, [teamIdsKey]);

  const load = async () => {
    if (loadPct !== null) return;
    setData(null); setHierarchy([]);
    setLoadPct(5); setError('');
    const timer = setInterval(() => setLoadPct(p => p !== null && p < 85 ? Math.min(85, p + 8 + Math.random() * 7) : p), 600);
    try {
      const ids = [...teamIds].join(',');
      const params = new URLSearchParams({ dateFrom, dateTo, userIds: ids, all: 'true' });
      const r = await apiFetch(`/api/team-time?${params}`);
      const d = await r.json();
      if (!r.ok) throw new Error(d.error);
      hasLoadedRef.current = true;
      setLoadPct(100);
      setData(d);
      setHierarchy(buildHierarchy(d.tickets || []));
      setProjectInclude([]); setProjectExclude([]);
      setPersonInclude([]);  setPersonExclude([]);
      setFacturacionFilter('');
      setColFilters({});
      setSearch('');
      if (d.allProjects) setCachedMeta({ projects: d.allProjects, users: d.allUsers || [] });
      if (d.warning) console.warn('[omni-tempo]', d.warning);
    } catch (e) { setError(e.message); }
    finally     { clearInterval(timer); setTimeout(() => setLoadPct(null), 400); }
  };
  loadRef.current = load;

  const loadDescriptions = async (tickets, force = false) => {
    const candidates = tickets.filter(t => t.title);
    const toLoad = force ? candidates : candidates.filter(t => !descriptions[t.key]);
    if (!toLoad.length) return;
    setDescLoading(true);
    const BATCH = 75;
    try {
      for (let i = 0; i < toLoad.length; i += BATCH) {
        const batch = toLoad.slice(i, i + BATCH);
        const r = await apiFetch('/api/summarize', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tickets: batch.map(t => ({ key: t.key, title: t.title, type: t.type, description: t.description || '' })) }),
        });
        const d = await r.json();
        if (d.error) throw new Error(d.error);
        if (d.descriptions) setDescriptions(prev => ({ ...prev, ...d.descriptions }));
      }
    } catch (err) { setError('Gemini: ' + err.message); }
    finally { setDescLoading(false); }
  };

  const toggleNode  = id  => setExpandedNodes(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const expandAll   = () => setExpandedNodes(new Set(collectGroupIds(groupTree)));
  const collapseAll = () => setExpandedNodes(new Set());

  useEffect(() => { setExpandedNodes(new Set()); }, [groupBy]);

  const filteredHierarchy = hierarchy.map(epic => ({
    ...epic,
    children: (epic.children || []).filter(t => {
      if (projectInclude.length > 0 && !projectInclude.includes(t.project)) return false;
      if (projectExclude.length > 0 &&  projectExclude.includes(t.project)) return false;
      if (facturacionFilter && t.facturacion !== facturacionFilter) return false;
      const cfTipo   = colFilters.tipo;
      const cfEstado = colFilters.estado;
      const cfQuart  = colFilters.quarter;
      if (cfTipo?.length   > 0 && !cfTipo.includes(t.type   || 'Sin tipo'))   return false;
      if (cfEstado?.length > 0 && !cfEstado.includes(t.status || 'Sin estado')) return false;
      if (cfQuart?.length  > 0) {
        const qs = new Set(Object.entries(t.wlByMonth || {}).filter(([, h]) => h > 0).map(([ym]) => {
          const mo = parseInt(ym.split('-')[1]); return mo <= 3 ? 'Q1' : mo <= 6 ? 'Q2' : mo <= 9 ? 'Q3' : 'Q4';
        }));
        if (!cfQuart.some(q => qs.has(q))) return false;
      }
      if (personInclude.length > 0 && !personInclude.some(p => (t.wlAuthors || []).includes(p))) return false;
      if (personExclude.length > 0 &&  personExclude.some(p => (t.wlAuthors || []).includes(p))) return false;
      if (search) { const q = search.toLowerCase(); return t.key.toLowerCase().includes(q) || t.title.toLowerCase().includes(q); }
      return true;
    }),
  })).filter(e => {
    if (search) { const q = search.toLowerCase(); return e.key.toLowerCase().includes(q) || e.title.toLowerCase().includes(q) || e.children.length > 0; }
    if (personInclude.length > 0) return e.children.length > 0 || personInclude.some(p => (e.wlAuthors || []).includes(p));
    return true;
  });

  const months = getMonths(dateFrom, dateTo);
  // Use cached metadata for dropdowns so filters don't lose options after re-query
  const allProjectsInfo = cachedMeta?.projects || (data
    ? [...new Map(data.tickets.map(t => [t.project, t.projectName || t.project])).entries()]
        .filter(([k]) => k)
        .sort((a, b) => (a[1] || a[0]).localeCompare(b[1] || b[0]))
        .map(([key, name]) => ({ key, name: name || key }))
    : []);
  const allUsers = cachedMeta?.users || (data ? [...new Set(data.tickets.flatMap(t => t.wlAuthors || []))] : []);

  const facturacionStats = (() => {
    let facturables = 0, noFacturable = 0, sinClasif = 0;
    filteredHierarchy.forEach(e => {
      (e.children || []).forEach(t => {
        const h = t.totalHours || 0;
        if (t.facturacion === 'Facturables') facturables += h;
        else if (t.facturacion === 'No facturable') noFacturable += h;
        else sinClasif += h;
      });
    });
    const r = v => Math.round(v * 100) / 100;
    return { facturables: r(facturables), noFacturable: r(noFacturable), sinClasif: r(sinClasif) };
  })();

  const flatTickets = filteredHierarchy.flatMap(epic => {
    const tasks = (epic.children || []).filter(t => (t.totalHours || 0) > 0 || (t.wlRangeHours || 0) > 0);
    if (tasks.length > 0) return tasks.map(t => ({ ...t, _epicKey: epic.key, _epicTitle: epic.title }));
    // Epic has its own logged hours — include it but only count its direct hours (not children)
    if ((epic.wlRangeHours || 0) > 0) {
      return [{ ...epic, _epicKey: epic.key, _epicTitle: epic.title, totalHours: epic.wlRangeHours }];
    }
    return [];
  });

  const totalH   = Math.round(flatTickets.reduce((s, t) => s + (t.totalHours || 0), 0) * 100) / 100;
  const totalTix = flatTickets.length;

  // Unique values for column filters — always from full loaded dataset
  const allTipos   = data ? [...new Set((data.tickets || []).map(t => t.type   || 'Sin tipo'))].sort()   : [];
  const allEstados = data ? [...new Set((data.tickets || []).map(t => t.status || 'Sin estado'))].sort() : [];

  const personData = data ? (() => {
    const byPerson = {};
    flatTickets.forEach(t => {
      Object.entries(t.wlByAuthor || {}).forEach(([name]) => {
        const monthlyHrs = months.reduce((s, m) => s + (t.wlByAuthorByMonth?.[name]?.[m] || 0), 0);
        if (monthlyHrs <= 0) return;
        if (!byPerson[name]) byPerson[name] = { hrs: 0, ticketMap: new Map() };
        byPerson[name].hrs += monthlyHrs;
        if (!byPerson[name].ticketMap.has(t.key)) {
          byPerson[name].ticketMap.set(t.key, { ...t, hoursForPerson: Math.round(monthlyHrs * 100) / 100 });
        }
      });
    });
    return Object.entries(byPerson)
      .filter(([, d]) => d.hrs > 0)
      .map(([name, d]) => {
        const tickets = [...d.ticketMap.values()].sort((a, b) => b.hoursForPerson - a.hoursForPerson);
        return {
          name,
          totalHours: Math.round(d.hrs * 100) / 100,
          tickets,
          projects: [...new Set(tickets.map(t => t.project))].join(', '),
        };
      })
      .sort((a, b) => b.totalHours - a.totalHours);
  })() : [];

  const sortedTickets = (() => {
    const col = [...REORDERABLE_COLS, ...FIXED_METRIC_COLS].find(c => c.id === sortCol);
    if (!col?.sortKey) return flatTickets;
    return [...flatTickets].sort((a, b) => {
      const av = a[col.sortKey] ?? '', bv = b[col.sortKey] ?? '';
      const cmp = (typeof av === 'number' && typeof bv === 'number') ? av - bv : String(av).localeCompare(String(bv));
      return sortDir === 'asc' ? cmp : -cmp;
    });
  })();

  const groupTree = buildGroupTree(sortedTickets, groupBy);

  const exportExcel = async () => {
    try {
      setExportPct(50);
      const r = await apiFetch('/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type: 'excel', data: { tickets: flatTickets, months, descriptions, dateFrom, dateTo, tabName: 'tiempo' } }),
      });
      setExportPct(95);
      const blob = await r.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a'); a.href = url; a.download = `omni-tiempo-${dateFrom}.xlsx`; a.click();
    } finally {
      setExportPct(null);
    }
  };

  return (
    <div className="space-y-4">
      <p className="text-xs text-gray-400 leading-relaxed">
        Agrupa los tickets con tiempo registrado en Tempo por épica o proyecto. Muestra horas estimadas vs registradas por mes y permite filtrar por equipo, facturación, tipo, estado y quarter. Usa <strong className="text-gray-500">Consultar Jira</strong> para cargar; los filtros y agrupaciones se aplican localmente sin nuevas consultas.
      </p>
      {/* Controls */}
      <div className="bg-white rounded-lg shadow-sm p-4 space-y-3">

        {/* 1. Fecha */}
        <div className="flex flex-wrap gap-3 items-end">
          <div>
            <label className="block text-xs font-semibold text-gray-500 mb-1">Desde</label>
            <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} disabled={teamIds.size === 0} className="border rounded px-2 py-1.5 text-sm disabled:bg-gray-100 disabled:text-gray-400 disabled:cursor-not-allowed" />
          </div>
          <div>
            <label className="block text-xs font-semibold text-gray-500 mb-1">Hasta</label>
            <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} disabled={teamIds.size === 0} className="border rounded px-2 py-1.5 text-sm disabled:bg-gray-100 disabled:text-gray-400 disabled:cursor-not-allowed" />
          </div>
          <button onClick={load} disabled={loadPct !== null || teamIds.size === 0}
            className={`px-5 py-2 rounded-lg text-sm font-medium flex items-center gap-2 min-w-[160px] justify-center transition-colors ${teamIds.size === 0 ? 'bg-gray-200 text-gray-400 cursor-not-allowed' : loadPct !== null ? 'bg-blue-500 text-white cursor-not-allowed' : 'bg-blue-600 text-white hover:bg-blue-700'}`}>
            {loadPct !== null
              ? <><svg className="w-4 h-4 shrink-0 animate-spin" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><circle cx="12" cy="12" r="10" strokeDasharray="40" strokeDashoffset="15"/></svg><span>Consultando…</span></>
              : '🔍 Consultar Jira'}
          </button>
          {teamIds.size === 0 && <span className="text-xs text-amber-600 font-medium">← Selecciona un equipo primero</span>}
          {data && (
            <>
              <select value={groupBy[0] || 'proyecto'} onChange={e => setGroupBy([e.target.value])}
                className="border rounded px-2 py-1.5 text-xs text-gray-600 bg-white">
                {GROUP_OPTIONS.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
              </select>
              <input placeholder="Buscar ticket…" value={search} onChange={e => setSearch(e.target.value)} className="border rounded px-2 py-1.5 text-sm w-40" />
              <div className="ml-auto flex gap-2 items-center">
                <button onClick={expandAll}   className="px-3 py-1.5 text-xs border rounded hover:bg-gray-50">Expandir</button>
                <button onClick={collapseAll} className="px-3 py-1.5 text-xs border rounded hover:bg-gray-50">Contraer</button>
                <button onClick={exportExcel} disabled={exportPct !== null}
                  className="px-4 py-1.5 bg-green-600 text-white text-sm rounded hover:bg-green-700 disabled:opacity-75 flex items-center gap-1.5 min-w-[90px] justify-center">
                  {exportPct !== null
                    ? <><svg className="w-3.5 h-3.5 animate-spin shrink-0" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" strokeDasharray="40" strokeDashoffset="10"/></svg>{exportPct}%</>
                    : '⬇ Excel'}
                </button>
              </div>
            </>
          )}
        </div>

        {/* Toggle filtros */}
        {data && (() => {
          const n = projectInclude.length + projectExclude.length + personInclude.length + personExclude.length + (facturacionFilter ? 1 : 0) + Object.values(colFilters).reduce((s, a) => s + (a?.length || 0), 0);
          return (
            <button onClick={() => setShowFilters(v => !v)}
              className={`flex items-center gap-1.5 text-xs pt-1 ${n > 0 ? 'text-blue-600 font-semibold' : 'text-gray-400'} hover:text-blue-500`}>
              <span>{showFilters ? '▼' : '▶'}</span>
              Filtros
              {n > 0 && <span className="bg-blue-500 text-white rounded-full px-1.5 py-0.5 text-xs leading-none">{n}</span>}
            </button>
          );
        })()}

        {showFilters && <>

        {/* 2. Proyecto */}
        {data && (
          <div className={`flex items-start gap-3 pt-2 border-t ${projectInclude.length > 0 || projectExclude.length > 0 ? 'bg-blue-50 -mx-4 px-4 rounded' : ''}`}>
            <span className={`text-xs font-bold uppercase tracking-wide w-16 mt-1.5 shrink-0 flex items-center gap-1 ${projectInclude.length > 0 || projectExclude.length > 0 ? 'text-blue-600' : 'text-gray-400'}`}>
              Proyecto
              {(projectInclude.length > 0 || projectExclude.length > 0) && <span className="w-1.5 h-1.5 rounded-full bg-blue-500 inline-block shrink-0" />}
            </span>
            <div className="flex flex-wrap gap-x-6 gap-y-1.5 flex-1">
              <div className="flex flex-wrap gap-1 items-center">
                <span className="text-xs font-semibold text-green-600 shrink-0 mr-0.5">Incluir</span>
                {projectInclude.map(p => {
                  const info = allProjectsInfo.find(i => i.key === p);
                  return (
                    <span key={p} className="inline-flex items-center bg-green-100 text-green-700 border border-green-300 px-1.5 py-0.5 rounded text-xs">
                      {info?.name || p}<button onClick={() => setProjectInclude(prev => prev.filter(x => x !== p))} className="ml-1 leading-none hover:text-red-600">×</button>
                    </span>
                  );
                })}
                <select value="" onChange={e => { if (e.target.value && !projectInclude.includes(e.target.value)) setProjectInclude(prev => [...prev, e.target.value]); e.target.value = ''; }} className="border rounded px-1.5 py-0.5 text-xs text-gray-500">
                  <option value="">+ añadir</option>
                  {allProjectsInfo.filter(p => !projectInclude.includes(p.key)).map(p => <option key={p.key} value={p.key}>{p.name}</option>)}
                </select>
              </div>
              <div className="flex flex-wrap gap-1 items-center">
                <span className="text-xs font-semibold text-red-500 shrink-0 mr-0.5">Excluir</span>
                {projectExclude.map(p => {
                  const info = allProjectsInfo.find(i => i.key === p);
                  return (
                    <span key={p} className="inline-flex items-center bg-red-100 text-red-600 border border-red-300 px-1.5 py-0.5 rounded text-xs">
                      {info?.name || p}<button onClick={() => setProjectExclude(prev => prev.filter(x => x !== p))} className="ml-1 leading-none hover:text-red-800">×</button>
                    </span>
                  );
                })}
                <select value="" onChange={e => { if (e.target.value && !projectExclude.includes(e.target.value)) setProjectExclude(prev => [...prev, e.target.value]); e.target.value = ''; }} className="border rounded px-1.5 py-0.5 text-xs text-gray-500">
                  <option value="">+ añadir</option>
                  {allProjectsInfo.filter(p => !projectExclude.includes(p.key)).map(p => <option key={p.key} value={p.key}>{p.name}</option>)}
                </select>
              </div>
            </div>
          </div>
        )}

        {/* 3. Facturación */}
        {data && (
          <div className={`flex items-start gap-3 pt-2 border-t ${facturacionFilter ? 'bg-blue-50 -mx-4 px-4 rounded' : ''}`}>
            <span className={`text-xs font-bold uppercase tracking-wide w-16 mt-1.5 shrink-0 flex items-center gap-1 ${facturacionFilter ? 'text-blue-600' : 'text-gray-400'}`}>
              Facturac.
              {facturacionFilter && <span className="w-1.5 h-1.5 rounded-full bg-blue-500 inline-block shrink-0" />}
            </span>
            <div className="flex gap-1 items-center mt-1">
              {['', 'Facturables', 'No facturable'].map(v => (
                <button key={v} onClick={() => setFacturacionFilter(v)} className={`px-2.5 py-0.5 rounded text-xs border transition-colors ${
                    facturacionFilter === v
                      ? v === 'Facturables'    ? 'bg-green-100 text-green-700 border-green-400 font-semibold'
                      : v === 'No facturable'  ? 'bg-orange-100 text-orange-700 border-orange-400 font-semibold'
                      : 'bg-blue-100 text-blue-700 border-blue-300 font-semibold'
                      : 'bg-white text-gray-500 border-gray-200 hover:bg-gray-50'
                  }`}>
                  {v === '' ? 'Todos' : v}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* 4. Personas */}
        {data && (
          <div className={`flex items-start gap-3 pt-2 border-t ${personInclude.length > 0 || personExclude.length > 0 ? 'bg-blue-50 -mx-4 px-4 rounded' : ''}`}>
            <span className={`text-xs font-bold uppercase tracking-wide w-16 mt-1.5 shrink-0 flex items-center gap-1 ${personInclude.length > 0 || personExclude.length > 0 ? 'text-blue-600' : 'text-gray-400'}`}>
              Personas
              {(personInclude.length > 0 || personExclude.length > 0) && <span className="w-1.5 h-1.5 rounded-full bg-blue-500 inline-block shrink-0" />}
            </span>
            <div className="flex flex-wrap gap-x-6 gap-y-1.5 flex-1">
              <div className="flex flex-wrap gap-1 items-center">
                <span className="text-xs font-semibold text-green-600 shrink-0 mr-0.5">Incluir</span>
                {personInclude.map(p => (
                  <span key={p} className="inline-flex items-center bg-green-100 text-green-700 border border-green-300 px-1.5 py-0.5 rounded text-xs">
                    {p.split(' ').slice(0, 2).join(' ')}<button onClick={() => setPersonInclude(prev => prev.filter(x => x !== p))} className="ml-1 leading-none hover:text-red-600">×</button>
                  </span>
                ))}
                <select value="" onChange={e => { if (e.target.value) { setPersonInclude(prev => prev.includes(e.target.value) ? prev : [...prev, e.target.value]); e.target.value = ''; } }}
                  className="border rounded px-1.5 py-0.5 text-xs text-gray-500">
                  <option value="">+ añadir</option>
                  {allUsers.filter(u => !personInclude.includes(u)).sort().map(u => <option key={u} value={u}>{u}</option>)}
                </select>
              </div>
              <div className="flex flex-wrap gap-1 items-center">
                <span className="text-xs font-semibold text-red-500 shrink-0 mr-0.5">Excluir</span>
                {personExclude.map(p => (
                  <span key={p} className="inline-flex items-center bg-red-100 text-red-600 border border-red-300 px-1.5 py-0.5 rounded text-xs">
                    {p.split(' ').slice(0, 2).join(' ')}<button onClick={() => setPersonExclude(prev => prev.filter(x => x !== p))} className="ml-1 leading-none hover:text-red-800">×</button>
                  </span>
                ))}
                <select value="" onChange={e => { if (e.target.value) { setPersonExclude(prev => prev.includes(e.target.value) ? prev : [...prev, e.target.value]); e.target.value = ''; } }}
                  className="border rounded px-1.5 py-0.5 text-xs text-gray-500">
                  <option value="">+ añadir</option>
                  {allUsers.filter(u => !personExclude.includes(u)).sort().map(u => <option key={u} value={u}>{u}</option>)}
                </select>
              </div>
            </div>
          </div>
        )}

        {/* 5. Filtros de columna */}
        {data && (
          <div className={`flex flex-wrap items-center gap-2 pt-2 border-t ${Object.values(colFilters).some(a => a?.length > 0) ? 'bg-blue-50 -mx-4 px-4 rounded' : ''}`}>
            <span className={`text-xs font-bold uppercase tracking-wide w-16 shrink-0 ${Object.values(colFilters).some(a => a?.length > 0) ? 'text-blue-600' : 'text-gray-400'}`}>Columnas</span>
            {[
              { id: 'tipo',    label: 'Tipo',    opts: allTipos },
              { id: 'estado',  label: 'Estado',  opts: allEstados },
              { id: 'quarter', label: 'Quarter', opts: ['Q1', 'Q2', 'Q3', 'Q4'] },
            ].map(({ id, label, opts }) => {
              const active = colFilters[id] || [];
              return (
                <div key={id} className="flex items-center gap-1 flex-wrap">
                  {active.map(v => (
                    <span key={v} className="inline-flex items-center bg-violet-100 text-violet-700 border border-violet-300 px-1.5 py-0.5 rounded text-xs">
                      {label}: {v}
                      <button onClick={() => setColFilters(p => ({ ...p, [id]: (p[id] || []).filter(x => x !== v) }))} className="ml-1 leading-none hover:text-red-600">×</button>
                    </span>
                  ))}
                  {opts.filter(v => !active.includes(v)).length > 0 && (
                    <select value="" onChange={e => { if (e.target.value) { setColFilters(p => ({ ...p, [id]: [...(p[id] || []), e.target.value] })); e.target.value = ''; } }}
                      className="border rounded px-1.5 py-0.5 text-xs text-gray-500 bg-white">
                      <option value="">+ {label}</option>
                      {opts.filter(v => !active.includes(v)).map(v => <option key={v} value={v}>{v}</option>)}
                    </select>
                  )}
                </div>
              );
            })}
            {Object.values(colFilters).some(a => a?.length > 0) && (
              <button onClick={() => setColFilters({})} className="text-xs text-red-400 hover:text-red-600 ml-1">Limpiar</button>
            )}
          </div>
        )}

        </>}
      </div>

      {error && <div className="bg-red-50 border border-red-200 rounded p-3 text-sm text-red-700">{error}</div>}
      {data?.mode === 'jira' && (
        <div className="bg-amber-50 border border-amber-200 rounded p-3 text-sm text-amber-800">
          <strong>Modo Jira:</strong> Se muestran horas de <strong>todos los equipos</strong> (Tempo registra el tiempo bajo su propia cuenta de servicio, no por persona). Para ver solo las horas del equipo Omni Solutions con desglose por persona, agrega <code className="bg-amber-100 px-1 rounded">TEMPO_API_TOKEN</code> en <code className="bg-amber-100 px-1 rounded">.env.local</code>.
        </div>
      )}

      {/* Summary cards */}
      {data && (
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          {[
            { label: 'Horas registradas', val: `${totalH}h`, color: '#3b82f6' },
            { label: 'Tickets con tiempo', val: totalTix, color: '#10b981' },
            { label: 'Épicas', val: hierarchy.filter(e => !e.isVirtual).length, color: '#8b5cf6' },
            { label: 'Proyectos', val: allProjectsInfo.length || new Set((data.tickets || []).map(t => t.project)).size, color: '#f59e0b' },
          ].map(c => (
            <div key={c.label} className="bg-white rounded-lg shadow-sm p-3 border-l-4" style={{ borderColor: c.color }}>
              <div className="text-xs text-gray-400">{c.label}</div>
              <div className="text-2xl font-bold" style={{ color: c.color }}>{c.val}</div>
            </div>
          ))}
        </div>
      )}

      {/* Facturación breakdown */}
      {data && (facturacionStats.facturables > 0 || facturacionStats.noFacturable > 0) && (
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          <button onClick={() => setShowFacturacion(v => !v)}
            className="w-full flex items-center gap-2 px-3 py-2 hover:bg-gray-50 transition-colors">
            <span className="text-gray-400 text-xs">{showFacturacion ? '▼' : '▶'}</span>
            <span className="text-xs font-bold text-gray-400 uppercase tracking-wide">Facturación</span>
          </button>
        {showFacturacion && <div className="px-3 pb-3 flex flex-wrap gap-6 items-center border-t pt-2">
          <span />
          <div className="flex items-center gap-1.5 cursor-pointer" onClick={() => setFacturacionFilter(facturacionFilter === 'Facturables' ? '' : 'Facturables')}>
            <div className="w-2.5 h-2.5 rounded-full bg-green-500" />
            <span className="text-xs text-gray-500">Facturables</span>
            <span className="font-bold text-sm text-green-700">{facturacionStats.facturables}h</span>
          </div>
          <div className="flex items-center gap-1.5 cursor-pointer" onClick={() => setFacturacionFilter(facturacionFilter === 'No facturable' ? '' : 'No facturable')}>
            <div className="w-2.5 h-2.5 rounded-full bg-orange-400" />
            <span className="text-xs text-gray-500">No facturable</span>
            <span className="font-bold text-sm text-orange-600">{facturacionStats.noFacturable}h</span>
          </div>
          {facturacionStats.sinClasif > 0 && (
            <div className="flex items-center gap-1.5">
              <div className="w-2.5 h-2.5 rounded-full bg-gray-300" />
              <span className="text-xs text-gray-400">Sin clasificar</span>
              <span className="font-bold text-sm text-gray-400">{facturacionStats.sinClasif}h</span>
            </div>
          )}
          {totalH > 0 && (
            <div className="ml-auto flex items-center gap-2 text-xs text-gray-400">
              <div className="w-32 h-2 rounded-full bg-gray-100 overflow-hidden flex">
                <div className="h-full bg-green-400" style={{ width: `${Math.round(facturacionStats.facturables / totalH * 100)}%` }} />
                <div className="h-full bg-orange-300" style={{ width: `${Math.round(facturacionStats.noFacturable / totalH * 100)}%` }} />
              </div>
              <span>{totalH > 0 ? Math.round(facturacionStats.facturables / totalH * 100) : 0}% facturable</span>
            </div>
          )}
        </div>}
        </div>
      )}

      {/* Person summary cards */}
      {data && personData.length > 0 && (
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          <button onClick={() => setShowPersonCards(v => !v)}
            className="w-full flex items-center gap-2 px-4 py-2.5 hover:bg-gray-50 transition-colors">
            <span className="text-gray-400 text-xs">{showPersonCards ? '▼' : '▶'}</span>
            <span className="text-xs font-semibold text-gray-400 uppercase tracking-wide">Horas por persona · período</span>
          </button>
          {showPersonCards && (
            <div className="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-7 gap-2 px-4 pb-4 border-t pt-3">
              {personData.map(p => (
                <button key={p.name}
                  onClick={() => { setPersonInclude(prev => prev.includes(p.name) ? prev.filter(x => x !== p.name) : [...prev, p.name]); }}
                  title={p.name}
                  className={`p-2 rounded-lg border text-left transition-all ${personInclude.includes(p.name) ? 'border-blue-500 bg-blue-50 shadow-sm' : 'border-gray-200 hover:border-blue-300 hover:bg-blue-50'}`}>
                  <div className="font-mono font-bold text-lg text-blue-700 leading-none">{p.totalHours}h</div>
                  <div className="text-xs font-medium text-gray-800 mt-1 truncate">{p.name.split(' ').slice(0, 2).join(' ')}</div>
                  <div className="text-xs text-gray-400">{p.tickets.length} tickets</div>
                </button>
              ))}
            </div>
          )}
        </div>
      )}


      {/* Hierarchy table */}
      {data && (
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead className="bg-gray-50 border-b sticky top-0 z-10">
              <tr>
                <th className="px-3 py-2 w-6" />
                {columnOrder.map(colId => {
                  const col = REORDERABLE_COLS.find(c => c.id === colId);
                  if (!col) return null;
                  const isSorted = sortCol === col.id;
                  return (
                    <th key={col.id}
                      draggable
                      onDragStart={() => { dragCol.current = col.id; }}
                      onDragOver={e => e.preventDefault()}
                      onDrop={() => {
                        if (!dragCol.current || dragCol.current === col.id) return;
                        const from = columnOrder.indexOf(dragCol.current);
                        const to   = columnOrder.indexOf(col.id);
                        const next = [...columnOrder];
                        next.splice(from, 1);
                        next.splice(to, 0, dragCol.current);
                        setColumnOrder(next);
                        dragCol.current = null;
                      }}
                      onClick={() => {
                        if (!col.sortKey) return;
                        if (sortCol === col.id) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
                        else { setSortCol(col.id); setSortDir('desc'); }
                      }}
                      className={`px-3 py-2 text-left font-semibold whitespace-nowrap select-none ${col.sortKey ? 'cursor-pointer hover:bg-gray-100' : 'cursor-grab'} ${isSorted ? 'text-blue-600' : 'text-gray-500'}`}>
                      <span className="flex items-center gap-1">
                        <span className="text-gray-300 text-xs mr-0.5" title="Arrastra para reordenar">⠿</span>
                        {col.label}
                        {isSorted && <span className="text-blue-500 text-xs">{sortDir === 'asc' ? '↑' : '↓'}</span>}
                      </span>
                    </th>
                  );
                })}
                {FIXED_METRIC_COLS.map(col => {
                  const isSorted = sortCol === col.id;
                  return (
                    <th key={col.id}
                      onClick={() => {
                        if (sortCol === col.id) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
                        else { setSortCol(col.id); setSortDir('desc'); }
                      }}
                      className={`px-3 py-2 text-left font-semibold whitespace-nowrap cursor-pointer select-none hover:bg-gray-100 ${isSorted ? 'text-blue-600' : 'text-gray-500'}`}>
                      <span className="flex items-center gap-1">
                        {col.label}
                        {isSorted && <span className="text-blue-500 text-xs">{sortDir === 'asc' ? '↑' : '↓'}</span>}
                      </span>
                    </th>
                  );
                })}
                {months.map(m => (
                  <th key={m} className="px-3 py-2 text-right font-semibold text-gray-400 whitespace-nowrap">{fmtMonthLabel(m)}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {renderGroupedTable(groupTree, groupTree === null ? sortedTickets.map(t => ({ ...t, _depth: 0 })) : null, expandedNodes, toggleNode, months, descriptions, descLoading, columnOrder, false, false)}
            </tbody>
            <tfoot className="border-t-2 border-gray-300 bg-gray-50">
              <tr>
                <td colSpan={columnOrder.length + 1} className="px-3 py-2 text-xs font-bold text-gray-600">
                  TOTAL ({totalTix} tickets · {groupTree?.length || 0} {GROUP_OPTIONS.find(o => o.value === groupBy[0])?.label.toLowerCase() || 'grupos'})
                </td>
                <td className="px-3 py-2 font-mono font-bold text-gray-500 text-xs">
                  {fmtH(Math.round(flatTickets.reduce((s, t) => s + (t.totalEstH || 0), 0) * 100) / 100)}
                </td>
                <td className="px-3 py-2 font-mono font-bold text-blue-800 text-sm">
                  {fmtH(totalH)}
                </td>
                {months.map(m => {
                  const mh = Math.round(flatTickets.reduce((s, t) => s + (t.wlByMonth?.[m] || 0), 0) * 100) / 100;
                  return <td key={m} className="px-3 py-2 font-mono font-bold text-right text-blue-700">{mh > 0 ? `${mh}h` : '—'}</td>;
                })}
              </tr>
            </tfoot>
          </table>
          </div>
        </div>
      )}

      {loadPct !== null && (
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          <div className="flex flex-col items-center gap-4 py-10">
            <div className="relative w-32 h-32">
              <svg className="w-full h-full" viewBox="0 0 100 100">
                <circle cx="50" cy="50" r="40" fill="none" stroke="#e5e7eb" strokeWidth="7"/>
                <circle cx="50" cy="50" r="40" fill="none" stroke="#3b82f6" strokeWidth="7"
                  strokeDasharray={`${2*Math.PI*40}`}
                  strokeDashoffset={`${2*Math.PI*40*(1-(loadPct||0)/100)}`}
                  strokeLinecap="round" transform="rotate(-90 50 50)"
                  style={{ transition: 'stroke-dashoffset 0.4s ease' }}/>
              </svg>
              <div className="absolute inset-0 flex flex-col items-center justify-center">
                <span className="text-3xl font-bold text-blue-600 tabular-nums">{Math.round(loadPct||0)}%</span>
                <span className="text-xs text-gray-400 mt-0.5">cargando</span>
              </div>
            </div>
            <p className="text-sm text-gray-500 animate-pulse">Consultando Jira y Tempo…</p>
          </div>
          <div className="px-6 pb-8 space-y-2.5 border-t pt-4">
            {[90,70,85,60,80,75,95,65].map((w, i) => (
              <div key={i} className="flex gap-3 animate-pulse" style={{ animationDelay: `${i*80}ms` }}>
                <div className="h-5 bg-gray-200 rounded" style={{ width: '9%' }}/>
                <div className="h-5 bg-gray-100 rounded" style={{ width: `${Math.round(w*0.3)}%` }}/>
                <div className="h-5 bg-gray-200 rounded flex-1"/>
                <div className="h-5 bg-gray-100 rounded w-16"/>
                <div className="h-5 bg-gray-200 rounded w-12"/>
              </div>
            ))}
          </div>
        </div>
      )}

      {!data && loadPct === null && (
        <div className="bg-white rounded-lg shadow-sm p-12 text-center text-gray-400">
          Selecciona el rango de fechas y haz clic en <strong>Consultar Jira</strong>.
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// TAB 2: DESGLOSE DE REGISTROS — todos los tickets con tiempo propio
// ══════════════════════════════════════════════════════════════════════════════
function DetailTimeTab({ teamIds, setTeamIds, teamMembers }) {
  const [dateFrom,   setDateFrom]   = useState('2026-01-01');
  const [dateTo,     setDateTo]     = useState('2026-03-31');
  const [data,       setData]       = useState(null);
  const [detailTickets, setDetailTickets] = useState([]);
  const [loadPct,    setLoadPct]    = useState(null);
  const [error,      setError]      = useState('');
  const [search,     setSearch]     = useState('');
  const [groupBy,    setGroupBy]    = useState(['proyecto']);
  const [columnOrder,setColumnOrder]= useState(DETAIL_REORDERABLE_COLS.map(c => c.id));
  const [sortCol,    setSortCol]    = useState('htotal');
  const [sortDir,    setSortDir]    = useState('desc');
  const dragCol = useRef(null);
  const [expandedNodes,     setExpandedNodes]     = useState(new Set());
  const [projectInclude,    setProjectInclude]    = useState([]);
  const [projectExclude,    setProjectExclude]    = useState([]);
  const [personInclude,     setPersonInclude]     = useState([]);
  const [personExclude,     setPersonExclude]     = useState([]);
  const [facturacionFilter, setFacturacionFilter] = useState('');
  const [colFilters,        setColFilters]        = useState({});
  const [showFilters,       setShowFilters]       = useState(false);
  const [showPersonCards,   setShowPersonCards]   = useState(true);
  const [showFacturacion,   setShowFacturacion]   = useState(true);
  const [cachedMeta,        setCachedMeta]        = useState(null);
  const [descriptions,      setDescriptions]      = useState({});
  const [descLoading,       setDescLoading]       = useState(false);
  const [exportPct,         setExportPct]         = useState(null);
  const hasLoadedRef  = useRef(false);
  const mountedRef    = useRef(false);
  const loadRef       = useRef(null);

  const resetView = () => { setData(null); setDetailTickets([]); setError(''); };

  const teamIdsKey = [...teamIds].sort().join(',');
  useEffect(() => { if (hasLoadedRef.current) resetView(); }, [dateFrom, dateTo]);
  useEffect(() => {
    if (!mountedRef.current) { mountedRef.current = true; return; }
    if (teamIds.size > 0) loadRef.current?.();
  }, [teamIdsKey]);

  const load = async () => {
    if (loadPct !== null) return;
    setData(null); setDetailTickets([]);
    setLoadPct(5); setError('');
    const timer = setInterval(() => setLoadPct(p => p !== null && p < 85 ? Math.min(85, p + 8 + Math.random() * 7) : p), 600);
    try {
      const ids = [...teamIds].join(',');
      const params = new URLSearchParams({ dateFrom, dateTo, userIds: ids, all: 'true' });
      const r = await apiFetch(`/api/team-time?${params}`);
      const d = await r.json();
      if (!r.ok) throw new Error(d.error);
      hasLoadedRef.current = true;
      setLoadPct(100);
      setData(d);
      setDetailTickets(buildDetailView(d.tickets || []));
      setProjectInclude([]); setProjectExclude([]);
      setPersonInclude([]);  setPersonExclude([]);
      setFacturacionFilter('');
      setColFilters({});
      setSearch('');
      if (d.allProjects) setCachedMeta({ projects: d.allProjects, users: d.allUsers || [] });
    } catch (e) { setError(e.message); }
    finally     { clearInterval(timer); setTimeout(() => setLoadPct(null), 400); }
  };
  loadRef.current = load;

  const loadDescriptions = async (tickets, force = false) => {
    const candidates = tickets.filter(t => t.title);
    const toLoad = force ? candidates : candidates.filter(t => !descriptions[t.key]);
    if (!toLoad.length) return;
    setDescLoading(true);
    const BATCH = 75;
    try {
      for (let i = 0; i < toLoad.length; i += BATCH) {
        const batch = toLoad.slice(i, i + BATCH);
        const r = await apiFetch('/api/summarize', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tickets: batch.map(t => ({ key: t.key, title: t.title, type: t.type, description: t.description || '' })) }),
        });
        const d = await r.json();
        if (d.error) throw new Error(d.error);
        if (d.descriptions) setDescriptions(prev => ({ ...prev, ...d.descriptions }));
      }
    } catch (err) { setError('Gemini: ' + err.message); }
    finally { setDescLoading(false); }
  };

  const toggleNode  = id  => setExpandedNodes(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const expandAll   = () => setExpandedNodes(new Set(collectGroupIds(groupTree)));
  const collapseAll = () => setExpandedNodes(new Set());

  useEffect(() => { setExpandedNodes(new Set()); }, [groupBy]);

  const months = getMonths(dateFrom, dateTo);

  const filteredTickets = detailTickets.filter(t => {
    if (projectInclude.length > 0 && !projectInclude.includes(t.project)) return false;
    if (projectExclude.length > 0 &&  projectExclude.includes(t.project)) return false;
    if (facturacionFilter && t.facturacion !== facturacionFilter) return false;
    const cfTipo   = colFilters.tipo;
    const cfEstado = colFilters.estado;
    const cfQuart  = colFilters.quarter;
    if (cfTipo?.length   > 0 && !cfTipo.includes(t.type   || 'Sin tipo'))   return false;
    if (cfEstado?.length > 0 && !cfEstado.includes(t.status || 'Sin estado')) return false;
    if (cfQuart?.length  > 0) {
      const qs = new Set(Object.entries(t.wlByMonth || {}).filter(([, h]) => h > 0).map(([ym]) => {
        const mo = parseInt(ym.split('-')[1]); return mo <= 3 ? 'Q1' : mo <= 6 ? 'Q2' : mo <= 9 ? 'Q3' : 'Q4';
      }));
      if (!cfQuart.some(q => qs.has(q))) return false;
    }
    if (personInclude.length > 0 && !personInclude.some(p => (t.wlAuthors || []).includes(p))) return false;
    if (personExclude.length > 0 &&  personExclude.some(p => (t.wlAuthors || []).includes(p))) return false;
    if (search) { const q = search.toLowerCase(); return t.key.toLowerCase().includes(q) || t.title.toLowerCase().includes(q); }
    return true;
  });

  const allProjectsInfo = cachedMeta?.projects || (data
    ? [...new Map(data.tickets.map(t => [t.project, t.projectName || t.project])).entries()]
        .filter(([k]) => k)
        .sort((a, b) => (a[1] || a[0]).localeCompare(b[1] || b[0]))
        .map(([key, name]) => ({ key, name: name || key }))
    : []);
  const allUsers = cachedMeta?.users || (data ? [...new Set(data.tickets.flatMap(t => t.wlAuthors || []))] : []);

  const facturacionStats = (() => {
    let facturables = 0, noFacturable = 0, sinClasif = 0;
    filteredTickets.forEach(t => {
      const h = t.totalHours || 0;
      if (t.facturacion === 'Facturables')  facturables  += h;
      else if (t.facturacion === 'No facturable') noFacturable += h;
      else sinClasif += h;
    });
    const r = v => Math.round(v * 100) / 100;
    return { facturables: r(facturables), noFacturable: r(noFacturable), sinClasif: r(sinClasif) };
  })();

  const totalH = Math.round(filteredTickets.reduce((s, t) => s + (t.totalHours || 0), 0) * 100) / 100;

  const personData = data ? (() => {
    const byPerson = {};
    filteredTickets.forEach(t => {
      Object.entries(t.wlByAuthor || {}).forEach(([name]) => {
        const monthlyHrs = months.reduce((s, m) => s + (t.wlByAuthorByMonth?.[name]?.[m] || 0), 0);
        if (monthlyHrs <= 0) return;
        if (!byPerson[name]) byPerson[name] = { hrs: 0, ticketMap: new Map() };
        byPerson[name].hrs += monthlyHrs;
        if (!byPerson[name].ticketMap.has(t.key)) {
          byPerson[name].ticketMap.set(t.key, { ...t, hoursForPerson: Math.round(monthlyHrs * 100) / 100 });
        }
      });
    });
    return Object.entries(byPerson)
      .filter(([, d]) => d.hrs > 0)
      .map(([name, d]) => {
        const tickets = [...d.ticketMap.values()].sort((a, b) => b.hoursForPerson - a.hoursForPerson);
        return { name, totalHours: Math.round(d.hrs * 100) / 100, tickets };
      })
      .sort((a, b) => b.totalHours - a.totalHours);
  })() : [];

  const allTipos   = data ? [...new Set((data.tickets || []).map(t => t.type   || 'Sin tipo'))].sort()   : [];
  const allEstados = data ? [...new Set((data.tickets || []).map(t => t.status || 'Sin estado'))].sort() : [];

  const sortedTickets = (() => {
    const col = [...DETAIL_REORDERABLE_COLS, ...FIXED_METRIC_COLS].find(c => c.id === sortCol);
    if (!col?.sortKey) return filteredTickets;
    return [...filteredTickets].sort((a, b) => {
      const av = a[col.sortKey] ?? '', bv = b[col.sortKey] ?? '';
      const cmp = (typeof av === 'number' && typeof bv === 'number') ? av - bv : String(av).localeCompare(String(bv));
      return sortDir === 'asc' ? cmp : -cmp;
    });
  })();

  const groupTree = buildGroupTree(sortedTickets, groupBy);

  const exportExcel = async () => {
    try {
      setExportPct(50);
      const r = await apiFetch('/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type: 'excel', data: { tickets: filteredTickets, months, descriptions, dateFrom, dateTo, tabName: 'detalle' } }),
      });
      setExportPct(95);
      const blob = await r.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a'); a.href = url; a.download = `omni-desglose-${dateFrom}.xlsx`; a.click();
    } finally {
      setExportPct(null);
    }
  };

  return (
    <div className="space-y-4">
      <p className="text-xs text-gray-400 leading-relaxed">
        Lista cada ticket individual con su tiempo registrado desglosado por mes. Cuando un ticket tiene múltiples personas que registraron horas, muestra una fila por autor. Ideal para auditar quién registró qué y cuándo. Los filtros se aplican sobre los datos ya cargados.
      </p>
      {/* Controls */}
      <div className="bg-white rounded-lg shadow-sm p-4 space-y-3">
        <div className="flex flex-wrap gap-3 items-end">
          <div>
            <label className="block text-xs font-semibold text-gray-500 mb-1">Desde</label>
            <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} disabled={teamIds.size === 0} className="border rounded px-2 py-1.5 text-sm disabled:bg-gray-100 disabled:text-gray-400 disabled:cursor-not-allowed" />
          </div>
          <div>
            <label className="block text-xs font-semibold text-gray-500 mb-1">Hasta</label>
            <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} disabled={teamIds.size === 0} className="border rounded px-2 py-1.5 text-sm disabled:bg-gray-100 disabled:text-gray-400 disabled:cursor-not-allowed" />
          </div>
          <button onClick={load} disabled={loadPct !== null || teamIds.size === 0}
            className={`px-5 py-2 rounded-lg text-sm font-medium flex items-center gap-2 min-w-[160px] justify-center transition-colors ${teamIds.size === 0 ? 'bg-gray-200 text-gray-400 cursor-not-allowed' : loadPct !== null ? 'bg-blue-500 text-white cursor-not-allowed' : 'bg-blue-600 text-white hover:bg-blue-700'}`}>
            {loadPct !== null
              ? <><svg className="w-4 h-4 shrink-0 animate-spin" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><circle cx="12" cy="12" r="10" strokeDasharray="40" strokeDashoffset="15"/></svg><span>Consultando…</span></>
              : '🔍 Consultar Jira'}
          </button>
          {teamIds.size === 0 && <span className="text-xs text-amber-600 font-medium">← Selecciona un equipo primero</span>}
          {data && (
            <>
              <select value={groupBy[0] || 'proyecto'} onChange={e => setGroupBy([e.target.value])}
                className="border rounded px-2 py-1.5 text-xs text-gray-600 bg-white">
                {GROUP_OPTIONS.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
              </select>
              <input placeholder="Buscar ticket…" value={search} onChange={e => setSearch(e.target.value)} className="border rounded px-2 py-1.5 text-sm w-40" />
              <div className="ml-auto flex gap-2 items-center">
                <button onClick={expandAll}   className="px-3 py-1.5 text-xs border rounded hover:bg-gray-50">Expandir</button>
                <button onClick={collapseAll} className="px-3 py-1.5 text-xs border rounded hover:bg-gray-50">Contraer</button>
                <button onClick={exportExcel} disabled={exportPct !== null}
                  className="px-4 py-1.5 bg-green-600 text-white text-sm rounded hover:bg-green-700 disabled:opacity-75 flex items-center gap-1.5 min-w-[90px] justify-center">
                  {exportPct !== null
                    ? <><svg className="w-3.5 h-3.5 animate-spin shrink-0" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" strokeDasharray="40" strokeDashoffset="10"/></svg>{exportPct}%</>
                    : '⬇ Excel'}
                </button>
              </div>
            </>
          )}
        </div>

        {/* Toggle filtros */}
        {data && (() => {
          const n = projectInclude.length + projectExclude.length + personInclude.length + personExclude.length + (facturacionFilter ? 1 : 0) + Object.values(colFilters).reduce((s, a) => s + (a?.length || 0), 0);
          return (
            <button onClick={() => setShowFilters(v => !v)}
              className={`flex items-center gap-1.5 text-xs pt-1 ${n > 0 ? 'text-blue-600 font-semibold' : 'text-gray-400'} hover:text-blue-500`}>
              <span>{showFilters ? '▼' : '▶'}</span>
              Filtros
              {n > 0 && <span className="bg-blue-500 text-white rounded-full px-1.5 py-0.5 text-xs leading-none">{n}</span>}
            </button>
          );
        })()}

        {showFilters && <>

        {/* Proyecto */}
        {data && (
          <div className={`flex items-start gap-3 pt-2 border-t ${projectInclude.length > 0 || projectExclude.length > 0 ? 'bg-blue-50 -mx-4 px-4 rounded' : ''}`}>
            <span className={`text-xs font-bold uppercase tracking-wide w-16 mt-1.5 shrink-0 ${projectInclude.length > 0 || projectExclude.length > 0 ? 'text-blue-600' : 'text-gray-400'}`}>Proyecto</span>
            <div className="flex flex-wrap gap-x-6 gap-y-1.5 flex-1">
              <div className="flex flex-wrap gap-1 items-center">
                <span className="text-xs font-semibold text-green-600 shrink-0 mr-0.5">Incluir</span>
                {projectInclude.map(p => {
                  const info = allProjectsInfo.find(i => i.key === p);
                  return (
                    <span key={p} className="inline-flex items-center bg-green-100 text-green-700 border border-green-300 px-1.5 py-0.5 rounded text-xs">
                      {info?.name || p}
                      <button onClick={() => setProjectInclude(prev => prev.filter(x => x !== p))} className="ml-1 leading-none hover:text-red-600">×</button>
                    </span>
                  );
                })}
                <select value="" onChange={e => { if (e.target.value && !projectInclude.includes(e.target.value)) setProjectInclude(prev => [...prev, e.target.value]); e.target.value = ''; }}
                  className="border rounded px-1.5 py-0.5 text-xs text-gray-500">
                  <option value="">+ añadir</option>
                  {allProjectsInfo.filter(p => !projectInclude.includes(p.key)).map(p => <option key={p.key} value={p.key}>{p.name}</option>)}
                </select>
              </div>
              <div className="flex flex-wrap gap-1 items-center">
                <span className="text-xs font-semibold text-red-500 shrink-0 mr-0.5">Excluir</span>
                {projectExclude.map(p => {
                  const info = allProjectsInfo.find(i => i.key === p);
                  return (
                    <span key={p} className="inline-flex items-center bg-red-100 text-red-600 border border-red-300 px-1.5 py-0.5 rounded text-xs">
                      {info?.name || p}
                      <button onClick={() => setProjectExclude(prev => prev.filter(x => x !== p))} className="ml-1 leading-none hover:text-red-800">×</button>
                    </span>
                  );
                })}
                <select value="" onChange={e => { if (e.target.value && !projectExclude.includes(e.target.value)) setProjectExclude(prev => [...prev, e.target.value]); e.target.value = ''; }}
                  className="border rounded px-1.5 py-0.5 text-xs text-gray-500">
                  <option value="">+ añadir</option>
                  {allProjectsInfo.filter(p => !projectExclude.includes(p.key)).map(p => <option key={p.key} value={p.key}>{p.name}</option>)}
                </select>
              </div>
            </div>
          </div>
        )}

        {/* Facturación */}
        {data && (
          <div className={`flex items-start gap-3 pt-2 border-t ${facturacionFilter ? 'bg-blue-50 -mx-4 px-4 rounded' : ''}`}>
            <span className={`text-xs font-bold uppercase tracking-wide w-16 mt-1.5 shrink-0 ${facturacionFilter ? 'text-blue-600' : 'text-gray-400'}`}>Facturac.</span>
            <div className="flex gap-1 items-center mt-1">
              {['', 'Facturables', 'No facturable'].map(v => (
                <button key={v} onClick={() => setFacturacionFilter(v)}
                  className={`px-2.5 py-0.5 rounded text-xs border transition-colors ${
                    facturacionFilter === v
                      ? v === 'Facturables'   ? 'bg-green-100 text-green-700 border-green-400 font-semibold'
                      : v === 'No facturable' ? 'bg-orange-100 text-orange-700 border-orange-400 font-semibold'
                      : 'bg-blue-100 text-blue-700 border-blue-300 font-semibold'
                      : 'bg-white text-gray-500 border-gray-200 hover:bg-gray-50'
                  }`}>
                  {v === '' ? 'Todos' : v}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Personas */}
        {data && (
          <div className={`flex items-start gap-3 pt-2 border-t ${personInclude.length > 0 || personExclude.length > 0 ? 'bg-blue-50 -mx-4 px-4 rounded' : ''}`}>
            <span className={`text-xs font-bold uppercase tracking-wide w-16 mt-1.5 shrink-0 ${personInclude.length > 0 || personExclude.length > 0 ? 'text-blue-600' : 'text-gray-400'}`}>Personas</span>
            <div className="flex flex-wrap gap-x-6 gap-y-1.5 flex-1">
              <div className="flex flex-wrap gap-1 items-center">
                <span className="text-xs font-semibold text-green-600 shrink-0 mr-0.5">Incluir</span>
                {personInclude.map(p => (
                  <span key={p} className="inline-flex items-center bg-green-100 text-green-700 border border-green-300 px-1.5 py-0.5 rounded text-xs">
                    {p.split(' ')[0]}<button onClick={() => setPersonInclude(prev => prev.filter(x => x !== p))} className="ml-1 leading-none hover:text-red-600">×</button>
                  </span>
                ))}
                <select value="" onChange={e => { if (e.target.value && !personInclude.includes(e.target.value)) setPersonInclude(prev => [...prev, e.target.value]); e.target.value = ''; }}
                  className="border rounded px-1.5 py-0.5 text-xs text-gray-500">
                  <option value="">+ añadir</option>
                  {allUsers.filter(u => !personInclude.includes(u)).map(u => <option key={u} value={u}>{u}</option>)}
                </select>
              </div>
              <div className="flex flex-wrap gap-1 items-center">
                <span className="text-xs font-semibold text-red-500 shrink-0 mr-0.5">Excluir</span>
                {personExclude.map(p => (
                  <span key={p} className="inline-flex items-center bg-red-100 text-red-600 border border-red-300 px-1.5 py-0.5 rounded text-xs">
                    {p.split(' ')[0]}<button onClick={() => setPersonExclude(prev => prev.filter(x => x !== p))} className="ml-1 leading-none hover:text-red-800">×</button>
                  </span>
                ))}
                <select value="" onChange={e => { if (e.target.value && !personExclude.includes(e.target.value)) setPersonExclude(prev => [...prev, e.target.value]); e.target.value = ''; }}
                  className="border rounded px-1.5 py-0.5 text-xs text-gray-500">
                  <option value="">+ añadir</option>
                  {allUsers.filter(u => !personExclude.includes(u)).map(u => <option key={u} value={u}>{u}</option>)}
                </select>
              </div>
            </div>
          </div>
        )}

        {/* Filtros de columna */}
        {data && (
          <div className={`flex flex-wrap items-center gap-2 pt-2 border-t ${Object.values(colFilters).some(a => a?.length > 0) ? 'bg-blue-50 -mx-4 px-4 rounded' : ''}`}>
            <span className={`text-xs font-bold uppercase tracking-wide w-16 shrink-0 ${Object.values(colFilters).some(a => a?.length > 0) ? 'text-blue-600' : 'text-gray-400'}`}>Columnas</span>
            {[
              { id: 'tipo',    label: 'Tipo',    opts: allTipos },
              { id: 'estado',  label: 'Estado',  opts: allEstados },
              { id: 'quarter', label: 'Quarter', opts: ['Q1', 'Q2', 'Q3', 'Q4'] },
            ].map(({ id, label, opts }) => {
              const active = colFilters[id] || [];
              return (
                <div key={id} className="flex items-center gap-1 flex-wrap">
                  {active.map(v => (
                    <span key={v} className="inline-flex items-center bg-violet-100 text-violet-700 border border-violet-300 px-1.5 py-0.5 rounded text-xs">
                      {label}: {v}
                      <button onClick={() => setColFilters(p => ({ ...p, [id]: (p[id] || []).filter(x => x !== v) }))} className="ml-1 leading-none hover:text-red-600">×</button>
                    </span>
                  ))}
                  {opts.filter(v => !active.includes(v)).length > 0 && (
                    <select value="" onChange={e => { if (e.target.value) { setColFilters(p => ({ ...p, [id]: [...(p[id] || []), e.target.value] })); e.target.value = ''; } }}
                      className="border rounded px-1.5 py-0.5 text-xs text-gray-500 bg-white">
                      <option value="">+ {label}</option>
                      {opts.filter(v => !active.includes(v)).map(v => <option key={v} value={v}>{v}</option>)}
                    </select>
                  )}
                </div>
              );
            })}
            {Object.values(colFilters).some(a => a?.length > 0) && (
              <button onClick={() => setColFilters({})} className="text-xs text-red-400 hover:text-red-600 ml-1">Limpiar</button>
            )}
          </div>
        )}

        </>}

        {error && <div className="text-red-600 text-sm bg-red-50 border border-red-200 rounded px-3 py-2">{error}</div>}
      </div>

      {/* Facturación breakdown */}
      {data && (facturacionStats.facturables > 0 || facturacionStats.noFacturable > 0) && (
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          <button onClick={() => setShowFacturacion(v => !v)} className="w-full flex items-center gap-2 px-3 py-2 hover:bg-gray-50">
            <span className="text-gray-400 text-xs">{showFacturacion ? '▼' : '▶'}</span>
            <span className="text-xs font-bold text-gray-400 uppercase tracking-wide">Facturación</span>
          </button>
          {showFacturacion && (
            <div className="px-3 pb-3 flex flex-wrap gap-6 items-center border-t pt-2">
              <div className="flex items-center gap-1.5 cursor-pointer" onClick={() => setFacturacionFilter(facturacionFilter === 'Facturables' ? '' : 'Facturables')}>
                <div className="w-2.5 h-2.5 rounded-full bg-green-500" /><span className="text-xs text-gray-500">Facturables</span>
                <span className="font-bold text-sm text-green-700">{facturacionStats.facturables}h</span>
              </div>
              <div className="flex items-center gap-1.5 cursor-pointer" onClick={() => setFacturacionFilter(facturacionFilter === 'No facturable' ? '' : 'No facturable')}>
                <div className="w-2.5 h-2.5 rounded-full bg-orange-400" /><span className="text-xs text-gray-500">No facturable</span>
                <span className="font-bold text-sm text-orange-600">{facturacionStats.noFacturable}h</span>
              </div>
              {facturacionStats.sinClasif > 0 && (
                <div className="flex items-center gap-1.5">
                  <div className="w-2.5 h-2.5 rounded-full bg-gray-300" /><span className="text-xs text-gray-400">Sin clasificar</span>
                  <span className="font-bold text-sm text-gray-400">{facturacionStats.sinClasif}h</span>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* Person cards */}
      {data && personData.length > 0 && (
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          <button onClick={() => setShowPersonCards(v => !v)} className="w-full flex items-center gap-2 px-4 py-2.5 hover:bg-gray-50">
            <span className="text-gray-400 text-xs">{showPersonCards ? '▼' : '▶'}</span>
            <span className="text-xs font-semibold text-gray-400 uppercase tracking-wide">Horas por persona · período</span>
          </button>
          {showPersonCards && (
            <div className="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-7 gap-2 px-4 pb-4 border-t pt-3">
              {personData.map(p => (
                <button key={p.name}
                  onClick={() => setPersonInclude(prev => prev.includes(p.name) ? prev.filter(x => x !== p.name) : [...prev, p.name])}
                  title={p.name}
                  className={`p-2 rounded-lg border text-left transition-all ${personInclude.includes(p.name) ? 'border-blue-500 bg-blue-50 shadow-sm' : 'border-gray-200 hover:border-blue-300 hover:bg-blue-50'}`}>
                  <div className="font-mono font-bold text-lg text-blue-700 leading-none">{p.totalHours}h</div>
                  <div className="text-xs font-medium text-gray-800 mt-1 truncate">{p.name.split(' ').slice(0, 2).join(' ')}</div>
                  <div className="text-xs text-gray-400">{p.tickets.length} tickets</div>
                </button>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Table */}
      {data && (
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead className="bg-gray-50 border-b sticky top-0 z-10">
              <tr>
                <th className="px-3 py-2 w-6" />
                {columnOrder.map(colId => {
                  const col = DETAIL_REORDERABLE_COLS.find(c => c.id === colId);
                  if (!col) return null;
                  const isSorted = sortCol === col.id;
                  return (
                    <th key={col.id}
                      draggable
                      onDragStart={() => { dragCol.current = col.id; }}
                      onDragOver={e => e.preventDefault()}
                      onDrop={() => {
                        if (!dragCol.current || dragCol.current === col.id) return;
                        const from = columnOrder.indexOf(dragCol.current);
                        const to   = columnOrder.indexOf(col.id);
                        const next = [...columnOrder];
                        next.splice(from, 1);
                        next.splice(to, 0, dragCol.current);
                        setColumnOrder(next);
                        dragCol.current = null;
                      }}
                      onClick={() => {
                        if (!col.sortKey) return;
                        if (sortCol === col.id) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
                        else { setSortCol(col.id); setSortDir('desc'); }
                      }}
                      className={`px-3 py-2 text-left font-semibold whitespace-nowrap select-none ${col.sortKey ? 'cursor-pointer hover:bg-gray-100' : 'cursor-grab'} ${isSorted ? 'text-blue-600' : 'text-gray-500'}`}>
                      <span className="flex items-center gap-1">
                        <span className="text-gray-300 text-xs mr-0.5" title="Arrastra para reordenar">⠿</span>
                        {col.label}
                        {isSorted && <span className="text-blue-500 text-xs">{sortDir === 'asc' ? '↑' : '↓'}</span>}
                      </span>
                    </th>
                  );
                })}
                {FIXED_METRIC_COLS.map(col => {
                  const isSorted = sortCol === col.id;
                  return (
                    <th key={col.id}
                      onClick={() => { if (sortCol === col.id) setSortDir(d => d === 'asc' ? 'desc' : 'asc'); else { setSortCol(col.id); setSortDir('desc'); } }}
                      className={`px-3 py-2 text-left font-semibold whitespace-nowrap cursor-pointer select-none hover:bg-gray-100 ${isSorted ? 'text-blue-600' : 'text-gray-500'}`}>
                      <span className="flex items-center gap-1">
                        {col.label}{isSorted && <span className="text-blue-500 text-xs">{sortDir === 'asc' ? '↑' : '↓'}</span>}
                      </span>
                    </th>
                  );
                })}
                {months.map(m => (
                  <th key={m} className="px-3 py-2 text-right font-semibold text-gray-400 whitespace-nowrap">{fmtMonthLabel(m)}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {renderGroupedTable(groupTree, groupTree === null ? sortedTickets.map(t => ({ ...t, _depth: 0 })) : null, expandedNodes, toggleNode, months, descriptions, descLoading, columnOrder, true, false)}
            </tbody>
            <tfoot className="border-t-2 border-gray-300 bg-gray-50">
              <tr>
                <td colSpan={columnOrder.length + 1} className="px-3 py-2 text-xs font-bold text-gray-600">
                  TOTAL ({filteredTickets.length} registros · {groupTree?.length || 0} {GROUP_OPTIONS.find(o => o.value === groupBy[0])?.label.toLowerCase() || 'grupos'})
                </td>
                <td className="px-3 py-2 font-mono font-bold text-gray-500 text-xs">
                  {fmtH(Math.round(filteredTickets.reduce((s, t) => s + (t.totalEstH || 0), 0) * 100) / 100)}
                </td>
                <td className="px-3 py-2 font-mono font-bold text-blue-800 text-sm">
                  {fmtH(totalH)}
                </td>
                {months.map(m => {
                  const mh = Math.round(filteredTickets.reduce((s, t) => s + (t.wlByMonth?.[m] || 0), 0) * 100) / 100;
                  return <td key={m} className="px-3 py-2 font-mono font-bold text-right text-blue-700">{mh > 0 ? `${mh}h` : '—'}</td>;
                })}
              </tr>
            </tfoot>
          </table>
          </div>
        </div>
      )}

      {loadPct !== null && (
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          <div className="flex flex-col items-center gap-4 py-10">
            <div className="relative w-32 h-32">
              <svg className="w-full h-full" viewBox="0 0 100 100">
                <circle cx="50" cy="50" r="40" fill="none" stroke="#e5e7eb" strokeWidth="7"/>
                <circle cx="50" cy="50" r="40" fill="none" stroke="#3b82f6" strokeWidth="7"
                  strokeDasharray={`${2*Math.PI*40}`}
                  strokeDashoffset={`${2*Math.PI*40*(1-(loadPct||0)/100)}`}
                  strokeLinecap="round" transform="rotate(-90 50 50)"
                  style={{ transition: 'stroke-dashoffset 0.4s ease' }}/>
              </svg>
              <div className="absolute inset-0 flex flex-col items-center justify-center">
                <span className="text-3xl font-bold text-blue-600 tabular-nums">{Math.round(loadPct||0)}%</span>
                <span className="text-xs text-gray-400 mt-0.5">cargando</span>
              </div>
            </div>
            <p className="text-sm text-gray-500 animate-pulse">Consultando Jira y Tempo…</p>
          </div>
          <div className="px-6 pb-8 space-y-2.5 border-t pt-4">
            {[90,70,85,60,80,75,95,65].map((w, i) => (
              <div key={i} className="flex gap-3 animate-pulse" style={{ animationDelay: `${i*80}ms` }}>
                <div className="h-5 bg-gray-200 rounded" style={{ width: '9%' }}/>
                <div className="h-5 bg-gray-100 rounded" style={{ width: `${Math.round(w*0.3)}%` }}/>
                <div className="h-5 bg-gray-200 rounded flex-1"/>
                <div className="h-5 bg-gray-100 rounded w-16"/>
                <div className="h-5 bg-gray-200 rounded w-12"/>
              </div>
            ))}
          </div>
        </div>
      )}

      {!data && loadPct === null && (
        <div className="bg-white rounded-lg shadow-sm p-12 text-center text-gray-400">
          Selecciona el rango de fechas y haz clic en <strong>Consultar Jira</strong>.
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// TAB 3: REGISTRO DE TIEMPO (GANTT)
// ══════════════════════════════════════════════════════════════════════════════
const GA_MONTH = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
const GA_DAYS  = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];

function isWknd(d) { const n = new Date(d + 'T00:00:00').getDay(); return n === 0 || n === 6; }
function sumH(byDay, dates) { return Math.round(dates.reduce((s, d) => s + (byDay[d] || 0), 0) * 10) / 10; }

function GanttTimeTab({ teamIds, teamMembers }) {
  const [dateFrom,     setDateFrom]     = useState('2026-01-01');
  const [dateTo,       setDateTo]       = useState('2026-03-31');
  const [zoom,         setZoom]         = useState('week');
  const [data,         setData]         = useState(null);
  const [loading,      setLoading]      = useState(false);
  const [error,        setError]        = useState('');
  const [activePeople, setActivePeople] = useState(new Set());
  const [expanded,     setExpanded]     = useState(new Set());
  const [exportPct,    setExportPct]    = useState(null);
  const [showGap,      setShowGap]      = useState(true);

  const load = async () => {
    setLoading(true); setError('');
    try {
      const ids = [...teamIds].join(',');
      const r = await apiFetch(`/api/gantt-time?${new URLSearchParams({ dateFrom, dateTo, userIds: ids })}`);
      const d = await r.json();
      if (!r.ok) throw new Error(d.error || 'Error');
      setData(d);
      setActivePeople(new Set());
      setExpanded(new Set());
    } catch(e) { setError(e.message); }
    finally { setLoading(false); }
  };

  const allDates = useMemo(() => {
    if (!dateFrom || !dateTo) return [];
    const out = [];
    let d = new Date(dateFrom + 'T00:00:00');
    const end = new Date(dateTo + 'T00:00:00');
    while (d <= end) { out.push(d.toISOString().split('T')[0]); d.setDate(d.getDate() + 1); }
    return out;
  }, [dateFrom, dateTo]);

  const columns = useMemo(() => {
    if (!allDates.length) return [];
    if (zoom === 'day') {
      return allDates.map(d => {
        const dt = new Date(d + 'T00:00:00');
        return { key: d, label: String(dt.getDate()), sub: GA_DAYS[dt.getDay()], monthKey: d.slice(0, 7), isWknd: isWknd(d), dates: [d] };
      });
    }
    if (zoom === 'week') {
      const wks = []; const seen = new Set();
      allDates.forEach(d => {
        const dt = new Date(d + 'T00:00:00');
        const diff = dt.getDay() === 0 ? -6 : 1 - dt.getDay();
        const mon = new Date(dt); mon.setDate(dt.getDate() + diff);
        const mk = mon.toISOString().split('T')[0];
        if (!seen.has(mk)) { seen.add(mk); wks.push({ key: mk, mon, dates: [] }); }
        wks.find(w => w.key === mk).dates.push(d);
      });
      return wks.map(w => {
        const sun = new Date(w.mon); sun.setDate(w.mon.getDate() + 6);
        return { key: w.key, label: `${w.mon.getDate()} ${GA_MONTH[w.mon.getMonth()]}`, sub: `→ ${sun.getDate()} ${GA_MONTH[sun.getMonth()]}`, monthKey: w.key.slice(0, 7), isWknd: false, dates: w.dates };
      });
    }
    if (zoom === 'month') {
      const mths = {};
      allDates.forEach(d => { const m = d.slice(0, 7); if (!mths[m]) mths[m] = { key: m, dates: [] }; mths[m].dates.push(d); });
      return Object.values(mths).map(m => {
        const [y, mo] = m.key.split('-');
        return { key: m.key, label: `${GA_MONTH[parseInt(mo) - 1]} '${y.slice(2)}`, sub: '', monthKey: m.key, isWknd: false, dates: m.dates };
      });
    }
    const qtrs = {};
    allDates.forEach(d => {
      const [y, mo] = d.split('-');
      const q = Math.ceil(parseInt(mo) / 3);
      const k = `${y}-Q${q}`;
      if (!qtrs[k]) qtrs[k] = { key: k, label: `Q${q} '${y.slice(2)}`, sub: '', monthKey: k, isWknd: false, dates: [] };
      qtrs[k].dates.push(d);
    });
    return Object.values(qtrs);
  }, [allDates, zoom]);

  const monthGroups = useMemo(() => {
    if (zoom !== 'day') return [];
    return columns.reduce((acc, col) => {
      if (!acc.length || acc[acc.length - 1].key !== col.monthKey) acc.push({ key: col.monthKey, count: 1 });
      else acc[acc.length - 1].count++;
      return acc;
    }, []);
  }, [columns, zoom]);

  const togglePerson = id => setActivePeople(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const toggleExpand = id => setExpanded(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });

  const persons  = data?.persons || [];
  const visible  = activePeople.size === 0 ? persons : persons.filter(p => activePeople.has(p.accountId));
  const totalH   = Math.round(visible.reduce((s, p) => s + Object.values(p.byDay).reduce((a, b) => a + b, 0), 0) * 10) / 10;
  const missDays = visible.reduce((s, p) => s + allDates.filter(d => !isWknd(d) && !(p.byDay[d] > 0)).length, 0);

  const workingDays = allDates.filter(d => !isWknd(d)).length;
  const expectedH   = workingDays * 8;
  const gapData = visible.map(p => {
    const registered = Math.round(Object.values(p.byDay).reduce((s, h) => s + h, 0) * 10) / 10;
    const gap = Math.round((expectedH - registered) * 10) / 10;
    const pct = expectedH > 0 ? Math.min(100, Math.round(registered / expectedH * 100)) : 0;
    return { name: p.name, registered, expectedH, gap, pct, workingDays };
  }).sort((a, b) => b.gap - a.gap);

  // Colombia: 8h por día hábil (lun–vie)
  const cellClass = (hrs, col) => {
    if (col.isWknd) return { bg: 'bg-gray-50', text: 'text-gray-200', tip: '' };
    const wDays = col.dates.filter(d => !isWknd(d));
    if (wDays.length === 0) return { bg: '', text: 'text-gray-200', tip: '' };
    const expected = wDays.length * 8;
    const pct = expected > 0 ? hrs / expected : 0;
    const tip = `${hrs}h de ${expected}h esperadas (${Math.round(pct * 100)}%)`;
    if (hrs === 0)    return { bg: 'bg-red-100',    text: 'text-red-500',    tip, bold: true };
    if (pct < 0.5)   return { bg: 'bg-orange-100', text: 'text-orange-700', tip, bold: true };
    if (pct < 0.8)   return { bg: 'bg-yellow-50',  text: 'text-yellow-700', tip };
    if (pct < 1.0)   return { bg: 'bg-blue-50',    text: 'text-blue-700',   tip };
    return               { bg: 'bg-green-50',   text: 'text-green-700',  tip, bold: true };
  };

  const exportExcel = () => {
    if (!data) return;
    setExportPct(30);
    const wb = XLSX.utils.book_new();

    // Sheet 1: resumen por persona × período (columnas = columns actuales)
    const colHeaders = columns.map(c => c.label + (c.sub ? ` ${c.sub}` : ''));
    const headerRow  = ['Persona', ...colHeaders, 'Total'];
    const rows = [headerRow];
    visible.forEach(person => {
      const row = [person.name, ...columns.map(col => sumH(person.byDay, col.dates) || ''), Math.round(Object.values(person.byDay).reduce((s, h) => s + h, 0) * 10) / 10];
      rows.push(row);
    });
    // Fila Total equipo
    rows.push(['Total equipo', ...columns.map(col => Math.round(visible.reduce((s, p) => s + sumH(p.byDay, col.dates), 0) * 10) / 10 || ''), totalH]);
    const ws1 = XLSX.utils.aoa_to_sheet(rows);
    ws1['!cols'] = [{ wch: 28 }, ...columns.map(() => ({ wch: 10 })), { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws1, 'Registro por período');

    setExportPct(60);

    // Sheet 2: detalle por persona × ticket × día (zoom día solamente tiene sentido)
    const allWkDates = allDates.filter(d => !isWknd(d));
    const detailHeaders = ['Persona', 'Ticket', 'Título', ...allWkDates, 'Total'];
    const detailRows = [detailHeaders];
    visible.forEach(person => {
      Object.values(person.byTicket || {}).forEach(tk => {
        const tkTotal = Math.round(Object.values(tk.byDay).reduce((s, h) => s + h, 0) * 10) / 10;
        if (tkTotal === 0) return;
        detailRows.push([person.name, tk.key, tk.title, ...allWkDates.map(d => tk.byDay[d] || ''), tkTotal]);
      });
    });
    const ws2 = XLSX.utils.aoa_to_sheet(detailRows);
    ws2['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 40 }, ...allWkDates.map(() => ({ wch: 7 })), { wch: 8 }];
    XLSX.utils.book_append_sheet(wb, ws2, 'Detalle por ticket');

    setExportPct(90);

    // Sheet 3: GAP por persona × día hábil
    const gapHeaders = ['Persona', ...allWkDates.map(d => { const dt = new Date(d + 'T00:00:00'); return `${dt.getDate()}/${dt.getMonth() + 1} ${GA_DAYS[dt.getDay()]}`; }), 'Total reg.', 'Esperado', 'GAP'];
    const gapRows = [gapHeaders];
    visible.forEach(person => {
      const reg   = Math.round(Object.values(person.byDay).reduce((s, h) => s + h, 0) * 10) / 10;
      const exp   = allWkDates.length * 8;
      const gap   = Math.round((exp - reg) * 10) / 10;
      gapRows.push([person.name, ...allWkDates.map(d => person.byDay[d] || 0), reg, exp, gap]);
    });
    const ws3 = XLSX.utils.aoa_to_sheet(gapRows);
    ws3['!cols'] = [{ wch: 24 }, ...allWkDates.map(() => ({ wch: 6 })), { wch: 10 }, { wch: 10 }, { wch: 8 }];
    XLSX.utils.book_append_sheet(wb, ws3, 'GAP de registro');

    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a'); a.href = url; a.download = `registro-periodo-${dateFrom}.xlsx`; a.click();
    setExportPct(null);
  };

  return (
    <div className="space-y-4">
      <p className="text-xs text-gray-400 leading-relaxed">
        Vista de calendario que muestra las horas registradas por persona en cada día, semana, mes o trimestre. Las celdas se colorean según el cumplimiento de la jornada laboral colombiana (8h/día hábil): rojo sin registro, naranja &lt;50%, amarillo 50–79%, azul 80–99%, verde ≥100%. Exporta a Excel con detalle por ticket y análisis de GAP.
      </p>
      {/* Controls */}
      <div className="bg-white rounded-lg shadow-sm p-4">
        <div className="flex flex-wrap gap-3 items-end">
          <div>
            <label className="block text-xs font-semibold text-gray-500 mb-1">Desde</label>
            <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} disabled={teamIds.size === 0} className="border rounded px-2 py-1.5 text-sm disabled:bg-gray-100 disabled:text-gray-400 disabled:cursor-not-allowed" />
          </div>
          <div>
            <label className="block text-xs font-semibold text-gray-500 mb-1">Hasta</label>
            <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} disabled={teamIds.size === 0} className="border rounded px-2 py-1.5 text-sm disabled:bg-gray-100 disabled:text-gray-400 disabled:cursor-not-allowed" />
          </div>
          <button onClick={load} disabled={loading || teamIds.size === 0} className="px-5 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed">
            {loading ? 'Cargando…' : '🔍 Consultar'}
          </button>
          {teamIds.size === 0 && <span className="text-xs text-amber-600 font-medium">← Selecciona un equipo primero</span>}
          <div className="flex items-center gap-1">
            <span className="text-xs font-semibold text-gray-500 mr-1">Zoom:</span>
            {[{ id: 'day', l: 'Día' }, { id: 'week', l: 'Semana' }, { id: 'month', l: 'Mes' }, { id: 'quarter', l: 'Trimestre' }].map(z => (
              <button key={z.id} onClick={() => setZoom(z.id)}
                className={`px-3 py-1.5 text-xs font-medium rounded border transition-colors ${zoom === z.id ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-gray-600 border-gray-300 hover:border-blue-400'}`}>
                {z.l}
              </button>
            ))}
          </div>
        </div>
        {error && <p className="text-xs text-red-500 mt-2">{error}</p>}
      </div>

      {/* Stats + persona filter */}
      {data && (
        <div className="flex flex-wrap gap-3 items-center">
          <div className="flex gap-2">
            <div className="bg-white rounded-lg px-4 py-2 shadow-sm">
              <div className="text-xs text-gray-400">Total horas</div>
              <div className="font-bold text-blue-700 text-sm">{totalH}h</div>
            </div>
            <div className="bg-white rounded-lg px-4 py-2 shadow-sm">
              <div className="text-xs text-gray-400">Personas</div>
              <div className="font-bold text-gray-800 text-sm">{visible.length}</div>
            </div>
            <div className="bg-red-50 border border-red-100 rounded-lg px-4 py-2 shadow-sm">
              <div className="text-xs text-red-400">Días sin registro</div>
              <div className="font-bold text-red-600 text-sm">{missDays}</div>
            </div>
            <button onClick={exportExcel} disabled={exportPct !== null}
              className="px-4 py-2 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700 disabled:opacity-75 flex items-center gap-1.5 font-medium">
              {exportPct !== null
                ? <><svg className="w-3.5 h-3.5 animate-spin shrink-0" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" strokeDasharray="40" strokeDashoffset="10"/></svg>{exportPct}%</>
                : '⬇ Excel'}
            </button>
          </div>
          <div className="flex flex-wrap gap-1 items-center ml-2">
            <span className="text-xs font-semibold text-gray-500">Equipo:</span>
            <button onClick={() => setActivePeople(new Set())}
              className={`px-2 py-0.5 rounded text-xs font-medium ${activePeople.size === 0 ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700'}`}>
              Todos
            </button>
            {persons.map(p => (
              <button key={p.accountId} onClick={() => togglePerson(p.accountId)}
                className={`px-2 py-0.5 rounded text-xs font-medium ${activePeople.has(p.accountId) ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700'}`}>
                {p.name.split(' ').slice(0, 2).join(' ')}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* Leyenda */}
      {data && (
        <div className="flex gap-3 text-xs text-gray-500 items-center flex-wrap">
          <span className="font-semibold text-gray-600">Colombia · 8h/día hábil:</span>
          <span><span className="inline-block w-3 h-3 rounded bg-red-100 mr-1 align-middle" />0h (sin registro)</span>
          <span><span className="inline-block w-3 h-3 rounded bg-orange-100 mr-1 align-middle" />&lt;50% esperado</span>
          <span><span className="inline-block w-3 h-3 rounded bg-yellow-50 border border-yellow-200 mr-1 align-middle" />50–79%</span>
          <span><span className="inline-block w-3 h-3 rounded bg-blue-50 border border-blue-200 mr-1 align-middle" />80–99%</span>
          <span><span className="inline-block w-3 h-3 rounded bg-green-50 border border-green-200 mr-1 align-middle" />≥100% (≥8h/día)</span>
          <span><span className="inline-block w-3 h-3 rounded bg-gray-100 mr-1 align-middle" />Fin de semana</span>
        </div>
      )}

      {/* KPI de registro */}
      {data && gapData.length > 0 && (() => {
        const deudores   = gapData.filter(p => p.gap > 0);
        const alDia      = gapData.filter(p => p.gap <= 0);
        const totalDeuda = Math.round(deudores.reduce((s, p) => s + p.gap, 0) * 10) / 10;
        return (
          <div className="bg-white rounded-lg shadow-sm overflow-hidden border border-gray-100">
            <button onClick={() => setShowGap(v => !v)}
              className="w-full flex items-center gap-3 px-4 py-2 hover:bg-gray-50 transition-colors">
              <span className="text-gray-300 text-xs">{showGap ? '▼' : '▶'}</span>
              <span className="text-xs font-bold text-gray-500 uppercase tracking-wide">KPI Registro</span>
              {deudores.length > 0
                ? <span className="flex items-center gap-1.5 bg-red-50 border border-red-200 text-red-700 rounded-full px-2.5 py-0.5 text-xs font-semibold">
                    <span className="w-1.5 h-1.5 rounded-full bg-red-500 inline-block" />
                    {deudores.length} con deuda · {totalDeuda}h pendientes
                  </span>
                : <span className="flex items-center gap-1.5 bg-green-50 border border-green-200 text-green-700 rounded-full px-2.5 py-0.5 text-xs font-semibold">
                    ✓ Todos al día
                  </span>}
              <span className="ml-auto text-xs text-gray-300">{workingDays}d hábiles · {expectedH}h/persona</span>
            </button>
            {showGap && (
              <div className="border-t px-4 py-3 space-y-3">
                {deudores.length > 0 && (
                  <div>
                    <p className="text-xs font-semibold text-red-500 uppercase tracking-wide mb-2">Deben registrar</p>
                    <div className="flex flex-wrap gap-2">
                      {deudores.map(p => {
                        const urgent = p.pct < 40;
                        return (
                          <div key={p.name} className={`flex items-center gap-2 rounded-lg border px-3 py-1.5 min-w-[160px] ${urgent ? 'border-red-300 bg-red-50' : 'border-yellow-300 bg-yellow-50'}`}>
                            <div className="flex-1 min-w-0">
                              <div className={`text-xs font-semibold truncate ${urgent ? 'text-red-800' : 'text-yellow-800'}`}>{p.name.split(' ').slice(0, 2).join(' ')}</div>
                              <div className="flex items-center gap-1.5 mt-0.5">
                                <div className="w-16 bg-gray-200 rounded-full h-1 overflow-hidden">
                                  <div className={`h-full rounded-full ${urgent ? 'bg-red-400' : 'bg-yellow-400'}`} style={{ width: `${p.pct}%` }} />
                                </div>
                                <span className={`text-xs ${urgent ? 'text-red-500' : 'text-yellow-600'}`}>{p.pct}%</span>
                              </div>
                            </div>
                            <div className="text-right shrink-0">
                              <div className={`font-mono font-bold text-sm leading-none ${urgent ? 'text-red-600' : 'text-yellow-600'}`}>-{p.gap}h</div>
                              <div className="text-xs text-gray-400 mt-0.5">{p.registered}h reg.</div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
                {alDia.length > 0 && (
                  <div>
                    <p className="text-xs font-semibold text-green-600 uppercase tracking-wide mb-2">Al día</p>
                    <div className="flex flex-wrap gap-1.5">
                      {alDia.map(p => (
                        <div key={p.name} className="flex items-center gap-2 rounded-full border border-green-200 bg-green-50 px-3 py-1 text-xs text-green-700">
                          <span className="font-medium">{p.name.split(' ').slice(0, 2).join(' ')}</span>
                          <span className="font-mono font-bold">{p.registered}h</span>
                          <span className="text-green-400">{p.pct}%</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        );
      })()}

      {/* Tabla Gantt */}
      {data && visible.length > 0 && (
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
            <table className="text-xs border-collapse" style={{ minWidth: '100%' }}>
              <thead>
                {zoom === 'day' && (
                  <tr className="bg-gray-100 border-b border-gray-200">
                    <th className="sticky left-0 z-20 bg-gray-100 border-r border-gray-200 px-3 py-1.5 min-w-[200px]" />
                    {monthGroups.map(mg => {
                      const [y, mo] = mg.key.split('-');
                      return (
                        <th key={mg.key} colSpan={mg.count} className="px-1 py-1.5 text-center font-semibold text-gray-600 border-r border-gray-300 uppercase tracking-wide">
                          {GA_MONTH[parseInt(mo) - 1]} {y}
                        </th>
                      );
                    })}
                    <th className="bg-gray-100 px-2 min-w-[60px]" />
                  </tr>
                )}
                <tr className="bg-gray-50 border-b border-gray-200">
                  <th className="sticky left-0 z-20 bg-gray-50 border-r border-gray-200 px-3 py-2 text-left font-semibold text-gray-600 min-w-[200px]">
                    Persona / Ticket
                  </th>
                  {columns.map(col => (
                    <th key={col.key} className={`px-1 py-1.5 text-center font-medium border-r border-gray-100 min-w-[52px] whitespace-nowrap ${col.isWknd ? 'text-gray-300 bg-gray-50' : 'text-gray-500'}`}>
                      <div className="font-semibold">{col.label}</div>
                      {col.sub && <div className="text-gray-400 font-normal">{col.sub}</div>}
                    </th>
                  ))}
                  <th className="bg-gray-50 border-l border-gray-200 px-3 py-2 text-center font-semibold text-gray-600 min-w-[60px]">Total</th>
                </tr>
              </thead>
              <tbody>
                {visible.map(person => {
                  const isExp = expanded.has(person.accountId);
                  const personTotal = Math.round(Object.values(person.byDay).reduce((s, h) => s + h, 0) * 10) / 10;
                  const tickets = Object.values(person.byTicket)
                    .map(tk => ({ ...tk, total: Math.round(Object.values(tk.byDay).reduce((s, h) => s + h, 0) * 10) / 10 }))
                    .sort((a, b) => b.total - a.total);

                  return (
                    <React.Fragment key={person.accountId}>
                      <tr className="border-b border-gray-100 hover:bg-blue-50 cursor-pointer" onClick={() => toggleExpand(person.accountId)}>
                        <td className="sticky left-0 z-10 bg-white border-r border-gray-200 px-3 py-2 hover:bg-blue-50">
                          <div className="flex items-center gap-2">
                            <span className="text-gray-300 text-xs w-3 shrink-0">{isExp ? '▾' : '▸'}</span>
                            <span className="w-6 h-6 rounded-full bg-blue-100 text-blue-700 font-bold text-xs flex items-center justify-center shrink-0">
                              {person.name.charAt(0)}
                            </span>
                            <span className="font-semibold text-gray-800 truncate">{person.name}</span>
                          </div>
                        </td>
                        {columns.map(col => {
                          const hrs = sumH(person.byDay, col.dates);
                          const wDays = col.dates.filter(d => !isWknd(d));
                          const { bg, text, tip, bold } = cellClass(hrs, col);
                          const cls = `text-center px-1 py-2 border-r border-gray-100 ${bg} ${text} ${bold ? 'font-semibold' : ''}`;
                          return (
                            <td key={col.key} className={cls} title={tip}>
                              {col.isWknd ? '' : hrs > 0 ? `${hrs}h` : (wDays.length > 0 ? <span className="opacity-30">—</span> : '—')}
                            </td>
                          );
                        })}
                        <td className="px-3 py-2 text-center font-bold text-blue-700 bg-blue-50 border-l border-gray-200">
                          {personTotal}h
                        </td>
                      </tr>
                      {isExp && tickets.map(tk => (
                        <tr key={`${person.accountId}-${tk.key}`} className="border-b border-gray-100 bg-gray-50">
                          <td className="sticky left-0 z-10 bg-gray-50 border-r border-gray-200 px-3 py-1.5 pl-10">
                            <div className="flex items-center gap-1.5">
                              <a href={`${JIRA_URL}/browse/${tk.key}`} target="_blank" rel="noreferrer"
                                className="text-blue-600 font-mono font-semibold hover:underline shrink-0 text-xs" onClick={e => e.stopPropagation()}>
                                {tk.key}
                              </a>
                              <span className="text-gray-500 truncate max-w-[110px] text-xs" title={tk.title}>{tk.title}</span>
                            </div>
                          </td>
                          {columns.map(col => {
                            const hrs = sumH(tk.byDay, col.dates);
                            return (
                              <td key={col.key} className={`text-center px-1 py-1.5 border-r border-gray-100 ${col.isWknd ? 'text-gray-100' : hrs > 0 ? 'text-blue-500' : 'text-gray-200'}`}>
                                {hrs > 0 ? `${hrs}h` : '—'}
                              </td>
                            );
                          })}
                          <td className="px-3 py-1.5 text-center text-blue-600 font-semibold border-l border-gray-200">{tk.total}h</td>
                        </tr>
                      ))}
                    </React.Fragment>
                  );
                })}
              </tbody>
              <tfoot>
                <tr className="bg-gray-100 border-t-2 border-gray-300">
                  <td className="sticky left-0 z-10 bg-gray-100 border-r border-gray-200 px-3 py-2 font-bold text-gray-700">Total equipo</td>
                  {columns.map(col => {
                    const t = Math.round(visible.reduce((s, p) => s + sumH(p.byDay, col.dates), 0) * 10) / 10;
                    const wDays = col.dates.filter(d => !isWknd(d));
                    const expectedTeam = wDays.length * 8 * visible.length;
                    const pct = expectedTeam > 0 ? t / expectedTeam : 0;
                    const footCls = col.isWknd ? 'text-gray-200' : t === 0 ? 'text-gray-300' : pct < 0.5 ? 'text-orange-600' : pct < 0.8 ? 'text-yellow-600' : pct < 1.0 ? 'text-blue-700' : 'text-green-700';
                    return (
                      <td key={col.key} className={`text-center px-1 py-2 border-r border-gray-100 font-semibold ${footCls}`}>
                        {t > 0 ? `${t}h` : '—'}
                      </td>
                    );
                  })}
                  <td className="px-3 py-2 text-center font-bold text-blue-800 bg-blue-100 border-l border-gray-200">{totalH}h</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {data && visible.length === 0 && (
        <div className="bg-white rounded-lg shadow-sm p-8 text-center text-gray-400">Sin registros para el período seleccionado.</div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// TAB 4: ANÁLISIS DE CUMPLIMIENTO (legacy)
// ══════════════════════════════════════════════════════════════════════════════
function ComplianceTab({ teamIds }) {
  const [dateFrom,    setDateFrom]    = useState('2026-01-01');
  const [dateTo,      setDateTo]      = useState('2026-03-31');
  const [planned,     setPlanned]     = useState([]);   // rows from Excel
  const [jiraData,    setJiraData]    = useState(null); // fetched from Jira
  const [loading,     setLoading]     = useState(false);
  const [error,       setError]       = useState('');
  const [overrides,   setOverrides]   = useState({});   // {key: {resultado, justificacion}}
  const [messages,    setMessages]    = useState([]);
  const [chatInput,   setChatInput]   = useState('');
  const [chatLoading, setChatLoading] = useState(false);
  const [wordText,    setWordText]    = useState('');
  const [showUnplan,  setShowUnplan]  = useState(false);
  const [exportPct,   setExportPct]   = useState(null);
  const chatEndRef = useRef(null);

  // Parse Excel file
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb   = XLSX.read(ev.target.result, { type: 'array' });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

        // Normalize column names (lower, trim, remove special chars)
        const norm = key => key.toLowerCase().trim().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\s+/g,'');
        const normalized = rows.map(row => {
          const r = {};
          Object.entries(row).forEach(([k, v]) => { r[norm(k)] = v; });
          return {
            funcionalidad:   r['funcionalidad'] || '',
            tarea:           r['tarea'] || '',
            descripcion:     r['descripciongeneral'] || r['descripcion'] || '',
            aporte:          r['aporte'] || '',
            responsable:     r['responsable'] || '',
            tiempoBack:      parseFloat(r['tiempoback'] || r['back'] || 0) || 0,
            tiempoFront:     parseFloat(r['tiempofront'] || r['front'] || 0) || 0,
            qa:              parseFloat(r['qa'] || 0) || 0,
            despliegue:      parseFloat(r['despliegue'] || 0) || 0,
            total:           parseFloat(r['total'] || 0) || 0,
            estatusFront:    r['estatusfront'] || r['estadofront'] || '',
            estatusBack:     r['estatusback']  || r['estadoback']  || '',
            responsableBack: r['responsableback']  || '',
            responsableFront:r['responsablefront'] || '',
            epica:           r['epica'] || '',
            ticket:          (r['ticket'] || '').toString().trim().toUpperCase(),
          };
        }).filter(r => r.ticket.match(/[A-Z]+-\d+/));

        setPlanned(normalized);
        setJiraData(null);
        setOverrides({});
        setMessages([]);
      } catch (err) {
        setError('Error leyendo Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Fetch Jira data for planned tickets
  const loadJira = async () => {
    if (planned.length === 0) return;
    setLoading(true); setError('');
    try {
      const keys = [...new Set(planned.map(p => p.ticket).filter(Boolean))].join(',');
      const r = await apiFetch(`/api/team-time?dateFrom=${dateFrom}&dateTo=${dateTo}&keys=${encodeURIComponent(keys)}`);
      const d = await r.json();
      if (!r.ok) throw new Error(d.error);

      // Also fetch unplanned tickets worked on by team
      const ids = [...teamIds].join(',');
      const r2  = await apiFetch(`/api/team-time?dateFrom=${dateFrom}&dateTo=${dateTo}&userIds=${encodeURIComponent(ids)}`);
      const d2  = await r2.json();

      setJiraData({ planned: d.tickets || [], all: d2.tickets || [] });
      setOverrides({});
      // Auto-send initial analysis message
      setTimeout(() => sendToChat('__init__', d.tickets || [], d2.tickets || []), 500);
    } catch (e) { setError(e.message); }
    finally     { setLoading(false); }
  };

  // Comparison logic
  const comparison = !jiraData ? [] : planned.map(p => {
    const jira = jiraData.planned.find(t => t.key === p.ticket);
    return {
      ...p,
      plannedKey:  p.ticket,
      jira,
      realKey:     jira?.key || p.ticket,
      realHours:   jira?.wlRangeHours || 0,
      status:      jira?.status || 'No encontrado',
      assignee:    jira?.assignee || p.responsable,
      resultado:   overrides[p.ticket]?.resultado || (jira ? autoResultado(jira) : 'Sin trabajo'),
      justificacion: overrides[p.ticket]?.justificacion || '—',
    };
  });

  const plannedKeys = new Set(planned.map(p => p.ticket));
  const unplanned   = jiraData ? jiraData.all.filter(t => !plannedKeys.has(t.key) && t.wlRangeHours > 0) : [];

  const setOv = (key, field, val) =>
    setOverrides(p => ({ ...p, [key]: { ...(p[key] || {}), [field]: val } }));

  const resCount = {};
  RESULTADO_OPTS.forEach(r => { resCount[r] = 0; });
  comparison.forEach(c => { resCount[c.resultado] = (resCount[c.resultado] || 0) + 1; });

  const totalEstH  = Math.round(comparison.reduce((s, c) => s + (c.total || 0), 0) * 100) / 100;
  const totalRegH  = Math.round(comparison.reduce((s, c) => s + (c.realHours || 0), 0) * 100) / 100;

  // ── Chat ─────────────────────────────────────────────────────────────────
  const sendToChat = async (content, plannedTix, allTix) => {
    const userMsg = content === '__init__' ? null : { role: 'user', content };
    const msgs    = content === '__init__'
      ? [{ role: 'user', content: 'Analiza el cumplimiento del período. Dame un resumen ejecutivo con métricas clave, tickets logrados, parciales, no logrados y nuevos. Identifica riesgos y recomendaciones.' }]
      : [...messages, userMsg];

    if (content !== '__init__') setMessages(msgs);
    setChatLoading(true);

    const ctx = {
      planned:  plannedTix || planned,
      jiraData: allTix     || (jiraData?.planned || []),
      dateFrom, dateTo,
    };

    try {
      const r = await apiFetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ messages: msgs, context: ctx }),
      });

      let fullText = '';
      const reader = r.body.getReader();
      const decoder = new TextDecoder();

      const assistantMsg = { role: 'assistant', content: '' };
      setMessages(prev => [...prev, assistantMsg]);

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        const chunk = decoder.decode(value);
        const lines = chunk.split('\n');
        for (const line of lines) {
          if (!line.startsWith('data: ')) continue;
          const payload = line.slice(6);
          if (payload === '[DONE]') break;
          try {
            const json = JSON.parse(payload);
            if (json.text) {
              fullText += json.text;
              setMessages(prev => {
                const updated = [...prev];
                updated[updated.length - 1] = { role: 'assistant', content: fullText };
                return updated;
              });
            }
          } catch { /* skip */ }
        }
      }

      // Check if Word document was generated
      const wordMatch = fullText.match(/=== INICIO DOCUMENTO WORD ===([\s\S]+?)=== FIN DOCUMENTO WORD ===/);
      if (wordMatch) setWordText(wordMatch[1].trim());
    } catch (e) {
      setMessages(prev => [...prev, { role: 'assistant', content: `❌ Error: ${e.message}` }]);
    } finally {
      setChatLoading(false);
      setTimeout(() => chatEndRef.current?.scrollIntoView({ behavior: 'smooth' }), 100);
    }
  };

  const handleChatSubmit = (e) => {
    e.preventDefault();
    if (!chatInput.trim() || chatLoading) return;
    const txt = chatInput.trim();
    setChatInput('');
    sendToChat(txt);
  };

  // ── Export ───────────────────────────────────────────────────────────────
  const exportExcel = async () => {
    try {
      setExportPct(30);
      const r = await apiFetch('/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type: 'excel', data: { comparison, unplanned, hierarchy: [], dateFrom, dateTo } }),
      });
      setExportPct(90);
      const blob = await r.blob();
      const url  = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'cumplimiento.xlsx'; a.click();
    } finally {
      setExportPct(null);
    }
  };

  const exportWord = async () => {
    if (!wordText) return alert('Primero pídele a Claude que genere el documento Word.');
    const r = await apiFetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'word', data: { docText: wordText, dateFrom, dateTo } }),
    });
    const blob = await r.blob();
    const url  = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'reporte-cumplimiento.docx'; a.click();
  };

  return (
    <div className="flex gap-4 h-[calc(100vh-120px)]">
      {/* Left panel: Upload + comparison */}
      <div className="flex-1 min-w-0 overflow-y-auto space-y-4 pr-1">

        {/* Step 1: Upload */}
        <div className="bg-white rounded-lg shadow-sm p-4">
          <h2 className="text-sm font-bold text-gray-700 mb-3">① Subir Excel de planificación</h2>
          <div className="flex flex-wrap gap-3 items-end">
            <div>
              <label className="block text-xs font-semibold text-gray-500 mb-1">Archivo Excel (.xlsx)</label>
              <input type="file" accept=".xlsx,.xls" onChange={handleFile} className="text-sm border rounded px-2 py-1" />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 mb-1">Desde</label>
              <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} className="border rounded px-2 py-1 text-sm" />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 mb-1">Hasta</label>
              <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} className="border rounded px-2 py-1 text-sm" />
            </div>
            {planned.length > 0 && (
              <button onClick={loadJira} disabled={loading}
                className="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700 disabled:bg-gray-400">
                {loading ? 'Cargando…' : '② Cargar datos Jira'}
              </button>
            )}
            {jiraData && (
              <>
                <button onClick={exportExcel} disabled={exportPct !== null}
                  className="px-3 py-2 bg-green-600 text-white text-sm rounded hover:bg-green-700 disabled:opacity-75 flex items-center gap-1.5 min-w-[90px] justify-center">
                  {exportPct !== null
                    ? <><svg className="w-3.5 h-3.5 animate-spin shrink-0" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" strokeDasharray="40" strokeDashoffset="10"/></svg>{exportPct}%</>
                    : '⬇ Excel'}
                </button>
                <button onClick={exportWord}  className="px-3 py-2 bg-indigo-600 text-white text-sm rounded hover:bg-indigo-700" title={wordText ? 'Descargar Word' : 'Pídele a Claude que genere el documento primero'}>
                  📄 Word{!wordText && ' (pide a Claude)'}
                </button>
              </>
            )}
          </div>
          {planned.length > 0 && <p className="text-xs text-green-600 mt-2">✓ {planned.length} tickets cargados del Excel</p>}
          {error && <p className="text-xs text-red-600 mt-2">{error}</p>}
        </div>

        {/* Step 2: Comparison */}
        {jiraData && (
          <>
            {/* Summary */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-2">
              {[
                { label: 'Planificados', val: planned.length, clr: '#3b82f6' },
                { label: 'Se logró', val: resCount['Se logró'], clr: '#16a34a' },
                { label: 'Parcial', val: resCount['Se logró parcialmente'], clr: '#ca8a04' },
                { label: 'No logrado', val: (resCount['No se logró']||0) + (resCount['Sin trabajo']||0), clr: '#dc2626' },
                { label: 'No planificados', val: unplanned.length, clr: '#7c3aed' },
              ].map(c => (
                <div key={c.label} className="bg-white rounded shadow-sm p-3 border-l-4" style={{ borderColor: c.clr }}>
                  <div className="text-xs text-gray-400">{c.label}</div>
                  <div className="text-2xl font-bold" style={{ color: c.clr }}>{c.val}</div>
                </div>
              ))}
            </div>

            {/* Hours bar */}
            <div className="bg-white rounded shadow-sm p-3 flex gap-6 items-center text-sm">
              <div><div className="text-xs text-gray-400">Planificado</div><div className="font-bold text-gray-700">{totalEstH}h</div></div>
              <div><div className="text-xs text-gray-400">Registrado</div><div className="font-bold text-blue-700">{totalRegH}h</div></div>
              <div><div className="text-xs text-gray-400">Diferencia</div>
                <div className={`font-bold ${totalRegH - totalEstH > 0 ? 'text-red-600' : 'text-green-600'}`}>
                  {totalRegH - totalEstH > 0 ? '+' : ''}{Math.round((totalRegH - totalEstH) * 100) / 100}h
                </div>
              </div>
              <div className="flex-1 max-w-xs">
                <div className="flex justify-between text-xs text-gray-400 mb-1">
                  <span>Ejecución</span><span>{totalEstH > 0 ? Math.min(200, Math.round(totalRegH / totalEstH * 100)) : 0}%</span>
                </div>
                <div className="w-full h-2 bg-gray-200 rounded-full overflow-hidden">
                  <div className="h-full rounded-full bg-blue-500" style={{ width: `${totalEstH > 0 ? Math.min(100, Math.round(totalRegH / totalEstH * 100)) : 0}%` }} />
                </div>
              </div>
              {unplanned.length > 0 && (
                <button onClick={() => setShowUnplan(v => !v)}
                  className={`px-3 py-1 text-xs rounded-lg border ${showUnplan ? 'bg-purple-100 text-purple-700' : 'text-gray-500 border-gray-300'}`}>
                  {showUnplan ? '▼' : '▶'} {unplanned.length} no planificados
                </button>
              )}
            </div>

            {/* Main comparison table */}
            <div className="bg-white rounded shadow-sm overflow-hidden">
              <div className="px-4 py-2 border-b bg-gray-50 text-xs font-bold text-gray-600">📊 Comparativo Planificado vs Real</div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead className="bg-gray-50 border-b">
                    <tr>
                      {['Ticket','Tarea / Iniciativa','Épica','H.Est','H.Reg','Δ','Estado','Resultado','Justificación'].map(h => (
                        <th key={h} className="px-3 py-2 text-left font-semibold text-gray-500 whitespace-nowrap">{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {[...comparison].sort((a,b) => {
                      const o = { 'No se logró':0,'Sin trabajo':1,'Se replanificó':2,'Se logró parcialmente':3,'Se logró':4 };
                      return (o[a.resultado]??5) - (o[b.resultado]??5);
                    }).map(c => {
                      const rClr = RES_CLR[c.resultado] || RES_CLR['Sin trabajo'];
                      const diff = Math.round((c.realHours - c.total) * 100) / 100;
                      return (
                        <tr key={c.plannedKey} style={{ backgroundColor: rClr.bg + '33' }}>
                          <td className="px-3 py-2 whitespace-nowrap">
                            <a href={`${JIRA_URL}/browse/${c.plannedKey}`} target="_blank" rel="noreferrer" className="text-blue-600 font-bold hover:underline">{c.plannedKey}</a>
                          </td>
                          <td className="px-3 py-2 max-w-[180px]"><div className="truncate" title={c.tarea}>{c.tarea || '—'}</div></td>
                          <td className="px-3 py-2 text-gray-500 whitespace-nowrap">{c.epica || '—'}</td>
                          <td className="px-3 py-2 font-mono text-gray-500">{c.total > 0 ? `${c.total}h` : '—'}</td>
                          <td className="px-3 py-2 font-mono font-bold text-blue-700">{c.realHours > 0 ? `${c.realHours}h` : '—'}</td>
                          <td className="px-3 py-2 font-mono font-bold">
                            {c.total > 0 ? <span className={diff > 0 ? 'text-red-600' : diff < 0 ? 'text-green-600' : 'text-gray-400'}>{diff > 0 ? `+${diff}h` : diff < 0 ? `${diff}h` : '='}</span> : '—'}
                          </td>
                          <td className="px-3 py-2 whitespace-nowrap"><StatusBadge s={c.status} /></td>
                          <td className="px-3 py-2">
                            <select value={c.resultado} onChange={e => setOv(c.plannedKey, 'resultado', e.target.value)}
                              className="border rounded px-1.5 py-1 text-xs w-full" style={{ backgroundColor: rClr.bg, color: rClr.tx, borderColor: rClr.bd }}>
                              {RESULTADO_OPTS.map(o => <option key={o} value={o}>{o}</option>)}
                            </select>
                          </td>
                          <td className="px-3 py-2" style={{ minWidth: '160px' }}>
                            <select value={c.justificacion} onChange={e => setOv(c.plannedKey, 'justificacion', e.target.value)}
                              className="border rounded px-1.5 py-1 text-xs w-full">
                              {JUSTIF_OPTS.map(o => <option key={o} value={o}>{o}</option>)}
                            </select>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot className="border-t-2 border-gray-300 bg-gray-50">
                    <tr>
                      <td colSpan={3} className="px-3 py-2 font-bold text-xs text-gray-600">TOTAL ({comparison.length})</td>
                      <td className="px-3 py-2 font-mono font-bold text-gray-600">{totalEstH}h</td>
                      <td className="px-3 py-2 font-mono font-bold text-blue-700">{totalRegH}h</td>
                      <td colSpan={4} className="px-3 py-2 text-xs text-gray-400">
                        ✅ {resCount['Se logró']} · 🟡 {resCount['Se logró parcialmente']} · 🔴 {resCount['No se logró']} · ⬜ {resCount['Sin trabajo']} · 🔵 {resCount['Se replanificó']}
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            {/* Unplanned section */}
            {showUnplan && unplanned.length > 0 && (
              <div className="bg-white rounded shadow-sm overflow-hidden">
                <div className="px-4 py-2 border-b bg-purple-50 text-xs font-bold text-purple-700">
                  🆕 Actividades no planificadas con tiempo registrado ({unplanned.length})
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead className="bg-gray-50 border-b">
                      <tr>{['Ticket','Título','Proyecto','Asignado','Estado','H. Reg.','Épica'].map(h => <th key={h} className="px-3 py-2 text-left font-semibold text-gray-500">{h}</th>)}</tr>
                    </thead>
                    <tbody className="divide-y">
                      {unplanned.sort((a,b) => b.wlRangeHours - a.wlRangeHours).map(t => (
                        <tr key={t.key} className="hover:bg-purple-50/30">
                          <td className="px-3 py-2"><a href={`${JIRA_URL}/browse/${t.key}`} target="_blank" rel="noreferrer" className="text-blue-600 font-bold hover:underline">{t.key}</a></td>
                          <td className="px-3 py-2 max-w-[200px] truncate" title={t.title}>{t.title}</td>
                          <td className="px-3 py-2 text-gray-500">{t.project}</td>
                          <td className="px-3 py-2 text-gray-600">{t.assignee}</td>
                          <td className="px-3 py-2"><StatusBadge s={t.status} /></td>
                          <td className="px-3 py-2 font-mono font-bold text-blue-700">{t.wlRangeHours}h</td>
                          <td className="px-3 py-2 text-gray-400">{t.parentKey || '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        )}
      </div>

      {/* Right panel: Chat */}
      <div className="w-96 flex-shrink-0 flex flex-col bg-white rounded-lg shadow-sm overflow-hidden">
        <div className="px-4 py-3 border-b bg-gray-50 flex items-center gap-2">
          <span className="text-lg">🤖</span>
          <div>
            <div className="text-sm font-bold text-gray-700">Análisis con Claude</div>
            <div className="text-xs text-gray-400">Sustenta el cumplimiento y genera el Word</div>
          </div>
        </div>

        {/* Messages */}
        <div className="flex-1 overflow-y-auto p-3 space-y-3 text-xs">
          {messages.length === 0 && (
            <div className="text-center text-gray-400 py-8">
              <div className="text-3xl mb-2">💬</div>
              <p>Sube el Excel y carga los datos de Jira.<br />Claude analizará automáticamente el cumplimiento.</p>
              <div className="mt-4 space-y-1 text-left text-gray-500">
                <p className="font-semibold">Puedes pedirle:</p>
                {[
                  '¿Cuáles tickets no se lograron y por qué?',
                  'El ticket NVOMS-XXXX no se terminó porque…',
                  'Genera el documento Word del reporte',
                  '¿Qué tareas se heredaron al Q2?',
                ].map(s => (
                  <button key={s} onClick={() => { setChatInput(s); }}
                    className="block w-full text-left px-2 py-1 rounded bg-gray-100 hover:bg-gray-200 text-xs">{s}</button>
                ))}
              </div>
            </div>
          )}
          {messages.map((m, i) => (
            <div key={i} className={`flex ${m.role === 'user' ? 'justify-end' : 'justify-start'}`}>
              <div className={`max-w-[90%] rounded-lg px-3 py-2 whitespace-pre-wrap leading-relaxed ${
                m.role === 'user' ? 'bg-blue-600 text-white' : 'bg-gray-100 text-gray-800'
              }`}>
                {m.content}
              </div>
            </div>
          ))}
          {chatLoading && (
            <div className="flex justify-start">
              <div className="bg-gray-100 rounded-lg px-3 py-2 text-gray-500 animate-pulse">Analizando…</div>
            </div>
          )}
          <div ref={chatEndRef} />
        </div>

        {wordText && (
          <div className="px-3 py-2 border-t bg-indigo-50">
            <button onClick={exportWord} className="w-full py-1.5 bg-indigo-600 text-white text-xs rounded hover:bg-indigo-700 font-semibold">
              📄 Descargar documento Word generado
            </button>
          </div>
        )}

        {/* Input */}
        <form onSubmit={handleChatSubmit} className="border-t p-2 flex gap-2">
          <input value={chatInput} onChange={e => setChatInput(e.target.value)}
            placeholder={jiraData ? 'Escribe aquí…' : 'Carga los datos primero'}
            disabled={!jiraData || chatLoading}
            className="flex-1 border rounded px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-blue-300 disabled:bg-gray-100" />
          <button type="submit" disabled={!jiraData || chatLoading || !chatInput.trim()}
            className="px-3 py-1.5 bg-blue-600 text-white rounded text-xs hover:bg-blue-700 disabled:bg-gray-300">
            ↑
          </button>
        </form>
      </div>
    </div>
  );
}

// ─── Shared components ────────────────────────────────────────────────────────
function FactBadge({ v }) {
  if (!v) return <span className="text-gray-300">—</span>;
  const isF = v === 'Facturables';
  return (
    <span className={`px-1.5 py-0.5 rounded text-xs font-medium whitespace-nowrap ${isF ? 'bg-green-100 text-green-700' : 'bg-orange-100 text-orange-600'}`}>
      {isF ? 'F' : 'NF'}
    </span>
  );
}

function StatusBadge({ s }) {
  const l = (s || '').toLowerCase();
  const [bg, tx] = l.includes('done') || l.includes('finaliz')
    ? ['#dcfce7','#15803d']
    : l.includes('cancel') || l.includes('rechaz')
    ? ['#fee2e2','#b91c1c']
    : l.includes('progress') || l.includes('curso')
    ? ['#dbeafe','#1d4ed8']
    : l.includes('qa') || l.includes('deploy')
    ? ['#ede9fe','#7e22ce']
    : ['#f3f4f6','#374151'];
  return <span className="px-1.5 py-0.5 rounded text-xs whitespace-nowrap" style={{ backgroundColor: bg, color: tx }}>{s || '—'}</span>;
}
