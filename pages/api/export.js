import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, HeadingLevel, AlignmentType, WidthType, BorderStyle } from 'docx';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { type, data } = req.body;

  if (type === 'excel') {
    const wb = XLSX.utils.book_new();

    const { tickets = [], months = [], descriptions = {}, dateFrom: df, dateTo: dt, tabName } = data;

    // Helpers
    const DONE_STATUSES = new Set(['done','done infraestructura','finalizado','finalizada','completado','completada','closed']);
    const cumplimiento = s => DONE_STATUSES.has((s || '').toLowerCase()) ? 'Cumplido' : 'No cumplido';
    const quarters = wlByMonth => {
      const qs = [...new Set(
        Object.entries(wlByMonth || {}).filter(([, h]) => h > 0).map(([ym]) => {
          const m = parseInt(ym.split('-')[1]);
          return m <= 3 ? 'Q1' : m <= 6 ? 'Q2' : m <= 9 ? 'Q3' : 'Q4';
        })
      )].sort();
      return qs.join(', ') || '—';
    };
    const monthLabels = months.map(ym => {
      const [y, mo] = ym.split('-');
      return `${['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'][+mo - 1]}'${y.slice(2)}`;
    });

    const isDetail = tabName === 'detalle';

    let ticketRows;
    let ws1ColWidths;

    if (isDetail) {
      // Desglose: same columns as UI (no Cumplimiento), one row per author when multiple authors
      const headers = [
        'Proyecto', 'Ticket', 'Tipo',
        'Título', 'Descripción', 'Épica', 'Título Épica', 'Objetivo (IA)',
        'Registró tiempo', 'Quarter', 'Estado', 'Facturación',
        'H. Estimadas', 'H. Total',
        ...monthLabels,
      ];
      ticketRows = [headers];
      tickets.forEach(t => {
        const authorEntries = Object.entries(t.wlByAuthor || {}).filter(([, h]) => h > 0);
        const shouldSplit = authorEntries.length > 1;
        if (!shouldSplit) {
          ticketRows.push([
            t.projectName || t.project || '',
            t.key,
            t.type || 'Task',
            t.title || '',
            t.description || '',
            t._epicKey || '',
            t._epicTitle || '',
            descriptions[t.key] || '',
            (t.wlAuthors || []).join(', '),
            quarters(t.wlByMonth),
            t.status || '',
            t.facturacion || '',
            t.totalEstH  || 0,
            t.totalHours || 0,
            ...months.map(m => t.wlByMonth?.[m] || 0),
          ]);
        } else {
          authorEntries.forEach(([author, authorHours], i) => {
            const authorMths = t.wlByAuthorByMonth?.[author] || {};
            ticketRows.push([
              t.projectName || t.project || '',
              t.key,
              t.type || 'Task',
              t.title || '',
              t.description || '',
              t._epicKey || '',
              t._epicTitle || '',
              descriptions[t.key] || '',
              author,
              quarters(t.wlByMonth),
              t.status || '',
              t.facturacion || '',
              i === 0 ? (t.totalEstH || 0) : '',
              Math.round(authorHours * 100) / 100,
              ...months.map(m => Math.round((authorMths[m] || 0) * 100) / 100),
            ]);
          });
        }
      });
      ws1ColWidths = [14, 16, 12, 40, 50, 14, 30, 50, 30, 8, 18, 14, 10, 10];
    } else {
      // Tiempo Registrado: includes Cumplimiento, one row per ticket
      const headers = [
        'Proyecto', 'Ticket', 'Tipo',
        'Título', 'Descripción', 'Épica', 'Título Épica', 'Objetivo (IA)',
        'Registró tiempo', 'Quarter', 'Estado', 'Cumplimiento', 'Facturación',
        'H. Estimadas', 'H. Total',
        ...monthLabels,
      ];
      ticketRows = [headers];
      tickets.forEach(t => {
        ticketRows.push([
          t.projectName || t.project || '',
          t.key,
          t.type || 'Task',
          t.title || '',
          t.description || '',
          t._epicKey || '',
          t._epicTitle || '',
          descriptions[t.key] || '',
          (t.wlAuthors || []).join(', '),
          quarters(t.wlByMonth),
          t.status || '',
          cumplimiento(t.status),
          t.facturacion || '',
          t.totalEstH  || 0,
          t.totalHours || 0,
          ...months.map(m => t.wlByMonth?.[m] || 0),
        ]);
      });
      ws1ColWidths = [14, 16, 12, 40, 50, 14, 30, 50, 30, 8, 18, 12, 14, 10, 10];
    }

    const ws1 = XLSX.utils.aoa_to_sheet(ticketRows);
    ws1['!cols'] = [...ws1ColWidths, ...months.map(() => 8)].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws1, isDetail ? 'Desglose' : 'Tiempo Registrado');

    // Sheet GAP: registro vs esperado (solo cuando viene del tab tiempo con gapData)
    if (!isDetail && Array.isArray(data.gapData) && data.gapData.length > 0) {
      const gapRows = [['Persona', 'Días hábiles', 'H. Esperadas', 'H. Registradas', 'GAP (h)', 'Completitud (%)']];
      data.gapData.forEach(p => {
        gapRows.push([p.name, p.workingDays, p.expectedH, p.registered, p.gap, p.pct]);
      });
      const wsGap = XLSX.utils.aoa_to_sheet(gapRows);
      wsGap['!cols'] = [32, 14, 14, 16, 12, 16].map(w => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, wsGap, 'GAP Registro');
    }

    // Sheet 2: Compliance
    const compRows = [
      ['Ticket Plan', 'Iniciativa', 'Épica', 'H. Back Est.', 'H. Front Est.', 'H. QA Est.', 'H. Deploy Est.', 'H. Total Est.',
       'Ticket Real', 'H. Registradas', 'Estado Real', 'Resultado', 'Diferencia h.', 'Justificación'],
    ];
    (data.comparison || []).forEach(row => {
      compRows.push([
        row.plannedKey, row.tarea, row.epica, row.tiempoBack, row.tiempoFront, row.qa, row.despliegue, row.total,
        row.realKey || row.plannedKey, row.realHours, row.status, row.resultado,
        (row.realHours || 0) - (row.total || 0), row.justificacion || '',
      ]);
    });
    const ws2 = XLSX.utils.aoa_to_sheet(compRows);
    ws2['!cols'] = [14,35,14,10,10,10,10,10,14,12,16,18,12,40].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws2, 'Cumplimiento');

    // Sheet 3: New (unplanned) work
    const newRows = [['Ticket', 'Título', 'Proyecto', 'Asignado', 'Estado', 'H. Registradas', 'Épica', 'Sprint']];
    (data.unplanned || []).forEach(t => {
      newRows.push([t.key, t.title, t.project, t.assignee, t.status, t.wlRangeHours, t.parentKey, t.sprint]);
    });
    const ws3 = XLSX.utils.aoa_to_sheet(newRows);
    ws3['!cols'] = [14,40,12,16,16,12,14,18].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws3, 'No Planificado');

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="omni-tempo-${data.dateFrom}-${data.dateTo}.xlsx"`);
    return res.send(Buffer.from(buf));
  }

  if (type === 'word') {
    const { docText, dateFrom, dateTo } = data;

    // Parse markdown-ish document text from Claude into Word paragraphs
    const lines = (docText || '').split('\n');
    const children = [];

    for (const line of lines) {
      if (line.startsWith('# ')) {
        children.push(new Paragraph({ text: line.slice(2), heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }));
      } else if (line.startsWith('## ')) {
        children.push(new Paragraph({ text: line.slice(3), heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 150 } }));
      } else if (line.startsWith('### ')) {
        children.push(new Paragraph({ text: line.slice(4), heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 100 } }));
      } else if (line.startsWith('- ') || line.startsWith('* ')) {
        children.push(new Paragraph({ text: line.slice(2), bullet: { level: 0 } }));
      } else if (line.startsWith('  - ') || line.startsWith('  * ')) {
        children.push(new Paragraph({ text: line.slice(4), bullet: { level: 1 } }));
      } else if (line.startsWith('| ') && line.includes(' | ')) {
        // Skip markdown table rows — tables are complex, render as text
        children.push(new Paragraph({ children: [new TextRun({ text: line.replace(/\|/g, '  ').trim(), font: 'Courier New', size: 18 })] }));
      } else if (line.trim() === '' || line.trim() === '---') {
        children.push(new Paragraph({ text: '' }));
      } else if (line.startsWith('**') && line.endsWith('**')) {
        children.push(new Paragraph({ children: [new TextRun({ text: line.slice(2, -2), bold: true })] }));
      } else {
        children.push(new Paragraph({ text: line }));
      }
    }

    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            text: `Reporte de Cumplimiento — Equipo Omnisolutions`,
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({
            children: [new TextRun({ text: `Período: ${dateFrom} — ${dateTo}`, color: '555555', size: 24 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),
          ...children,
        ],
      }],
    });

    const buf = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="reporte-cumplimiento-${dateFrom}-${dateTo}.docx"`);
    return res.send(buf);
  }

  return res.status(400).json({ error: 'type must be excel or word' });
}
