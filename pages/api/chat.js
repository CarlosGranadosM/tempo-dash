import Anthropic from '@anthropic-ai/sdk';

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

export const config = { api: { responseLimit: false } };

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { messages, context } = req.body;
  if (!messages?.length) return res.status(400).json({ error: 'messages required' });

  const systemPrompt = buildSystemPrompt(context);

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  try {
    const stream = await client.messages.stream({
      model: 'claude-sonnet-4-6',
      max_tokens: 4096,
      system: systemPrompt,
      messages: messages.map(m => ({ role: m.role, content: m.content })),
    });

    for await (const chunk of stream) {
      if (chunk.type === 'content_block_delta' && chunk.delta?.type === 'text_delta') {
        res.write(`data: ${JSON.stringify({ text: chunk.delta.text })}\n\n`);
      }
    }

    res.write('data: [DONE]\n\n');
    res.end();
  } catch (err) {
    res.write(`data: ${JSON.stringify({ error: err.message })}\n\n`);
    res.end();
  }
}

function buildSystemPrompt(ctx) {
  if (!ctx) return `Eres un analista de proyectos de software. Ayudas a analizar el cumplimiento de tickets de Jira vs la planificación inicial. Responde siempre en español.`;

  const { planned = [], jiraData = [], dateFrom, dateTo, comparison } = ctx;

  const totalPlanned  = planned.length;
  const totalWorked   = jiraData.filter(t => t.wlRangeHours > 0).length;
  const totalNewWork  = jiraData.filter(t => !planned.find(p => p.ticket === t.key)).length;
  const totalEstH     = planned.reduce((s, p) => s + (parseFloat(p.total) || 0), 0);
  const totalRegH     = jiraData.reduce((s, t) => s + t.wlRangeHours, 0);

  const plannedSummary = planned.slice(0, 50).map(p =>
    `- ${p.ticket || '?'} | ${p.tarea || ''} | Est: ${p.total || 0}h | Back: ${p.tiempoBack || 0}h | Front: ${p.tiempoFront || 0}h | QA: ${p.qa || 0}h | Epica: ${p.epica || ''}`
  ).join('\n');

  const realSummary = jiraData.slice(0, 80).map(t =>
    `- ${t.key} | ${t.title} | Reg: ${t.wlRangeHours}h | Est: ${t.estimatedHours}h | Estado: ${t.status} | Asignado: ${t.assignee}`
  ).join('\n');

  return `Eres un analista de proyectos de software especializado en reportes de cumplimiento para equipos de desarrollo. Responde siempre en español. Eres preciso, estructurado y orientado a sustentación ante gerencia.

## PERIODO ANALIZADO
${dateFrom} → ${dateTo}

## RESUMEN
- Tickets planificados: ${totalPlanned}
- Horas planificadas: ${Math.round(totalEstH * 100) / 100}h
- Tickets con tiempo registrado: ${totalWorked}
- Horas realmente registradas: ${Math.round(totalRegH * 100) / 100}h
- Tickets nuevos (no planificados): ${totalNewWork}

## PLAN INICIAL (primeros 50 tickets)
${plannedSummary}

## EJECUCIÓN REAL (primeros 80 tickets con registro)
${realSummary}

## TU ROL
Cuando el usuario te pregunte, puedes:
1. Explicar el cumplimiento del periodo.
2. Justificar por qué ciertos tickets no se completaron (el usuario te dará contexto).
3. Identificar tickets nuevos vs planificados.
4. Generar el documento Word cuando el usuario lo pida (emite el texto completo del documento en formato Markdown estructurado, comenzando con "=== INICIO DOCUMENTO WORD ===" y terminando con "=== FIN DOCUMENTO WORD ===").

El documento Word debe incluir: portada con periodo, resumen ejecutivo, tabla de tickets planificados vs ejecutados, análisis de desviaciones, justificaciones, conclusiones y métricas clave.`;
}
