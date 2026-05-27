import { GoogleGenerativeAI } from '@google/generative-ai';

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { tickets } = req.body;
  if (!tickets?.length) return res.status(200).json({ descriptions: {} });

  const list = tickets
    .slice(0, 75) // frontend already batches in groups of 75
    .map(t => {
      const typeTag = t.type && t.type !== 'Task' ? ` [${t.type}]` : '';
      const desc = t.description ? `\n  Descripción: ${t.description.slice(0, 400)}` : '';
      return `${t.key}: ${t.title}${typeTag}${desc}`;
    })
    .join('\n');

  const prompt = `Eres un asistente técnico de proyectos de software. Para cada ticket de Jira, genera una descripción del OBJETIVO en máximo 3 líneas (~50 palabras), en español, orientada al valor de negocio o funcionalidad. Sé concreto y directo. Usa el título y la descripción del ticket para generar un objetivo preciso. Si no hay descripción, infiere el objetivo desde el título.

Responde ÚNICAMENTE con un JSON válido así:
{"KEY-1": "descripción...", "KEY-2": "descripción..."}

Tickets:
${list}`;

  try {
    const model = genAI.getGenerativeModel({ model: 'gemini-2.0-flash-lite' });
    const result = await model.generateContent(prompt);
    const text = result.response.text() || '{}';
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    const descriptions = jsonMatch ? JSON.parse(jsonMatch[0]) : {};
    return res.status(200).json({ descriptions });
  } catch (err) {
    return res.status(500).json({ error: err.message, descriptions: {} });
  }
}
