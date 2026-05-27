import fs from 'fs';
import path from 'path';

const ENV_PATH = path.join(process.cwd(), '.env.local');

const KEYS = ['JIRA_URL', 'JIRA_EMAIL', 'JIRA_API_TOKEN', 'TEMPO_API_TOKEN'];

function parseEnv(content) {
  const result = {};
  for (const line of content.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eq = trimmed.indexOf('=');
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    const val = trimmed.slice(eq + 1).trim();
    result[key] = val;
  }
  return result;
}

function serializeEnv(vars) {
  return Object.entries(vars).map(([k, v]) => `${k}=${v}`).join('\n') + '\n';
}

export default function handler(req, res) {
  if (req.method === 'GET') {
    try {
      const content = fs.existsSync(ENV_PATH) ? fs.readFileSync(ENV_PATH, 'utf8') : '';
      const all = parseEnv(content);
      const filtered = {};
      for (const k of KEYS) filtered[k] = all[k] || '';
      return res.status(200).json({ config: filtered });
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  if (req.method === 'POST') {
    try {
      const content = fs.existsSync(ENV_PATH) ? fs.readFileSync(ENV_PATH, 'utf8') : '';
      const all = parseEnv(content);
      const updates = req.body || {};
      for (const k of KEYS) {
        if (k in updates) all[k] = updates[k];
      }
      fs.writeFileSync(ENV_PATH, serializeEnv(all), 'utf8');
      return res.status(200).json({ ok: true });
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  res.status(405).end();
}
