// Extrae las credenciales que el cliente envía por headers (cada usuario tiene
// las suyas en su navegador). Si no llegan, hace fallback a process.env para
// desarrollo local / defaults del servidor.
export function getCredentials(req) {
  const h = req.headers;
  return {
    jiraUrl:    h['x-jira-url']      || process.env.JIRA_URL || 'https://omnipro.atlassian.net',
    jiraEmail:  h['x-jira-email']    || process.env.JIRA_EMAIL || '',
    jiraToken:  h['x-jira-token']    || process.env.JIRA_API_TOKEN || '',
    tempoToken: h['x-tempo-token']   || process.env.TEMPO_API_TOKEN || '',
  };
}
