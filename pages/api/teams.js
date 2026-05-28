import { jiraGet, OMNI_EX_MEMBERS, OMNI_SOLUTIONS_GROUP_ID } from '../../lib/jira';
import { getCredentials } from '../../lib/credentials';

const EXCLUDED = new Set([
  'org-admins', 'site-admins', 'administrators', 'confluence-users', 'tempo-SLA',
]);

export default async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).end();
  const creds = getCredentials(req);
  try {
    const { groupId } = req.query;

    if (groupId) {
      const data = await jiraGet(`/rest/api/3/group/member?groupId=${encodeURIComponent(groupId)}&maxResults=100`, creds);
      const members = (data.values || []).map(u => ({
        accountId:   u.accountId,
        displayName: u.displayName,
      }));
      if (groupId === OMNI_SOLUTIONS_GROUP_ID) {
        const existingIds = new Set(members.map(m => m.accountId));
        OMNI_EX_MEMBERS.forEach(ex => {
          if (!existingIds.has(ex.accountId)) members.push({ ...ex, isExMember: true });
        });
      }
      return res.status(200).json({ members });
    }

    const data = await jiraGet('/rest/api/3/groups/picker?maxResults=50', creds);
    const groups = (data.groups || [])
      .filter(g => !EXCLUDED.has(g.name))
      .map(g => ({ name: g.name, groupId: g.groupId }))
      .sort((a, b) => a.name.localeCompare(b.name));
    return res.status(200).json({ groups });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
