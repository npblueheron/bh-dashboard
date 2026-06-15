// api/projects/index.js
// Azure Function - runs server-side, secret never reaches browser

const TENANT_ID    = process.env.TENANT_ID;
const CLIENT_ID    = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const FABRIC_ENDPOINT = process.env.FABRIC_ENDPOINT;
const DATABASE     = process.env.DATABASE || 'lh_blueheron_dev';

let cachedToken = null;
let tokenExpiry = null;

async function getToken() {
  if (cachedToken && tokenExpiry && Date.now() < tokenExpiry) {
    return cachedToken;
  }
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://database.windows.net/.default'
  });
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });
  if (!res.ok) throw new Error(`Token fetch failed: ${await res.text()}`);
  const data = await res.json();
  cachedToken = data.access_token;
  tokenExpiry = Date.now() + (data.expires_in - 60) * 1000;
  return cachedToken;
}

async function queryFabric(token, sql) {
  const res = await fetch(`https://${FABRIC_ENDPOINT}/query`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ queryText: sql, databaseName: DATABASE })
  });
  if (!res.ok) throw new Error(`Fabric query failed: ${await res.text()}`);
  const data = await res.json();
  if (!data.results?.[0]) return [];
  const { columnNames, rows } = data.results[0];
  return rows.map(row =>
    Object.fromEntries(columnNames.map((col, i) => [col, row[i]]))
  );
}

const PROJECT_SQL = `
SELECT *
FROM (
  SELECT *,
    ROW_NUMBER() OVER (PARTITION BY project_no ORDER BY _modified_at DESC) AS rn
  FROM [lh_blueheron_dev].[smartsheets].[bronze_db_sheets_project_roll_up_summary_ascaya]
) AS ranked
WHERE rn = 1
  AND project_no IS NOT NULL
ORDER BY project_name
`;

module.exports = async function (context, req) {
  context.res = { headers: { 'Content-Type': 'application/json' } };
  try {
    const token = await getToken();
    const projects = await queryFabric(token, PROJECT_SQL);
    context.res.status = 200;
    context.res.body = JSON.stringify({ data: projects, timestamp: new Date().toISOString() });
  } catch (err) {
    context.log.error('API error:', err.message);
    context.res.status = 500;
    context.res.body = JSON.stringify({ error: err.message });
  }
};
