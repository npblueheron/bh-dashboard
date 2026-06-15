// src/services/powerbiApi.js
import { loginRequest } from './authConfig';

const PBI_BASE = 'https://api.powerbi.com/v1.0/myorg';

async function getToken(msalInstance) {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) throw new Error('Not logged in');
  
  const response = await msalInstance.acquireTokenSilent({
    ...loginRequest,
    account: accounts[0]
  });
  return response.accessToken;
}

async function pbiGet(msalInstance, url) {
  const token = await getToken(msalInstance);
  const res = await fetch(`${PBI_BASE}${url}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(`PBI API error: ${res.status} ${await res.text()}`);
  return res.json();
}

async function pbiPost(msalInstance, url, body) {
  const token = await getToken(msalInstance);
  const res = await fetch(`${PBI_BASE}${url}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(`PBI API error: ${res.status} ${await res.text()}`);
  return res.json();
}

// Find workspace and dataset IDs
async function findDataset(msalInstance) {
  const groups = await pbiGet(msalInstance, '/groups');
  const ws = groups.value.find(g => g.name === 'WS_BH_DATA_DEV');
  if (!ws) throw new Error('Workspace WS_BH_DATA_DEV not found');

  const datasets = await pbiGet(msalInstance, `/groups/${ws.id}/datasets`);
  const ds = datasets.value.find(d => d.name === 'Project_Dashboard_Model');
  if (!ds) throw new Error('Project_Dashboard_Model not found');

  return { workspaceId: ws.id, datasetId: ds.id };
}

// Execute DAX query against the semantic model
export async function fetchProjects(msalInstance) {
  const { workspaceId, datasetId } = await findDataset(msalInstance);

  const dax = `
    EVALUATE
    SELECTCOLUMNS(
      bh_project_details_tbl,
      "project_name", bh_project_details_tbl[project_name],
      "project_no", bh_project_details_tbl[project_no],
      "lot_no", bh_project_details_tbl[lot_no],
      "community", bh_project_details_tbl[community],
      "lot_address_full", bh_project_details_tbl[lot_address_full],
      "current_phase", bh_project_details_tbl[current_phase],
      "project_health", bh_project_details_tbl[project_health],
      "project_completion", bh_project_details_tbl[project_completion],
      "in_construction", bh_project_details_tbl[in_construction],
      "in_pre_con", bh_project_details_tbl[in_pre_con],
      "home_base_price", bh_project_details_tbl[home_base_price],
      "initial_budget", bh_project_details_tbl[initial_budget],
      "total_project_expenses", bh_project_details_tbl[total_project_expenses],
      "lot_cost", bh_project_details_tbl[lot_cost],
      "tasks_completed", bh_project_details_tbl[tasks_completed],
      "tasks_in_progress", bh_project_details_tbl[tasks_in_progress],
      "tasks_overdue", bh_project_details_tbl[tasks_overdue],
      "open_issues", bh_project_details_tbl[open_issues],
      "open_risks", bh_project_details_tbl[open_risks],
      "schedule_variance", bh_project_details_tbl[schedule_variance],
      "baseline_start_date", bh_project_details_tbl[baseline_start_date],
      "baseline_end_date", bh_project_details_tbl[baseline_end_date],
      "actual_start_date", bh_project_details_tbl[actual_start_date],
      "current_forcasted_turnover", bh_project_details_tbl[current_forcasted_turnover],
      "of_bedrooms", bh_project_details_tbl[of_bedrooms],
      "of_bathrooms", bh_project_details_tbl[of_bathrooms],
      "current_sqft", bh_project_details_tbl[current_sqft],
      "division", bh_project_details_tbl[division]
    )
  `;

  const result = await pbiPost(
    msalInstance,
    `/groups/${workspaceId}/datasets/${datasetId}/executeQueries`,
    {
      queries: [{ query: dax }],
      serializerSettings: { includeNulls: true }
    }
  );

  // Parse response into row objects
  const table = result.results?.[0]?.tables?.[0];
  if (!table) return [];

  return table.rows.map(row => {
    const obj = {};
    Object.entries(row).forEach(([key, val]) => {
      // Strip table prefix like "[project_name]" -> "project_name"
      const clean = key.replace(/^\[|\]$/g, '').replace(/^.*\[|\]$/g, '');
      obj[clean] = val;
    });
    return obj;
  });
}
