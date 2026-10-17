const GRAPH = 'https://graph.microsoft.com/v1.0';

async function getAccessToken() {
  const refreshToken = process.env.MS_REFRESH_TOKEN;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;
  if (!refreshToken || !clientId || !clientSecret) {
    const e = new Error('Microsoft-koppeling is nog niet geconfigureerd.');
    e.code = 'NOT_CONFIGURED';
    throw e;
  }
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope: 'offline_access Files.ReadWrite'
  });
  const r = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
    method: 'POST',
    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
    body
  });
  const d = await r.json();
  if (!r.ok || !d.access_token) throw new Error(d.error_description || 'Microsoft-token kon niet worden vernieuwd.');
  return d.access_token;
}

async function graph(path, options = {}) {
  const token = await getAccessToken();
  const r = await fetch(GRAPH + path, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {})
    }
  });
  const text = await r.text();
  let data = {};
  try { data = text ? JSON.parse(text) : {}; } catch { data = {raw:text}; }
  if (!r.ok) {
    const msg = data?.error?.message || data?.error_description || `Graph HTTP ${r.status}`;
    const e = new Error(msg); e.status = r.status; throw e;
  }
  return data;
}

function fileName() { return process.env.EXCEL_FILENAME || 'MRM-01 Openstaande schouw .xlsx'; }
function sheetName() { return process.env.EXCEL_SHEET || 'MRM-01'; }
function encodePath(s) { return encodeURIComponent(s); }

export async function getWorkbookUsedRange() {
  const file = await graph(`/me/drive/root:/${encodePath(fileName())}:`);
  const sheet = encodeURIComponent(sheetName());
  return graph(`/me/drive/items/${file.id}/workbook/worksheets/${sheet}/usedRange(valuesOnly=true)`);
}

export async function setCell(row, col, value) {
  const file = await graph(`/me/drive/root:/${encodePath(fileName())}:`);
  const sheet = encodeURIComponent(sheetName());
  return graph(`/me/drive/items/${file.id}/workbook/worksheets/${sheet}/range(address='${col}${row}:${col}${row}')`, {
    method: 'PATCH',
    body: JSON.stringify({values:[[value]]})
  });
}

export { fileName, sheetName };
