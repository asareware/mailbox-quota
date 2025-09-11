// MailFoldersFunction - multi-tenant aware
const msal = require('@azure/msal-node');
const { DefaultAzureCredential } = require('@azure/identity');
const { SecretClient } = require('@azure/keyvault-secrets');
const fetch = global.fetch || require('node-fetch');
const pLimit = require('p-limit');

const DEFAULT_TENANT = process.env.AZURE_TENANT_ID || null;
const BACKEND_CLIENT_ID = process.env.BACKEND_CLIENT_ID;
const KEY_VAULT_URL = process.env.KEY_VAULT_URL;
const KEY_VAULT_SECRET_NAME = process.env.KEY_VAULT_SECRET_NAME || 'backend-client-secret';
const CONCURRENCY = parseInt(process.env.CONCURRENCY || '4', 10);
const limit = pLimit(CONCURRENCY);

const tenantCcaCache = new Map();
let cachedSecret = null;

// Decode JWT to read tid and aud (no signature verification here)
function parseJwt(token) {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;
    const payload = parts[1];
    const b64 = payload.replace(/-/g, '+').replace(/_/g, '/');
    const padded = b64.padEnd(b64.length + (4 - (b64.length % 4)) % 4, '=');
    return JSON.parse(Buffer.from(padded, 'base64').toString('utf8'));
  } catch (e) {
    return null;
  }
}

async function getClientSecret() {
  if (process.env.BACKEND_CLIENT_SECRET) return process.env.BACKEND_CLIENT_SECRET;
  if (cachedSecret) return cachedSecret;
  if (!KEY_VAULT_URL) throw new Error('KEY_VAULT_URL not configured');
  const cred = new DefaultAzureCredential();
  const client = new SecretClient(KEY_VAULT_URL, cred);
  const resp = await client.getSecret(KEY_VAULT_SECRET_NAME);
  cachedSecret = resp.value;
  return cachedSecret;
}

async function getMsalForTenant(tenantId) {
  const tid = tenantId || DEFAULT_TENANT;
  if (!tid) throw new Error('No tenant available (tid missing and no default AZURE_TENANT_ID configured)');
  if (tenantCcaCache.has(tid)) return tenantCcaCache.get(tid);

  const clientSecret = await getClientSecret();
  const msalConfig = {
    auth: {
      clientId: BACKEND_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${tid}`,
      clientSecret
    },
    system: {
      loggerOptions: {
        loggerCallback() {},
        piiLoggingEnabled: false,
        logLevel: msal.LogLevel.Warning
      }
    }
  };
  const cca = new msal.ConfidentialClientApplication(msalConfig);
  tenantCcaCache.set(tid, cca);
  return cca;
}

async function acquireGraphTokenOnBehalf(incomingAccessToken) {
  const payload = parseJwt(incomingAccessToken);
  const tid = payload && payload.tid ? payload.tid : null;
  const cca = await getMsalForTenant(tid);
  const oboRequest = { oboAssertion: incomingAccessToken, scopes: ["https://graph.microsoft.com/.default"] };
  const resp = await cca.acquireTokenOnBehalfOf(oboRequest);
  if (!resp || !resp.accessToken) throw new Error('OBO failed - no access token returned');
  return resp.accessToken;
}

async function graphGet(url, token) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }});
  if (!res.ok) {
    const txt = await res.text();
    const err = new Error(`Graph error ${res.status}: ${txt}`);
    err.status = res.status;
    throw err;
  }
  return res.json();
}

async function graphGetAll(url, token) {
  const items = [];
  let next = url;
  while (next) {
    const page = await graphGet(next, token);
    if (Array.isArray(page.value)) items.push(...page.value);
    next = page['@odata.nextLink'] || null;
  }
  return items;
}

async function computeFolderRecursive(folder, graphToken) {
  const folderBytes = Number(folder.sizeInBytes || 0);
  const children = await graphGetAll(`https://graph.microsoft.com/v1.0/me/mailFolders/${folder.id}/childFolders?$top=50`, graphToken);
  const childResults = await Promise.all(children.map(c => limit(() => computeFolderRecursive(c, graphToken))));
  const childrenSum = childResults.reduce((s, r) => s + (r.cumulativeBytes || 0), 0);
  return {
    id: folder.id,
    displayName: folder.displayName,
    bytes: folderBytes,
    cumulativeBytes: folderBytes + childrenSum,
    totalItemCount: folder.totalItemCount || 0,
    unreadItemCount: folder.unreadItemCount || 0,
    children: childResults
  };
}

function flattenFolders(node, parentPath = '') {
  const path = parentPath ? `${parentPath}/${node.displayName}` : node.displayName;
  const me = {
    id: node.id,
    displayName: node.displayName,
    folderPath: path,
    bytes: node.bytes,
    cumulativeBytes: node.cumulativeBytes,
    totalItemCount: node.totalItemCount,
    unreadItemCount: node.unreadItemCount
  };
  return [me, ...node.children.flatMap(c => flattenFolders(c, path))];
}

module.exports = async function (context, req) {
  context.log('MailFoldersFunction (multi-tenant) triggered');
  try {
    const auth = req.headers?.authorization;
    if (!auth || !auth.startsWith('Bearer ')) {
      context.res = { status: 401, body: { error: 'Missing Authorization Bearer token' } };
      return;
    }
    const incoming = auth.split(' ')[1];

    // quick validation: check aud claim is our backend API
    const payload = parseJwt(incoming) || {};
    const aud = payload.aud || payload.appid || payload.oid;
    const expectedAud = `api://${BACKEND_CLIENT_ID}`;
    if (!aud || (Array.isArray(aud) ? !aud.includes(expectedAud) : aud !== expectedAud)) {
      // allow tokens where aud equals client id as fallback
      const altAud = payload.aud === BACKEND_CLIENT_ID;
      if (!altAud) {
        context.log.warn('Token audience mismatch', { aud, expectedAud });
        // proceed — OBO will likely fail, but we let msal handle more strict checks
      }
    }

    const graphToken = await acquireGraphTokenOnBehalf(incoming);

    const topFolders = await graphGetAll('https://graph.microsoft.com/v1.0/me/mailFolders?$top=50', graphToken);

    const results = [];
    for (const f of topFolders) {
      results.push(await computeFolderRecursive(f, graphToken));
    }

    const flat = results.flatMap(r => flattenFolders(r));
    const totalBytes = flat.reduce((s, f) => s + (f.cumulativeBytes || 0), 0);

    const response = {
      totalMailboxBytes: totalBytes,
      totalMailboxGB: +(totalBytes / (1024 ** 3)).toFixed(3),
      folders: flat.map(f => ({
        id: f.id,
        displayName: f.displayName,
        folderPath: f.folderPath,
        bytes: f.bytes,
        cumulativeBytes: f.cumulativeBytes,
        bytesGB: +(f.bytes / (1024 ** 3)).toFixed(3),
        cumulativeGB: +(f.cumulativeBytes / (1024 ** 3)).toFixed(3),
        totalItemCount: f.totalItemCount
      }))
    };

    context.res = { status: 200, body: response, headers: { 'Content-Type': 'application/json' } };
  } catch (err) {
    context.log.error('Error in MailFoldersFunction:', err);
    context.res = { status: err.status || 500, body: { error: err.message || 'Server error' } };
  }
};