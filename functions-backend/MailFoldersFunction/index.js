// MailFoldersFunction - multi-tenant aware version
const msal = require('@azure/msal-node');
const { DefaultAzureCredential } = require('@azure/identity');
const { SecretClient } = require('@azure/keyvault-secrets');
const fetch = global.fetch || require('node-fetch');
const pLimit = require('p-limit');

const DEFAULT_TENANT = process.env.AZURE_TENANT_ID || null; // fallback tenant if needed
const BACKEND_CLIENT_ID = process.env.BACKEND_CLIENT_ID;
const KEY_VAULT_URL = process.env.KEY_VAULT_URL;
const KEY_VAULT_SECRET_NAME = process.env.KEY_VAULT_SECRET_NAME || 'backend-client-secret';
const CONCURRENCY = parseInt(process.env.CONCURRENCY || '4', 10);
const limit = pLimit(CONCURRENCY);

// caches
const tenantCcaCache = new Map(); // tenantId -> ConfidentialClientApplication
let cachedSecret = null;

/** Decode a JWT (no verification) and return payload as object.
 *  Used ONLY to read the 'tid' claim from the incoming token.
 */
function parseJwt(token) {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;
    const payload = parts[1];
    // base64 decode - replace url-safe chars
    const b64 = payload.replace(/-/g, '+').replace(/_/g, '/');
    const padded = b64.padEnd(b64.length + (4 - (b64.length % 4)) % 4, '=');
    const json = Buffer.from(padded, 'base64').toString('utf8');
    return JSON.parse(json);
  } catch (e) {
    return null;
  }
}

/** Get client secret from environment (local dev) or Key Vault (prod) */
async function getClientSecret() {
  if (process.env.BACKEND_CLIENT_SECRET) return process.env.BACKEND_CLIENT_SECRET; // local/dev
  if (cachedSecret) return cachedSecret;
  if (!KEY_VAULT_URL) throw new Error('KEY_VAULT_URL not set');
  const cred = new DefaultAzureCredential();
  const client = new SecretClient(KEY_VAULT_URL, cred);
  const resp = await client.getSecret(KEY_VAULT_SECRET_NAME);
  cachedSecret = resp.value;
  return cachedSecret;
}

/** Create or return a cached ConfidentialClientApplication for the tenant */
async function getMsalForTenant(tenantId) {
  // fallback to default tenant if none
  const tid = tenantId || DEFAULT_TENANT;
  if (!tid) throw new Error('No tenant available (tid not in token and no default AZURE_TENANT_ID configured)');

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
        loggerCallback(loglevel, message, containsPii) {
          // avoid logging PII
        },
        piiLoggingEnabled: false,
        logLevel: msal.LogLevel.Warning,
      }
    }
  };

  const cca = new msal.ConfidentialClientApplication(msalConfig);
  tenantCcaCache.set(tid, cca);
  return cca;
}

/** Acquire Graph token (OBO) using the MSAL client for the tenant in the incoming token */
async function acquireGraphTokenOnBehalf(incomingAccessToken) {
  // parse tid from incoming token payload
  const payload = parseJwt(incomingAccessToken);
  const tid = payload && (payload.tid || payload.t ? payload.tid || payload.t : null) || null;

  const cca = await getMsalForTenant(tid);
  const oboRequest = {
    oboAssertion: incomingAccessToken,
    scopes: ["https://graph.microsoft.com/.default"]
  };

  // acquireTokenOnBehalfOf will fail if the user/tenant hasn't granted required delegated permissions
  const resp = await cca.acquireTokenOnBehalfOf(oboRequest);
  if (!resp || !resp.accessToken) {
    throw new Error('OBO failed: no access token returned');
  }
  return resp.accessToken;
}

/* ---------- Graph helpers (same as before) ---------- */
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
  const childPromises = children.map(c => limit(() => computeFolderRecursive(c, graphToken)));
  const childResults = await Promise.all(childPromises);
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

/* ---------- Azure Function entry ---------- */
module.exports = async function (context, req) {
  context.log('MailFoldersFunction (multi-tenant) triggered');

  try {
    const auth = req.headers?.authorization;
    if (!auth || !auth.startsWith('Bearer ')) {
      context.res = { status: 401, body: { error: 'Missing Authorization Bearer token' } };
      return;
    }
    const incoming = auth.split(' ')[1];

    // Acquire Graph token via OBO for the tenant indicated in the incoming token
    const graphToken = await acquireGraphTokenOnBehalf(incoming);

    // Get top-level mail folders (paged)
    const topFolders = await graphGetAll('https://graph.microsoft.com/v1.0/me/mailFolders?$top=50', graphToken);

    // Recursively compute with bounded concurrency
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