// lf-client-api — LF Advisory CLIENT portal API (Cloudflare Worker).
//
// Sits BEHIND Cloudflare Access on the route  www.lfadvisory.com.au/client-api/*
// (and the apex). Access authenticates the client (one-time PIN to their email)
// and injects a signed JWT header (Cf-Access-Jwt-Assertion). This Worker:
//   1. VALIDATES that JWT (RS256, JWKS from the team cert endpoint, aud + issuer
//      + expiry) — never trusts the header blindly.
//   2. Maps the caller's email → a client + the apps they're entitled to, via
//      REGISTRY (an env var; one line per client email).
//   3. Proxies a SCOPED read/write of that client's own data file in LF
//      Advisory's SharePoint, using app-only Graph (client-credentials).
//
// The client's browser NEVER receives a Graph token, and can NEVER address
// another client's file: the filename is derived from their registry entry
// (clientId), not from anything in the request. A stranger who reaches the
// route without a valid Access session is rejected at step 1.
//
// Routes (all same-origin, no CORS; the page just fetch()es them):
//   GET /client-api/me                → { email, client, apps:[{id,name,url,desc,icon}] }
//   GET /client-api/data/<app>        → that client's stored JSON ({} if none yet)
//   PUT /client-api/data/<app>        → overwrite that client's stored JSON
//
// Env vars: SP_HOSTNAME, FOLDER, TENANT_ID, CLIENT_ID, ACCESS_TEAM, ACCESS_AUD,
//           REGISTRY (JSON). Secret: CLIENT_SECRET (Azure app secret).
// Deploy: REST API upload, same method as register-mailer (see CLAUDE.md).

const GRAPH = 'https://graph.microsoft.com/v1.0';

// App catalogue — appId → card metadata. /me returns only the entitled subset,
// so adding a client to an app is a REGISTRY edit, not a code change.
const APPS = {
  'loan-tracker': {
    name: 'Investor Loan Tracker',
    url: '/clients/loan-tracker',
    desc: 'Your register of investor loans — live interest accruals, upcoming expiries and reminder dates.',
    icon: '📈', // 📈
  },
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname.replace(/^\/client-api/, '') || '/';

    // 1. Identify the caller from the validated Access JWT.
    let email;
    try { email = await callerEmail(request, env); }
    catch (e) { return json({ error: 'Not signed in — ' + e.message }, 401); }

    // 2. Map to a client + entitlements.
    const entry = lookup(env, email);
    if (!entry) return json({
      error: 'Your account (' + email + ') is not set up for the LF Advisory client portal. Please contact LF Advisory.'
    }, 403);

    // GET /me — identity + entitled app cards
    if (path === '/me' && request.method === 'GET') {
      const apps = (entry.apps || [])
        .filter(a => APPS[a])
        .map(a => ({ id: a, ...APPS[a] }));
      return json({ email, client: entry.client, apps });
    }

    // /data/<app> — scoped read / write of this client's file
    const m = path.match(/^\/data\/([a-z0-9-]+)$/);
    if (m) {
      const app = m[1];
      if (!APPS[app]) return json({ error: 'Unknown app.' }, 404);
      if (!(entry.apps || []).includes(app)) return json({ error: 'You do not have access to this tool.' }, 403);

      // Filename is derived from the VERIFIED identity, never from the request.
      const file = entry.clientId + '-' + app + '.json';
      let token;
      try { token = await getToken(env); }
      catch (e) { return json({ error: 'Storage temporarily unavailable.' }, 502); }

      if (request.method === 'GET') {
        try { return json((await readFile(env, token, file)) || {}); }
        catch (e) { return json({ error: 'Could not read your data.' }, 502); }
      }
      if (request.method === 'PUT') {
        const bodyText = await request.text();
        if (bodyText.length > 5 * 1024 * 1024) return json({ error: 'Data too large.' }, 413);
        try { JSON.parse(bodyText); } catch { return json({ error: 'Invalid JSON body.' }, 400); }
        try { await writeFile(env, token, file, bodyText); return json({ ok: true }); }
        catch (e) { return json({ error: 'Could not save your data.' }, 502); }
      }
      return json({ error: 'Method not allowed.' }, 405);
    }

    return json({ error: 'Not found.' }, 404);
  },
};

/* ---------- registry ---------- */
function lookup(env, email) {
  let reg = {};
  try { reg = JSON.parse(env.REGISTRY || '{}'); } catch (e) {}
  return reg[email] || null;
}

/* ---------- Cloudflare Access JWT validation ---------- */
let _jwks = null, _jwksAt = 0;

async function callerEmail(request, env) {
  const cookie = request.headers.get('Cookie') || '';
  const jwt = request.headers.get('Cf-Access-Jwt-Assertion')
    || (cookie.match(/CF_Authorization=([^;]+)/) || [])[1];
  if (!jwt) throw new Error('no Access session (is Cloudflare Access enabled on this route?)');
  const payload = await verifyJwt(jwt, env);
  const email = String(payload.email || '').toLowerCase();
  if (!email) throw new Error('no email in Access token');
  return email;
}

async function verifyJwt(jwt, env) {
  const parts = jwt.split('.');
  if (parts.length !== 3) throw new Error('malformed token');
  const [h, p, s] = parts;
  const header = JSON.parse(b64urlToStr(h));
  const payload = JSON.parse(b64urlToStr(p));

  const iss = 'https://' + env.ACCESS_TEAM + '.cloudflareaccess.com';
  if (payload.iss && payload.iss !== iss) throw new Error('wrong issuer');
  if (env.ACCESS_AUD && ![].concat(payload.aud || []).includes(env.ACCESS_AUD)) throw new Error('wrong audience');
  if (payload.exp && Math.floor(Date.now() / 1000) > payload.exp) throw new Error('session expired — refresh the page');

  const key = await getKey(header.kid, env);
  const ok = await crypto.subtle.verify(
    { name: 'RSASSA-PKCS1-v1_5' }, key,
    b64urlToBytes(s), new TextEncoder().encode(h + '.' + p)
  );
  if (!ok) throw new Error('bad signature');
  return payload;
}

async function getKey(kid, env) {
  const now = Date.now();
  if (!_jwks || now - _jwksAt > 3600000) await refreshJwks(env, now);
  let jwk = _jwks.find(k => k.kid === kid);
  if (!jwk) { await refreshJwks(env, now); jwk = _jwks.find(k => k.kid === kid); } // key rotation
  if (!jwk) throw new Error('signing key not found');
  return crypto.subtle.importKey('jwk', jwk, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
}
async function refreshJwks(env, now) {
  const r = await fetch('https://' + env.ACCESS_TEAM + '.cloudflareaccess.com/cdn-cgi/access/certs');
  if (!r.ok) throw new Error('JWKS fetch ' + r.status);
  _jwks = (await r.json()).keys || [];
  _jwksAt = now;
}

function b64urlToBytes(s) {
  s = s.replace(/-/g, '+').replace(/_/g, '/');
  while (s.length % 4) s += '=';
  const bin = atob(s), bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}
function b64urlToStr(s) { return new TextDecoder().decode(b64urlToBytes(s)); }

/* ---------- Microsoft Graph (app-only) ---------- */
let _siteId = null;
async function getToken(env) {
  const body = new URLSearchParams({
    client_id: env.CLIENT_ID, client_secret: env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials'
  });
  const r = await fetch('https://login.microsoftonline.com/' + env.TENANT_ID + '/oauth2/v2.0/token',
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  if (!r.ok) throw new Error('token ' + r.status);
  return (await r.json()).access_token;
}
async function siteId(env, token) {
  if (_siteId) return _siteId;
  const r = await fetch(GRAPH + '/sites/' + env.SP_HOSTNAME, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('site ' + r.status);
  _siteId = (await r.json()).id;
  return _siteId;
}
function fileUrl(sid, env, file) {
  return GRAPH + '/sites/' + sid + '/drive/root:/'
    + encodeURIComponent(env.FOLDER) + '/' + encodeURIComponent(file) + ':/content';
}
async function readFile(env, token, file) {
  const sid = await siteId(env, token);
  const r = await fetch(fileUrl(sid, env, file), { headers: { Authorization: 'Bearer ' + token } });
  if (r.status === 404) return null;
  if (!r.ok) throw new Error('read ' + r.status);
  return r.json();
}
async function writeFile(env, token, file, text) {
  const sid = await siteId(env, token);
  const r = await fetch(fileUrl(sid, env, file),
    { method: 'PUT', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: text });
  if (!r.ok) throw new Error('write ' + r.status);
}

function json(obj, status) {
  return new Response(JSON.stringify(obj), {
    status: status || 200,
    headers: { 'content-type': 'application/json', 'cache-control': 'no-store' }
  });
}
