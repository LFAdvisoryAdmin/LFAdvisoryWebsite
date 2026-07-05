// LF Advisory — daily task email
// Cloudflare Worker. Cron 17:00 UTC = 03:00 Brisbane (AEST, UTC+10, no DST).
// Reads the register JSON from SharePoint via Microsoft Graph (app-only),
// groups active jobs by preparer, and emails each person their Overdue/Today
// (priority, top) + This week (secondary, below). Liam is CC'd on each.

const GRAPH = 'https://graph.microsoft.com/v1.0';
const WD = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const MO = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

export default {
  async scheduled(event, env, ctx) { ctx.waitUntil(run(env)); },
  async fetch(req, env) {
    const url = new URL(req.url);
    // Manual test trigger: /run?key=YOUR_TRIGGER_KEY
    if (url.pathname === '/run' && env.TRIGGER_KEY && url.searchParams.get('key') === env.TRIGGER_KEY) {
      try { return new Response(await run(env), { status: 200 }); }
      catch (e) { return new Response('Error: ' + (e && e.message || e), { status: 500 }); }
    }
    return new Response('LF register mailer — cron 17:00 UTC (3am Brisbane).', { status: 200 });
  }
};

async function run(env) {
  const token = await getToken(env);
  const data = await readRegister(env, token);
  const tasks = (data.tasks || []).filter(t => t && t.dueDate);
  const today = brisbaneToday();
  const weekEnd = addDays(today, 7);

  // Active work only: drop completed and not-ready; keep anything due up to a week out.
  const relevant = tasks.filter(t => t.stage !== 'completed' && t.stage !== 'notready' && t.dueDate <= weekEnd);

  // Group by preparer (blank -> "Unassigned").
  const groups = {};
  for (const t of relevant) {
    const who = (t.preparer || '').trim() || 'Unassigned';
    (groups[who] = groups[who] || []).push(t);
  }

  const people = parsePeople(env);
  let sent = 0;
  const notes = [];
  for (const who of Object.keys(groups)) {
    const list = groups[who];
    const overdue = list.filter(t => t.dueDate < today).sort(byDue);
    const todayL = list.filter(t => t.dueDate === today).sort(byDue);
    const week = list.filter(t => t.dueDate > today && t.dueDate <= weekEnd).sort(byDue);
    if (!overdue.length && !todayL.length && !week.length) continue;

    const to = people[who.toLowerCase()] || env.FALLBACK_EMAIL;
    const subject = `LF Advisory — your tasks, ${prettyDate(today)}`;
    const html = buildHtml(who, overdue, todayL, week, today, env.TOOL_URL);
    await sendMail(env, token, to, subject, html);
    sent++;
    if (!people[who.toLowerCase()] && who !== 'Unassigned') notes.push(`no email mapped for "${who}" -> sent to fallback`);
  }
  return `Sent ${sent} email(s) for ${today}.` + (notes.length ? ' Notes: ' + notes.join('; ') : '');
}

/* ---------- Microsoft Graph ---------- */
async function getToken(env) {
  const body = new URLSearchParams({
    client_id: env.CLIENT_ID, client_secret: env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials'
  });
  const r = await fetch(`https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`,
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  if (!r.ok) throw new Error('token ' + r.status + ' ' + await r.text());
  return (await r.json()).access_token;
}
async function readRegister(env, token) {
  const site = await gget(`${GRAPH}/sites/${env.SP_HOSTNAME}`, token);
  const url = `${GRAPH}/sites/${site.id}/drive/root:/${encodeURIComponent(env.FOLDER_NAME)}/${encodeURIComponent(env.DATA_FILE)}:/content`;
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('read register ' + r.status);
  const data = await r.json();
  return Array.isArray(data) ? { tasks: data } : data;
}
async function gget(url, token) {
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('GET ' + url + ' ' + r.status);
  return r.json();
}
async function sendMail(env, token, to, subject, html) {
  const cc = (env.CC_EMAIL && env.CC_EMAIL.toLowerCase() !== String(to).toLowerCase())
    ? [{ emailAddress: { address: env.CC_EMAIL } }] : [];
  const msg = { message: {
    subject, body: { contentType: 'HTML', content: html },
    toRecipients: [{ emailAddress: { address: to } }], ccRecipients: cc
  }, saveToSentItems: true };
  const r = await fetch(`${GRAPH}/users/${encodeURIComponent(env.SENDER_UPN)}/sendMail`,
    { method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify(msg) });
  if (!r.ok) throw new Error('sendMail ' + r.status + ' ' + await r.text());
}

/* ---------- helpers ---------- */
function brisbaneToday() { return new Date(Date.now() + 10 * 3600 * 1000).toISOString().slice(0, 10); }
function addDays(iso, n) { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function byDue(a, b) { return a.dueDate.localeCompare(b.dueDate); }
function prettyDate(iso) { const d = new Date(iso + 'T00:00:00Z'); return `${WD[d.getUTCDay()]} ${d.getUTCDate()} ${MO[d.getUTCMonth()]}`; }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])); }
function parsePeople(env) {
  const m = {}; try { const o = JSON.parse(env.PEOPLE || '{}'); for (const k in o) m[k.toLowerCase()] = o[k]; } catch (e) {}
  return m;
}
function jobLink(base, t) {
  if (!base) return '';
  const q = `client=${encodeURIComponent(t.client)}` + (t.id ? `&job=${encodeURIComponent(t.id)}` : '');
  return `${base}${base.indexOf('?') < 0 ? '?' : '&'}${q}`;
}
function taskLine(t, today, toolUrl) {
  const tag = t.dueDate < today
    ? ` <span style="color:#b0432c">(overdue — was due ${prettyDate(t.dueDate)})</span>`
    : ` <span style="color:#888">(due ${prettyDate(t.dueDate)})</span>`;
  const link = jobLink(toolUrl, t);
  const open = link ? `<td style="padding:4px 0;text-align:right;white-space:nowrap;vertical-align:top"><a href="${link}" style="color:#2d6670;text-decoration:underline;font-size:12px">Go to job &#8599;</a></td>` : '';
  return `<tr><td style="padding:4px 12px 4px 0;font-size:13px;color:#222;vertical-align:top">&bull; <strong>${esc(t.client)}</strong> — ${esc(t.task)}${tag}</td>${open}</tr>`;
}
function section(title, items, today, muted, toolUrl) {
  if (!items.length) return '';
  const color = muted ? '#666' : '#15324a';
  return `<h3 style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:${color};margin:18px 0 6px">${title}</h3>`
    + `<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;font-family:Arial,Helvetica,sans-serif">${items.map(t => taskLine(t, today, toolUrl)).join('')}</table>`;
}
function buildHtml(who, overdue, todayL, week, today, toolUrl) {
  const named = who && who !== 'Unassigned';
  const toolLink = toolUrl ? `<a href="${toolUrl}" style="color:#2d6670;text-decoration:underline">the tool</a>` : 'the tool';
  let h = `<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#222;max-width:680px">`;
  h += `<p>Good morning${named ? ' ' + esc(who) : ' — unassigned tasks'},</p>`;
  h += `<p>Your task list for <strong>${prettyDate(today)}</strong>:</p>`;
  h += section('Overdue', overdue, today, false, toolUrl);
  h += section('Due today', todayL, today, false, toolUrl);
  if (!overdue.length && !todayL.length) h += `<p style="color:#666">Nothing due today.</p>`;
  h += `<hr style="border:none;border-top:1px solid #ddd;margin:18px 0">`;
  h += section('Due this week', week, today, true, toolUrl) || `<p style="color:#999">Nothing else due this week.</p>`;
  h += `<p style="color:#555;font-size:12px;margin-top:20px">If the data above is incorrect, please go to ${toolLink} and ensure it is up to date. If you require assistance with prioritising tasks, please speak to your manager.</p>`;
  h += `<p style="color:#999;font-size:11px;margin-top:12px">Automated daily from the LF Advisory Workflow Management Tool.</p></div>`;
  return h;
}
