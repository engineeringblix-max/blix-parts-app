// BLIX Parts Review – server (Netlify Function, route /api/*)
//
// What it does
//  • Accounts: the admin creates accounts (username + password) for the mechanics.
//    Passwords are stored as scrypt hashes, never in readable form.
//  • Progress per account: drafts (status, note, responsible, date) and their photos are
//    stored in Netlify Blobs under the account, so a mechanic can continue on any phone.
//  • Send: e-mails the Excel report through Microsoft 365 (Graph, Mail.Send) and records
//    what was sent, so every phone shows which parts are done, by whom and when.
//
// Environment variables (Netlify → Site configuration → Environment variables)
//  MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET   – from the app registration (IT)
//  MAIL_FROM        – mailbox the reports are sent from
//  MAIL_TO          – where reports go (comma separated)
//  ADMIN_PASSWORD   – password for the built-in "admin" login (used to create accounts)
//  SESSION_SECRET   – long random text used to sign logins

import { getStore } from '@netlify/blobs';
import crypto from 'node:crypto';
import { XLSX_MINI } from '../lib/xlsx-mini.mjs';

const DATA = () => getStore({ name: 'parts-review', consistency: 'strong' });
const PHOTOS = () => getStore({ name: 'parts-photos', consistency: 'strong' });
const SESSION_DAYS = 120;
const MAX_PHOTO = 4_000_000;
const MAX_FILE = 3_300_000;
const PART_RE = /^p[0-9a-f]{10}(-r[1-9]\d?)?$/; // original part, or revision -r1 (= -A), -r2 (= -B) …
const parseId = (id) => { const m = /^(p[0-9a-f]{10})(?:-r([1-9]\d?))?$/.exec(id || ''); return m ? { base: m[1], k: m[2] ? +m[2] : 0 } : null; };
const revId = (base, k) => (k ? `${base}-r${k}` : base);
const openRev = (parts, id) => { const { base } = parseId(id); let k = 0; while (parts[revId(base, k)]) k++; return revId(base, k); };
const USER_RE = /^[a-z0-9._-]{2,32}$/;

const json = (status, body) => new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' } });
const err = (status, error) => json(status, { error });
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const clip = (v, n) => String(v ?? '').slice(0, n);

/* ---------- passwords & sessions ---------- */
const hashPw = (pw, salt = crypto.randomBytes(16).toString('hex')) => ({ salt, hash: crypto.scryptSync(String(pw), salt, 32).toString('hex') });
const checkPw = (pw, rec) => { if (!rec || !rec.salt) return false; const h = crypto.scryptSync(String(pw), rec.salt, 32); const b = Buffer.from(rec.hash, 'hex'); return b.length === h.length && crypto.timingSafeEqual(b, h); };
const b64u = (s) => Buffer.from(s).toString('base64url');
function sign(payload) { const body = b64u(JSON.stringify(payload)); const sig = crypto.createHmac('sha256', process.env.SESSION_SECRET).update(body).digest('base64url'); return `${body}.${sig}`; }
function verify(token) {
  if (!token || !token.includes('.')) return null;
  const [body, sig] = token.split('.');
  const good = crypto.createHmac('sha256', process.env.SESSION_SECRET).update(body).digest('base64url');
  if (sig.length !== good.length || !crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(good))) return null;
  try { const p = JSON.parse(Buffer.from(body, 'base64url').toString()); return p.exp > Date.now() ? p : null; } catch { return null; }
}
async function auth(req) {
  const p = verify((req.headers.get('authorization') || '').replace(/^Bearer /, ''));
  if (!p) return null;
  if (p.u === 'admin') return { u: 'admin', name: 'Administrator', role: 'admin' };
  const rec = await DATA().get(`users/${p.u}`, { type: 'json' });
  if (!rec || !rec.active || rec.v !== p.v) return null; // blocked or password reset → logged out
  return { u: rec.u, name: rec.name, role: rec.role };
}
const publicUser = (r) => ({ u: r.u, name: r.name, role: r.role, active: r.active, created: r.created, lastLogin: r.lastLogin || null });

/* ---------- mail ---------- */
async function graphToken(env) {
  const r = await fetch(`https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: env.MS_CLIENT_ID, client_secret: env.MS_CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
  });
  if (!r.ok) throw new Error(`Microsoft sign-in failed (${r.status}): ${await r.text()}`);
  return (await r.json()).access_token;
}
function mailHtml(p, reporter) {
  const td = (v, x = '') => `<td style="padding:6px 8px;border-bottom:1px solid #dbe3e8;${x}">${esc(v)}</td>`;
  const rows = (p.rows || []).map((r) => `<tr>${td(r.vehicle, 'font-weight:600;color:#012135')}${td(r.part, 'font-family:Consolas,monospace;font-weight:600')}${td(r.next || '', 'font-family:Consolas,monospace;color:#0b7fa6')}${td(r.status)}${td(r.priority)}${td(r.note)}${td(r.responsible)}${td(r.date, 'white-space:nowrap')}${td(r.photos || 0, 'text-align:center')}</tr>`).join('');
  const th = (t) => `<th style="text-align:left;padding:7px 8px;background:#012135;color:#fff;font-weight:600">${t}</th>`;
  return `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#14232c"><div style="background:#012135;color:#fff;padding:14px 16px;border-bottom:4px solid #0bbbef"><b style="font-size:17px">BLIX Parts Review</b><br><span style="color:#a9dff2">Report from ${esc(reporter)}${p.batches > 1 ? ` · file ${p.batch} of ${p.batches}` : ''}</span></div>
  <p>${(p.rows || []).length} part(s) reported. The attached Excel file contains every report with the part preview and the photos.</p>
  <table style="border-collapse:collapse;font-size:13px;width:100%"><tr>${th('Vehicle')}${th('Part')}${th('New revision')}${th('Status')}${th('Priority')}${th('What to change')}${th('Responsible')}${th('Date')}${th('Photos')}</tr>${rows}</table>
  <p style="color:#5d6e78;font-size:12px">Sent automatically by the BLIX Parts app.</p></div>`;
}

const mailMissing = (env) => {
  const need = env.RESEND_API_KEY ? ['RESEND_API_KEY', 'MAIL_TO'] : env.MS_TENANT_ID ? ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'MAIL_FROM', 'MAIL_TO'] : ['RESEND_API_KEY', 'MAIL_TO'];
  const miss = need.filter((k) => !env[k]); return miss.length ? miss.join(', ') : '';
};
/* send one e-mail (Resend if RESEND_API_KEY is set, otherwise Microsoft 365). Returns '' or an error text. */
async function sendMail(env, { subject, html, file, files, to: toOverride }) {
  files = files || (file ? [file] : []);
  const to = toOverride || env.MAIL_TO.split(',').map((a) => a.trim()).filter(Boolean);
  if (env.RESEND_API_KEY) {
    const r = await fetch('https://api.resend.com/emails', {
      method: 'POST', headers: { Authorization: `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ from: env.RESEND_FROM || 'BLIX Parts App <onboarding@resend.dev>', to, subject, html,
        ...(files.length ? { attachments: files.map((f) => ({ filename: f.name, content: f.b64 })) } : {}) }),
    });
    if (r.ok) return '';
    const t = await r.text(); console.error('Resend failed', r.status, t);
    let m = ''; try { m = JSON.parse(t).message || ''; } catch (e) {}
    return `The e-mail could not be sent (Resend ${r.status}${m ? ': ' + m : ''}).`;
  }
  const token = await graphToken(env);
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(env.MAIL_FROM)}/sendMail`, {
    method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ message: { subject, body: { contentType: 'HTML', content: html },
      toRecipients: to.map((address) => ({ emailAddress: { address } })),
      attachments: files.map((f) => ({ '@odata.type': '#microsoft.graph.fileAttachment', name: f.name, contentType: f.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes: f.b64 })) },
      saveToSentItems: true }),
  });
  if (r.ok) return '';
  console.error('Graph sendMail failed', r.status, await r.text()); return 'The e-mail could not be sent by Microsoft 365.';
}
const fmtNL = (iso) => new Date(iso).toLocaleString('en-GB', { timeZone: 'Europe/Amsterdam', weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }).replace(/\//g, '-');
function clearMailHtml(ev) {
  const td = (v, x = '') => `<td style="padding:6px 8px;border-bottom:1px solid #dbe3e8;vertical-align:top;${x}">${esc(v)}</td>`;
  const th = (t) => `<th style="text-align:left;padding:7px 8px;background:#b3470c;color:#fff;font-weight:600">${t}</th>`;
  const rows = ev.parts.map((p) => { const rs = ev.rows.filter((r) => r.base === p.id);
    const what = rs.length ? rs.map((r) => `${r.part}: ${r.status}${r.note ? ' – “' + r.note + '”' : ''} (sent by ${r.by}, ${fmtNL(r.at)})`).join('\n') : 'No sent reports';
    return `<tr>${td(p.name, 'font-family:Consolas,monospace;font-weight:600')}${td(p.from ? `From ${p.fromName} onwards` : 'Everything (back to the original part)')}${td(what, 'white-space:pre-line')}${td(p.drafts || 0, 'text-align:center')}</tr>`; }).join('');
  return `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#14232c"><div style="background:#b3470c;color:#fff;padding:14px 16px"><b style="font-size:17px">BLIX Parts Review – changes CLEARED</b><br><span style="color:#fde5d6">${esc(ev.by)} cleared ${ev.parts.length} part${ev.parts.length === 1 ? '' : 's'} on ${esc(fmtNL(ev.at))}</span></div>
  <p><b>Reason:</b> ${esc(ev.reason)}</p>
  <table style="border-collapse:collapse;font-size:13px;width:100%"><tr>${th('Part')}${th('What was cleared')}${th('Removed reports')}${th('Unsent drafts removed')}</tr>${rows}</table>
  <p>These parts are unlocked again for every user. Revisions created by the removed reports no longer exist in the app. The earlier report e-mails for these changes are cancelled – do not use them.</p>
  <p style="color:#5d6e78;font-size:12px">Sent automatically by the BLIX Parts app.</p></div>`;
}
async function clearedMap(store) {
  const out = {}; for (const c of await listJSON(store, 'clear/')) if (c && c.last) out[c.base] = c.last; return out;
}

async function releasedMap(store) {
  const out = {}; for (const r of await listJSON(store, 'release/')) if (r && r.id) out[r.id] = r; return out;
}
const ROLES = { admin: 'Admin', engineer: 'Engineering', mechanic: 'Mechanic' };

/* ---------- weekly backup (also on demand from Settings) ---------- */
export async function runBackup(env, trigger = 'weekly') {
  const store = DATA(), miss = mailMissing(env);
  if (miss) return `The mail service is not set up yet (missing: ${miss}).`;
  const sends = await listJSON(store, 'send/'), released = await releasedMap(store), clears = await listJSON(store, 'clear/');
  const withKey = async (prefix, f) => { const { blobs } = await store.list({ prefix }); return (await Promise.all(blobs.map(async (b) => { const v = await store.get(b.key, { type: 'json' }); return v && { ...v, ...f(b.key.split('/')) }; }))).filter(Boolean); };
  const maint = await listJSON(store, 'maint/'), users = (await listJSON(store, 'users/')).map(publicUser), drafts = await withKey('drafts/', (k) => ({ user: k[1] }));
  const team = ((await store.get('config/team', { type: 'json' })) || {}).names || [];
  const { blobs: dk } = await store.list({ prefix: 'drafts/' });
  const draftUsers = {}; dk.forEach((b) => { const u = b.key.split('/')[1]; draftUsers[u] = (draftUsers[u] || 0) + 1; });
  const X = XLSX_MINI, day = (iso) => (iso ? X.dateSerial(new Date(iso).toLocaleDateString('sv-SE', { timeZone: 'Europe/Amsterdam' })) : null);
  const H = (arr) => ({ h: 24, cells: arr.map((v) => ({ v, s: 1 })) }), T = (v) => ({ h: 28, cells: [{ v, s: 4 }] });
  const cell = (v, s = 0) => (v === null || v === undefined || v === '' ? { s } : { v, s }), dcell = (iso) => (iso ? { v: day(iso), t: 'd', s: 2 } : { s: 0 });
  const nextId = (id) => { const q = parseId(id); return q ? revId(q.base, q.k + 1) : ''; };
  const rows = []; sends.forEach((s) => (s.rows || []).forEach((r) => rows.push({ ...r, by: s.by, at: s.at })));
  rows.sort((a, b) => String(b.at).localeCompare(String(a.at)));
  const stamp = new Date().toLocaleDateString('en-GB', { timeZone: 'Europe/Amsterdam' }).replace(/\//g, '-');
  const sheets = [
    { name: 'Sent reports', freeze: 2, cols: [16, 16, 9, 26, 22, 15, 10, 48, 16, 14, 8, 16, 16].map((w) => ({ w })),
      rows: [T(`BLIX Parts Review – backup ${stamp}`), H(['Sent on', 'Sent by', 'Vehicle', 'Part', 'New revision', 'Status', 'Priority', 'What to change', 'Responsible', 'Planned date', 'Photos', 'Released in CAD', 'Released by']),
        ...rows.map((r) => { const rl = released[nextId(r.id)]; return { cells: [dcell(r.at), cell(r.by), cell(r.vehicle, 3), cell(r.part, 3), cell(r.next), cell(r.status, 3), cell(r.priority), cell(r.note), cell(r.responsible), r.dateIso ? { v: X.dateSerial(r.dateIso), t: 'd', s: 2 } : { s: 0 }, cell(r.photos || 0), rl ? dcell(rl.at) : { s: 0 }, cell(rl && rl.by)] }; })] },
    { name: 'Released in CAD', freeze: 2, cols: [16, 26, 18, 40].map((w) => ({ w })),
      rows: [T('Released in CAD'), H(['Released on', 'Part (revision)', 'Released by', 'Note']), ...Object.values(released).sort((a, b) => String(b.at).localeCompare(String(a.at))).map((r) => ({ cells: [dcell(r.at), cell(r.name, 3), cell(r.by), cell(r.note)] }))] },
    { name: 'Cleared', freeze: 2, cols: [16, 24, 18, 44, 22, 10, 10].map((w) => ({ w })),
      rows: [T('Cleared changes'), H(['Cleared on', 'Part', 'Cleared by', 'Reason', 'Scope', 'Reports', 'Drafts']), ...clears.flatMap((c) => (c.history || []).map((e) => ({ cells: [dcell(e.at), cell(e.name || c.base, 3), cell(e.by), cell(e.reason), cell(e.from ? `From ${e.fromName} onwards` : 'Everything'), cell(e.reports || 0), cell(e.drafts || 0)] })))] },
    { name: 'Maintenance', freeze: 2, cols: [16, 30, 18, 10, 10, 10, 10].map((w) => ({ w })),
      rows: [T('Maintenance resets'), H(['Reset on', 'Path', 'By', 'Parts', 'Reports', 'Drafts', 'Notes']), ...maint.sort((a, b) => String(b.at).localeCompare(String(a.at))).map((m) => ({ cells: [dcell(m.at), cell(m.scope, 3), cell(m.by), cell(m.parts), cell(m.reports), cell(m.drafts), cell(m.notes)] }))] },
    { name: 'Accounts', freeze: 2, cols: [22, 14, 14, 10, 16, 16, 14].map((w) => ({ w })),
      rows: [T('Accounts (no passwords)'), H(['Name', 'Username', 'Role', 'Active', 'Created', 'Last sign-in', 'Unsent drafts']), ...users.map((u) => ({ cells: [cell(u.name, 3), cell(u.u), cell(ROLES[u.role] || u.role), cell(u.active ? 'Yes' : 'Blocked'), dcell(u.created), dcell(u.lastLogin), cell(draftUsers[u.u] || 0)] }))] },
  ];
  const ncrAll = await listJSON(store, 'ncr/');
  sheets.push({ name: 'Problem reports', freeze: 2, cols: [16, 34, 14, 10, 12, 16, 18, 44, 30, 30, 16].map((w) => ({ w })),
    rows: [T('Problem reports'), H(['Reported', 'Title', 'Category', 'Severity', 'Status', 'Chassis', 'Part', 'Description', 'Cause', 'Action', 'Closed']),
      ...ncrAll.sort((a, b) => String(b.at).localeCompare(String(a.at))).map((n) => ({ cells: [dcell(n.at), cell(n.title, 3), cell(n.cat), cell(n.sev), cell(n.status), cell(n.chassis), cell(n.part && n.part.name), cell(n.desc), cell(n.cause), cell(n.action), dcell(n.closedAt)] }))] });
  const xlsx = Buffer.from(X.build(sheets)).toString('base64');
  const threads = await listJSON(store, 'thread/'), messages = await withKey('msg/', (k) => ({ tid: k[1] }));
  const meta = Object.fromEntries(await Promise.all([...META, 'chk', 'ncrindex', 'inspindex', 'vehicles'].map(async (k) => [k, await getMeta(store, k)])));
  const ncr = await listJSON(store, 'ncr/'), insp = await listJSON(store, 'insp/'), instr = await withKey('instr/', (k) => ({ base: k[1] }));
  const raw = { exportedAt: new Date().toISOString(), trigger, sends, released, clears, maint, users, team, drafts, threads, messages, meta, ncr, insp, instr };
  const json64 = Buffer.from(JSON.stringify(raw)).toString('base64');
  const html = `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#14232c"><div style="background:#012135;color:#fff;padding:14px 16px;border-bottom:4px solid #0bbbef"><b style="font-size:17px">BLIX Parts Review – ${trigger === 'weekly' ? 'weekly' : 'manual'} backup</b><br><span style="color:#a9dff2">${esc(fmtNL(raw.exportedAt))}</span></div>
    <p>${rows.length} sent report${rows.length === 1 ? '' : 's'} · ${Object.keys(released).length} released in CAD · ${dk.length} unsent draft${dk.length === 1 ? '' : 's'} · ${users.length} account${users.length === 1 ? '' : 's'}.</p>
    <p>Attached: the overview in Excel, and a full data file (.json) that can be used to restore the app if ever needed. Keep this e-mail.</p><p style="color:#5d6e78;font-size:12px">Sent automatically by the BLIX Parts app.</p></div>`;
  const d = new Date().toISOString().slice(0, 10);
  return sendMail(env, { subject: `BLIX parts – ${trigger === 'weekly' ? 'weekly' : 'manual'} backup – ${stamp}`, html,
    files: [{ name: `BLIX_Parts_Backup_${d}.xlsx`, b64: xlsx }, { name: `BLIX_Parts_Backup_${d}.json`, b64: json64, type: 'application/json' }] });
}

/* ---------- v12 helpers: combined meta docs, audit log ---------- */
const META = ['partx', 'stock', 'dwg', 'appr', 'fit', 'veh', 'done'];
const getMeta = async (store, k) => (await store.get(`meta/${k}`, { type: 'json' })) || {};
async function updMeta(store, k, fn) { const m = await getMeta(store, k); const r = fn(m); await store.setJSON(`meta/${k}`, m); return r; }
async function audit(store, id, me, act, detail = '') {
  const q = parseId(id), base = q ? q.base : String(id || 'general'), at = new Date().toISOString();
  await store.setJSON(`audit/${base}/${at}-${crypto.randomBytes(2).toString('hex')}`, { at, by: me.name, u: me.u, act, detail: clip(detail, 400), id });
}
const isEA = (me) => me.role === 'engineer' || me.role === 'admin';
const CHASSIS_RE = /^[A-Za-z0-9][A-Za-z0-9._-]{1,39}$/;
const REF_RE = /^[a-z0-9]{6,40}$/;
const newId = (pfx) => pfx + crypto.randomBytes(6).toString('hex');
const today = () => new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Amsterdam' });
/* weekly overdue summary: sent changes with a planned date in the past that are not released or rejected */
export async function runOverdue(env, trigger = 'weekly') {
  const store = DATA(), miss = mailMissing(env);
  if (miss) return `The mail service is not set up yet (missing: ${miss}).`;
  const parts = await sharedStatus(store), released = await releasedMap(store), appr = await getMeta(store, 'appr'), t = today();
  const over = Object.values(parts).filter((r) => r.dateIso && r.dateIso < t && !released[revId(parseId(r.id).base, parseId(r.id).k + 1)] && !(appr[r.id] && appr[r.id].status === 'rejected'));
  if (!over.length) return trigger === 'weekly' ? '' : 'Nothing is overdue – no e-mail sent.';
  const groups = {}; over.forEach((r) => (groups[r.responsible || 'No responsible person'] = groups[r.responsible || 'No responsible person'] || []).push(r));
  const td = (v, x = '') => `<td style="padding:6px 8px;border-bottom:1px solid #dbe3e8;${x}">${esc(v)}</td>`;
  const html = `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#14232c"><div style="background:#012135;color:#fff;padding:14px 16px;border-bottom:4px solid #d9a441"><b style="font-size:17px">BLIX Parts Review – overdue changes</b><br><span style="color:#a9dff2">${over.length} change${over.length === 1 ? '' : 's'} past the planned date</span></div>` +
    Object.entries(groups).map(([who, rs]) => `<h3 style="margin:18px 0 6px;color:#012135">${esc(who)} (${rs.length})</h3><table style="border-collapse:collapse;font-size:13px;width:100%">${rs.sort((a, b) => a.dateIso.localeCompare(b.dateIso)).map((r) => `<tr>${td(r.vehicle, 'font-weight:600')}${td(r.part, 'font-family:Consolas,monospace;font-weight:600')}${td(r.next || '')}${td(r.note)}${td('Planned ' + r.dateIso.split('-').reverse().join('-'), 'white-space:nowrap;color:#b3470c;font-weight:600')}</tr>`).join('')}</table>`).join('') +
    `<p style="color:#5d6e78;font-size:12px">Sent automatically by the BLIX Parts app.</p></div>`;
  return sendMail(env, { subject: `BLIX parts – ${over.length} overdue change${over.length === 1 ? '' : 's'}`, html });
}

/* ---------- chat ---------- */
const TID_RE = /^t[0-9a-f]{12}$/;
async function myThreads(store, u) {
  return (await listJSON(store, 'thread/')).filter((t) => t && (t.members || []).includes(u) && !(t.hidden && t.hidden[u] && String(t.hidden[u]) >= String((t.last || {}).at || t.created)))
    .map((t) => ({ ...t, unread: !!(t.last && t.last.u !== u && String(t.last.at) > String((t.read || {})[u] || '')) }))
    .sort((a, b) => String((b.last || {}).at || b.created).localeCompare(String((a.last || {}).at || a.created)));
}
async function addMessage(store, t, me, text, photos) {
  const at = new Date().toISOString();
  const m = { at, u: me.u, name: me.name, text: clip(String(text || '').trim(), 4000), photos: (photos || []).filter((x) => /^[0-9a-f]{24}$/.test(x)).slice(0, 10) };
  await store.setJSON(`msg/${t.id}/${at}-${crypto.randomBytes(3).toString('hex')}`, m);
  t.last = { at, u: me.u, name: me.name, text: m.text || (m.photos.length ? '📷 Photo' : '') }; t.count = (t.count || 0) + 1;
  t.read = { ...(t.read || {}), [me.u]: at };
  await store.setJSON(`thread/${t.id}`, t); return m;
}

/* ---------- drafts ---------- */
function cleanDraft(d, id) {
  return {
    id, status: clip(d.status, 6), priority: clip(d.priority, 4), note: clip(d.note, 4000), resp: clip(d.resp, 80),
    date: /^\d{4}-\d{2}-\d{2}$/.test(d.date || '') ? d.date : '',
    photos: (Array.isArray(d.photos) ? d.photos : []).filter((x) => typeof x === 'string' && /^[a-z0-9]{12,40}$/.test(x)).slice(0, 30),
    updatedAt: new Date().toISOString(),
  };
}
async function listJSON(store, prefix) {
  const { blobs } = await store.list({ prefix });
  return (await Promise.all(blobs.map((b) => store.get(b.key, { type: 'json' })))).filter(Boolean);
}
async function sharedStatus(store) {
  const sends = await listJSON(store, 'send/');
  sends.sort((a, b) => String(a.at).localeCompare(String(b.at)));
  const parts = {};
  for (const s of sends) for (const r of s.rows || []) parts[r.id] = { ...r, by: s.by, u: s.u, at: s.at };
  return parts;
}
/* a sent part is locked: move a draft on it to the part's open revision (the newer draft wins) */
async function moveDraft(u, from, to) {
  const store = DATA(), d = await store.get(`drafts/${u}/${from}`, { type: 'json' });
  if (!d || from === to) return;
  const t = await store.get(`drafts/${u}/${to}`, { type: 'json' });
  const [keep, drop] = t && String(t.updatedAt) > String(d.updatedAt) ? [t, d] : [d, t];
  if (drop) { const k = new Set(keep.photos || []); await Promise.all((drop.photos || []).filter((x) => !k.has(x)).map((pid) => PHOTOS().delete(`${u}/${pid}`))); }
  await store.setJSON(`drafts/${u}/${to}`, { ...keep, id: to });
  await store.delete(`drafts/${u}/${from}`);
}
async function deleteDraft(u, id) {
  const store = DATA(), key = `drafts/${u}/${id}`;
  const d = await store.get(key, { type: 'json' });
  if (d) { await Promise.all((d.photos || []).map((pid) => PHOTOS().delete(`${u}/${pid}`))); await store.delete(key); }
}

/* ---------- router ---------- */
export default async (req) => {
  const env = process.env;
  if (!env.SESSION_SECRET || env.SESSION_SECRET.length < 16 || !env.ADMIN_PASSWORD) return err(500, 'The server is not set up yet (SESSION_SECRET / ADMIN_PASSWORD missing).');
  const url = new URL(req.url);
  const path = url.pathname.replace(/^.*\/api\//, '').replace(/\/+$/, '');
  const seg = path.split('/');
  const M = req.method;
  const store = DATA();

  try {
    /* login */
    if (path === 'login' && M === 'POST') {
      const { username, password } = await req.json().catch(() => ({}));
      const u = String(username || '').trim().toLowerCase();
      if (u === 'admin') {
        if (String(password) === env.ADMIN_PASSWORD) return json(200, { token: sign({ u: 'admin', v: 0, exp: Date.now() + 1 * 864e5 }), user: { u: 'admin', name: 'Administrator', role: 'admin' } });
      } else if (USER_RE.test(u)) {
        const rec = await store.get(`users/${u}`, { type: 'json' });
        if (rec && rec.active && checkPw(password, rec)) {
          rec.lastLogin = new Date().toISOString(); await store.setJSON(`users/${u}`, rec);
          return json(200, { token: sign({ u, v: rec.v, exp: Date.now() + SESSION_DAYS * 864e5 }), user: { u, name: rec.name, role: rec.role } });
        }
      }
      await new Promise((r) => setTimeout(r, 600)); // slow down guessing
      return err(401, 'Username or password is not correct.');
    }

    const me = await auth(req);
    if (!me) return err(401, 'Please sign in again.');

    /* everything the app needs after login */
    if (path === 'state' && M === 'GET') {
      const parts = await sharedStatus(store);
      let drafts = me.u === 'admin' ? [] : await listJSON(store, `drafts/${me.u}/`);
      const stale = drafts.filter((d) => parts[d.id]);
      if (stale.length) { for (const d of stale) await moveDraft(me.u, d.id, openRev(parts, d.id)); drafts = await listJSON(store, `drafts/${me.u}/`); }
      const team = (await store.get('config/team', { type: 'json' })) || { names: [] };
      const maint = (await listJSON(store, 'maint/')).sort((a, b) => String(b.at).localeCompare(String(a.at)))[0] || null;
      const meta = Object.fromEntries(await Promise.all(META.map(async (k) => [k, await getMeta(store, k)])));
      const used = parseInt(url.searchParams.get('used') || '', 10);
      if (used > 0 && me.u) await store.setJSON(`usage/${today()}/${me.u}`, { n: Math.min(used, 1e6), at: new Date().toISOString() });
      const ncrOpen = await getMeta(store, 'ncrindex');
      return json(200, { user: me, drafts, parts, team: team.names, cleared: await clearedMap(store), maint, released: await releasedMap(store), unread: me.u === 'admin' ? 0 : (await myThreads(store, me.u)).filter((t) => t.unread).length,
        prefs: me.u === 'admin' ? { clr: {}, hide: {} } : ((await store.get(`prefs/${me.u}`, { type: 'json' })) || { clr: {}, hide: {} }), meta, ncrMine: Object.values(ncrOpen).filter((n) => n.status !== 'closed' && n.assignee === me.u).length, ncrOpen: Object.values(ncrOpen).filter((n) => n.status !== 'closed').length });
    }

    /* drafts */
    if ((seg[0] === 'drafts' || path === 'send') && me.role === 'engineer') return err(403, 'Engineering accounts do not send reports.');
    if (seg[0] === 'drafts' && seg.length === 2 && PART_RE.test(seg[1]) && me.u !== 'admin') {
      if (M === 'PUT') {
        const parts = await sharedStatus(store), q = parseId(seg[1]);
        if (q.k > 0 && !parts[revId(q.base, q.k - 1)]) return err(400, 'This revision does not exist yet.');
        const id = openRev(parts, seg[1]); // part already sent → the draft goes to its new revision
        const d = cleanDraft(await req.json(), id);
        if (id !== seg[1]) await moveDraft(me.u, seg[1], id);
        await store.setJSON(`drafts/${me.u}/${id}`, d); return json(200, { draft: d, moved: id !== seg[1] });
      }
      if (M === 'DELETE') { await deleteDraft(me.u, seg[1]); return json(200, { ok: true }); }
    }

    /* chat: people, threads, messages, chat photos */
    if (seg[0] === 'people' && M === 'GET') {
      return json(200, { people: (await listJSON(store, 'users/')).filter((x) => x.active && x.u !== me.u).map((x) => ({ u: x.u, name: x.name, role: x.role })).sort((a, b) => a.name.localeCompare(b.name)) });
    }
    if (seg[0] === 'threads' && me.u !== 'admin') {
      if (seg.length === 1 && M === 'GET') return json(200, { threads: await myThreads(store, me.u) });
      if (seg.length === 2 && TID_RE.test(seg[1]) && M === 'DELETE') {   /* remove a conversation from my list (comes back if a new message arrives) */
        const key = `thread/${seg[1]}`, t = await store.get(key, { type: 'json' }); if (!t || !(t.members || []).includes(me.u)) return err(404, 'Conversation not found.');
        t.hidden = { ...(t.hidden || {}), [me.u]: new Date().toISOString() }; t.read = { ...(t.read || {}), [me.u]: (t.last || {}).at || t.read?.[me.u] };
        if (t.members.every((m) => t.hidden[m])) { const { blobs } = await store.list({ prefix: `msg/${t.id}/` }); await Promise.all(blobs.map((x) => store.delete(x.key))); const { blobs: ph } = await PHOTOS().list({ prefix: `chat/${t.id}/` }); await Promise.all(ph.map((x) => PHOTOS().delete(x.key))); await store.delete(key); }
        else await store.setJSON(key, t);
        return json(200, { ok: true });
      }
      if (seg.length === 1 && M === 'POST') {
        const b = await req.json().catch(() => ({}));
        const to = [...new Set((Array.isArray(b.to) ? b.to : []).map(String).filter((x) => USER_RE.test(x) && x !== me.u))].slice(0, 10);
        if (!to.length) return err(400, 'Choose who the message is for.');
        const recs = await Promise.all(to.map((x) => store.get(`users/${x}`, { type: 'json' })));
        if (recs.some((r) => !r || !r.active)) return err(400, 'One of the chosen people has no active account.');
        const part = b.part && PART_RE.test(b.part.id || '') ? { id: b.part.id, name: clip(b.part.name, 80) } : null;
        const members = [me.u, ...to].sort();
        let t = (await myThreads(store, me.u)).find((x) => x.members.slice().sort().join() === members.join() && ((x.part && x.part.id) || '') === ((part && part.id) || ''));
        if (!t) {
          t = { id: 't' + crypto.randomBytes(6).toString('hex'), members, names: Object.fromEntries([[me.u, me.name], ...recs.map((r) => [r.u, r.name])]), part, created: new Date().toISOString(), by: me.u, read: {}, count: 0 };
          delete t.unread; await store.setJSON(`thread/${t.id}`, t);
        } else delete t.unread;
        if (String(b.text || '').trim()) await addMessage(store, t, me, b.text, []);
        return json(200, { thread: t });
      }
      if (seg.length >= 2 && TID_RE.test(seg[1])) {
        const t = await store.get(`thread/${seg[1]}`, { type: 'json' });
        if (!t || !(t.members || []).includes(me.u)) return err(404, 'Conversation not found.');
        if (seg.length === 2 && M === 'GET') {
          const since = url.searchParams.get('since') || '';
          const { blobs } = await store.list({ prefix: `msg/${t.id}/` });
          const keys = blobs.map((x) => x.key).filter((k) => k.slice(k.lastIndexOf('/') + 1) > since).sort();
          const msgs = (await Promise.all(keys.map((k) => store.get(k, { type: 'json' })))).filter(Boolean);
          if (t.last && String((t.read || {})[me.u] || '') < String(t.last.at)) { t.read = { ...(t.read || {}), [me.u]: t.last.at }; await store.setJSON(`thread/${t.id}`, t); }
          return json(200, { thread: t, messages: msgs });
        }
        if (seg.length === 3 && seg[2] === 'messages' && M === 'POST') {
          const b = await req.json().catch(() => ({}));
          if (!String(b.text || '').trim() && !(b.photos || []).length) return err(400, 'Write a message or add a photo.');
          return json(200, { message: await addMessage(store, t, me, b.text, b.photos) });
        }
        if (seg.length === 3 && seg[2] === 'photos' && M === 'POST') {
          const buf = await req.arrayBuffer();
          if (!buf.byteLength || buf.byteLength > MAX_PHOTO) return err(413, 'Photo is too large.');
          const id = crypto.randomBytes(12).toString('hex');
          await PHOTOS().set(`chat/${t.id}/${id}`, buf, { metadata: { type: 'image/jpeg', by: me.u } });
          return json(200, { id });
        }
        if (seg.length === 4 && seg[2] === 'photos' && /^[0-9a-f]{24}$/.test(seg[3]) && M === 'GET') {
          const buf = await PHOTOS().get(`chat/${t.id}/${seg[3]}`, { type: 'arrayBuffer' });
          if (!buf) return err(404, 'Photo not found.');
          return new Response(buf, { status: 200, headers: { 'Content-Type': 'image/jpeg', 'Cache-Control': 'private, max-age=31536000' } });
        }
      }
    }

    /* photos (stored per account) */
    if (path === 'photos' && M === 'POST' && me.u !== 'admin') {
      const buf = await req.arrayBuffer();
      if (!buf.byteLength || buf.byteLength > MAX_PHOTO) return err(413, 'Photo is too large.');
      const id = crypto.randomBytes(12).toString('hex');
      await PHOTOS().set(`${me.u}/${id}`, buf, { metadata: { type: 'image/jpeg', at: new Date().toISOString() } });
      return json(200, { id });
    }
    if (seg[0] === 'photos' && seg.length === 2 && M === 'GET') {
      const buf = await PHOTOS().get(`${me.u}/${seg[1]}`, { type: 'arrayBuffer' });
      if (!buf) return err(404, 'Photo not found.');
      return new Response(buf, { status: 200, headers: { 'Content-Type': 'image/jpeg', 'Cache-Control': 'private, max-age=31536000' } });
    }

    /* send report by e-mail */
    if (path === 'send' && M === 'POST' && me.u !== 'admin') {
      const miss = mailMissing(env);
      if (miss) return err(500, `The mail service is not set up yet (missing: ${miss}).`);
      const p = await req.json().catch(() => null);
      if (!p || typeof p.file !== 'string' || !/^[\w\-. ()]+\.xlsx$/.test(p.filename || '')) return err(400, 'The report is incomplete.');
      if (p.file.length * 0.75 > MAX_FILE) return err(413, 'The report is too large to e-mail. Send fewer parts at once.');
      const rows = (p.rows || []).filter((r) => r && PART_RE.test(r.id || ''));
      const parts = await sharedStatus(store);
      const taken = rows.filter((r) => parts[r.id]);
      if (taken.length) {
        for (const r of taken) await moveDraft(me.u, r.id, openRev(parts, r.id));
        return err(409, `Already sent by someone else: ${taken.map((r) => `${r.part} (${parts[r.id].by})`).join(', ')}. Your change was moved to the new revision – check it and send again.`);
      }
      if (rows.some((r) => { const q = parseId(r.id); return q.k > 0 && !parts[revId(q.base, q.k - 1)]; })) return err(400, 'The report contains a revision that does not exist yet.');
      const vehicles = [...new Set(rows.map((r) => r.vehicle))].join(', ');
      const subject = `BLIX parts report – ${me.name} – ${vehicles || 'parts'} (${rows.length} part${rows.length === 1 ? '' : 's'})${p.batches > 1 ? ` [${p.batch}/${p.batches}]` : ''}`;
      const mailErr = await sendMail(env, { subject, html: mailHtml({ ...p, rows }, me.name), file: { name: p.filename, b64: p.file } });
      if (mailErr) return err(502, mailErr);
      const at = new Date().toISOString();
      const rec = rows.map((x) => ({ id: x.id, statusKey: clip(x.statusKey, 6), status: clip(x.status, 40), priorityKey: clip(x.priorityKey, 4), priority: clip(x.priority, 20),
        note: clip(x.note, 4000), responsible: clip(x.responsible, 80), dateIso: /^\d{4}-\d{2}-\d{2}$/.test(x.dateIso || '') ? x.dateIso : '', date: clip(x.date, 30),
        photos: Number(x.photos) || 0, vehicle: clip(x.vehicle, 10), part: clip(x.part, 80), next: clip(x.next, 80) }));
      await store.setJSON(`send/${at}-${crypto.randomBytes(3).toString('hex')}`, { by: me.name, u: me.u, at, rows: rec });
      for (const x of rec) await audit(store, x.id, me, `Sent ${x.part}`, `${x.status}${x.note ? ' – ' + x.note : ''}${x.next ? ' → new revision ' + x.next : ''}`);
      await Promise.all(rows.map((x) => deleteDraft(me.u, x.id))); // photos went out by e-mail; clear the draft
      return json(200, { ok: true, at });
    }

    /* Engineering: mark a revision as released in CAD (or undo) */
    if (seg[0] === 'release' && seg.length === 2 && PART_RE.test(seg[1])) {
      if (me.role !== 'engineer') return err(403, 'Only an Engineering account can release parts.');
      const id = seg[1], q = parseId(id), key = `release/${id}`;
      if (M === 'POST') {
        const parts = await sharedStatus(store);
        if (q.k < 1 || !parts[revId(q.base, q.k - 1)]) return err(400, 'This revision does not exist (yet).');
        const b = await req.json().catch(() => ({}));
        const rec = { id, name: clip(b.name, 80), note: clip(String(b.note || '').trim(), 300), at: new Date().toISOString(), by: me.name, u: me.u };
        await store.setJSON(key, rec); await audit(store, id, me, `Released ${rec.name} in CAD`, rec.note); return json(200, { release: rec });
      }
      if (M === 'DELETE') { await store.delete(key); for (const k of ['fit', 'done']) await updMeta(store, k, (m) => { delete m[id]; }); await audit(store, id, me, 'Release undone'); return json(200, { ok: true }); }
    }

    /* clear changes of one part (any account) or a whole section (Admin): unlocks for everyone, e-mails Engineering first */
    if (path === 'clear' && M === 'POST') {
      const b = await req.json().catch(() => ({}));
      const reason = clip(String(b.reason || '').trim(), 500);
      if (reason.length < 3) return err(400, 'Write a short reason – it is sent to Engineering.');
      const want = new Map();
      for (const x of Array.isArray(b.items) ? b.items.slice(0, 600) : []) {
        const id = String(x && x.id || ''); if (!/^p[0-9a-f]{10}$/.test(id) || want.has(id)) continue;
        want.set(id, { id, from: Math.max(0, Math.min(60, parseInt(x.from, 10) || 0)), name: clip(x.name, 80), fromName: clip(x.fromName, 80), drafts: 0 });
      }
      if (!want.size) return err(400, 'No parts chosen.');
      if (want.size > 1 && me.role !== 'admin') return err(403, 'Only an Admin account can clear a whole section.');
      const miss = mailMissing(env);
      if (miss) return err(500, `Clearing needs the e-mail service so Engineering is notified (missing: ${miss}).`);
      const inScope = (id) => { const q = parseId(id); const w = q && want.get(q.base); return !!w && q.k >= w.from; };
      const parts = await sharedStatus(store);
      const rows = Object.values(parts).filter((r) => inScope(r.id)).map((r) => ({ base: parseId(r.id).base, id: r.id, part: r.part, status: r.status, note: r.note, by: r.by, at: r.at }));
      const { blobs: dr } = await store.list({ prefix: 'drafts/' });
      const drafts = dr.map((x) => x.key.split('/')).filter((k) => k.length === 3 && inScope(k[2]));
      for (const k of drafts) want.get(parseId(k[2]).base).drafts++;
      if (!rows.length && !drafts.length) return err(400, 'There is nothing to clear – no sent reports or drafts.');
      const ev = { at: new Date().toISOString(), by: me.name, u: me.u, reason, parts: [...want.values()], rows };
      ev.parts = ev.parts.filter((p) => p.drafts || rows.some((r) => r.base === p.id));
      const mailErr = await sendMail(env, { subject: `BLIX parts CLEARED – ${ev.parts.map((p) => p.name).slice(0, 6).join(', ')}${ev.parts.length > 6 ? ` +${ev.parts.length - 6}` : ''} – by ${me.name}`, html: clearMailHtml(ev) });
      if (mailErr) return err(502, `${mailErr} Nothing was cleared.`);
      const { blobs } = await store.list({ prefix: 'send/' });
      for (const x of blobs) { const rec = await store.get(x.key, { type: 'json' }); if (!rec || !(rec.rows || []).some((r) => inScope(r.id))) continue;
        rec.rows = rec.rows.filter((r) => !inScope(r.id)); if (rec.rows.length) await store.setJSON(x.key, rec); else await store.delete(x.key); }
      await Promise.all(drafts.map((k) => deleteDraft(k[1], k[2])));
      await updMeta(store, 'appr', (m) => { for (const k of Object.keys(m)) if (inScope(k)) delete m[k]; });
      for (const mk of ['fit', 'done']) await updMeta(store, mk, (m) => { for (const k of Object.keys(m)) { const q = parseId(k), w = q && want.get(q.base); if (w && q.k > w.from) delete m[k]; } });
      for (const p of ev.parts) await audit(store, p.id, me, 'Changes cleared', `${p.from ? 'From ' + p.fromName + ' onwards' : 'Everything'} – ${reason}`);
      const { blobs: rl } = await store.list({ prefix: 'release/' }); // revisions created by removed reports no longer exist
      for (const x of rl) { const q = parseId(x.key.slice(8)); const w = q && want.get(q.base); if (w && q.k > w.from) await store.delete(x.key); }
      for (const p of ev.parts) {
        const key = `clear/${p.id}`, old = (await store.get(key, { type: 'json' })) || { base: p.id, history: [] };
        const last = { at: ev.at, by: ev.by, name: p.name, reason, from: p.from, fromName: p.fromName, reports: rows.filter((r) => r.base === p.id).length, drafts: p.drafts };
        await store.setJSON(key, { base: p.id, last, history: [...(old.history || []), last].slice(-30) });
      }
      return json(200, { ok: true, at: ev.at, parts: ev.parts.length });
    }

    /* ================= v12: parts, stock, drawings, approvals, fitted, workshop ================= */
    /* add / edit parts (Engineering + Admin) */
    if (seg[0] === 'partx' && M === 'PUT' && seg.length === 2 && PART_RE.test(seg[1])) {
      if (!isEA(me)) return err(403, 'Only Engineering or Admin can edit parts.');
      const b = await req.json().catch(() => ({})), id = seg[1];
      const rec = await updMeta(store, 'partx', (m) => { const r = { ...(m[id] || {}) };
        for (const k of ['n', 'v', 'c', 't']) if (b[k] !== undefined) r[k] = clip(String(b[k]).trim(), k === 'c' ? 60 : 40);
        if (b.img !== undefined) r.img = typeof b.img === 'string' && b.img.startsWith('data:image/') && b.img.length < 120000 ? b.img : r.img;
        if (b.hidden !== undefined) r.hidden = !!b.hidden;
        if (b.supplier !== undefined) r.supplier = { name: clip(b.supplier.name, 80), email: clip(b.supplier.email, 120) };
        if (b.isNew) r.custom = true; r.by = me.name; r.at = new Date().toISOString(); m[id] = r; return r; });
      if (b.vehName && b.v) await updMeta(store, 'veh', (m) => { m[clip(b.v, 20)] = clip(b.vehName, 80); });
      await audit(store, id, me, b.isNew ? `Part added: ${rec.n}` : (b.hidden !== undefined ? (b.hidden ? 'Part hidden' : 'Part shown again') : 'Part details edited'), [rec.n, rec.v, rec.c].filter(Boolean).join(' · '));
      return json(200, { part: rec });
    }
    if (path === 'partx/import' && M === 'POST') {
      if (!isEA(me)) return err(403, 'Only Engineering or Admin can import parts.');
      const b = await req.json().catch(() => ({}));
      const rows = (Array.isArray(b.rows) ? b.rows : []).slice(0, 1000).filter((r) => r && String(r.n || '').trim() && String(r.v || '').trim());
      if (!rows.length) return err(400, 'No rows with at least a part number and vehicle.');
      const at = new Date().toISOString(); const made = [];
      await updMeta(store, 'partx', (m) => { for (const r of rows) { const id = 'p' + crypto.randomBytes(5).toString('hex'); m[id] = { n: clip(String(r.n).trim(), 40), v: clip(String(r.v).trim().toUpperCase(), 20), c: clip(String(r.c || 'Other').trim(), 60), t: String(r.t || '').toUpperCase().startsWith('A') ? 'A' : 'P', custom: true, by: me.name, at }; made.push(id); } });
      if (b.vehNames) await updMeta(store, 'veh', (m) => { for (const [k, v] of Object.entries(b.vehNames)) if (v) m[clip(k, 20)] = clip(v, 80); });
      await audit(store, 'general', me, `Imported ${made.length} parts`);
      return json(200, { added: made.length });
    }
    /* stock (any account) */
    if (seg[0] === 'stock' && M === 'POST' && seg.length === 2 && PART_RE.test(seg[1])) {
      const b = await req.json().catch(() => ({})), id = seg[1];
      const rec = await updMeta(store, 'stock', (m) => { const r = { qty: 0, min: 0, loc: '', ...(m[id] || {}) };
        if (Number.isFinite(+b.delta) && b.delta !== null && b.delta !== '') r.qty = Math.max(0, (+r.qty || 0) + Math.round(+b.delta));
        if (b.qty !== undefined && b.qty !== '' && Number.isFinite(+b.qty)) r.qty = Math.max(0, Math.round(+b.qty));
        if (b.min !== undefined && b.min !== '' && Number.isFinite(+b.min)) r.min = Math.max(0, Math.round(+b.min));
        if (b.loc !== undefined) r.loc = clip(String(b.loc).trim(), 40);
        r.by = me.name; r.at = new Date().toISOString(); m[id] = r; return r; });
      await audit(store, id, me, 'Stock updated', `${b.delta ? (b.delta > 0 ? '+' : '') + b.delta + ' → ' : ''}${rec.qty} in stock${rec.min ? ', min ' + rec.min : ''}${rec.loc ? ', location ' + rec.loc : ''}`);
      return json(200, { stock: rec });
    }
    /* drawings (PDF) per part or revision */
    if (seg[0] === 'dwg' && seg.length === 2 && PART_RE.test(seg[1])) {
      const id = seg[1];
      if (M === 'GET') { const buf = await PHOTOS().get(`dwg/${id}`, { type: 'arrayBuffer' }); if (!buf) return err(404, 'No drawing.'); return new Response(buf, { status: 200, headers: { 'Content-Type': 'application/pdf', 'Cache-Control': 'private, max-age=60' } }); }
      if (!isEA(me)) return err(403, 'Only Engineering or Admin can change drawings.');
      if (M === 'POST') {
        const buf = await req.arrayBuffer(); if (!buf.byteLength || buf.byteLength > 5_500_000) return err(413, 'The PDF is too large (max 5 MB).');
        if (Buffer.from(buf.slice(0, 5)).toString() !== '%PDF-') return err(400, 'This is not a PDF file.');
        await PHOTOS().set(`dwg/${id}`, buf);
        const name = clip(url.searchParams.get('name') || 'drawing.pdf', 100);
        await updMeta(store, 'dwg', (m) => { m[id] = { name, size: buf.byteLength, at: new Date().toISOString(), by: me.name }; });
        await audit(store, id, me, 'Drawing uploaded', name); return json(200, { ok: true });
      }
      if (M === 'DELETE') { await PHOTOS().delete(`dwg/${id}`); await updMeta(store, 'dwg', (m) => { delete m[id]; }); await audit(store, id, me, 'Drawing removed'); return json(200, { ok: true }); }
    }
    /* send drawing / change to the supplier */
    if (seg[0] === 'supplier' && seg.length === 2 && PART_RE.test(seg[1]) && M === 'POST') {
      if (!isEA(me)) return err(403, 'Only Engineering or Admin can e-mail suppliers.');
      const miss = mailMissing(env); if (miss) return err(500, `The mail service is not set up yet (missing: ${miss}).`);
      const b = await req.json().catch(() => ({})), id = seg[1];
      const email = String(b.email || '').trim(); if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return err(400, 'Enter a valid supplier e-mail address.');
      const dwgId = PART_RE.test(b.dwg || '') ? b.dwg : id, buf = await PHOTOS().get(`dwg/${dwgId}`, { type: 'arrayBuffer' }), dm = (await getMeta(store, 'dwg'))[dwgId];
      const html = `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#14232c"><p>${esc(b.message || '').replace(/\n/g, '<br>')}</p><p><b>Part:</b> ${esc(b.part || '')}</p>${buf ? '<p>The drawing is attached.</p>' : ''}<p>Kind regards,<br>${esc(me.name)}<br>BLIX Automotive – Engineering</p></div>`;
      const e = await sendMail(env, { subject: clip(b.subject || `Drawing ${b.part || ''}`, 150), html, to: [email, ...env.MAIL_TO.split(',').map((a) => a.trim()).filter(Boolean)],
        files: buf ? [{ name: dm ? dm.name : `${b.part || 'drawing'}.pdf`, b64: Buffer.from(buf).toString('base64'), type: 'application/pdf' }] : [] });
      if (e) return err(502, e);
      await audit(store, id, me, 'Sent to supplier', `${email}${buf ? ' (with drawing)' : ''}`); return json(200, { ok: true });
    }
    /* approval of a sent change (Engineering) */
    if (seg[0] === 'appr' && seg.length === 2 && PART_RE.test(seg[1])) {
      if (me.role !== 'engineer') return err(403, 'Only an Engineering account can approve or reject changes.');
      const id = seg[1];
      if (M === 'POST') {
        const parts = await sharedStatus(store); if (!parts[id]) return err(400, 'This part has not been sent.');
        const b = await req.json().catch(() => ({}));
        if (!['approved', 'rejected'].includes(b.status)) return err(400, 'Choose approve or reject.');
        if (b.status === 'rejected' && String(b.reason || '').trim().length < 3) return err(400, 'Write why the change is rejected.');
        const rec = { status: b.status, reason: clip(String(b.reason || '').trim(), 500), cost: clip(b.cost, 30), stock: ['use', 'scrap', 'rework', ''].includes(b.stock) ? b.stock : '', effect: clip(b.effect, 60),
          deadline: /^\d{4}-\d{2}-\d{2}$/.test(b.deadline || '') ? b.deadline : '', at: new Date().toISOString(), by: me.name, u: me.u };
        await updMeta(store, 'appr', (m) => { m[id] = rec; });
        await audit(store, id, me, b.status === 'approved' ? 'Change approved' : 'Change rejected', [rec.reason, rec.cost && 'cost ' + rec.cost, rec.stock && 'old stock: ' + rec.stock, rec.effect && 'from ' + rec.effect, rec.deadline && 'deadline ' + rec.deadline].filter(Boolean).join(' · '));
        return json(200, { appr: rec });
      }
      if (M === 'DELETE') { await updMeta(store, 'appr', (m) => { delete m[id]; }); await audit(store, id, me, 'Approval undone'); return json(200, { ok: true }); }
    }
    /* fitted on vehicle (after release) */
    if (seg[0] === 'fit' && seg.length === 2 && PART_RE.test(seg[1]) && M === 'POST') {   /* fit check by the workshop after release */
      const id = seg[1], b = await req.json().catch(() => ({}));
      if (b.remove) { if (!isEA(me)) return err(403, 'Only Engineering can ask for a new fit check.'); await updMeta(store, 'fit', (m) => { delete m[id]; }); await updMeta(store, 'done', (m) => { delete m[id]; }); await audit(store, id, me, 'New fit check requested', clip(b.note, 300)); return json(200, { ok: true }); }
      const released = await releasedMap(store); if (!released[id]) return err(400, 'This revision is not released in CAD yet.');
      if (!['ok', 'work'].includes(b.result)) return err(400, 'Choose “Fits OK” or “Extra work needed”.');
      if (b.result === 'work' && String(b.note || '').trim().length < 3) return err(400, 'Describe the extra work that is needed.');
      const chassis = String(b.chassis || '').trim();
      const rec = { result: b.result, note: clip(String(b.note || '').trim(), 1000), chassis: CHASSIS_RE.test(chassis) ? chassis : '', at: new Date().toISOString(), by: me.name, u: me.u };
      await updMeta(store, 'fit', (m) => { m[id] = rec; });
      await audit(store, id, me, rec.result === 'ok' ? 'Fit check: fits OK' : 'Fit check: extra work needed', [rec.note, rec.chassis && 'chassis ' + rec.chassis].filter(Boolean).join(' · '));
      return json(200, { fit: rec });
    }
    /* final approval by Engineering: drawing, CAD and fitment OK → change completed */
    if (seg[0] === 'done' && seg.length === 2 && PART_RE.test(seg[1])) {
      if (me.role !== 'engineer') return err(403, 'Only an Engineering account can give the final approval.');
      const id = seg[1];
      if (M === 'POST') {
        const fit = (await getMeta(store, 'fit'))[id]; if (!fit || Array.isArray(fit)) return err(400, 'The workshop has not done the fit check yet.');
        const b = await req.json().catch(() => ({}));
        const rec = { note: clip(String(b.note || '').trim(), 500), name: clip(b.name, 80), at: new Date().toISOString(), by: me.name, u: me.u };
        await updMeta(store, 'done', (m) => { m[id] = rec; }); await audit(store, id, me, `Final approval – ${rec.name || 'revision'} completed`, rec.note); return json(200, { done: rec });
      }
      if (M === 'DELETE') { await updMeta(store, 'done', (m) => { delete m[id]; }); await audit(store, id, me, 'Final approval undone'); return json(200, { ok: true }); }
    }
    /* personal preferences: cleared / hidden lists (per account, on every phone) */
    if (path === 'prefs' && M === 'PUT' && me.u !== 'admin') {
      const b = await req.json().catch(() => ({})), key = `prefs/${me.u}`, pr = (await store.get(key, { type: 'json' })) || { clr: {}, hide: {} };
      for (const [k, v] of Object.entries(b.clr || {})) if (/^[a-z]{2,20}$/.test(k)) pr.clr[k] = v ? new Date().toISOString() : undefined;
      for (const [k, obj] of Object.entries(b.hide || {})) { if (!/^[a-z]{2,20}$/.test(k)) continue; pr.hide[k] = pr.hide[k] || {}; for (const [id, sig] of Object.entries(obj || {})) { if (!PART_RE.test(id)) continue; if (sig === null) delete pr.hide[k][id]; else pr.hide[k][id] = clip(sig, 40); } const e = Object.entries(pr.hide[k]); if (e.length > 800) pr.hide[k] = Object.fromEntries(e.slice(-800)); }
      if (b.reset) pr.hide = {}, pr.clr = {};
      await store.setJSON(key, pr); return json(200, { prefs: pr });
    }
    /* per-part history */
    if (seg[0] === 'audit' && seg.length === 2 && /^p[0-9a-f]{10}$|^general$/.test(seg[1]) && M === 'GET') {
      const list = (await listJSON(store, `audit/${seg[1]}/`)).sort((a, b) => String(b.at).localeCompare(String(a.at))).slice(0, 150);
      return json(200, { audit: list });
    }
    /* photos for problem reports, inspections and instructions */
    if (seg[0] === 'up' && ['ncr', 'insp', 'instr'].includes(seg[1]) && REF_RE.test(seg[2] || '')) {
      if (seg.length === 3 && M === 'POST') {
        const buf = await req.arrayBuffer(); if (!buf.byteLength || buf.byteLength > MAX_PHOTO) return err(413, 'Photo is too large.');
        const id = crypto.randomBytes(12).toString('hex'); await PHOTOS().set(`${seg[1]}/${seg[2]}/${id}`, buf); return json(200, { id });
      }
      if (seg.length === 4 && /^[0-9a-f]{24}$/.test(seg[3]) && M === 'GET') {
        const buf = await PHOTOS().get(`${seg[1]}/${seg[2]}/${seg[3]}`, { type: 'arrayBuffer' }); if (!buf) return err(404, 'Photo not found.');
        return new Response(buf, { status: 200, headers: { 'Content-Type': 'image/jpeg', 'Cache-Control': 'private, max-age=31536000' } });
      }
    }
    /* problem reports (non-conformance) */
    if (seg[0] === 'ncr') {
      const cleanP = (a) => (Array.isArray(a) ? a : []).filter((x) => /^[0-9a-f]{24}$/.test(x)).slice(0, 12);
      const idx = async (n) => updMeta(store, 'ncrindex', (m) => { m[n.id] = { id: n.id, status: n.status, assignee: (n.assignee || {}).u || '', title: n.title, sev: n.sev, at: n.at, updated: n.updated, chassis: n.chassis, part: n.part, cat: n.cat, by: n.by, closedAt: n.closedAt || '' }; });
      if (seg.length === 1 && M === 'GET') return json(200, { ncr: Object.values(await getMeta(store, 'ncrindex')).sort((a, b) => String(b.updated || b.at).localeCompare(String(a.updated || a.at))) });
      if (seg.length === 1 && M === 'POST') {
        const b = await req.json().catch(() => ({})); if (String(b.title || '').trim().length < 3) return err(400, 'Give the problem a short title.');
        const at = new Date().toISOString(), id = newId('n');
        const asg = USER_RE.test(b.assignee || '') ? await store.get(`users/${b.assignee}`, { type: 'json' }) : null;
        const n = { id, title: clip(String(b.title).trim(), 120), cat: clip(b.cat, 40), sev: ['low', 'medium', 'high'].includes(b.sev) ? b.sev : 'medium', chassis: CHASSIS_RE.test(b.chassis || '') ? b.chassis : '',
          part: b.part && PART_RE.test(b.part.id || '') ? { id: b.part.id, name: clip(b.part.name, 80) } : null, desc: clip(String(b.desc || '').trim(), 4000), photos: cleanP(b.photos), ref: REF_RE.test(b.ref || '') ? b.ref : id,
          status: 'open', assignee: asg ? { u: asg.u, name: asg.name } : null, cause: '', action: '', comments: [], by: me.name, u: me.u, at, updated: at };
        await store.setJSON(`ncr/${id}`, n); await idx(n); if (n.part) await audit(store, n.part.id, me, `Problem reported: ${n.title}`, n.desc);
        return json(200, { ncr: n });
      }
      if (seg.length === 2 && /^n[0-9a-f]{12}$/.test(seg[1])) {
        const key = `ncr/${seg[1]}`, n = await store.get(key, { type: 'json' }); if (!n) return err(404, 'Problem report not found.');
        if (M === 'GET') return json(200, { ncr: n });
        if (M === 'PUT') {
          const b = await req.json().catch(() => ({})), at = new Date().toISOString(), changes = [];
          if (b.status && ['open', 'progress', 'closed'].includes(b.status) && b.status !== n.status) { if (b.status === 'closed' && !String(b.action ?? n.action).trim()) return err(400, 'Write the action taken before closing.'); changes.push(`Status → ${b.status}`); n.status = b.status; if (b.status === 'closed') { n.closedAt = at; n.closedBy = me.name; } else { n.closedAt = ''; } }
          if (b.cause !== undefined && b.cause !== n.cause) { n.cause = clip(String(b.cause).trim(), 2000); changes.push('Cause updated'); }
          if (b.action !== undefined && b.action !== n.action) { n.action = clip(String(b.action).trim(), 2000); changes.push('Action updated'); }
          if (b.assignee !== undefined) { const a = USER_RE.test(b.assignee || '') ? await store.get(`users/${b.assignee}`, { type: 'json' }) : null; n.assignee = a ? { u: a.u, name: a.name } : null; changes.push(`Assigned to ${a ? a.name : 'nobody'}`); }
          if (String(b.comment || '').trim() || cleanP(b.photos).length) n.comments = [...(n.comments || []), { at, by: me.name, u: me.u, text: clip(String(b.comment || '').trim(), 2000), photos: cleanP(b.photos) }].slice(-200);
          if (changes.length) n.comments = [...(n.comments || []), { at, by: me.name, u: me.u, sys: changes.join(' · ') }];
          n.updated = at; await store.setJSON(key, n); await idx(n);
          if (n.part && changes.length) await audit(store, n.part.id, me, `Problem “${n.title}”`, changes.join(' · '));
          return json(200, { ncr: n });
        }
      }
    }
    /* checklist templates */
    if (seg[0] === 'chk') {
      if (seg.length === 1 && M === 'GET') return json(200, { chk: Object.values(await getMeta(store, 'chk')).sort((a, b) => a.name.localeCompare(b.name)) });
      if (seg.length === 2 && /^c[0-9a-f]{12}$|^new$/.test(seg[1]) && ['PUT', 'DELETE'].includes(M)) {
        if (!isEA(me)) return err(403, 'Only Engineering or Admin can edit checklists.');
        if (M === 'DELETE') { await updMeta(store, 'chk', (m) => { delete m[seg[1]]; }); return json(200, { ok: true }); }
        const b = await req.json().catch(() => ({})), id = seg[1] === 'new' ? newId('c') : seg[1];
        const items = (Array.isArray(b.items) ? b.items : []).map((x) => ({ t: clip(String(x.t || '').trim(), 200), photo: !!x.photo, val: !!x.val })).filter((x) => x.t).slice(0, 150);
        if (String(b.name || '').trim().length < 2 || !items.length) return err(400, 'Give the checklist a name and at least one item.');
        const rec = { id, name: clip(String(b.name).trim(), 80), veh: clip(b.veh || '', 20), items, by: me.name, at: new Date().toISOString() };
        await updMeta(store, 'chk', (m) => { m[id] = rec; }); return json(200, { chk: rec });
      }
    }
    /* inspections (filled-in checklists) */
    if (seg[0] === 'insp') {
      const idx = async (x) => updMeta(store, 'inspindex', (m) => { m[x.id] = { id: x.id, name: x.name, chassis: x.chassis, veh: x.veh, status: x.status, by: x.by, at: x.at, doneAt: x.doneAt || '', ok: x.items.filter((i) => i.ok === true).length, nok: x.items.filter((i) => i.ok === false).length, total: x.items.length }; });
      if (seg.length === 1 && M === 'GET') return json(200, { insp: Object.values(await getMeta(store, 'inspindex')).sort((a, b) => String(b.at).localeCompare(String(a.at))) });
      if (seg.length === 1 && M === 'POST') {
        const b = await req.json().catch(() => ({})), tpl = (await getMeta(store, 'chk'))[b.tpl]; if (!tpl) return err(400, 'Choose a checklist.');
        const chassis = String(b.chassis || '').trim(); if (!CHASSIS_RE.test(chassis)) return err(400, 'Enter the chassis number.');
        const x = { id: newId('i'), tpl: tpl.id, name: tpl.name, veh: tpl.veh, chassis, items: tpl.items.map((i) => ({ ...i, ok: null, value: '', note: '', photos: [] })), status: 'open', by: me.name, u: me.u, at: new Date().toISOString() };
        await store.setJSON(`insp/${x.id}`, x); await idx(x); return json(200, { insp: x });
      }
      if (seg.length === 2 && /^i[0-9a-f]{12}$/.test(seg[1])) {
        const key = `insp/${seg[1]}`, x = await store.get(key, { type: 'json' }); if (!x) return err(404, 'Inspection not found.');
        if (M === 'GET') return json(200, { insp: x });
        if (M === 'PUT') {
          if (x.status === 'done' && !isEA(me)) return err(400, 'This inspection is signed off and locked.');
          const b = await req.json().catch(() => ({}));
          if (Array.isArray(b.items)) b.items.slice(0, x.items.length).forEach((i, k) => { const t = x.items[k]; t.ok = i.ok === true ? true : i.ok === false ? false : null; t.value = clip(i.value, 60); t.note = clip(i.note, 500); t.photos = (Array.isArray(i.photos) ? i.photos : []).filter((p) => /^[0-9a-f]{24}$/.test(p)).slice(0, 6); });
          if (b.status === 'done') { if (x.items.some((i) => i.ok === null)) return err(400, 'Check every item (OK or Not OK) before signing off.'); x.status = 'done'; x.doneAt = new Date().toISOString(); x.doneBy = me.name; }
          if (b.status === 'open' && isEA(me)) { x.status = 'open'; x.doneAt = ''; x.doneBy = ''; }
          x.updated = new Date().toISOString(); await store.setJSON(key, x); await idx(x); return json(200, { insp: x });
        }
      }
    }
    /* assembly instructions per part */
    if (seg[0] === 'instr' && seg.length === 2 && /^p[0-9a-f]{10}$/.test(seg[1])) {
      const key = `instr/${seg[1]}`;
      if (M === 'GET') return json(200, { instr: (await store.get(key, { type: 'json' })) || { steps: [] } });
      if (M === 'PUT') {
        const b = await req.json().catch(() => ({}));
        const steps = (Array.isArray(b.steps) ? b.steps : []).map((x) => ({ text: clip(String(x.text || '').trim(), 1000), torque: clip(String(x.torque || '').trim(), 40), photo: /^[0-9a-f]{24}$/.test(x.photo || '') ? x.photo : '' })).filter((x) => x.text || x.photo).slice(0, 60);
        const rec = { steps, by: me.name, at: new Date().toISOString() }; await store.setJSON(key, rec); await audit(store, seg[1], me, 'Assembly instructions updated', `${steps.length} step${steps.length === 1 ? '' : 's'}`);
        return json(200, { instr: rec });
      }
    }
    /* vehicles (chassis numbers) with build record */
    if (seg[0] === 'vehicles') {
      if (seg.length === 1 && M === 'GET') return json(200, { vehicles: Object.values(await getMeta(store, 'vehicles')).sort((a, b) => String(b.start || b.at).localeCompare(String(a.start || a.at))) });
      const ch = decodeURIComponent(seg[1] || '');
      if (seg.length === 2 && CHASSIS_RE.test(ch)) {
        if (M === 'PUT') {
          const b = await req.json().catch(() => ({})); const d = (x) => (/^\d{4}-\d{2}-\d{2}$/.test(x || '') ? x : '');
          const rec = await updMeta(store, 'vehicles', (m) => { const r = { chassis: ch, ...(m[ch] || {}) };
            if (b.type !== undefined) r.type = clip(b.type, 20); if (b.customer !== undefined) r.customer = clip(b.customer, 80); if (b.note !== undefined) r.note = clip(b.note, 1000);
            if (b.status !== undefined && ['build', 'done', 'delivered'].includes(b.status)) r.status = b.status; if (b.start !== undefined) r.start = d(b.start); if (b.delivered !== undefined) r.delivered = d(b.delivered);
            r.status = r.status || 'build'; r.by = me.name; r.at = new Date().toISOString(); m[ch] = r; return r; });
          return json(200, { vehicle: rec });
        }
        if (M === 'GET') {
          const v = (await getMeta(store, 'vehicles'))[ch] || null, fit = await getMeta(store, 'fit');
          const fitted = Object.entries(fit).flatMap(([id, arr]) => (arr || []).filter((x) => x.chassis.toLowerCase() === ch.toLowerCase()).map((x) => ({ ...x, id })));
          const insp = Object.values(await getMeta(store, 'inspindex')).filter((x) => String(x.chassis).toLowerCase() === ch.toLowerCase());
          const ncr = Object.values(await getMeta(store, 'ncrindex')).filter((x) => String(x.chassis).toLowerCase() === ch.toLowerCase());
          return json(200, { vehicle: v, fitted, insp, ncr });
        }
      }
    }
    if (path === 'admin/usage' && M === 'GET' && me.role === 'admin') {
      const out = []; for (let i = 0; i < 7; i++) { const d = new Date(Date.now() - i * 864e5).toLocaleDateString('sv-SE', { timeZone: 'Europe/Amsterdam' }); const l = await listJSON(store, `usage/${d}/`); out.push({ day: d, n: l.reduce((s, x) => s + (x.n || 0), 0), users: l.length }); }
      return json(200, { usage: out });
    }
    if (path === 'admin/overdue' && M === 'POST' && me.role === 'admin') { const e = await runOverdue(env, 'manual'); return e && !e.startsWith('Nothing') ? err(502, e) : json(200, { ok: true, msg: e || 'Overdue summary sent.' }); }

    /* ---- admin ---- */
    if (seg[0] === 'admin') {
      if (me.role !== 'admin') return err(403, 'Only the administrator can do this.');
      if (path === 'admin/users' && M === 'GET') return json(200, { users: (await listJSON(store, 'users/')).map(publicUser).sort((a, b) => a.name.localeCompare(b.name)) });
      if (path === 'admin/users' && M === 'POST') {
        const b = await req.json().catch(() => ({}));
        const u = String(b.username || '').trim().toLowerCase(), name = clip(String(b.name || '').trim(), 60), pw = String(b.password || '');
        if (!USER_RE.test(u) || u === 'admin') return err(400, 'Username: 2–32 letters, digits, dot, dash or underscore.');
        if (!name) return err(400, 'Enter the full name.');
        if (pw.length < 6) return err(400, 'Password: at least 6 characters.');
        if (await store.get(`users/${u}`, { type: 'json' })) return err(409, 'This username already exists.');
        const rec = { u, name, role: ROLES[b.role] ? b.role : 'mechanic', active: true, v: 1, created: new Date().toISOString(), ...hashPw(pw) };
        await store.setJSON(`users/${u}`, rec); return json(200, { user: publicUser(rec) });
      }
      if (seg[1] === 'users' && seg.length === 3 && USER_RE.test(seg[2]) && M === 'DELETE') {
        const u = seg[2], key = `users/${u}`;
        if (u === me.u) return err(400, 'You cannot delete the account you are signed in with.');
        if (!(await store.get(key, { type: 'json' }))) return err(404, 'Account not found.');
        const { blobs: dr } = await store.list({ prefix: `drafts/${u}/` });
        await Promise.all(dr.map((b) => deleteDraft(u, b.key.split('/').pop())));
        const { blobs: ph } = await PHOTOS().list({ prefix: `${u}/` }); // loose photos not linked to a draft
        await Promise.all(ph.map((b) => PHOTOS().delete(b.key)));
        await store.delete(key); // sent reports stay in the history ("Sent · name · date")
        return json(200, { ok: true });
      }
      if (seg[1] === 'users' && seg.length === 4 && USER_RE.test(seg[2])) {
        const key = `users/${seg[2]}`, rec = await store.get(key, { type: 'json' });
        if (!rec) return err(404, 'Account not found.');
        const b = await req.json().catch(() => ({}));
        if (seg[3] === 'password' && M === 'POST') { if (String(b.password || '').length < 6) return err(400, 'Password: at least 6 characters.'); Object.assign(rec, hashPw(b.password)); rec.v = (rec.v || 1) + 1; }
        else if (seg[3] === 'active' && M === 'POST') { rec.active = !!b.active; rec.v = (rec.v || 1) + 1; }
        else if (seg[3] === 'role' && M === 'POST') { if (!ROLES[b.role]) return err(400, 'Unknown role.'); if (seg[2] === me.u && b.role !== 'admin') return err(400, 'You cannot remove your own Admin role.'); rec.role = b.role; }
        else return err(404, 'Unknown action.');
        await store.setJSON(key, rec); return json(200, { user: publicUser(rec) });
      }
      /* maintenance: reset a vehicle / category completely (no e-mail, logged in the app) */
      if (path === 'admin/reset' && M === 'POST') {
        const b = await req.json().catch(() => ({}));
        if (String(b.confirm || '').trim().toUpperCase() !== 'CONFIRM') return err(400, 'Type CONFIRM to reset the list.');
        const ids = new Set((Array.isArray(b.ids) ? b.ids : []).slice(0, 2000).map(String).filter((x) => /^p[0-9a-f]{10}$/.test(x)));
        if (!ids.size) return err(400, 'No parts in this selection.');
        const inScope = (id) => { const q = parseId(id); return !!q && ids.has(q.base); };
        let reports = 0, drafts = 0, notes = 0;
        const { blobs } = await store.list({ prefix: 'send/' });
        for (const x of blobs) { const rec = await store.get(x.key, { type: 'json' }); if (!rec) continue; const keep = (rec.rows || []).filter((r) => !inScope(r.id));
          if (keep.length === (rec.rows || []).length) continue; reports += rec.rows.length - keep.length;
          if (keep.length) await store.setJSON(x.key, { ...rec, rows: keep }); else await store.delete(x.key); }
        const { blobs: dr } = await store.list({ prefix: 'drafts/' });
        for (const k of dr.map((x) => x.key.split('/')).filter((k) => k.length === 3 && inScope(k[2]))) { await deleteDraft(k[1], k[2]); drafts++; }
        const { blobs: cl } = await store.list({ prefix: 'clear/' });
        for (const x of cl.filter((x) => ids.has(x.key.slice(6)))) { await store.delete(x.key); notes++; }
        const { blobs: rl } = await store.list({ prefix: 'release/' });
        for (const x of rl) { const q = parseId(x.key.slice(8)); if (q && ids.has(q.base)) await store.delete(x.key); }
        for (const k of ['appr', 'fit', 'done']) await updMeta(store, k, (m) => { for (const id of Object.keys(m)) if (inScope(id)) delete m[id]; });
        const ev = { at: new Date().toISOString(), by: me.name, u: me.u, scope: clip(b.scope, 160), parts: ids.size, reports, drafts, notes };
        await store.setJSON(`maint/${ev.at}-${crypto.randomBytes(3).toString('hex')}`, ev);
        return json(200, { ok: true, ...ev });
      }
      if (path === 'admin/backup' && M === 'POST') {
        const e = await runBackup(env, 'manual'); return e ? err(502, e) : json(200, { ok: true });
      }
      if (path === 'admin/maint' && M === 'GET') {
        const log = (await listJSON(store, 'maint/')).sort((a, b) => String(b.at).localeCompare(String(a.at))).slice(0, 40);
        return json(200, { log });
      }
      if (path === 'admin/team' && M === 'PUT') {
        const b = await req.json().catch(() => ({}));
        const names = [...new Set((b.names || []).map((n) => clip(String(n).trim(), 60)).filter(Boolean))].slice(0, 60);
        await store.setJSON('config/team', { names }); return json(200, { names });
      }
    }
    return err(404, 'Not found.');
  } catch (e) {
    console.error(e);
    return err(500, 'Something went wrong on the server. Try again.');
  }
};

export const config = { path: '/api/*' };
