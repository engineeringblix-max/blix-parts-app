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
async function sendMail(env, { subject, html, file }) {
  const to = env.MAIL_TO.split(',').map((a) => a.trim()).filter(Boolean);
  if (env.RESEND_API_KEY) {
    const r = await fetch('https://api.resend.com/emails', {
      method: 'POST', headers: { Authorization: `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ from: env.RESEND_FROM || 'BLIX Parts App <onboarding@resend.dev>', to, subject, html,
        ...(file ? { attachments: [{ filename: file.name, content: file.b64 }] } : {}) }),
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
      attachments: file ? [{ '@odata.type': '#microsoft.graph.fileAttachment', name: file.name, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes: file.b64 }] : [] },
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
  for (const s of sends) for (const r of s.rows || []) parts[r.id] = { ...r, by: s.by, at: s.at };
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
      return json(200, { user: me, drafts, parts, team: team.names, cleared: await clearedMap(store) });
    }

    /* drafts */
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
      await Promise.all(rows.map((x) => deleteDraft(me.u, x.id))); // photos went out by e-mail; clear the draft
      return json(200, { ok: true, at });
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
      for (const p of ev.parts) {
        const key = `clear/${p.id}`, old = (await store.get(key, { type: 'json' })) || { base: p.id, history: [] };
        const last = { at: ev.at, by: ev.by, reason, from: p.from, fromName: p.fromName, reports: rows.filter((r) => r.base === p.id).length, drafts: p.drafts };
        await store.setJSON(key, { base: p.id, last, history: [...(old.history || []), last].slice(-30) });
      }
      return json(200, { ok: true, at: ev.at, parts: ev.parts.length });
    }

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
        const rec = { u, name, role: b.role === 'admin' ? 'admin' : 'mechanic', active: true, v: 1, created: new Date().toISOString(), ...hashPw(pw) };
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
        else return err(404, 'Unknown action.');
        await store.setJSON(key, rec); return json(200, { user: publicUser(rec) });
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
