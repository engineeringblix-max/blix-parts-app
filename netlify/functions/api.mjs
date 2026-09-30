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
const PART_RE = /^p[0-9a-f]{10}$/;
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
  const rows = (p.rows || []).map((r) => `<tr>${td(r.vehicle, 'font-weight:600;color:#012135')}${td(r.part, 'font-family:Consolas,monospace;font-weight:600')}${td(r.status)}${td(r.priority)}${td(r.note)}${td(r.responsible)}${td(r.date, 'white-space:nowrap')}${td(r.photos || 0, 'text-align:center')}</tr>`).join('');
  const th = (t) => `<th style="text-align:left;padding:7px 8px;background:#012135;color:#fff;font-weight:600">${t}</th>`;
  return `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#14232c"><div style="background:#012135;color:#fff;padding:14px 16px;border-bottom:4px solid #0bbbef"><b style="font-size:17px">BLIX Parts Review</b><br><span style="color:#a9dff2">Report from ${esc(reporter)}${p.batches > 1 ? ` · file ${p.batch} of ${p.batches}` : ''}</span></div>
  <p>${(p.rows || []).length} part(s) reported. The attached Excel file contains every report with the part preview and the photos.</p>
  <table style="border-collapse:collapse;font-size:13px;width:100%"><tr>${th('Vehicle')}${th('Part')}${th('Status')}${th('Priority')}${th('What to change')}${th('Responsible')}${th('Date')}${th('Photos')}</tr>${rows}</table>
  <p style="color:#5d6e78;font-size:12px">Sent automatically by the BLIX Parts app.</p></div>`;
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
      const drafts = me.u === 'admin' ? [] : await listJSON(store, `drafts/${me.u}/`);
      const team = (await store.get('config/team', { type: 'json' })) || { names: [] };
      return json(200, { user: me, drafts, parts: await sharedStatus(store), team: team.names });
    }

    /* drafts */
    if (seg[0] === 'drafts' && seg.length === 2 && PART_RE.test(seg[1]) && me.u !== 'admin') {
      if (M === 'PUT') { const d = cleanDraft(await req.json(), seg[1]); await store.setJSON(`drafts/${me.u}/${seg[1]}`, d); return json(200, { draft: d }); }
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
      const useResend = !!env.RESEND_API_KEY;
      const need = useResend ? ['RESEND_API_KEY', 'MAIL_TO'] : ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'MAIL_FROM', 'MAIL_TO'];
      const miss = need.filter((k) => !env[k]);
      if (miss.length) return err(500, useResend || !env.MS_TENANT_ID
        ? `The mail service is not set up yet (missing: ${useResend ? miss.join(', ') : 'RESEND_API_KEY, MAIL_TO'}).`
        : `The mail service is not set up yet (missing: ${miss.join(', ')}).`);
      const p = await req.json().catch(() => null);
      if (!p || typeof p.file !== 'string' || !/^[\w\-. ()]+\.xlsx$/.test(p.filename || '')) return err(400, 'The report is incomplete.');
      if (p.file.length * 0.75 > MAX_FILE) return err(413, 'The report is too large to e-mail. Send fewer parts at once.');
      const rows = (p.rows || []).filter((r) => r && PART_RE.test(r.id || ''));
      const vehicles = [...new Set(rows.map((r) => r.vehicle))].join(', ');
      const subject = `BLIX parts report – ${me.name} – ${vehicles || 'parts'} (${rows.length} part${rows.length === 1 ? '' : 's'})${p.batches > 1 ? ` [${p.batch}/${p.batches}]` : ''}`;
      const to = env.MAIL_TO.split(',').map((a) => a.trim()).filter(Boolean);
      const html = mailHtml({ ...p, rows }, me.name);
      if (useResend) {
        const r = await fetch('https://api.resend.com/emails', {
          method: 'POST', headers: { Authorization: `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ from: env.RESEND_FROM || 'BLIX Parts App <onboarding@resend.dev>', to, subject, html,
            attachments: [{ filename: p.filename, content: p.file }] }),
        });
        if (!r.ok) { const t = await r.text(); console.error('Resend failed', r.status, t);
          let m = ''; try { m = JSON.parse(t).message || ''; } catch (e) {}
          return err(502, `The e-mail could not be sent (Resend ${r.status}${m ? ': ' + m : ''}).`); }
      } else {
        const token = await graphToken(env);
        const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(env.MAIL_FROM)}/sendMail`, {
          method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ message: { subject, body: { contentType: 'HTML', content: html },
            toRecipients: to.map((address) => ({ emailAddress: { address } })),
            attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: p.filename, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes: p.file }] },
            saveToSentItems: true }),
        });
        if (!r.ok) { console.error('Graph sendMail failed', r.status, await r.text()); return err(502, 'The e-mail could not be sent by Microsoft 365.'); }
      }
      const at = new Date().toISOString();
      const rec = rows.map((x) => ({ id: x.id, statusKey: clip(x.statusKey, 6), status: clip(x.status, 40), priorityKey: clip(x.priorityKey, 4), priority: clip(x.priority, 20),
        note: clip(x.note, 4000), responsible: clip(x.responsible, 80), dateIso: /^\d{4}-\d{2}-\d{2}$/.test(x.dateIso || '') ? x.dateIso : '', date: clip(x.date, 30),
        photos: Number(x.photos) || 0, vehicle: clip(x.vehicle, 10), part: clip(x.part, 80) }));
      await store.setJSON(`send/${at}-${crypto.randomBytes(3).toString('hex')}`, { by: me.name, u: me.u, at, rows: rec });
      await Promise.all(rows.map((x) => deleteDraft(me.u, x.id))); // photos went out by e-mail; clear the draft
      return json(200, { ok: true, at });
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
