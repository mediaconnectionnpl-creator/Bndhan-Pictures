/**
 * Bandhan Pictures - backend (Google Apps Script)
 * Paste this whole file into: Google Sheet > Extensions > Apps Script.
 * Then follow LAUNCH-GUIDE.md (Step 2).
 */

// ====== SET THESE THREE LINES ======
const OWNER_EMAIL = '';                         // your email: you get a message for every new booking, order, message, review
const ADMIN_PASSWORD = 'change-this-password';  // your admin password (you can change it later inside the admin panel)
const ROOT_FOLDER_ID = '1STmsfp0BYTasgqiDimKVNTl5fcfIduKG';   // your BANDHAN PICTURES folder on Drive
// ====================================

const PHOTOS_FOLDER_ID = '1V5n1sS84B_lJrCH844xGZCfEdoB4Ttml';  // Photos sub-folder (uploads are stored here)
const COLLECTIONS = ['bookings', 'orders', 'messages', 'reviews', 'subscribers'];
const PERM = { bookings: ['bookings', 'track'], orders: ['orders'], messages: ['msgs'], reviews: ['inbox'], subscribers: ['subs'] };
const MAX_JSON = 6000;

function doGet() { return out({ ok: true, name: 'Bandhan Pictures API' }); }

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, err: 'bad' }); }
  try {
    switch (d.a) {
      case 'content': return out(getContent());
      case 'add': return out(addRecord(d.c, d.rec));
      case 'track': return out(track(d.ref, d.phone));
      case 'login': return out(login(d.user, d.pass));
    }
    const s = session(d.t);
    if (!s) return out({ ok: false, err: 'auth' });
    switch (d.a) {
      case 'list': return out(listRecords(d.k, s));
      case 'replace': return out(replaceRecords(d.k, d.list, d.since, s));
      case 'saveContent': return out(saveContent(d, s));
      case 'upload': return out(upload(d, s));
      case 'setPassword': return out(setPassword(d, s));
    }
    return out({ ok: false, err: 'unknown' });
  } catch (err) {
    return out({ ok: false, err: String(err) });
  }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function setup() { COLLECTIONS.forEach(sheet); DriveApp.getFolderById(ROOT_FOLDER_ID); if (OWNER_EMAIL) MailApp.sendEmail(OWNER_EMAIL, 'Bandhan Pictures backend is ready', 'Your Google Sheet and Drive are connected.'); }

// ---------- storage helpers ----------
function sheet(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(n);
  if (!sh) { sh = ss.insertSheet(n); sh.appendRow(['id', 'ts', 'json']); }
  return sh;
}
function rows(n) {
  const v = sheet(n).getDataRange().getValues();
  return v.slice(1).filter(r => r[0]).map(r => { let o = {}; try { o = JSON.parse(r[2]); } catch (e) {} o.id = String(r[0]); o.ts = Number(r[1]); return o; });
}
function jf(name) { const it = DriveApp.getFolderById(ROOT_FOLDER_ID).getFilesByName(name); return it.hasNext() ? it.next() : null; }
function readJson(n) { const f = jf(n); if (!f) return null; try { return JSON.parse(f.getBlob().getDataAsString()); } catch (e) { return null; } }
function writeJson(n, o) { const f = jf(n), t = JSON.stringify(o); if (f) f.setContent(t); else DriveApp.getFolderById(ROOT_FOLDER_ID).createFile(n, t, 'application/json'); }
function props() { return PropertiesService.getScriptProperties(); }
function getStaff() { try { return JSON.parse(props().getProperty('STAFF') || '[]'); } catch (e) { return []; } }
function pubStaff(a) { return a.map(x => ({ id: x.id, user: x.user, name: x.name, perm: x.perm || [], ro: !!x.ro, on: x.on !== false })); }
function hashPw(pw, salt) {
  let b = pw + '|' + salt;
  for (let i = 0; i < 1000; i++) b = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, b));
  return b;
}
function ownerPw() { return props().getProperty('ADMIN_PASSWORD') || ADMIN_PASSWORD; }

// ---------- public ----------
function addRecord(c, rec) {
  if (COLLECTIONS.indexOf(c) < 0 || !rec || typeof rec !== 'object') return { ok: false, err: 'bad' };
  const cache = CacheService.getScriptCache(), k = 'rl' + Math.floor(Date.now() / 60000), n = Number(cache.get(k) || 0);
  if (n > 80) return { ok: false, err: 'busy' };
  cache.put(k, String(n + 1), 120);
  rec = JSON.parse(JSON.stringify(rec));
  if (c === 'reviews') rec.ok = false;
  if (c === 'bookings') { rec.st = 'Pending'; delete rec.paid; delete rec.stage; delete rec.hist; delete rec.inote; delete rec.tnote; delete rec.eta; }
  rec.id = String(rec.id || Utilities.getUuid()).slice(0, 40);
  delete rec.ts;
  const json = JSON.stringify(rec);
  if (json.length > MAX_JSON) return { ok: false, err: 'big' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (rows(c).some(r => r.id === rec.id)) return { ok: true, dup: true };
    sheet(c).appendRow([rec.id, Date.now(), json]);
  } finally { lock.releaseLock(); }
  if (OWNER_EMAIL) {
    try {
      const label = { bookings: 'booking', orders: 'order', messages: 'message', reviews: 'review/feedback', subscribers: 'subscriber' }[c];
      MailApp.sendEmail(OWNER_EMAIL, 'New ' + label + ' - Bandhan Pictures', Object.keys(rec).map(k2 => k2 + ': ' + (typeof rec[k2] === 'object' ? JSON.stringify(rec[k2]) : rec[k2])).join('\n'));
    } catch (e) {}
  }
  return { ok: true };
}

function getContent() {
  const cache = CacheService.getScriptCache(), hit = cache.get('content');
  if (hit) return JSON.parse(hit);
  const cfg = readJson('content.json') || {};
  const reviews = rows('reviews').filter(r => r.ty === 'review' && r.ok).map(r => ({ n: r.n, s: r.s, t: r.t, ty: 'review' }));
  const taken = rows('bookings').filter(r => r.st === 'Confirmed' || r.st === 'Completed').map(r => r.date);
  const res = { ok: true, cfg: cfg, reviews: reviews, taken: taken };
  try { cache.put('content', JSON.stringify(res), 30); } catch (e) {}
  return res;
}

function track(ref, phone) {
  ref = String(ref || '').trim().toUpperCase();
  const ph = String(phone || '').replace(/\D/g, '').slice(-10);
  if (!ref || ph.length < 7) return { ok: false };
  const r = rows('bookings').find(x => String(x.ref || '').toUpperCase() === ref && String(x.ph || '').replace(/\D/g, '').slice(-10) === ph);
  if (!r) return { ok: false };
  return { ok: true, rec: { ref: r.ref, ev: r.ev, stage: Number(r.stage) || 0, note: r.tnote || '', eta: r.eta || '', hist: r.hist || [], upd: ((r.hist || []).slice(-1)[0] || [0, r.date])[1] } };
}

// ---------- login / session ----------
function login(user, pass) {
  const u = String(user || 'admin').toLowerCase().trim() || 'admin', cache = CacheService.getScriptCache(), fk = 'f_' + u;
  if (Number(cache.get(fk) || 0) >= 8) return { ok: false, err: 'locked' };
  pass = String(pass || '');
  let s = null, me = null;
  if (u === 'admin') { if (pass === ownerPw()) s = { role: 'owner', id: 'owner' }; }
  else {
    const st = getStaff().find(x => x.user === u && x.on !== false);
    if (st && hashPw(pass, st.salt) === st.h) { s = { role: 'staff', id: st.id, perm: st.perm || [], ro: !!st.ro }; me = pubStaff([st])[0]; }
  }
  if (!s) { cache.put(fk, String(Number(cache.get(fk) || 0) + 1), 900); return { ok: false, err: 'bad' }; }
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('t_' + token, JSON.stringify(s), 21600);
  return { ok: true, token: token, role: s.role, me: me, staff: s.role === 'owner' ? pubStaff(getStaff()) : null };
}
function session(t) {
  if (!t) return null;
  const v = CacheService.getScriptCache().get('t_' + t);
  return v ? JSON.parse(v) : null;
}
function canColl(s, k, write) {
  if (COLLECTIONS.indexOf(k) < 0) return false;
  if (s.role === 'owner') return true;
  if (write && s.ro) return false;
  return PERM[k].some(p => (s.perm || []).indexOf(p) >= 0);
}

// ---------- admin ----------
function listRecords(k, s) {
  if (!canColl(s, k, false)) return { ok: false, err: 'perm' };
  return { ok: true, list: rows(k), now: Date.now() };
}
function replaceRecords(k, list, since, s) {
  if (!canColl(s, k, true) || !Array.isArray(list)) return { ok: false, err: 'perm' };
  since = Number(since) || 0;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const old = rows(k), oldMap = {};
    old.forEach(r => { oldMap[r.id] = r; });
    const seen = {}, final = [];
    list.forEach(r => {
      const id = String(r.id || Utilities.getUuid()).slice(0, 40);
      if (seen[id]) return; seen[id] = 1;
      const ts = oldMap[id] ? oldMap[id].ts : Date.now();
      const o = JSON.parse(JSON.stringify(r)); delete o.id; delete o.ts;
      if (k === 'reviews' && s.role !== 'owner' && !(s.perm || []).includes('inbox')) return;
      final.push([id, ts, JSON.stringify(o)]);
    });
    old.forEach(r => { if (!seen[r.id] && r.ts > since) { const o = JSON.parse(JSON.stringify(r)); delete o.id; delete o.ts; final.push([r.id, r.ts, JSON.stringify(o)]); } });
    const sh = sheet(k);
    sh.clearContents();
    sh.getRange(1, 1, 1, 3).setValues([['id', 'ts', 'json']]);
    if (final.length) sh.getRange(2, 1, final.length, 3).setValues(final);
    CacheService.getScriptCache().remove('content');
  } finally { lock.releaseLock(); }
  return { ok: true, now: Date.now() };
}
function saveContent(d, s) {
  if (s.role !== 'owner' && s.ro) return { ok: false, err: 'perm' };
  const cfg = d.cfg && typeof d.cfg === 'object' ? d.cfg : {};
  ['staff', 'pw', 'pwh', 'track'].forEach(k => delete cfg[k]);
  if (s.role !== 'owner') { const old = readJson('content.json') || {}; cfg.site = old.site; cfg.blocked = old.blocked; }
  writeJson('content.json', cfg);
  CacheService.getScriptCache().remove('content');
  let staff = null;
  if (s.role === 'owner' && Array.isArray(d.staff)) {
    const old = getStaff(), oldMap = {};
    old.forEach(x => { oldMap[x.id] = x; });
    const list = d.staff.map(x => {
      const prev = oldMap[x.id] || {};
      const salt = prev.salt || Utilities.getUuid();
      const h = x.pw ? hashPw(String(x.pw), salt) : prev.h;
      return { id: String(x.id), user: String(x.user).toLowerCase().replace(/[^a-z0-9]/g, ''), name: String(x.name || ''), perm: x.perm || [], ro: !!x.ro, on: x.on !== false, salt: salt, h: h };
    }).filter(x => x.h && x.user && x.user !== 'admin');
    props().setProperty('STAFF', JSON.stringify(list));
    staff = pubStaff(list);
  }
  return { ok: true, staff: staff };
}
function upload(d, s) {
  if (s.role !== 'owner' && s.ro) return { ok: false, err: 'perm' };
  const mime = String(d.mime || '');
  if (['image/jpeg', 'image/png', 'image/webp', 'image/gif'].indexOf(mime) < 0) return { ok: false, err: 'type' };
  if (!d.b64 || d.b64.length > 8000000) return { ok: false, err: 'size' };
  const blob = Utilities.newBlob(Utilities.base64Decode(d.b64), mime, String(d.name || 'photo.jpg').replace(/[^\w.\-]/g, '_'));
  const f = DriveApp.getFolderById(PHOTOS_FOLDER_ID).createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: 'https://lh3.googleusercontent.com/d/' + f.getId() };
}
function setPassword(d, s) {
  if (s.role !== 'owner') return { ok: false, err: 'perm' };
  if (String(d.old) !== ownerPw() || String(d.pw || '').length < 6) return { ok: false, err: 'bad' };
  props().setProperty('ADMIN_PASSWORD', String(d.pw));
  return { ok: true };
}
