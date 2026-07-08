// ── FIREBASE CONFIG ───────────────────────────────────
const firebaseConfig = {
  apiKey:            "AIzaSyABDc89YeU0QRYRtayDQHJmCocMg5MQARw",
  authDomain:        "cs-inventory-6906b.firebaseapp.com",
  projectId:         "cs-inventory-6906b",
  storageBucket:     "cs-inventory-6906b.firebasestorage.app",
  messagingSenderId: "709498580037",
  appId:             "1:709498580037:web:584307178d4f253268072c"
};
firebase.initializeApp(firebaseConfig);
const fdb   = firebase.firestore();
const fauth = firebase.auth();

// ── CONSTANTS ─────────────────────────────────────────
const STATUSES    = ['Active','Available','Reserved','Inactive'];
const ACT_LABELS  = {Added:'b-active',Updated:'b-reserved',Deleted:'b-inactive','CSV Upload':'b-available',Exported:'b-available',Login:'b-available',Backup:'b-active'};
const CSV_HEADERS = ['Client','Product','Number','Status','Remarks','Posted Status','Posted Date','Posted Time','Client OSF','Client MRC','Client OTRF','Client Channel Fee','Client CPM','Effective Date','Activated Date','Provider','Arrival Date','Provider Activation Date','Provider OSF','Provider MRC','Provider OTRF','Provider CPM','Type / Session','Route Request by','Deactivation Date','Previous Client'];
const CSV_FIELD_MAP = {'Client':'client','Product':'product','Number':'number','Status':'status','Remarks':'remarks','Posted Status':'postedStatus','Posted Date':'postedDate','Client OSF':'clientOSF','Client MRC':'clientMRC','Client OTRF':'clientOTRF','Client Channel Fee':'clientCF','Client CPM':'clientCPM','Effective Date':'effDate','Activated Date':'actDate','Provider':'provider','Arrival Date':'arrDate','Provider Activation Date':'provActDate','Provider OSF':'provOSF','Provider MRC':'provMRC','Provider OTRF':'provOTRF','Provider CPM':'provCPM','Type / Session':'typeSession','Route Request by':'route','Deactivation Date':'deactDate','Previous Client':'prevClient'};
const FIELD_LABELS = {client:'Client',product:'Product',number:'Number',status:'Status',remarks:'Remarks',postedStatus:'Posted Status',postedDate:'Posted Date',postedHour:'Posted Hour',postedMin:'Posted Minute',clientOSF:'Client OSF',clientMRC:'Client MRC',clientOTRF:'Client OTRF',clientCF:'Client Channel Fee',clientCPM:'Client CPM',effDate:'Effective Date',actDate:'Activated Date',provider:'Provider',arrDate:'Arrival Date',provActDate:'Provider Activation Date',provOSF:'Provider OSF',provMRC:'Provider MRC',provOTRF:'Provider OTRF',provCPM:'Provider CPM',typeSession:'Type / Session',route:'Route Request by',deactDate:'Deactivation Date',prevClient:'Previous Client'};
const DATE_FIELDS = new Set(['mPostedDate','mEffDate','mActDate','mArrDate','mProvActDate','mDeactDate']);
const VALID_STATUSES = new Set(['Active','Available','Reserved','Inactive','']);
const DATE_CSV_FIELDS = ['postedDate','effDate','actDate','arrDate','provActDate','deactDate'];

// ── AUTO BACKUP (weekly email) ────────────────────────
// A full CSV of the inventory is emailed once a week (Friday) to the logged-in
// address. Because this is a pure client-side app, "every Friday" means: the
// first time an editor opens the app on Friday/Sat/Sun of a week it hasn't sent
// yet, it sends once (a Firestore transaction ensures only ONE client sends).
// Mail is delivered by a tiny Google Apps Script web app (see setup notes) that
// sends the file as a real attachment from your own Gmail — no third-party, no
// Firebase billing. Paste your deployment URL + shared secret below.
const BACKUP_MAILER_URL = 'https://script.google.com/macros/s/AKfycbyGdK5FJIR5-XZoKifgqeUvXBX28SJ5pb_akszkjh7Qtlc1xZ7kAEqm2gG-EIYq-MrJ/exec';
const BACKUP_MAILER_KEY = 'cs-inv-9f3k2p7q-backup-2026';   // ⚠ PASTE your Apps Script SECRET here (the var SECRET='...' value) — must match it EXACTLY
let   AB = null;                 // cached meta/autoBackup config { enabled, recipient, lastSentWeek, lastSentAt }
let   _abTimer = null, _abDeferT = null, _abSending = false;
// ── UTILITIES ─────────────────────────────────────────
function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
const fmt = iso => iso ? iso.replace(/(\d{4})-(\d{2})-(\d{2})/,'$2/$3/$1') : '—';
function sanitizeDate(v) {
  if (!v) return '';
  const s = String(v).trim();
  // YYYY-MM-DD
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  // MM/DD/YYYY
  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m2) return `${m2[3]}-${m2[1].padStart(2,'0')}-${m2[2].padStart(2,'0')}`;
  // D-Mon-YY or D-Mon-YYYY (e.g. 8-Dec-14, 27-May-2020)
  const MONTHS = {jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'};
  const m3 = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
  if (m3) {
    const mon = MONTHS[m3[2].toLowerCase()];
    if (mon) {
      const yr = m3[3].length === 2 ? (parseInt(m3[3]) >= 50 ? '19' : '20') + m3[3] : m3[3];
      return `${yr}-${mon}-${m3[1].padStart(2,'0')}`;
    }
  }
  return '';
}
function parseCSVLine(line) {
  const res=[]; let cur='', q=false;
  for (let i=0; i<line.length; i++) {
    const c=line[i];
    if (c==='"') { if(q&&line[i+1]==='"'){cur+='"';i++;} else q=!q; }
    else if (c===','&&!q) { res.push(cur); cur=''; }
    else cur+=c;
  }
  res.push(cur); return res;
}
function decodeText(bytes, encoding, options) {
  return new TextDecoder(encoding, options).decode(bytes).replace(/^\uFEFF/, '');
}
async function readCSVText(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (bytes[0]===0xFF && bytes[1]===0xFE) return decodeText(bytes, 'utf-16le');
  if (bytes[0]===0xFE && bytes[1]===0xFF) return decodeText(bytes, 'utf-16be');
  try {
    return decodeText(bytes, 'utf-8', {fatal:true});
  } catch(e) {
    // Excel CSV files saved with an ANSI code page commonly contain Windows-1252 bytes.
    return decodeText(bytes, 'windows-1252');
  }
}
// Strip formatting and leading country/area codes to get a bare local number for comparison.
// Treats +63/63/0/02/2 prefixes as equivalent so different regional formats match the same entry.
function normalizePhone(n) {
  let digits = String(n == null ? '' : n).replace(/\D/g, '');
  // Step 1: strip country code or trunk prefix
  if (digits.startsWith('63'))       digits = digits.slice(2);
  else if (digits.startsWith('02'))  digits = digits.slice(2);
  else if (digits.startsWith('0'))   digits = digits.slice(1);
  // Step 2: strip Metro Manila area code '2' from 9-digit numbers (e.g. 279182881 → 79182881)
  if (digits.startsWith('2') && digits.length === 9) digits = digits.slice(1);
  return digits;
}
function bclass(s) { return {Active:'b-active',Available:'b-available',Reserved:'b-reserved',Inactive:'b-inactive'}[s]||''; }
// Canonicalize Posted Status to the current labels, migrating legacy values on the fly:
//   "Yes" → "Posted", "Not Yet" → "For Posting". Blank/unknown values pass through unchanged.
function canonPostedStatus(v) {
  const raw = String(v == null ? '' : v).trim();
  const s = raw.toLowerCase();
  if (s === 'posted' || s === 'yes')          return 'Posted';
  if (s === 'for posting' || s === 'not yet') return 'For Posting';
  if (s === 'no')                             return 'No';
  return raw;
}
function postedClass(s) {
  switch (canonPostedStatus(s)) {
    case 'Posted':      return 'b-posted-yes';
    case 'For Posting': return 'b-posted-notyet';
    default:            return 'b-posted-no';
  }
}
function pad2(n) { return String(n).padStart(2, '0'); }
// Advance a {h,m} time by one minute, wrapping 23:59 → 00:00.
function stepMinute(h, m) {
  m += 1;
  if (m > 59) { m = 0; h += 1; }
  if (h > 23) { h = 0; }
  return { h, m };
}
// Highest time among entries marked "Posted", as {h,m} (or null if none). This is the point
// the "For Posting" run continues from (+1 minute).
function latestPostedAnchor() {
  let best = null; // minutes-of-day
  for (const r of DB) {
    if (canonPostedStatus(r.postedStatus) !== 'Posted') continue;
    if (r.postedHour === '' || r.postedHour == null) continue;
    const mins = (parseInt(r.postedHour, 10) || 0) * 60 + (parseInt(r.postedMin || '0', 10) || 0);
    if (best === null || mins > best) best = mins;
  }
  return best === null ? null : { h: Math.floor(best / 60), m: best % 60 };
}
// Renumber every "For Posting" entry into one chronological run based on table position:
// the bottom row is earliest and each row above is +1 minute (wrapping 23:59 → 00:00). The
// bottom row continues from the latest "Posted" time (+1); if nothing is posted yet it keeps
// its own current time, otherwise it seeds at the current clock time. "Posted" entries are
// never touched, and only records whose time actually changes are written to Firestore.
async function resequencePostingTimes() {
  // DB is in display order (top→bottom) after refreshInventoryRecent(); bottom = last row.
  const ordered = DB.filter(r => canonPostedStatus(r.postedStatus) === 'For Posting').reverse();
  if (!ordered.length) return [];

  let cur;
  const anchor = latestPostedAnchor();
  if (anchor) {
    cur = stepMinute(anchor.h, anchor.m);                       // bottom = latest Posted + 1
  } else {
    const bottom = ordered[0];
    if (bottom.postedHour !== '' && bottom.postedHour != null) {
      cur = { h: parseInt(bottom.postedHour, 10) || 0, m: parseInt(bottom.postedMin || '0', 10) || 0 };
    } else {
      const n = new Date();
      cur = { h: n.getHours(), m: n.getMinutes() };             // seed at the current clock time
    }
  }

  const base = Date.now();
  const changed = [];
  for (let i = 0; i < ordered.length; i++) {
    if (i > 0) cur = stepMinute(cur.h, cur.m);
    const r = ordered[i];
    const hh = pad2(cur.h), mm = pad2(cur.m);
    if ((r.postedHour || '') !== hh || (r.postedMin || '') !== mm) {
      r.postedHour = hh;
      r.postedMin = mm;
      r.postedTimeAt = new Date(base + i).toISOString();
      changed.push(r);
    }
  }
  if (!changed.length) return [];

  try {
    const CHUNK = 400;
    for (let i = 0; i < changed.length; i += CHUNK) {
      const b = fdb.batch();
      changed.slice(i, i + CHUNK).forEach(r =>
        b.update(fdb.collection('inventory').doc(r.id), {
          postedHour: r.postedHour, postedMin: r.postedMin, postedTimeAt: r.postedTimeAt
        }));
      await b.commit();
    }
  } catch (e) {
    console.error('resequencePostingTimes:', e);
    showToast('Could not update all posting times: ' + e.message, 'error');
  }
  return changed.map(r => r.id);
}
function roleBadge(role) {
  const m = {admin:['rb-admin','Admin'],'semi-admin':['rb-semi','Semi-Admin'],viewer:['rb-viewer','Viewer']};
  const [cls,lbl] = m[role] || ['rb-viewer', role];
  return `<span class="role-badge ${cls}">${esc(lbl)}</span>`;
}
function dr(label, val) {
  return `<div class="dr"><span class="dl">${esc(label)}</span><span class="dv">${esc(val == null ? '—' : String(val))}</span></div>`;
}
function drHTML(label, valHTML) {
  return `<div class="dr"><span class="dl">${esc(label)}</span><span class="dv">${valHTML}</span></div>`;
}

// ── DOM CACHE ─────────────────────────────────────────
const EL = {};
function initEL() {
  ['invBody','tInfo','pgInfo','pgFirst','pgPrev','pgNext','pgLast','pgSize','selAll','selBar','selCount',
   'logBody','lInfo','lPgInfo','lPgPrev','lPgNext','lPgSize'].forEach(id => EL[id] = document.getElementById(id));
  EL.sTotal    = document.getElementById('s-total');
  EL.sActive   = document.getElementById('s-active');
  EL.sAvail    = document.getElementById('s-avail');
  EL.sReserved = document.getElementById('s-reserved');
  EL.dRecent   = document.getElementById('d-recent');
  EL.dStatus   = document.getElementById('d-status');
  EL.dClients  = document.getElementById('d-clients');
  EL.dProducts = document.getElementById('d-products');
}

// ── STATE ─────────────────────────────────────────────
let DB=[], LOGS=[], recentViewed=[];
let fd=[], fl=[];
let pg=1, sortCol=null, sortDir=1;
let lpg=1, lSortCol=null, lSortDir=1;
let curRec=null, editId=null, moreOpen=false, showDupes=false;
let _editUpdatedAt=null;
let currentUser=null, currentRole='viewer';
let USERS=[];
let SELECTIONS={clients:[],products:[],providers:[],routes:[]};
let persistentSelIds = new Set();
let pinnedIds = new Set();
let umEditUid=null, _secondApp=null;
let effDateTouched=false, actDateTouched=false, bulkEffDateTouched=false, bulkActDateTouched=false;
let _syncUnsub=null, _syncPrimed=false, _lastSyncAt='', _logCursor='', _syncRetry=0;
let _invLoading=false;   // guards against a live-sync ping racing an in-flight inventory load

function updateThemeButton() {
  const btn = document.getElementById('themeBtn');
  if (!btn) return;
  const dark = document.documentElement.hasAttribute('data-dark');
  const label = dark ? 'Switch to light mode' : 'Switch to dark mode';
  btn.title = label;
  btn.setAttribute('aria-label', label);
}

// ── THEME INIT (runs immediately on script load) ──────
(function() {
  const t = localStorage.getItem('cs-inv-theme');
  if (t === 'light') {
    document.documentElement.removeAttribute('data-dark');
  }
  updateThemeButton();
})();

// ── TOAST ─────────────────────────────────────────────
const TOAST_ICONS  = {success:'✔',error:'✖',info:'ℹ',warning:'⚠'};
const TOAST_TITLES = {success:'Success',error:'Error',info:'Info',warning:'Warning'};

function showToast(msg, type='success', duration=4000) {
  const c = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = `toast t-${type}`;
  t.innerHTML = `<span class="toast-icon">${TOAST_ICONS[type]||'•'}</span><div class="toast-body"><div class="toast-title">${TOAST_TITLES[type]||type}</div><div class="toast-msg">${esc(msg)}</div></div><button class="toast-close" onclick="dismissToast(this.closest('.toast'))">✕</button>`;
  c.appendChild(t);
  const timer = setTimeout(() => dismissToast(t), duration);
  t._timer = timer;
}

function showUndoToast(msg, onUndo, duration=6000, title='Deleted') {
  const c = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = 'toast t-info';
  t.innerHTML = `<span class="toast-icon">ℹ</span><div class="toast-body"><div class="toast-title">${esc(title)}</div><div class="toast-msg">${esc(msg)}</div></div><button class="toast-undo">↩ Undo</button><button class="toast-close" onclick="dismissToast(this.closest('.toast'))">✕</button>`;
  t.querySelector('.toast-undo').onclick = () => { clearTimeout(t._timer); dismissToast(t); onUndo(); };
  c.appendChild(t);
  const timer = setTimeout(() => dismissToast(t), duration);
  t._timer = timer;
}

function dismissToast(t) {
  if (!t || t._dismissed) return;
  t._dismissed = true;
  clearTimeout(t._timer);
  t.classList.add('hiding');
  setTimeout(() => t.remove(), 210);
}

// ── THEME ─────────────────────────────────────────────
function toggleTheme() {
  const h = document.documentElement;
  const dark = h.hasAttribute('data-dark');
  dark ? h.removeAttribute('data-dark') : h.setAttribute('data-dark','');
  updateThemeButton();
  localStorage.setItem('cs-inv-theme', dark ? 'light' : 'dark');
  setTimeout(drawChart, 40);
}

// ── AUTH ──────────────────────────────────────────────
fauth.onAuthStateChanged(async user => {
  if (user) {
    currentUser = user;
    await loadUserRole(user);
    await addLog('Login', `Signed in as ${user.email}`);
    document.getElementById('authOv').style.display = 'none';
    document.getElementById('appNav').style.display = '';
    document.getElementById('appMain').style.display = '';
    document.getElementById('navUser').textContent = user.email;
    applyRoleRestrictions();
    loadInventory();
    loadLogs();
    await loadSelections();
    if (currentRole === 'admin') loadUsers();
    startSyncListener();
    initAutoBackup();
  } else {
    stopSyncListener();
    if (_abTimer) clearInterval(_abTimer);
    _abTimer = null; clearTimeout(_abDeferT); AB = null;
    currentUser = null; currentRole = 'viewer';
    DB=[]; LOGS=[]; fd=[]; fl=[]; recentViewed=[];
    USERS=[]; SELECTIONS={clients:[],products:[],providers:[],routes:[]};
    persistentSelIds = new Set();
    document.getElementById('authOv').style.display = 'flex';
    document.getElementById('appNav').style.display = 'none';
    document.getElementById('appMain').style.display = 'none';
    document.getElementById('navUser').textContent = '—';
    renderDash(); renderTbl(); renderLogs();
  }
});

async function doSignIn() {
  const email = document.getElementById('authEmail').value.trim();
  const pass  = document.getElementById('authPass').value;
  const err   = document.getElementById('authErr');
  if (!email || !pass) { err.textContent = 'Enter email and password.'; return; }
  err.textContent = 'Signing in…';
  try { await fauth.signInWithEmailAndPassword(email, pass); }
  catch(e) { err.textContent = e.message; }
}
function doSignOut() { document.getElementById('soOv').classList.add('on'); }
function confirmSignOut() { document.getElementById('soOv').classList.remove('on'); fauth.signOut(); }

// ── LOCAL CACHE + DELTA SYNC ──────────────────────────────
// The inventory is large (thousands of docs), so re-reading the whole collection
// on every login / reload / Sync is the dominant Firebase read cost. Instead we keep
// a local IndexedDB copy and, after one cold load, fetch only the records that changed
// since last time (updatedAt > cursor). A cheap count() catches remote deletions; a
// daily full reload is the ultimate safety net. Everything degrades gracefully to a
// full read if IndexedDB or the delta path ever fails, so the app can't get stuck.
const IDB_NAME = 'cs-inv-cache', IDB_VER = 1;
const DELTA_REWIND_MS = 2 * 60 * 1000;        // re-fetch a 2-min overlap to tolerate clock skew
const FULL_REFRESH_MS = 24 * 60 * 60 * 1000;  // force a full reconcile at least once a day
let _idb = null;

function idbOpen() {
  if (_idb) return Promise.resolve(_idb);
  if (!('indexedDB' in window)) return Promise.reject(new Error('no indexedDB'));
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, IDB_VER);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('inventory')) db.createObjectStore('inventory', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv');
    };
    req.onsuccess = () => { _idb = req.result; resolve(_idb); };
    req.onerror = () => reject(req.error);
  });
}
function _idbDone(tx) { return new Promise((res, rej) => { tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); tx.onabort = () => rej(tx.error); }); }
function _idbReq(r)   { return new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); }

async function cacheGetAllInv() {
  try { const db = await idbOpen(); return (await _idbReq(db.transaction('inventory','readonly').objectStore('inventory').getAll())) || []; }
  catch(e) { return []; }
}
async function cacheReplaceInv(recs) {
  try { const db = await idbOpen(); const tx = db.transaction('inventory','readwrite'); const os = tx.objectStore('inventory');
        os.clear(); recs.forEach(r => os.put(r)); await _idbDone(tx); }
  catch(e) { console.error('cacheReplaceInv:', e); }
}
async function cachePutInv(recs) {
  if (!recs || !recs.length) return;
  try { const db = await idbOpen(); const tx = db.transaction('inventory','readwrite'); const os = tx.objectStore('inventory');
        recs.forEach(r => os.put(r)); await _idbDone(tx); }
  catch(e) { console.error('cachePutInv:', e); }
}
async function cacheDelInv(ids) {
  if (!ids || !ids.length) return;
  try { const db = await idbOpen(); const tx = db.transaction('inventory','readwrite'); const os = tx.objectStore('inventory');
        ids.forEach(id => os.delete(id)); await _idbDone(tx); }
  catch(e) { console.error('cacheDelInv:', e); }
}
async function kvGet(key) {
  try { const db = await idbOpen(); return await _idbReq(db.transaction('kv','readonly').objectStore('kv').get(key)); }
  catch(e) { return undefined; }
}
async function kvSet(key, val) {
  try { const db = await idbOpen(); const tx = db.transaction('kv','readwrite'); tx.objectStore('kv').put(val, key); await _idbDone(tx); }
  catch(e) { console.error('kvSet:', e); }
}
function maxUpdatedAt(arr) {
  let m = '';
  for (const r of arr) { const u = r.updatedAt || r.createdAt || ''; if (u > m) m = u; }
  return m;
}
function rewindIso(iso, ms) {
  const t = Date.parse(iso); if (isNaN(t)) return '';
  return new Date(t - ms).toISOString();
}
async function serverInvCount() {
  try {
    const col = fdb.collection('inventory');
    if (typeof col.count !== 'function') return null;   // aggregate count unsupported → caller falls back
    const agg = await col.count().get();
    return agg.data().count;
  } catch(e) { console.error('serverInvCount:', e); return null; }
}
// Mirror a local change into the cache AND notify other open clients (one call per write path).
function propagateChange(ids = [], del = [], full = false) {
  cachePutInv(ids.map(id => DB.find(r => r.id === id)).filter(Boolean));
  cacheDelInv(del);
  if (del.length) recordDeletions(del);   // persistent tombstone so closed clients catch the delete
  broadcastSync(ids, del, full);
}
// Deletions don't show up in an `updatedAt >` delta query, so we also append a small
// tombstone. A client that was closed when a record was deleted reads this on its next
// load and drops the record locally — no full re-read needed. (count() isn't available
// in this Firestore build, so this is how remote deletes reach reopened clients.)
async function recordDeletions(ids) {
  const clean = (ids || []).filter(Boolean);
  if (!clean.length) return;
  try {
    await fdb.collection('meta').doc('deletions').set({
      items: firebase.firestore.FieldValue.arrayUnion({ ids: clean, at: new Date().toISOString() })
    }, { merge: true });
  } catch(e) { console.error('recordDeletions:', e); }
}
async function applyRemoteDeletions(sinceIso) {
  try {
    const snap = await fdb.collection('meta').doc('deletions').get();
    if (!snap.exists) return;
    const items = snap.data().items || [];
    const goneIds = [];
    for (const it of items) if (it && it.at && it.at > sinceIso && Array.isArray(it.ids)) goneIds.push(...it.ids);
    if (goneIds.length) {
      const gone = new Set(goneIds);
      DB = DB.filter(r => !gone.has(r.id));
      goneIds.forEach(id => persistentSelIds.delete(id));
      await cacheDelInv(goneIds);
    }
    // Keep the tombstone doc bounded — trim oldest once it grows large (best effort).
    if (items.length > 1500) {
      const trimmed = items.slice().sort((a,b) => (a.at||'').localeCompare(b.at||'')).slice(-1000);
      await fdb.collection('meta').doc('deletions').set({ items: trimmed });
    }
  } catch(e) { console.error('applyRemoteDeletions:', e); }
}

// ── FIRESTORE LOAD ────────────────────────────────────
async function loadInventory() {
  _invLoading = true;
  try {
    const cached  = await cacheGetAllInv();
    const cursor  = await kvGet('invCursor');
    const lastFull= await kvGet('invFullAt');
    const stale   = !lastFull || (Date.now() - Date.parse(lastFull) > FULL_REFRESH_MS);
    if (!cached.length || !cursor || stale) { await fullLoadInventory(); return; }

    // Warm path: pull only records changed since last sync.
    DB = cached;
    const since = rewindIso(cursor, DELTA_REWIND_MS) || cursor;
    // Apply remote deletions FIRST, so an undo-restore in this same window (which arrives
    // as an add/update below) wins over its own tombstone.
    await applyRemoteDeletions(since);
    const snap  = await fdb.collection('inventory').where('updatedAt', '>', since).get();
    const changed = snap.docs.map(d => ({...d.data(), id:d.id}));
    if (changed.length) {
      changed.forEach(rec => { const i = DB.findIndex(r => r.id===rec.id); if (i>-1) DB[i]=rec; else DB.push(rec); });
      await cachePutInv(changed);
    }
    // Secondary net (only fires if this Firestore build ever gains count()): totals disagree → full reconcile.
    const cnt = await serverInvCount();
    if (cnt != null && cnt !== DB.length) { await fullLoadInventory(); return; }
    await kvSet('invCursor', maxUpdatedAt(DB) || cursor);
    refreshInventoryRecent();
  } catch(e) {
    console.error('loadInventory (delta):', e);
    try { await fullLoadInventory(); }                 // any delta failure → safe full reload
    catch(e2) { console.error('loadInventory (full fallback):', e2); if (DB.length) refreshInventoryRecent(); }
  } finally { _invLoading = false; }
}
async function fullLoadInventory() {
  const snap = await fdb.collection('inventory').orderBy('client').get();
  DB = snap.docs.map(d => ({...d.data(), id:d.id}));
  await cacheReplaceInv(DB);
  await kvSet('invCursor', maxUpdatedAt(DB));
  await kvSet('invFullAt', new Date().toISOString());
  refreshInventoryRecent();
}
function activityStamp(r) {
  return r?.updatedAt || r?.createdAt || '';
}
function loadPinned() {
  try { pinnedIds = new Set(JSON.parse(localStorage.getItem('cs-inv-pinned') || '[]')); } catch(e) { pinnedIds = new Set(); }
}
function savePinned() {
  localStorage.setItem('cs-inv-pinned', JSON.stringify([...pinnedIds]));
}
function sortInventoryByActivity() {
  const pinned = DB.filter(r => pinnedIds.has(r.id));
  const unpinned = DB.filter(r => !pinnedIds.has(r.id));
  pinned.sort((a,b) => activityStamp(b).localeCompare(activityStamp(a)) || String(b.id||'').localeCompare(String(a.id||'')));
  unpinned.sort((a,b) => activityStamp(b).localeCompare(activityStamp(a)) || String(b.id||'').localeCompare(String(a.id||'')));
  DB.length = 0;
  for (const r of [...pinned, ...unpinned]) DB.push(r);
}
function clearInventorySortState() {
  sortCol = null;
  sortDir = 1;
  document.querySelectorAll('#invTbl th').forEach(th => th.classList.remove('asc','desc'));
}
function refreshInventoryRecent(resetPage=true) {
  sortInventoryByActivity();
  clearInventorySortState();
  fd = [...DB];
  if (resetPage) pg = 1;
  renderTbl();
  renderDash();
}
async function syncData() {
  const btn = document.getElementById('syncBtn');
  btn.classList.add('syncing');
  // Refreshes THIS client only. Real changes already auto-broadcast to other open
  // users, so the button intentionally does NOT force everyone to full-reload — at
  // this inventory size that would cost ~8k reads per open user per click.
  try { await Promise.all([loadInventory(), loadLogs()]); }
  finally { btn.classList.remove('syncing'); }
}
async function loadLogs() {
  try {
    const cachedLogs = await kvGet('logsCache');
    const cursor     = await kvGet('logsCursor');
    const lastFull   = await kvGet('logsFullAt');
    const stale      = !lastFull || (Date.now() - Date.parse(lastFull) > FULL_REFRESH_MS);
    if (!cachedLogs || !cachedLogs.length || !cursor || stale) {
      const snap = await fdb.collection('logs').orderBy('datetime','desc').limit(500).get();
      LOGS = snap.docs.map(d => ({...d.data(), id:d.id}));
      await kvSet('logsFullAt', new Date().toISOString());
    } else {
      LOGS = cachedLogs;
      const snap = await fdb.collection('logs').where('datetime','>', cursor).orderBy('datetime','desc').get();
      const fresh = snap.docs.map(d => ({...d.data(), id:d.id}));
      if (fresh.length) {
        const known = new Set(LOGS.map(l => l.id));
        const add = fresh.filter(l => !known.has(l.id));
        if (add.length) LOGS = [...add, ...LOGS].slice(0, 500);
      }
    }
    _logCursor = LOGS[0]?.datetime || '';
    kvSet('logsCache', LOGS); kvSet('logsCursor', _logCursor);
    fl = [...LOGS]; renderLogs();
  } catch(e) {
    console.error('loadLogs:', e);
    if (!LOGS.length) {
      try { const snap = await fdb.collection('logs').orderBy('datetime','desc').limit(500).get();
            LOGS = snap.docs.map(d => ({...d.data(), id:d.id})); _logCursor = LOGS[0]?.datetime||''; fl=[...LOGS]; renderLogs(); }
      catch(e2) { console.error('loadLogs (fallback):', e2); }
    }
  }
}

// ── LIVE SYNC (lightweight cross-client refresh) ──────────
// Every open client watches ONE tiny doc (meta/syncSignal). When someone
// changes inventory they write the affected record ids here; other clients
// fetch just those docs and patch them in place. This is deliberately NOT a
// real-time listener on the whole inventory collection — only this single doc
// is watched, so idle cost is ~zero and each change costs other clients only
// the reads for the records that actually changed.
// Adds/updates beyond SYNC_LIMIT → full reload (each changed row costs a read to fetch,
// so past a few hundred it's cheaper to just reload all). Deletes are free for receivers
// to apply (the ids ride inside the signal), so they get a much higher cap — bounded only
// by the signal doc's size, not by read cost.
const SYNC_LIMIT = 300;
const DEL_LIMIT  = 1000;
// If the ONE watched doc's listener dies by ERROR (not a normal network blip — the SDK
// auto-recovers those for free), re-arm it with bounded exponential backoff so live sync
// heals itself instead of silently staying dead until someone reloads. Terminal auth errors
// are NOT retried (they'd loop forever, burning reads); a healthy re-attach resets the count.
const SYNC_RETRY_MAX     = 5;
const SYNC_RETRY_BASE_MS = 2000;
async function broadcastSync(ids = [], del = [], full = false) {
  if (!currentUser) return;
  const tooBig = ids.length > SYNC_LIMIT || del.length > DEL_LIMIT;
  try {
    await fdb.collection('meta').doc('syncSignal').set({
      at:   new Date().toISOString(),
      by:   currentUser.uid || '',
      ids:  (full || tooBig) ? [] : ids.filter(Boolean),
      del:  (full || tooBig) ? [] : del.filter(Boolean),
      full: full || tooBig
    });
  } catch(e) { console.error('broadcastSync:', e); }
}
function startSyncListener() {
  if (_syncUnsub) return;
  _syncPrimed = false;
  _syncUnsub = fdb.collection('meta').doc('syncSignal').onSnapshot(snap => {
    const sig = snap.data();
    // First callback is the doc's current state at attach — baseline only; a healthy attach
    // also means we're connected, so clear any backoff left over from a prior failure.
    if (!_syncPrimed) { _syncPrimed = true; _syncRetry = 0; _lastSyncAt = sig?.at || ''; return; }
    if (!sig || !sig.at || sig.at === _lastSyncAt) return;
    _lastSyncAt = sig.at;
    if (sig.by && sig.by === currentUser?.uid) return;   // our own change, already applied locally
    applyRemoteSync(sig);
  }, err => {
    console.error('sync listener:', err);
    // The listener is dead now. Clear the handle so startSyncListener() can re-arm — its
    // `if (_syncUnsub) return` guard would otherwise refuse forever. Then retry with backoff,
    // unless the error is terminal (auth) or we've signed out / exhausted attempts.
    _syncUnsub = null; _syncPrimed = false;
    if (err?.code === 'permission-denied' || err?.code === 'unauthenticated') return;
    if (!currentUser || _syncRetry >= SYNC_RETRY_MAX) return;
    const delay = Math.min(SYNC_RETRY_BASE_MS * 2 ** _syncRetry, 30000);
    _syncRetry++;
    setTimeout(() => { if (currentUser && !_syncUnsub) startSyncListener(); }, delay);
  });
}
function stopSyncListener() {
  if (_syncUnsub) { _syncUnsub(); _syncUnsub = null; }
  _syncPrimed = false; _lastSyncAt = ''; _syncRetry = 0;
}
function isInvFilterActive() {
  return ['fSearch','fClient','fStatus','fProduct','fProvider','fDateFrom','fDateTo']
    .some(id => document.getElementById(id)?.value) || showDupes;
}
async function applyRemoteSync(sig) {
  if (_invLoading) return;   // an initial/explicit load is in flight and will capture this change
  const btn = document.getElementById('syncBtn');
  btn?.classList.add('syncing');
  try {
    if (sig.full) {
      const keepPg = pg;
      await loadInventory();            // delta reload (picks up the changed rows cheaply)
      await refreshLogsIncremental();   // pull only NEW logs, not a full 500-doc re-read
      pg = keepPg; remoteRerender();    // keep the viewer's page/search instead of jumping to page 1
      return;
    }
    // Deletes: ids ride in the signal, so no reads needed to drop them locally.
    const goneIds = [...(sig.del || [])];
    if (sig.del?.length) {
      const gone = new Set(sig.del);
      DB = DB.filter(r => !gone.has(r.id));
      sig.del.forEach(id => persistentSelIds.delete(id));
    }
    // Adds/updates: fetch just the affected records (documentId 'in' → chunk by 10, in parallel).
    const ids = (sig.ids || []).filter(Boolean);
    const chunks = [];
    for (let i = 0; i < ids.length; i += 10) chunks.push(ids.slice(i, i + 10));
    const results = await Promise.all(chunks.map(chunk =>
      fdb.collection('inventory').where(firebase.firestore.FieldPath.documentId(), 'in', chunk).get()
        .then(qs => ({ chunk, qs }))
    ));
    const fetched = [];
    for (const { chunk, qs } of results) {
      const found = new Set();
      qs.forEach(d => {
        found.add(d.id);
        const rec = {...d.data(), id:d.id};
        fetched.push(rec);
        const idx = DB.findIndex(r => r.id === d.id);
        if (idx > -1) DB[idx] = rec; else DB.push(rec);
      });
      // Requested but not returned → deleted after the ping; drop it locally.
      chunk.forEach(id => { if (!found.has(id)) { goneIds.push(id); DB = DB.filter(r => r.id !== id); persistentSelIds.delete(id); } });
    }
    cachePutInv(fetched);            // keep the local cache in step with the remote change
    cacheDelInv(goneIds);
    remoteRerender();
    await refreshLogsIncremental();
  } catch(e) { console.error('applyRemoteSync:', e); }
  finally { btn?.classList.remove('syncing'); }
}
// Re-render after a remote patch without yanking the viewer around: keep their
// active search/filter and their current page (clamped) instead of resetting.
function remoteRerender() {
  const keepPg = pg;
  if (isInvFilterActive()) applyF();          // rebuilds fd from current filters (sets pg=1)
  else refreshInventoryRecent(false);         // no filter: fd=DB, keep page
  const sz = parseInt(EL?.pgSize?.value || 50);
  const tp = Math.max(1, Math.ceil(fd.length / sz));
  pg = Math.min(Math.max(1, keepPg), tp);
  renderTbl();
  renderDash();
}
async function refreshLogsIncremental() {
  try {
    if (!_logCursor) { _logCursor = LOGS[0]?.datetime || ''; return; }
    const snap = await fdb.collection('logs')
      .where('datetime', '>', _logCursor).orderBy('datetime', 'desc').get();
    if (snap.empty) return;
    const known = new Set(LOGS.map(l => l.id));
    const add = snap.docs.map(d => ({...d.data(), id:d.id})).filter(l => !known.has(l.id));
    if (!add.length) return;
    LOGS = [...add, ...LOGS].slice(0, 500);
    _logCursor = LOGS[0]?.datetime || _logCursor;
    kvSet('logsCache', LOGS); kvSet('logsCursor', _logCursor);
    fl = [...LOGS]; renderLogs();
  } catch(e) { console.error('refreshLogsIncremental:', e); }
}

// ── NAVIGATION ────────────────────────────────────────
function go(tab, btn) {
  if (tab==='admin' && currentRole!=='admin') return;
  if (tab==='logs'  && currentRole==='viewer') return;
  document.querySelectorAll('.page').forEach(el => el.classList.remove('on'));
  document.querySelectorAll('.nav-btn').forEach(el => el.classList.remove('on'));
  document.getElementById('page-'+tab).classList.add('on');
  btn.classList.add('on');
  if (tab==='dashboard') renderDash();
  if (tab==='inventory') renderTbl();
  if (tab==='logs')      renderLogs();
  if (tab==='admin')     { loadUsers(); renderAutoBackupCard(); }
}

// ── DASHBOARD ─────────────────────────────────────────
function renderDash() {
  const counts = {}; STATUSES.forEach(s => counts[s]=0);
  DB.forEach(r => counts[r.status] = (counts[r.status]||0)+1);
  if (EL.sTotal)    EL.sTotal.textContent    = DB.length.toLocaleString();
  if (EL.sActive)   EL.sActive.textContent   = (counts['Active']||0).toLocaleString();
  if (EL.sAvail)    EL.sAvail.textContent    = (counts['Available']||0).toLocaleString();
  if (EL.sReserved) EL.sReserved.textContent = (counts['Reserved']||0).toLocaleString();

  if (EL.dRecent) {
    if (!recentViewed.length) {
      EL.dRecent.innerHTML = '<p style="color:var(--t3);font-size:12px">No recently viewed numbers.</p>';
    } else {
      EL.dRecent.innerHTML = recentViewed.slice(0,6).map(r => `
        <div class="li" onclick="openSP('${esc(r.id)}')" style="cursor:pointer">
          <div class="li-left"><div class="li-name">${esc(r.number)}</div><div class="li-sub">${esc(r.client)} · ${esc(r.product)}</div></div>
          <span class="badge ${bclass(r.status)}">${esc(r.status)}</span>
        </div>`).join('');
    }
  }

  const total = DB.length || 1;
  const colors = {Active:'#4f8ef7',Available:'#34d399',Reserved:'#fbbf24',Inactive:'#f87171'};
  if (EL.dStatus) {
    EL.dStatus.innerHTML = Object.entries(counts).map(([s,c]) => `
      <div class="sbar">
        <div class="sbar-row"><span>${esc(s)}</span><span>${c} (${Math.round(c/total*100)}%)</span></div>
        <div class="sbar-track"><div class="sbar-fill" style="width:${c/total*100}%;background:${colors[s]}"></div></div>
      </div>`).join('');
  }

  if (EL.dClients) {
    const cc = {}; DB.forEach(r => cc[r.client] = (cc[r.client]||0)+1);
    EL.dClients.innerHTML = Object.entries(cc).sort((a,b) => b[1]-a[1]).slice(0,5)
      .map(([c,n]) => `<div class="li"><span class="li-name">${esc(c)}</span><span style="color:var(--t2);font-size:12px">${n} numbers</span></div>`).join('');
  }

  if (EL.dProducts) {
    const pc = {}; DB.forEach(r => pc[r.product] = (pc[r.product]||0)+1);
    EL.dProducts.innerHTML = Object.entries(pc).sort((a,b) => b[1]-a[1]).slice(0,5)
      .map(([p,n]) => `<div class="li"><span class="li-name">${esc(p)}</span><span style="color:var(--t2);font-size:12px">${n} numbers</span></div>`).join('');
  }

  setTimeout(drawChart, 60);
}

function drawChart() {
  const canvas = document.getElementById('chart');
  if (!canvas) return;
  const days = parseInt(document.getElementById('chartDays')?.value || '14');
  const rect  = canvas.parentElement.getBoundingClientRect();
  const dpr   = window.devicePixelRatio || 1;
  canvas.width  = rect.width  * dpr;
  canvas.height = rect.height * dpr;
  const ctx = canvas.getContext('2d'); ctx.scale(dpr, dpr);
  const W = rect.width, H = rect.height;
  const dark = document.documentElement.hasAttribute('data-dark');
  const tc = dark ? '#64748b' : '#9ca3af';
  const gc = dark ? '#2d3f52' : '#e5e7eb';
  const labels=[], act=[], deact=[];
  const now = Date.now();
  for (let i = days-1; i >= 0; i--) {
    const d  = new Date(now - i*864e5);
    const ds = d.toISOString().split('T')[0];
    labels.push(d.toLocaleDateString('en-US',{month:'short',day:'numeric'}));
    act.push(DB.filter(r => r.actDate===ds).length);
    deact.push(DB.filter(r => r.deactDate===ds).length);
  }
  const pad = {l:32,r:10,t:8,b:26};
  const cW = W-pad.l-pad.r, cH = H-pad.t-pad.b;
  const maxV = Math.max(...act,...deact,4)+2;
  const bW   = (cW/days)*0.33;
  const labelEvery = Math.max(1, Math.round(days/7));
  ctx.clearRect(0,0,W,H);
  for (let i=0; i<=4; i++) {
    const y = pad.t+(cH/4)*i;
    ctx.strokeStyle=gc; ctx.lineWidth=1;
    ctx.beginPath(); ctx.moveTo(pad.l,y); ctx.lineTo(W-pad.r,y); ctx.stroke();
    ctx.fillStyle=tc; ctx.font=`9px 'DM Sans',system-ui`; ctx.textAlign='right';
    ctx.fillText(Math.round(maxV-(maxV/4)*i), pad.l-3, y+3);
  }
  labels.forEach((day,i) => {
    const x  = pad.l+(cW/days)*i+(cW/days)*0.1;
    const aH = act[i]/maxV*cH, dH = deact[i]/maxV*cH;
    ctx.fillStyle='#4f8ef7'; ctx.fillRect(x, pad.t+cH-aH, bW, aH);
    ctx.fillStyle='#f87171'; ctx.fillRect(x+bW+2, pad.t+cH-dH, bW, dH);
    if (i % labelEvery === 0) {
      ctx.fillStyle=tc; ctx.font=`8px 'DM Sans',system-ui`; ctx.textAlign='center';
      ctx.fillText(day, x+bW, H-5);
    }
  });
}

// ── FILTERS ───────────────────────────────────────────
function wildcardToRegex(s) {
  if (!s.includes('*')) return new RegExp(s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'i');
  const esc2 = s.split('*').map(p => p.replace(/[.+?^${}()|[\]\\]/g,'\\$&')).join('.*');
  const anchored = (s[0]!=='*' ? '^' : '') + esc2 + (s[s.length-1]!=='*' ? '$' : '');
  return new RegExp(anchored, 'i');
}
function getPhoneNorms(numberField) {
  if (!numberField || String(numberField).trim().toUpperCase() === 'NA') return [];
  return String(numberField).split('/').map(p => normalizePhone(p.trim())).filter(Boolean);
}
function getDupeSet() {
  const counts = {};
  DB.forEach(r => {
    getPhoneNorms(r.number).forEach(key => {
      counts[key] = (counts[key] || 0) + 1;
    });
  });
  const dupeKeys = new Set(Object.keys(counts).filter(k => counts[k] > 1));
  return dupeKeys;
}
function applyF() {
  const s  = document.getElementById('fSearch').value.toLowerCase();
  const cl = document.getElementById('fClient').value;
  const st = document.getElementById('fStatus').value;
  const pr = document.getElementById('fProduct').value;
  const pv = document.getElementById('fProvider').value;
  const df = document.getElementById('fDateFrom').value;
  const dt = document.getElementById('fDateTo').value;
  const sRe = s ? wildcardToRegex(s) : null;
  const dupeSet = showDupes ? getDupeSet() : null;
  fd = DB.filter(r => {
    if (sRe) {
      const SKIP = new Set(['id','createdBy','updatedBy','createdAt','updatedAt','clientOSF','clientMRC','clientOTRF','clientCF','clientCPM','prevClient']);
      // Match against the canonical Posted Status label shown in the table (e.g. legacy
      // "Not Yet"/"Yes" display as "For Posting"/"Posted"), not just the raw stored value.
      const hit = Object.entries(r).some(([k,v]) => !SKIP.has(k) && v != null && sRe.test(String(v).toLowerCase()))
                || sRe.test(canonPostedStatus(r.postedStatus).toLowerCase());
      if (!hit) return false;
    }
    if (cl && r.client!==cl)   return false;
    if (st && r.status!==st)   return false;
    if (pr && r.product!==pr)  return false;
    if (pv && r.provider!==pv) return false;
    if (df && (!r.actDate || r.actDate < df)) return false;
    if (dt && (!r.actDate || r.actDate > dt)) return false;
    if (dupeSet) {
      const norms = getPhoneNorms(r.number);
      if (!norms.length || !norms.some(k => dupeSet.has(k))) return false;
    }
    return true;
  });
  pg=1; renderTbl();
}
function clearF() {
  ['fSearch','fDateFrom','fDateTo'].forEach(id => document.getElementById(id).value='');
  ['fClient','fStatus','fProduct','fProvider'].forEach(id => document.getElementById(id).value='');
  if (showDupes) {
    showDupes = false;
    document.getElementById('btnDupes').classList.remove('active');
  }
  sortCol=null; sortDir=1;
  document.querySelectorAll('#invTbl th').forEach(th => th.classList.remove('asc','desc'));
  fd=[...DB]; pg=1; renderTbl();
}
function toggleDupes() {
  showDupes = !showDupes;
  document.getElementById('btnDupes').classList.toggle('active', showDupes);
  if (showDupes) {
    document.getElementById('fDateFrom').value = '';
    document.getElementById('fDateTo').value = '';
  }
  applyF();
}
function toggleMore() {
  moreOpen = !moreOpen;
  document.getElementById('moreRow').classList.toggle('on', moreOpen);
  document.getElementById('moreBtn').textContent = moreOpen ? 'Less ▴' : 'More ▾';
}

// ── SORT ──────────────────────────────────────────────
const colIdx = {client:2,product:3,number:4,status:5,postedStatus:6,remarks:7};
function sortBy(col) {
  if (sortCol===col) sortDir*=-1; else { sortCol=col; sortDir=1; }
  const pinnedFd = fd.filter(r => pinnedIds.has(r.id));
  const unpinnedFd = fd.filter(r => !pinnedIds.has(r.id));
  unpinnedFd.sort((a,b) => (a[col]||'').localeCompare(b[col]||'')*sortDir);
  fd = [...pinnedFd, ...unpinnedFd];
  document.querySelectorAll('#invTbl th').forEach(th => th.classList.remove('asc','desc'));
  const ths = [...document.querySelectorAll('#invTbl th')];
  if (colIdx[col]) ths[colIdx[col]].classList.add(sortDir===1?'asc':'desc');
  renderTbl();
}

// ── RENDER TABLE ──────────────────────────────────────
function renderTbl() {
  const sz = parseInt(EL.pgSize?.value || 50);
  const s = (pg-1)*sz, e = s+sz, total = fd.length, tp = Math.ceil(total/sz)||1;
  if (EL.tInfo)    EL.tInfo.textContent    = `Showing ${Math.min(s+1,total)}–${Math.min(e,total)} of ${total} records`;
  if (EL.pgInfo)   EL.pgInfo.textContent   = `Page ${pg} of ${tp}`;
  if (EL.pgFirst)  EL.pgFirst.disabled     = pg<=1;
  if (EL.pgPrev)   EL.pgPrev.disabled      = pg<=1;
  if (EL.pgNext)   EL.pgNext.disabled      = pg>=tp;
  if (EL.pgLast)   EL.pgLast.disabled      = pg>=tp;
  if (EL.invBody)  EL.invBody.innerHTML    = fd.slice(s,e).map((r,i) => {
    const isPinned = pinnedIds.has(r.id);
    return `
    <tr style="--row-i:${i}" class="${isPinned?'tr-pinned':''}" onclick="rowClick(event,'${esc(r.id)}')">
      <td onclick="event.stopPropagation()"><input type="checkbox" class="rcb" data-id="${esc(r.id)}" ${persistentSelIds.has(r.id)?'checked':''} onchange="toggleRowSel(this)"></td>
      <td class="row-num">${isPinned?'<span class="pin-ind" title="Pinned">📌</span>':s+i+1}</td>
      <td>${esc(r.client)}</td>
      <td>${esc(r.product)}</td>
      <td class="num-cell"><span class="num-val">${esc(r.number)}</span><button type="button" class="num-copy" title="Copy number" aria-label="Copy number" onclick="event.stopPropagation();copyNumber(this)"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg></button></td>
      <td><span class="badge ${bclass(r.status)}">${esc(r.status)}</span></td>
      <td><span class="badge ${postedClass(r.postedStatus)}">${esc(canonPostedStatus(r.postedStatus) || 'No')}</span>${r.postedHour ? `<span class="posted-time">${esc(r.postedHour)}:${esc(r.postedMin || '00')}</span>` : ''}</td>
      <td>${esc(r.remarks)}</td>
      <td onclick="event.stopPropagation()">
        <div class="act-btns">
          <button class="act-btn pin-btn${isPinned?' pinned':''}" title="${isPinned?'Unpin this entry':'Pin this entry'}" onclick="togglePin('${esc(r.id)}')">📌</button>
          ${currentRole!=='viewer'?`<button class="act-btn" title="Edit" onclick="openEditById('${esc(r.id)}')">✎</button><button class="act-btn del" title="Delete" onclick="delRec('${esc(r.id)}')">⊗</button>`:''}
        </div>
      </td>
    </tr>`;
  }).join('');
  updateSelBar();
}
// Copy a single phone number from its table cell to the clipboard.
function copyNumber(btn) {
  const cell = btn.closest('.num-cell');
  const val = cell ? (cell.querySelector('.num-val')?.textContent || '').trim() : '';
  if (!val) { showToast('No number to copy.', 'warning'); return; }
  const done = () => showToast('Number copied to clipboard.', 'success');
  const fail = () => showToast('Could not copy to clipboard.', 'error');
  if (navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(val).then(done).catch(fail);
  } else {
    try {
      const ta = document.createElement('textarea');
      ta.value = val; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      document.execCommand('copy'); ta.remove(); done();
    } catch(e) { fail(); }
  }
}
function changePg(d) {
  const sz = parseInt(EL.pgSize?.value || 50);
  const tp = Math.ceil(fd.length/sz)||1;
  pg = Math.max(1, Math.min(pg+d, tp)); renderTbl();
}
function goToPage(target) {
  const sz = parseInt(EL.pgSize?.value || 50);
  const tp = Math.ceil(fd.length/sz)||1;
  pg = target === 'first' ? 1 : tp; renderTbl();
}
function selAllRows(cb) {
  if (cb.checked) {
    document.querySelectorAll('.rcb').forEach(c => { c.checked=true; persistentSelIds.add(c.dataset.id); });
  } else {
    persistentSelIds.clear();
    document.querySelectorAll('.rcb').forEach(c => c.checked=false);
  }
  updateSelBar();
}
function toggleRowSel(cb) {
  if (cb.checked) persistentSelIds.add(cb.dataset.id);
  else persistentSelIds.delete(cb.dataset.id);
  updateSelBar();
}
function getCheckedIds() { return [...persistentSelIds]; }

function togglePin(id) {
  if (pinnedIds.has(id)) pinnedIds.delete(id);
  else pinnedIds.add(id);
  savePinned();
  refreshInventoryRecent(false);
  updatePinBtnState(editId);
  updateSPPinBtn();
}
function pinEntries(ids) {
  ids.forEach(id => pinnedIds.add(id));
  savePinned();
  refreshInventoryRecent(false);
}
function unpinEntries(ids) {
  ids.forEach(id => pinnedIds.delete(id));
  savePinned();
  refreshInventoryRecent(false);
}
function updatePinBtnState(id) {
  const btn = document.getElementById('mPinBtn');
  if (!btn) return;
  if (!id) { btn.style.display = 'none'; return; }
  btn.style.display = '';
  btn.textContent = pinnedIds.has(id) ? '📌 Unpin' : '📌 Pin';
}
function togglePinModal() {
  if (!editId) return;
  togglePin(editId);
}
function togglePinSP() {
  if (!curRec) return;
  togglePin(curRec.id);
}
function updateSPPinBtn() {
  const btn = document.getElementById('btnSpPin');
  if (!btn || !curRec) return;
  btn.textContent = pinnedIds.has(curRec.id) ? '📌 Unpin' : '📌 Pin';
}
function pinSelected() {
  const ids = getCheckedIds(); if (!ids.length) return;
  pinEntries(ids);
  showToast(`Pinned ${ids.length} record${ids.length!==1?'s':''}`, 'success');
}
function unpinSelected() {
  const ids = getCheckedIds(); if (!ids.length) return;
  unpinEntries(ids);
  showToast(`Unpinned ${ids.length} record${ids.length!==1?'s':''}`, 'success');
}
function updateSelBar() {
  const count = persistentSelIds.size;
  const total = document.querySelectorAll('.rcb').length;
  const checkedOnPage = document.querySelectorAll('.rcb:checked').length;
  if (EL.selCount) EL.selCount.textContent = `${count} row${count!==1?'s':''} selected`;
  if (EL.selBar)   EL.selBar.classList.toggle('on', count>0);
  if (EL.selAll) {
    EL.selAll.indeterminate = checkedOnPage>0 && checkedOnPage<total;
    EL.selAll.checked = total>0 && checkedOnPage===total;
  }
}
function rowClick(e, id) {
  if (e.target.tagName==='INPUT') return;
  openSP(id);
  const r = DB.find(x => x.id===id);
  if (r && !recentViewed.find(x => x.id===id)) { recentViewed.unshift(r); recentViewed=recentViewed.slice(0,6); }
}

// ── SIDE PANEL ────────────────────────────────────────
function activationSnapshot(r={}) {
  return {
    client: r.client || '',
    product: r.product || '',
    status: r.status || '',
    effDate: r.effDate || '',
    actDate: r.actDate || '',
    provider: r.provider || '',
    arrDate: r.arrDate || '',
    provActDate: r.provActDate || '',
    route: r.route || ''
  };
}
function hasActivationSnapshot(a={}) {
  return !!(a.client || a.product || a.effDate || a.actDate || a.provider || a.arrDate || a.provActDate || a.route);
}
function activationRowsHTML(a={}) {
  if (!hasActivationSnapshot(a)) return '';
  return `
    ${a.product ? `<div class="deact-hist-row"><span style="color:var(--t3)">Product</span> ${esc(a.product)}</div>` : ''}
    ${a.status ? `<div class="deact-hist-row"><span style="color:var(--t3)">Status</span> ${esc(a.status)}</div>` : ''}
    ${a.effDate ? `<div class="deact-hist-row"><span style="color:var(--t3)">Effective Date</span> ${fmt(a.effDate)}</div>` : ''}
    ${a.actDate ? `<div class="deact-hist-row"><span style="color:var(--t3)">Activated Date</span> ${fmt(a.actDate)}</div>` : ''}
    ${a.provider ? `<div class="deact-hist-row"><span style="color:var(--t3)">Provider</span> ${esc(a.provider)}</div>` : ''}
    ${a.arrDate ? `<div class="deact-hist-row"><span style="color:var(--t3)">Arrival Date</span> ${fmt(a.arrDate)}</div>` : ''}
    ${a.provActDate ? `<div class="deact-hist-row"><span style="color:var(--t3)">Provider Activation</span> ${fmt(a.provActDate)}</div>` : ''}
    ${a.route ? `<div class="deact-hist-row"><span style="color:var(--t3)">Route Request by</span> ${esc(a.route)}</div>` : ''}`;
}
function metaDate(v) {
  return v ? new Date(v).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}) : '—';
}
function currentActivationHTML(r) {
  const a = activationSnapshot(r);
  if (!hasActivationSnapshot(a)) return '';
  return `
    <div class="deact-hist-entry act-hist-entry">
      <div class="deact-hist-top">
        <span class="deact-hist-client">${esc(a.client || 'Activation Details')}</span>
        <span class="deact-hist-date">${fmt(a.actDate || a.effDate)}</span>
      </div>
      <div class="hist-subtitle">Current Activation</div>
      ${activationRowsHTML(a)}
      <div class="deact-hist-meta">last updated by ${esc(r.updatedBy || r.createdBy || '—')} · ${metaDate(r.updatedAt || r.createdAt)}</div>
    </div>`;
}
function historySectionHTML(r) {
  const current = currentActivationHTML(r);
  const deact = (r.deactivationHistory?.length) ? [...r.deactivationHistory].reverse().map(h => {
    const activation = h.activation || {};
    return `
      <div class="deact-hist-entry">
        <div class="deact-hist-top">
          <span class="deact-hist-client">${esc(h.previousClient||'—')}</span>
          <span class="deact-hist-date">${fmt(h.deactDate)}</span>
        </div>
        ${hasActivationSnapshot(activation) ? `<div class="hist-subtitle">Activation Details</div>${activationRowsHTML(activation)}<div class="hist-subtitle">Deactivation Details</div>` : ''}
        ${h.requestedBy ? `<div class="deact-hist-row"><span style="color:var(--t3)">Requested by</span> ${esc(h.requestedBy)}</div>` : ''}
        ${h.remarks ? `<div class="deact-hist-row">${esc(h.remarks)}</div>` : ''}
        <div class="deact-hist-meta">by ${esc(h.deactivatedBy||'—')} · ${metaDate(h.deactivatedAt)}</div>
      </div>`;
  }).join('') : '';
  return (current || deact) ? `<div class="ds"><div class="ds-title">Activation &amp; Deactivation History</div>${current}${deact}</div>` : '';
}
function openSP(id) {
  const r = DB.find(x => x.id===id); if (!r) return;
  curRec = r;
  document.getElementById('spTitle').textContent = r.number;
  document.getElementById('spBody').innerHTML = `
    <div class="ds"><div class="ds-title">Client Information</div>
      ${dr('Client',r.client)}${dr('Product',r.product)}${dr('Number',r.number)}
      ${drHTML('Status',`<span class="badge ${bclass(r.status)}">${esc(r.status)}</span>`)}
      ${dr('Remarks',r.remarks||'—')}${dr('Posted Status',canonPostedStatus(r.postedStatus)||'—')}
      ${dr('Posted Date & Time', (r.postedDate || r.postedHour) ? `${r.postedDate ? fmt(r.postedDate) : ''}${r.postedHour ? ` ${r.postedHour}:${r.postedMin||'00'}` : ''}`.trim() : '—')}${dr('Client OSF','$'+(r.clientOSF||'—'))}
      ${dr('Client MRC','$'+(r.clientMRC||'—'))}${dr('Client OTRF','$'+(r.clientOTRF||'—'))}
      ${dr('Client Channel Fee','$'+(r.clientCF||'—'))}${dr('Client CPM',r.clientCPM||'—')}
      ${dr('Effective Date',fmt(r.effDate))}${dr('Activated Date',fmt(r.actDate))}
    </div>
    <div class="ds"><div class="ds-title">Provider Information</div>
      ${dr('Provider',r.provider||'—')}${dr('Arrival Date',fmt(r.arrDate))}
      ${dr('Provider Activation Date',fmt(r.provActDate))}
      ${dr('Provider OSF','$'+(r.provOSF||'—'))}${dr('Provider MRC','$'+(r.provMRC||'—'))}
      ${dr('Provider OTRF','$'+(r.provOTRF||'—'))}${dr('Provider CPM',r.provCPM||'—')}
      ${dr('Type / Session',r.typeSession||'—')}
    </div>
    <div class="ds"><div class="ds-title">Routing &amp; History</div>
      ${dr('Route Request by',r.route||'—')}
      ${dr('Deactivation Date (Prev Client)',fmt(r.deactDate))}
      ${dr('Previous Client',r.prevClient||'—')}
    </div>
    ${historySectionHTML(r)}
    <div class="ds"><div class="ds-title">Meta</div>
      ${dr('Created by',r.createdBy||'—')}${dr('Updated by',r.updatedBy||'—')}
    </div>`;
  document.getElementById('spOv').classList.add('on');
  document.getElementById('sp').classList.add('on');
  updateSPPinBtn();
}
function closeSP() {
  document.getElementById('spOv').classList.remove('on');
  document.getElementById('sp').classList.remove('on');
  curRec = null;
}

// ── MODAL ─────────────────────────────────────────────
const mMap = {
  mClient:'client',mProduct:'product',mNumber:'number',mStatus:'status',mRemarks:'remarks',
  mPosted:'postedStatus',mPostedDate:'postedDate',mPostedHour:'postedHour',mPostedMin:'postedMin',
  mClientOSF:'clientOSF',mClientMRC:'clientMRC',
  mClientOTRF:'clientOTRF',mClientCF:'clientCF',mClientCPM:'clientCPM',mEffDate:'effDate',
  mActDate:'actDate',mProvider:'provider',mArrDate:'arrDate',mProvActDate:'provActDate',
  mProvOSF:'provOSF',mProvMRC:'provMRC',mProvOTRF:'provOTRF',mProvCPM:'provCPM',
  mTypeSession:'typeSession',mRoute:'route',mDeactDate:'deactDate',mPrevClient:'prevClient'
};
const FEE_FIELDS = [
  ['mClientOSFSel','mClientOSF'],['mClientMRCSel','mClientMRC'],
  ['mClientOTRFSel','mClientOTRF'],['mClientCPMSel','mClientCPM'],
  ['mProvOSFSel','mProvOSF'],['mProvMRCSel','mProvMRC'],
  ['mProvOTRFSel','mProvOTRF'],['mProvCPMSel','mProvCPM']
];
const BE_FEE_FIELDS = [
  ['beClientOSFSel','beClientOSF'],['beClientMRCSel','beClientMRC'],
  ['beClientOTRFSel','beClientOTRF'],['beClientCPMSel','beClientCPM'],
  ['beProvOSFSel','beProvOSF'],['beProvMRCSel','beProvMRC'],
  ['beProvOTRFSel','beProvOTRF'],['beProvCPMSel','beProvCPM']
];
function bindDateMirror(effId, actId, isActTouched, setEffTouched, setActTouched) {
  const eff = document.getElementById(effId);
  const act = document.getElementById(actId);
  if (!eff || !act || eff.dataset.mirrorBound) return;
  const mirror = () => {
    setEffTouched(true);
    if (!isActTouched()) act.value = eff.value;
  };
  eff.addEventListener('input', mirror);
  eff.addEventListener('change', mirror);
  act.addEventListener('input', () => setActTouched(true));
  act.addEventListener('change', () => setActTouched(true));
  eff.dataset.mirrorBound = '1';
}
function fillTimeSelects(hourId, minId) {
  const hourSel = document.getElementById(hourId);
  const minSel  = document.getElementById(minId);
  if (!hourSel || !minSel || hourSel.dataset.init) return;
  for (let h = 0; h < 24; h++) {
    const o = document.createElement('option');
    o.value = o.textContent = String(h).padStart(2,'0');
    hourSel.appendChild(o);
  }
  for (let m = 0; m < 60; m++) {
    const o = document.createElement('option');
    o.value = o.textContent = String(m).padStart(2,'0');
    minSel.appendChild(o);
  }
  hourSel.dataset.init = '1';
}
function initPostedTimeSelects() {
  fillTimeSelects('mPostedHour', 'mPostedMin');
  fillTimeSelects('bePostedHour', 'bePostedMin');
}
function initDateMirrors() {
  bindDateMirror('mEffDate','mActDate',() => actDateTouched,v => { effDateTouched=v; },v => { actDateTouched=v; });
  bindDateMirror('beEffDate','beActDate',() => bulkActDateTouched,v => { bulkEffDateTouched=v; },v => { bulkActDateTouched=v; });
}
function resetDateMirror(scope) {
  if (scope === 'bulk') {
    bulkEffDateTouched = false;
    bulkActDateTouched = false;
  } else {
    effDateTouched = false;
    actDateTouched = false;
  }
}
function onFeeSel(sel) {
  const inputId = sel.id.replace('Sel','');
  const inp = document.getElementById(inputId);
  if (!inp) return;
  if (sel.value === '__amt__') {
    inp.style.display = '';
    inp.value = '';
    inp.focus();
  } else {
    inp.style.display = 'none';
    inp.value = sel.value;
  }
}
function initFeeField(selId, inputId, val) {
  const sel = document.getElementById(selId);
  const inp = document.getElementById(inputId);
  if (!sel || !inp) return;
  if (!val) {
    sel.value = ''; inp.value = ''; inp.style.display = 'none';
  } else if (['Waived','POC','NA'].includes(val)) {
    sel.value = val; inp.value = val; inp.style.display = 'none';
  } else {
    sel.value = '__amt__'; inp.value = val; inp.style.display = '';
  }
}
function resetFeeSelects(fieldPairs) {
  fieldPairs.forEach(([selId, inputId]) => {
    const sel = document.getElementById(selId);
    const inp = document.getElementById(inputId);
    if (sel) sel.value = '';
    if (inp) { inp.value = ''; inp.style.display = 'none'; }
  });
}
function setSelectVal(el, val) {
  el.value = val;
  if (el.tagName==='SELECT' && el.value!==val) {
    const opt = document.createElement('option'); opt.value=val; opt.textContent=val;
    el.appendChild(opt); el.value=val;
  }
}
function clearMo() {
  resetDateMirror('single');
  Object.keys(mMap).forEach(id => {
    const el = document.getElementById(id); if (el) el.value = id==='mStatus'?'Available':id==='mPosted'?'No':'';
  });
  resetFeeSelects(FEE_FIELDS);
  document.getElementById('mNumber')?.classList.remove('err');
}
function fillMo(r) {
  _editUpdatedAt = r.updatedAt || null;
  resetDateMirror('single');
  Object.entries(mMap).forEach(([id,key]) => {
    const el = document.getElementById(id); if (!el) return;
    const raw = r[key];
    // Reset (don't skip) fields the record doesn't define, so no stale value from a
    // previously-opened entry lingers in the form — e.g. imported records with no posted time.
    let v = (raw === undefined || raw === null) ? '' : (DATE_FIELDS.has(id) ? sanitizeDate(raw) : raw);
    if (id === 'mStatus') v = v || 'Available';
    if (id === 'mPosted') v = canonPostedStatus(v) || 'No';
    if (el.tagName==='SELECT') setSelectVal(el,v); else el.value=v;
  });
  FEE_FIELDS.forEach(([selId, inputId]) => {
    const key = mMap[inputId];
    if (key !== undefined) initFeeField(selId, inputId, r[key] || '');
  });
}
function openAdd() {
  editId=null; document.getElementById('moTitle').textContent='Add Number';
  clearMo(); resetDeactSection('single');
  const btn = document.getElementById('mDeactBtn'); if (btn) btn.style.display = 'none';
  updatePinBtnState(null);
  document.getElementById('moOv').classList.add('on');
}
function openEdit() { if (curRec) openEditById(curRec.id); }
function openEditById(id) {
  const r = DB.find(x => x.id===id); if (!r) return;
  editId=id; document.getElementById('moTitle').textContent='Edit Number';
  fillMo(r); resetDeactSection('single');
  const btn = document.getElementById('mDeactBtn'); if (btn) btn.style.display = '';
  updatePinBtnState(id);
  document.getElementById('moOv').classList.add('on');
}
function closeMo() { document.getElementById('moOv').classList.remove('on'); }
function resetDeactSection(mode) {
  const isSingle = mode === 'single';
  const sec = document.getElementById(isSingle ? 'mDeactSection' : 'bDeactSection');
  const btn = document.getElementById(isSingle ? 'mDeactBtn' : 'bDeactBtn');
  if (sec) sec.style.display = 'none';
  if (btn) { btn.classList.remove('active'); btn.textContent = isSingle ? 'Deactivate' : 'Deactivate Selected'; }
  const d = document.getElementById(isSingle ? 'dDeactDate' : 'bdDeactDate');
  const r = document.getElementById(isSingle ? 'dRoute' : 'bdRoute');
  const m = document.getElementById(isSingle ? 'dRemarks' : 'bdRemarks');
  if (d) d.value = ''; if (r) r.value = ''; if (m) m.value = '';
}
function toggleDeactivate(mode) {
  const isSingle = mode === 'single';
  const sec = document.getElementById(isSingle ? 'mDeactSection' : 'bDeactSection');
  const btn = document.getElementById(isSingle ? 'mDeactBtn' : 'bDeactBtn');
  const isOn = sec.style.display === 'block';
  if (isOn) {
    resetDeactSection(mode);
  } else {
    sec.style.display = 'block';
    btn.classList.add('active');
    btn.textContent = '✕ Cancel Deactivate';
  }
}

async function saveRec() {
  // ── Validation ───
  const numEl = document.getElementById('mNumber');
  const numVal = numEl?.value.trim();
  if (!numVal) {
    numEl?.classList.add('err');
    showToast('Number field is required.', 'warning');
    numEl?.focus();
    return;
  }
  numEl?.classList.remove('err');

  if (effDateTouched && !actDateTouched) {
    document.getElementById('mActDate').value = document.getElementById('mEffDate').value;
  }
  const nd = {};
  Object.entries(mMap).forEach(([id,key]) => { const el=document.getElementById(id); if(el) nd[key]=el.value; });
  nd.updatedBy = currentUser?.email || 'system';
  nd.updatedAt = new Date().toISOString();

  // ── Deactivation ───
  const isDeact = editId && document.getElementById('mDeactSection')?.style.display === 'block';
  if (isDeact) {
    const deactDateVal = document.getElementById('dDeactDate').value;
    if (!deactDateVal) { showToast('Deactivation date is required.', 'warning'); document.getElementById('dDeactDate').focus(); return; }
    const currentRec = DB.find(r => r.id === editId);
    const histEntry = {
      previousClient: currentRec?.client || '',
      activation: activationSnapshot(currentRec || {}),
      deactDate: deactDateVal,
      requestedBy: document.getElementById('dRoute').value,
      remarks: document.getElementById('dRemarks').value,
      deactivatedBy: currentUser?.email || 'system',
      deactivatedAt: nd.updatedAt
    };
    nd.client = ''; nd.status = 'Available'; nd.remarks = ''; nd.postedStatus = '';
    nd.postedDate = ''; nd.postedHour = ''; nd.postedMin = ''; nd.postedTimeAt = '';
    nd.clientOSF = ''; nd.clientMRC = ''; nd.clientOTRF = '';
    nd.clientCF = ''; nd.clientCPM = ''; nd.effDate = ''; nd.actDate = '';
    nd.deactDate = deactDateVal;
    nd.route = document.getElementById('dRoute').value;
    nd.prevClient = currentRec?.client || '';
    nd.deactivationHistory = [...(currentRec?.deactivationHistory || []), histEntry];
  }

  // Persist Posted Status in canonical form. Posting times for the whole "For Posting" set
  // are (re)generated together by resequencePostingTimes() after the save succeeds below.
  nd.postedStatus = canonPostedStatus(nd.postedStatus);

  try {
    if (editId) {
      // Save old state for undo
      const idx = DB.findIndex(r => r.id===editId);
      const oldRec = idx>-1 ? {...DB[idx]} : null;
      // ── Concurrent edit detection ───
      try {
        const snap = await fdb.collection('inventory').doc(editId).get();
        if (snap.exists && snap.data().updatedAt && _editUpdatedAt && snap.data().updatedAt !== _editUpdatedAt) {
          showToast('This record was modified by another user. Please reload and try again.', 'error', 7000);
          return;
        }
      } catch(e) { /* proceed on check failure */ }
      await fdb.collection('inventory').doc(editId).update(nd);
      if (idx>-1) DB[idx] = {...DB[idx], ...nd};
      await addLog('Updated', `Updated number ${nd.number}`);
      refreshInventoryRecent();
      const reseq = await resequencePostingTimes();
      renderTbl(); closeMo();
      openSP(editId);
      propagateChange([editId, ...reseq]);
      if (oldRec) {
        showUndoToast(`Updated ${nd.number}`, async () => {
          try {
            const {id:rid, ...oldData} = oldRec;
            await fdb.collection('inventory').doc(rid).update(oldData);
            const i = DB.findIndex(r => r.id===rid);
            if (i>-1) DB[i] = {...oldRec};
            refreshInventoryRecent();
            propagateChange([rid]);
            await addLog('Updated', `Reverted ${oldRec.number} (undo edit)`);
            showToast(`Reverted ${oldRec.number}`, 'success');
          } catch(e) { showToast('Revert failed: '+e.message, 'error'); }
        }, 6000, 'Updated');
      } else {
        showToast(`Updated ${nd.number}`, 'success');
      }
    } else {
      nd.createdBy = currentUser?.email || 'system';
      nd.createdAt = new Date().toISOString();
      const ref = await fdb.collection('inventory').add(nd);
      nd.id = ref.id; DB.push(nd);
      await addLog('Added', `Added number ${nd.number}`);
      refreshInventoryRecent();
      const reseq = await resequencePostingTimes();
      renderTbl(); closeMo();
      propagateChange([ref.id, ...reseq]);
      showToast(`Added ${nd.number}`, 'success');
    }
  } catch(e) { showToast('Save error: '+e.message, 'error'); }
}

// ── DELETE ────────────────────────────────────────────
function delRec(id) {
  const r = DB.find(x => x.id===id);
  document.getElementById('delRecTitle').textContent = 'Delete this record?';
  document.getElementById('delRecInfo').innerHTML = `
    <div><span style="color:var(--t2)">Number:</span> <strong>${esc(r?.number||id)}</strong></div>
    ${r?.client  ? `<div><span style="color:var(--t2)">Client:</span> ${esc(r.client)}</div>`  : ''}
    ${r?.product ? `<div><span style="color:var(--t2)">Product:</span> ${esc(r.product)}</div>` : ''}
    ${r?.status  ? `<div><span style="color:var(--t2)">Status:</span> ${esc(r.status)}</div>`  : ''}`.trim();
  document.getElementById('delRecOv').classList.add('on');
  const btn   = document.getElementById('delRecConfirmBtn');
  const fresh = btn.cloneNode(true); btn.replaceWith(fresh);
  fresh.textContent = 'Delete';
  fresh.onclick = async () => {
    document.getElementById('delRecOv').classList.remove('on');
    const savedRec = r ? {...r} : null;
    try {
      await fdb.collection('inventory').doc(id).delete();
      DB = DB.filter(x => x.id!==id); fd = fd.filter(x => x.id!==id);
      persistentSelIds.delete(id);
      if (r) await addLog('Deleted', `Deleted number ${r.number}`);
      renderTbl(); closeSP();
      propagateChange([], [id]);
      if (savedRec) {
        showUndoToast(`Deleted ${savedRec.number}`, async () => {
          try {
            const {id:rid, ...data} = savedRec;
            await fdb.collection('inventory').doc(rid).set({...data, id:rid});
            DB.push(savedRec);
            refreshInventoryRecent();
            propagateChange([rid]);
            await addLog('Added', `Restored ${savedRec.number} (undo delete)`);
            showToast(`Restored ${savedRec.number}`, 'success');
          } catch(e) { showToast('Restore failed: '+e.message, 'error'); }
        });
      }
    } catch(e) { showToast('Delete error: '+e.message, 'error'); }
  };
}

// ── CSV / EXPORT ──────────────────────────────────────
async function handleCSV(e) {
  const f = e.target.files[0]; if (!f) return;
  const text  = await readCSVText(f);
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (lines.length < 2) { showToast('CSV has no data rows.','warning'); e.target.value=''; return; }
  const hdr = parseCSVLine(lines[0]);
  const colMap = {};
  hdr.forEach((h,i) => { const field=CSV_FIELD_MAP[h.trim()]; if(field) colMap[i]=field; });
  if (!Object.values(colMap).includes('number')) { showToast('CSV must have a "Number" column.','warning'); e.target.value=''; return; }

  const ops=[], warnings=[];
  for (let i=1; i<lines.length; i++) {
    if (!lines[i].trim()) continue;
    const cols = parseCSVLine(lines[i]);
    const nd   = {};
    Object.entries(colMap).forEach(([idx,field]) => {
      const v = cols[idx]||'';
      nd[field] = (['client','product','provider','route'].includes(field)) ? v.toUpperCase() : v;
    });
    if (!nd.number) { warnings.push(`Row ${i+1}: missing Number — skipped`); continue; }
    if (nd.status) {
      const trimmed = nd.status.trim();
      const matched = ['Active','Available','Reserved','Inactive'].find(s => s.toLowerCase() === trimmed.toLowerCase());
      if (matched) {
        nd.status = matched;
      } else {
        warnings.push(`Row ${i+1}: invalid status "${nd.status}" — cleared`);
        nd.status = '';
      }
    }
    // Normalize Posted Status to a current label ("Not Yet" → "For Posting", "Yes" → "Posted")
    // so imported rows store consistently and are searchable by the label shown in the table.
    if (nd.postedStatus) nd.postedStatus = canonPostedStatus(nd.postedStatus);
    DATE_CSV_FIELDS.forEach(field => {
      if (nd[field]) {
        const clean = sanitizeDate(nd[field]);
        if (!clean) { warnings.push(`Row ${i+1}: invalid date in ${field} — cleared`); nd[field]=''; }
        else nd[field] = clean;
      }
    });
    nd.updatedBy = currentUser?.email||'system';
    nd.updatedAt = new Date().toISOString();
    const ndNorm = normalizePhone(nd.number);
    const isNA = String(nd.number).trim().toUpperCase() === 'NA';
    const existing = isNA ? null : DB.find(r => normalizePhone(r.number) === ndNorm);
    if (existing) {
      // Strip empty-string values so existing non-blank data is not overwritten.
      // Always keep the existing number format — never overwrite with the CSV's format.
      const updateData = {updatedBy: nd.updatedBy, updatedAt: nd.updatedAt};
      Object.entries(nd).forEach(([k, v]) => {
        if (k === 'number') return;
        if (k !== 'updatedBy' && k !== 'updatedAt' && v !== '' && v !== null && v !== undefined) {
          updateData[k] = v;
        }
      });
      ops.push({type:'update', ref:fdb.collection('inventory').doc(existing.id), data:updateData, id:existing.id});
    } else {
      nd.createdBy = currentUser?.email||'system';
      nd.createdAt = new Date().toISOString();
      const ref = fdb.collection('inventory').doc();
      nd.id = ref.id;
      ops.push({type:'set', ref, data:nd});
    }
  }
  if (!ops.length) { showToast('No valid rows found.','warning'); e.target.value=''; return; }
  if (warnings.length) {
    console.warn('CSV import warnings:', warnings);
    showToast(`${warnings.length} row(s) had issues and were skipped or corrected. Check the browser console for details.`, 'warning', 7000);
  }
  try {
    const CHUNK = 400;
    for (let i=0; i<ops.length; i+=CHUNK) {
      const b = fdb.batch();
      ops.slice(i,i+CHUNK).forEach(op => op.type==='update' ? b.update(op.ref,op.data) : b.set(op.ref,op.data));
      await b.commit();
    }
    let added=0, updated=0;
    ops.forEach(op => {
      if (op.type==='update') { const idx=DB.findIndex(r=>r.id===op.id); if(idx>-1) DB[idx]={...DB[idx],...op.data}; updated++; }
      else { DB.push(op.data); added++; }
    });
    refreshInventoryRecent();
    // Uploaded "For Posting" rows arrive without a time — sequence the whole set now
    // (same as a modal save) so they get chronological HH:MM instead of staying blank.
    const reseq = await resequencePostingTimes();
    renderTbl();
    await addLog('CSV Upload', `"${f.name}": ${added} added, ${updated} updated`);
    // Propagate only the rows this import actually touched; broadcastSync escalates
    // to a full reload on its own if that set exceeds the incremental cap.
    const csvIds = ops.map(o => o.type === 'update' ? o.id : o.data.id).filter(Boolean);
    propagateChange([...csvIds, ...reseq]);
    showToast(`Upload complete — ${added} added, ${updated} updated`, 'success');
  } catch(err) { showToast('Import error: '+err.message, 'error'); }
  e.target.value='';
}

function dlSample() {
  const row = ['TOKU','DID Local','+15550001234','Active','Sample','No','','','100.00','50.00','25.00','10.00','0.0050','2024-01-01','2024-01-15','Twilio','2023-12-15','2024-01-15','80.00','40.00','20.00','0.0040','SIP','Katherine Serrano','','DIDLOGIC'];
  dlCSV([CSV_HEADERS,row], 'sample_inventory.csv');
  addLog('Exported','Downloaded sample CSV template');
}

function exportAll() {
  const rows = DB.map(r => [r.client,r.product,r.number,r.status,r.remarks,canonPostedStatus(r.postedStatus),r.postedDate,r.postedHour?(r.postedHour+':'+(r.postedMin||'00')):'',r.clientOSF,r.clientMRC,r.clientOTRF,r.clientCF,r.clientCPM,r.effDate,r.actDate,r.provider,r.arrDate,r.provActDate,r.provOSF,r.provMRC,r.provOTRF,r.provCPM,r.typeSession,r.route,r.deactDate,r.prevClient]);
  dlCSV([CSV_HEADERS,...rows], 'inventory_export.csv');
  closeExportMenu();
  addLog('Exported', `Exported ${DB.length} records to CSV`);
}

// ── SHARED EXCEL STYLING ──────────────────────────────
// Black header w/ white bold text, sensible column widths, and a narrow wrapped Remarks column.
// (Requires the xlsx-js-style build; the plain SheetJS build silently ignores the `.s` styles.)
const XL_HEADER_STYLE = {
  fill: { patternType: 'solid', fgColor: { rgb: '000000' } },
  font: { color: { rgb: 'FFFFFF' }, bold: true },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
};
// Column width (in characters) by header name — Remarks kept narrow and wrapped so it doesn't hog space.
const XL_COL_WIDTHS = {
  'Client': 16, 'Product': 18, 'Number': 16, 'Status': 12, 'Remarks': 32,
  'Posted Status': 13, 'Posted Date': 13, 'Posted Time': 11,
  'Client OSF': 11, 'Client MRC': 11, 'Client OTRF': 11, 'Client Channel Fee': 16, 'Client CPM': 11,
  'Effective Date': 14, 'Activated Date': 14, 'Provider': 16, 'Arrival Date': 13,
  'Provider Activation Date': 22, 'Provider OSF': 12, 'Provider MRC': 12, 'Provider OTRF': 12, 'Provider CPM': 12,
  'Type / Session': 14, 'Route Request by': 18, 'Deactivation Date': 16, 'Previous Client': 16
};
function styleExcelSheet(ws) {
  if (!ws || !ws['!ref']) return ws;
  const range = XLSX.utils.decode_range(ws['!ref']);
  // Read header names from the first row to map columns
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: range.s.r, c });
    headers[c] = ws[addr] ? String(ws[addr].v) : '';
  }
  const remarksCol = headers.indexOf('Remarks');
  // Column widths
  ws['!cols'] = headers.map(h => ({ wch: XL_COL_WIDTHS[h] || 14 }));
  // Style the header row (black background, white bold text)
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: range.s.r, c });
    if (ws[addr]) ws[addr].s = XL_HEADER_STYLE;
  }
  // Wrap the Remarks column body cells
  if (remarksCol !== -1) {
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const addr = XLSX.utils.encode_cell({ r, c: remarksCol });
      if (ws[addr]) ws[addr].s = { alignment: { wrapText: true, vertical: 'top' } };
    }
  }
  return ws;
}

function exportExcel() {
  if (typeof XLSX === 'undefined') { showToast('Excel library not loaded yet. Try again in a moment.','warning'); return; }
  const rows = DB.map(r => ({'Client':r.client,'Product':r.product,'Number':r.number,'Status':r.status,'Remarks':r.remarks,'Posted Status':canonPostedStatus(r.postedStatus),'Posted Date':r.postedDate,'Posted Time':r.postedHour?(r.postedHour+':'+(r.postedMin||'00')):'','Client OSF':r.clientOSF,'Client MRC':r.clientMRC,'Client OTRF':r.clientOTRF,'Client Channel Fee':r.clientCF,'Client CPM':r.clientCPM,'Effective Date':r.effDate,'Activated Date':r.actDate,'Provider':r.provider,'Arrival Date':r.arrDate,'Provider Activation Date':r.provActDate,'Provider OSF':r.provOSF,'Provider MRC':r.provMRC,'Provider OTRF':r.provOTRF,'Provider CPM':r.provCPM,'Type / Session':r.typeSession,'Route Request by':r.route,'Deactivation Date':r.deactDate,'Previous Client':r.prevClient}));
  const ws = XLSX.utils.json_to_sheet(rows);
  styleExcelSheet(ws);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Inventory');
  XLSX.writeFile(wb, 'inventory_export.xlsx');
  closeExportMenu();
  addLog('Exported', `Exported ${DB.length} records to Excel`);
}

// ── GOOGLE SHEETS SYNC ────────────────────────────────
const GS_INTL_PREFIXES = [
  'USA','AUSTRALIA','UK','UNITED KINGDOM','SINGAPORE','CANADA','JAPAN',
  'HONG KONG','MALAYSIA','INDONESIA','INDIA','CHINA','KOREA','TAIWAN',
  'THAILAND','VIETNAM','NEW ZEALAND','GERMANY','FRANCE','ITALY','SPAIN',
  'BRAZIL','MEXICO','SAUDI','UAE','DUBAI','INTERNATIONAL','INTL'
];
const gsIsIntl  = p => { const u = String(p||'').toUpperCase().trim(); return u.endsWith(' DID') || GS_INTL_PREFIXES.some(x => u.startsWith(x)); };
const gsIsNANum = r => String(r.number||'').trim().toUpperCase() === 'NA';

let _gsTokenClient = null;
let _gsAccessToken = null;

function gsRecordToRow(r) {
  return [
    r.client||'', r.product||'', r.number||'', r.status||'', r.remarks||'',
    canonPostedStatus(r.postedStatus), r.postedDate||'', r.postedHour?(r.postedHour+':'+(r.postedMin||'00')):'', r.clientOSF||'', r.clientMRC||'',
    r.clientOTRF||'', r.clientCF||'', r.clientCPM||'', r.effDate||'', r.actDate||'',
    r.provider||'', r.arrDate||'', r.provActDate||'', r.provOSF||'', r.provMRC||'',
    r.provOTRF||'', r.provCPM||'', r.typeSession||'', r.route||'',
    r.deactDate||'', r.prevClient||''
  ];
}

async function gsAPI(method, url, body) {
  const opts = { method, headers: { 'Authorization': `Bearer ${_gsAccessToken}` } };
  if (body !== undefined) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  }
  const resp = await fetch(url, opts);
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err?.error?.message || `HTTP ${resp.status}`);
  }
  return resp.json();
}

function openGSSettings() {
  closeExportMenu();
  const { clientId, spreadsheetId } = gsGetSettings();
  document.getElementById('gsClientId').value = clientId;
  document.getElementById('gsSpreadsheetId').value = spreadsheetId;
  document.getElementById('gsOriginHint').textContent = location.origin;
  document.getElementById('gsSettingsOv').classList.add('on');
}
function closeGSSettings() { document.getElementById('gsSettingsOv').classList.remove('on'); }

function gsGetSettings() {
  return {
    clientId: localStorage.getItem('gs-client-id') || '',
    spreadsheetId: localStorage.getItem('gs-spreadsheet-id') || ''
  };
}

function saveGSSettings() {
  const clientId = document.getElementById('gsClientId').value.trim();
  let raw = document.getElementById('gsSpreadsheetId').value.trim();
  const m = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  const spreadsheetId = m ? m[1] : raw;
  if (!clientId || !spreadsheetId) { showToast('Both fields are required.', 'warning'); return; }
  localStorage.setItem('gs-client-id', clientId);
  localStorage.setItem('gs-spreadsheet-id', spreadsheetId);
  _gsTokenClient = null; // reset so it re-initialises with new client ID
  closeGSSettings();
  showToast('Google Sheets settings saved.', 'success');
}

function syncToGoogleSheets() {
  closeExportMenu();
  const { clientId, spreadsheetId } = gsGetSettings();
  if (!clientId || !spreadsheetId) {
    openGSSettings();
    showToast('Configure your Google Sheets settings first.', 'info');
    return;
  }
  if (typeof google === 'undefined' || !google.accounts) {
    showToast('Google library not loaded yet. Try again in a moment.', 'warning'); return;
  }
  if (!_gsTokenClient) {
    _gsTokenClient = google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: 'https://www.googleapis.com/auth/spreadsheets',
      callback: async resp => {
        if (resp.error) { showToast('Google auth failed: ' + resp.error, 'error'); return; }
        _gsAccessToken = resp.access_token;
        await doGSSync(spreadsheetId);
      }
    });
  }
  _gsTokenClient.requestAccessToken({ prompt: _gsAccessToken ? '' : 'consent' });
}

function gsDismissSyncingToast() {
  document.querySelectorAll('.toast.t-info').forEach(t => {
    if (t.querySelector('.toast-msg')?.textContent?.includes('Syncing')) dismissToast(t);
  });
}

async function doGSSync(spreadsheetId) {
  try {
    showToast('Syncing to Google Sheets…', 'info', 60000);
    const base = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;

    const HEADERS = ['Client','Product','Number','Status','Remarks','Posted Status','Posted Date','Posted Time',
      'Client OSF','Client MRC','Client OTRF','Client Channel Fee','Client CPM',
      'Effective Date','Activated Date','Provider','Arrival Date','Provider Activation Date',
      'Provider OSF','Provider MRC','Provider OTRF','Provider CPM',
      'Type / Session','Route Request by','Deactivation Date','Previous Client'];

    // Build tab list
    const intlRecords = DB.filter(r => gsIsIntl(r.product));
    const naRecords   = DB.filter(r => gsIsNANum(r));
    const localProds  = [...new Set(DB.map(r => r.product).filter(p => p && !gsIsIntl(p)))].sort();
    const tabs = [
      { name: 'All Data',      records: DB },
      ...(intlRecords.length ? [{ name: 'International', records: intlRecords }] : []),
      ...(naRecords.length   ? [{ name: 'NA Numbers',    records: naRecords    }] : []),
      ...localProds.map(p => ({ name: p.slice(0,100), records: DB.filter(r => r.product===p && !gsIsNANum(r)) })).filter(t => t.records.length)
    ];

    // 1 — Get existing sheets
    const info = await gsAPI('GET', base);
    const existing = new Map(info.sheets.map(s => [s.properties.title, s.properties.sheetId]));

    // 2 — Create missing tabs in one batch call
    const toCreate = tabs.filter(t => !existing.has(t.name));
    if (toCreate.length) {
      const created = await gsAPI('POST', `${base}:batchUpdate`, {
        requests: toCreate.map(t => ({ addSheet: { properties: { title: t.name } } }))
      });
      (created.replies || []).forEach(rep => {
        const p = rep.addSheet && rep.addSheet.properties;
        if (p) existing.set(p.title, p.sheetId);
      });
    }

    // 3 — Batch clear all tab ranges (1 API call)
    await gsAPI('POST', `${base}/values:batchClear`, {
      ranges: tabs.map(t => `'${t.name}'!A:Z`)
    });

    // 4 — Batch write all tabs (1 API call)
    await gsAPI('POST', `${base}/values:batchUpdate`, {
      valueInputOption: 'RAW',
      data: tabs.map(t => ({
        range: `'${t.name}'!A1`,
        values: [HEADERS, ...t.records.map(gsRecordToRow)]
      }))
    });

    // 5 — Format every tab: black header w/ white text, auto-sized columns, narrow wrapped Remarks (1 API call)
    const remarksCol = HEADERS.indexOf('Remarks');
    const fmtRequests = [];
    tabs.forEach(t => {
      const sheetId = existing.get(t.name);
      if (sheetId === undefined) return;
      // Header row: black fill, white bold text, centered
      fmtRequests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
          cell: { userEnteredFormat: {
            backgroundColor: { red: 0, green: 0, blue: 0 },
            textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
          } },
          fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
        }
      });
      // Freeze the header row
      fmtRequests.push({
        updateSheetProperties: {
          properties: { sheetId, gridProperties: { frozenRowCount: 1 } },
          fields: 'gridProperties.frozenRowCount'
        }
      });
      // Auto-size every column except Remarks so nothing looks compact
      if (remarksCol > 0) {
        fmtRequests.push({ autoResizeDimensions: { dimensions: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: remarksCol } } });
      }
      fmtRequests.push({ autoResizeDimensions: { dimensions: { sheetId, dimension: 'COLUMNS', startIndex: remarksCol + 1, endIndex: HEADERS.length } } });
      // Remarks: fixed width + wrap
      fmtRequests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: 'COLUMNS', startIndex: remarksCol, endIndex: remarksCol + 1 },
          properties: { pixelSize: 260 }, fields: 'pixelSize'
        }
      });
      fmtRequests.push({
        repeatCell: {
          range: { sheetId, startColumnIndex: remarksCol, endColumnIndex: remarksCol + 1 },
          cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } },
          fields: 'userEnteredFormat.wrapStrategy'
        }
      });
    });
    if (fmtRequests.length) {
      await gsAPI('POST', `${base}:batchUpdate`, { requests: fmtRequests });
    }

    gsDismissSyncingToast();
    showToast(`Synced — ${DB.length} records across ${tabs.length} tabs`, 'success', 6000);
    addLog('Exported', `Synced ${DB.length} records to Google Sheets (${tabs.length} tabs)`);
  } catch(e) {
    gsDismissSyncingToast();
    showToast('Sync failed: ' + e.message, 'error', 8000);
    console.error('GS sync error:', e);
  }
}

function exportGoogleSheets() {
  if (typeof XLSX === 'undefined') { showToast('Excel library not loaded yet. Try again in a moment.','warning'); return; }

  function recordToRow(r) {
    return {
      'Client': r.client||'', 'Product': r.product||'', 'Number': r.number||'',
      'Status': r.status||'', 'Remarks': r.remarks||'', 'Posted Status': canonPostedStatus(r.postedStatus),
      'Posted Date': r.postedDate||'', 'Posted Time': r.postedHour?(r.postedHour+':'+(r.postedMin||'00')):'', 'Client OSF': r.clientOSF||'', 'Client MRC': r.clientMRC||'',
      'Client OTRF': r.clientOTRF||'', 'Client Channel Fee': r.clientCF||'', 'Client CPM': r.clientCPM||'',
      'Effective Date': r.effDate||'', 'Activated Date': r.actDate||'', 'Provider': r.provider||'',
      'Arrival Date': r.arrDate||'', 'Provider Activation Date': r.provActDate||'',
      'Provider OSF': r.provOSF||'', 'Provider MRC': r.provMRC||'', 'Provider OTRF': r.provOTRF||'',
      'Provider CPM': r.provCPM||'', 'Type / Session': r.typeSession||'',
      'Route Request by': r.route||'', 'Deactivation Date': r.deactDate||'', 'Previous Client': r.prevClient||''
    };
  }

  // Country-name prefixes that indicate an international product
  const INTL_PREFIXES = [
    'USA','AUSTRALIA','UK','UNITED KINGDOM','SINGAPORE','CANADA','JAPAN',
    'HONG KONG','MALAYSIA','INDONESIA','INDIA','CHINA','KOREA','TAIWAN',
    'THAILAND','VIETNAM','NEW ZEALAND','GERMANY','FRANCE','ITALY','SPAIN',
    'BRAZIL','MEXICO','SAUDI','UAE','DUBAI','INTERNATIONAL','INTL'
  ];
  const isIntl  = p => { const u = String(p||'').toUpperCase().trim(); return u.endsWith(' DID') || INTL_PREFIXES.some(x => u.startsWith(x)); };
  const isNANum = r => String(r.number||'').trim().toUpperCase() === 'NA';

  // Safe Excel sheet name: max 31 chars, no \ / ? * [ ] :
  const usedNames = new Set();
  function sheetName(raw) {
    let n = String(raw).replace(/[\\\/\?\*\[\]:]/g,'_').slice(0,31);
    if (!usedNames.has(n)) { usedNames.add(n); return n; }
    // deduplicate with a numeric suffix
    for (let i=2; i<100; i++) {
      const s = n.slice(0,28)+'_'+i;
      if (!usedNames.has(s)) { usedNames.add(s); return s; }
    }
    return n;
  }

  function makeSheet(records) {
    const rows = records.map(recordToRow);
    return styleExcelSheet(XLSX.utils.json_to_sheet(rows.length ? rows : [recordToRow({})]));
  }

  const wb = XLSX.utils.book_new();

  // Tab 1 — All Data
  XLSX.utils.book_append_sheet(wb, makeSheet(DB), sheetName('All Data'));

  // International tab
  const intlRecords = DB.filter(r => isIntl(r.product));
  if (intlRecords.length) {
    XLSX.utils.book_append_sheet(wb, makeSheet(intlRecords), sheetName('International'));
  }

  // NA Numbers tab
  const naRecords = DB.filter(r => isNANum(r));
  if (naRecords.length) {
    XLSX.utils.book_append_sheet(wb, makeSheet(naRecords), sheetName('NA Numbers'));
  }

  // Per-product tabs (local/domestic products only, sorted — NA numbers excluded, they're in NA Numbers tab)
  const localProducts = [...new Set(DB.map(r => r.product).filter(p => p && !isIntl(p)))].sort();
  localProducts.forEach(product => {
    const recs = DB.filter(r => r.product === product && !isNANum(r));
    if (!recs.length) return;
    XLSX.utils.book_append_sheet(wb, makeSheet(recs), sheetName(product));
  });

  XLSX.writeFile(wb, 'inventory_sheets.xlsx');
  closeExportMenu();
  const tabCount = 1 + (intlRecords.length?1:0) + (naRecords.length?1:0) + localProducts.length;
  addLog('Exported', `Exported to Google Sheets format — ${DB.length} records across ${tabCount} tabs`);
}

function exportPDF() {
  const jsPDFCls = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
  if (!jsPDFCls) { showToast('PDF library not loaded yet. Try again in a moment.','warning'); return; }
  const doc  = new jsPDFCls({orientation:'landscape',unit:'mm',format:'a4'});
  const cols = ['#','Client','Product','Number','Status','Remarks','Act. Date','Provider'];
  const rows = DB.map((r,i) => [i+1,r.client||'',r.product||'',r.number||'',r.status||'',r.remarks||'',r.actDate||'',r.provider||'']);
  doc.setFontSize(14); doc.text('CS Inventory', 14, 14);
  doc.setFontSize(9);  doc.text(`Exported: ${new Date().toLocaleString()} — ${DB.length} records`, 14, 20);
  doc.autoTable({head:[cols],body:rows,startY:25,styles:{fontSize:8},headStyles:{fillColor:[26,115,232]}});
  doc.save('inventory_export.pdf');
  closeExportMenu();
  addLog('Exported', `Exported ${DB.length} records to PDF`);
}

// Turn a 2-D array of rows into a UTF-8 CSV string (BOM + quoted/escaped cells).
function csvString(rows) {
  return '\uFEFF' + rows.map(r => r.map(c => `"${String(c==null?'':c).replace(/"/g,'""')}"`).join(',')).join('\n');
}
function dlCSV(rows, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csvString(rows)],{type:'text/csv;charset=utf-8'}));
  a.download = name; a.click();
}

// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550
// AUTO BACKUP \u2014 weekly inventory CSV emailed to the logged-in address
// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550

// ISO-8601 week key, e.g. "2026-W28". Fri/Sat/Sun all share one week's key.
function isoWeekKey(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = d.getUTCDay() || 7;                 // Mon=1 \u2026 Sun=7
  d.setUTCDate(d.getUTCDate() + 4 - day);         // shift to the week's Thursday
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return d.getUTCFullYear() + '-W' + String(week).padStart(2, '0');
}

// Due only on Friday, Saturday or Sunday of a week we haven't emailed yet.
function isBackupDue() {
  if (!AB || !AB.enabled) return false;
  const now = new Date();
  const dow = now.getDay();                        // 0 Sun \u2026 5 Fri \u2026 6 Sat
  if (!(dow === 5 || dow === 6 || dow === 0)) return false;   // Mon\u2013Thu: not yet
  return AB.lastSentWeek !== isoWeekKey(now);
}

// Whole-inventory CSV (same columns/order as the manual "Export to CSV").
function buildBackupCSV() {
  const rows = DB.map(r => [r.client,r.product,r.number,r.status,r.remarks,canonPostedStatus(r.postedStatus),r.postedDate,r.postedHour?(r.postedHour+':'+(r.postedMin||'00')):'',r.clientOSF,r.clientMRC,r.clientOTRF,r.clientCF,r.clientCPM,r.effDate,r.actDate,r.provider,r.arrDate,r.provActDate,r.provOSF,r.provMRC,r.provOTRF,r.provCPM,r.typeSession,r.route,r.deactDate,r.prevClient]);
  return csvString([CSV_HEADERS, ...rows]);
}

// UTF-8-safe base64 (handles the BOM and any non-Latin1 characters).
function toB64Utf8(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

// Read config once at sign-in, refresh the Admin card, and arm the checks.
async function initAutoBackup() {
  try {
    const snap = await fdb.collection('meta').doc('autoBackup').get();
    AB = snap.exists ? (snap.data() || {}) : { enabled:false, recipient:'', lastSentWeek:'', lastSentAt:'' };
  } catch (e) { console.error('initAutoBackup:', e); AB = { enabled:false }; }
  if (!AB.recipient && currentUser?.email) AB.recipient = currentUser.email;
  renderAutoBackupCard();
  if (_abTimer) clearInterval(_abTimer);
  _abTimer = setInterval(maybeRunAutoBackup, 60 * 60 * 1000);   // re-check hourly for long-open sessions
  maybeRunAutoBackup();
}

// Gate everything, then wait until the inventory is actually loaded before sending.
function maybeRunAutoBackup() {
  if (!AB || !AB.enabled) return;
  if (!(currentRole === 'admin' || currentRole === 'semi-admin')) return;  // only editors can write meta / send
  if (!isBackupDue()) return;
  if (_invLoading || !DB.length) { clearTimeout(_abDeferT); _abDeferT = setTimeout(maybeRunAutoBackup, 5000); return; }
  runWeeklyBackup();
}

// Claim the week (transaction \u2192 one sender), build the file, email it, log it.
async function runWeeklyBackup(opts = {}) {
  if (_abSending) return;
  const test = !!opts.test;
  const recipient = (AB && AB.recipient) || currentUser?.email || '';
  if (!recipient)        { if (test) showToast('No recipient email found.', 'warning'); return; }
  if (!BACKUP_MAILER_URL || !BACKUP_MAILER_KEY){ showToast('Backup email isn\u2019t set up yet \u2014 paste your SECRET into BACKUP_MAILER_KEY.', 'warning'); return; }
  if (!DB.length)        { if (test) showToast('No inventory loaded yet.', 'warning'); return; }

  _abSending = true;
  const ref = fdb.collection('meta').doc('autoBackup');
  const wk  = isoWeekKey(new Date());
  let prevWeek = (AB && AB.lastSentWeek) || '';
  try {
    if (!test) {
      // Atomically claim this week so a second open tab / device won't double-send.
      let claimed = false;
      await fdb.runTransaction(async tx => {
        const snap = await tx.get(ref);
        const d = snap.exists ? (snap.data() || {}) : {};
        if (!d.enabled) return;
        if (d.lastSentWeek === wk) return;         // someone already sent this week
        prevWeek = d.lastSentWeek || '';
        tx.set(ref, { lastSentWeek: wk }, { merge:true });
        claimed = true;
      });
      if (!claimed) { _abSending = false; return; }
      AB.lastSentWeek = wk;
    }

    const csv   = buildBackupCSV();
    const fname = `CS-Inventory-Backup-${new Date().toISOString().slice(0,10)}.csv`;
    await deliverBackupEmail(recipient, fname, csv, DB.length, test);

    const stamp = new Date().toISOString();
    if (!test) {
      try { await ref.set({ lastSentAt: stamp, lastSentBy: currentUser?.email || '' }, { merge:true }); } catch(_){}
      AB.lastSentAt = stamp;
    }
    await addLog('Backup', `${test ? 'Test' : 'Weekly'} backup emailed to ${recipient} (${DB.length} records)`);
    showToast(`Backup emailed to ${recipient} \u2713`, 'success');
    renderAutoBackupCard();
  } catch (err) {
    console.error('runWeeklyBackup:', err);
    if (!test) {                                    // network failure \u2192 release the claim so it retries next open
      try { await ref.set({ lastSentWeek: prevWeek }, { merge:true }); } catch(_){}
      AB.lastSentWeek = prevWeek;
    }
    showToast('Backup could not be sent \u2014 will retry.', 'error');
    renderAutoBackupCard();
  } finally { _abSending = false; }
}

// Fire-and-forget POST to the Apps Script mailer. `no-cors` keeps it a "simple"
// cross-origin request (no preflight); it resolves unless the network truly fails.
async function deliverBackupEmail(to, filename, csvText, count, test) {
  const today = new Date();
  const payload = {
    token: BACKUP_MAILER_KEY,
    to,
    subject: `CS Inventory \u2014 ${test ? 'Test ' : ''}Weekly Backup (${today.toISOString().slice(0,10)})`,
    body: `Automated ${test ? 'test ' : ''}weekly backup of CS Inventory.\n\n`
        + `Records: ${count}\nGenerated: ${today.toLocaleString()}\n\n`
        + `The full inventory is attached as a CSV file.`,
    filename,
    mimeType: 'text/csv',
    dataB64: toB64Utf8(csvText)
  };
  await fetch(BACKUP_MAILER_URL, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },   // avoids CORS preflight
    body: JSON.stringify(payload)
  });
}

// Admin toggle handler.
async function setAutoBackupEnabled(on) {
  if (currentRole !== 'admin') { showToast('Only admins can change this.', 'warning'); renderAutoBackupCard(); return; }
  try {
    await fdb.collection('meta').doc('autoBackup').set({
      enabled: !!on,
      recipient: (AB && AB.recipient) || currentUser?.email || '',
      updatedAt: new Date().toISOString(),
      updatedBy: currentUser?.email || ''
    }, { merge:true });
    AB = AB || {}; AB.enabled = !!on; if (!AB.recipient) AB.recipient = currentUser?.email || '';
    await addLog('Backup', `Automatic weekly backup turned ${on ? 'ON' : 'OFF'}`);
    showToast(`Automatic backup ${on ? 'enabled' : 'disabled'}.`, 'success');
    renderAutoBackupCard();
    if (on) maybeRunAutoBackup();
  } catch (e) {
    console.error('setAutoBackupEnabled:', e);
    showToast('Could not update the setting.', 'error');
    renderAutoBackupCard();
  }
}

// Manual "Send test backup now" button.
function sendTestBackup() {
  if (currentRole !== 'admin') { showToast('Only admins can do this.', 'warning'); return; }
  runWeeklyBackup({ test:true });
}

// Paint the Admin \u25B8 Automatic Backup card from current state.
function renderAutoBackupCard() {
  const t = document.getElementById('abToggle');
  if (!t) return;                                   // card not on screen yet
  t.checked = !!(AB && AB.enabled);
  const rec = document.getElementById('abRecipient');
  if (rec) rec.textContent = (AB && AB.recipient) || currentUser?.email || '\u2014';
  const st = document.getElementById('abStatus');
  if (st) {
    const last = (AB && AB.lastSentAt) ? new Date(AB.lastSentAt).toLocaleString() : 'Never';
    const head = (AB && AB.enabled) ? 'On \u2014 a backup is emailed every Friday.' : 'Off \u2014 no automatic backups.';
    st.innerHTML = `${esc(head)}<br>Last sent: ${esc(last)}`;
  }
  const warn = document.getElementById('abWarn');
  if (warn) warn.style.display = (BACKUP_MAILER_URL && BACKUP_MAILER_KEY) ? 'none' : '';
}

function toggleExportMenu(e) {
  e.stopPropagation();
  document.getElementById('exportMenu').classList.toggle('on');
}
function closeExportMenu() {
  document.getElementById('exportMenu')?.classList.remove('on');
}

// ── LOGS ──────────────────────────────────────────────
async function addLog(action, details, extra={}) {
  const log = {datetime:new Date().toISOString(), user:currentUser?.email||'system', action, details, ...extra};
  try {
    const ref = await fdb.collection('logs').add(log);
    log.id = ref.id; LOGS.unshift(log); fl=[...LOGS]; renderLogs();
    _logCursor = LOGS[0]?.datetime || _logCursor;
    kvSet('logsCache', LOGS.slice(0, 500)); kvSet('logsCursor', _logCursor);   // keep cache in step
  } catch(e) { console.error('addLog:', e); }
}
function toggleLF() {
  document.getElementById('lfBody').classList.toggle('on');
  document.getElementById('lfArrow').classList.toggle('on');
}
function applyLF() {
  const act = document.getElementById('lAction').value;
  const df  = document.getElementById('lFrom').value;
  const dt  = document.getElementById('lTo').value;
  fl = LOGS.filter(r => {
    if (act && r.action!==act) return false;
    const d = r.datetime.slice(0,10);
    if (df && d<df) return false;
    if (dt && d>dt) return false;
    return true;
  });
  lpg=1; renderLogs();
}
function clearLF() {
  ['lAction','lFrom','lTo'].forEach(id => document.getElementById(id).value='');
  fl=[...LOGS]; lpg=1; renderLogs();
}
function exportLogs() {
  dlCSV([['#','Date & Time','User','Action','Details'],...fl.map((r,i) => [i+1,r.datetime,r.user,r.action,r.details])], 'logs_export.csv');
}
function getLogFilterState() {
  const act = document.getElementById('lAction').value;
  const df  = document.getElementById('lFrom').value;
  const dt  = document.getElementById('lTo').value;
  const parts = [];
  if (act) parts.push(`Action: ${act}`);
  if (df) parts.push(`From: ${fmt(df)}`);
  if (dt) parts.push(`To: ${fmt(dt)}`);
  return {act, df, dt, summary: parts.join(', ')};
}
async function deleteLogsByIds(ids) {
  const CHUNK = 400;
  for (let i=0; i<ids.length; i+=CHUNK) {
    const batch = fdb.batch();
    ids.slice(i,i+CHUNK).forEach(id => batch.delete(fdb.collection('logs').doc(id)));
    await batch.commit();
  }
  const gone = new Set(ids);
  LOGS = LOGS.filter(r => !gone.has(r.id));
  applyLF();
}
function openClearLogsConfirm({title, question, desc, note, buttonText, onConfirm}) {
  const ov = document.getElementById('clearLogsOv');
  document.getElementById('clearLogsTitle').textContent = title;
  document.getElementById('clearLogsQuestion').textContent = question;
  document.getElementById('clearLogsDesc').textContent = desc;
  document.getElementById('clearLogsNote').textContent = note;
  const btn = document.getElementById('clearLogsConfirmBtn');
  const fresh = btn.cloneNode(true); btn.replaceWith(fresh);
  fresh.textContent = buttonText;
  fresh.onclick = async function() {
    ov.classList.remove('on');
    await onConfirm();
  };
  ov.classList.add('on');
}
function clearFilteredLogs() {
  const {act, df, dt, summary} = getLogFilterState();
  if (!act && !df && !dt) {
    showToast('Choose a log filter first, or use Clear All.', 'warning');
    return;
  }
  const ids = fl.map(r => r.id).filter(Boolean);
  if (!ids.length) {
    showToast('No matching logs to clear.', 'warning');
    return;
  }
  openClearLogsConfirm({
    title: 'Clear Matching Logs',
    question: `Clear ${ids.length} matching log${ids.length!==1?'s':''}?`,
    desc: `This will permanently delete logs matching: ${summary}.`,
    note: 'This cannot be undone.',
    buttonText: `Clear ${ids.length}`,
    onConfirm: async () => {
      try {
        await deleteLogsByIds(ids);
        showToast(`Cleared ${ids.length} matching log${ids.length!==1?'s':''}.`, 'info');
      } catch(e) { showToast('Error: '+e.message, 'error'); }
    }
  });
}
function clearAllLogs() {
  openClearLogsConfirm({
    title: 'Clear All Logs',
    question: 'Clear all logs?',
    desc: 'This will permanently delete all log entries.',
    note: 'This cannot be undone.',
    buttonText: 'Clear All',
    onConfirm: async () => {
      try {
        const snap = await fdb.collection('logs').get();
        await deleteLogsByIds(snap.docs.map(d => d.id));
        showToast('All logs cleared.', 'info');
      } catch(e) { showToast('Error: '+e.message, 'error'); }
    }
  });
}
function logRecordSummary(r) {
  return {
    id: r?.id || '',
    number: r?.number || '',
    client: r?.client || '',
    product: r?.product || '',
    status: r?.status || '',
    changes: Array.isArray(r?.changes) ? r.changes : []
  };
}
function formatLogValue(v) {
  const s = String(v == null ? '' : v).trim();
  return s || 'blank';
}
function changeSummary(changes) {
  return changes.map(c => {
    if (c.field === 'status') return `From ${esc(formatLogValue(c.from))} status to ${esc(formatLogValue(c.to))} status`;
    return `${esc(c.label || FIELD_LABELS[c.field] || c.field || 'Field')}: ${esc(formatLogValue(c.from))} &rarr; ${esc(formatLogValue(c.to))}`;
  }).join('<br>');
}
function inferredLogChanges(r, log) {
  if (log?.action !== 'Updated' || !Array.isArray(log.fields)) return [];
  const current = DB.find(x => (r.id && x.id === r.id) || (r.number && x.number === r.number));
  if (!current) return [];
  return log.fields.map(field => {
    if (!(field in r)) return null;
    const from = r[field] ?? '';
    const to = current[field] ?? '';
    if (String(from) === String(to)) return null;
    return {field, label:FIELD_LABELS[field] || field, from, to};
  }).filter(Boolean);
}
function logRecordDetail(r, log) {
  if (Array.isArray(r.changes) && r.changes.length) return changeSummary(r.changes);
  const inferred = inferredLogChanges(r, log);
  if (inferred.length) return changeSummary(inferred);
  if (log?.action === 'Updated' && Array.isArray(log.fields) && log.fields.length) {
    return `${esc(log.fields.map(f => FIELD_LABELS[f] || f).join(', '))} updated`;
  }
  return r.status ? `<span class="badge ${bclass(r.status)}">${esc(r.status)}</span>` : '—';
}
function logRecordList(log) {
  if (Array.isArray(log?.records) && log.records.length) {
    return log.records.map(logRecordSummary).filter(r => r.number || r.id);
  }
  const oldDelete = String(log?.details || '').match(/^Bulk deleted \d+ records?:\s*(.+)$/i);
  if (oldDelete) {
    return oldDelete[1].split(',').map(n => ({number:n.trim(), client:'', product:'', status:''})).filter(r => r.number);
  }
  return [];
}
function renderLogDetails(log) {
  const details = String(log.details || '');
  const records = logRecordList(log);
  const m = details.match(/^(.*?)(\d+\s+records?)(.*)$/i);
  if (!records.length || !m || !log.id) return esc(details);
  return `${esc(m[1])}<button type="button" class="log-rec-link" onclick="event.stopPropagation();openLogRecords('${esc(log.id)}')">${esc(m[2])}</button>${esc(m[3])}`;
}
function openLogRecords(logId) {
  const log = LOGS.find(r => r.id === logId) || fl.find(r => r.id === logId);
  if (!log) return;
  const records = logRecordList(log);
  const ov = document.getElementById('logRecordsOv');
  const title = document.getElementById('logRecordsTitle');
  const meta = document.getElementById('logRecordsMeta');
  const body = document.getElementById('logRecordsBody');
  const lastHead = document.getElementById('logRecordsLastHead');
  if (!ov || !title || !meta || !body) return;
  const showDetails = log.action === 'Updated';
  if (lastHead) lastHead.textContent = showDetails ? 'Details' : 'Status';
  title.textContent = `${log.action || 'Log'} - ${records.length} record${records.length!==1?'s':''}`;
  meta.textContent = log.details || '';
  body.innerHTML = records.length ? records.map((r,i) => `
    <tr>
      <td class="row-num">${i+1}</td>
      <td class="num-cell">${esc(r.number || '—')}</td>
      <td>${esc(r.client || '—')}</td>
      <td>${esc(r.product || '—')}</td>
      <td>${showDetails ? logRecordDetail(r, log) : (r.status ? `<span class="badge ${bclass(r.status)}">${esc(r.status)}</span>` : '—')}</td>
    </tr>`).join('') : '<tr><td colspan="5" style="text-align:center;color:var(--t3);padding:20px">No record list stored for this log.</td></tr>';
  ov.classList.add('on');
}
function closeLogRecords() {
  document.getElementById('logRecordsOv')?.classList.remove('on');
}
function bulkChangeSummary(r, fields, updates) {
  return {
    ...logRecordSummary(r),
    changes: fields.map(field => ({
      field,
      label: FIELD_LABELS[field] || field,
      from: r?.[field] ?? '',
      to: updates?.[field] ?? ''
    }))
  };
}
function reverseBulkChanges(records) {
  return records.map(r => ({
    ...r,
    changes: (r.changes || []).map(c => ({...c, from:c.to, to:c.from}))
  }));
}
function sortL(col) {
  if (lSortCol===col) lSortDir*=-1; else { lSortCol=col; lSortDir=1; }
  fl.sort((a,b) => (a[col]||'').localeCompare(b[col]||'')*lSortDir);
  renderLogs();
}
function renderLogs() {
  const sz = parseInt(EL.lPgSize?.value || 25);
  const total = fl.length, tp = Math.ceil(total/sz)||1;
  lpg = Math.max(1, Math.min(lpg, tp));
  const s=(lpg-1)*sz, e=s+sz;
  if (EL.logBody) EL.logBody.innerHTML = fl.slice(s,e).map((r,i) => `
    <tr>
      <td class="row-num">${s+i+1}</td>
      <td>${new Date(r.datetime).toLocaleString()}</td>
      <td>${esc(r.user)}</td>
      <td><span class="badge ${ACT_LABELS[r.action]||'b-available'}">${esc(r.action)}</span></td>
      <td>${renderLogDetails(r)}</td>
    </tr>`).join('');
  if (EL.lInfo)    EL.lInfo.textContent    = `Showing ${Math.min(s+1,total)||0}–${Math.min(e,total)} of ${total} records`;
  if (EL.lPgInfo)  EL.lPgInfo.textContent  = `Page ${lpg} of ${tp}`;
  if (EL.lPgPrev)  EL.lPgPrev.disabled     = lpg<=1;
  if (EL.lPgNext)  EL.lPgNext.disabled     = lpg>=tp;
}
function changeLPg(d) {
  const sz = parseInt(EL.lPgSize?.value || 25);
  const tp = Math.ceil(fl.length/sz)||1;
  lpg = Math.max(1, Math.min(lpg+d, tp)); renderLogs();
}

// ── DOWNLOAD SELECTED ─────────────────────────────────
function dlSelected() {
  const ids = getCheckedIds(); if (!ids.length) return;
  const rows = ids.map(id => DB.find(r => r.id===id)).filter(Boolean);
  const data = rows.map(r => [r.client,r.product,r.number,r.status,r.remarks,canonPostedStatus(r.postedStatus),r.postedDate,r.postedHour?(r.postedHour+':'+(r.postedMin||'00')):'',r.clientOSF,r.clientMRC,r.clientOTRF,r.clientCF,r.clientCPM,r.effDate,r.actDate,r.provider,r.arrDate,r.provActDate,r.provOSF,r.provMRC,r.provOTRF,r.provCPM,r.typeSession,r.route,r.deactDate,r.prevClient]);
  dlCSV([CSV_HEADERS,...data], `selected_${ids.length}_entries.csv`);
  addLog('Exported', `Downloaded ${ids.length} selected record${ids.length!==1?'s':''}`);
}

// ── BULK DELETE ───────────────────────────────────────
function delSelected() {
  const ids = getCheckedIds(); if (!ids.length) return;
  const count   = ids.length;
  const preview = ids.slice(0,5).map(id => DB.find(r => r.id===id)?.number).filter(Boolean);
  document.getElementById('delRecTitle').textContent = `Delete ${count} selected record${count!==1?'s':''}?`;
  document.getElementById('delRecInfo').innerHTML = `
    <div><span style="color:var(--t2)">Records to delete:</span> <strong>${count}</strong></div>
    ${preview.length?`<div style="margin-top:4px"><span style="color:var(--t2)">Numbers:</span> ${preview.map(n=>esc(n)).join(', ')}${count>5?` <em>+${count-5} more</em>`:''}</div>`:''}`.trim();
  document.getElementById('delRecOv').classList.add('on');
  const btn   = document.getElementById('delRecConfirmBtn');
  const fresh = btn.cloneNode(true); btn.replaceWith(fresh);
  fresh.textContent = `Delete ${count}`;
  fresh.onclick = async () => {
    document.getElementById('delRecOv').classList.remove('on');
    const savedRecs = ids.map(id => ({...DB.find(r => r.id===id)})).filter(r => r.id);
    const affectedRecords = savedRecs.map(logRecordSummary);
    try {
      const CHUNK = 400;
      for (let i=0; i<ids.length; i+=CHUNK) {
        const b = fdb.batch();
        ids.slice(i,i+CHUNK).forEach(id => b.delete(fdb.collection('inventory').doc(id)));
        await b.commit();
      }
      DB = DB.filter(r => !ids.includes(r.id)); fd = fd.filter(r => !ids.includes(r.id));
      ids.forEach(id => persistentSelIds.delete(id));
      await addLog('Deleted', `Bulk deleted ${ids.length} records`, {records: affectedRecords});
      renderTbl(); closeSP();
      propagateChange([], ids);
      showUndoToast(`Deleted ${ids.length} records`, async () => {
        try {
          const CHUNK = 400;
          for (let i=0; i<savedRecs.length; i+=CHUNK) {
            const b = fdb.batch();
            savedRecs.slice(i,i+CHUNK).forEach(rec => {
              const {id, ...data} = rec;
              b.set(fdb.collection('inventory').doc(id), {...data, id});
            });
            await b.commit();
          }
          savedRecs.forEach(rec => { if (!DB.find(r => r.id===rec.id)) DB.push(rec); });
          refreshInventoryRecent();
          propagateChange(savedRecs.map(r => r.id));
          await addLog('Added', `Restored ${savedRecs.length} records (undo bulk delete)`, {records: affectedRecords});
          showToast(`Restored ${savedRecs.length} records`, 'success');
        } catch(e) { showToast('Restore failed: '+e.message, 'error'); }
      });
    } catch(err) { showToast('Delete error: '+err.message, 'error'); }
  };
}

// ── BULK EDIT ─────────────────────────────────────────
const BE_FIELD_MAP = {
  beStatus:'status',bePosted:'postedStatus',beClient:'client',beProduct:'product',beProvider:'provider',
  beRoute:'route',bePrevClient:'prevClient',beRemarks:'remarks',
  beClientOSF:'clientOSF',beClientMRC:'clientMRC',beClientOTRF:'clientOTRF',beClientCF:'clientCF',beClientCPM:'clientCPM',
  beEffDate:'effDate',beActDate:'actDate',bePostedDate:'postedDate',bePostedHour:'postedHour',bePostedMin:'postedMin',
  beArrDate:'arrDate',beProvActDate:'provActDate',beProvOSF:'provOSF',beProvMRC:'provMRC',
  beProvOTRF:'provOTRF',beProvCPM:'provCPM',beTypeSession:'typeSession',beDeactDate:'deactDate'
};
function openBulkEdit() {
  const ids = getCheckedIds(); if (!ids.length) return;
  document.getElementById('beTitle').textContent = `Bulk Edit — ${ids.length} record${ids.length!==1?'s':''}`;
  resetDateMirror('bulk');
  Object.keys(BE_FIELD_MAP).forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  resetFeeSelects(BE_FEE_FIELDS);
  resetDeactSection('bulk');
  document.getElementById('beOv').classList.add('on');
}
function closeBE() { document.getElementById('beOv').classList.remove('on'); }
async function saveBulkEdit() {
  const ids = getCheckedIds(); if (!ids.length) return;

  // ── Bulk deactivation ───
  if (document.getElementById('bDeactSection')?.style.display === 'block') {
    const deactDateVal = document.getElementById('bdDeactDate').value;
    if (!deactDateVal) { showToast('Deactivation date is required.', 'warning'); document.getElementById('bdDeactDate').focus(); return; }
    const requestedBy = document.getElementById('bdRoute').value;
    const bdRemarks = document.getElementById('bdRemarks').value;
    const deactivatedBy = currentUser?.email || 'system';
    const deactivatedAt = new Date().toISOString();
    const deactUpdates = {client:'', status:'Available', deactDate:deactDateVal, route:requestedBy};
    const affectedRecords = ids.map(id => DB.find(r => r.id===id)).filter(Boolean).map(r => bulkChangeSummary(r, ['client','status','deactDate','route'], deactUpdates));
    try {
      const CHUNK = 400;
      for (let i=0; i<ids.length; i+=CHUNK) {
        const b = fdb.batch();
        ids.slice(i,i+CHUNK).forEach(id => {
          const rec = DB.find(r => r.id===id);
          const histEntry = { previousClient: rec?.client||'', activation: activationSnapshot(rec || {}), deactDate: deactDateVal, requestedBy, remarks: bdRemarks, deactivatedBy, deactivatedAt };
          b.update(fdb.collection('inventory').doc(id), {
            client:'', status:'Available', remarks:'', postedStatus:'', postedDate:'',
            postedHour:'', postedMin:'', postedTimeAt:'',
            clientOSF:'', clientMRC:'', clientOTRF:'', clientCF:'', clientCPM:'',
            effDate:'', actDate:'', deactDate: deactDateVal, route: requestedBy,
            prevClient: rec?.client||'',
            deactivationHistory: [...(rec?.deactivationHistory||[]), histEntry],
            updatedBy: deactivatedBy, updatedAt: deactivatedAt
          });
        });
        await b.commit();
      }
      ids.forEach(id => {
        const idx = DB.findIndex(r => r.id===id);
        if (idx>-1) {
          const rec = DB[idx];
          const histEntry = { previousClient: rec.client||'', activation: activationSnapshot(rec), deactDate: deactDateVal, requestedBy, remarks: bdRemarks, deactivatedBy, deactivatedAt };
          DB[idx] = {...rec, client:'', status:'Available', remarks:'', postedStatus:'', postedDate:'', postedHour:'', postedMin:'', postedTimeAt:'', clientOSF:'', clientMRC:'', clientOTRF:'', clientCF:'', clientCPM:'', effDate:'', actDate:'', deactDate: deactDateVal, route: requestedBy, prevClient: rec.client||'', deactivationHistory:[...(rec.deactivationHistory||[]),histEntry], updatedBy:deactivatedBy, updatedAt:deactivatedAt};
        }
      });
      refreshInventoryRecent();
      await addLog('Updated', `Bulk deactivated ${ids.length} record${ids.length!==1?'s':''}`, {records:affectedRecords, fields:['client','status','deactDate','route']});
      propagateChange(ids);
      closeBE();
      showToast(`Deactivated ${ids.length} record${ids.length!==1?'s':''}`, 'success');
    } catch(err) { showToast('Bulk deactivation error: '+err.message, 'error'); }
    return;
  }

  if (bulkEffDateTouched && !bulkActDateTouched) {
    document.getElementById('beActDate').value = document.getElementById('beEffDate').value;
  }
  const updates = {};
  Object.entries(BE_FIELD_MAP).forEach(([elId,field]) => {
    const el = document.getElementById(elId); if (!el) return;
    const v = el.value.trim ? el.value.trim() : el.value;
    if (v) updates[field] = v;
  });
  if (!Object.keys(updates).length) { closeBE(); return; }
  if (updates.postedStatus) updates.postedStatus = canonPostedStatus(updates.postedStatus);
  updates.updatedBy = currentUser?.email||'system';
  updates.updatedAt = new Date().toISOString();

  // Save original field values for undo
  const dataFields = Object.keys(updates).filter(k => k!=='updatedBy' && k!=='updatedAt');
  const savedRecs = ids.map(id => {
    const r = DB.find(x => x.id===id); if (!r) return null;
    const saved = {id};
    dataFields.forEach(f => { saved[f] = r[f] !== undefined ? r[f] : ''; });
    return saved;
  }).filter(Boolean);
  const affectedRecords = ids.map(id => DB.find(r => r.id===id)).filter(Boolean).map(r => bulkChangeSummary(r, dataFields, updates));

  try {
    const CHUNK = 400;
    for (let i=0; i<ids.length; i+=CHUNK) {
      const b = fdb.batch();
      ids.slice(i,i+CHUNK).forEach(id => b.update(fdb.collection('inventory').doc(id), updates));
      await b.commit();
    }
    ids.forEach(id => { const idx=DB.findIndex(r=>r.id===id); if(idx>-1) DB[idx]={...DB[idx],...updates}; });
    refreshInventoryRecent();
    const reseq = await resequencePostingTimes();
    renderTbl();
    await addLog('Updated', `Bulk edited ${ids.length} records: ${dataFields.join(', ')}`, {records: affectedRecords, fields:dataFields});
    propagateChange([...ids, ...reseq]);
    closeBE();
    showUndoToast(`Bulk updated ${ids.length} record${ids.length!==1?'s':''}`, async () => {
      try {
        const CHUNK2 = 400;
        for (let i=0; i<savedRecs.length; i+=CHUNK2) {
          const b = fdb.batch();
          savedRecs.slice(i,i+CHUNK2).forEach(rec => {
            const {id, ...data} = rec;
            b.update(fdb.collection('inventory').doc(id), {
              ...data,
              updatedBy: currentUser?.email||'system',
              updatedAt: new Date().toISOString()
            });
          });
          await b.commit();
        }
        savedRecs.forEach(rec => {
          const idx = DB.findIndex(r => r.id===rec.id);
          if (idx>-1) DB[idx] = {...DB[idx], ...rec};
        });
        refreshInventoryRecent();
        propagateChange(savedRecs.map(r => r.id));
        await addLog('Updated', `Reverted bulk edit of ${savedRecs.length} records (undo)`, {records: reverseBulkChanges(affectedRecords), fields:dataFields});
        showToast(`Reverted ${savedRecs.length} record${savedRecs.length!==1?'s':''}`, 'success');
      } catch(e) { showToast('Revert failed: '+e.message, 'error'); }
    }, 6000, 'Updated');
  } catch(err) { showToast('Bulk edit error: '+err.message, 'error'); }
}

// ── ROLES & RESTRICTIONS ──────────────────────────────
function getSecondApp() {
  if (!_secondApp) _secondApp = firebase.initializeApp(firebaseConfig,'secondary');
  return _secondApp;
}
async function loadUserRole(user) {
  try {
    const doc = await fdb.collection('users').doc(user.uid).get();
    if (doc.exists) {
      currentRole = doc.data().role || 'viewer';
    } else {
      const snap = await fdb.collection('users').get();
      if (snap.empty) {
        currentRole = 'admin';
        await fdb.collection('users').doc(user.uid).set({uid:user.uid,email:user.email,alias:'',role:'admin',addedDate:new Date().toISOString(),addedBy:'system'});
      } else {
        currentRole = 'viewer';
        await fdb.collection('users').doc(user.uid).set({uid:user.uid,email:user.email,alias:'',role:'viewer',addedDate:new Date().toISOString(),addedBy:'system'});
      }
    }
  } catch(e) { console.error('loadUserRole:', e); currentRole='viewer'; }
}
function applyRoleRestrictions() {
  const isViewer = currentRole==='viewer';
  const isAdmin  = currentRole==='admin';
  const nl = document.getElementById('navLogs');
  const na = document.getElementById('navAdmin');
  if (nl) nl.style.display = isViewer ? 'none' : '';
  if (na) na.style.display = isAdmin  ? '' : 'none';
  ['btnAdd','btnUpload','btnExportWrap'].forEach(id => {
    const el = document.getElementById(id); if (el) el.style.display = isViewer ? 'none' : '';
  });
  ['btnBulkEdit','btnBulkDel'].forEach(id => {
    const el = document.getElementById(id); if (el) el.style.display = isViewer ? 'none' : '';
  });
  const se = document.getElementById('btnSpEdit');
  if (se) se.style.display = isViewer ? 'none' : '';
  if (!isAdmin && document.getElementById('page-admin').classList.contains('on')) {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('on'));
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('on'));
    document.getElementById('page-dashboard').classList.add('on');
    document.querySelector('.nav-btn').classList.add('on');
  }
}

// ── USER MANAGEMENT ───────────────────────────────────
async function loadUsers() {
  try {
    const snap = await fdb.collection('users').get();
    USERS = snap.docs.map(d => ({...d.data(), uid:d.id}));
    renderUsers();
  } catch(e) { console.error('loadUsers:', e); }
}
function renderUsers() {
  const tbody = document.getElementById('umBody'); if (!tbody) return;
  if (!USERS.length) { tbody.innerHTML='<tr><td colspan="7" style="text-align:center;color:var(--t3);padding:20px">No users found.</td></tr>'; return; }
  const self = currentUser?.uid;
  tbody.innerHTML = USERS.map((u,i) => `
    <tr>
      <td class="row-num">${i+1}</td>
      <td>${esc(u.email||'—')}</td>
      <td style="font-family:'DM Mono',monospace;color:var(--t3);font-size:12px">••••••••</td>
      <td>${esc(u.alias||'—')}</td>
      <td>${roleBadge(u.role)}</td>
      <td style="font-size:12px;color:var(--t2)">${u.addedDate?new Date(u.addedDate).toLocaleDateString():'—'}</td>
      <td>
        <div class="act-btns">
          <button class="act-btn" title="Edit" onclick="openEditUser('${esc(u.uid)}')">✎</button>
          ${u.uid===self?'<span style="color:var(--t3);padding:3px 5px;font-size:12px" title="Cannot delete own account">—</span>':`<button class="act-btn del" title="Delete" onclick="deleteUser('${esc(u.uid)}')">⊗</button>`}
        </div>
      </td>
    </tr>`).join('');
}
function openAddUser() {
  umEditUid=null;
  document.getElementById('umTitle').textContent='Add User';
  document.getElementById('umEmail').value=''; document.getElementById('umEmail').disabled=false;
  document.getElementById('umPass').value='';
  document.getElementById('umPassFg').style.display=''; document.getElementById('umResetFg').style.display='none';
  document.getElementById('umAlias').value=''; document.getElementById('umAliasFg').style.display='';
  document.getElementById('umRole').value='viewer'; document.getElementById('umRole').disabled=false;
  document.getElementById('umSelfNote').style.display='none';
  document.getElementById('umOv').classList.add('on');
}
function openEditUser(uid) {
  const u = USERS.find(x => x.uid===uid); if (!u) return;
  umEditUid=uid;
  document.getElementById('umTitle').textContent='Edit User';
  document.getElementById('umEmail').value=u.email; document.getElementById('umEmail').disabled=true;
  document.getElementById('umPassFg').style.display='none'; document.getElementById('umResetFg').style.display='';
  document.getElementById('umAlias').value=u.alias||''; document.getElementById('umAliasFg').style.display='';
  const isSelf = currentUser && currentUser.uid===uid;
  document.getElementById('umRole').value=u.role||'viewer'; document.getElementById('umRole').disabled=isSelf;
  document.getElementById('umSelfNote').style.display=isSelf?'':'none';
  document.getElementById('umOv').classList.add('on');
}
function closeUM() { document.getElementById('umOv').classList.remove('on'); document.getElementById('umRole').disabled=false; umEditUid=null; }
async function sendUserResetEmail() {
  const email = document.getElementById('umEmail').value; if (!email) return;
  try { await fauth.sendPasswordResetEmail(email); showToast(`Password reset email sent to ${email}`,'info'); }
  catch(e) { showToast('Error: '+e.message,'error'); }
}
async function saveUser() {
  const email = document.getElementById('umEmail').value.trim();
  const alias = document.getElementById('umAlias').value.trim();
  const role  = document.getElementById('umRole').value;
  if (!umEditUid) {
    const pass = document.getElementById('umPass').value;
    if (!email||!pass) { showToast('Email and password are required.','warning'); return; }
    if (pass.length<6) { showToast('Password must be at least 6 characters.','warning'); return; }
    try {
      const auth2 = getSecondApp().auth();
      const cred  = await auth2.createUserWithEmailAndPassword(email,pass);
      const uid   = cred.user.uid;
      await auth2.signOut();
      const userData = {uid,email,alias,role,addedDate:new Date().toISOString(),addedBy:currentUser?.email||'system'};
      await fdb.collection('users').doc(uid).set(userData);
      USERS.push(userData); renderUsers();
      await addLog('Added',`Added user ${email} (${role})`);
      showToast(`User ${email} created`,'success'); closeUM();
    } catch(e) { showToast('Error creating user: '+e.message,'error'); }
  } else {
    const updates = {alias};
    if (!document.getElementById('umRole').disabled) updates.role=role;
    try {
      await fdb.collection('users').doc(umEditUid).update(updates);
      const idx = USERS.findIndex(u => u.uid===umEditUid);
      if (idx>-1) USERS[idx]={...USERS[idx],...updates};
      renderUsers();
      await addLog('Updated',`Updated user ${email}`);
      showToast(`User ${email} updated`,'success'); closeUM();
    } catch(e) { showToast('Error updating user: '+e.message,'error'); }
  }
}
async function deleteUser(uid) {
  const u = USERS.find(x => x.uid===uid); if (!u) return;
  if (currentUser && currentUser.uid===uid) { showToast('You cannot delete your own account.','warning'); return; }
  document.getElementById('delRecTitle').textContent = 'Remove this user?';
  document.getElementById('delRecInfo').innerHTML = `
    <div><span style="color:var(--t2)">Email:</span> <strong>${esc(u.email)}</strong></div>
    <div><span style="color:var(--t2)">Role:</span> ${esc(u.role)}</div>
    <div style="margin-top:8px;font-size:11.5px;color:var(--t3)">This removes their access. To fully delete from Firebase Auth, use the Firebase Console.</div>`.trim();
  document.getElementById('delRecOv').classList.add('on');
  const btn   = document.getElementById('delRecConfirmBtn');
  const fresh = btn.cloneNode(true); btn.replaceWith(fresh);
  fresh.textContent = 'Remove User';
  fresh.onclick = async () => {
    document.getElementById('delRecOv').classList.remove('on');
    try {
      await fdb.collection('users').doc(uid).delete();
      USERS = USERS.filter(x => x.uid!==uid); renderUsers();
      await addLog('Deleted',`Removed user ${u.email}`);
      showToast(`User ${u.email} removed`,'info');
    } catch(e) { showToast('Error: '+e.message,'error'); }
  };
}

// ── SELECTION MANAGEMENT ──────────────────────────────
async function loadSelections() {
  try {
    const types = ['clients','products','providers','routes'];
    const snaps = await Promise.all(types.map(t => fdb.collection('selections').doc(t).get()));
    snaps.forEach((s,i) => { SELECTIONS[types[i]] = s.exists ? (s.data().items||[]) : []; });
    populateDropdowns(); renderSelections();
  } catch(e) { console.error('loadSelections:', e); }
}
function renderSelections() {
  ['clients','products','providers','routes'].forEach(type => {
    const el = document.getElementById('selItems-'+type); if (!el) return;
    const sorted = [...SELECTIONS[type]].sort();
    el.innerHTML = sorted.length
      ? sorted.map(v => `<span class="sel-chip"><span>${esc(v)}</span><button class="sel-chip-del" title="Remove" data-type="${esc(type)}" data-val="${esc(v)}" onclick="removeSelItemBtn(this)">✕</button></span>`).join('')
      : '<span style="font-size:12px;color:var(--t3)">No items added.</span>';
  });
}
function populateDropdowns() {
  [['clients','fClient','All Clients'],['products','fProduct','All Products'],['providers','fProvider','All Providers']].forEach(([type,id,lbl]) => {
    const el = document.getElementById(id); if (!el) return;
    const val = el.value;
    el.innerHTML = `<option value="">${lbl}</option>` + [...SELECTIONS[type]].sort().map(v => `<option>${esc(v)}</option>`).join('');
    el.value = val;
  });
  const modalSelMap = [['clients','mClient','— select client —'],['products','mProduct','— select product —'],['providers','mProvider','— select provider —'],['routes','mRoute','— select —'],['clients','mPrevClient','— select prev client —'],['routes','dRoute','— select —'],['routes','bdRoute','— select —']];
  modalSelMap.forEach(([type,id,lbl]) => {
    const el = document.getElementById(id); if (!el) return;
    const val = el.value;
    el.innerHTML = `<option value="">${lbl}</option>` + [...SELECTIONS[type]].sort().map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
    if (val) setSelectVal(el, val);
  });
  const beSelMap = [['clients','beClient'],['products','beProduct'],['providers','beProvider'],['routes','beRoute'],['clients','bePrevClient']];
  beSelMap.forEach(([type,id]) => {
    const el = document.getElementById(id); if (!el) return;
    el.innerHTML = `<option value="">— keep existing —</option>` + [...SELECTIONS[type]].sort().map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  });
}
function toggleSel(type) {
  document.getElementById('selBody-'+type).classList.toggle('on');
  document.getElementById('selArrow-'+type).classList.toggle('on');
}
async function addSelItem(type) {
  const inp = document.getElementById('selInput-'+type);
  const raw = inp.value.trim(); if (!raw) return;
  const val = raw.toUpperCase();
  if (SELECTIONS[type].includes(val)) { showToast('Item already exists.','warning'); return; }
  try {
    await fdb.collection('selections').doc(type).set({items:firebase.firestore.FieldValue.arrayUnion(val)},{merge:true});
    SELECTIONS[type].push(val); inp.value='';
    populateDropdowns(); renderSelections();
    await addLog('Updated', `Added "${val}" to ${type}`);
  } catch(e) { showToast('Error: '+e.message,'error'); }
}
function removeSelItemBtn(btn) { removeSelItem(btn.dataset.type, btn.dataset.val); }
function removeSelItem(type, val) {
  document.getElementById('rmSelMsg').textContent = `"${val}"`;
  const ov  = document.getElementById('rmSelOv');
  ov.classList.add('on');
  const btn    = document.getElementById('rmSelConfirmBtn');
  const newBtn = btn.cloneNode(true); btn.parentNode.replaceChild(newBtn, btn);
  newBtn.onclick = async function() {
    ov.classList.remove('on');
    try {
      await fdb.collection('selections').doc(type).update({items:firebase.firestore.FieldValue.arrayRemove(val)});
      SELECTIONS[type] = SELECTIONS[type].filter(v => v!==val);
      populateDropdowns(); renderSelections();
      await addLog('Updated', `Removed "${val}" from ${type}`);
    } catch(e) { showToast('Error: '+e.message,'error'); }
  };
}
async function handleSelCSV(e, type) {
  const f = e.target.files[0]; if (!f) return;
  const text  = await readCSVText(f);
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  const rawItems = lines.slice(1).map(l => parseCSVLine(l)[0]?.trim()).filter(Boolean);
  const items    = rawItems.map(v => v.toUpperCase());
  const toAdd    = items.filter(v => !SELECTIONS[type].includes(v));
  if (!toAdd.length) { showToast('No new items found.','warning'); e.target.value=''; return; }
  try {
    await fdb.collection('selections').doc(type).set({items:firebase.firestore.FieldValue.arrayUnion(...toAdd)},{merge:true});
    SELECTIONS[type].push(...toAdd);
    populateDropdowns(); renderSelections();
    await addLog('Updated', `CSV upload: added ${toAdd.length} item(s) to ${type}`);
    showToast(`Added ${toAdd.length} item${toAdd.length!==1?'s':''} to ${type}.`,'success');
  } catch(err) { showToast('Error: '+err.message,'error'); }
  e.target.value='';
}
function dlSelSample(type) {
  const labels  = {clients:'Client',products:'Product',providers:'Provider',routes:'Route Requested by'};
  const samples = {clients:'TOKU',products:'DID Local',providers:'Twilio',routes:'Katherine Serrano'};
  dlCSV([[labels[type]],[samples[type]]], `sample_${type}.csv`);
}
async function delAllSelItems(type) {
  const labels = {clients:'Client',products:'Product',providers:'Provider',routes:'Route Requested by'};
  const count  = SELECTIONS[type].length;
  if (!count) { showToast(`No ${labels[type]} items to delete.`, 'warning'); return; }
  document.getElementById('delRecTitle').textContent = `Delete all ${labels[type]} items?`;
  document.getElementById('delRecInfo').innerHTML = `
    <div><span style="color:var(--t2)">Items to delete:</span> <strong>${count}</strong></div>
    <div style="margin-top:4px;font-size:11.5px;color:var(--t3)">All ${count} item${count!==1?'s':''} will be permanently removed from the ${labels[type]} selection list.</div>`.trim();
  document.getElementById('delRecOv').classList.add('on');
  const btn   = document.getElementById('delRecConfirmBtn');
  const fresh = btn.cloneNode(true); btn.replaceWith(fresh);
  fresh.textContent = 'Delete All';
  fresh.onclick = async () => {
    document.getElementById('delRecOv').classList.remove('on');
    try {
      await fdb.collection('selections').doc(type).set({items:[]});
      SELECTIONS[type] = [];
      populateDropdowns(); renderSelections();
      await addLog('Updated', `Deleted all ${count} item${count!==1?'s':''} from ${type}`);
      showToast(`Deleted all ${count} item${count!==1?'s':''} from ${labels[type]}.`, 'success');
    } catch(e) { showToast('Error: '+e.message, 'error'); }
  };
}

// ── EVENT LISTENERS ───────────────────────────────────
document.getElementById('beOv').addEventListener('click', function(e) { if(e.target===this) closeBE(); });
document.getElementById('umOv').addEventListener('click', function(e) { if(e.target===this) closeUM(); });
document.getElementById('logRecordsOv').addEventListener('click', function(e) { if(e.target===this) closeLogRecords(); });
document.addEventListener('click', () => closeExportMenu());
window.addEventListener('resize', () => {
  if (document.getElementById('page-dashboard').classList.contains('on')) drawChart();
});

// ── KEYBOARD SHORTCUTS ────────────────────────────────
document.addEventListener('keydown', e => {
  const typing = document.activeElement && ['INPUT','SELECT','TEXTAREA'].includes(document.activeElement.tagName);
  if (e.key === 'Escape') {
    if (document.getElementById('moOv').classList.contains('on'))        { closeMo();  return; }
    if (document.getElementById('beOv').classList.contains('on'))        { closeBE();  return; }
    if (document.getElementById('umOv').classList.contains('on'))        { closeUM();  return; }
    if (document.getElementById('logRecordsOv').classList.contains('on')){ closeLogRecords(); return; }
    if (document.getElementById('sp').classList.contains('on'))          { closeSP();  return; }
    if (document.getElementById('exportMenu').classList.contains('on'))  { closeExportMenu(); return; }
  }
  if (typing) return;
  if (e.key === '/' || (e.ctrlKey && e.key === 'k')) {
    if (document.getElementById('page-inventory').classList.contains('on')) {
      e.preventDefault(); document.getElementById('fSearch').focus();
    }
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 'n') {
    if (document.getElementById('page-inventory').classList.contains('on') && currentRole !== 'viewer') {
      e.preventDefault(); openAdd();
    }
  }
});

// ── INIT ──────────────────────────────────────────────
initEL();
initDateMirrors();
initPostedTimeSelects();
loadPinned();
if (EL.pgSize) EL.pgSize.value = '50';
renderDash(); renderTbl(); renderLogs();
