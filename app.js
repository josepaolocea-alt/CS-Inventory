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
const ACT_LABELS  = {Added:'b-active',Updated:'b-reserved',Deleted:'b-inactive','CSV Upload':'b-available',Exported:'b-available',Login:'b-available'};
const CSV_HEADERS = ['Client','Product','Number','Status','Remarks','Posted Status','Posted Date','Client OSF','Client MRC','Client OTRF','Client Channel Fee','Client CPM','Effective Date','Activated Date','Provider','Arrival Date','Provider Activation Date','Provider OSF','Provider MRC','Provider OTRF','Provider CPM','Type / Session','Route Request by','Deactivation Date','Previous Client'];
const CSV_FIELD_MAP = {'Client':'client','Product':'product','Number':'number','Status':'status','Remarks':'remarks','Posted Status':'postedStatus','Posted Date':'postedDate','Client OSF':'clientOSF','Client MRC':'clientMRC','Client OTRF':'clientOTRF','Client Channel Fee':'clientCF','Client CPM':'clientCPM','Effective Date':'effDate','Activated Date':'actDate','Provider':'provider','Arrival Date':'arrDate','Provider Activation Date':'provActDate','Provider OSF':'provOSF','Provider MRC':'provMRC','Provider OTRF':'provOTRF','Provider CPM':'provCPM','Type / Session':'typeSession','Route Request by':'route','Deactivation Date':'deactDate','Previous Client':'prevClient'};
const DATE_FIELDS = new Set(['mPostedDate','mEffDate','mActDate','mArrDate','mProvActDate','mDeactDate']);
const VALID_STATUSES = new Set(['Active','Available','Reserved','Inactive','']);
const DATE_CSV_FIELDS = ['postedDate','effDate','actDate','arrDate','provActDate','deactDate'];

// ── UTILITIES ─────────────────────────────────────────
function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
const fmt = iso => iso ? iso.replace(/(\d{4})-(\d{2})-(\d{2})/,'$2/$3/$1') : '—';
function sanitizeDate(v) {
  if (!v) return '';
  const s = String(v);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m2) return `${m2[3]}-${m2[1].padStart(2,'0')}-${m2[2].padStart(2,'0')}`;
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
function bclass(s) { return {Active:'b-active',Available:'b-available',Reserved:'b-reserved',Inactive:'b-inactive'}[s]||''; }
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
  ['invBody','tInfo','pgInfo','pgPrev','pgNext','pgSize','selAll','selBar','selCount',
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
let curRec=null, editId=null, moreOpen=false;
let _editUpdatedAt=null;
let currentUser=null, currentRole='viewer';
let USERS=[];
let SELECTIONS={clients:[],products:[],providers:[],routes:[]};
let umEditUid=null, _secondApp=null;

// ── THEME INIT (runs immediately on script load) ──────
(function() {
  const t = localStorage.getItem('cs-inv-theme');
  if (t === 'light') {
    document.documentElement.removeAttribute('data-dark');
    const btn = document.getElementById('themeBtn');
    if (btn) btn.textContent = '◑';
  }
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

function showUndoToast(msg, onUndo, duration=6000) {
  const c = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = 'toast t-info';
  t.innerHTML = `<span class="toast-icon">ℹ</span><div class="toast-body"><div class="toast-title">Deleted</div><div class="toast-msg">${esc(msg)}</div></div><button class="toast-undo">↩ Undo</button><button class="toast-close" onclick="dismissToast(this.closest('.toast'))">✕</button>`;
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
  document.getElementById('themeBtn').textContent = dark ? '◑' : '◐';
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
  } else {
    currentUser = null; currentRole = 'viewer';
    DB=[]; LOGS=[]; fd=[]; fl=[]; recentViewed=[];
    USERS=[]; SELECTIONS={clients:[],products:[],providers:[],routes:[]};
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

// ── FIRESTORE LOAD ────────────────────────────────────
async function loadInventory() {
  try {
    const snap = await fdb.collection('inventory').orderBy('client').get();
    DB = snap.docs.map(d => ({...d.data(), id:d.id}));
    fd = [...DB]; renderTbl(); renderDash();
  } catch(e) { console.error('loadInventory:', e); }
}
async function syncData() {
  const btn = document.getElementById('syncBtn');
  btn.classList.add('syncing');
  try { await Promise.all([loadInventory(), loadLogs()]); }
  finally { btn.classList.remove('syncing'); }
}
async function loadLogs() {
  try {
    const snap = await fdb.collection('logs').orderBy('datetime','desc').limit(500).get();
    LOGS = snap.docs.map(d => ({...d.data(), id:d.id}));
    fl = [...LOGS]; renderLogs();
  } catch(e) { console.error('loadLogs:', e); }
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
  if (tab==='admin')     loadUsers();
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
function applyF() {
  const s  = document.getElementById('fSearch').value.toLowerCase();
  const cl = document.getElementById('fClient').value;
  const st = document.getElementById('fStatus').value;
  const pr = document.getElementById('fProduct').value;
  const pv = document.getElementById('fProvider').value;
  const df = document.getElementById('fDateFrom').value;
  const dt = document.getElementById('fDateTo').value;
  const sRe = s ? wildcardToRegex(s) : null;
  fd = DB.filter(r => {
    if (sRe && !sRe.test(Object.values(r).join(' ').toLowerCase())) return false;
    if (cl && r.client!==cl)   return false;
    if (st && r.status!==st)   return false;
    if (pr && r.product!==pr)  return false;
    if (pv && r.provider!==pv) return false;
    if (df && r.actDate && r.actDate<df) return false;
    if (dt && r.actDate && r.actDate>dt) return false;
    return true;
  });
  pg=1; renderTbl();
}
function clearF() {
  ['fSearch','fDateFrom','fDateTo'].forEach(id => document.getElementById(id).value='');
  ['fClient','fStatus','fProduct','fProvider'].forEach(id => document.getElementById(id).value='');
  fd=[...DB]; pg=1; renderTbl();
}
function toggleMore() {
  moreOpen = !moreOpen;
  document.getElementById('moreRow').classList.toggle('on', moreOpen);
  document.getElementById('moreBtn').textContent = moreOpen ? 'Less ▴' : 'More ▾';
}

// ── SORT ──────────────────────────────────────────────
const colIdx = {client:2,product:3,number:4,status:5,remarks:6};
function sortBy(col) {
  if (sortCol===col) sortDir*=-1; else { sortCol=col; sortDir=1; }
  fd.sort((a,b) => (a[col]||'').localeCompare(b[col]||'')*sortDir);
  document.querySelectorAll('#invTbl th').forEach(th => th.classList.remove('asc','desc'));
  const ths = [...document.querySelectorAll('#invTbl th')];
  if (colIdx[col]) ths[colIdx[col]].classList.add(sortDir===1?'asc':'desc');
  renderTbl();
}

// ── RENDER TABLE ──────────────────────────────────────
function renderTbl() {
  const sz = parseInt(EL.pgSize?.value || 100);
  const s = (pg-1)*sz, e = s+sz, total = fd.length, tp = Math.ceil(total/sz)||1;
  if (EL.tInfo)    EL.tInfo.textContent    = `Showing ${Math.min(s+1,total)}–${Math.min(e,total)} of ${total} records`;
  if (EL.pgInfo)   EL.pgInfo.textContent   = `Page ${pg} of ${tp}`;
  if (EL.pgPrev)   EL.pgPrev.disabled      = pg<=1;
  if (EL.pgNext)   EL.pgNext.disabled      = pg>=tp;
  if (EL.selAll)   EL.selAll.checked       = false;
  if (EL.selBar)   EL.selBar.classList.remove('on');
  if (EL.invBody)  EL.invBody.innerHTML    = fd.slice(s,e).map((r,i) => `
    <tr onclick="rowClick(event,'${esc(r.id)}')">
      <td onclick="event.stopPropagation()"><input type="checkbox" class="rcb" data-id="${esc(r.id)}" onchange="updateSelBar()"></td>
      <td class="row-num">${s+i+1}</td>
      <td>${esc(r.client)}</td>
      <td>${esc(r.product)}</td>
      <td class="num-cell" style="color:var(--accent);font-weight:500">${esc(r.number)}</td>
      <td><span class="badge ${bclass(r.status)}">${esc(r.status)}</span></td>
      <td>${esc(r.remarks)}</td>
      <td onclick="event.stopPropagation()">
        <div class="act-btns">
          ${currentRole!=='viewer'?`<button class="act-btn" title="Edit" onclick="openEditById('${esc(r.id)}')">✎</button><button class="act-btn del" title="Delete" onclick="delRec('${esc(r.id)}')">⊗</button>`:''}
        </div>
      </td>
    </tr>`).join('');
}
function changePg(d) {
  const sz = parseInt(EL.pgSize?.value || 100);
  const tp = Math.ceil(fd.length/sz)||1;
  pg = Math.max(1, Math.min(pg+d, tp)); renderTbl();
}
function selAllRows(cb) { document.querySelectorAll('.rcb').forEach(c => c.checked=cb.checked); updateSelBar(); }
function getCheckedIds() { return [...document.querySelectorAll('.rcb:checked')].map(c => c.dataset.id); }
function updateSelBar() {
  const ids   = getCheckedIds();
  const total = document.querySelectorAll('.rcb').length;
  if (EL.selCount) EL.selCount.textContent = `${ids.length} row${ids.length!==1?'s':''} selected`;
  if (EL.selBar)   EL.selBar.classList.toggle('on', ids.length>0);
  if (EL.selAll) {
    EL.selAll.indeterminate = ids.length>0 && ids.length<total;
    if (total>0) EL.selAll.checked = ids.length===total;
  }
}
function rowClick(e, id) {
  if (e.target.tagName==='INPUT') return;
  openSP(id);
  const r = DB.find(x => x.id===id);
  if (r && !recentViewed.find(x => x.id===id)) { recentViewed.unshift(r); recentViewed=recentViewed.slice(0,6); }
}

// ── SIDE PANEL ────────────────────────────────────────
function openSP(id) {
  const r = DB.find(x => x.id===id); if (!r) return;
  curRec = r;
  document.getElementById('spTitle').textContent = r.number;
  document.getElementById('spBody').innerHTML = `
    <div class="ds"><div class="ds-title">Client Information</div>
      ${dr('Client',r.client)}${dr('Product',r.product)}${dr('Number',r.number)}
      ${drHTML('Status',`<span class="badge ${bclass(r.status)}">${esc(r.status)}</span>`)}
      ${dr('Remarks',r.remarks||'—')}${dr('Posted Status',r.postedStatus||'—')}
      ${dr('Posted Date',fmt(r.postedDate))}${dr('Client OSF','$'+(r.clientOSF||'—'))}
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
    <div class="ds"><div class="ds-title">Meta</div>
      ${dr('Created by',r.createdBy||'—')}${dr('Updated by',r.updatedBy||'—')}
    </div>`;
  document.getElementById('spOv').classList.add('on');
  document.getElementById('sp').classList.add('on');
}
function closeSP() {
  document.getElementById('spOv').classList.remove('on');
  document.getElementById('sp').classList.remove('on');
  curRec = null;
}

// ── MODAL ─────────────────────────────────────────────
const mMap = {
  mClient:'client',mProduct:'product',mNumber:'number',mStatus:'status',mRemarks:'remarks',
  mPosted:'postedStatus',mPostedDate:'postedDate',mClientOSF:'clientOSF',mClientMRC:'clientMRC',
  mClientOTRF:'clientOTRF',mClientCF:'clientCF',mClientCPM:'clientCPM',mEffDate:'effDate',
  mActDate:'actDate',mProvider:'provider',mArrDate:'arrDate',mProvActDate:'provActDate',
  mProvOSF:'provOSF',mProvMRC:'provMRC',mProvOTRF:'provOTRF',mProvCPM:'provCPM',
  mTypeSession:'typeSession',mRoute:'route',mDeactDate:'deactDate',mPrevClient:'prevClient'
};
function setSelectVal(el, val) {
  el.value = val;
  if (el.tagName==='SELECT' && el.value!==val) {
    const opt = document.createElement('option'); opt.value=val; opt.textContent=val;
    el.appendChild(opt); el.value=val;
  }
}
function clearMo() {
  Object.keys(mMap).forEach(id => {
    const el = document.getElementById(id); if (el) el.value = id==='mStatus'?'Available':id==='mPosted'?'No':'';
  });
  document.getElementById('mNumber')?.classList.remove('err');
}
function fillMo(r) {
  _editUpdatedAt = r.updatedAt || null;
  Object.entries(mMap).forEach(([id,key]) => {
    const el = document.getElementById(id); if (!el || r[key]===undefined) return;
    const v = DATE_FIELDS.has(id) ? sanitizeDate(r[key]) : r[key];
    if (el.tagName==='SELECT') setSelectVal(el,v); else el.value=v;
  });
}
function openAdd() {
  editId=null; document.getElementById('moTitle').textContent='Add Number';
  clearMo(); document.getElementById('moOv').classList.add('on');
}
function openEdit() { if (curRec) openEditById(curRec.id); }
function openEditById(id) {
  const r = DB.find(x => x.id===id); if (!r) return;
  editId=id; document.getElementById('moTitle').textContent='Edit Number';
  fillMo(r); document.getElementById('moOv').classList.add('on');
}
function closeMo() { document.getElementById('moOv').classList.remove('on'); }

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

  const nd = {};
  Object.entries(mMap).forEach(([id,key]) => { const el=document.getElementById(id); if(el) nd[key]=el.value; });
  nd.updatedBy = currentUser?.email || 'system';
  nd.updatedAt = new Date().toISOString();
  try {
    if (editId) {
      // ── Concurrent edit detection ───
      try {
        const snap = await fdb.collection('inventory').doc(editId).get();
        if (snap.exists && snap.data().updatedAt && _editUpdatedAt && snap.data().updatedAt !== _editUpdatedAt) {
          showToast('This record was modified by another user. Please reload and try again.', 'error', 7000);
          return;
        }
      } catch(e) { /* proceed on check failure */ }
      await fdb.collection('inventory').doc(editId).update(nd);
      const idx = DB.findIndex(r => r.id===editId);
      if (idx>-1) DB[idx] = {...DB[idx], ...nd};
      await addLog('Updated', `Updated number ${nd.number}`);
    } else {
      nd.createdBy = currentUser?.email || 'system';
      nd.createdAt = new Date().toISOString();
      const ref = await fdb.collection('inventory').add(nd);
      nd.id = ref.id; DB.push(nd);
      await addLog('Added', `Added number ${nd.number}`);
    }
    fd=[...DB]; pg=1; renderTbl(); closeMo();
    showToast(editId ? `Updated ${nd.number}` : `Added ${nd.number}`, 'success');
    if (editId) openSP(editId);
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
      if (r) await addLog('Deleted', `Deleted number ${r.number}`);
      renderTbl(); closeSP();
      if (savedRec) {
        showUndoToast(`Deleted ${savedRec.number}`, async () => {
          try {
            const {id:rid, ...data} = savedRec;
            await fdb.collection('inventory').doc(rid).set({...data, id:rid});
            DB.push(savedRec); fd = [...DB];
            renderTbl(); renderDash();
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
  const text  = await f.text();
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
    if (nd.status && !VALID_STATUSES.has(nd.status)) {
      warnings.push(`Row ${i+1}: invalid status "${nd.status}" — cleared`);
      nd.status = '';
    }
    DATE_CSV_FIELDS.forEach(field => {
      if (nd[field]) {
        const clean = sanitizeDate(nd[field]);
        if (!clean) { warnings.push(`Row ${i+1}: invalid date in ${field} — cleared`); nd[field]=''; }
        else nd[field] = clean;
      }
    });
    nd.updatedBy = currentUser?.email||'system';
    nd.updatedAt = new Date().toISOString();
    const existing = DB.find(r => r.number===nd.number);
    if (existing) {
      ops.push({type:'update', ref:fdb.collection('inventory').doc(existing.id), data:nd, id:existing.id});
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
    fd=[...DB]; pg=1; renderTbl(); renderDash();
    await addLog('CSV Upload', `"${f.name}": ${added} added, ${updated} updated`);
    showToast(`Upload complete — ${added} added, ${updated} updated`, 'success');
  } catch(err) { showToast('Import error: '+err.message, 'error'); }
  e.target.value='';
}

function dlSample() {
  const row = ['TOKU','DID Local','+15550001234','Active','Sample','No','','100.00','50.00','25.00','10.00','0.0050','2024-01-01','2024-01-15','Twilio','2023-12-15','2024-01-15','80.00','40.00','20.00','0.0040','SIP','Katherine Serrano','','DIDLOGIC'];
  dlCSV([CSV_HEADERS,row], 'sample_inventory.csv');
  addLog('Exported','Downloaded sample CSV template');
}

function exportAll() {
  const rows = DB.map(r => [r.client,r.product,r.number,r.status,r.remarks,r.postedStatus,r.postedDate,r.clientOSF,r.clientMRC,r.clientOTRF,r.clientCF,r.clientCPM,r.effDate,r.actDate,r.provider,r.arrDate,r.provActDate,r.provOSF,r.provMRC,r.provOTRF,r.provCPM,r.typeSession,r.route,r.deactDate,r.prevClient]);
  dlCSV([CSV_HEADERS,...rows], 'inventory_export.csv');
  closeExportMenu();
  addLog('Exported', `Exported ${DB.length} records to CSV`);
}

function exportExcel() {
  if (typeof XLSX === 'undefined') { showToast('Excel library not loaded yet. Try again in a moment.','warning'); return; }
  const rows = DB.map(r => ({'Client':r.client,'Product':r.product,'Number':r.number,'Status':r.status,'Remarks':r.remarks,'Posted Status':r.postedStatus,'Posted Date':r.postedDate,'Client OSF':r.clientOSF,'Client MRC':r.clientMRC,'Client OTRF':r.clientOTRF,'Client Channel Fee':r.clientCF,'Client CPM':r.clientCPM,'Effective Date':r.effDate,'Activated Date':r.actDate,'Provider':r.provider,'Arrival Date':r.arrDate,'Provider Activation Date':r.provActDate,'Provider OSF':r.provOSF,'Provider MRC':r.provMRC,'Provider OTRF':r.provOTRF,'Provider CPM':r.provCPM,'Type / Session':r.typeSession,'Route Request by':r.route,'Deactivation Date':r.deactDate,'Previous Client':r.prevClient}));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Inventory');
  XLSX.writeFile(wb, 'inventory_export.xlsx');
  closeExportMenu();
  addLog('Exported', `Exported ${DB.length} records to Excel`);
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

function dlCSV(rows, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([rows.map(r => r.map(c => `"${String(c||'').replace(/"/g,'""')}"`).join(',')).join('\n')],{type:'text/csv'}));
  a.download = name; a.click();
}

function toggleExportMenu(e) {
  e.stopPropagation();
  document.getElementById('exportMenu').classList.toggle('on');
}
function closeExportMenu() {
  document.getElementById('exportMenu')?.classList.remove('on');
}

// ── LOGS ──────────────────────────────────────────────
async function addLog(action, details) {
  const log = {datetime:new Date().toISOString(), user:currentUser?.email||'system', action, details};
  try {
    const ref = await fdb.collection('logs').add(log);
    log.id = ref.id; LOGS.unshift(log); fl=[...LOGS]; renderLogs();
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
function clearAllLogs() {
  const ov = document.getElementById('clearLogsOv');
  document.getElementById('clearLogsConfirmBtn').onclick = async function() {
    ov.classList.remove('on');
    try {
      const snap = await fdb.collection('logs').get();
      const batch = fdb.batch();
      snap.docs.forEach(d => batch.delete(d.ref));
      await batch.commit();
      LOGS=[]; fl=[]; renderLogs();
      showToast('All logs cleared.', 'info');
    } catch(e) { showToast('Error: '+e.message, 'error'); }
  };
  ov.classList.add('on');
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
      <td>${esc(r.details)}</td>
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
  const data = rows.map(r => [r.client,r.product,r.number,r.status,r.remarks,r.postedStatus,r.postedDate,r.clientOSF,r.clientMRC,r.clientOTRF,r.clientCF,r.clientCPM,r.effDate,r.actDate,r.provider,r.arrDate,r.provActDate,r.provOSF,r.provMRC,r.provOTRF,r.provCPM,r.typeSession,r.route,r.deactDate,r.prevClient]);
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
    try {
      const CHUNK = 400;
      for (let i=0; i<ids.length; i+=CHUNK) {
        const b = fdb.batch();
        ids.slice(i,i+CHUNK).forEach(id => b.delete(fdb.collection('inventory').doc(id)));
        await b.commit();
      }
      const nums = ids.map(id => DB.find(r => r.id===id)?.number).filter(Boolean).join(', ');
      DB = DB.filter(r => !ids.includes(r.id)); fd = fd.filter(r => !ids.includes(r.id));
      await addLog('Deleted', `Bulk deleted ${ids.length} records: ${nums}`);
      renderTbl(); closeSP();
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
          fd=[...DB]; renderTbl(); renderDash();
          await addLog('Added', `Restored ${savedRecs.length} records (undo bulk delete)`);
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
  beEffDate:'effDate',beActDate:'actDate',bePostedDate:'postedDate',
  beArrDate:'arrDate',beProvActDate:'provActDate',beProvOSF:'provOSF',beProvMRC:'provMRC',
  beProvOTRF:'provOTRF',beProvCPM:'provCPM',beTypeSession:'typeSession',beDeactDate:'deactDate'
};
function openBulkEdit() {
  const ids = getCheckedIds(); if (!ids.length) return;
  document.getElementById('beTitle').textContent = `Bulk Edit — ${ids.length} record${ids.length!==1?'s':''}`;
  Object.keys(BE_FIELD_MAP).forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  document.getElementById('beOv').classList.add('on');
}
function closeBE() { document.getElementById('beOv').classList.remove('on'); }
async function saveBulkEdit() {
  const ids = getCheckedIds(); if (!ids.length) return;
  const updates = {};
  Object.entries(BE_FIELD_MAP).forEach(([elId,field]) => {
    const el = document.getElementById(elId); if (!el) return;
    const v = el.value.trim ? el.value.trim() : el.value;
    if (v) updates[field] = v;
  });
  if (!Object.keys(updates).length) { closeBE(); return; }
  updates.updatedBy = currentUser?.email||'system';
  updates.updatedAt = new Date().toISOString();
  try {
    const CHUNK = 400;
    for (let i=0; i<ids.length; i+=CHUNK) {
      const b = fdb.batch();
      ids.slice(i,i+CHUNK).forEach(id => b.update(fdb.collection('inventory').doc(id), updates));
      await b.commit();
    }
    ids.forEach(id => { const idx=DB.findIndex(r=>r.id===id); if(idx>-1) DB[idx]={...DB[idx],...updates}; });
    fd=[...DB]; pg=1; renderTbl(); renderDash();
    await addLog('Updated', `Bulk edited ${ids.length} records: ${Object.keys(updates).filter(k=>k!=='updatedBy'&&k!=='updatedAt').join(', ')}`);
    closeBE();
    showToast(`Bulk updated ${ids.length} records`, 'success');
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
  const modalSelMap = [['clients','mClient','— select client —'],['products','mProduct','— select product —'],['providers','mProvider','— select provider —'],['routes','mRoute','— select —'],['clients','mPrevClient','— select prev client —']];
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
  const text  = await f.text();
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

// ── EVENT LISTENERS ───────────────────────────────────
document.getElementById('beOv').addEventListener('click', function(e) { if(e.target===this) closeBE(); });
document.getElementById('umOv').addEventListener('click', function(e) { if(e.target===this) closeUM(); });
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
renderDash(); renderTbl(); renderLogs();
