// =====================================================
// Dashboard Nominasi Operator RTG + CC – TPS
// dashboard.js (gabungan, dengan equipment switcher)
// =====================================================

// ═══════════════════════════════════════════════════
// STATE — terpisah total per jenis alat
// ═══════════════════════════════════════════════════
let currentEquip = 'rtg';   // 'rtg' | 'cc'
let currentPane  = 'upload';

const ST = {
  rtg: { wb:null, sheets:{}, allData:{}, curPeriode:null },
  cc:  { wb:null, sheets:{}, allData:{}, curPeriode:null },
};
const S = () => ST[currentEquip];

// ═══════════════════════════════════════════════════
// CONFIG per jenis alat — sheet yang dicari & bobot
// ═══════════════════════════════════════════════════
const CFG = {
  rtg: {
    label:'RTG', icon:'🏗️', fileHint:'FINAL BEST RTG, List Operator RTG, Presensi, K3, dll.',
    k3Keyword:'RTG',
    sheetCfg: {
      raw:      { label:'FINAL BEST RTG',    kw:['FINAL BEST RTG'],                     req:true,  ico:'📊' },
      manning:  { label:'List Operator RTG', kw:['LIST OPERATOR RTG'],                  req:true,  ico:'👥' },
      presensi: { label:'Presensi',          kw:['PRESENSI','ABSENSI'],                 req:false, ico:'📋' },
      k3:       { label:'K3',                kw:['K3','INSIDEN','INCIDENT'],            req:false, ico:'⚠️' },
    },
    bobot: {D:35, K:20, Kal:20, M:15, Man:10},
    // type:'ratio' → poin = bobot * MIN(nilai/avg-grup,1)
    // type:'match' → poin = bobot * (match/(match+notMatch))
    components: [
      {key:'D',  label:'RTG Dinson', ico:'🏗️', color:'#1e4080', type:'ratio', field:'mvDinson', isVolume:true, dec:0},
      {key:'K',  label:'RTG Kone',   ico:'🏗️', color:'#00A896', type:'ratio', field:'mvKone',   isVolume:true, dec:0},
      {key:'Kal',label:'RTG Kalmar', ico:'🏗️', color:'#D4A017', type:'ratio', field:'mvKalmar', isVolume:true, dec:0},
      {key:'M',  label:'Match Penempatan', ico:'🎯', color:'#7b5ea7', type:'match', fields:['matchPlan','notMatchPlan']},
      {key:'Man',label:'Match Manning',    ico:'👷', color:'#C0392B', type:'match', fields:['matchMan','notMatchMan']},
    ],
    volumeLabel:'Total Moves',
  },
  cc: {
    label:'CC', icon:'🚢', fileHint:'Final Best CC, List Operator CC, Presensi, K3, dll.',
    k3Keyword:'CC',
    sheetCfg: {
      raw:      { label:'Final Best CC',    kw:['FINAL BEST CC'],       req:true,  ico:'📊' },
      manning:  { label:'List Operator CC', kw:['LIST OPERATOR CC'],    req:true,  ico:'👥' },
      presensi: { label:'Presensi',         kw:['PRESENSI','ABSENSI'],  req:false, ico:'📋' },
      k3:       { label:'K3',               kw:['K3','INSIDEN','INCIDENT'], req:false, ico:'⚠️' },
    },
    bobot: {MO:30, MN:20, NO:30, NN:20},
    components: [
      {key:'MO',label:'Moves Old', ico:'📦', color:'#1e4080', type:'ratio', field:'movesOld', isVolume:true, dec:0},
      {key:'MN',label:'Moves New', ico:'📦', color:'#00A896', type:'ratio', field:'movesNew', isVolume:true, dec:0},
      {key:'NO',label:'NMPH Old',  ico:'⚡', color:'#D4A017', type:'ratio', field:'nmphOld',  isVolume:false, dec:2},
      {key:'NN',label:'NMPH New',  ico:'⚡', color:'#7b5ea7', type:'ratio', field:'nmphNew',  isVolume:false, dec:2},
    ],
    volumeLabel:'Total Moves',
  },
};

// ═══════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════
const toN = v => { const n=parseFloat(String(v??'').replace(/[^0-9.\-]/g,'')); return isNaN(n)?0:n; };
const norm = s => String(s||'').replace(/^\d+\s*[-–]\s*/,'').trim().toUpperCase().replace(/\s+/g,' ');

function findSheet(wb, keywords) {
  if (!wb) return null;
  for (const kw of keywords) {
    const sn = wb.SheetNames.find(s =>
      s.toUpperCase().replace(/[\s_\-]/g,'').includes(kw.toUpperCase().replace(/[\s_\-]/g,''))
    );
    if (sn) return sn;
  }
  return null;
}

function getRows(key) {
  const st = S(); const sn = st.sheets[key]; if (!sn || !st.wb) return null;
  return XLSX.utils.sheet_to_json(st.wb.Sheets[sn], {header:1, defval:null, raw:false});
}

// ═══════════════════════════════════════════════════
// EQUIPMENT SWITCHER
// ═══════════════════════════════════════════════════
function setEquip(eq) {
  if (eq === currentEquip) return;
  currentEquip = eq;
  document.getElementById('eq-rtg').classList.toggle('on', eq==='rtg');
  document.getElementById('eq-cc').classList.toggle('on', eq==='cc');
  renderUploadPane();
  updStatus();
  // re-render pane yang sedang aktif supaya nunjukin data equip yang baru dipilih
  if (currentPane==='dash') rDash();
  if (currentPane==='tbl')  rTbl();
  if (currentPane==='det')  rDet();
  if (currentPane==='thn')  rThn();
}

// ═══════════════════════════════════════════════════
// BOBOT BOX (render dinamis sesuai equip)
// ═══════════════════════════════════════════════════
function renderBobot() {
  const cfg = CFG[currentEquip];
  const items = cfg.components.map(c=>
    `<div class="bi"><label>${c.label}</label><div class="bi-row"><span class="bval">${cfg.bobot[c.key]}</span><span style="font-size:12px;color:var(--muted)">poin</span></div></div>`
  ).join('');
  const total = Object.values(cfg.bobot).reduce((a,b)=>a+b,0);
  document.getElementById('fbox').innerHTML = `
    <h3>BOBOT PENILAIAN ${cfg.icon} ${cfg.label}</h3>
    <div class="fgrid">${items}</div>
    <div class="fi-d">
      Operator <b>GUGUR</b> jika :<br>
      • Terlambat / Pulang Cepat / Tanpa Keterangan / Tidak Absen Masuk atau Keluar<br>
      • Terlibat insiden K3 (Operator ${cfg.label})</div>
    <div class="btotal ${total===100?'ok':'warn'}">${total===100?'✅':'⚠️'} Total: ${total} poin</div>`;
}

// ═══════════════════════════════════════════════════
// UPLOAD ZONE
// ═══════════════════════════════════════════════════
function handleDrop(e) { const f=e.dataTransfer.files[0]; if(f) loadFile(f); }

function resetFile() {
  const st = S(); st.wb=null; st.sheets={};
  document.getElementById('uz').className='upload-zone';
  document.getElementById('uz').onclick=()=>document.getElementById('fi').click();
  document.getElementById('uz-ico').textContent='📂';
  document.getElementById('uz-title').textContent='Drag & drop file Excel ke sini';
  document.getElementById('uz-desc').textContent='1 file berisi semua sheet: '+CFG[currentEquip].fileHint;
  document.getElementById('uz-chk').innerHTML='';
  document.getElementById('uz-extra').textContent='';
  document.getElementById('fi').value='';
  updStatus();
}

function renderUploadPane() {
  renderBobot();
  const st = S();
  if (st.wb) {
    renderChecklist(st._fname || 'File ter-upload', st.wb.SheetNames);
  } else {
    resetFile();
  }
}

function loadFile(file) {
  if (!file) return;
  const eq = currentEquip; // kunci ke equip yang aktif saat upload
  const rd = new FileReader();
  rd.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {type:'array', cellDates:true});
      const sheets = {};
      for (const [k,c] of Object.entries(CFG[eq].sheetCfg)) sheets[k] = findSheet(wb, c.kw);
      ST[eq].wb = wb; ST[eq].sheets = sheets; ST[eq]._fname = file.name;
      renderChecklist(file.name, wb.SheetNames);
      updStatus();
    } catch(err) {
      document.getElementById('uz').className='upload-zone err';
      document.getElementById('uz-title').textContent='❌ Gagal membaca file';
      document.getElementById('uz-desc').textContent=err.message;
    }
  };
  rd.readAsArrayBuffer(file);
}

function renderChecklist(fname, sheetNames) {
  const uz = document.getElementById('uz');
  uz.className = 'upload-zone ok';
  uz.onclick = null;
  document.getElementById('uz-ico').textContent = '✅';
  document.getElementById('uz-title').textContent = fname;
  document.getElementById('uz-desc').textContent = sheetNames.length + ' sheet ditemukan';

  const sheets = S().sheets;
  let h = '<div class="chk-grid">';
  for (const [k,c] of Object.entries(CFG[currentEquip].sheetCfg)) {
    const found = sheets[k];
    const cls = found ? 'found' : (c.req ? 'miss' : 'opt');
    h += `<div class="chk-item ${cls}"><div class="chk-dot"></div>
      <span>${c.ico} ${c.label}</span>
      <span style="margin-left:auto;font-size:9px">${found ? '→ '+found : (c.req ? '❌ WAJIB' : 'opsional')}</span>
    </div>`;
  }
  h += '</div>';
  document.getElementById('uz-chk').innerHTML = h;

  const used = new Set(Object.values(sheets).filter(Boolean));
  const unk = sheetNames.filter(s=>!used.has(s));
  document.getElementById('uz-extra').textContent =
    unk.length ? 'Sheet tidak dikenali: '+unk.join(', ') : '';
}

function updStatus() {
  const el=document.getElementById('pstat'), btn=document.getElementById('go-btn');
  const st = S(); const cfg = CFG[currentEquip];
  if (!st.wb) { el.textContent=`Upload file Excel ${cfg.label} terlebih dahulu.`; btn.disabled=true; return; }
  const miss = Object.entries(cfg.sheetCfg).filter(([k,c])=>c.req&&!st.sheets[k]).map(([,c])=>c.label);
  if (miss.length) { el.textContent='⚠️ Sheet wajib tidak ditemukan: '+miss.join(', '); btn.disabled=true; }
  else {
    const n = Object.values(st.sheets).filter(Boolean).length;
    el.textContent = `✅ Siap! ${n}/${Object.keys(cfg.sheetCfg).length} sheet terdeteksi. Klik Proses.`;
    btn.disabled = false;
  }
}

// ═══════════════════════════════════════════════════
// PARSERS — RAW (beda struktur RTG vs CC)
// ═══════════════════════════════════════════════════

// RTG — FINAL BEST RTG: kolom tetap sesuai formula Excel RTG
// col0=No,1=PIN,2=Nama,3=NIPP,4=Group,5=Dinson,6=Kone,7=Kalmar,
// 9=MatchPlan,10=NotMatchPlan,12=MatchMan,13=NotMatchMan,15=PointTotal
function parseRawRTG() {
  const rows = getRows('raw'); if (!rows) return [];
  const ops = [];
  for (let i=1; i<rows.length; i++) {
    const r = rows[i]; if (!r || !r[1]) continue;
    const pin = String(r[1]).trim().replace(/\D/g,'');
    if (!pin) continue;
    const grup = String(r[4]||'').trim().toUpperCase();
    if (!['A','B','C','D'].includes(grup)) continue;
    const namaRaw = String(r[2]||'').trim();
    const nama = namaRaw.replace(/^\d+\s*[-–]\s*/,'').trim();
    ops.push({
      no: toN(r[0]), pin, nama, nipp: String(r[3]||'').trim(), group: grup,
      mvDinson: toN(r[5]), mvKone: toN(r[6]), mvKalmar: toN(r[7]),
      matchPlan: toN(r[9]), notMatchPlan: toN(r[10]),
      matchMan: toN(r[12]), notMatchMan: toN(r[13]),
      rawTotalPoint: toN(r[15]),
    });
  }
  return ops;
}

// CC — Final Best CC: header dicari otomatis (tahan perubahan urutan kolom)
function parseRawCC() {
  const rows = getRows('raw'); if (!rows) return [];
  let hdr = 0;
  for (let i=0; i<Math.min(rows.length,5); i++) {
    const r = rows[i]; if (!r) continue;
    const s = r.map(v=>String(v||'')).join('|').toUpperCase();
    if (s.includes('PIN') && (s.includes('GROUP')||s.includes('GRUP')) && s.includes('MOVES')) { hdr=i; break; }
  }
  const h = rows[hdr] || [];
  const col = {};
  h.forEach((v,i)=>{
    const s = String(v||'').toUpperCase().replace(/[\s_\-]+/g,'');
    if (s==='PIN') col.pin=i;
    else if (s.includes('NAMAOPERATOR')||s==='NAMA') col.nama=i;
    else if (s==='NIPP') col.nipp=i;
    else if (s==='GROUP'||s==='GRUP') col.group=i;
    else if (s==='MOVESOLD') col.movesOld=i;
    else if (s==='MOVESNEW') col.movesNew=i;
    else if (s==='NMPHOLD') col.nmphOld=i;
    else if (s==='NMPHNEW') col.nmphNew=i;
    else if (s.includes('TOTALPOINT')) col.totalPoint=i;
  });
  const c = k => col[k] !== undefined ? col[k] : ({
    pin:1, nama:2, nipp:3, group:4, movesOld:5, movesNew:6,
    nmphOld:8, nmphNew:9, totalPoint:11
  })[k];

  const ops = [];
  for (let i=hdr+1; i<rows.length; i++) {
    const r = rows[i]; if (!r || !r[c('pin')]) continue;
    const pin = String(r[c('pin')]).trim().replace(/\D/g,'');
    if (!pin) continue;
    const grup = String(r[c('group')]||'').trim().toUpperCase();
    if (!['A','B','C','D'].includes(grup)) continue;
    const namaRaw = String(r[c('nama')]||'').trim();
    const nama = namaRaw.replace(/^\d+\s*[-–]\s*/,'').trim();
    ops.push({
      no: i, pin, nama, nipp: String(r[c('nipp')]||'').trim(), group: grup,
      movesOld: toN(r[c('movesOld')]), movesNew: toN(r[c('movesNew')]),
      nmphOld: toN(r[c('nmphOld')]), nmphNew: toN(r[c('nmphNew')]),
      rawTotalPoint: toN(r[c('totalPoint')]),
    });
  }
  return ops;
}

function parseRaw() { return currentEquip==='rtg' ? parseRawRTG() : parseRawCC(); }

// MANNING — sama persis strukturnya untuk RTG & CC: R1=header(NO|NAMA|NIPP|GRUP)
function parseManning() {
  const rows = getRows('manning'); if (!rows) return {};
  const map = {};
  let hdr = 0;
  for (let i=0; i<Math.min(rows.length,5); i++) {
    const r = rows[i]; if (!r) continue;
    const s = r.map(v=>String(v||'')).join('|').toUpperCase();
    if (s.includes('NAMA') && s.includes('NIPP') && s.includes('GRUP')) { hdr=i; break; }
  }
  for (let i=hdr+1; i<rows.length; i++) {
    const r = rows[i]; if (!r || !r[0]) continue;
    const namaRaw = String(r[1]||'').trim();
    const pm = namaRaw.match(/^(\d{4})\s*[-–]/);
    if (!pm) continue;
    const gs = String(r[3]||'').trim().toUpperCase();
    if (!['A','B','C','D'].includes(gs)) continue;
    const pin = pm[1];
    map[pin] = {
      pin, nama: namaRaw.replace(/^\d+\s*[-–]\s*/,'').trim(),
      nipp: String(r[2]||'').trim(), group: gs,
    };
  }
  return map;
}

// PRESENSI — sama persis untuk RTG & CC, kolom dicari lewat header
function parsePresensi() {
  const rows = getRows('presensi'); if (!rows) return {};
  const map = {};
  let hdr = 2;
  for (let i=0; i<Math.min(rows.length,5); i++) {
    const r = rows[i]; if (!r) continue;
    const s = r.map(v=>String(v||'')).join('|').toUpperCase();
    if (s.includes('NIPP') && s.includes('TERLAMBAT')) { hdr=i; break; }
  }
  const h = rows[hdr] || [];
  let cNipp=3, cTK=15, cTel=16, cPC=17, cTM=18, cTK2=19;
  h.forEach((v,i)=>{
    const s = String(v||'').toUpperCase().replace(/[\s\n]+/g,'');
    if (s==='NIPP') cNipp=i;
    if (s.includes('TANPAKETERANGAN')) cTK=i;
    if (s.includes('TERLAMBATMASUK')||s==='TERLAMBAT') cTel=i;
    if (s.includes('PULANGCEPAT')) cPC=i;
    if (s.includes('TIDAKABSENMASUK')) cTM=i;
    if (s.includes('TIDAKABSENKELUAR')||s.includes('TIDAKABSENPULANG')) cTK2=i;
  });
  for (let i=hdr+1; i<rows.length; i++) {
    const r = rows[i]; if (!r) continue;
    const nipp = String(r[cNipp]||'').trim().replace(/\D/g,'');
    if (!nipp || nipp.length<5) continue;
    map[nipp] = {
      tanpaKet: toN(r[cTK]), telat: toN(r[cTel]),
      pc: toN(r[cPC]), tdkMasuk: toN(r[cTM]), tdkKeluar: toN(r[cTK2]),
    };
  }
  return map;
}

// K3 — filter kata kunci equip ('RTG' atau 'CC') supaya insiden yang
// tercampur di 1 sheet K3 tidak saling nyampur antar dashboard
function parseK3(manMap) {
  const rows = getRows('k3'); if (!rows) return new Set();
  const kw = CFG[currentEquip].k3Keyword;

  const nameToPinMap = {};
  for (const [pin, info] of Object.entries(manMap)) {
    nameToPinMap[norm(info.nama)] = pin;
  }

  function lookupPin(namaRaw) {
    const key = norm(namaRaw);
    if (nameToPinMap[key]) return nameToPinMap[key];
    const wA = key.split(' ');
    for (const [mKey, pin] of Object.entries(nameToPinMap)) {
      const wB = mKey.split(' ');
      const w2A = wA.slice(0,2).join(' '), w2B = wB.slice(0,2).join(' ');
      if (w2A.length > 3 && w2A === w2B) return pin;
      if (wA.length >= 2 && wB.length >= 2) {
        const t = s => s.slice(0,4);
        if (t(wA[0]) === t(wB[0]) && t(wA[1]) === t(wB[1])) return pin;
      }
    }
    return null;
  }

  let formatB = false;
  let cNama = 7, cJabatan = 8, cKronologi = 6;
  let cEquip = 11, cOp = 12;

  for (let i = 0; i < Math.min(rows.length, 8); i++) {
    const r = rows[i]; if (!r) continue;
    const joined = r.map(v => String(v||'').toUpperCase().replace(/\s+/g,'')).join('|');
    if (joined.includes('NAMA') && joined.includes('JABATAN')) {
      formatB = true;
      r.forEach((v, ci) => {
        const s = String(v||'').toUpperCase().replace(/\s+/g,'');
        if (s === 'NAMA')            cNama      = ci;
        if (s === 'JABATAN')         cJabatan   = ci;
        if (s.includes('KRONOLOGI')) cKronologi = ci;
      });
      break;
    }
    if (joined.includes('EQUIP')) {
      r.forEach((v, ci) => {
        const s = String(v||'').toUpperCase().replace(/\s+/g,'');
        if ((s==='EQUIPMENT'||s.includes('EQUIP')) && ci>=8 && ci<=14) cEquip = ci;
        if ((s.includes('NAMAEQUIP')||s.includes('OPERATORN')) && ci>=9) cOp = ci;
      });
      break;
    }
  }

  const k3Pins = new Set();

  for (const r of rows) {
    if (!r || !r[0]) continue;
    const no = parseInt(String(r[0]||'').replace(/\D/g,'')); if (isNaN(no) || no <= 0) continue;

    let namaRaw = '';
    if (formatB) {
      const jabatan   = String(r[cJabatan]   || '').toUpperCase();
      const kronologi = String(r[cKronologi] || '').toUpperCase();
      namaRaw         = String(r[cNama]      || '').trim();
      if (!namaRaw) continue;
      const isTargetOp = jabatan.includes(kw) || (kronologi.includes(kw) && jabatan.includes('OPERATOR'));
      if (!isTargetOp) continue;
    } else {
      const equip = String(r[cEquip] || '').toUpperCase();
      if (!equip.includes(kw)) continue;
      namaRaw = String(r[cOp] || '').trim();
      if (!namaRaw) continue;
    }

    const pin = lookupPin(namaRaw);
    if (pin) k3Pins.add(pin);
    else k3Pins.add('#' + norm(namaRaw));
  }

  return k3Pins;
}

// ═══════════════════════════════════════════════════
// SCORING GENERIC — dipakai untuk RTG maupun CC
// ═══════════════════════════════════════════════════
function computeScores(op, avgMap) {
  const cfg = CFG[currentEquip];
  const scores = {};
  let total = 0;
  for (const comp of cfg.components) {
    const bobot = cfg.bobot[comp.key];
    let val = 0, sub = '';
    if (comp.type === 'ratio') {
      const avg = avgMap[comp.field] || 0;
      val = avg > 0 ? bobot * Math.min(1, op[comp.field] / avg) : 0;
      const pct = avg > 0 ? Math.min(100, op[comp.field]/avg*100).toFixed(0) : 0;
      sub = `${op[comp.field].toLocaleString('id',{maximumFractionDigits:comp.dec})} · avg grup ${avg.toFixed(comp.dec)} · ratio ${pct}%`;
    } else if (comp.type === 'match') {
      const [mf, nf] = comp.fields;
      const tot = op[mf] + op[nf];
      val = tot > 0 ? bobot * (op[mf]/tot) : 0;
      sub = `${op[mf]} match · ${op[nf]} tidak · ${tot>0?((op[mf]/tot)*100).toFixed(1):0}%`;
    }
    scores[comp.key] = { val, max: bobot, sub };
    total += val;
  }
  return { scores, total };
}

function computeGroupAvg(ops) {
  const cfg = CFG[currentEquip];
  const avgMap = {};
  for (const comp of cfg.components) {
    if (comp.type !== 'ratio') continue;
    avgMap[comp.field] = ops.length ? ops.reduce((s,o)=>s+o[comp.field],0)/ops.length : 0;
  }
  return avgMap;
}

function volumeOf(op) {
  const cfg = CFG[currentEquip];
  return cfg.components.filter(c=>c.isVolume).reduce((s,c)=>s+(op[c.field]||0), 0);
}

// ═══════════════════════════════════════════════════
// MAIN PROCESS
// ═══════════════════════════════════════════════════
function processAll() {
  const periode = document.getElementById('periode-in').value.trim() ||
    ('PERIODE '+new Date().toLocaleDateString('id-ID',{month:'long',year:'numeric'}).toUpperCase());
  const st = S(); const cfg = CFG[currentEquip];
  if (!st.wb) { showAlert('err',`❌ Belum ada file ${cfg.label}!`); return; }
  if (!st.sheets.raw)     { showAlert('err',`❌ Sheet ${cfg.sheetCfg.raw.label} tidak ditemukan!`); return; }
  if (!st.sheets.manning) { showAlert('err',`❌ Sheet ${cfg.sheetCfg.manning.label} tidak ditemukan!`); return; }

  loading('Membaca data '+cfg.label+'...');
  setTimeout(()=>{
    try {
      const rawOps = parseRaw();
      if (!rawOps.length) { hide(); showAlert('err','❌ Data kosong atau format tidak sesuai.'); return; }

      loading('Membaca Manning...');
      const manMap = parseManning();

      loading('Membaca Presensi & K3...');
      const presMap = st.sheets.presensi ? parsePresensi() : {};
      const k3Pins  = st.sheets.k3       ? parseK3(manMap) : new Set();

      loading('Menghitung poin...');

      const groups = {A:[],B:[],C:[],D:[]};
      for (const op of rawOps) {
        const manInfo = manMap[op.pin];
        const nipp = manInfo?.nipp || op.nipp || '';
        const presData = nipp ? presMap[nipp] : null;
        const k3count = k3Pins.has(op.pin) ? 1
                      : k3Pins.has('#' + norm(op.nama)) ? 1
                      : 0;
        const g = op.group;
        if (!groups[g]) continue;
        groups[g].push({...op, nipp, presData, k3count});
      }

      for (const g of ['A','B','C','D']) {
        const ops = groups[g]; if (!ops.length) continue;
        const avgMap = computeGroupAvg(ops);

        for (const op of ops) {
          const {scores, total} = computeScores(op, avgMap);
          const pr = op.presData;
          const absenViol = pr
            ? (pr.tanpaKet>0||pr.telat>0||pr.pc>0||pr.tdkMasuk>0||pr.tdkKeluar>0)
            : false;
          const eligible = !absenViol && op.k3count===0;

          Object.assign(op, {
            avgMap, scores, totalPoin: total,
            absenViol, eligible,
            tanpaKet:  pr?.tanpaKet  || 0,
            telat:     pr?.telat     || 0,
            pc:        pr?.pc        || 0,
            tdkMasuk:  pr?.tdkMasuk  || 0,
            tdkKeluar: pr?.tdkKeluar || 0,
            hasPres:   !!pr,
          });
        }
        ops.sort((a,z)=>z.totalPoin-a.totalPoin);
      }

      let totalOps=0, totalMoves=0, totalK3=0;
      for (const g of ['A','B','C','D']) {
        groups[g].forEach(o=>{ totalOps++; totalMoves+=volumeOf(o); if(o.k3count) totalK3++; });
      }

      st.allData[periode] = {
        periode, uploadedAt: new Date().toISOString(),
        groups, bobot:{...cfg.bobot},
        stats:{ totalOps, totalMoves, totalK3,
          eligible: ['A','B','C','D'].flatMap(g=>groups[g]).filter(o=>o.eligible).length },
        sheetsFound: Object.entries(st.sheets).filter(([,v])=>v).map(([k])=>k),
      };
      st.curPeriode = periode;

      hide();
      showAlert('suc', `✅ <strong>${cfg.icon} ${periode}</strong> (${cfg.label}) berhasil! ${totalOps} operator, ${['A','B','C','D'].filter(g=>groups[g].length).length} grup.`);
      sw('dash', document.querySelectorAll('.tab')[1]);

    } catch(err) {
      hide();
      showAlert('err','❌ Error: '+err.message);
      console.error(err);
    }
  }, 80);
}

// ═══════════════════════════════════════════════════
// TAB SWITCH
// ═══════════════════════════════════════════════════
function sw(name, btn) {
  currentPane = name;
  document.querySelectorAll('.pane').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.tab').forEach(b=>b.classList.remove('on'));
  document.getElementById('pane-'+name).classList.add('on');
  if (btn) btn.classList.add('on');
  if (name==='upload') renderUploadPane();
  if (name==='dash') rDash();
  if (name==='tbl')  rTbl();
  if (name==='det')  rDet();
  if (name==='thn')  rThn();
}

// ═══════════════════════════════════════════════════
// ALERT & LOADING
// ═══════════════════════════════════════════════════
function showAlert(t,m){const e=document.getElementById('alertBox');e.className='alert on '+t;e.innerHTML=m;setTimeout(()=>e.classList.remove('on'),9000);}
function loading(t){document.getElementById('ovt').textContent=t||'Memproses...';document.getElementById('ov').classList.add('on');}
function hide(){document.getElementById('ov').classList.remove('on');}

// ═══════════════════════════════════════════════════
// RENDER DASHBOARD
// ═══════════════════════════════════════════════════
function rDash() {
  const st = S(); const cfg = CFG[currentEquip];
  const keys = Object.keys(st.allData);
  if (!keys.length) { document.getElementById('dash-content').innerHTML=`<div class="empty"><div class="ei">📊</div><h3>Belum ada data ${cfg.label}</h3><p>Upload & proses file ${cfg.label} di tab pertama.</p></div>`; return; }
  if (!st.curPeriode||!st.allData[st.curPeriode]) st.curPeriode=keys[keys.length-1];
  const pd = st.allData[st.curPeriode];
  const {groups,stats} = pd;

  const chips = keys.map(p=>`<div class="chip ${p===st.curPeriode?'on':''}" onclick="ST['${currentEquip}'].curPeriode='${p}';rDash()">
    📅 ${p}<span class="chip-x" onclick="event.stopPropagation();delP('${p}')"> ✕</span></div>`).join('');

  const statHtml = `<div class="stats">
    <div class="sc navy"><div class="sc-ico">👤</div><div><div class="sc-lbl">Total Operator</div><div class="sc-val">${stats.totalOps}</div><div class="sc-sub">${['A','B','C','D'].filter(g=>groups[g]?.length).length} grup aktif</div></div></div>
    <div class="sc gold"><div class="sc-ico">📦</div><div><div class="sc-lbl">${cfg.volumeLabel}</div><div class="sc-val">${stats.totalMoves.toLocaleString('id')}</div></div></div>
    <div class="sc green"><div class="sc-ico">✅</div><div><div class="sc-lbl">Memenuhi Syarat</div><div class="sc-val">${stats.eligible}</div><div class="sc-sub">Bisa dinominasikan</div></div></div>
    <div class="sc red"><div class="sc-ico">🚫</div><div><div class="sc-lbl">Gugur</div><div class="sc-val">${stats.totalOps-stats.eligible}</div><div class="sc-sub">Presensi / K3</div></div></div>
    <div class="sc teal"><div class="sc-ico">⚠️</div><div><div class="sc-lbl">Terlibat K3</div><div class="sc-val">${stats.totalK3}</div><div class="sc-sub">Insiden ${cfg.label}</div></div></div>
  </div>`;

  const gHtml = `<div class="grids">${['A','B','C','D'].map(g=>{
    const ops=groups[g]||[]; if(!ops.length) return '';
    const winner=ops.find(o=>o.eligible);
    const maxPt=ops[0]?.totalPoin||100;
    const medals=['🥇','🥈','🥉'];
    const elig=ops.filter(o=>o.eligible).length;
    const rows=ops.slice(0,10).map((op,i)=>{
      const rc=i===0?'r1':i===1?'r2':i===2?'r3':'';
      const pct=maxPt>0?Math.round(op.totalPoin/maxPt*100):0;
      const why=op.k3count>0?'K3':(op.absenViol?'Absen':'');
      return `<div class="op ${rc} ${op.eligible?'':'disq'}" onclick="showDet('${st.curPeriode}','${op.pin}')">
        <div class="rdot">${i<3?medals[i]:i+1}</div>
        <div class="op-inf">
          <div class="op-name">${op.nama}</div>
          <div class="op-meta">PIN ${op.pin} · ${volumeOf(op).toLocaleString('id')} moves</div>
        </div>
        ${op.eligible?`<span class="tag tag-ok">✅</span>`:`<span class="tag tag-no">🚫${why}</span>`}
        <div class="mini-bar"><div class="mb-bg"><div class="mb-fill" style="width:${pct}%"></div></div></div>
        <div class="op-pt ${i===0&&op.eligible?'hi':''}">${op.totalPoin.toFixed(1)}</div>
      </div>`;
    }).join('');
    return `<div class="gc">
      <div class="gc-hdr">
        <div><h3>GRUP ${g}</h3><div class="gc-sub">${ops.length} operator · ${elig} memenuhi syarat</div></div>
        <div class="gc-win">🏆 ${winner?.nama.split(' ').slice(0,2).join(' ')||'—'}</div>
      </div>${rows}
    </div>`;
  }).join('')}</div>`;

  const compCols = cfg.components.map(c=>`<th>${c.label}</th>`).join('');
  const wRows=['A','B','C','D'].map(g=>{
    const winner=(groups[g]||[]).find(o=>o.eligible); if(!winner) return '';
    const compVals = cfg.components.map(c=>`<td>${winner.scores[c.key].val.toFixed(2)}</td>`).join('');
    return `<tr class="tr-win">
      <td><span class="tag tag-${g}">GRUP ${g}</span></td>
      <td>${winner.pin}</td><td><strong>${winner.nama}</strong></td>
      ${compVals}
      <td><strong>${winner.totalPoin.toFixed(2)}</strong></td>
      <td><span class="tag tag-win">🏆 NOMINASI</span></td>
    </tr>`;
  }).join('');

  document.getElementById('dash-content').innerHTML=`
    <div style="margin-bottom:16px"><div class="stitle"><h2>📅 Periode — ${cfg.icon} ${cfg.label}</h2></div><div class="chips">${chips}</div></div>
    ${statHtml}
    <div class="stitle"><h2>📊 Ranking per Grup</h2><div class="sbadge">${st.curPeriode}</div></div>
    <p style="font-size:11px;color:var(--muted);margin-bottom:12px">Klik nama operator untuk detail · ✅ = Memenuhi syarat · 🚫 = Gugur</p>
    ${gHtml}
    <div class="stitle"><h2>🏅 Nominasi ${st.curPeriode}</h2></div>
    <div class="twrap"><div style="overflow-x:auto"><table>
      <thead><tr><th>Grup</th><th>PIN</th><th>Nama</th>${compCols}<th>TOTAL</th><th>Status</th></tr></thead>
      <tbody>${wRows||`<tr><td colspan="${3+cfg.components.length+2}" style="text-align:center;padding:20px;color:var(--muted)">Tidak ada operator yang memenuhi syarat</td></tr>`}</tbody>
    </table></div></div>`;
}

function delP(p){if(!confirm(`Hapus data "${p}"?`))return;delete S().allData[p];if(S().curPeriode===p)S().curPeriode=Object.keys(S().allData).slice(-1)[0]||null;rDash();}

// ═══════════════════════════════════════════════════
// RENDER DATA LENGKAP
// ═══════════════════════════════════════════════════
function rTbl() {
  const st = S(); const cfg = CFG[currentEquip];
  const keys=Object.keys(st.allData);
  if(!keys.length){document.getElementById('tbl-content').innerHTML=`<div class="empty"><div class="ei">📋</div><h3>Belum ada data ${cfg.label}</h3></div>`;return;}
  document.getElementById('tbl-content').innerHTML=`
    <div class="sbar">
      <input class="sinput" id="ts" placeholder="🔍 Cari nama atau PIN..." oninput="fTbl()">
      <select class="ssel" id="tg" onchange="fTbl()"><option value="">Semua Grup</option><option>A</option><option>B</option><option>C</option><option>D</option></select>
      <select class="ssel" id="tp" onchange="fTbl()">${keys.sort().map(p=>`<option ${p===st.curPeriode?'selected':''}>${p}</option>`).join('')}</select>
      <select class="ssel" id="tf" onchange="fTbl()"><option value="">Semua</option><option value="ok">Memenuhi Syarat</option><option value="no">Gugur</option></select>
    </div>
    <div id="tbl-rows"></div>`;
  fTbl();
}

function fTbl(){
  const st = S(); const cfg = CFG[currentEquip];
  const q=(document.getElementById('ts')?.value||'').toUpperCase();
  const gf=document.getElementById('tg')?.value||'';
  const pf=document.getElementById('tp')?.value||st.curPeriode||'';
  const ff=document.getElementById('tf')?.value||'';
  const pd=st.allData[pf]; if(!pd) return;
  let rows='';
  for(const g of ['A','B','C','D']){
    if(gf&&gf!==g) continue;
    (pd.groups[g]||[]).forEach((op,i)=>{
      if(q&&!op.nama.toUpperCase().includes(q)&&!op.pin.includes(q)) return;
      if(ff==='ok'&&!op.eligible) return;
      if(ff==='no'&&op.eligible) return;
      const why=!op.eligible?(op.k3count>0?'K3':(op.absenViol?'Absen':'')):'' ;
      const compTds = cfg.components.map(c=>`<td>${op.scores[c.key].val.toFixed(2)}</td>`).join('');
      rows+=`<tr class="${i===0&&op.eligible?'tr-win':''} ${op.eligible?'':'tr-dq'}" style="cursor:pointer" onclick="showDet('${pf}','${op.pin}')">
        <td>${i+1}</td><td><span class="tag tag-${g}">GRUP ${g}</span></td>
        <td>${op.pin}</td><td>${op.nipp||'—'}</td><td><strong>${op.nama}</strong></td>
        ${compTds}
        <td><strong>${op.totalPoin.toFixed(2)}</strong></td>
        <td>${op.hasPres?(op.tanpaKet+'/'+op.telat+'/'+op.pc+'/'+op.tdkMasuk+'/'+op.tdkKeluar):'—'}</td>
        <td>${op.k3count>0?`<span class="tag tag-no">⚠️${op.k3count}</span>`:'—'}</td>
        <td>${op.eligible?'<span class="tag tag-ok">✅</span>':`<span class="tag tag-no">🚫 ${why}</span>`}</td>
      </tr>`;
    });
  }
  const compHdr = cfg.components.map(c=>`<th>${c.label}</th>`).join('');
  document.getElementById('tbl-rows').innerHTML=`<div class="twrap">
    <div class="thdr"><h3>Data Lengkap ${cfg.label} – ${pf}</h3><div class="sbadge">${pd.stats.totalOps} Operator</div><span style="font-size:11px;color:var(--muted);margin-left:6px">Klik baris untuk detail</span></div>
    <div style="overflow-x:auto"><table>
      <thead><tr><th>Rank</th><th>Grup</th><th>PIN</th><th>NIPP</th><th>Nama</th>
        ${compHdr}
        <th>TOTAL</th><th>Absen(TK/T/PC/TM/TK)</th><th>K3</th><th>Status</th>
      </tr></thead>
      <tbody>${rows||`<tr><td colspan="${9+cfg.components.length}" style="text-align:center;padding:20px;color:var(--muted)">Tidak ada data</td></tr>`}</tbody>
    </table></div>
  </div>`;
}

// ═══════════════════════════════════════════════════
// RENDER DETAIL
// ═══════════════════════════════════════════════════
function rDet() {
  const st = S(); const cfg = CFG[currentEquip];
  const keys=Object.keys(st.allData);
  if(!keys.length){document.getElementById('det-content').innerHTML=`<div class="empty"><div class="ei">🔍</div><h3>Belum ada data ${cfg.label}</h3></div>`;return;}
  const pd=st.allData[st.curPeriode||keys[keys.length-1]];
  const allOps=['A','B','C','D'].flatMap(g=>(pd.groups[g]||[]));
  document.getElementById('det-content').innerHTML=`
    <div class="sbar">
      <select class="ssel" id="dp" style="min-width:160px" onchange="rDetOp()">${keys.sort().map(p=>`<option ${p===(st.curPeriode||'')?'selected':''}>${p}</option>`).join('')}</select>
      <select class="ssel" id="dop" style="flex:1;min-width:200px" onchange="rDetOp()">
        <option value="">— Pilih Operator —</option>
        ${allOps.map(o=>`<option value="${o.pin}">${o.pin} – ${o.nama} (Grup ${o.group})</option>`).join('')}
      </select>
    </div>
    <div id="det-op"></div>`;
}

function showDet(periode, pin) {
  const detTab = Array.from(document.querySelectorAll('.tab')).find(b=>b.textContent.includes('Detail'));
  sw('det', detTab);
  setTimeout(function(){
    const selP=document.getElementById('dp');
    if(selP) selP.value=periode;
    const sel=document.getElementById('dop');
    if(sel){ sel.value=pin; rDetOp(); }
  },50);
}

function rDetOp(){
  const st = S(); const cfg = CFG[currentEquip];
  const pf=document.getElementById('dp')?.value;
  const pin=document.getElementById('dop')?.value;
  if(!pf||!pin){document.getElementById('det-op').innerHTML='';return;}
  const pd=st.allData[pf]; if(!pd) return;
  const op=['A','B','C','D'].flatMap(g=>(pd.groups[g]||[])).find(o=>o.pin===pin);
  if(!op) return;
  const rank=(pd.groups[op.group]||[]).findIndex(o=>o.pin===pin)+1;

  const sCards=cfg.components.map(c=>{
    const s=op.scores[c.key];
    return `<div class="scomp">
      <div class="sc2-ico">${c.ico}</div>
      <div style="flex:1">
        <div class="sc2-lbl">${c.label}</div>
        <div class="sc2-val">${s.val.toFixed(2)}<span style="font-size:11px;color:var(--muted)"> / ${s.max}</span></div>
        <div class="sc2-sub">${s.sub}</div>
        <div class="sc2-bar"><div class="sc2-fill" style="width:${s.max>0?Math.min(100,s.val/s.max*100).toFixed(0):0}%;background:${c.color}"></div></div>
      </div></div>`;
  }).join('');

  const absenItems=[
    ['Tanpa Keterangan',op.tanpaKet],['Terlambat',op.telat],
    ['Pulang Cepat',op.pc],['Tidak Absen Masuk',op.tdkMasuk],['Tidak Absen Keluar',op.tdkKeluar],
  ].map(([l,v])=>`<div class="abitem" style="background:${v>0?'var(--red-pale)':'var(--green-pale)'}">
    <div class="abitem-lbl">${l}</div>
    <div class="abitem-val" style="color:${v>0?'var(--red)':'var(--green)'}">${v}</div>
  </div>`).join('');

  const eligBox=op.eligible
    ?`<div class="sybox sy-ok">✅ Memenuhi semua syarat — Presensi bersih & tidak ada insiden ${cfg.label}</div>`
    :`<div class="sybox sy-no">🚫 Gugur: ${op.absenViol?'Ada pelanggaran presensi':''}${op.absenViol&&op.k3count?' + ':''}${op.k3count?`Terlibat ${op.k3count} insiden ${cfg.label}`:''}</div>`;

  document.getElementById('det-op').innerHTML=`<div class="det">
    <div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:16px">
      <div style="font-size:46px">👤</div>
      <div>
        <h2 style="font-size:20px;font-weight:800;color:var(--navy)">${op.nama}</h2>
        <p style="font-size:13px;color:var(--muted)">PIN: ${op.pin} · NIPP: ${op.nipp||'—'} · <span class="tag tag-${op.group}">GRUP ${op.group}</span> · Rank #${rank} · ${cfg.icon} ${cfg.label}</p>
      </div>
      <div style="margin-left:auto;text-align:right">
        <div style="font-size:36px;font-weight:900;color:var(--teal)">${op.totalPoin.toFixed(2)}</div>
        <div style="font-size:11px;color:var(--muted)">TOTAL POIN</div>
        ${Math.abs(op.totalPoin-op.rawTotalPoint)>0.5?`<div style="font-size:10px;color:var(--red)">Excel: ${op.rawTotalPoint.toFixed(2)}</div>`:''}
      </div>
    </div>
    <div class="score-grid">${sCards}</div>
    ${eligBox}
    <div style="margin-top:14px">
      <h3 style="font-size:13px;font-weight:700;color:var(--navy);margin-bottom:6px">📋 Presensi ${op.hasPres?'':'(data tidak tersedia)'}</h3>
      ${op.hasPres?`<div class="abgrid">${absenItems}</div>`:'<p style="color:var(--muted);font-size:12px">Sheet Presensi tidak diupload.</p>'}
    </div>
    ${op.k3count>0?`<div style="margin-top:12px;padding:10px 14px;background:var(--red-pale);border-radius:8px;font-size:12px;color:var(--red);font-weight:600">⚠️ Terlibat ${op.k3count} insiden ${cfg.label} pada K3 — otomatis gugur</div>`:''}
  </div>`;
}

// ═══════════════════════════════════════════════════
// RENDER TAHUNAN
// ═══════════════════════════════════════════════════
function rThn(){
  const st = S(); const cfg = CFG[currentEquip];
  const keys=Object.keys(st.allData).sort();
  if(!keys.length){document.getElementById('thn-content').innerHTML=`<div class="empty"><div class="ei">🏅</div><h3>Belum ada data ${cfg.label}</h3></div>`;return;}
  const track={};
  for(const p of keys){
    const pd=st.allData[p];
    for(const g of ['A','B','C','D']){
      const w=(pd.groups[g]||[]).find(o=>o.eligible); if(!w) continue;
      if(!track[w.pin]) track[w.pin]={...w,wins:[],pts:[]};
      track[w.pin].wins.push(p); track[w.pin].pts.push(w.totalPoin);
    }
  }
  const sorted=Object.values(track).sort((a,z)=>z.wins.length-a.wins.length||(z.pts.reduce((s,v)=>s+v,0)/z.pts.length)-(a.pts.reduce((s,v)=>s+v,0)/a.pts.length));
  const rows=sorted.map((w,i)=>{
    const avg=(w.pts.reduce((s,v)=>s+v,0)/w.pts.length).toFixed(1);
    return `<tr><td>${i+1}</td><td><span class="tag tag-${w.group}">GRUP ${w.group}</span></td>
      <td>${w.pin}</td><td>${w.nipp||'—'}</td><td><strong>${w.nama}</strong></td>
      <td><strong style="color:var(--teal)">${w.wins.length}×</strong></td>
      <td style="font-size:11px">${w.wins.join(', ')}</td><td>${avg}</td>
      <td>${i<4?'<span class="tag tag-win">🏅</span>':''}</td></tr>`;
  }).join('');
  const byP=keys.map(p=>{
    const pd=st.allData[p];
    return `<tr><td><strong>${p}</strong></td>${['A','B','C','D'].map(g=>{
      const w=(pd.groups[g]||[]).find(o=>o.eligible);
      return `<td>${w?`<strong>${w.nama}</strong><br><small style="color:var(--muted)">${w.totalPoin.toFixed(1)} poin</small>`:'<small style="color:var(--muted)">—</small>'}</td>`;
    }).join('')}</tr>`;
  }).join('');
  document.getElementById('thn-content').innerHTML=`
    <div class="stitle"><h2>🏅 Rekap Nominasi Tahunan — ${cfg.icon} ${cfg.label}</h2><div class="sbadge">${keys.length} Periode</div></div>
    <div class="twrap"><div class="thdr"><h3>Operator Paling Sering Menang</h3></div>
      <table><thead><tr><th>Rank</th><th>Grup</th><th>PIN</th><th>NIPP</th><th>Nama</th><th>Menang</th><th>Periode</th><th>Avg Poin</th><th>Status</th></tr></thead>
      <tbody>${rows||'<tr><td colspan="9" style="text-align:center;padding:20px">Belum cukup data</td></tr>'}</tbody></table></div>
    <div class="stitle" style="margin-top:18px"><h2>📅 Pemenang per Periode</h2></div>
    <div class="twrap"><table>
      <thead><tr><th>Periode</th><th>Pemenang Grup A</th><th>Pemenang Grup B</th><th>Pemenang Grup C</th><th>Pemenang Grup D</th></tr></thead>
      <tbody>${byP}</tbody></table></div>`;
}

// ═══════════════════════════════════════════════════
// INIT
// ═══════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', renderUploadPane);