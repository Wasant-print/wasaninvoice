// ============================================================
//  วสันต์ปริ้นเตอร์ — Apps Script API Backend
//  ทำงานเป็น REST API ให้ GitHub Pages เรียกใช้
// ============================================================

const SS          = () => SpreadsheetApp.getActiveSpreadsheet();
const SHEET_DOCS  = 'ฐานข้อมูลเอกสาร';
const SHEET_ITEMS = 'ฐานข้อมูลรายการ';
const SHEET_CFG   = 'ตั้งค่าระบบ';
const SHEET_USERS = 'ผู้ใช้งาน';

// ── CORS HEADERS ─────────────────────────────────────────────
function setCORS(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ── RESPONSE HELPERS ─────────────────────────────────────────
function ok(data) {
  return setCORS(ContentService
    .createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON));
}
function err(msg) {
  return setCORS(ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON));
}

// ── ROUTING ──────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';
    switch (action) {
      case 'getDocs':     return ok(getAllDocs_());
      case 'getSettings': return ok(getSettings_());
      case 'getUsers':    return ok(getUsers_());
      default:
        // ถ้าไม่มี action = เปิดจาก Apps Script โดยตรง
        return HtmlService.createHtmlOutput('<h2>API พร้อมใช้งาน ✅</h2><p>ใช้งานผ่าน GitHub Pages</p>');
    }
  } catch(e) { return err(e.message); }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || '';
    switch (action) {
      case 'login':        return ok(login_(body.username, body.password));
      case 'saveDoc':      return ok(saveDoc_(body.data));
      case 'deleteDoc':    return ok(deleteDoc_(body.number));
      case 'saveSettings': return ok(saveSettings_(body.data));
      case 'saveUser':     return ok(saveUser_(body.data));
      case 'deleteUser':   return ok(deleteUser_(body.username));
      default:             return err('Unknown action: ' + action);
    }
  } catch(e) { return err(e.message); }
}

// ── SHEET HELPER ─────────────────────────────────────────────
function getSheet_(name, headers) {
  const ss = SS(); let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) sh.getRange(1,1,1,headers.length).setValues([headers]);
    if ([SHEET_CFG, SHEET_USERS].includes(name)) sh.hideSheet();
  }
  return sh;
}
function delRows_(sh, val, col) {
  if (!sh) return;
  const d = sh.getDataRange().getValues();
  for (let i = d.length-1; i >= 1; i--) if (d[i][col] == val) sh.deleteRow(i+1);
}
function typeCode_(n) { return {ใบเสนอราคา:'QT',ใบเสร็จรับเงิน:'RT',ใบแจ้งหนี้:'IV',ใบกำกับภาษี:'TV',ใบส่งของ:'DO'}[n]||'QT'; }
function typeName_(c) { return {QT:'ใบเสนอราคา',RT:'ใบเสร็จรับเงิน',IV:'ใบแจ้งหนี้',TV:'ใบกำกับภาษี',DO:'ใบส่งของ'}[c]||'ใบเสนอราคา'; }

// ── AUTH ─────────────────────────────────────────────────────
function login_(username, password) {
  const sh = getSheet_(SHEET_USERS, ['username','password','role','name','active']);
  const rows = sh.getDataRange().getValues().slice(1);
  if (rows.length === 0) {
    sh.appendRow(['admin','admin1234','admin','เจ้าของ',true]);
    sh.appendRow(['staff1','staff1234','staff','พนักงาน 1',true]);
  }
  for (const r of sh.getDataRange().getValues().slice(1)) {
    if (r[0] == username && r[1] == password && r[4] !== false)
      return { ok: true, role: r[2], name: r[3] };
  }
  return { ok: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

// ── USERS ─────────────────────────────────────────────────────
function getUsers_() {
  const sh = getSheet_(SHEET_USERS, []);
  return sh.getDataRange().getValues().slice(1)
    .map(r => ({ username: r[0], role: r[2], name: r[3], active: r[4] }));
}
function saveUser_(u) {
  const sh = getSheet_(SHEET_USERS, []);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == u.username) {
      sh.getRange(i+1,1,1,5).setValues([[u.username, u.password||rows[i][1], u.role, u.name, u.active !== false]]);
      return { success: true };
    }
  }
  sh.appendRow([u.username, u.password||'staff1234', u.role||'staff', u.name||u.username, true]);
  return { success: true };
}
function deleteUser_(username) {
  if (username === 'admin') return { error: 'ไม่สามารถลบ admin หลักได้' };
  const sh = getSheet_(SHEET_USERS, []);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length-1; i >= 1; i--) {
    if (rows[i][0] == username) { sh.deleteRow(i+1); return { success: true }; }
  }
  return { error: 'ไม่พบผู้ใช้' };
}

// ── SETTINGS ─────────────────────────────────────────────────
function getSettings_() {
  const sh = getSheet_(SHEET_CFG, ['key','value']), map = {};
  sh.getDataRange().getValues().slice(1).forEach(r => { if(r[0]) map[r[0]] = r[1]; });
  return { logo: map['logo']||'', sig: map['sig']||'' };
}
function saveSettings_(p) {
  const sh = getSheet_(SHEET_CFG, ['key','value']);
  const rows = sh.getDataRange().getValues();
  const ups = (key, val) => {
    if (val === undefined || val === null) return;
    for (let i = 1; i < rows.length; i++) { if (rows[i][0] === key) { sh.getRange(i+1,2).setValue(val); return; } }
    sh.appendRow([key, val]);
  };
  ups('logo', p.logo); ups('sig', p.sig);
  return { success: true };
}

// ── DOCUMENTS ─────────────────────────────────────────────────
function getAllDocs_() {
  const ss = SS();
  const shD = ss.getSheetByName(SHEET_DOCS), shI = ss.getSheetByName(SHEET_ITEMS);
  if (!shD || !shI) return [];
  const dR = shD.getDataRange().getValues(), iR = shI.getDataRange().getValues();
  const dH = dR[0], iH = iR[0], iMap = {};
  iR.slice(1).forEach(r => {
    const o = {}; iH.forEach((h,i) => o[h] = r[i]);
    const n = o['เลขที่เอกสาร']; if(!n) return;
    if(!iMap[n]) iMap[n] = [];
    iMap[n].push({ name: o['สินค้า / บริการ']||'', qty: parseFloat(o['จำนวน'])||0, price: parseFloat(o['ราคาต่อหน่วย'])||0 });
  });
  return dR.slice(1).filter(r => r[0]).map(r => {
    const o = {}; dH.forEach((h,i) => o[h] = r[i]);
    const n = o['เลขที่เอกสาร'], items = iMap[n]||[];
    const total = items.reduce((s,it) => s+it.qty*it.price, 0);
    const hasVat = (o['มี VAT ?']||'').toString().includes('มี');
    let date = o['วันที่สร้าง'];
    if (date instanceof Date) date = Utilities.formatDate(date,'Asia/Bangkok','yyyy-MM-dd');
    return { number: n, type: typeCode_(o['ประเภทเอกสาร']), typeName: o['ประเภทเอกสาร']||'',
      date, ref: o['อ้างอิงเอกสาร']||'', customerName: o['ชื่อลูกค้า']||'',
      customerAddr1: o['ที่อยู่ลูกค้า 1']||'', customerAddr2: o['ที่อยู่ลูกค้า 2']||'',
      customerTax: o['เลขที่เสียภาษีลูกค้า']||'', customerPhone: o['เบอร์โทรศัพท์ลูกค้า']||'',
      note: o['หมายเหตุ']||'', hasVat, items, total };
  });
}
function saveDoc_(d) {
  const shD = getSheet_(SHEET_DOCS, ['เลขที่เอกสาร','ประเภทเอกสาร','วันที่สร้าง','อ้างอิงเอกสาร','ชื่อลูกค้า','ที่อยู่ลูกค้า 1','ที่อยู่ลูกค้า 2','เลขที่เสียภาษีลูกค้า','เบอร์โทรศัพท์ลูกค้า','หมายเหตุ','ส่วนลด','มี VAT ?']);
  const shI = getSheet_(SHEET_ITEMS, ['เลขที่เอกสาร','สินค้า / บริการ','จำนวน','ราคาต่อหน่วย']);
  const n = d.number, date = d.date ? new Date(d.date) : new Date();
  delRows_(shD, n, 0); delRows_(shI, n, 0);
  shD.appendRow([n, d.typeName||typeName_(d.type), date, d.ref||'', d.customerName||'',
    d.customerAddr1||'', d.customerAddr2||'', d.customerTax||'', d.customerPhone||'',
    d.note||'', 0, d.hasVat ? 'มี VAT' : 'ไม่มี VAT']);
  (d.items||[]).forEach(it => { if(it.name) shI.appendRow([n, it.name, it.qty, it.price]); });
  return { success: true, number: n };
}
function deleteDoc_(num) {
  const ss = SS();
  delRows_(ss.getSheetByName(SHEET_DOCS), num, 0);
  delRows_(ss.getSheetByName(SHEET_ITEMS), num, 0);
  return { success: true };
}

// ── MENU ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📄 ออกเอกสาร')
    .addItem('เปิด API Status', 'showApiUrl_').addToUi();
}
function showApiUrl_() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('API URL:\n' + url + '\n\nคัดลอกไปวางใน index.html ที่ตัวแปร API_URL');
}
