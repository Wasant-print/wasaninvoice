// ============================================================
//  วสันต์ปริ้นเตอร์ — Apps Script Backend  |  Code.gs
// ============================================================

const SS          = () => SpreadsheetApp.getActiveSpreadsheet();
const SHEET_DOCS  = 'ฐานข้อมูลเอกสาร';
const SHEET_ITEMS = 'ฐานข้อมูลรายการ';
const SHEET_CFG   = 'ตั้งค่าระบบ';
const SHEET_USERS = 'ผู้ใช้งาน';

// ── WEB APP ──────────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('วสันต์ปริ้นเตอร์')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width,initial-scale=1');
}

// ── SHEET HELPER ─────────────────────────────────────────────
function getSheet_(name, headers) {
  const ss = SS(); let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) sh.getRange(1,1,1,headers.length).setValues([headers]);
    if ([SHEET_CFG,SHEET_USERS].includes(name)) sh.hideSheet();
  }
  return sh;
}
function delRows_(sh, val, col) {
  if (!sh) return;
  const d = sh.getDataRange().getValues();
  for (let i = d.length-1; i >= 1; i--) if (d[i][col]==val) sh.deleteRow(i+1);
}
function typeCode_(n){return {ใบเสนอราคา:'QT',ใบเสร็จรับเงิน:'RT',ใบแจ้งหนี้:'IV',ใบกำกับภาษี:'TV',ใบส่งของ:'DO'}[n]||'QT';}
function typeName_(c){return {QT:'ใบเสนอราคา',RT:'ใบเสร็จรับเงิน',IV:'ใบแจ้งหนี้',TV:'ใบกำกับภาษี',DO:'ใบส่งของ'}[c]||'ใบเสนอราคา';}

// ── AUTH ─────────────────────────────────────────────────────
function login(username, password) {
  try {
    const sh   = getSheet_(SHEET_USERS,['username','password','role','name','active']);
    const rows = sh.getDataRange().getValues().slice(1);
    if (rows.length === 0) {
      sh.appendRow(['admin','admin1234','admin','เจ้าของ',true]);
      sh.appendRow(['staff1','staff1234','staff','พนักงาน 1',true]);
    }
    for (const r of sh.getDataRange().getValues().slice(1)) {
      if (r[0]==username && r[1]==password && r[4]!==false)
        return { ok:true, role:r[2], name:r[3] };
    }
    return { ok:false, error:'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch(e){ return { ok:false, error:e.message }; }
}

// ── USERS (admin only) ───────────────────────────────────────
function getUsers() {
  const sh = getSheet_(SHEET_USERS,[]);
  return sh.getDataRange().getValues().slice(1)
    .map(r=>({username:r[0],role:r[2],name:r[3],active:r[4]}));
}
function saveUser(u) {
  const sh   = getSheet_(SHEET_USERS,[]);
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][0]==u.username) {
      sh.getRange(i+1,1,1,5).setValues([[u.username,u.password||rows[i][1],u.role,u.name,u.active!==false]]);
      return {success:true};
    }
  }
  sh.appendRow([u.username,u.password||'staff1234',u.role||'staff',u.name||u.username,true]);
  return {success:true};
}
function deleteUser(username) {
  if (username==='admin') return {error:'ไม่สามารถลบ admin หลักได้'};
  const sh=getSheet_(SHEET_USERS,[]);
  const rows=sh.getDataRange().getValues();
  for (let i=rows.length-1;i>=1;i--) {
    if(rows[i][0]==username){sh.deleteRow(i+1);return {success:true};}
  }
  return {error:'ไม่พบผู้ใช้'};
}

// ── SETTINGS ─────────────────────────────────────────────────
function getSettings() {
  try {
    const sh=getSheet_(SHEET_CFG,['key','value']),map={};
    sh.getDataRange().getValues().slice(1).forEach(r=>{if(r[0])map[r[0]]=r[1];});
    return {logo:map['logo']||'',sig:map['sig']||''};
  } catch(e){return {logo:'',sig:'',error:e.message};}
}
function saveSettings(p) {
  try {
    const sh=getSheet_(SHEET_CFG,['key','value']);
    const rows=sh.getDataRange().getValues();
    const ups=(key,val)=>{
      if(val===undefined||val===null)return;
      for(let i=1;i<rows.length;i++){if(rows[i][0]===key){sh.getRange(i+1,2).setValue(val);return;}}
      sh.appendRow([key,val]);
    };
    ups('logo',p.logo); ups('sig',p.sig);
    return {success:true};
  } catch(e){return {error:e.message};}
}

// ── DOCUMENTS ────────────────────────────────────────────────
function getAllDocs() {
  try {
    const ss=SS();
    const shD=ss.getSheetByName(SHEET_DOCS),shI=ss.getSheetByName(SHEET_ITEMS);
    if(!shD||!shI) return {error:'ไม่พบ Sheet ฐานข้อมูล กรุณาสร้าง Sheet ก่อน'};
    const dR=shD.getDataRange().getValues(),iR=shI.getDataRange().getValues();
    const dH=dR[0],iH=iR[0],iMap={};
    iR.slice(1).forEach(r=>{
      const o={};iH.forEach((h,i)=>o[h]=r[i]);
      const n=o['เลขที่เอกสาร'];if(!n)return;
      if(!iMap[n])iMap[n]=[];
      iMap[n].push({name:o['สินค้า / บริการ']||'',qty:parseFloat(o['จำนวน'])||0,price:parseFloat(o['ราคาต่อหน่วย'])||0});
    });
    return dR.slice(1).filter(r=>r[0]).map(r=>{
      const o={};dH.forEach((h,i)=>o[h]=r[i]);
      const n=o['เลขที่เอกสาร'],items=iMap[n]||[];
      const total=items.reduce((s,it)=>s+it.qty*it.price,0);
      const hasVat=(o['มี VAT ?']||'').toString().includes('มี');
      let date=o['วันที่สร้าง'];
      if(date instanceof Date)date=Utilities.formatDate(date,'Asia/Bangkok','yyyy-MM-dd');
      return {number:n,type:typeCode_(o['ประเภทเอกสาร']),typeName:o['ประเภทเอกสาร']||'',
        date,ref:o['อ้างอิงเอกสาร']||'',customerName:o['ชื่อลูกค้า']||'',
        customerAddr1:o['ที่อยู่ลูกค้า 1']||'',customerAddr2:o['ที่อยู่ลูกค้า 2']||'',
        customerTax:o['เลขที่เสียภาษีลูกค้า']||'',customerPhone:o['เบอร์โทรศัพท์ลูกค้า']||'',
        note:o['หมายเหตุ']||'',hasVat,items,total};
    });
  } catch(e){return {error:e.message};}
}

function saveDoc(d) {
  try {
    const shD=getSheet_(SHEET_DOCS,['เลขที่เอกสาร','ประเภทเอกสาร','วันที่สร้าง','อ้างอิงเอกสาร','ชื่อลูกค้า','ที่อยู่ลูกค้า 1','ที่อยู่ลูกค้า 2','เลขที่เสียภาษีลูกค้า','เบอร์โทรศัพท์ลูกค้า','หมายเหตุ','ส่วนลด','มี VAT ?']);
    const shI=getSheet_(SHEET_ITEMS,['เลขที่เอกสาร','สินค้า / บริการ','จำนวน','ราคาต่อหน่วย']);
    const n=d.number,date=d.date?new Date(d.date):new Date();
    delRows_(shD,n,0);delRows_(shI,n,0);
    shD.appendRow([n,d.typeName||typeName_(d.type),date,d.ref||'',d.customerName||'',
      d.customerAddr1||'',d.customerAddr2||'',d.customerTax||'',d.customerPhone||'',
      d.note||'',0,d.hasVat?'มี VAT':'ไม่มี VAT']);
    (d.items||[]).forEach(it=>{if(it.name)shI.appendRow([n,it.name,it.qty,it.price]);});
    return {success:true,number:n};
  } catch(e){return {error:e.message};}
}

function deleteDoc(num) {
  try {
    const ss=SS();
    delRows_(ss.getSheetByName(SHEET_DOCS),num,0);
    delRows_(ss.getSheetByName(SHEET_ITEMS),num,0);
    return {success:true};
  } catch(e){return {error:e.message};}
}

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📄 ออกเอกสาร')
    .addItem('เปิดระบบ','openApp_').addToUi();
}
function openApp_() {
  const url=ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<script>window.open('${url}','_blank');google.script.host.close();</script>`).setWidth(1).setHeight(1),
    'กำลังเปิด...');
}
