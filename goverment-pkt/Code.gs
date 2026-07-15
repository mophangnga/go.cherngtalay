/*************************************************************
 * ระบบออกใบรับรองผู้ประกอบการ - องค์การบริหารส่วนตำบลเชิงทะเล
 * Backend (Google Apps Script)  v2
 *
 * เก็บข้อมูลใน Google Sheets + รูป (โลโก้/ลายเซ็น) ใน Google Drive
 * ชีต/โฟลเดอร์ทั้งหมดสร้างอัตโนมัติเมื่อใช้งานครั้งแรก
 *************************************************************/

var CONFIG_TITLE = 'ระบบออกใบรับรอง องค์การบริหารส่วนตำบลเชิงทะเล';
var SH_CERT = 'ใบรับรอง';
var SH_USER = 'ผู้ใช้งาน';
var SH_REG  = 'ทะเบียนผู้ประกอบการ';
var SH_APP  = 'คำขอใบอนุญาต';

/* ---------- ประเภทบริการ ---------- */
function getServiceTypes() {
  return [
    { code: 'ANM', name: 'การเลี้ยงสัตว์ (เพาะพันธุ์ เลี้ยง รวบรวมสัตว์)',            icon: 'paw',       color: '#C98A3D' },
    { code: 'APD', name: 'สัตว์และผลิตภัณฑ์ (ฆ่าสัตว์ ฟอกหนัง เคี่ยวไขมัน)',           icon: 'box',       color: '#C25B54' },
    { code: 'FOD', name: 'อาหาร เครื่องดื่ม น้ำดื่ม (ผลิต สะสม จำหน่าย)',              icon: 'utensils',  color: '#E8A23D' },
    { code: 'MED', name: 'ยา เวชภัณฑ์ เครื่องสำอาง (ผลิต สะสม แบ่งบรรจุ)',            icon: 'pill',      color: '#4FA3A0' },
    { code: 'AGR', name: 'การเกษตร (โรงสีข้าว อบ/รม/รมควันผลไม้ เก็บรักษาผลิตผล)',      icon: 'leaf',      color: '#5BAE6E' },
    { code: 'MTL', name: 'โลหะหรือแร่ (ถลุง หล่อ รีด ตัด กลึงโลหะ)',                  icon: 'hammer',    color: '#7A8794' },
    { code: 'MEC', name: 'ยานยนต์ เครื่องจักร (ทำ ซ่อม พ่นสี เครื่องกล)',             icon: 'gear',      color: '#5B8DC7' },
    { code: 'CON', name: 'ดิน หิน ทราย ซีเมนต์ (ทำ บด ผสม สะสมวัสดุก่อสร้าง)',         icon: 'brick',     color: '#B98A5E' },
    { code: 'WOD', name: 'ไม้ (เลื่อย ไส อบไม้ ผลิตภัณฑ์จากไม้)',                      icon: 'tree',      color: '#8A6D3B' },
    { code: 'CHM', name: 'สารเคมี วัตถุอันตราย (ผลิต สะสม ใช้วัตถุมีพิษ)',             icon: 'flask',     color: '#C77FB8' },
    { code: 'TEX', name: 'สิ่งทอและเครื่องหนัง (ปั่น ทอ ย้อม ฟอก ทำเครื่องหนัง)',       icon: 'shirt',     color: '#D97FA0' },
    { code: 'PPR', name: 'กระดาษ พลาสติก ยาง (ทำ สะสม หลอม)',                        icon: 'layers',    color: '#6FAF8E' },
    { code: 'OTH', name: 'อื่นๆ (พิมพ์ ท่าเทียบเรือ กำจัดแมลง บรรจุหีบห่อด้วยเครื่องจักร)', icon: 'dots',    color: '#8C9AA6' },
    { code: 'CTY', name: 'ลานสะสมตู้บรรจุสินค้า / ลานจอดรถหัวลากตู้บรรจุสินค้า',         icon: 'container', color: '#4F8BA0' }
  ];
}

var CERT_HEADERS = ['เลขที่ใบรับรอง','วันที่ออก','รหัสประเภท','ประเภทบริการ','ชื่อผู้ประกอบการ',
  'ชื่อสถานประกอบการ','เลขบัตร ปชช./ทะเบียน','ที่ตั้งสถานประกอบการ','เบอร์โทร','วันหมดอายุ',
  'สถานะ','ผู้ออกใบรับรอง','หมายเหตุ','บันทึกเมื่อ','ลิงก์เอกสาร','ค่าธรรมเนียม','สถานะชำระ','เลขที่ใบเสร็จ','พิกัด'];
var USER_HEADERS = ['ชื่อผู้ใช้','รหัสผ่าน(เข้ารหัส)','ชื่อ-สกุล','สิทธิ์','สร้างเมื่อ'];
var REG_HEADERS  = ['รหัส','ชื่อผู้ประกอบการ','ชื่อสถานประกอบการ','ประเภทกิจการ','เลขบัตร/ทะเบียน',
  'ที่ตั้ง','เบอร์โทร','วันที่ขึ้นทะเบียน','สถานะ','หมายเหตุ'];

/* ---------- หน้าเว็บ ---------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.cert) return renderVerify_(e.parameter.cert);
  if (e && e.parameter && e.parameter.lic) return renderVerifyLic_(e.parameter.lic);
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(CONFIG_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getWebAppUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }
function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }

/* ---------- สเปรดชีต / โฟลเดอร์ ---------- */
function getSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SS_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  ss = SpreadsheetApp.create(CONFIG_TITLE + ' (ฐานข้อมูล)');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}
function getSheet_(name, headers) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#0B3D4F').setFontColor('#fff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 150);
  }
  return sheet;
}
/* โฟลเดอร์แม่ของไฟล์ระบบ — สร้างไว้ "ที่เดียวกับสเปรดชีตฐานข้อมูล" */
function ssParentFolder_() {
  try {
    var ss = getSS_();
    var file = DriveApp.getFileById(ss.getId());
    var parents = file.getParents();
    if (parents.hasNext()) return parents.next();
  } catch (e) {}
  return DriveApp.getRootFolder();
}
function getFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var parent = ssParentFolder_();          // ← สร้างข้างสเปรดชีต
  var name = CONFIG_TITLE + ' - ไฟล์ระบบ';
  // ถ้ามีโฟลเดอร์ชื่อนี้อยู่แล้วในที่เดียวกัน ให้ใช้ตัวเดิม
  var it = parent.getFoldersByName(name);
  var f = it.hasNext() ? it.next() : parent.createFolder(name);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

/* ย้ายโฟลเดอร์ไฟล์ระบบไปอยู่ที่เดียวกับสเปรดชีต (เรียกได้จากหน้าตั้งค่า) */
function moveSystemFolderToSheet() {
  try {
    var folder = getFolder_();
    var target = ssParentFolder_();
    var fFile = DriveApp.getFolderById(folder.getId());

    // อยู่ถูกที่แล้วหรือยัง
    var ps = fFile.getParents();
    while (ps.hasNext()) { if (ps.next().getId() === target.getId()) return { ok: true, moved: false, folder: folder.getName(), where: target.getName() }; }

    target.addFolder(fFile);
    var olds = fFile.getParents();
    while (olds.hasNext()) { var p = olds.next(); if (p.getId() !== target.getId()) { try { p.removeFolder(fFile); } catch (e) {} } }
    return { ok: true, moved: true, folder: folder.getName(), where: target.getName() };
  } catch (e) { return { ok: false, error: e.message }; }
}
/* ข้อมูลตำแหน่งจัดเก็บ (แสดงในหน้าตั้งค่า) */
function getStorageInfo() {
  try {
    var folder = getFolder_();
    var target = ssParentFolder_();
    var slip = null;
    try { slip = getSlipFolder_(); } catch (e) {}
    var same = false;
    var ps = DriveApp.getFolderById(folder.getId()).getParents();
    while (ps.hasNext()) { if (ps.next().getId() === target.getId()) same = true; }
    return { ok: true, folderName: folder.getName(), folderUrl: folder.getUrl(),
             slipName: slip ? slip.getName() : '', slipUrl: slip ? slip.getUrl() : '',
             sheetFolder: target.getName(), sheetFolderUrl: target.getUrl(), sameLocation: same };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- ผู้ใช้งาน / ล็อกอิน ---------- */
function hash_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}
function seedAdmin_(sheet) {
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(['admin', hash_('admin1234'), 'ผู้ดูแลระบบ', 'admin',
      Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss')]);
  }
}
function login(username, password) {
  var sheet = getSheet_(SH_USER, USER_HEADERS);
  seedAdmin_(sheet);
  var data = sheet.getDataRange().getValues();
  var h = hash_(password);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(username) && String(data[i][1]) === h) {
      return { ok: true, user: { username: data[i][0], name: data[i][2], role: data[i][3] } };
    }
  }
  return { ok: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}
function getUsers() {
  var sheet = getSheet_(SH_USER, USER_HEADERS);
  seedAdmin_(sheet);
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({ row: i + 1, username: data[i][0], name: data[i][2], role: data[i][3], created: fmt_(data[i][4]) });
  }
  return out;
}
function addUser(u) {
  try {
    var sheet = getSheet_(SH_USER, USER_HEADERS);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++)
      if (String(data[i][0]) === String(u.username)) return { ok: false, error: 'มีชื่อผู้ใช้นี้แล้ว' };
    sheet.appendRow([u.username, hash_(u.password), u.name || '', u.role || 'staff',
      Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss')]);
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
}
function updateUser(u) {
  try {
    var sheet = getSheet_(SH_USER, USER_HEADERS);
    sheet.getRange(u.row, 1).setValue(u.username);
    sheet.getRange(u.row, 3).setValue(u.name || '');
    sheet.getRange(u.row, 4).setValue(u.role || 'staff');
    if (u.password) sheet.getRange(u.row, 2).setValue(hash_(u.password));
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
}
function deleteUser(row) {
  try {
    var sheet = getSheet_(SH_USER, USER_HEADERS);
    var data = sheet.getDataRange().getValues();
    var admins = 0;
    for (var i = 1; i < data.length; i++) if (data[i][3] === 'admin') admins++;
    if (data[row - 1] && data[row - 1][3] === 'admin' && admins <= 1)
      return { ok: false, error: 'ต้องมีผู้ดูแลระบบอย่างน้อย 1 คน' };
    sheet.deleteRow(row);
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- ตั้งค่า + รูปภาพ ----------
 * เก็บรูป (base64) ใน ScriptProperties โดยแบ่งเป็นชิ้นละ 8000 ตัวอักษร
 * -> อัปโหลดได้ทันทีโดยไม่ต้องขอสิทธิ์ Google Drive
 * และพยายามสำรองไฟล์ลง "โฟลเดอร์ระบบ" ใน Drive ให้ด้วย (ถ้าได้รับสิทธิ์)
 */
function saveBig_(key, str) {
  var p = PropertiesService.getScriptProperties();
  clearBig_(key);
  var size = 8000, n = Math.ceil(str.length / size), obj = {};
  obj[key + '_n'] = String(n);
  for (var i = 0; i < n; i++) obj[key + '_' + i] = str.substr(i * size, size);
  p.setProperties(obj, false);
}
function getBig_(key) {
  var p = PropertiesService.getScriptProperties();
  var n = parseInt(p.getProperty(key + '_n') || '0', 10);
  if (!n) return '';
  var out = '';
  for (var i = 0; i < n; i++) out += (p.getProperty(key + '_' + i) || '');
  return out;
}
function clearBig_(key) {
  var p = PropertiesService.getScriptProperties();
  var n = parseInt(p.getProperty(key + '_n') || '0', 10);
  for (var i = 0; i < n; i++) p.deleteProperty(key + '_' + i);
  p.deleteProperty(key + '_n');
}
/* สำรองรูปลงโฟลเดอร์ใน Drive (best-effort: ถ้าไม่มีสิทธิ์ก็ข้ามไป ไม่ทำให้อัปโหลดล้มเหลว) */
function backupToDrive_(kind, dataUrl) {
  try {
    var m = String(dataUrl).match(/^data:([^;]+);base64,(.*)$/);
    if (!m) return;
    var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], kind + '.png');
    var folder = getFolder_();
    var p = PropertiesService.getScriptProperties();
    var key = kind === 'logo' ? 'LOGO_ID' : 'SIGN_ID';
    var old = p.getProperty(key);
    if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
    var file = folder.createFile(blob);
    p.setProperty(key, file.getId());
  } catch (e) { /* ข้าม */ }
}
function getSettings() {
  var p = PropertiesService.getScriptProperties();
  return {
    orgName:    p.getProperty('orgName')    || 'องค์การบริหารส่วนตำบลเชิงทะเล',
    orgSub:     p.getProperty('orgSub')     || 'อำเภอถลาง จังหวัดภูเก็ต',
    mayorName:  p.getProperty('mayorName')  || '',
    mayorTitle: p.getProperty('mayorTitle') || 'นายกองค์การบริหารส่วนตำบลเชิงทะเล',
    receiverName: p.getProperty('receiverName') || '',
    feeHaz: p.getProperty('feeHaz') || '', feeFood: p.getProperty('feeFood') || '',
    feeMarket: p.getProperty('feeMarket') || '', feeFoodnotify: p.getProperty('feeFoodnotify') || '',
    ppEnabled: p.getProperty('ppEnabled') || '',
    ppType:    p.getProperty('ppType')    || 'taxid',
    ppId:      p.getProperty('ppId')      || '',
    ppName:    p.getProperty('ppName')    || '',
    ppBank:    p.getProperty('ppBank')    || '',
    ppAccount: p.getProperty('ppAccount') || '',
    ppNote:    p.getProperty('ppNote')    || '',
    slipEnabled: p.getProperty('slipEnabled') || '',
    themeColor: p.getProperty('themeColor') || '#0E7C86',
    notifyEmail:p.getProperty('notifyEmail') || '',
    notifyDays: p.getProperty('notifyDays')  || '30',
    logoUrl:    getBig_('LOGO'),
    signUrl:    getBig_('SIGN'),
    garudaUrl:  getBig_('GARUDA'),
    sealUrl:    getBig_('SEAL'),
    loginBgUrl: getBig_('LOGINBG')
  };
}
function saveSettings(s) {
  try {
    var p = PropertiesService.getScriptProperties();
    ['orgName','orgSub','mayorName','mayorTitle','receiverName','feeHaz','feeFood','feeMarket','feeFoodnotify','ppEnabled','ppType','ppId','ppName','ppBank','ppAccount','ppNote','slipEnabled','themeColor','notifyEmail','notifyDays'].forEach(function (k) {
      if (s[k] != null) p.setProperty(k, String(s[k]));
    });
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
}
/* kind = 'logo' | 'sign' | 'garuda' ; dataUrl = "data:image/png;base64,...." */
function imgKey_(kind){ var m={logo:'LOGO',sign:'SIGN',garuda:'GARUDA',seal:'SEAL',loginbg:'LOGINBG'}; return m[kind]||'LOGO'; }
function getLoginBg() { try { return getBig_('LOGINBG') || ''; } catch (e) { return ''; } }
function setLoginBgUrl(url) { try { saveBig_('LOGINBG', String(url || '')); return { ok: true, url: String(url || '') }; } catch (e) { return { ok: false, error: e.message }; } }
function uploadImage(kind, dataUrl) {
  try {
    if (!/^data:[^;]+;base64,/.test(String(dataUrl))) return { ok: false, error: 'รูปแบบไฟล์ไม่ถูกต้อง' };
    saveBig_(imgKey_(kind), dataUrl);
    backupToDrive_(kind, dataUrl);
    return { ok: true, url: dataUrl };
  } catch (e) { return { ok: false, error: e.message }; }
}
function removeImage(kind) {
  try {
    clearBig_(imgKey_(kind));
    var p = PropertiesService.getScriptProperties();
    var key = imgKey_(kind) + '_ID';
    var old = p.getProperty(key);
    if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
    p.deleteProperty(key);
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- ใบรับรอง ---------- */
function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'GMT+7', 'yyyy-MM-dd');
  return v ? String(v) : '';
}
function generateCertNo_(code, sheet) {
  var year = new Date().getFullYear() + 543;
  var data = sheet.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    var no = String(data[i][0] || '');
    if (no.indexOf(code + '-') === 0 && no.indexOf('/' + year) > -1) n++;
  }
  return code + '-' + ('0000' + (n + 1)).slice(-4) + '/' + year;
}
/* ---------- สร้างเอกสาร PDF เก็บใน Drive + ลิงก์สำหรับ QR ---------- */
function esc_(s){ return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }
function thDate_(str){
  if(!str) return '-';
  var d = (str instanceof Date) ? str : new Date(str);
  if(isNaN(d)) return String(str);
  var m=['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return d.getDate()+' '+m[d.getMonth()]+' พ.ศ. '+(d.getFullYear()+543);
}
function fileIdFromUrl_(url){ var m=String(url||'').match(/\/d\/([^\/]+)/); return m?m[1]:''; }
function ensureLinkHeader_(sheet){
  try{ for(var i=14;i<CERT_HEADERS.length;i++){ if(!sheet.getRange(1,i+1).getValue()) sheet.getRange(1,i+1).setValue(CERT_HEADERS[i]); } }catch(e){}
}
function certDocHtml_(c, qrDataUrl){
  var s = getSettings();
  var logoTag = s.logoUrl ? '<img src="'+s.logoUrl+'" style="width:96px;height:96px;object-fit:contain">' : '';
  var signTag = s.signUrl ? '<img src="'+s.signUrl+'" style="height:64px;object-fit:contain"><br>' : '<br><br>';
  var qrTag = qrDataUrl ? '<div style="margin-top:18px"><img src="'+qrDataUrl+'" style="width:96px;height:96px"><div style="font-size:11px;color:#888">สแกนเพื่อตรวจสอบเอกสาร</div></div>' : '';
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    + 'body{font-family:sans-serif;color:#14323a;margin:0}'
    + '.wrap{border:6px double #C8A24B;margin:16px;padding:34px 40px;text-align:center}'
    + 'h1{font-size:24px;margin:8px 0 2px;color:#0B3D4F}'
    + '.sub{color:#666;font-size:13px}'
    + '.hr{height:1px;background:#E4C97E;margin:14px auto;width:55%}'
    + '.title{font-size:22px;color:#0B3D4F;font-weight:bold;margin:8px 0 6px}'
    + '.type{display:inline-block;border:1px solid #0E7C86;color:#0E7C86;border-radius:16px;padding:4px 16px;font-size:15px;margin-bottom:16px}'
    + '.lead{font-size:17px;line-height:2.1}'
    + '.lead b{color:#0B3D4F}'
    + '.foot{margin-top:40px;width:100%;border-collapse:collapse}'
    + '.foot td{font-size:14px;color:#444;vertical-align:bottom;text-align:center;width:50%}'
    + '.no{margin-top:16px;font-size:12px;color:#888}'
    + '</style></head><body><div class="wrap">'
    + logoTag
    + '<h1>'+esc_(s.orgName)+'</h1><div class="sub">'+esc_(s.orgSub)+'</div>'
    + '<div class="hr"></div>'
    + '<div class="title">ใบรับรองการประกอบกิจการ</div>'
    + '<div class="type">'+esc_(c.serviceName)+'</div>'
    + '<div class="lead">หนังสือฉบับนี้ให้ไว้เพื่อรับรองว่า<br><b>'+esc_(c.operatorName||'-')+'</b><br>'
    + 'ผู้ประกอบกิจการ <b>'+esc_(c.businessName||'-')+'</b><br>'
    + 'ซึ่งตั้งอยู่ '+esc_(c.address||'-')+'<br>'
    + 'ได้ผ่านการตรวจประเมินและรับรองตามหลักเกณฑ์ของ'+esc_(s.orgName)+'</div>'
    + '<table class="foot"><tr>'
    + '<td>ให้ไว้ ณ วันที่<br><b>'+thDate_(c.issueDate)+'</b>'+(c.expiryDate?'<br>มีผลถึงวันที่ '+thDate_(c.expiryDate):'')+'</td>'
    + '<td>'+signTag+'('+esc_(s.mayorName||'..............................')+')<br>'+esc_(s.mayorTitle)+'</td>'
    + '</tr></table>'
    + '<div class="no">เลขที่ '+esc_(c.certNo)+'</div>'
    + qrTag
    + '</div></body></html>';
}
function buildCertPdf_(c, qrDataUrl){
  var html = certDocHtml_(c, qrDataUrl);
  var pdf = Utilities.newBlob(html, 'text/html', 'cert.html').getAs('application/pdf')
    .setName('ใบรับรอง ' + String(c.certNo).replace(/\//g, '-') + '.pdf');
  var file = getFolder_().createFile(pdf);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { url: file.getUrl(), id: file.getId() };
}
/* สร้าง/อัปเดต PDF ของใบรับรอง (เรียกจากฝั่งเว็บ พร้อม QR) — ต้องมีสิทธิ์ Drive */
function makeCertPdf(row, qrDataUrl){
  try{
    var sheet = getSheet_(SH_CERT, CERT_HEADERS);
    ensureLinkHeader_(sheet);
    var r = sheet.getRange(row, 1, 1, CERT_HEADERS.length).getValues()[0];
    var cert = { certNo:r[0], issueDate:fmt_(r[1]), serviceName:r[3], operatorName:r[4],
      businessName:r[5], address:r[7], expiryDate:fmt_(r[9]) };
    var old = r[14];
    if (old) { var oid = fileIdFromUrl_(old); if (oid) { try { DriveApp.getFileById(oid).setTrashed(true); } catch (e3) {} } }
    var info = buildCertPdf_(cert, qrDataUrl);
    sheet.getRange(row, 15).setValue(info.url);
    return { ok:true, docUrl: info.url };
  } catch (e) { return { ok:false, error:e.message }; }
}

/* ---------- หน้าตรวจสอบสาธารณะ (เปิดจาก QR) ---------- */
function findCertByNo_(no){
  var sheet = getSheet_(SH_CERT, CERT_HEADERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(no)) {
      var r = data[i];
      return { certNo:r[0], issueDate:fmt_(r[1]), serviceCode:r[2], serviceName:r[3], operatorName:r[4],
        businessName:r[5], idNumber:r[6], address:r[7], phone:r[8], expiryDate:fmt_(r[9]),
        status:r[10], issuedBy:r[11], note:r[12], docUrl:r[14] || '', coords:r[18] || '' };
    }
  }
  return null;
}
function vRow_(k,v){ return '<tr><td>'+esc_(k)+'</td><td>'+esc_(v||'-')+'</td></tr>'; }
function verifyCss_(){
  return 'body{font-family:sans-serif;background:#eef5f5;margin:0;padding:18px;color:#14323a}'
   + '.card{max-width:460px;margin:18px auto;background:#fff;border-radius:16px;padding:26px;box-shadow:0 16px 44px -18px rgba(0,0,0,.3);text-align:center}'
   + '.ok{width:66px;height:66px;line-height:66px;border-radius:50%;background:#2E9E6B;color:#fff;font-size:36px;margin:0 auto 8px}'
   + '.bad{width:66px;height:66px;line-height:66px;border-radius:50%;background:#D65A4F;color:#fff;font-size:36px;margin:0 auto 8px}'
   + 'h2{margin:6px 0;color:#0B3D4F}.muted{color:#777;font-size:13px;margin:0 0 12px}'
   + '.logo{width:74px;height:74px;object-fit:contain;margin:8px auto;display:block}'
   + '.title{font-weight:bold;color:#0E7C86;font-size:18px;margin-top:8px}'
   + '.type{display:inline-block;border:1px solid #0E7C86;color:#0E7C86;border-radius:14px;padding:3px 14px;font-size:13px;margin:8px 0 14px}'
   + 'table{width:100%;border-collapse:collapse;text-align:left;font-size:14px}'
   + 'td{padding:9px 6px;border-bottom:1px solid #eee;vertical-align:top}'
   + 'td:first-child{color:#888;width:40%}'
   + '.btn{display:inline-block;margin-top:18px;background:#0E7C86;color:#fff;text-decoration:none;padding:11px 22px;border-radius:10px;font-weight:bold}'
   + '.foot{margin-top:16px;font-size:12px;color:#aaa}';
}
function renderVerify_(no){
  var s = getSettings();
  var c = findCertByNo_(no);
  var head = '<!DOCTYPE html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ตรวจสอบใบรับรอง</title><style>'+verifyCss_()+'</style></head><body>';
  var html;
  if (!c) {
    html = head + '<div class="card"><div class="bad">&#10007;</div><h2>ไม่พบใบรับรอง</h2>'
      + '<p class="muted">ไม่พบใบรับรองเลขที่ ' + esc_(no) + ' ในระบบของ ' + esc_(s.orgName) + '</p></div></body></html>';
  } else {
    var expired = false;
    if (c.expiryDate) { var d = new Date(c.expiryDate); if (!isNaN(d) && d < new Date()) expired = true; }
    var stColor = (c.status === 'ใช้งาน' && !expired) ? '#2E9E6B' : (c.status === 'ยกเลิก' ? '#E0972C' : '#D65A4F');
    var driveBtn = c.docUrl ? '<a class="btn" href="' + c.docUrl + '" target="_blank">เปิดไฟล์ใบรับรอง (PDF)</a>' : '';
    html = head + '<div class="card">'
      + '<div class="ok">&#10003;</div><h2>เอกสารถูกต้อง</h2>'
      + '<p class="muted">ออกโดย ' + esc_(s.orgName) + '</p>'
      + (s.logoUrl ? '<img class="logo" src="' + s.logoUrl + '">' : '')
      + '<div class="title">ใบรับรองการประกอบกิจการ</div>'
      + '<div class="type">' + esc_(c.serviceName) + '</div>'
      + '<table>'
      + vRow_('เลขที่ใบรับรอง', c.certNo)
      + vRow_('ผู้ประกอบการ', c.operatorName)
      + vRow_('สถานประกอบการ', c.businessName)
      + vRow_('ที่ตั้ง', c.address)
      + vRow_('วันที่ออก', thDate_(c.issueDate))
      + vRow_('วันหมดอายุ', c.expiryDate ? thDate_(c.expiryDate) : '-')
      + '<tr><td>สถานะ</td><td><b style="color:' + stColor + '">' + esc_(c.status) + (expired ? ' (เลยกำหนด)' : '') + '</b></td></tr>'
      + '</table>'
      + (c.coords ? '<div style="margin-top:14px"><iframe title="แผนที่" width="100%" height="200" style="border:0;border-radius:12px" loading="lazy" referrerpolicy="no-referrer-when-downgrade" src="https://maps.google.com/maps?q=' + encodeURIComponent(c.coords) + '&z=16&hl=th&output=embed"></iframe></div>'
          + '<a class="btn" style="margin-top:10px" href="https://www.google.com/maps?q=' + encodeURIComponent(c.coords) + '" target="_blank">เปิดตำแหน่งในแผนที่</a>' : '')
      + driveBtn
      + '<p class="foot">ตรวจสอบเมื่อ ' + thDate_(Utilities.formatDate(new Date(),'GMT+7','yyyy-MM-dd')) + '</p>'
      + '</div></body></html>';
  }
  return HtmlService.createHtmlOutput(html).setTitle('ตรวจสอบใบรับรอง');
}

function saveCertificate(f) {
  try {
    var sheet = getSheet_(SH_CERT, CERT_HEADERS);
    ensureLinkHeader_(sheet);
    var t = getServiceTypes().filter(function (x) { return x.code === f.serviceCode; })[0];
    if (!t) throw new Error('ไม่พบประเภทบริการ');
    var certNo = generateCertNo_(f.serviceCode, sheet);
    var now = new Date();
    var issueDate = f.issueDate || Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    sheet.appendRow([certNo, issueDate, f.serviceCode, t.name, f.operatorName || '', f.businessName || '',
      f.idNumber || '', f.address || '', f.phone || '', f.expiryDate || '', 'ใช้งาน',
      f.issuedBy || '', f.note || '', Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd HH:mm:ss')]);
    var row = sheet.getLastRow();
    // คอลัมน์ ค่าธรรมเนียม/สถานะชำระ/เลขที่ใบเสร็จ (16-18)
    sheet.getRange(row, 16, 1, 3).setValues([[ f.fee || '', f.paid || '', f.receiptNo || '' ]]);
    sheet.getRange(row, 19).setValue(f.coords || '');
    var verifyUrl = getWebAppUrl_(); if (verifyUrl) verifyUrl += (verifyUrl.indexOf('?')>-1?'&':'?') + 'cert=' + encodeURIComponent(certNo);
    return { ok: true, certNo: certNo, row: row, verifyUrl: verifyUrl };
  } catch (e) { return { ok: false, error: e.message }; }
}
function updateCertificate(f) {
  try {
    var sheet = getSheet_(SH_CERT, CERT_HEADERS);
    ensureLinkHeader_(sheet);
    var t = getServiceTypes().filter(function (x) { return x.code === f.serviceCode; })[0];
    // คอลัมน์ 2..13
    sheet.getRange(f.row, 2, 1, 12).setValues([[
      f.issueDate || '', f.serviceCode, t ? t.name : f.serviceName, f.operatorName || '',
      f.businessName || '', f.idNumber || '', f.address || '', f.phone || '',
      f.expiryDate || '', f.status || 'ใช้งาน', f.issuedBy || '', f.note || ''
    ]]);
    // ค่าธรรมเนียม/สถานะชำระ/เลขที่ใบเสร็จ (16-18)
    sheet.getRange(f.row, 16, 1, 3).setValues([[ f.fee || '', f.paid || '', f.receiptNo || '' ]]);
    sheet.getRange(f.row, 19).setValue(f.coords || '');
    var certNo = sheet.getRange(f.row, 1).getValue();
    var verifyUrl = getWebAppUrl_(); if (verifyUrl) verifyUrl += (verifyUrl.indexOf('?')>-1?'&':'?') + 'cert=' + encodeURIComponent(certNo);
    return { ok: true, row: f.row, verifyUrl: verifyUrl };
  } catch (e) { return { ok: false, error: e.message }; }
}
function getCertificates() {
  var sheet = getSheet_(SH_CERT, CERT_HEADERS);
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i]; if (!r[0]) continue;
    out.push({ row: i + 1, certNo: r[0], issueDate: fmt_(r[1]), serviceCode: r[2], serviceName: r[3],
      operatorName: r[4], businessName: r[5], idNumber: r[6], address: r[7], phone: r[8],
      expiryDate: fmt_(r[9]), status: r[10], issuedBy: r[11], note: r[12], docUrl: r[14] || '',
      fee: r[15] || '', paid: r[16] || '', receiptNo: r[17] || '', coords: r[18] || '' });
  }
  return out.reverse();
}
function updateStatus(row, status) {
  try { getSheet_(SH_CERT, CERT_HEADERS).getRange(row, 11).setValue(status); return { ok: true }; }
  catch (e) { return { ok: false, error: e.message }; }
}
function deleteCertificate(row) {
  try { getSheet_(SH_CERT, CERT_HEADERS).deleteRow(row); return { ok: true }; }
  catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- ทะเบียนผู้ประกอบการ ---------- */
function getOperators() {
  var sheet = getSheet_(SH_REG, REG_HEADERS);
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i]; if (!r[1] && !r[2]) continue;
    out.push({ row: i + 1, code: r[0], operatorName: r[1], businessName: r[2], bizType: r[3],
      idNumber: r[4], address: r[5], phone: r[6], regDate: fmt_(r[7]), status: r[8] || 'ปกติ', note: r[9] });
  }
  return out.reverse();
}
function addOperator(o) {
  try {
    var sheet = getSheet_(SH_REG, REG_HEADERS);
    var code = 'OP-' + ('0000' + Math.max(0, sheet.getLastRow() - 1) + 1).slice(-4);
    sheet.appendRow([code, o.operatorName || '', o.businessName || '', o.bizType || '', o.idNumber || '',
      o.address || '', o.phone || '', o.regDate || Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd'),
      o.status || 'ปกติ', o.note || '']);
    return { ok: true, code: code };
  } catch (e) { return { ok: false, error: e.message }; }
}
function updateOperator(o) {
  try {
    var sheet = getSheet_(SH_REG, REG_HEADERS);
    sheet.getRange(o.row, 2, 1, 9).setValues([[o.operatorName || '', o.businessName || '', o.bizType || '',
      o.idNumber || '', o.address || '', o.phone || '', o.regDate || '', o.status || 'ปกติ', o.note || '']]);
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
}
function deleteOperator(row) {
  try { getSheet_(SH_REG, REG_HEADERS).deleteRow(row); return { ok: true }; }
  catch (e) { return { ok: false, error: e.message }; }
}
function addOperators(list) {
  try {
    var sheet = getSheet_(SH_REG, REG_HEADERS);
    var base = Math.max(0, sheet.getLastRow() - 1);
    var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
    var rows = (list || []).map(function (o, i) {
      return ['OP-' + ('0000' + (base + i + 1)).slice(-4), o.operatorName || '', o.businessName || '',
        o.bizType || '', o.idNumber || '', o.address || '', o.phone || '', o.regDate || today,
        o.status || 'ปกติ', o.note || ''];
    });
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, REG_HEADERS.length).setValues(rows);
    return { ok: true, count: rows.length };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- บันทึกการใช้งาน (audit log) ---------- */
var SH_LOG = 'บันทึกการใช้งาน';
var LOG_HEADERS = ['เวลา', 'ผู้ใช้', 'การกระทำ', 'รายละเอียด'];
function addLog(user, action, detail) {
  try {
    getSheet_(SH_LOG, LOG_HEADERS).appendRow([
      Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss'),
      user || '-', action || '', detail || '']);
    return { ok: true };
  } catch (e) { return { ok: false }; }
}
function getLogs(limit) {
  var sheet = getSheet_(SH_LOG, LOG_HEADERS);
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({ time: fmtDT_(data[i][0]), user: data[i][1], action: data[i][2], detail: data[i][3] });
  }
  out.reverse();
  return out.slice(0, limit || 200);
}
function fmtDT_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  return v ? String(v) : '';
}

/* ---------- เปลี่ยนรหัสผ่านตนเอง ---------- */
function changeOwnPassword(username, oldPass, newPass) {
  try {
    if (!newPass || String(newPass).length < 4) return { ok: false, error: 'รหัสผ่านใหม่ต้องยาวอย่างน้อย 4 ตัวอักษร' };
    var sheet = getSheet_(SH_USER, USER_HEADERS);
    var data = sheet.getDataRange().getValues();
    var ho = hash_(oldPass);
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(username)) {
        if (String(data[i][1]) !== ho) return { ok: false, error: 'รหัสผ่านเดิมไม่ถูกต้อง' };
        sheet.getRange(i + 1, 2).setValue(hash_(newPass));
        return { ok: true };
      }
    }
    return { ok: false, error: 'ไม่พบผู้ใช้' };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- แจ้งเตือนใกล้หมดอายุทางอีเมล ---------- */
function notifyExpiring() {
  try {
    var s = getSettings();
    if (!s.notifyEmail) return { ok: false, error: 'ยังไม่ได้ตั้งค่าอีเมลผู้รับในหน้าตั้งค่า' };
    var days = parseInt(s.notifyDays || '30', 10); if (isNaN(days) || days <= 0) days = 30;
    var today = new Date(); today.setHours(0,0,0,0);
    var limit = new Date(today); limit.setDate(today.getDate() + days);
    var list = getCertificates().filter(function (c) {
      if (c.status !== 'ใช้งาน' || !c.expiryDate) return false;
      var d = new Date(c.expiryDate); if (isNaN(d)) return false;
      return d <= limit; // รวมที่เลยกำหนดด้วย
    }).sort(function (a, b) { return new Date(a.expiryDate) - new Date(b.expiryDate); });

    // ใบอนุญาต/หนังสือรับรอง (จากคำขอที่อนุมัติ)
    var licList = [];
    try {
      licList = getApplications().filter(function (a) {
        if (a.status !== 'อนุมัติ' || !a.expireDate) return false;
        var d = new Date(a.expireDate); if (isNaN(d)) return false;
        return d <= limit;
      }).sort(function (a, b) { return new Date(a.expireDate) - new Date(b.expireDate); });
    } catch (e) { licList = []; }

    if (!list.length && !licList.length) {
      MailApp.sendEmail({ to: s.notifyEmail, subject: '[' + s.orgName + '] ไม่มีเอกสารใกล้หมดอายุ',
        htmlBody: 'ขณะนี้ไม่มีใบรับรองหรือใบอนุญาตที่ใกล้หมดอายุภายใน ' + days + ' วัน' });
      return { ok: true, count: 0 };
    }
    var rows = list.map(function (c) {
      var d = new Date(c.expiryDate); var diff = Math.ceil((d - today) / 86400000);
      var note = diff < 0 ? ('<span style="color:#c0392b">เลยกำหนด ' + Math.abs(diff) + ' วัน</span>') : ('เหลือ ' + diff + ' วัน');
      return '<tr><td style="padding:6px 10px;border:1px solid #ddd">' + c.certNo + '</td>'
        + '<td style="padding:6px 10px;border:1px solid #ddd">' + c.businessName + '</td>'
        + '<td style="padding:6px 10px;border:1px solid #ddd">' + c.operatorName + '</td>'
        + '<td style="padding:6px 10px;border:1px solid #ddd">' + c.expiryDate + '</td>'
        + '<td style="padding:6px 10px;border:1px solid #ddd">' + note + '</td>'
        + '<td style="padding:6px 10px;border:1px solid #ddd">' + (c.phone || '-') + '</td></tr>';
    }).join('');
    var html = '<div style="font-family:sans-serif;color:#222">'
      + '<h2 style="color:#0B3D4F">แจ้งเตือนเอกสารใกล้หมดอายุ</h2>';
    if (list.length) {
      html += '<p>' + esc_(s.orgName) + ' — ใบรับรอง <b>' + list.length + '</b> รายการที่จะหมดอายุภายใน ' + days + ' วัน</p>'
      + '<table style="border-collapse:collapse;font-size:14px"><tr style="background:#0E7C86;color:#fff">'
      + '<th style="padding:8px 10px">เลขที่</th><th style="padding:8px 10px">สถานประกอบการ</th>'
      + '<th style="padding:8px 10px">ผู้ประกอบการ</th><th style="padding:8px 10px">วันหมดอายุ</th>'
      + '<th style="padding:8px 10px">คงเหลือ</th><th style="padding:8px 10px">โทร</th></tr>'
      + rows + '</table>';
    }
    if (licList.length) {
      var lrows = licList.map(function (a) {
        var d = new Date(a.expireDate); var diff = Math.ceil((d - today) / 86400000);
        var note = diff < 0 ? ('<span style="color:#c0392b">เลยกำหนด ' + Math.abs(diff) + ' วัน</span>') : ('เหลือ ' + diff + ' วัน');
        return '<tr><td style="padding:6px 10px;border:1px solid #ddd">' + a.appNo + '</td>'
          + '<td style="padding:6px 10px;border:1px solid #ddd">' + (a.bizName || '-') + '</td>'
          + '<td style="padding:6px 10px;border:1px solid #ddd">' + (a.applicantName || '-') + '</td>'
          + '<td style="padding:6px 10px;border:1px solid #ddd">' + a.expireDate + '</td>'
          + '<td style="padding:6px 10px;border:1px solid #ddd">' + note + '</td></tr>';
      }).join('');
      html += '<h3 style="color:#0B3D4F;margin-top:18px">ใบอนุญาต/หนังสือรับรองใกล้หมดอายุ (' + licList.length + ' รายการ)</h3>'
        + '<table style="border-collapse:collapse;font-size:14px"><tr style="background:#B7791F;color:#fff">'
        + '<th style="padding:8px 10px">เลขที่คำขอ</th><th style="padding:8px 10px">สถานประกอบการ</th>'
        + '<th style="padding:8px 10px">ผู้ขอ</th><th style="padding:8px 10px">วันหมดอายุ</th>'
        + '<th style="padding:8px 10px">คงเหลือ</th></tr>'
        + lrows + '</table>';
    }
    html += '<p style="color:#888;font-size:12px;margin-top:14px">อีเมลอัตโนมัติจากระบบออกใบรับรอง ' + esc_(s.orgName) + '</p></div>';
    var totalN = list.length + licList.length;
    MailApp.sendEmail({ to: s.notifyEmail, subject: '[' + s.orgName + '] เอกสารใกล้หมดอายุ ' + totalN + ' รายการ', htmlBody: html });
    return { ok: true, count: totalN };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- สำรองข้อมูล (ใช้สิทธิ์ Spreadsheet ไม่ต้องใช้ Drive) ---------- */
function backupNow() {
  try {
    var ss = getSS_();
    var name = CONFIG_TITLE + ' - สำรอง ' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH-mm');
    var copy = ss.copy(name);            // สร้างสำเนาด้วยสิทธิ์ Spreadsheet
    var url = '';
    try { url = copy.getUrl(); } catch (e) {}
    // พยายามย้ายเข้าโฟลเดอร์ระบบถ้ามีสิทธิ์ Drive (ถ้าไม่มีก็อยู่ใน "ไดรฟ์ของฉัน")
    try {
      var file = DriveApp.getFileById(copy.getId());
      getFolder_().addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (e2) {}
    return { ok: true, name: name, url: url };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- สำรองอัตโนมัติรายสัปดาห์ (Trigger) ---------- */
var BACKUP_HANDLER = 'autoBackup';
function autoBackup() {
  var res = backupNow();
  var p = PropertiesService.getScriptProperties();
  p.setProperty('BK_LAST', Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm'));
  p.setProperty('BK_LAST_OK', res && res.ok ? '1' : '0');
  cleanOldBackups_();
  // แจ้งผลทางอีเมล (ถ้าตั้งค่าอีเมลไว้)
  try {
    var s = getSettings();
    if (s.notifyEmail) {
      MailApp.sendEmail({ to: s.notifyEmail,
        subject: '[' + s.orgName + '] ผลการสำรองข้อมูลอัตโนมัติ',
        htmlBody: res && res.ok
          ? ('สำรองข้อมูลสำเร็จ<br>ไฟล์: ' + res.name + (res.url ? ('<br><a href="' + res.url + '">เปิดไฟล์สำรอง</a>') : ''))
          : ('<b style="color:#c0392b">สำรองข้อมูลไม่สำเร็จ</b><br>' + (res ? res.error : '')) });
    }
  } catch (e) {}
  return res;
}
/* เก็บไฟล์สำรองล่าสุด N ชุด (ค่าเริ่มต้น 8 สัปดาห์) */
function cleanOldBackups_() {
  try {
    var keep = parseInt(PropertiesService.getScriptProperties().getProperty('BK_KEEP') || '8', 10);
    var folder = getFolder_();
    var it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    var arr = [];
    while (it.hasNext()) {
      var f = it.next();
      if (f.getName().indexOf(' - สำรอง ') > -1) arr.push({ f: f, t: f.getDateCreated().getTime() });
    }
    arr.sort(function (a, b) { return b.t - a.t; });
    for (var i = keep; i < arr.length; i++) { try { arr[i].f.setTrashed(true); } catch (e) {} }
  } catch (e) {}
}
function getBackupStatus() {
  try {
    var enabled = false, day = '', hour = '';
    var t = ScriptApp.getProjectTriggers();
    for (var i = 0; i < t.length; i++) if (t[i].getHandlerFunction() === BACKUP_HANDLER) enabled = true;
    var p = PropertiesService.getScriptProperties();
    return { enabled: enabled, last: p.getProperty('BK_LAST') || '', lastOk: p.getProperty('BK_LAST_OK') === '1',
             day: p.getProperty('BK_DAY') || '1', hour: p.getProperty('BK_HOUR') || '2', keep: p.getProperty('BK_KEEP') || '8' };
  } catch (e) { return { enabled: false, error: e.message }; }
}
function enableWeeklyBackup(day, hour, keep) {
  try {
    disableWeeklyBackup();
    var h = parseInt(hour, 10); if (isNaN(h) || h < 0 || h > 23) h = 2;
    var d = parseInt(day, 10); if (isNaN(d) || d < 1 || d > 7) d = 1;
    var days = [ScriptApp.WeekDay.SUNDAY, ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY,
                ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY, ScriptApp.WeekDay.SATURDAY];
    ScriptApp.newTrigger(BACKUP_HANDLER).timeBased().onWeekDay(days[d % 7]).atHour(h).create();
    var p = PropertiesService.getScriptProperties();
    p.setProperty('BK_DAY', String(d)); p.setProperty('BK_HOUR', String(h));
    p.setProperty('BK_KEEP', String(parseInt(keep, 10) || 8));
    return { ok: true, enabled: true, day: d, hour: h };
  } catch (e) { return { ok: false, error: e.message }; }
}
function disableWeeklyBackup() {
  try {
    var t = ScriptApp.getProjectTriggers();
    for (var i = 0; i < t.length; i++) if (t[i].getHandlerFunction() === BACKUP_HANDLER) ScriptApp.deleteTrigger(t[i]);
    return { ok: true, enabled: false };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- ขออนุญาตสิทธิ์ (Run ฟังก์ชันนี้ครั้งเดียวในตัวแก้ไข) ----------
   สำคัญ: ห้ามใส่ try/catch ในฟังก์ชันนี้
   เพราะ Apps Script จะไม่ขอสิทธิ์ถ้าโค้ดถูกครอบด้วย try/catch ทั้งหมด        */
function AUTHORIZE() {
  // Drive — บังคับขอสิทธิ์ https://www.googleapis.com/auth/drive
  var folder = DriveApp.createFolder('ทดสอบสิทธิ์ - ลบได้');
  folder.setTrashed(true);
  DriveApp.getRootFolder().getName();

  // Mail
  MailApp.getRemainingDailyQuota();

  // Trigger
  ScriptApp.getProjectTriggers();

  // Spreadsheet
  SpreadsheetApp.getActiveSpreadsheet().getName();

  Logger.log('✅ อนุญาตสิทธิ์ครบถ้วนแล้ว — Drive / Mail / Trigger / Spreadsheet');
  return '✅ อนุญาตสิทธิ์ครบถ้วนแล้ว';
}

/* ---------- ตรวจสอบสิทธิ์ที่ระบบต้องใช้ ---------- */
function checkPermissions() {
  var out = { drive: false, mail: false, trigger: false, errors: [] };
  try { var f = getSlipFolder_(); out.drive = !!f; }
  catch (e) { out.errors.push('Drive: ' + e.message); }
  try { MailApp.getRemainingDailyQuota(); out.mail = true; }
  catch (e) { out.errors.push('Mail: ' + e.message); }
  try { ScriptApp.getProjectTriggers(); out.trigger = true; }
  catch (e) { out.errors.push('Trigger: ' + e.message); }
  out.ok = out.drive && out.mail && out.trigger;
  return out;
}

/* ---------- สลิปการโอนเงิน (เก็บใน Google Drive) ---------- */
function getSlipFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SLIP_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var parent = getFolder_();
  var f = parent.createFolder('สลิปการโอนเงิน');
  props.setProperty('SLIP_FOLDER_ID', f.getId());
  return f;
}
/* อัปโหลดสลิป: เก็บไฟล์ลง Drive + บันทึก id/url ใน JSON ของคำขอ */
function uploadSlip(row, dataUrl, filename, username) {
  try {
    var m = String(dataUrl || '').match(/^data:([^;]+);base64,(.*)$/);
    if (!m) return { ok: false, error: 'รูปแบบไฟล์ไม่ถูกต้อง' };
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    var cell = sheet.getRange(row, 7).getValue();
    var d = {}; try { d = JSON.parse(cell || '{}'); } catch (e) {}

    var ext = (m[1].indexOf('pdf') > -1) ? '.pdf' : ((m[1].indexOf('png') > -1) ? '.png' : '.jpg');
    var name = 'สลิป_' + (d.appNo || ('row' + row)).replace(/[\/\\]/g, '-') + '_' +
               Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd-HHmmss') + ext;
    var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
    var file = getSlipFolder_().createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

    var slip = { id: file.getId(), url: file.getUrl(), name: name,
      on: Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm'), by: username || '' };
    d.slips = d.slips || [];
    d.slips.push(slip);
    sheet.getRange(row, 7).setValue(JSON.stringify(d));
    return { ok: true, slip: slip, count: d.slips.length };
  } catch (e) { return { ok: false, error: e.message }; }
}
/* ดึงรูปสลิปเป็น dataUrl (สำหรับแสดงในระบบ) */
function getSlipImage(fileId) {
  try {
    var f = DriveApp.getFileById(fileId);
    var b = f.getBlob();
    return { ok: true, dataUrl: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()), name: f.getName() };
  } catch (e) { return { ok: false, error: e.message }; }
}
/* ลบสลิป (เจ้าหน้าที่ที่อัปโหลดเอง หรือ admin) */
function deleteSlip(row, fileId, username) {
  try {
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    var cell = sheet.getRange(row, 7).getValue();
    var d = {}; try { d = JSON.parse(cell || '{}'); } catch (e) {}
    var list = d.slips || [];
    var target = null;
    for (var i = 0; i < list.length; i++) if (list[i].id === fileId) target = list[i];
    if (!target) return { ok: false, error: 'ไม่พบสลิป' };
    if (!isAdminUser_(username) && String(target.by) !== String(username))
      return { ok: false, error: 'ลบได้เฉพาะสลิปที่ตนเองอัปโหลด (หรือผู้ดูแลระบบ)' };
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
    d.slips = list.filter(function (s) { return s.id !== fileId; });
    sheet.getRange(row, 7).setValue(JSON.stringify(d));
    return { ok: true, count: d.slips.length };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- แจ้งเตือนอัตโนมัติรายวัน (Trigger) ---------- */
var NOTIFY_HANDLER = 'notifyExpiring';
function getNotifyStatus() {
  try {
    var t = ScriptApp.getProjectTriggers();
    for (var i = 0; i < t.length; i++) if (t[i].getHandlerFunction() === NOTIFY_HANDLER) return { enabled: true };
    return { enabled: false };
  } catch (e) { return { enabled: false, error: e.message }; }
}
function enableDailyNotify(hour) {
  try {
    disableDailyNotify();
    var h = parseInt(hour, 10); if (isNaN(h) || h < 0 || h > 23) h = 8;
    ScriptApp.newTrigger(NOTIFY_HANDLER).timeBased().everyDays(1).atHour(h).create();
    return { ok: true, enabled: true, hour: h };
  } catch (e) { return { ok: false, error: e.message }; }
}
function disableDailyNotify() {
  try {
    var t = ScriptApp.getProjectTriggers();
    for (var i = 0; i < t.length; i++) if (t[i].getHandlerFunction() === NOTIFY_HANDLER) ScriptApp.deleteTrigger(t[i]);
    return { ok: true, enabled: false };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- ส่งออกข้อมูลเป็น CSV (ไม่ใช้ Drive) ---------- */
function getCertsCsv() {
  try {
    var sheet = getSheet_(SH_CERT, CERT_HEADERS);
    var data = sheet.getDataRange().getValues();
    var csv = data.map(function (r) {
      return r.map(function (x) {
        x = (x instanceof Date) ? Utilities.formatDate(x, 'GMT+7', 'yyyy-MM-dd') : (x == null ? '' : String(x));
        return '"' + x.replace(/"/g, '""') + '"';
      }).join(',');
    }).join('\r\n');
    var name = 'ใบรับรอง_' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd') + '.csv';
    return { ok: true, csv: csv, name: name };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- สถิติ + bootstrap ---------- */
function getStats() {
  var list = getCertificates();
  var types = getServiceTypes();
  var byType = {}; types.forEach(function (t) { byType[t.code] = 0; });
  var today = new Date(); var soon = new Date(); soon.setDate(today.getDate() + 30);
  var active = 0, expiringSoon = 0, expired = 0;
  list.forEach(function (c) {
    if (byType[c.serviceCode] !== undefined) byType[c.serviceCode]++;
    if (c.status === 'ใช้งาน') active++;
    if (c.expiryDate) { var d = new Date(c.expiryDate);
      if (!isNaN(d)) { if (d < today) expired++; else if (d <= soon) expiringSoon++; } }
  });
  return { total: list.length, active: active, expiringSoon: expiringSoon, expired: expired, byType: byType,
    recent: list.slice(0, 6), operators: getOperators().length };
}
/* ---------- คำขอใบอนุญาต ---------- */
var APP_HEADERS = ['เลขที่คำขอ','วันที่รับคำขอ','ประเภทแบบฟอร์ม','ชื่อผู้ขอ','ชื่อสถานประกอบกิจการ','สถานะ','ข้อมูล','บันทึกเมื่อ'];
function generateAppNo_(sheet) {
  var year = new Date().getFullYear() + 543;
  var data = sheet.getDataRange().getValues(); var n = 0;
  for (var i = 1; i < data.length; i++) { var no = String(data[i][0] || ''); if (no.indexOf('คข-') === 0 && no.indexOf('/' + year) > -1) n++; }
  return 'คข-' + ('0000' + (n + 1)).slice(-4) + '/' + year;
}
function getApplications() {
  var sheet = getSheet_(SH_APP, APP_HEADERS);
  var data = sheet.getDataRange().getValues(); var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i]; if (!r[0]) continue;
    var d = {}; try { d = JSON.parse(r[6] || '{}'); } catch (e) { d = {}; }
    d.row = i + 1; d.appNo = r[0]; d.receivedDate = fmt_(r[1]); d.formKey = r[2];
    d.applicantName = r[3]; d.bizName = r[4]; d.status = r[5] || 'รับเรื่อง';
    out.push(d);
  }
  return out.reverse();
}
function saveApplication(d) {
  try {
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    var no = generateAppNo_(sheet);
    var now = new Date();
    d.appNo = no; d.receivedDate = d.receivedDate || Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    sheet.appendRow([no, d.receivedDate, d.formKey || '', d.applicantName || '', d.bizName || '',
      d.status || 'รับเรื่อง', JSON.stringify(d), Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd HH:mm:ss')]);
    return { ok: true, appNo: no };
  } catch (e) { return { ok: false, error: e.message }; }
}
function updateApplication(d) {
  try {
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    sheet.getRange(d.row, 2, 1, 5).setValues([[ d.receivedDate || '', d.formKey || '', d.applicantName || '', d.bizName || '', d.status || 'รับเรื่อง' ]]);
    sheet.getRange(d.row, 7).setValue(JSON.stringify(d));
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
}
function deleteApplication(row) {
  try { getSheet_(SH_APP, APP_HEADERS).deleteRow(row); return { ok: true }; }
  catch (e) { return { ok: false, error: e.message }; }
}
function addApplications(list) {
  try {
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    var now = new Date(); var cnt = 0;
    (list || []).forEach(function (d) {
      var no = generateAppNo_(sheet);
      d.appNo = no;
      d.receivedDate = d.receivedDate || Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
      sheet.appendRow([no, d.receivedDate, d.formKey || '', d.applicantName || d.name || '', d.bizName || d.ename || '',
        d.status || 'รับเรื่อง', JSON.stringify(d), Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd HH:mm:ss')]);
      cnt++;
    });
    return { ok: true, count: cnt };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- ใบเสร็จรับเงิน (Receipt) ---------- */
function fiscalYearBE_() {
  var d = new Date(); var y = d.getFullYear() + 543;
  return (d.getMonth() + 1 >= 10) ? (y + 1) : y; // ปีงบเริ่ม 1 ต.ค.
}
function getReceiptState() {
  var p = PropertiesService.getScriptProperties();
  var fy = fiscalYearBE_();
  if (p.getProperty('RC_FY') && p.getProperty('RC_FY') !== String(fy)) {
    // ขึ้นปีงบใหม่ → รีเซ็ตเลขอัตโนมัติ
    p.setProperty('RC_SEQ', '0'); p.setProperty('RC_FY', String(fy));
  }
  if (!p.getProperty('RC_FY')) p.setProperty('RC_FY', String(fy));
  return { book: p.getProperty('RC_BOOK') || '1', lastNo: +(p.getProperty('RC_SEQ') || 0), fiscalYear: fy };
}
function setReceiptBook(book) {
  PropertiesService.getScriptProperties().setProperty('RC_BOOK', String(book || '1'));
  return { ok: true, book: String(book || '1') };
}
function payApplication(row, book) {
  try {
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    var p = PropertiesService.getScriptProperties();
    var fy = fiscalYearBE_();
    if (p.getProperty('RC_FY') !== String(fy)) { p.setProperty('RC_SEQ', '0'); p.setProperty('RC_FY', String(fy)); }
    book = String(book || p.getProperty('RC_BOOK') || '1');
    var seq = (+(p.getProperty('RC_SEQ') || 0)) + 1;
    var cell = sheet.getRange(row, 7).getValue();
    var d = {}; try { d = JSON.parse(cell || '{}'); } catch (e) {}
    d.rcBook = book; d.rcNo = String(seq); d.rcDate = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd'); d.paid = 'ชำระแล้ว'; d.rcFY = fy;
    sheet.getRange(row, 7).setValue(JSON.stringify(d));
    p.setProperty('RC_SEQ', String(seq)); p.setProperty('RC_BOOK', book);
    return { ok: true, rcBook: book, rcNo: String(seq), rcDate: d.rcDate, fiscalYear: fy };
  } catch (e) { return { ok: false, error: e.message }; }
}
function unpayApplication(row) {
  try {
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    var cell = sheet.getRange(row, 7).getValue();
    var d = {}; try { d = JSON.parse(cell || '{}'); } catch (e) {}
    d.rcBook = ''; d.rcNo = ''; d.rcDate = ''; d.paid = 'ยังไม่ชำระ';
    sheet.getRange(row, 7).setValue(JSON.stringify(d));
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ตรวจสิทธิ์ผู้เรียกจากชีตผู้ใช้ (กันเรียกข้ามสิทธิ์) */
function isAdminUser_(username) {
  try {
    if (!username) return false;
    var sheet = getSheet_(SH_USER, USER_HEADERS);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(username)) return String(data[i][3]) === 'admin';
    }
    return false;
  } catch (e) { return false; }
}
/* ลบใบเสร็จ (เฉพาะ admin) — ล้างเลขที่/เล่ม/วันที่ชำระ ออกจากคำขอ */
function deleteReceipt(row, username) {
  try {
    if (!isAdminUser_(username)) return { ok: false, error: 'ไม่มีสิทธิ์ลบใบเสร็จ (เฉพาะผู้ดูแลระบบ)' };
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    var cell = sheet.getRange(row, 7).getValue();
    var d = {}; try { d = JSON.parse(cell || '{}'); } catch (e) {}
    var old = { book: d.rcBook || '', no: d.rcNo || '', date: d.rcDate || '' };
    d.rcVoided = { book: old.book, no: old.no, date: old.date, by: username, on: Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss') };
    d.rcBook = ''; d.rcNo = ''; d.rcDate = ''; d.paid = 'ยังไม่ชำระ';
    sheet.getRange(row, 7).setValue(JSON.stringify(d));
    return { ok: true, voided: old };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- อนุมัติ / ออกใบอนุญาต / ต่ออายุ ---------- */
function fmtD_(d) { return Utilities.formatDate(d, 'GMT+7', 'yyyy-MM-dd'); }
function plusOneYear_(base) { var d = new Date(base.getTime()); d.setFullYear(d.getFullYear() + 1); d.setDate(d.getDate() - 1); return d; }
function approveApplication(row) {
  try {
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    var cell = sheet.getRange(row, 7).getValue();
    var d = {}; try { d = JSON.parse(cell || '{}'); } catch (e) {}
    var p = PropertiesService.getScriptProperties();
    var be = new Date().getFullYear() + 543;
    if (!d.licNo) {
      var key = 'LIC_SEQ_' + be;
      var seq = (+(p.getProperty(key) || 0)) + 1;
      p.setProperty(key, String(seq));
      d.licNo = String(seq);
    }
    if (!d.licYear) d.licYear = String(be);
    var today = new Date();
    if (!d.issueDate) d.issueDate = fmtD_(today);
    if (!d.expireDate) d.expireDate = fmtD_(plusOneYear_(today));
    d.status = 'อนุมัติ';
    sheet.getRange(row, 6).setValue('อนุมัติ');
    sheet.getRange(row, 7).setValue(JSON.stringify(d));
    return { ok: true, licNo: d.licNo, licYear: d.licYear, issueDate: d.issueDate, expireDate: d.expireDate };
  } catch (e) { return { ok: false, error: e.message }; }
}
function renewApplication(row) {
  try {
    var sheet = getSheet_(SH_APP, APP_HEADERS);
    var cell = sheet.getRange(row, 7).getValue();
    var d = {}; try { d = JSON.parse(cell || '{}'); } catch (e) {}
    d.renewals = d.renewals || [];
    d.renewals.push({ on: fmtD_(new Date()), issue: d.issueDate || '', expire: d.expireDate || '', rcBook: d.rcBook || '', rcNo: d.rcNo || '' });
    var today = new Date(); today.setHours(0,0,0,0);
    var base = today;
    if (d.expireDate) { var ex = new Date(d.expireDate); if (!isNaN(ex) && ex > today) base = ex; }
    d.issueDate = fmtD_(new Date());
    d.expireDate = fmtD_(plusOneYear_(base));
    d.rcBook = ''; d.rcNo = ''; d.rcDate = ''; d.paid = 'ยังไม่ชำระ';
    d.status = 'อนุมัติ';
    sheet.getRange(row, 6).setValue('อนุมัติ');
    sheet.getRange(row, 7).setValue(JSON.stringify(d));
    return { ok: true, issueDate: d.issueDate, expireDate: d.expireDate, count: d.renewals.length };
  } catch (e) { return { ok: false, error: e.message }; }
}

function bootstrap() {
  return {
    serviceTypes: getServiceTypes(),
    settings: getSettings(),
    webAppUrl: getWebAppUrl_(),
    stats: getStats(),
    certificates: getCertificates(),
    operators: getOperators()
  };
}

/* ---------- หน้าตรวจสอบใบอนุญาต (QR ?lic=เลขที่คำขอ) ---------- */
function renderVerifyLic_(no) {
  var s = getSettings();
  var found = null;
  try {
    var list = getApplications();
    for (var i = 0; i < list.length; i++) { if (String(list[i].appNo) === String(no)) { found = list[i]; break; } }
  } catch (e) {}
  var body;
  if (!found) {
    body = '<div class="box bad"><h2>ไม่พบข้อมูล</h2><p>ไม่พบคำขอ/ใบอนุญาตเลขที่ <b>' + esc_(no) + '</b> ในระบบ</p></div>';
  } else {
    var names = { haz: 'ใบอนุญาตประกอบกิจการที่เป็นอันตรายต่อสุขภาพ', food: 'ใบอนุญาตจัดตั้งสถานที่จำหน่ายอาหารและสะสมอาหาร', market: 'ใบอนุญาตประกอบกิจการตลาด', foodnotify: 'หนังสือรับรองการแจ้งจัดตั้งสถานที่จำหน่ายอาหารและสะสมอาหาร' };
    var expired = false;
    if (found.expireDate) { var ex = new Date(found.expireDate); var t = new Date(); t.setHours(0,0,0,0); if (!isNaN(ex) && ex < t) expired = true; }
    var ok = (found.status === 'อนุมัติ') && !expired;
    body = '<div class="box ' + (ok ? 'ok' : 'bad') + '"><h2>' + (ok ? '✓ เอกสารถูกต้อง' : (expired ? '✕ หมดอายุแล้ว' : '⚠ สถานะ: ' + esc_(found.status || '-'))) + '</h2>'
      + '<table>'
      + '<tr><td>เลขที่คำขอ</td><td><b>' + esc_(found.appNo) + '</b></td></tr>'
      + '<tr><td>ประเภทเอกสาร</td><td>' + esc_(names[found.formKey] || found.formKey || '-') + '</td></tr>'
      + (found.licNo ? ('<tr><td>เลขที่ใบอนุญาต</td><td>' + esc_(found.licNo) + (found.licYear ? (' ปี ' + esc_(found.licYear)) : '') + '</td></tr>') : '')
      + '<tr><td>ผู้ขอ</td><td>' + esc_(found.applicantName || '-') + '</td></tr>'
      + '<tr><td>สถานประกอบการ</td><td>' + esc_(found.bizName || '-') + '</td></tr>'
      + '<tr><td>สถานะคำขอ</td><td>' + esc_(found.status || '-') + '</td></tr>'
      + (found.issueDate ? ('<tr><td>วันที่ออก</td><td>' + esc_(found.issueDate) + '</td></tr>') : '')
      + (found.expireDate ? ('<tr><td>วันหมดอายุ</td><td>' + esc_(found.expireDate) + (expired ? ' <span style="color:#c0392b">(หมดอายุ)</span>' : '') + '</td></tr>') : '')
      + '</table></div>';
  }
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>ตรวจสอบใบอนุญาต</title>'
    + '<style>body{font-family:sans-serif;background:#f2efe8;margin:0;padding:24px;color:#223}'
    + '.wrap{max-width:520px;margin:0 auto}.hd{text-align:center;margin-bottom:14px}.hd h1{font-size:18px;margin:6px 0 2px;color:#0B3D4F}.hd p{margin:0;color:#777;font-size:13px}'
    + '.box{background:#fff;border-radius:14px;padding:20px 22px;box-shadow:0 4px 18px rgba(0,0,0,.07)}'
    + '.box.ok h2{color:#177245}.box.bad h2{color:#c0392b}.box h2{margin:0 0 12px;font-size:20px}'
    + 'table{width:100%;border-collapse:collapse;font-size:14.5px}td{padding:7px 4px;vertical-align:top}td:first-child{color:#789;white-space:nowrap;padding-right:12px}'
    + '</style></head><body><div class="wrap"><div class="hd"><h1>' + esc_(s.orgName) + '</h1><p>ระบบตรวจสอบใบอนุญาต/หนังสือรับรอง</p></div>' + body + '</div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('ตรวจสอบใบอนุญาต');
}
