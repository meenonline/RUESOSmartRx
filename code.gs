// ============================================================
//  SUB-STOCK MANAGEMENT SYSTEM — Google Apps Script
//  code.gs  |  โรงพยาบาลรือเสาะ
// ============================================================
const SHEETS = {
  TRANSACTION   : 'บันทึกรับจ่าย',
  SUMMARY       : 'สรุป',
  SUMMARY_NOLOT : 'สรุปแบบไม่แยก Lot',
  DISPENSE      : 'เบิกยา',
  FORMULARY     : 'บัญชีโรงพยาบาล',
  PRINT         : 'ใบเบิกยา',
  LOG           : 'Log',
  SETTINGS      : 'Settings',
  AUDIT_TRAIL   : 'Audit Trail',
  DASHBOARD     : 'Dashboard',
  EXPIRY_UNTRACK: 'ติดตามหมดอายุ',   // REQ #11 — ยกเลิกติดตาม
};
const HOSPITAL_NAME = 'โรงพยาบาลรือเสาะ อำเภอรือเสาะ จังหวัดนราธิวาส';
const FISCAL_YEAR   = '2569';
// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏥 Sub-Stock')
    .addItem('📥 นำเข้าข้อมูล Excel',  'showImportDialog')
    .addItem('🖨️  พิมพ์ใบเบิกยา',      'showPrintDialog')
    .addSeparator()
    .addItem('🔄 รีเฟรชทุกชีท',         'refreshAll')
    .addItem('🔄 รีเฟรช สรุป',           'refreshSummarySheets')
    .addItem('🔄 รีเฟรช เบิกยา',         'refreshDispenseSheet')
    .addSeparator()
    .addItem('🏗️  สร้างโครงสร้าง Sheets','setupAllSheets')
    .addToUi();
}
// ============================================================
//  SETUP SHEETS
// ============================================================
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('⚠️ ยืนยัน',
    'จะสร้าง/รีเซ็ต Header ของทุก Sheet\n(ข้อมูลเดิมยังคงอยู่)',
    ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) return;
  _setupTransactionSheet(ss);
  _setupSummarySheet(ss);
  _setupSummaryNoLotSheet(ss);
  _setupDispenseSheet(ss);
  _setupFormularySheet(ss);
  _setupLogSheet(ss);
  _setupSettingsSheet(ss);
  _setupAuditTrailSheet(ss);
  ui.alert('✅ เสร็จแล้ว', 'สร้างโครงสร้าง Sheets ทั้งหมดเรียบร้อย', ui.ButtonSet.OK);
}
function _setupTransactionSheet(ss) {
  const sh = _getOrCreate(ss, SHEETS.TRANSACTION);
  const headers = ['Disp No','วันที่','แผนก','รหัสยา','ชื่อยาเต็ม',
    'จำนวน','Pack','ราคา/หน่วย','มูลค่ารวม','Lot No',
    'Barcode','วันหมดอายุ','ประเภทธุรกรรม','ผู้บันทึก','เวลาบันทึก'];
  _writeHeader(sh, headers, '#1a56db');
  sh.setColumnWidth(5,280); sh.setColumnWidth(2,100); sh.setColumnWidth(12,110);
}
function _setupSummarySheet(ss) {
  const sh = _getOrCreate(ss, SHEETS.SUMMARY);
  const headers = ['รหัสยา','ชื่อยาเต็ม','Lot No','วันหมดอายุ','Pack Size',
    'คงเหลือปัจจุบัน','รับเข้าทั้งหมด','ตัดจ่ายทั้งหมด',
    'มูลค่ารวม','ราคาเฉลี่ย','สถานะสต็อก','สถานะหมดอายุ',
    'วันจนหมดอายุ','อัปเดตล่าสุด','Minimum Stock'];
  _writeHeader(sh, headers, '#0ea472');
  sh.setColumnWidth(2,280);
}
function _setupSummaryNoLotSheet(ss) {
  const sh = _getOrCreate(ss, SHEETS.SUMMARY_NOLOT);
  const headers = ['ชื่อยาเต็ม','Pack Size','คงเหลือปัจจุบัน (Pack)',
    'รับเข้าทั้งหมด (Pack)','ตัดจ่ายทั้งหมด (Pack)',
    'มูลค่ารวม','ราคาเฉลี่ย/Pack','สถานะสต็อก',
    'อัปเดตล่าสุด','Minimum Stock','ตู้ยา','Reorder Point'];
  _writeHeader(sh, headers, '#d97706');
  sh.setColumnWidth(1,280);
}
function _setupDispenseSheet(ss) {
  const sh = _getOrCreate(ss, SHEETS.DISPENSE);
  const headers = ['รหัสยา','ชื่อยาเต็ม','Lot No','วันหมดอายุ','Pack Size',
    'คงเหลือปัจจุบัน','Minimum Stock','สถานะสต็อก','ตู้ยา','จำนวนที่เบิก'];
  _writeHeader(sh, headers, '#7c3aed');
  sh.setColumnWidth(2,280); sh.setColumnWidth(10,110);
}
function _setupFormularySheet(ss) {
  const sh = _getOrCreate(ss, SHEETS.FORMULARY);
  if (sh.getLastRow() > 1) return;
  sh.getRange(1,1,1,5).merge().setValue('🏥 บัญชีโรงพยาบาล - การกำหนด Minimum Stock')
    .setBackground('#1a56db').setFontColor('#fff').setFontWeight('bold').setFontSize(13);
  const h2 = ['📋 ชื่อยา','⚠️ Minimum Stock','ตู้ยา','🔔 Reorder Point','📍 ตำแหน่งบนชั้น'];
  sh.getRange(2,1,1,5).setValues([h2]);
  _styleRange(sh.getRange(2,1,1,5), '#374151');
  sh.setFrozenRows(2); sh.setColumnWidth(1,300); sh.setColumnWidth(2,130);
  sh.setColumnWidth(3,100); sh.setColumnWidth(4,130); sh.setColumnWidth(5,160);
}
function _setupSettingsSheet(ss) {
  const sh = _getOrCreate(ss, SHEETS.SETTINGS);
  if (sh.getLastRow() > 0) return;
  sh.getRange(1,1,1,2).merge().setValue('⚙️ Settings — LINE Notification & System')
    .setBackground('#1a56db').setFontColor('#fff').setFontWeight('bold').setFontSize(13);
  const rows = [
    ['KEY','VALUE'],
    ['LINE_CHANNEL_ACCESS_TOKEN',''],
    ['LINE_TARGET_TYPE','groupId'],  // groupId | userId
    ['LINE_TARGET_ID',''],
    ['NOTIFY_HOUR','8'],             // ส่งเวลากี่โมง (0-23)
    ['NOTIFY_LOW_STOCK','TRUE'],     // แจ้งเตือนสต็อกต่ำ
    ['NOTIFY_EXPIRY','TRUE'],        // แจ้งเตือนใกล้หมดอายุ
    ['EXPIRY_WARN_DAYS','90'],       // แจ้งเตือนล่วงหน้า X วัน
    ['NOTIFY_STALE_DATA','TRUE'],    // แจ้งเตือนข้อมูลไม่อัพเดท
    ['STALE_DATA_DAYS','7'],         // กี่วันถือว่า stale
    ['STALE_NOTIFY_DAYS','1111100'], // แจ้งเตือนวันไหน (จันทร์-อาทิตย์ 1=แจ้ง 0=ไม่แจ้ง)
    ['NOTIFY_EMAIL','TRUE'],          // ส่งสรุปทาง Email
    ['EMAIL_RECIPIENTS',''],          // email1@example.com,email2@example.com
  ];
  sh.getRange(2,1,rows.length,2).setValues(rows);
  sh.getRange(2,1,1,2).setBackground('#374151').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(2); sh.setColumnWidth(1,250); sh.setColumnWidth(2,350);
}
function _setupAuditTrailSheet(ss) {
  const sh = _getOrCreate(ss, SHEETS.AUDIT_TRAIL);
  if (sh.getLastRow() > 0) return;
  _writeHeader(sh, ['วันที่เวลา','ผู้แก้ไข','หน้า/Sheet','Field ที่แก้','ค่าเดิม','ค่าใหม่','หมายเหตุ'], '#7c3aed');
  sh.setColumnWidth(1,160); sh.setColumnWidth(3,160);
  sh.setColumnWidth(4,140); sh.setColumnWidth(5,220); sh.setColumnWidth(6,220);
}
function _setupLogSheet(ss) {
  const sh = _getOrCreate(ss, SHEETS.LOG);
  _writeHeader(sh, ['วันที่เวลา','กิจกรรม','จำนวนรายการ','ผู้บันทึก','รายละเอียด'], '#374151');
}
// ============================================================
//  IMPORT DIALOG
// ============================================================
function showImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('import')
    .setWidth(640).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '📥 นำเข้าข้อมูล Excel');
}
// ============================================================
//  IMPORT DATA — เรียกจาก HTML (import.html)
//  payload = { rows: [...], txnType: 'รับเข้า'|'เบิกยา', recorder: string }
// ============================================================
function importData(payload) {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const sh       = _getOrCreate(ss, SHEETS.TRANSACTION);
    const txnType  = payload.txnType;
    const recorder = payload.recorder || 'ระบบ';
    const rows     = payload.rows || [];
    const sign     = txnType === 'รับเข้า' ? 1 : -1;
    const now      = new Date();
    if (!rows.length) return { success: false, message: 'ไม่มีข้อมูล' };
    const writeRows = rows.map(r => {
      const fullName = [r.NAME, r.TYPE, r.CONTENT].filter(Boolean).join(' ');
      const qty   = sign * Math.abs(Number(r.amount) || 0);
      const price = Number(r.price) || 0;
      return [
        String(r.dispno || ''),
        _parseExcelDate(r.date),
        r.department || '',
        r.code1 || '',
        fullName,
        qty,
        Number(r.pack) || 1,
        price,
        qty * price,
        String(r.LotNo || ''),
        r.Barcode || '',
        _parseExcelDate(r.ExpDate),
        txnType,
        recorder,
        now,
      ];
    });
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, writeRows.length, 15).setValues(writeRows);
    sh.getRange(startRow, 2, writeRows.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(startRow,12, writeRows.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(startRow, 8, writeRows.length, 2).setNumberFormat('#,##0.00');
    sh.getRange(startRow, 6, writeRows.length, 1).setNumberFormat('#,##0');
    // สีแถว ตาม txnType
    const bg = txnType === 'รับเข้า' ? '#f0fdf4' : '#fef9f0';
    sh.getRange(startRow,1,writeRows.length,15).setBackground(bg);
    _logActivity(`นำเข้า (${txnType})`, writeRows.length, recorder);
    refreshSummarySheets();
    refreshDispenseSheet();
    return { success: true, count: writeRows.length, txnType };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}
// ============================================================
//  REFRESH ALL
// ============================================================
function refreshAll() {
  refreshSummarySheets();
  refreshDispenseSheet();
  SpreadsheetApp.getUi().alert('✅ รีเฟรชเสร็จแล้ว','อัปเดตทุกชีทเรียบร้อย', SpreadsheetApp.getUi().ButtonSet.OK);
}
// ============================================================
//  REFRESH SUMMARY SHEETS
// ============================================================
function refreshSummarySheets() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName(SHEETS.TRANSACTION);
  if (!txnSheet || txnSheet.getLastRow() < 2) return;
  const data    = txnSheet.getDataRange().getValues();
  const hdr     = data[0];
  const idx     = h => hdr.indexOf(h);
  const lotMap  = {};  // key = code|name|lot|exp
  const nameMap = {};  // key = name
  for (let i = 1; i < data.length; i++) {
    const r    = data[i];
    const code = r[idx('รหัสยา')];
    const name = String(r[idx('ชื่อยาเต็ม')] || '').trim();
    const lot  = String(r[idx('Lot No')] || '');
    const exp  = r[idx('วันหมดอายุ')];
    const pack = Number(r[idx('Pack')]) || 1;
    const qty  = Number(r[idx('จำนวน')]) || 0;
    const price= Number(r[idx('ราคา/หน่วย')]) || 0;
    const time = r[idx('เวลาบันทึก')];
    if (!name) continue;
    // LOT MAP
    const lk = `${code}||${name}||${lot}||${exp}`;
    if (!lotMap[lk]) lotMap[lk] = {code,name,lot,exp,pack,bal:0,inTotal:0,outTotal:0,value:0,prices:[],latestTime:null};
    const lm = lotMap[lk];
    lm.bal   += qty;
    if (qty>0) lm.inTotal += qty; else lm.outTotal += Math.abs(qty);
    lm.value += qty * price;
    if (price>0) lm.prices.push(price);
    if (time && (!lm.latestTime || time > lm.latestTime)) lm.latestTime = time;
    // NAME MAP
    if (!nameMap[name]) nameMap[name] = {name,pack,bal:0,inTotal:0,outTotal:0,value:0,latestTime:null};
    const nm = nameMap[name];
    nm.bal   += qty;
    if (qty>0) nm.inTotal += qty; else nm.outTotal += Math.abs(qty);
    nm.value += qty * price;
    if (time && (!nm.latestTime || time > nm.latestTime)) nm.latestTime = time;
  }
  const minMap = _getFormularyMap(ss);
  _writeSummary(ss, lotMap, minMap);
  _writeSummaryNoLot(ss, nameMap, minMap);
}
function _writeSummary(ss, lotMap, minMap) {
  const sh  = _getOrCreate(ss, SHEETS.SUMMARY);
  const now = new Date();
  // Always ensure header is correct
  const hdr = ['รหัสยา','ชื่อยาเต็ม','Lot No','วันหมดอายุ','Pack Size',
    'คงเหลือปัจจุบัน','รับเข้าทั้งหมด','ตัดจ่ายทั้งหมด',
    'มูลค่ารวม','ราคาเฉลี่ย','สถานะสต็อก','สถานะหมดอายุ',
    'วันจนหมดอายุ','อัปเดตล่าสุด','Minimum Stock'];
  sh.getRange(1,1,1,hdr.length).setValues([hdr]);
  _styleRange(sh.getRange(1,1,1,hdr.length), '#0ea472');
  sh.setFrozenRows(1);
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,15).clearContent().setBackground('#ffffff');
  const rows = Object.values(lotMap).map(lm => {
    const minStock = minMap[lm.name] ? minMap[lm.name].min : '';
    const avgPrice = lm.prices.length ? lm.prices.reduce((a,b)=>a+b,0)/lm.prices.length : 0;
    const expDate  = lm.exp instanceof Date ? lm.exp : (lm.exp ? new Date(lm.exp) : null);
    const daysLeft = expDate ? Math.round((expDate-now)/86400000) : '';
    const stockSt  = _stockStatus(lm.bal, minStock);
    const expSt    = _expiryStatus(daysLeft);
    return [lm.code, lm.name, lm.lot, expDate||'', lm.pack,
      lm.bal, lm.inTotal, lm.outTotal,
      _r2(lm.value), _r2(avgPrice),
      stockSt, expSt, daysLeft, lm.latestTime||now, minStock];
  }).sort((a,b)=>String(a[1]).localeCompare(String(b[1]),'th'));
  if (!rows.length) return;
  sh.getRange(2,1,rows.length,15).setValues(rows);
  sh.getRange(2,4,rows.length,1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2,14,rows.length,1).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(2,6,rows.length,4).setNumberFormat('#,##0.00');
  _applyStockColors(sh, rows, 10, 2);
}
function _writeSummaryNoLot(ss, nameMap, minMap) {
  const sh = _getOrCreate(ss, SHEETS.SUMMARY_NOLOT);
  const hdr = ['ชื่อยาเต็ม','Pack Size','คงเหลือปัจจุบัน (Pack)',
    'รับเข้าทั้งหมด (Pack)','ตัดจ่ายทั้งหมด (Pack)',
    'มูลค่ารวม','ราคาเฉลี่ย/Pack','สถานะสต็อก',
    'อัปเดตล่าสุด','Minimum Stock','ตู้ยา','Reorder Point'];
  sh.getRange(1,1,1,hdr.length).setValues([hdr]);
  _styleRange(sh.getRange(1,1,1,hdr.length), '#d97706');
  sh.setFrozenRows(1);
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,12).clearContent().setBackground('#ffffff');
  // โหลด reorder point ที่ตั้งค่า manual ไว้ก่อน (จาก Formulary col 4)
  const reorderMap = _getReorderMap(ss);
  const rows = Object.values(nameMap).map(nm => {
    const fm        = _lookupFormulary(minMap, nm.name);
    const minStock  = fm.min !== undefined ? fm.min : '';
    const cabinet   = fm.cabinet || '';
    const stockSt   = _stockStatus(nm.bal, minStock);
    // Reorder Point: ใช้ค่า manual ถ้ามี ไม่งั้น auto = min * 1.5
    const autoRp    = minStock !== '' ? Math.ceil(Number(minStock) * 1.5) : '';
    const reorderPt = reorderMap[nm.name] !== undefined ? reorderMap[nm.name] : autoRp;
    return [nm.name, nm.pack, nm.bal, nm.inTotal, nm.outTotal,
      _r2(nm.value), _r2(nm.inTotal ? nm.value/nm.inTotal : 0),
      stockSt, nm.latestTime||new Date(), minStock, cabinet, reorderPt];
  }).sort((a,b)=>String(a[0]).localeCompare(String(b[0]),'th'));
  if (!rows.length) return;
  sh.getRange(2,1,rows.length,12).setValues(rows);
  sh.getRange(2,9,rows.length,1).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(2,3,rows.length,5).setNumberFormat('#,##0.00');
  _applyStockColors(sh, rows, 7, 2);
}
// ============================================================
//  REFRESH DISPENSE SHEET
// ============================================================
function refreshDispenseSheet() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sumSheet = ss.getSheetByName(SHEETS.SUMMARY);
  const dispSheet= _getOrCreate(ss, SHEETS.DISPENSE);
  if (!sumSheet || sumSheet.getLastRow() < 2) return;
  const data   = sumSheet.getDataRange().getValues();
  const hdr    = data[0];
  const idx    = h => hdr.indexOf(h);
  const minMap = _getFormularyMap(ss);
  // Always re-write header
  const dispHdr = ['รหัสยา','ชื่อยาเต็ม','Lot No','วันหมดอายุ','Pack Size',
    'คงเหลือปัจจุบัน','Minimum Stock','สถานะสต็อก','ตู้ยา','จำนวนที่เบิก'];
  dispSheet.getRange(1,1,1,dispHdr.length).setValues([dispHdr]);
  _styleRange(dispSheet.getRange(1,1,1,dispHdr.length), '#0f766e');
  dispSheet.setFrozenRows(1);
  if (dispSheet.getLastRow() > 1) dispSheet.getRange(2,1,dispSheet.getLastRow()-1,10).clearContent().setBackground('#ffffff');
  const rows = [];
  for (let i=1; i<data.length; i++) {
    const r   = data[i];
    const bal = Number(r[idx('คงเหลือปัจจุบัน')]);
    if (bal <= 0) continue;
    const name = r[idx('ชื่อยาเต็ม')];
    const fm   = _lookupFormulary(minMap, name);
    rows.push([r[idx('รหัสยา')], name, r[idx('Lot No')], r[idx('วันหมดอายุ')],
      r[idx('Pack Size')], bal, fm.min||'', r[idx('สถานะสต็อก')], fm.cabinet||'', '']);
  }
  rows.sort((a,b)=>String(a[8]).localeCompare(String(b[8]),'th')||String(a[1]).localeCompare(String(b[1]),'th'));
  if (!rows.length) return;
  dispSheet.getRange(2,1,rows.length,10).setValues(rows);
  dispSheet.getRange(2,4,rows.length,1).setNumberFormat('dd/mm/yyyy');
  dispSheet.getRange(2,6,rows.length,2).setNumberFormat('#,##0');
  for (let i=0; i<rows.length; i++) {
    const st = rows[i][7];
    let bg = '#f0fdf4';
    if (st === 'ต่ำกว่า 50% ของ min stock') bg = '#fef3c7';
    else if (st === 'ต่ำกว่า minimum') bg = '#fff7ed';
    else if (st === 'หมดสต็อก') bg = '#fee2e2';
    dispSheet.getRange(i+2,1,1,10).setBackground(bg);
  }
  dispSheet.getRange(2,10,rows.length,1).setBackground('#dbeafe');
}
// ============================================================
//  PRINT DIALOG
// ============================================================
function showPrintDialog() {
  const html = HtmlService.createHtmlOutputFromFile('print_dialog')
    .setWidth(520).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '🖨️ พิมพ์ใบเบิกยา');
}
function getCabinetList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.FORMULARY);
  if (!sh || sh.getLastRow() < 3) return [];
  const data = sh.getRange(3,3,sh.getLastRow()-2,1).getValues();
  return [...new Set(data.map(r=>String(r[0]||'')).filter(Boolean))].sort();
}
// ============================================================
//  GENERATE PRINT SHEET + DOWNLOAD
// ============================================================
function processSelectionAndGenerate(selection) {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const dispSheet= ss.getSheetByName(SHEETS.DISPENSE);
    if (!dispSheet || dispSheet.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('ไม่มีข้อมูล','กรุณารีเฟรชชีทเบิกยาก่อน',SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const data = dispSheet.getDataRange().getValues();
    const hdr  = data[0];
    const idx  = h => hdr.indexOf(h);
    const medicines = [];
    for (let i=1; i<data.length; i++) {
      const r       = data[i];
      const bal     = Number(r[idx('คงเหลือปัจจุบัน')]);
      const cabinet = String(r[idx('ตู้ยา')]||'');
      if (bal <= 0) continue;
      if (selection.type === 'cabinet' && selection.cabinets.length > 0) {
        if (!selection.cabinets.includes(cabinet)) continue;
      }
      medicines.push({
        code: r[idx('รหัสยา')], name: r[idx('ชื่อยาเต็ม')],
        lot: r[idx('Lot No')], expDate: r[idx('วันหมดอายุ')],
        pack: r[idx('Pack Size')], balance: bal, cabinet,
      });
    }
    if (!medicines.length) {
      SpreadsheetApp.getUi().alert('ไม่พบรายการ','ไม่มียาตามเงื่อนไขที่เลือก',SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    // ===== เรียงชื่อยา A→Z เสมอ =====
    medicines.sort((a,b) => String(a.name||'').localeCompare(String(b.name||''),'th'));
    const cabLabel = selection.type === 'ipd'
      ? 'IPD ทั้งหมด'
      : selection.cabinets.join(', ');
    const summary = `📋 ใบเบิกยา — ตู้: ${cabLabel}`;
    let psh = ss.getSheetByName(SHEETS.PRINT);
    if (!psh) psh = ss.insertSheet(SHEETS.PRINT); else psh.clear();
    _logActivity(`พิมพ์ใบเบิก (${summary})`, medicines.length, 'ระบบ');
    // ===== สร้าง HTML ใบเบิกยา — table เดียว เรียง A→Z =====
    const now      = new Date();
    const dateStr  = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMdd_HHmm');
    const thaiDate = _thaiDateTime(now);
    const thaiMon  = _thaiMonth(now.getMonth());
    const year     = now.getFullYear() + 543;
    const dataRows = medicines.map(function(m, i) {
      const bg = i % 2 === 0 ? '#f9fafb' : '#ffffff';
      return '<tr style="background:' + bg + '">'
        + '<td style="text-align:center;border:1px solid #d1d5db;padding:4px 6px">' + (i+1) + '</td>'
        + '<td style="text-align:left;border:1px solid #d1d5db;padding:4px 8px;font-weight:500">' + (m.name||'') + '</td>'
        + '<td style="text-align:center;border:1px solid #d1d5db;padding:4px 6px;font-size:12px">' + (m.lot||'') + '</td>'
        + '<td style="text-align:center;border:1px solid #d1d5db;padding:4px 6px">' + (m.pack||'') + '</td>'
        + '<td style="text-align:center;border:1px solid #d1d5db;padding:4px 6px;font-weight:600;color:#1a56db">' + (m.balance||0) + '</td>'
        + '<td style="text-align:center;border:1px solid #d1d5db;padding:4px 6px">........</td>'
        + '<td style="text-align:center;border:1px solid #d1d5db;padding:4px 6px">........</td>'
        + '<td style="text-align:center;border:1px solid #d1d5db;padding:4px 6px">........</td>'
        + '<td style="text-align:center;border:1px solid #d1d5db;padding:4px 6px"></td>'
        + '</tr>';
    }).join('');
    const tableHtml = '<table style="width:100%;border-collapse:collapse;font-size:13px;font-family:\'TH Sarabun New\',Sarabun,sans-serif">'
      + '<thead><tr style="background:linear-gradient(135deg,#1e3a5f,#1a56db)">'
      + '<th style="width:36px;color:#fff;border:1px solid #1a56db;padding:6px 4px;text-align:center;font-size:12px">ลำดับ</th>'
      + '<th style="text-align:left;color:#fff;border:1px solid #1a56db;padding:6px 8px;font-size:13px">รายการยาที่ขอเบิก</th>'
      + '<th style="width:90px;color:#fff;border:1px solid #1a56db;padding:6px 4px;text-align:center;font-size:12px">Lot No.</th>'
      + '<th style="width:55px;color:#fff;border:1px solid #1a56db;padding:6px 4px;text-align:center;font-size:12px">Pack</th>'
      + '<th style="width:55px;color:#fff;border:1px solid #1a56db;padding:6px 4px;text-align:center;font-size:12px">คงเหลือ</th>'
      + '<th style="width:65px;color:#fff;border:1px solid #1a56db;padding:6px 4px;text-align:center;font-size:12px">วันที่ 1</th>'
      + '<th style="width:65px;color:#fff;border:1px solid #1a56db;padding:6px 4px;text-align:center;font-size:12px">วันที่ 2</th>'
      + '<th style="width:65px;color:#fff;border:1px solid #1a56db;padding:6px 4px;text-align:center;font-size:12px">วันที่ 3</th>'
      + '<th style="width:80px;color:#fff;border:1px solid #1a56db;padding:6px 4px;text-align:center;font-size:12px">หมายเหตุ</th>'
      + '</tr></thead>'
      + '<tbody>' + dataRows
      + '<tr style="background:#e8f0fe"><td colspan="2" style="text-align:center;font-weight:700;border:1px solid #d1d5db;padding:5px 6px;color:#1a56db">รวมทั้งหมด ' + medicines.length + ' รายการ</td>'
      + '<td colspan="7" style="border:1px solid #d1d5db"></td></tr>'
      + '</tbody></table>';
    const htmlContent = '<!DOCTYPE html><html><head><meta charset="UTF-8"/>'
      + '<title>ใบเบิกยา ' + dateStr + '</title>'
      + '<style>'
      + '*{box-sizing:border-box;margin:0;padding:0}'
      + 'body{font-family:"TH Sarabun New",Sarabun,Arial,sans-serif;font-size:15px;background:#fff;padding:8mm 10mm}'
      + '.no-print{margin-bottom:12px;text-align:center;padding:10px}'
      + 'button{padding:10px 24px;border:none;border-radius:8px;font-size:15px;cursor:pointer;font-family:inherit;margin-right:8px;font-weight:600}'
      + '.btn-print{background:linear-gradient(135deg,#1a56db,#1e3a5f);color:#fff;box-shadow:0 2px 8px rgba(26,86,219,.3)}'
      + '.btn-close{background:#6b7280;color:#fff}'
      + '.page-header{text-align:center;margin-bottom:14px;border-bottom:3px solid #1a56db;padding-bottom:10px}'
      + '.page-header h1{font-size:24px;color:#1a56db;margin-bottom:4px;letter-spacing:-0.02em}'
      + '.page-header p{font-size:13px;color:#555}'
      + '.summary-bar{background:linear-gradient(135deg,#e8f0fe,#dbeafe);color:#1a56db;font-weight:700;font-size:14px;text-align:center;padding:8px 12px;border-radius:6px;margin-bottom:14px;border:1px solid #bfdbfe}'
      + '.footer-sign{display:flex;justify-content:space-around;margin-top:28px;font-size:14px}'
      + '.footer-sign>div{min-width:200px}'
      + '@media print{@page{size:A4 landscape;margin:8mm}body{padding:0;font-size:14px}.no-print{display:none}tr{page-break-inside:avoid}}'
      + '</style></head><body>'
      + '<div class="no-print">'
      + '<button class="btn-print" onclick="window.print()">🖨️ พิมพ์ / บันทึก PDF</button>'
      + '<button class="btn-close" onclick="window.close()">✕ ปิด</button>'
      + '</div>'
      + '<div class="page-header">'
      + '<h1>ใบเบิกยา</h1>'
      + '<p>' + HOSPITAL_NAME + ' | ปีงบประมาณ ' + FISCAL_YEAR + '</p>'
      + '<p>วันที่พิมพ์: ' + thaiDate + ' | ประจำเดือน ' + thaiMon + ' ' + year + '</p>'
      + '</div>'
      + '<div class="summary-bar">' + summary + ' — รวม ' + medicines.length + ' รายการ (เรียงตามชื่อยา)</div>'
      + tableHtml
      + '<div class="footer-sign">'
      + '<div style="text-align:center"><p>ลงชื่อ ผู้เบิก ...............................<br/>(...............................)<br/>วันที่ ....../....../......</p></div>'
      + '<div style="text-align:center"><p>ลงชื่อ ผู้จ่าย ...............................<br/>(...............................)<br/>วันที่ ....../....../......</p></div>'
      + '</div></body></html>';
    return {
      success:     true,
      htmlContent: htmlContent,
      medicines:   medicines,
      count:       medicines.length,
      summary:     summary,
    };
  } catch(e) {
    _logActivity('Error พิมพ์ใบเบิก', 0, e.toString());
    return { success: false, error: e.toString() };
  }
}
// ============================================================
//  BUILD PRINT SHEET — โครงสร้างใบเบิก (ตามรูปที่ 5)
// ============================================================
function _buildPrintSheet(sh, medicines, summary, cabinets) {
  const now      = new Date();
  const thaiDate = _thaiDateTime(now);
  const thaiMon  = _thaiMonth(now.getMonth());
  const year     = now.getFullYear()+543;
  sh.setColumnWidths(1,1,40);  sh.setColumnWidths(2,1,280);
  sh.setColumnWidths(3,1,110); sh.setColumnWidths(4,1,80);
  for (let c=5;c<=9;c++) sh.setColumnWidths(c,1,70);
  sh.setColumnWidths(10,1,100);
  let r = 1;
  // Header block
  _mergeSet(sh,r,1,1,10,'ใบเบิกยา',{bg:'#1a56db',fc:'#fff',fw:'bold',fs:16,ha:'center'}); r++;
  _mergeSet(sh,r,1,1,10,`${HOSPITAL_NAME} ปีงบ ${FISCAL_YEAR}`,{ha:'center',fs:12}); r++;
  _mergeSet(sh,r,1,1,10,`วันที่พิมพ์: ${thaiDate}`,{ha:'center',fc:'#555'}); r++;
  _mergeSet(sh,r,1,1,10,`ประจำเดือน ${thaiMon} ${year}`,{ha:'center',fc:'#555'}); r++;
  _mergeSet(sh,r,1,1,10,summary,{ha:'center',fw:'bold',bg:'#e8f0fe',fc:'#1a56db',fs:13}); r++;
  r++; // blank
  // กลุ่มตาม ตู้ยา
  const groups = {};
  if (cabinets.length > 1) {
    medicines.forEach(m => {
      const c = m.cabinet||'ไม่ระบุ';
      if (!groups[c]) groups[c]=[];
      groups[c].push(m);
    });
  } else {
    const c = cabinets[0]||'ทั้งหมด';
    groups[c] = medicines;
  }
  const colHdr = ['ลำดับ','รายการยาที่ขอเบิก','Lot No.','Pack Size','วันที่ 1','วันที่ 2','วันที่ 3','วันที่ 4','วันที่ 5','หมายเหตุ'];
  Object.entries(groups).forEach(([cab, meds]) => {
    _mergeSet(sh,r,1,1,10,`ใบเบิกยา ตู้: ${cab}`,{bg:'#374151',fc:'#fff',fw:'bold',fs:12,ha:'center'}); r++;
    sh.getRange(r,1,1,10).setValues([colHdr]);
    _styleRange(sh.getRange(r,1,1,10),'#4b5563'); r++;
    meds.forEach((m,i) => {
      sh.getRange(r,1,1,10).setValues([[i+1, m.name, m.lot, m.pack,'........','........','........','........','........','']]);
      sh.getRange(r,1,1,10).setBackground(i%2===0?'#f9fafb':'#ffffff');
      sh.getRange(r,1).setHorizontalAlignment('center');
      r++;
    });
    _mergeSet(sh,r,1,1,2,`รวมทั้งหมด ${meds.length} รายการ`,{fw:'bold',bg:'#e8f0fe',ha:'center'});
    sh.getRange(r,3,1,8).setBackground('#e8f0fe'); r++;
    r++; // blank between groups
  });
  // Footer
  r++;
  sh.getRange(r,1,1,5).merge().setValue('ลงชื่อ ผู้เบิก ...............................').setHorizontalAlignment('center');
  sh.getRange(r,6,1,5).merge().setValue('ลงชื่อ ผู้จ่าย ...............................').setHorizontalAlignment('center'); r++;
  sh.getRange(r,1,1,5).merge().setValue('วันที่ ...............................').setHorizontalAlignment('center');
  sh.getRange(r,6,1,5).merge().setValue('วันที่ ...............................').setHorizontalAlignment('center');
  sh.getRange(7,1,r-6,10).setBorder(true,true,true,true,true,true,'#d1d5db',SpreadsheetApp.BorderStyle.SOLID);
}
// ============================================================
//  SERVER-SIDE DATA FOR HTML (called via google.script.run)
// ============================================================
function getSheetPreview(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 1) return {headers:[],rows:[]};
  const data = sh.getDataRange().getValues();
  return {
    headers: data[0],
    rows: data.slice(1,51).map(row =>
      row.map(cell => cell instanceof Date ? Utilities.formatDate(cell,'Asia/Bangkok','dd/MM/yyyy') : cell)
    )
  };
}
function getStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txn = ss.getSheetByName(SHEETS.TRANSACTION);
  const sum = ss.getSheetByName(SHEETS.SUMMARY_NOLOT);
  const txnRows = txn ? Math.max(0, txn.getLastRow()-1) : 0;
  let lowStock=0, outStock=0;
  if (sum && sum.getLastRow() > 1) {
    const data = sum.getRange(2,8,sum.getLastRow()-1,1).getValues();
    data.forEach(r=>{
      if(r[0]==='หมดสต็อก') outStock++;
      else if(r[0]==='ต่ำกว่า 50% ของ min stock' || r[0]==='ต่ำกว่า minimum') lowStock++;
    });
  }
  return { txnRows, lowStock, outStock };
}
// ============================================================
//  HELPERS
// ============================================================
function _getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
// ── ยกเลิก/เปิดใช้งานรายการยาในบัญชีโรงพยาบาล ──
function _setInactive(body) {
  const drugName = String(body.drugName || '').trim();
  const inactive = body.inactive === true || body.inactive === 'true';
  if (!drugName) return { success: false, error: 'ไม่มีชื่อยา' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.FORMULARY);
  if (!sh) return { success: false, error: 'ไม่พบชีท ' + SHEETS.FORMULARY };
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const data    = sh.getRange(1, 1, lastRow, lastCol).getValues();
  // หา header row
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(data.length, 5); i++) {
    const cell = String(data[i][0] || '').trim();
    if (cell === '' || cell.includes('บัญชี') || cell.toLowerCase().includes('ชื่อยา')) {
      headerRowIdx = i;
    }
  }
  const headers = data[headerRowIdx].map(h => String(h || '').trim());
  // ตรวจว่ามีคอลัมน์ INACTIVE ไหม ถ้าไม่มีให้เพิ่ม
  let inactiveCol = headers.indexOf('INACTIVE');
  if (inactiveCol < 0) {
    inactiveCol = lastCol; // จะเพิ่มคอลัมน์ใหม่
    sh.getRange(headerRowIdx + 1, inactiveCol + 1).setValue('INACTIVE');
  }
  // หาแถวที่ตรงกับชื่อยา
  const norm = s => String(s || '').trim().replace(/^[📋⚠️]\s*/u, '').toLowerCase();
  const targetNorm = norm(drugName);
  let found = false;
  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const cellName = norm(data[i][0]);
    if (cellName === targetNorm || cellName === norm(data[i][0])) {
      if (String(data[i][0] || '').trim().replace(/^[📋⚠️]\s*/u, '').toLowerCase() === targetNorm) {
        sh.getRange(i + 1, inactiveCol + 1).setValue(inactive ? 'TRUE' : '');
        found = true;
        break;
      }
    }
  }
  if (!found) return { success: false, error: 'ไม่พบยา: ' + drugName };
  _logActivity(inactive ? 'ตัดรายการยาออก' : 'เปิดใช้งานรายการยา', drugName, inactive ? 'INACTIVE' : 'ACTIVE');
  return { success: true, message: (inactive ? 'ตัดรายการ' : 'เปิดใช้งาน') + ' ' + drugName + ' เรียบร้อย' };
}
function _getFormularyMap(ss) {
  const sh = ss.getSheetByName(SHEETS.FORMULARY);
  if (!sh || sh.getLastRow() < 3) return {};
  const allData = sh.getDataRange().getValues();
  
  let dataStartRow = 2;
  for (let i = 0; i < Math.min(allData.length, 5); i++) {
    const cell = String(allData[i][0] || '').trim();
    if (cell === '' || cell.includes('บัญชี') || cell.toLowerCase() === 'ชื่อยา' ||
        cell.toLowerCase().includes('name') || cell.includes('รายการ') ||
        cell.includes('📋')) {
      dataStartRow = i + 1;
    }
  }
  const map = {};
  // normalize: lowercase + collapse spaces + ลบ spaces รอบหน่วย (50 mg → 50mg)
  const norm = s => String(s || '').trim().toLowerCase()
    .replace(/\s+/g, ' ')                    // collapse multiple spaces
    .replace(/(\d)\s+(mg|mcg|ml|g|iu|u)\b/gi, '$1$2')  // "50 mg" → "50mg"
    .replace(/\s*(tab|cap|inj|sol|syp|susp|cream|oint|eye|ear|drop|lot|gel)\s*/gi, ' $1 ')
    .replace(/\s+/g, ' ').trim();
  for (let i = dataStartRow; i < allData.length; i++) {
    const r      = allData[i];
    const rawName = String(r[0] || '').trim().replace(/^[📋⚠️]\s*/u,'');
    const minVal = r[1];
    const cab    = String(r[2] || '').trim();
    if (!rawName) continue;
    // Skip INACTIVE entries — ตรวจสอบคอลัมน์ INACTIVE (ใด ๆ ใน row)
    const rowStr = r.map(c => String(c||'').toUpperCase());
    if (rowStr.includes('TRUE') && r.some((c,ci) => ci > 2 && String(c||'').toUpperCase() === 'TRUE')) continue;
    const minNum = typeof minVal === 'number' ? minVal :
                   (minVal !== '' ? Number(String(minVal).replace(/,/g,'')) : 0);
    const entry = { min: isNaN(minNum) ? 0 : minNum, cabinet: cab };
    map[rawName]        = entry;   // exact key
    map[norm(rawName)]  = entry;   // normalized key
  }
  return map;
}
// Helper: ค้นหาใน formulary map แบบ fuzzy
function _lookupFormulary(fmap, drugName) {
  if (!drugName) return { min: 0, cabinet: '' };
  // 1. exact
  if (fmap[drugName]) return fmap[drugName];
  // 2. normalized (spaces + units)
  const norm = s => String(s || '').trim().toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/(\d)\s+(mg|mcg|ml|g|iu|u)\b/gi, '$1$2')
    .replace(/\s*(tab|cap|inj|sol|syp|susp|cream|oint|eye|ear|drop|lot|gel)\s*/gi, ' $1 ')
    .replace(/\s+/g, ' ').trim();
  if (fmap[norm(drugName)]) return fmap[norm(drugName)];
  // 3. ultra-aggressive: ลบ spaces ทั้งหมดแล้วเทียบ
  const strip = s => String(s || '').toLowerCase().replace(/\s+/g, '');
  const stripped = strip(drugName);
  for (const key of Object.keys(fmap)) {
    if (strip(key) === stripped) return fmap[key];
  }
  return { min: 0, cabinet: '' };
}
function _parseExcelDate(val) {
  if (!val) return '';
  if (val instanceof Date) return val;
  if (typeof val === 'number') return new Date(Math.round((val-25569)*86400000));
  const d = new Date(val);
  return isNaN(d) ? String(val) : d;
}
function _stockStatus(bal, min) {
  if (min === '' || min === undefined) return 'ยังมีอยู่';
  if (bal <= 0)        return 'หมดสต็อก';
  if (bal < min*0.5)   return 'ต่ำกว่า 50% ของ min stock';
  if (bal < min)       return 'ต่ำกว่า minimum';
  return 'ยังมีอยู่';
}
function _expiryStatus(days) {
  if (days === '') return 'ไม่ระบุ';
  if (days < 0)    return 'หมดอายุ';
  if (days <= 30)  return 'ใกล้หมดอายุ';
  if (days <= 90)  return 'ควรเฝ้าระวัง';
  return 'ปกติ';
}
function _applyStockColors(sh, rows, statusColIdx, startRow) {
  for (let i=0; i<rows.length; i++) {
    const s = rows[i][statusColIdx];
    let bg = '#ffffff';
    if (s==='หมดสต็อก')                   bg='#fee2e2';
    else if (s==='ต่ำกว่า 50% ของ min stock') bg='#fef3c7';
    else if (s==='ต่ำกว่า minimum')         bg='#fff7ed';
    sh.getRange(startRow+i,1,1,rows[0].length).setBackground(bg);
  }
}
function _writeHeader(sh, headers, bg) {
  if (sh.getLastRow() > 0) sh.getRange(1,1,1,headers.length).clearContent();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  _styleRange(sh.getRange(1,1,1,headers.length), bg);
  sh.setFrozenRows(1);
}
function _styleRange(range, bg) {
  range.setBackground(bg).setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(12).setHorizontalAlignment('center');
}
function _mergeSet(sh,r,c,nr,nc,val,opts={}) {
  const rng = sh.getRange(r,c,nr,nc).merge().setValue(val);
  if (opts.bg)  rng.setBackground(opts.bg);
  if (opts.fc)  rng.setFontColor(opts.fc);
  if (opts.fw)  rng.setFontWeight(opts.fw);
  if (opts.fs)  rng.setFontSize(opts.fs);
  if (opts.ha)  rng.setHorizontalAlignment(opts.ha);
}
function _thaiDateTime(d) {
  const m = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return `${d.getDate()} ${m[d.getMonth()]} ${d.getFullYear()+543} เวลา ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
}
function _thaiMonth(m) {
  return ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'][m];
}
function _r2(n) { return Math.round(n*100)/100; }
function _logActivity(action, count, user) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = _getOrCreate(ss, SHEETS.LOG);
    if (sh.getLastRow()===0) sh.getRange(1,1,1,5).setValues([['วันที่เวลา','กิจกรรม','จำนวน','ผู้บันทึก','รายละเอียด']]);
    sh.appendRow([new Date(), action, count, user||'ระบบ', '']);
  } catch(e){}
}
// ============================================================
//  WEB APP ENDPOINTS — สำหรับ index.html
// ============================================================
// doGet router สำหรับ Web App
function doGet(e) {
  // ถ้าไม่มี action ให้ serve index.html
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Sub-Stock | รพ.รือเสาะ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const action = e.parameter.action;
  const sheet  = e.parameter.sheet || '';
  try {
    let result;
    switch(action) {
      case 'getData':        result = _getSheetData(sheet);  break;
      case 'getStats':       result = _getStatsData();        break;
      case 'getCabinetList': result = { data: getCabinetList() }; break;
      case 'getSettings':    result = _getSettings();         break;
      case 'getAuditTrail':  result = _getAuditTrailData();   break;
      case 'getMonthlyData': result = _getMonthlyData(e.parameter.month, e.parameter.year); break;
      case 'getNextDispNo':  result = _getNextDispNo();    break;
      case 'getDispUsers':   result = _getDispUsers();     break;
      case 'getProcureList': result = _getProcureList();   break;
      case 'getUsageData':   result = _getUsageData(e.parameter.from, e.parameter.to); break;
      case 'getRecentDrugList': result = _getRecentDrugList(); break;
      case 'getUntrackedExpiry': result = _getUntrackedExpiry(); break;
      case 'getMonthlyTrend': result = _getMonthlyTrend(); break;
      default: result = { success: false, error: 'Unknown action: ' + action };
    }
    return _jsonOut(result);
  } catch(err) {
    return _jsonOut({ success: false, error: err.toString() });
  }
}
// doPost router สำหรับ Web App
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    switch(action) {
      case 'importData':
        result = importData(body);
        break;
      case 'refreshAll':
        refreshSummarySheets();
        refreshDispenseSheet();
        result = { success: true };
        break;
      case 'processSelectionAndGenerate':
        result = processSelectionAndGenerate({ type: body.type, cabinets: body.cabinets||[], isPDF: body.isPDF });
        if (!result) result = { success: false, error: 'ไม่มีข้อมูลยา' };
        break;
      case 'adjustStock':
        result = _adjustStock(body);
        break;
      case 'saveAuditResult':
        result = _saveAuditResult(body);
        break;
      case 'updateCabinet':
        result = _updateCabinet(body);
        break;
      case 'updateTransaction':
        result = _updateTransaction(body);
        break;
      case 'updateSummaryField':
        result = _updateSummaryField(body);
        break;
      case 'updateReorderPoint':
        result = _updateReorderPoint(body);
        break;
      case 'saveSettings':
        result = _saveSettings(body);
        break;
      case 'sendLineTest':
        result = _sendLineTest();
        break;
      case 'addManualTransaction':
        result = _addManualTransaction(body);
        break;
      case 'batchDispense':
        result = _batchDispense(body);
        break;
      case 'updateMinStock':
        result = _updateMinStock(body);
        break;
      case 'setInactive':
        result = _setInactive(body);
        break;
      case 'writeDashboard':
        result = writeDashboardSheet();
        result.success = true;
        break;
      case 'sendEmailTest': result = _sendEmailTest(body); break;
      case 'untrackExpiry': result = _untrackExpiry(body); break;
      case 'updateFormulary': result = _updateFormulary(body); break;
      case 'getMonthlyTrend': result = _getMonthlyTrend(); break;
      case 'retrackExpiry': result = _retrackExpiry(body); break;
      default:
        result = { success: false, error: 'Unknown action' };
    }
    return _jsonOut(result);
  } catch(e) {
    return _jsonOut({ success: false, error: e.toString() });
  }
}
// ============================================================
//  ADJUST STOCK — ปรับสต็อกจาก Web App
// ============================================================
function _adjustStock(body) {
  // body: { drugName, lotNo, currentQty, newQty, recorder }
  const { drugName, lotNo, currentQty, newQty, recorder } = body;
  if (!drugName) return { success: false, error: 'ไม่มีชื่อยา' };
  const diff = Number(newQty) - Number(currentQty);
  if (diff === 0) return { success: false, error: 'จำนวนเท่าเดิม ไม่มีการเปลี่ยนแปลง' };
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const txn = _getOrCreate(ss, SHEETS.TRANSACTION);
  // หา header row ของ บันทึกรับจ่าย
  // headers: Disp No|วันที่|แผนก|รหัสยา|ชื่อยาเต็ม|จำนวน|Pack|ราคา/หน่วย|มูลค่ารวม|Lot No|Barcode|วันหมดอายุ|ประเภทธุรกรรม|ผู้บันทึก|เวลาบันทึก
  const txnHdr = txn.getRange(1, 1, 1, txn.getLastColumn()).getValues()[0];
  const tIdx   = h => txnHdr.findIndex(c => String(c).trim() === h);
  // ดึงข้อมูลยาจาก sheet สรุป (แยก Lot) เพื่อเอา code, ราคา, วันหมดอายุ
  const sumSh = ss.getSheetByName(SHEETS.SUMMARY);
  let code = '', expDate = '', unitPrice = 0, barcode = '';
  if (sumSh && sumSh.getLastRow() > 1) {
    const sumData = sumSh.getDataRange().getValues();
    const sumHdr  = sumData[0];
    const sIdx    = h => sumHdr.findIndex(c => String(c).trim() === h);
    for (let i = 1; i < sumData.length; i++) {
      const r       = sumData[i];
      const rName   = String(r[sIdx('ชื่อยาเต็ม')] || '').trim();
      const rLot    = String(r[sIdx('Lot No')]       || '').trim();
      const lotMatch = !lotNo || rLot === String(lotNo).trim();
      if (rName === drugName.trim() && lotMatch) {
        code      = String(r[sIdx('รหัสยา')]      || '');
        expDate   = r[sIdx('วันหมดอายุ')]          || '';
        unitPrice = Number(r[sIdx('ราคา/หน่วย')]   || r[sIdx('ราคาเฉลี่ย')] || 0);
        barcode   = String(r[sIdx('Barcode')]       || '');
        break;
      }
    }
  }
  const now      = new Date();
  const dateOnly = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // วันที่ไม่มีเวลา
  const timeStr  = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');     // เวลาบันทึก
  // จำนวน: diff มีเครื่องหมาย (+เพิ่ม / -ลด)
  // มูลค่ารวม = |diff| × ราคา/หน่วย
  const totalValue = Math.abs(diff) * unitPrice;
  // สร้าง row ตาม header จริง
  const newRow = new Array(txnHdr.length).fill('');
  const set    = (col, val) => { const i = tIdx(col); if (i >= 0) newRow[i] = val; };
  set('Disp No',         '');           // ไม่มี Disp No สำหรับการปรับ stock
  set('วันที่',           dateOnly);
  set('แผนก',            '');
  set('รหัสยา',           code);
  set('ชื่อยาเต็ม',       drugName);
  set('จำนวน',            diff);         // ← มีเครื่องหมาย + หรือ -
  set('Pack',            1);
  set('ราคา/หน่วย',       unitPrice);
  set('มูลค่ารวม',        totalValue);
  set('Lot No',          lotNo || '');
  set('Barcode',         barcode);
  set('วันหมดอายุ',       expDate);
  set('ประเภทธุรกรรม',    'ปรับ stock'); // ← ตามที่กำหนด
  set('ผู้บันทึก',         recorder || 'ระบบ');
  set('เวลาบันทึก',        now);  // ใช้ Date object แทน string
  txn.appendRow(newRow);
  // Double flush pattern — ให้ข้อมูลลง sheet ก่อน refresh
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  SpreadsheetApp.flush();
  // Recalculate summary sheets
  refreshSummarySheets();
  refreshDispenseSheet();
  _logActivity(
    `ปรับ stock: ${drugName} (${diff > 0 ? '+' : ''}${diff})`,
    Math.abs(diff),
    recorder
  );
   return {
    success: true,
    message: `ปรับ stock ${drugName} สำเร็จ (${diff > 0 ? '+' : ''}${diff})`,
    newQty:  Number(newQty),
    diff:    diff,
    refreshed: true,   // ← เพิ่มบรรทัดนี้
  };
}
// ============================================================
//  ADD MANUAL TRANSACTION — เพิ่มรายการ Manual ใน บันทึกรับจ่าย
// ============================================================
function _addManualTransaction(body) {
  const { dispNo, date, dept, code, drugName, qty, pack, price, total, lotNo, expiry, recorder } = body;
  if (!drugName) return { success: false, error: 'ไม่มีชื่อยา' };
  if (!date)     return { success: false, error: 'ไม่มีวันที่' };
  if (!recorder) return { success: false, error: 'ไม่มีชื่อผู้บันทึก' };
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const txn = _getOrCreate(ss, SHEETS.TRANSACTION);
  // parse date string to date object
  let dateObj;
  try { dateObj = new Date(date); } catch(e) { dateObj = new Date(); }
  const dateOnly = Utilities.formatDate(dateObj, 'Asia/Bangkok', 'dd/MM/yyyy');
  // parse expiry
  let expDate = '';
  if (expiry) {
    try {
      const ed = new Date(expiry);
      expDate = Utilities.formatDate(ed, 'Asia/Bangkok', 'dd/MM/yyyy');
    } catch(e) { expDate = expiry; }
  }
  const now     = new Date();
  const timeStr = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  const qtyNum  = Number(qty) || 0;
  const priceNum = Number(price) || 0;
  const totalNum = Number(total) || Math.abs(qtyNum) * priceNum;
  // ประเภทธุรกรรม: ถ้าจำนวน > 0 = รับเข้า, < 0 = เบิกยา
  const txnType = qtyNum > 0 ? 'รับเข้า' : (qtyNum < 0 ? 'เบิกยา' : 'manual');
  const txnHdr = txn.getLastRow() > 0
    ? txn.getRange(1, 1, 1, txn.getLastColumn()).getValues()[0]
    : [];
  const tIdx = h => txnHdr.findIndex(c => String(c).trim() === h);
  const newRow = new Array(Math.max(txnHdr.length, 15)).fill('');
  const set = (col, val) => { const i = tIdx(col); if (i >= 0) newRow[i] = val; };
  set('Disp No',         dispNo   || '');
  set('วันที่',           dateOnly);
  set('แผนก',            dept     || '');
  set('รหัสยา',           code     || '');
  set('ชื่อยาเต็ม',       drugName);
  set('จำนวน',            qtyNum);
  set('Pack',            Number(pack) || 1);
  set('ราคา/หน่วย',       priceNum);
  set('มูลค่ารวม',        totalNum);
  set('Lot No',          lotNo    || '');
  set('วันหมดอายุ',       expDate  || '');
  set('ประเภทธุรกรรม',    txnType);
  set('ผู้บันทึก',         recorder);
  set('เวลาบันทึก',        timeStr);
  txn.appendRow(newRow);
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  SpreadsheetApp.flush();
  refreshSummarySheets();
  refreshDispenseSheet();
  _logActivity(`เพิ่ม Manual: ${drugName} (${qtyNum > 0 ? '+' : ''}${qtyNum})`, Math.abs(qtyNum), recorder);
  return { success: true, message: `บันทึก ${drugName} สำเร็จ (${qtyNum > 0 ? '+' : ''}${qtyNum})` };
}
function _saveAuditResult(body) {
  // body: { round, recorder, items: [{ drugName, lotNo, sysQty, counted, diff, cabinet, type }] }
  const { round, recorder, items } = body;
  if (!items || !items.length) return { success: false, error: 'ไม่มีข้อมูลการตรวจนับ' };
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const AUDIT  = 'สุ่มนับจำนวน';
  const sh     = _getOrCreate(ss, AUDIT);
  // สร้าง header ถ้ายังไม่มี
  const auditHeaders = [
    'รอบที่','วันที่','เวลาบันทึก','ผู้ตรวจนับ',
    'ชื่อยา','Lot No','ตู้ยา','ประเภท',
    'คงเหลือระบบ','นับได้','ผลต่าง','สถานะ'
  ];
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, auditHeaders.length).setValues([auditHeaders]);
    _styleRange(sh.getRange(1, 1, 1, auditHeaders.length), '#7c3aed');
  }
  const now     = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy');
  const timeStr = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');
  const rows = items.map(it => [
    round,
    dateStr,
    timeStr,
    recorder,
    it.drugName  || '',
    it.lotNo     || '',
    it.cabinet   || '',
    it.type === 'critical' ? 'วิกฤต' : 'ปกติ',
    Number(it.sysQty  || 0),
    Number(it.counted || 0),
    Number(it.diff    || 0),
    Number(it.diff || 0) === 0 ? 'ตรง' : 'ต่าง',
  ]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, auditHeaders.length).setValues(rows);
  _logActivity(`บันทึกผลสุ่มนับรอบที่ ${round}`, items.length, recorder);
  return {
    success: true,
    count:   items.length,
    round:   round,
    message: `บันทึกผลตรวจนับรอบที่ ${round} จำนวน ${items.length} รายการสำเร็จ`
  };
}
// ดึงข้อมูลจาก Sheet แล้วแปลงเป็น Array of Objects
function _getSheetData(sheetName) {
  if (!sheetName) return { success: false, error: 'sheet parameter required' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return { success: true, data: [] };
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return { success: true, data: [] };
  const raw = sh.getRange(1, 1, lastRow, lastCol).getValues();
  // บัญชีโรงพยาบาล มี merged title row 1, header จริงอยู่ row 2
  let headerRowIdx = 0;
  if (sheetName === 'บัญชีโรงพยาบาล' && raw.length > 2) {
    headerRowIdx = 1;
  }
  const headers = raw[headerRowIdx].map(h => String(h || '').trim());
  const dataRows = raw.slice(headerRowIdx + 1);
  const data = dataRows
    .filter(row => row.some(cell => {
      if (cell === '' || cell === null || cell === undefined) return false;
      if (typeof cell === 'number' && cell === 0) return true; // 0 ถือว่ามีค่า
      return true;
    }))
    .map(row => {
      const obj = {};
      headers.forEach((key, i) => {
        if (!key) return;
        const v = row[i];
        if (v instanceof Date) {
          obj[key] = Utilities.formatDate(v, 'Asia/Bangkok', 'dd/MM/yyyy');
        } else if (typeof v === 'number' && v > 40000 && (key.includes('วัน') || key.includes('date') || key.includes('Date'))) {
          // Excel serial date
          try {
            obj[key] = Utilities.formatDate(new Date(Math.round((v - 25569) * 86400000)), 'Asia/Bangkok', 'dd/MM/yyyy');
          } catch(e) { obj[key] = v; }
        } else {
          obj[key] = v;
        }
      });
      return obj;
    });
  return { success: true, data, count: data.length };
}
// ═══════════════════════════════════════════════════════════════════
//  _getStatsData — REQ #11: กรอง untracked expiry + แสดง stock=0
// ═══════════════════════════════════════════════════════════════════
function _getStatsData() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const txn    = ss.getSheetByName(SHEETS.TRANSACTION);
  const sumLot = ss.getSheetByName(SHEETS.SUMMARY);        // แยก Lot
  const sumNo  = ss.getSheetByName(SHEETS.SUMMARY_NOLOT);  // ไม่แยก Lot
  const txnRows = txn ? Math.max(0, txn.getLastRow() - 1) : 0;
  // ── สต็อก จาก NoLot sheet ──
  let activeItems = 0, lowStock = 0, outStock = 0, halfMinStock = 0;
  const stockOkList  = [];
  const lowStockList = [];
  const outStockList = [];
  const halfMinList  = [];
  // โหลด inactive drugs จาก formulary
  const _inactiveSet = new Set();
  const fmSh = ss.getSheetByName(SHEETS.FORMULARY);
  if (fmSh && fmSh.getLastRow() > 1) {
    const fmData   = fmSh.getDataRange().getValues();
    const fmHdr    = fmData.length > 1 ? fmData[1] : fmData[0];
    const inactCol = fmHdr.map(h => String(h || '').toUpperCase()).indexOf('INACTIVE');
    if (inactCol >= 0) {
      for (let fi = 2; fi < fmData.length; fi++) {
        if (String(fmData[fi][inactCol] || '').toUpperCase() === 'TRUE') {
          const nm = String(fmData[fi][0] || '').trim().replace(/^[📋⚠️]\s*/u, '').toLowerCase();
          if (nm) _inactiveSet.add(nm);
        }
      }
    }
  }
  // ── REQ #11: โหลด untracked expiry set ──
  const _untrackedSet = new Set();
  try {
    const untrackedRes = _getUntrackedExpiry();
    (untrackedRes.data || []).forEach(r => {
      _untrackedSet.add(r.drugName + '||' + r.lotNo);
    });
  } catch (e) {
    // ถ้าโหลดไม่ได้ก็ข้ามไป
  }
  if (sumNo && sumNo.getLastRow() > 1) {
    const data    = sumNo.getDataRange().getValues();
    const hdr     = data[0];
    const nameIdx = hdr.indexOf('ชื่อยาเต็ม');
    const balIdx  = hdr.findIndex(h => String(h).includes('คงเหลือปัจจุบัน'));
    const minIdx  = hdr.indexOf('Minimum Stock');
    const stIdx   = hdr.indexOf('สถานะสต็อก');
    const cabIdx  = hdr.indexOf('ตู้ยา');
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][nameIdx] || '');
      const bal  = Number(data[i][balIdx]  || 0);
      const min  = Number(data[i][minIdx]  || 0);
      const st   = String(data[i][stIdx]   || '');
      const cab  = String(data[i][cabIdx]  || '');
      if (!name) continue;
      if (_inactiveSet.has(name.toLowerCase())) continue;
      const entry = { name, bal, min, cab };
      if (st === 'หมดสต็อก') {
        outStock++;
        outStockList.push(entry);
      } else if (st === 'ต่ำกว่า 50% ของ min stock') {
        halfMinStock++;
        halfMinList.push(entry);
      } else if (st === 'ต่ำกว่า minimum') {
        lowStock++;
        lowStockList.push(entry);
      } else if (bal > 0) {
        activeItems++;
        stockOkList.push(entry);
      }
    }
  }
  // ── อายุยา จาก Summary (แยก Lot) ──
  let exp0 = 0, exp3 = 0, exp6 = 0, expUnknown = 0;
  const expiredList    = [];
  const exp3List       = [];
  const exp6List       = [];
  const expUnknownList = [];
  if (sumLot && sumLot.getLastRow() > 1) {
    const data    = sumLot.getDataRange().getValues();
    const hdr     = data[0];
    const nameIdx = hdr.indexOf('ชื่อยาเต็ม');
    const lotIdx  = hdr.indexOf('Lot No');
    const expIdx  = hdr.indexOf('วันหมดอายุ');
    const daysIdx = hdr.indexOf('วันจนหมดอายุ');
    const stExpIdx = hdr.indexOf('สถานะหมดอายุ');
    const balIdx  = hdr.findIndex(h => String(h).includes('คงเหลือปัจจุบัน'));
    for (let i = 1; i < data.length; i++) {
      const name  = String(data[i][nameIdx] || '');
      const lot   = String(data[i][lotIdx]  || '');
      const bal   = Number(data[i][balIdx]  || 0);
      const stExp = String(data[i][stExpIdx] || '');
      let   days  = data[i][daysIdx];
      days = (days === '' || days === undefined) ? null : Number(days);
      let expStr = '';
      const ev = data[i][expIdx];
      if (ev instanceof Date) expStr = Utilities.formatDate(ev, 'Asia/Bangkok', 'dd/MM/yyyy');
      else expStr = String(ev || '');
      if (!name) continue;
      // ── REQ #11: กรอง untracked ──
      if (_untrackedSet.has(name + '||' + lot)) continue;
      const entry = { name, lot, bal, exp: expStr };
      // ── REQ #11: ไม่กรองด้วย bal > 0 — แสดงแม้ stock = 0 ──
      if (days === null || expStr === '' || stExp === 'ไม่ระบุ') {
        expUnknown++;
        expUnknownList.push(entry);
      } else if (days < 0 || stExp === 'หมดอายุ') {
        exp0++;
        expiredList.push({ ...entry, days: Math.abs(days) });
      } else if (days <= 90) {
        exp3++;
        exp3List.push({ ...entry, days });
      } else if (days <= 180) {
        exp6++;
        exp6List.push({ ...entry, days });
      }
    }
  }
  return {
    success: true,
    txnRows,
    // สต็อก counts
    activeItems,
    lowStock,
    outStock,
    halfMinStock,
    totalNoLot: stockOkList.length + lowStock + outStock + halfMinStock,
    // อายุ counts
    exp0, exp3, exp6, expUnknown,
    // รายชื่อ (จำกัด 200 รายการต่อ list)
    stockOkList:     stockOkList.slice(0, 200),
    lowStockList:    lowStockList.slice(0, 200),
    outStockList:    outStockList.slice(0, 200),
    halfMinList:     halfMinList.slice(0, 200),
    expiredList:     expiredList.slice(0, 200),
    exp3List:        exp3List.slice(0, 200),
    exp6List:        exp6List.slice(0, 200),
    expUnknownList:  expUnknownList.slice(0, 200),
  };
}
function _jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
// ============================================================
//  UPDATE CABINET — แก้ไขตู้ยาจาก NoLot page → sync Formulary
// ============================================================
function _updateCabinet(body) {
  // body: { drugName, cabinet }
  const { drugName, cabinet } = body;
  if (!drugName) return { success: false, error: 'ไม่มีชื่อยา' };
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const formSh   = _getOrCreate(ss, SHEETS.FORMULARY);
  // หา header row จริงของ Formulary (row 1 คือ title merged, row 2 คือ header)
  // ตรวจว่า row 2 คือ header
  let hdrRow = 1;
  if (formSh.getLastRow() >= 2) {
    const r1 = String(formSh.getRange(1,1).getValue() || '');
    if (r1.includes('บัญชี') || r1.includes('🏥')) hdrRow = 2;
  }
  const dataStart = hdrRow + 1;
  // โหลด formulary ทั้งหมด
  const lastRow  = formSh.getLastRow();
  let   foundRow = -1;
  if (lastRow >= dataStart) {
    const vals = formSh.getRange(dataStart, 1, lastRow - dataStart + 1, 3).getValues();
    for (let i = 0; i < vals.length; i++) {
      const name = String(vals[i][0] || '').replace(/^[📋⚠️]\s*/u,'').trim();
      if (name === drugName.trim()) {
        foundRow = dataStart + i;
        break;
      }
    }
  }
  if (foundRow > 0) {
    // อัปเดต column ตู้ยา (col 3)
    formSh.getRange(foundRow, 3).setValue(cabinet);
  } else {
    // เพิ่มแถวใหม่ใน Formulary
    const minStock = _getMinStockForDrug(ss, drugName);
    formSh.appendRow([drugName, minStock || '', cabinet]);
    // Style แถวใหม่
    const newRow = formSh.getLastRow();
    formSh.getRange(newRow, 1, 1, 3).setBackground('#f9fafb').setBorder(true,true,true,true,true,true);
  }
  // Sync กลับไปที่ sheet สรุปแบบไม่แยก Lot ด้วย (อัปเดต column ตู้ยา ที่ col 11)
  const noLotSh = ss.getSheetByName(SHEETS.SUMMARY_NOLOT);
  if (noLotSh && noLotSh.getLastRow() > 1) {
    const nlData = noLotSh.getRange(2, 1, noLotSh.getLastRow()-1, 11).getValues();
    for (let i = 0; i < nlData.length; i++) {
      if (String(nlData[i][0]||'').trim() === drugName.trim()) {
        noLotSh.getRange(i + 2, 11).setValue(cabinet);
        break;
      }
    }
  }
  _logActivity(`อัปเดตตู้ยา: ${drugName} → ${cabinet}`, 1, 'ผู้ใช้');
  return { success: true, message: `บันทึกตู้ยา "${cabinet}" สำหรับ ${drugName} สำเร็จ` };
}
// helper: ดึง minStock ปัจจุบันจาก sheet สรุปแบบไม่แยก Lot
function _getMinStockForDrug(ss, drugName) {
  const sh = ss.getSheetByName(SHEETS.SUMMARY_NOLOT);
  if (!sh || sh.getLastRow() < 2) return '';
  const vals = sh.getRange(2, 1, sh.getLastRow()-1, 10).getValues();
  for (const r of vals) {
    if (String(r[0]||'').trim() === drugName.trim()) return r[9] || '';
  }
  return '';
}
// ============================================================
//  UPDATE TRANSACTION — แก้ไข Lot No / วันหมดอายุ / ชื่อยาเต็ม
//  ใน sheet บันทึกรับจ่าย แบบ inline จาก Web App
//  ใช้ txnKey = "DispNo|วันที่|ชื่อยาเต็ม|LotNo" ค้นหาแถว
// ============================================================
function _updateTransaction(body) {
  // body: { txnKey, field, value }
  // field: 'ชื่อยาเต็ม' | 'Lot No' | 'วันหมดอายุ'
  const { txnKey, field, value } = body;
  if (!txnKey || !field) return { success: false, error: 'ข้อมูลไม่ครบ' };
  const allowed = ['ชื่อยาเต็ม', 'Lot No', 'วันหมดอายุ'];
  if (!allowed.includes(field)) return { success: false, error: 'ไม่อนุญาตให้แก้ไข field นี้' };
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const txn = _getOrCreate(ss, SHEETS.TRANSACTION);
  if (txn.getLastRow() < 2) return { success: false, error: 'ไม่มีข้อมูล' };
  const allData = txn.getDataRange().getValues();
  const hdr     = allData[0].map(h => String(h).trim());
  const tIdx    = h => hdr.indexOf(h);
  // parse txnKey
  const parts    = String(txnKey).split('|');
  const keyDisp  = parts[0] || '';
  const keyDate  = parts[1] || '';
  const keyName  = parts[2] || '';
  const keyLot   = parts[3] || '';
  // ค้นหาแถวที่ตรงกัน (ค้นจากท้ายเผื่อซ้ำ → เอาแถวล่าสุด)
  let foundRow = -1;
  for (let i = allData.length - 1; i >= 1; i--) {
    const r = allData[i];
    const rDisp = String(r[tIdx('Disp No')]    || '').trim();
    const rDate = String(r[tIdx('วันที่')]      || '').trim();
    const rName = String(r[tIdx('ชื่อยาเต็ม')]  || '').trim();
    const rLot  = String(r[tIdx('Lot No')]      || '').trim();
    if (rDisp === keyDisp && rDate === keyDate && rName === keyName && rLot === keyLot) {
      foundRow = i + 1; // 1-based sheet row
      break;
    }
  }
  if (foundRow < 0) return { success: false, error: 'ไม่พบแถวที่ต้องการแก้ไข (อาจมีการรีเฟรชข้อมูล)' };
  const col = tIdx(field) + 1; // 1-based
  if (col <= 0) return { success: false, error: `ไม่พบ column: ${field}` };
  txn.getRange(foundRow, col).setValue(value);
  SpreadsheetApp.flush();
  _logActivity(`แก้ไข ${field} แถว ${foundRow}: ${value}`, 1, 'ผู้ใช้');
  return { success: true, message: `บันทึก ${field} สำเร็จ (แถว ${foundRow})` };
}
function _getRecentDrugList() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const txn = ss.getSheetByName(SHEETS.TRANSACTION);
  if (!txn || txn.getLastRow() < 2) return { success: true, data: [] };
  const allData = txn.getDataRange().getValues();
  const hdr     = allData[0].map(h => String(h).trim());
  const nameIdx = hdr.indexOf('ชื่อยาเต็ม');
  const packIdx = hdr.indexOf('Pack');
  const priceIdx= hdr.indexOf('ราคา/หน่วย');
  const lotIdx  = hdr.indexOf('Lot No');
  const expIdx  = hdr.indexOf('วันหมดอายุ');
  const codeIdx = hdr.indexOf('รหัสยา');
  const dateIdx = hdr.indexOf('วันที่');
  // สร้าง map: key = "ชื่อยา||packSize" → ข้อมูลล่าสุด
  // วนจากล่างขึ้น เพื่อเอาแถวล่าสุดก่อน
  const drugMap = {};
  for (let i = allData.length - 1; i >= 1; i--) {
    const r       = allData[i];
    const name    = String(r[nameIdx] || '').trim();
    const pack    = Number(r[packIdx] || 1) || 1;
    const price   = Number(r[priceIdx] || 0);
    const lot     = String(r[lotIdx]  || '');
    const code    = String(r[codeIdx] || '');
    const rawDate = r[dateIdx];
    const rawExp  = r[expIdx];
    if (!name) continue;
    const key = code + '||' + name + '||' + pack;
    if (!drugMap[key]) {
      // แปลงวันหมดอายุ
      let expStr = '';
      if (rawExp instanceof Date) {
        expStr = Utilities.formatDate(rawExp, 'Asia/Bangkok', 'yyyy-MM-dd');
      } else if (rawExp && String(rawExp).includes('/')) {
        // dd/MM/yyyy → yyyy-MM-dd
        const parts = String(rawExp).split('/');
        if (parts.length === 3) expStr = `${parts[2]}-${parts[1]}-${parts[0]}`;
      } else {
        expStr = String(rawExp || '');
      }
      drugMap[key] = { name, pack, price, lot, code, expDate: expStr };
    }
  }
  // เรียงตามชื่อ
  const data = Object.values(drugMap).sort((a, b) =>
    String(a.name).localeCompare(String(b.name), 'th')
  );
  return { success: true, data };
}
// ============================================================
//  REORDER POINT MAP — อ่านจาก Formulary col 4
// ============================================================
function _getReorderMap(ss) {
  const sh = ss.getSheetByName(SHEETS.FORMULARY);
  if (!sh || sh.getLastRow() < 3) return {};
  const allData = sh.getDataRange().getValues();
  let dataStart = 2;
  for (let i = 0; i < Math.min(allData.length, 5); i++) {
    const cell = String(allData[i][0] || '').trim();
    if (cell === '' || cell.includes('บัญชี') || cell.toLowerCase().includes('ชื่อยา') ||
        cell.includes('📋')) { dataStart = i + 1; }
  }
  const map = {};
  for (let i = dataStart; i < allData.length; i++) {
    const name = String(allData[i][0] || '').trim().replace(/^[📋⚠️]\s*/u,'');
    const rp   = allData[i][3]; // col 4 = Reorder Point
    if (!name) continue;
    if (rp !== '' && rp !== undefined && rp !== null) {
      map[name] = Number(rp) || 0;
    }
  }
  return map;
}
// ============================================================
//  UPDATE REORDER POINT — แก้ไข manual จาก UI
// ============================================================
function _updateReorderPoint(body) {
  const { drugName, reorderPoint, oldValue } = body;
  if (!drugName) return { success: false, error: 'ไม่มีชื่อยา' };
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const formSh = _getOrCreate(ss, SHEETS.FORMULARY);
  const allData = formSh.getDataRange().getValues();
  let dataStart = 2;
  for (let i = 0; i < Math.min(allData.length, 5); i++) {
    const cell = String(allData[i][0] || '').trim();
    if (cell === '' || cell.includes('บัญชี') || cell.toLowerCase().includes('ชื่อยา') ||
        cell.includes('📋')) { dataStart = i + 1; }
  }
  let foundRow = -1;
  for (let i = dataStart; i < allData.length; i++) {
    const name = String(allData[i][0] || '').trim().replace(/^[📋⚠️]\s*/u,'');
    if (name === drugName.trim()) { foundRow = i + 1; break; }
  }
  if (foundRow > 0) {
    formSh.getRange(foundRow, 4).setValue(reorderPoint === '' ? '' : Number(reorderPoint));
  } else {
    formSh.appendRow([drugName, '', '', reorderPoint === '' ? '' : Number(reorderPoint)]);
  }
  // Sync กลับ NoLot sheet col 12
  const nlSh = ss.getSheetByName(SHEETS.SUMMARY_NOLOT);
  if (nlSh && nlSh.getLastRow() > 1) {
    const nlData = nlSh.getRange(2,1,nlSh.getLastRow()-1,12).getValues();
    for (let i = 0; i < nlData.length; i++) {
      if (String(nlData[i][0]||'').trim() === drugName.trim()) {
        nlSh.getRange(i+2, 12).setValue(reorderPoint === '' ? '' : Number(reorderPoint));
        break;
      }
    }
  }
  _writeAuditTrail('Reorder Point', drugName, String(oldValue||''), String(reorderPoint||''), 'ผู้ใช้');
  _logActivity(`อัปเดต Reorder Point: ${drugName} → ${reorderPoint}`, 1, 'ผู้ใช้');
  return { success: true, message: `บันทึก Reorder Point สำเร็จ` };
}
// ============================================================
//  UPDATE SUMMARY FIELD — inline edit ใน summary sheets
//  รองรับ: ชื่อยาเต็ม (sync ทุก sheet), Lot No, วันหมดอายุ
// ============================================================
function _updateSummaryField(body) {
  // body: { sheet, field, oldValue, newValue, lotNo, recorder }
  // sheet: 'สรุป' | 'สรุปแบบไม่แยก Lot'
  const { sheet, field, oldValue, newValue, lotNo, recorder } = body;
  if (!field || newValue === undefined) return { success: false, error: 'ข้อมูลไม่ครบ' };
  const allowedFields = ['ชื่อยาเต็ม', 'Lot No', 'วันหมดอายุ'];
  if (!allowedFields.includes(field)) return { success: false, error: 'ไม่อนุญาตให้แก้ไข field นี้' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let updatedCount = 0;
  if (sheet === SHEETS.SUMMARY || sheet === 'สรุปแบบไม่แยก Lot') {
    // แก้ใน sheet ที่ระบุ
    const sh = ss.getSheetByName(sheet);
    if (sh && sh.getLastRow() > 1) {
      const data = sh.getDataRange().getValues();
      const hdr  = data[0].map(h => String(h).trim());
      const fIdx = hdr.indexOf(field);
      if (fIdx >= 0) {
        for (let i = 1; i < data.length; i++) {
          const nameVal = String(data[i][hdr.indexOf('ชื่อยาเต็ม')] || '').trim();
          const lotVal  = String(data[i][hdr.indexOf('Lot No')]      || '').trim();
          const match   = nameVal === String(oldValue).trim()
            && (field !== 'Lot No' || !lotNo || lotVal === String(lotNo).trim());
          if (match) {
            sh.getRange(i+1, fIdx+1).setValue(newValue);
            updatedCount++;
          }
        }
      }
    }
  }
  // ถ้าแก้ "ชื่อยาเต็ม" → sync ไปทุก sheet ที่เกี่ยวข้อง
  if (field === 'ชื่อยาเต็ม') {
    const sheetsToSync = [SHEETS.SUMMARY, SHEETS.SUMMARY_NOLOT, SHEETS.TRANSACTION, SHEETS.DISPENSE];
    sheetsToSync.forEach(sName => {
      if (sName === sheet) return; // ทำไปแล้ว
      const sh = ss.getSheetByName(sName);
      if (!sh || sh.getLastRow() < 2) return;
      const data = sh.getDataRange().getValues();
      const hdr  = data[0].map(h => String(h).trim());
      const fIdx = hdr.indexOf('ชื่อยาเต็ม');
      if (fIdx < 0) return;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][fIdx]||'').trim() === String(oldValue).trim()) {
          sh.getRange(i+1, fIdx+1).setValue(newValue);
          updatedCount++;
        }
      }
    });
    // sync Formulary ด้วย
    const formSh = ss.getSheetByName(SHEETS.FORMULARY);
    if (formSh && formSh.getLastRow() > 2) {
      const fData = formSh.getDataRange().getValues();
      for (let i = 2; i < fData.length; i++) {
        const nm = String(fData[i][0]||'').trim().replace(/^[📋⚠️]\s*/u,'');
        if (nm === String(oldValue).trim()) {
          formSh.getRange(i+1, 1).setValue(newValue);
          updatedCount++;
        }
      }
    }
  }
  SpreadsheetApp.flush();
  _writeAuditTrail(sheet, field === 'ชื่อยาเต็ม' ? oldValue : (oldValue + ' / ' + field),
    String(oldValue), String(newValue), recorder || 'ผู้ใช้');
  _logActivity(`แก้ไข ${field}: "${oldValue}" → "${newValue}" (${updatedCount} rows)`, updatedCount, recorder || 'ผู้ใช้');
  return { success: true, updatedCount, message: `บันทึกสำเร็จ — อัปเดต ${updatedCount} รายการ` };
}
// ============================================================
//  AUDIT TRAIL WRITER
// ============================================================
function _writeAuditTrail(sheetOrPage, fieldOrItem, oldVal, newVal, user) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = _getOrCreate(ss, SHEETS.AUDIT_TRAIL);
    if (sh.getLastRow() === 0) {
      _writeHeader(sh, ['วันที่เวลา','ผู้แก้ไข','หน้า/Sheet','Field ที่แก้','ค่าเดิม','ค่าใหม่','หมายเหตุ'], '#7c3aed');
    }
    sh.appendRow([new Date(), user||'ระบบ', sheetOrPage, fieldOrItem, oldVal, newVal, '']);
  } catch(e){}
}
function _getAuditTrailData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.AUDIT_TRAIL);
  if (!sh || sh.getLastRow() < 2) return { success: true, data: [] };
  const data = sh.getDataRange().getValues();
  const hdr  = data[0];
  const rows = data.slice(1).reverse().slice(0, 200).map(r => {
    const obj = {};
    hdr.forEach((k,i) => {
      obj[String(k)] = r[i] instanceof Date
        ? Utilities.formatDate(r[i], 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
        : r[i];
    });
    return obj;
  });
  return { success: true, data: rows };
}
// ============================================================
//  SETTINGS — อ่าน/เขียน Settings sheet
// ============================================================
function _getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sh) { _setupSettingsSheet(ss); sh = ss.getSheetByName(SHEETS.SETTINGS); }
  if (!sh || sh.getLastRow() < 3) return { success: true, data: {} };
  const data = sh.getRange(3, 1, sh.getLastRow()-2, 2).getValues();
  const obj  = {};
  data.forEach(r => { if (r[0]) obj[String(r[0])] = String(r[1]||''); });
  return { success: true, data: obj };
}
function _saveSettings(body) {
  const { settings } = body;
  if (!settings) return { success: false, error: 'ไม่มีข้อมูล' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sh) { _setupSettingsSheet(ss); sh = ss.getSheetByName(SHEETS.SETTINGS); }
  if (sh.getLastRow() < 3) return { success: false, error: 'ไม่มีข้อมูล Settings sheet' };
  const data = sh.getRange(3, 1, sh.getLastRow()-2, 2).getValues();
  data.forEach((r, i) => {
    const key = String(r[0] || '');
    if (key && settings[key] !== undefined) {
      sh.getRange(i+3, 2).setValue(settings[key]);
    }
  });
  // ไม่เรียก ScriptApp ใน Web App mode — ผู้ใช้ต้องตั้ง Trigger ผ่าน GAS Editor แยกต่างหาก
  _logActivity('บันทึก Settings', 1, 'ผู้ใช้');
  return {
    success: true,
    message: 'บันทึก Settings สำเร็จ',
    triggerNote: 'กรุณาตั้ง Time Trigger ด้วยตนเองใน GAS Editor โดยรันฟังก์ชัน setupLineTriggerManual()'
  };
}
// ============================================================
//  LINE NOTIFICATION
// ============================================================
function _getSettingsMap() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sh || sh.getLastRow() < 3) return {};
  const data = sh.getRange(3, 1, sh.getLastRow()-2, 2).getValues();
  const map  = {};
  data.forEach(r => { if (r[0]) map[String(r[0])] = String(r[1]||''); });
  return map;
}
// ============================================================
//  LINE MESSAGING — ส่งข้อความผ่าน LINE Messaging API
// ============================================================
// ฟังก์ชันนี้รันจาก GAS Editor เพื่อ grant permission UrlFetchApp
// ก่อนใช้งานครั้งแรก ต้องรันฟังก์ชันนี้ใน GAS Editor ก่อน
function grantPermissionsForLine() {
  const results = [];
  // 1. UrlFetchApp — สำคัญที่สุด
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
      method: 'GET',
      muteHttpExceptions: true,
      headers: { 'Authorization': 'Bearer dummy' }
    });
    results.push('✅ UrlFetchApp: OK');
  } catch(e) {
    results.push('UrlFetchApp: ' + e.message);
  }
  // 2. ScriptApp
  try { ScriptApp.getProjectTriggers(); results.push('✅ ScriptApp: OK'); }
  catch(e) { results.push('ScriptApp: ' + e.message); }
  // 3. GmailApp (สำหรับ email)
  try { GmailApp.getInboxUnreadCount(); results.push('✅ GmailApp: OK'); }
  catch(e) { results.push('GmailApp: ' + e.message); }
  const msg = '✅ Authorization เสร็จแล้ว\n\n' + results.join('\n') +
    '\n\n⚠️ ถ้ายังส่ง LINE ไม่ได้:\n' +
    '1. ไปที่ GAS Editor → Project Settings → Scopes\n' +
    '2. ตรวจสอบว่ามี script.external_request\n' +
    '3. ถ้าไม่มี → ลบ trigger เก่า แล้วรัน setupLineTriggerManual() ใหม่';
  try { SpreadsheetApp.getUi().alert('Authorization Result', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch(e) { Logger.log(msg); }
  return msg;
}
function _sendLine(message) {
  const cfg    = _getSettingsMap();
  const token  = (cfg['LINE_CHANNEL_ACCESS_TOKEN'] || '').trim();
  const target = (cfg['LINE_TARGET_ID'] || '').trim();
  const type   = cfg['LINE_TARGET_TYPE'] || 'groupId';
  if (!token) return { ok: false, error: 'ไม่มี Channel Access Token', code: 0 };
  if (!target) return { ok: false, error: 'ไม่มี Target ID', code: 0 };
  // Validate token format - should be a long string
  if (token.length < 100) {
    return { ok: false, error: 'Token สั้นผิดปกติ (' + token.length + ' chars) — ตรวจสอบว่าใช้ Channel Access Token (Long-lived) ไม่ใช่ Channel Secret', code: 0 };
  }
  const payload = { to: target, messages: [{ type: 'text', text: message }] };
  try {
    const resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method:             'POST',
      contentType:        'application/json',
      headers: {
        'Authorization': 'Bearer ' + token,
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    _logActivity('LINE API Response: ' + code + ' | ' + body.substring(0,200), 0, 'ระบบ');
    if (code === 200) return { ok: true };
    let lineMsg = body;
    try { lineMsg = JSON.parse(body).message || body; } catch(e2) {}
    return { ok: false, error: 'LINE API ' + code + ': ' + lineMsg.substring(0,300), code: code, detail: body };
  } catch(e) {
    const msg = e.message || '';
    _logActivity('LINE Error: ' + msg, 0, 'ระบบ');
    if (msg.includes('script.external_request') || msg.includes('UrlFetchApp') || msg.includes('authorization')) {
      return { ok: false, error: 'PERMISSION_ERROR', code: 0, detail: msg };
    }
    return { ok: false, error: 'Network error: ' + msg, code: 0 };
  }
}
// ═══════════════════════════════════════════════════════════════════
//  sendDailyLineNotification — REQ #8: เพิ่มส่ง Email หลัง LINE
// ═══════════════════════════════════════════════════════════════════
function sendDailyLineNotification() {
  const cfg          = _getSettingsMap();
  const now          = new Date();
  const thaiDate     = _thaiDateTime(now);
  const notifyLow    = cfg['NOTIFY_LOW_STOCK']  !== 'FALSE';
  const notifyExpiry = cfg['NOTIFY_EXPIRY']     !== 'FALSE';
  const notifyStale  = cfg['NOTIFY_STALE_DATA'] !== 'FALSE';
  const expiryDays   = parseInt(cfg['EXPIRY_WARN_DAYS'] || '90');
  const staleDays    = parseInt(cfg['STALE_DATA_DAYS']  || '7');
  // อัพเดท Dashboard sheet ก่อน — เพื่อให้ข้อมูลเป็น single source of truth
  const dashData = writeDashboardSheet();
  // ตรวจวันในสัปดาห์ สำหรับ stale notify
  const stalePattern = cfg['STALE_NOTIFY_DAYS'] || '1111100';
  const dayOfWeek    = now.getDay();
  const dayIdx       = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  const shouldStale  = notifyStale && (stalePattern[dayIdx] === '1');
  const lines = [];
  lines.push('🏥 Sub-Stock รพ.รือเสาะ');
  lines.push('📅 ' + thaiDate);
  lines.push('─────────────────');
  // 1. ยาต่ำกว่า min / หมดสต็อก (อ่านจาก dashData)
  if (notifyLow) {
    const outList = dashData.outList   || [];
    const near50  = dashData.near50List || [];
    const lowList = dashData.lowList   || [];
    if (outList.length > 0) {
      lines.push('❌ หมดสต็อก (' + outList.length + ' รายการ)');
      outList.slice(0,10).forEach(r => lines.push('  • ' + r.name + ' (' + r.bal + '/' + r.min + ')'));
      if (outList.length > 10) lines.push('  ...และอีก ' + (outList.length - 10) + ' รายการ');
    }
    if (near50.length > 0) {
      lines.push('🔴 ต่ำกว่า 50% Min (' + near50.length + ' รายการ)');
      near50.slice(0,8).forEach(r => lines.push('  • ' + r.name + ' (' + r.bal + '/' + r.min + ')'));
      if (near50.length > 8) lines.push('  ...และอีก ' + (near50.length - 8) + ' รายการ');
    }
    if (lowList.length > 0) {
      lines.push('⚠️ ต่ำกว่า Min Stock (' + lowList.length + ' รายการ)');
      lowList.slice(0,8).forEach(r => lines.push('  • ' + r.name + ' (' + r.bal + '/' + r.min + ')'));
      if (lowList.length > 8) lines.push('  ...และอีก ' + (lowList.length - 8) + ' รายการ');
    }
    if (outList.length === 0 && near50.length === 0 && lowList.length === 0) {
      lines.push('✅ สต็อกปกติทุกรายการ');
    }
  }
  // 2. ยาใกล้หมดอายุ (REQ #11: รวมยาที่ stock=0 แต่ยังไม่ได้ยกเลิกติดตาม)
  if (notifyExpiry && dashData.expList && dashData.expList.length > 0) {
    const expList = dashData.expList;
    lines.push('─────────────────');
    lines.push('⏰ ยาใกล้หมดอายุ ≤' + expiryDays + ' วัน (' + expList.length + ' lot)');
    expList.slice(0,10).forEach(r => lines.push('  • ' + r.name + ' (หมด ' + r.expDate + ')'));
    if (expList.length > 10) lines.push('  ...และอีก ' + (expList.length - 10) + ' lot');
  }
  // 3. ข้อมูลไม่ได้อัพเดทนาน
  if (shouldStale) {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const txnSh = ss.getSheetByName(SHEETS.TRANSACTION);
    if (txnSh && txnSh.getLastRow() > 1) {
      const lastRow  = txnSh.getLastRow();
      const dateData = txnSh.getRange(2, 2, lastRow - 1, 1).getValues();
      let lastDate   = null;
      for (let i = dateData.length - 1; i >= 0; i--) {
        const v = dateData[i][0];
        if (v instanceof Date && !isNaN(v)) { lastDate = v; break; }
      }
      if (lastDate) {
        const diffDays = Math.floor((now - lastDate) / 86400000);
        if (diffDays >= staleDays) {
          lines.push('─────────────────');
          lines.push('📋 ยังไม่มีการอัพเดทข้อมูล');
          lines.push('   นาน ' + diffDays + ' วัน (ล่าสุด ' +
            Utilities.formatDate(lastDate, 'Asia/Bangkok', 'dd/MM/yyyy') + ')');
          lines.push('   กรุณาอัพโหลดข้อมูลบันทึกรับจ่ายด้วยนะครับ');
        }
      }
    }
  }
  lines.push('─────────────────');
  lines.push('🔗 Sub-Stock System');
  const msg    = lines.join('\n');
  const result = _sendLine(msg);
  _logActivity('ส่ง LINE แจ้งเตือนรายวัน', lines.length, result.ok ? 'สำเร็จ' : ('ล้มเหลว: ' + result.error));
  // ── REQ #8: ส่ง Email ด้วย (ถ้าตั้งค่าไว้) ──
  try {
    sendDailyEmailReport();
  } catch (eEmail) {
    _logActivity('Email error: ' + eEmail.message, 0, 'ระบบ');
  }
  return result;
}
function _sendLineTest() {
  const cfg    = _getSettingsMap();
  const token  = (cfg['LINE_CHANNEL_ACCESS_TOKEN'] || '').trim();
  const target = (cfg['LINE_TARGET_ID'] || '').trim();
  const type   = cfg['LINE_TARGET_TYPE'] || 'groupId';
  if (!token) return { success: false, message: 'ไม่มี Channel Access Token — กรุณากรอกและบันทึก Settings' };
  if (!target) return { success: false, message: 'ไม่มี Target ID — กรุณากรอก Group ID หรือ User ID' };
  if (token.length < 50) return { success: false, message: 'Token ดูสั้นผิดปกติ (' + token.length + ' chars) — ตรวจสอบ Token อีกครั้ง' };
  if (type === 'groupId' && !target.startsWith('C')) {
    return { success: false, message: 'Group ID ต้องขึ้นต้นด้วย "C" แต่ได้รับ "' + target.substring(0,4) + '..." — ตรวจสอบ Target ID' };
  }
  if (type === 'userId' && !target.startsWith('U')) {
    return { success: false, message: 'User ID ต้องขึ้นต้นด้วย "U" แต่ได้รับ "' + target.substring(0,4) + '..." — ตรวจสอบ Target ID' };
  }
  const msg = '🔔 ทดสอบ Sub-Stock รพ.รือเสาะ\n✅ ระบบเชื่อมต่อสำเร็จ\n📅 ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') + '\nTarget: ' + type + ' = ' + target.substring(0,10) + '...';
  const result = _sendLine(msg);
  if (result.ok) return { success: true, message: 'ส่งสำเร็จ! ตรวจสอบ LINE ได้เลย (Target: ' + type + ')' };
  // ตรวจ permission error
  if (result.error === 'PERMISSION_ERROR') {
    return {
      success: false,
      message: 'PERMISSION_REQUIRED',
      detail:  result.detail || '',
    };
  }
  let hint = '';
  if (result.code === 400) hint = ' | ⚠️ Bot อาจยังไม่ได้เป็นสมาชิกใน Group — เพิ่ม Bot เข้า Group ก่อน';
  if (result.code === 401) hint = ' | ⚠️ Token ไม่ถูกต้องหรือหมดอายุ';
  if (result.code === 403) hint = ' | ⚠️ Token ไม่มีสิทธิ์ push message';
  if (result.code === 429) hint = ' | ⚠️ Rate limit — รอสักครู่แล้วลองใหม่';
  return {
    success: false,
    message: (result.error || 'ส่งไม่สำเร็จ') + hint,
  };
}
// ============================================================
//  TIME TRIGGER — ตั้ง/ลบ trigger อัตโนมัติ
// ============================================================
function _setupLineTrigger(hourStr) {
  const hour = parseInt(hourStr) || 8;
  // ลบ trigger เดิมที่ชื่อ sendDailyLineNotification
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyLineNotification') ScriptApp.deleteTrigger(t);
  });
  // สร้าง trigger ใหม่
  ScriptApp.newTrigger('sendDailyLineNotification')
    .timeBased().everyDays(1).atHour(hour).create();
}
function setupLineTriggerManual() {
  const cfg  = _getSettingsMap();
  const hour = parseInt(cfg['NOTIFY_HOUR'] || '8');
  _setupLineTrigger(hour);
  SpreadsheetApp.getUi().alert('✅ ตั้ง Trigger สำเร็จ', 'จะส่ง LINE ทุกวัน เวลา ' + hour + ':00 น.', SpreadsheetApp.getUi().ButtonSet.OK);
}
// ============================================================
//  MONTHLY EXPORT DATA — ส่งข้อมูลรายเดือนไปให้ frontend
// ============================================================
function _getMonthlyData(monthStr, yearStr) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const txnSh = ss.getSheetByName(SHEETS.TRANSACTION);
  if (!txnSh || txnSh.getLastRow() < 2) return { success: true, data: [], summary: {} };
  const month = parseInt(monthStr) || (new Date().getMonth() + 1);
  const year  = parseInt(yearStr)  || new Date().getFullYear();
  const allData = txnSh.getDataRange().getValues();
  const hdr     = allData[0].map(h => String(h).trim());
  const dateIdx = hdr.indexOf('วันที่');
  const nameIdx = hdr.indexOf('ชื่อยาเต็ม');
  const qtyIdx  = hdr.indexOf('จำนวน');
  const valIdx  = hdr.indexOf('มูลค่ารวม');
  const typeIdx = hdr.indexOf('ประเภทธุรกรรม');
  const lotIdx  = hdr.indexOf('Lot No');
  const codeIdx = hdr.indexOf('รหัสยา');
  const rows = [];
  let totalIn = 0, totalOut = 0, totalVal = 0;
  for (let i = 1; i < allData.length; i++) {
    const r   = allData[i];
    const raw = r[dateIdx];
    let d;
    if (raw instanceof Date) d = raw;
    else { d = new Date(raw); if (isNaN(d)) continue; }
    if (d.getFullYear() !== year || d.getMonth() + 1 !== month) continue;
    const qty  = Number(r[qtyIdx] || 0);
    const val  = Number(r[valIdx] || 0);
    const type = String(r[typeIdx] || '');
    if (type === 'รับเข้า' || qty > 0) totalIn  += Math.abs(qty);
    else                                 totalOut += Math.abs(qty);
    totalVal += val;
    rows.push({
      วันที่:        Utilities.formatDate(d, 'Asia/Bangkok', 'dd/MM/yyyy'),
      รหัสยา:        String(r[codeIdx] || ''),
      ชื่อยาเต็ม:    String(r[nameIdx] || ''),
      จำนวน:         qty,
      มูลค่ารวม:     val,
      LotNo:         String(r[lotIdx]  || ''),
      ประเภทธุรกรรม: type,
    });
  }
  // สรุปรายยา
  const drugSummary = {};
  rows.forEach(r => {
    const k = r['ชื่อยาเต็ม'];
    if (!drugSummary[k]) drugSummary[k] = { name: k, in: 0, out: 0, val: 0 };
    if (r['จำนวน'] > 0) drugSummary[k].in  += r['จำนวน'];
    else                 drugSummary[k].out += Math.abs(r['จำนวน']);
    drugSummary[k].val += r['มูลค่ารวม'];
  });
  return {
    success: true,
    month, year,
    hospitalName: HOSPITAL_NAME,
    fiscalYear:   FISCAL_YEAR,
    data: rows,
    drugSummary: Object.values(drugSummary).sort((a,b) => String(a.name).localeCompare(String(b.name),'th')),
    summary: { totalRows: rows.length, totalIn, totalOut, totalVal: _r2(totalVal) },
  };
}
// ============================================================
//  GET NEXT DISP NO — Auto running number ปีงบประมาณ
// ============================================================
function _getNextDispNo() {
  // ปีงบประมาณไทย: ต.ค.(10)–ก.ย.(9) ของปีถัดไป
  // เช่น ต.ค.2568–ก.ย.2569 = ปีงบ 2569 → short = "69"
  const now  = new Date();
  const m    = now.getMonth() + 1; // 1–12
  const y    = now.getFullYear() + 543; // พ.ศ.
  const fy   = m >= 10 ? y + 1 : y;
  const fyShort = String(fy).slice(-2); // "68","69"...
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const txn = ss.getSheetByName(SHEETS.TRANSACTION);
  let maxNum = 0;
  if (txn && txn.getLastRow() > 1) {
    const vals = txn.getRange(2, 1, txn.getLastRow() - 1, 1).getValues();
    vals.forEach(r => {
      const v = String(r[0] || '');
      const m2 = v.match(/^(\d{2})-(\d+)$/);
      if (m2 && m2[1] === fyShort) {
        maxNum = Math.max(maxNum, parseInt(m2[2]) || 0);
      }
    });
  }
  const next = String(maxNum + 1).padStart(3, '0');
  return { success: true, dispNo: `${fyShort}-${next}` };
}
// ============================================================
//  GET DISP USERS — ดึงรายชื่อผู้ใช้จาก Settings sheet
//  (key: DISP_USERS, value: "ชื่อ1|ชื่อ2|ชื่อ3")
//  และ departments (key: DEPARTMENTS)
// ============================================================
function _getDispUsers() {
  const cfg    = _getSettingsMap();
  const raw    = cfg['DISP_USERS'] || '';
  const users  = raw ? raw.split('|').map(u => u.trim()).filter(Boolean) : [];
  const rawD   = cfg['DEPARTMENTS'] || '';
  const depts  = rawD ? rawD.split('|').map(d => d.trim()).filter(Boolean)
               : ['ห้องยาใน','ห้องยานอก','Ward 1','Ward 2','LR','ER'];
  return { success: true, users, departments: depts };
}
// ============================================================
//  BATCH DISPENSE — บันทึกการเบิกยาหลายรายการพร้อมกัน
// ============================================================
function _batchDispense(body) {
  /*
    body: {
      dispNo, date, dept, recorder,
      items: [{ drugName, lotNo, qty, packSize, unitPrice, expDate, code }]
    }
  */
  const { dispNo, date, dept, recorder, items } = body;
  if (!dispNo)                 return { success: false, error: 'ไม่มี Disp No' };
  if (!recorder)               return { success: false, error: 'ไม่มีผู้บันทึก' };
  if (!items || !items.length) return { success: false, error: 'ไม่มีรายการยา' };
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const txn = _getOrCreate(ss, SHEETS.TRANSACTION);
  const txnHdr = txn.getLastRow() > 0
    ? txn.getRange(1, 1, 1, txn.getLastColumn()).getValues()[0]
    : [];
  const tIdx = h => txnHdr.findIndex(c => String(c).trim() === h);
  // parse date
  let dateObj;
  try { dateObj = date ? new Date(date) : new Date(); }
  catch(e) { dateObj = new Date(); }
  const dateOnly = Utilities.formatDate(dateObj, 'Asia/Bangkok', 'dd/MM/yyyy');
  const timeStr  = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  const rowsToAdd = [];
  for (const item of items) {
    const qty = Number(item.qty) || 0;
    if (!qty) continue;
    const dispQty   = -Math.abs(qty); // เบิก = ติดลบ
    const unitPrice = Number(item.unitPrice) || 0;
    const total     = Math.abs(dispQty) * unitPrice;
    const newRow = new Array(Math.max(txnHdr.length, 15)).fill('');
    const set    = (col, val) => { const i = tIdx(col); if (i >= 0) newRow[i] = val; };
    set('Disp No',        dispNo);
    set('วันที่',          dateOnly);
    set('แผนก',           dept || '');
    set('รหัสยา',          item.code || '');
    set('ชื่อยาเต็ม',      item.drugName);
    set('จำนวน',           dispQty);
    set('Pack',           Number(item.packSize) || 1);
    set('ราคา/หน่วย',      unitPrice);
    set('มูลค่ารวม',       total);
    set('Lot No',         item.lotNo || '');
    set('วันหมดอายุ',      item.expDate || '');
    set('ประเภทธุรกรรม',   'เบิกยา');
    set('ผู้บันทึก',        recorder);
    set('เวลาบันทึก',       timeStr);
    rowsToAdd.push(newRow);
  }
  if (!rowsToAdd.length) return { success: false, error: 'ไม่มีรายการที่มีจำนวน' };
  // Batch append ในครั้งเดียว (เร็วกว่า appendRow ทีละบรรทัด)
  const startRow = txn.getLastRow() + 1;
  txn.getRange(startRow, 1, rowsToAdd.length, txnHdr.length || 15).setValues(rowsToAdd);
  SpreadsheetApp.flush();
  refreshSummarySheets();
  refreshDispenseSheet();
  _logActivity(
    `เบิกยา ${dispNo} — ${rowsToAdd.length} รายการ (${dept})`,
    rowsToAdd.length,
    recorder
  );
  _writeAuditTrail('บันทึกรับจ่าย', 'เบิกยา Disp No: ' + dispNo,
    '', rowsToAdd.length + ' รายการ', recorder);
  return {
    success: true,
    dispNo,
    count: rowsToAdd.length,
    message: `บันทึกการเบิก ${dispNo} สำเร็จ — ${rowsToAdd.length} รายการ`,
  };
}
// ============================================================
//  UPDATE MIN STOCK — แก้ไข Minimum Stock ลง Formulary
// ============================================================
function _updateMinStock(body) {
  // body: { drugName, newMin }
  const { drugName, newMin } = body;
  if (!drugName) return { success: false, error: 'ไม่มีชื่อยา' };
  const minVal = newMin === '' ? '' : Number(newMin);
  if (newMin !== '' && isNaN(minVal)) return { success: false, error: 'Min Stock ต้องเป็นตัวเลข' };
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const formSh = _getOrCreate(ss, SHEETS.FORMULARY);
  const allData = formSh.getDataRange().getValues();
  // หา dataStart (เหมือน _updateReorderPoint)
  let dataStart = 2;
  for (let i = 0; i < Math.min(allData.length, 5); i++) {
    const cell = String(allData[i][0] || '').trim();
    if (cell === '' || cell.includes('บัญชี') || cell.toLowerCase().includes('ชื่อยา') ||
        cell.includes('📋')) { dataStart = i + 1; }
  }
  const oldMin = (() => {
    for (let i = dataStart; i < allData.length; i++) {
      const n = String(allData[i][0]||'').trim().replace(/^[📋⚠️]\s*/u,'');
      if (n === drugName.trim()) return allData[i][1];
    }
    return '';
  })();
  let foundRow = -1;
  for (let i = dataStart; i < allData.length; i++) {
    const n = String(allData[i][0]||'').trim().replace(/^[📋⚠️]\s*/u,'');
    if (n === drugName.trim()) { foundRow = i + 1; break; }
  }
  if (foundRow > 0) {
    formSh.getRange(foundRow, 2).setValue(minVal);
  } else {
    formSh.appendRow([drugName, minVal, '', '']);
  }
  SpreadsheetApp.flush();
  // Recalc เพื่ออัปเดต สถานะสต็อก
  refreshSummarySheets();
  _writeAuditTrail('บัญชีโรงพยาบาล', 'Min Stock: ' + drugName,
    String(oldMin), String(minVal), 'ผู้ใช้');
  _logActivity(`อัปเดต Min Stock: ${drugName} → ${minVal}`, 1, 'ผู้ใช้');
  return { success: true, message: `บันทึก Min Stock "${drugName}" = ${minVal} สำเร็จ` };
}
// ============================================================
//  GET PROCURE LIST — ยาที่ต่ำกว่า Min Stock
// ============================================================
function _getProcureList() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const noLot = ss.getSheetByName(SHEETS.SUMMARY_NOLOT);
  if (!noLot || noLot.getLastRow() < 2) return { success: true, data: [] };
  const data = noLot.getDataRange().getValues();
  const hdr  = data[0].map(h => String(h).trim());
  const nIdx  = h => hdr.indexOf(h);
  const nameI = nIdx('ชื่อยาเต็ม');
  const balI  = hdr.findIndex(h => h.includes('คงเหลือปัจจุบัน'));
  const minI  = nIdx('Minimum Stock');
  const stI   = nIdx('สถานะสต็อก');
  const codeI = nIdx('รหัสยา');
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameI] || '').trim();
    const bal  = Number(data[i][balI]  || 0);
    const min  = Number(data[i][minI]  || 0);
    const st   = String(data[i][stI]   || '');
    const code = codeI >= 0 ? String(data[i][codeI] || '') : '';
    if (!name || min <= 0) continue;
    if (bal >= min)        continue; // ยังพอ ไม่ต้องเบิก
    rows.push({ name, bal, min, status: st, code });
  }
  return { success: true, data: rows };
}
// ============================================================
//  GET USAGE DATA — อัตราการใช้ยาในช่วงวันที่ระบุ
// ============================================================
function _getUsageData(fromStr, toStr) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const txn = ss.getSheetByName(SHEETS.TRANSACTION);
  if (!txn || txn.getLastRow() < 2) return { success: true, usage: {} };
  let fromDate, toDate;
  try {
    fromDate = fromStr ? new Date(fromStr) : new Date(Date.now() - 30*24*60*60*1000);
    toDate   = toStr   ? new Date(toStr)   : new Date();
    // normalize toDate ให้เป็น end of day
    toDate.setHours(23,59,59,999);
  } catch(e) {
    fromDate = new Date(Date.now() - 30*24*60*60*1000);
    toDate   = new Date();
  }
  const days = Math.max(1, Math.ceil((toDate - fromDate) / (1000*60*60*24)));
  const data    = txn.getDataRange().getValues();
  const hdr     = data[0].map(h => String(h).trim());
  const nameIdx = hdr.indexOf('ชื่อยาเต็ม');
  const dateIdx = hdr.indexOf('วันที่');
  const qtyIdx  = hdr.indexOf('จำนวน');
  const typeIdx = hdr.indexOf('ประเภทธุรกรรม');
  const usage = {}; // { drugName: totalOut }
  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][typeIdx] || '');
    if (type !== 'เบิกยา') continue;
    const rv = data[i][dateIdx];
    if (!rv) continue;
    const d = rv instanceof Date ? rv : new Date(rv);
    if (d < fromDate || d > toDate) continue;
    const name = String(data[i][nameIdx] || '').trim();
    const qty  = Number(data[i][qtyIdx]  || 0);
    if (!name || qty >= 0) continue; // qty < 0 = เบิก
    usage[name] = (usage[name] || 0) + Math.abs(qty);
  }
  return { success: true, usage, days };
}
// สร้าง/อ่าน sheet ติดตามหมดอายุ
function _setupExpiryUntrackSheet(ss) {
  const sh = _getOrCreate(ss, SHEETS.EXPIRY_UNTRACK);
  if (sh.getLastRow() === 0) {
    _writeHeader(sh, ['ชื่อยาเต็ม', 'Lot No', 'วันหมดอายุ', 'วันที่ยกเลิก', 'ผู้ยกเลิก'], '#6b7280');
  }
  return sh;
}
// ยกเลิกติดตาม — เพิ่มลง sheet ติดตามหมดอายุ
function _untrackExpiry(body) {
  const { drugName, lotNo, expDate, recorder } = body;
  if (!drugName) return { success: false, error: 'ไม่มีชื่อยา' };
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = _setupExpiryUntrackSheet(ss);
  const now = new Date();
  // ตรวจซ้ำ
  if (sh.getLastRow() > 1) {
    const existing = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
    for (const row of existing) {
      if (String(row[0]).trim() === drugName.trim() && String(row[1]).trim() === String(lotNo||'').trim()) {
        return { success: true, message: 'รายการนี้ยกเลิกติดตามไปแล้ว' };
      }
    }
  }
  sh.appendRow([drugName, lotNo||'', expDate||'',
    Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'),
    recorder||'ผู้ใช้']);
  _logActivity(`ยกเลิกติดตาม: ${drugName} (${lotNo})`, 1, recorder||'ผู้ใช้');
  return { success: true, message: `ยกเลิกติดตาม ${drugName} (${lotNo}) แล้ว` };
}
// ดึง untracked list
function _getUntrackedExpiry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.EXPIRY_UNTRACK);
  if (!sh || sh.getLastRow() < 2) return { success: true, data: [] };
  const data = sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues();
  return {
    success: true,
    data: data.map(r => ({
      drugName: String(r[0]||'').trim(),
      lotNo:    String(r[1]||'').trim(),
      expDate:  String(r[2]||''),
      cancelDate: String(r[3]||''),
      recorder:   String(r[4]||''),
    })).filter(r => r.drugName),
  };
}
// คืนค่าการติดตาม (ลบออกจาก sheet)
function _retrackExpiry(body) {
  const { drugName, lotNo } = body;
  if (!drugName) return { success: false, error: 'ไม่มีชื่อยา' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.EXPIRY_UNTRACK);
  if (!sh || sh.getLastRow() < 2) return { success: false, error: 'ไม่พบข้อมูล' };
  const data = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === drugName.trim() &&
        String(data[i][1]).trim() === String(lotNo||'').trim()) {
      sh.deleteRow(i + 2);
      return { success: true, message: `คืนการติดตาม ${drugName} แล้ว` };
    }
  }
  return { success: false, error: 'ไม่พบรายการที่จะคืน' };
}
function sendDailyEmailReport() {
  const cfg      = _getSettingsMap();
  const rawEmails = (cfg['EMAIL_RECIPIENTS'] || '').trim();
  const notifyEmail = cfg['NOTIFY_EMAIL'] !== 'FALSE';
  if (!notifyEmail || !rawEmails) {
    _logActivity('ข้ามส่ง Email (ปิดหรือไม่มี recipients)', 0, 'ระบบ');
    return { ok: false, error: 'Email ไม่ได้ตั้งค่า' };
  }
  const emails = rawEmails.split(/[,;|\n]/).map(e => e.trim()).filter(Boolean);
  if (!emails.length) return { ok: false, error: 'ไม่มี email recipients' };
  const dashData = writeDashboardSheet();
  const now      = new Date();
  const thaiDate = _thaiDateTime(now);
  const cfg2     = _getSettingsMap();
  const expiryDays = parseInt(cfg2['EXPIRY_WARN_DAYS'] || '90');
  const outList  = dashData.outList   || [];
  const near50   = dashData.near50List || [];
  const lowList  = dashData.lowList   || [];
  const expList  = dashData.expList   || [];
  // ── สร้าง HTML email ──
  const html = _buildEmailHtml({
    thaiDate, outList, near50, lowList, expList, expiryDays,
    hospitalName: 'โรงพยาบาลรือเสาะ',
    fiscalYear:   FISCAL_YEAR,
  });
  const subject = `[Sub-Stock] รายงานสต็อกยา ${thaiDate} — รพ.รือเสาะ`;
  let ok = 0, fail = 0;
  emails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, '', { htmlBody: html, name: 'Sub-Stock System' });
      ok++;
    } catch(e) { fail++; _logActivity('ส่ง Email ล้มเหลว: ' + email, 0, e.message); }
  });
  _logActivity(`ส่ง Email รายวัน (${ok} สำเร็จ, ${fail} ล้มเหลว)`, ok, 'ระบบ');
  return { ok: ok > 0, sent: ok, failed: fail };
}
// ทดสอบส่ง email
function _sendEmailTest(body) {
  const cfg    = _getSettingsMap();
  const rawEmails = (body.testEmail || cfg['EMAIL_RECIPIENTS'] || '').trim();
  if (!rawEmails) return { success: false, message: 'ไม่มี Email Recipients — กรุณากรอกและบันทึก Settings ก่อน' };
  const emails = rawEmails.split(/[,;|\n]/).map(e => e.trim()).filter(Boolean);
  const dashData = writeDashboardSheet();
  const now      = new Date();
  const thaiDate = _thaiDateTime(now);
  const cfg2     = _getSettingsMap();
  const expiryDays = parseInt(cfg2['EXPIRY_WARN_DAYS'] || '90');
  const html = _buildEmailHtml({
    thaiDate,
    outList:  dashData.outList   || [],
    near50:   dashData.near50List || [],
    lowList:  dashData.lowList   || [],
    expList:  dashData.expList   || [],
    expiryDays,
    hospitalName: 'โรงพยาบาลรือเสาะ',
    fiscalYear:   FISCAL_YEAR,
    isTest: true,
  });
  const subject = `[ทดสอบ] Sub-Stock รายงานสต็อกยา — ${thaiDate}`;
  let ok = 0, fail = 0, failList = [];
  emails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, '', { htmlBody: html, name: 'Sub-Stock System' });
      ok++;
    } catch(e) { fail++; failList.push(email + ': ' + e.message); }
  });
  if (ok > 0) return {
    success: true,
    message: `ส่งทดสอบสำเร็จ ${ok} email (${emails.join(', ')})`,
  };
  return {
    success: false,
    message: 'ส่งไม่สำเร็จ: ' + failList.join('; '),
  };
}
// สร้าง HTML email body
function _buildEmailHtml(opts) {
  const { thaiDate, outList, near50, lowList, expList, expiryDays, hospitalName, isTest } = opts;
  const totalAlert = outList.length + near50.length + lowList.length;
  // ── color palette ──
  const C = {
    bg:     '#0f172a',
    card:   '#1e293b',
    border: '#334155',
    teal:   '#0ff4c6',
    red:    '#f43f5e',
    amber:  '#f59e0b',
    orange: '#f97316',
    green:  '#10b981',
    text:   '#e2e8f0',
    muted:  '#94a3b8',
    white:  '#ffffff',
  };
  // ── helper: สร้าง pill badge ──
  const pill = (label, color, bg) =>
    `<span style="display:inline-block;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700;color:${color};background:${bg};border:1px solid ${color}40">${label}</span>`;
  // ── helper: section header ──
  const secHead = (icon, label, count, color) =>
    `<div style="background:${color}18;border-left:3px solid ${color};border-radius:0 8px 8px 0;padding:10px 14px;margin:16px 0 8px;display:flex;align-items:center;gap:10px">
      <span style="font-size:18px">${icon}</span>
      <span style="color:${color};font-weight:700;font-size:14px">${label}</span>
      <span style="margin-left:auto;background:${color};color:#fff;border-radius:12px;padding:1px 9px;font-size:11px;font-weight:700">${count}</span>
    </div>`;
  // ── helper: drug row ──
  const drugRow = (name, sub, badge, i) =>
    `<tr style="background:${i%2===0?C.card:'#253047'}">
      <td style="padding:7px 14px;font-size:13px;color:${C.text};border-bottom:1px solid ${C.border}30;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${name}</td>
      <td style="padding:7px 10px;font-size:12px;color:${C.muted};border-bottom:1px solid ${C.border}30;white-space:nowrap">${sub}</td>
      <td style="padding:7px 10px;text-align:right;border-bottom:1px solid ${C.border}30">${badge}</td>
    </tr>`;
  // ── stat card ──
  const statCard = (label, val, color) =>
    `<td style="width:25%;padding:0 6px">
      <div style="background:${color}18;border:1px solid ${color}40;border-radius:10px;padding:14px 10px;text-align:center">
        <div style="font-size:26px;font-weight:800;color:${color};font-family:monospace">${val}</div>
        <div style="font-size:11px;color:${C.muted};margin-top:3px">${label}</div>
      </div>
    </td>`;
  // ── content sections ──
  let content = '';
  // out of stock
  if (outList.length > 0) {
    content += secHead('❌', 'หมดสต็อก', outList.length, C.red);
    content += `<table width="100%" cellspacing="0" cellpadding="0" style="border-radius:8px;overflow:hidden;border:1px solid ${C.border}">`;
    outList.slice(0,15).forEach((r,i) => {
      content += drugRow(r.name, `คงเหลือ: 0 / Min: ${r.min}`, pill('หมดสต็อก', C.red, C.red+'18'), i);
    });
    if (outList.length > 15) content += `<tr><td colspan="3" style="padding:6px 14px;font-size:12px;color:${C.muted};text-align:center">...และอีก ${outList.length-15} รายการ</td></tr>`;
    content += '</table>';
  }
  // < 50% min
  if (near50.length > 0) {
    content += secHead('🔴', 'ต่ำกว่า 50% Min Stock', near50.length, C.orange);
    content += `<table width="100%" cellspacing="0" cellpadding="0" style="border-radius:8px;overflow:hidden;border:1px solid ${C.border}">`;
    near50.slice(0,12).forEach((r,i) => {
      content += drugRow(r.name, `คงเหลือ: ${r.bal} / Min: ${r.min}`, pill('เฝ้าระวัง', C.orange, C.orange+'18'), i);
    });
    if (near50.length > 12) content += `<tr><td colspan="3" style="padding:6px 14px;font-size:12px;color:${C.muted};text-align:center">...และอีก ${near50.length-12} รายการ</td></tr>`;
    content += '</table>';
  }
  // below min
  if (lowList.length > 0) {
    content += secHead('⚠️', 'ต่ำกว่า Minimum Stock', lowList.length, C.amber);
    content += `<table width="100%" cellspacing="0" cellpadding="0" style="border-radius:8px;overflow:hidden;border:1px solid ${C.border}">`;
    lowList.slice(0,12).forEach((r,i) => {
      content += drugRow(r.name, `คงเหลือ: ${r.bal} / Min: ${r.min}`, pill('ต่ำกว่า Min', C.amber, C.amber+'18'), i);
    });
    if (lowList.length > 12) content += `<tr><td colspan="3" style="padding:6px 14px;font-size:12px;color:${C.muted};text-align:center">...และอีก ${lowList.length-12} รายการ</td></tr>`;
    content += '</table>';
  }
  if (totalAlert === 0) {
    content += `<div style="text-align:center;padding:30px;background:${C.green}18;border:1px solid ${C.green}40;border-radius:12px;margin:16px 0">
      <div style="font-size:32px;margin-bottom:8px">✅</div>
      <div style="color:${C.green};font-weight:700;font-size:16px">สต็อกปกติทุกรายการ</div>
      <div style="color:${C.muted};font-size:13px;margin-top:4px">ไม่มีรายการที่ต้องดำเนินการ</div>
    </div>`;
  }
  // expiry
  if (expList.length > 0) {
    content += secHead('⏰', `ยาใกล้หมดอายุ ≤${expiryDays} วัน`, expList.length, C.teal);
    content += `<table width="100%" cellspacing="0" cellpadding="0" style="border-radius:8px;overflow:hidden;border:1px solid ${C.border}">`;
    expList.slice(0,12).forEach((r,i) => {
      const dColor = r.days <= 30 ? C.red : r.days <= 60 ? C.orange : C.amber;
      content += drugRow(r.name, `หมด: ${r.expDate}`, pill(`${r.days} วัน`, dColor, dColor+'18'), i);
    });
    if (expList.length > 12) content += `<tr><td colspan="3" style="padding:6px 14px;font-size:12px;color:${C.muted};text-align:center">...และอีก ${expList.length-12} lot</td></tr>`;
    content += '</table>';
  }
  // ── full HTML ──
  return `<!DOCTYPE html>
<html lang="th">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sub-Stock Report</title></head>
<body style="margin:0;padding:0;background:#0a0f1e;font-family:'Helvetica Neue',Arial,sans-serif">
<div style="max-width:640px;margin:0 auto;padding:24px 16px">
  ${isTest ? `<div style="background:#f59e0b18;border:1px solid #f59e0b;border-radius:8px;padding:10px 16px;margin-bottom:16px;color:#f59e0b;font-size:13px;font-weight:700">🧪 นี่คือ Email ทดสอบ — ไม่ใช่รายงานจริง</div>` : ''}
  <!-- Header -->
  <div style="background:linear-gradient(135deg,#1e293b,#0f172a);border:1px solid #334155;border-radius:16px;padding:24px;margin-bottom:16px;text-align:center">
    <div style="font-size:28px;margin-bottom:6px">💊</div>
    <div style="font-size:20px;font-weight:800;color:#ffffff;letter-spacing:-0.02em">Sub-Stock System</div>
    <div style="font-size:13px;color:#94a3b8;margin-top:4px">${hospitalName} · ปีงบ ${FISCAL_YEAR}</div>
    <div style="font-size:12px;color:#64748b;margin-top:2px">📅 ${thaiDate}</div>
  </div>
  <!-- Stat cards -->
  <table width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:16px">
    <tr>
      ${statCard('หมดสต็อก', outList.length, C.red)}
      ${statCard('ต่ำกว่า 50%', near50.length, C.orange)}
      ${statCard('ต่ำกว่า Min', lowList.length, C.amber)}
      ${statCard('ใกล้หมดอายุ', expList.length, C.teal)}
    </tr>
  </table>
  <!-- Main content -->
  <div style="background:#1e293b;border:1px solid #334155;border-radius:16px;padding:20px;margin-bottom:16px">
    ${content || `<div style="text-align:center;color:#64748b;padding:20px">ไม่มีรายการแจ้งเตือน</div>`}
  </div>
  <!-- Footer -->
  <div style="text-align:center;color:#475569;font-size:11px;padding:12px 0">
    <div>รายงานจาก Sub-Stock System — ${hospitalName}</div>
    <div style="margin-top:4px">ส่งอัตโนมัติโดย Google Apps Script · ${thaiDate}</div>
  </div>
</div>
</body></html>`;
}

// ============================================================
//  UPDATE FORMULARY — แก้ไข Min Stock, ตู้ยา, Reorder Point inline
// ============================================================
function _updateFormulary(body) {
  const { drugName, field, value } = body;
  if (!drugName || !field) return { success: false, error: 'ข้อมูลไม่ครบ' };
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const formSh = _getOrCreate(ss, SHEETS.FORMULARY);
  const allData = formSh.getDataRange().getValues();
  // หา dataStart
  let dataStart = 2;
  for (let i = 0; i < Math.min(allData.length, 5); i++) {
    const cell = String(allData[i][0] || '').trim();
    if (cell === '' || cell.includes('บัญชี') || cell.toLowerCase().includes('ชื่อยา') || cell.includes('📋')) {
      dataStart = i + 1;
    }
  }
  // field → column index mapping
  const colMap = { 'minStock': 2, 'cabinet': 3, 'reorderPoint': 4, 'shelfPosition': 5 };
  const col = colMap[field];
  if (!col) return { success: false, error: 'Field ไม่ถูกต้อง: ' + field };
  // หาแถวที่ตรงกับ drugName
  let foundRow = -1;
  let oldVal = '';
  for (let i = dataStart; i < allData.length; i++) {
    const n = String(allData[i][0] || '').trim().replace(/^[📋⚠️🔔📍]\s*/u, '');
    if (n === drugName.trim()) {
      foundRow = i + 1;
      oldVal = String(allData[i][col - 1] || '');
      break;
    }
  }
  const newVal = (field === 'minStock' || field === 'reorderPoint') && value !== ''
    ? Number(value) : value;
  if (foundRow > 0) {
    formSh.getRange(foundRow, col).setValue(newVal);
  } else {
    // สร้างแถวใหม่
    const newRow = [drugName, '', '', '', ''];
    newRow[col - 1] = newVal;
    formSh.appendRow(newRow);
  }
  SpreadsheetApp.flush();
  if (field === 'minStock') {
    refreshSummarySheets();
  }
  _writeAuditTrail('บัญชีโรงพยาบาล', field + ': ' + drugName, oldVal, String(newVal), 'ผู้ใช้');
  return { success: true, message: 'อัปเดต ' + field + ' สำเร็จ' };
}

// ============================================================
//  MONTHLY TREND — ค่าใช้จ่ายยารายเดือน (12 เดือนล่าสุด)
// ============================================================
function _getMonthlyTrend() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const txn = ss.getSheetByName(SHEETS.TRANSACTION);
  if (!txn || txn.getLastRow() < 2) return { success: true, data: [] };
  const allData = txn.getDataRange().getValues();
  const hdr     = allData[0].map(function(h) { return String(h).trim(); });
  const dateIdx = hdr.indexOf('วันที่');
  const valIdx  = hdr.indexOf('มูลค่ารวม');
  const typeIdx = hdr.indexOf('ประเภทธุรกรรม');
  const qtyIdx  = hdr.indexOf('จำนวน');
  // Aggregate by month
  var monthMap = {};
  for (var i = 1; i < allData.length; i++) {
    var raw = allData[i][dateIdx];
    var d;
    if (raw instanceof Date) d = raw;
    else { d = new Date(raw); if (isNaN(d)) continue; }
    var key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
    if (!monthMap[key]) monthMap[key] = { inVal: 0, outVal: 0, inQty: 0, outQty: 0 };
    var val = Math.abs(Number(allData[i][valIdx] || 0));
    var qty = Number(allData[i][qtyIdx] || 0);
    if (qty > 0) {
      monthMap[key].inVal += val;
      monthMap[key].inQty += qty;
    } else {
      monthMap[key].outVal += val;
      monthMap[key].outQty += Math.abs(qty);
    }
  }
  // Sort by key and take last 12 months
  var keys = Object.keys(monthMap).sort();
  var last12 = keys.slice(-12);
  var thaiMonths = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  var result = last12.map(function(k) {
    var parts = k.split('-');
    var m = parseInt(parts[1]) - 1;
    var y = parseInt(parts[0]) + 543;
    return {
      month: thaiMonths[m] + ' ' + String(y).slice(-2),
      key: k,
      inVal: Math.round(monthMap[k].inVal),
      outVal: Math.round(monthMap[k].outVal),
      inQty: monthMap[k].inQty,
      outQty: monthMap[k].outQty,
    };
  });
  return { success: true, data: result };
}
function writeDashboardSheet() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  let sh = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sh) sh = ss.insertSheet(SHEETS.DASHBOARD);
  sh.clearContents();
  const nolotSh = ss.getSheetByName(SHEETS.SUMMARY_NOLOT);
  const sumSh   = ss.getSheetByName(SHEETS.SUMMARY);
  const today   = _thaiDateTime(now);
  let outList = [], lowList = [], near50List = [], expList = [];
  // โหลด untracked set
  const untrackedRes = _getUntrackedExpiry();
  const untrackedSet = new Set(
    (untrackedRes.data || []).map(r => r.drugName + '||' + r.lotNo)
  );
  if (nolotSh && nolotSh.getLastRow() > 1) {
    const data    = nolotSh.getDataRange().getValues();
    const hdr     = data[0];
    const nameIdx = hdr.indexOf('ชื่อยาเต็ม');
    const stIdx   = hdr.indexOf('สถานะสต็อก');
    const balIdx  = hdr.indexOf('คงเหลือปัจจุบัน (Pack)');
    const minIdx  = hdr.indexOf('Minimum Stock');
    const cabIdx  = hdr.indexOf('ตู้ยา');
    for (let i = 1; i < data.length; i++) {
      const st  = String(data[i][stIdx]  || '');
      const nm  = String(data[i][nameIdx] || '');
      const bal = Number(data[i][balIdx] || 0);
      const min = Number(data[i][minIdx] || 0);
      const cab = String(data[i][cabIdx] || '');
      if (!nm) continue;
      if (st === 'หมดสต็อก')                       outList.push({name:nm, bal, min, cab});
      else if (st === 'ต่ำกว่า 50% ของ min stock') near50List.push({name:nm, bal, min, cab});
      else if (st === 'ต่ำกว่า minimum')            lowList.push({name:nm, bal, min, cab});
    }
  }
  if (sumSh && sumSh.getLastRow() > 1) {
    const cfg        = _getSettingsMap();
    const expiryDays = parseInt(cfg['EXPIRY_WARN_DAYS'] || '90');
    const data    = sumSh.getDataRange().getValues();
    const hdr     = data[0];
    const nameIdx = hdr.indexOf('ชื่อยาเต็ม');
    const lotIdx  = hdr.indexOf('Lot No');
    const daysIdx = hdr.indexOf('วันจนหมดอายุ');
    const expIdx  = hdr.indexOf('วันหมดอายุ');
    for (let i = 1; i < data.length; i++) {
      const days = Number(data[i][daysIdx] || 9999);
      // แสดงเฉพาะที่ คงเหลือ > 0 เท่านั้น (ถ้า = 0 แล้วไม่ต้องแสดง)
      if (days >= 0 && days <= expiryDays) {
        const nm  = String(data[i][nameIdx] || '');
        const lot = String(data[i][lotIdx] || '');
        const balIdx = hdr.indexOf('คงเหลือปัจจุบัน');
        const bal = balIdx >= 0 ? Number(data[i][balIdx] || 0) : 0;
        // แสดงเฉพาะ stock > 0
        if (bal <= 0) continue;
        // ข้าม untracked
        if (untrackedSet.has(nm + '||' + lot)) continue;
        const exp = data[i][expIdx] instanceof Date
          ? Utilities.formatDate(data[i][expIdx], 'Asia/Bangkok', 'dd/MM/yyyy')
          : String(data[i][expIdx] || '');
        expList.push({name: nm, lot, expDate: exp, days, bal});
      }
    }
  }
  // เขียน Dashboard sheet
  const rows = [];
  rows.push(['อัพเดทล่าสุด', today]);
  rows.push(['']);
  rows.push(['=== สรุปสถานะสต็อก ===']);
  rows.push(['หมดสต็อก',    outList.length]);
  rows.push(['ต่ำกว่า 50% Min', near50List.length]);
  rows.push(['ต่ำกว่า Min',  lowList.length]);
  rows.push(['']);
  if (outList.length > 0) {
    rows.push(['=== ❌ หมดสต็อก ===']);
    rows.push(['ชื่อยา', 'คงเหลือ', 'Min Stock', 'ตู้ยา']);
    outList.forEach(r => rows.push([r.name, r.bal, r.min, r.cab]));
    rows.push(['']);
  }
  if (near50List.length > 0) {
    rows.push(['=== 🔴 ต่ำกว่า 50% Min Stock ===']);
    rows.push(['ชื่อยา', 'คงเหลือ', 'Min Stock', 'ตู้ยา']);
    near50List.forEach(r => rows.push([r.name, r.bal, r.min, r.cab]));
    rows.push(['']);
  }
  if (lowList.length > 0) {
    rows.push(['=== ⚠️ ต่ำกว่า Min Stock ===']);
    rows.push(['ชื่อยา', 'คงเหลือ', 'Min Stock', 'ตู้ยา']);
    lowList.forEach(r => rows.push([r.name, r.bal, r.min, r.cab]));
    rows.push(['']);
  }
  if (expList.length > 0) {
    rows.push(['=== ⏰ ยาใกล้หมดอายุ ===']);
    rows.push(['ชื่อยา', 'Lot No', 'วันหมดอายุ', 'วันที่เหลือ']);
    expList.forEach(r => rows.push([r.name, r.lot||'', r.expDate, r.days]));
  }
  if (rows.length > 0) {
    const maxCols = Math.max(...rows.map(r => r.length), 1);
    sh.getRange(1, 1, rows.length, maxCols).setValues(
      rows.map(r => { while (r.length < maxCols) r.push(''); return r; })
    );
  }
  try {
    sh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1B3A6B').setFontColor('#FFFFFF');
  } catch(e) {}
  return { outList, near50List, lowList, expList };
}