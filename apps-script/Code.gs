// ============================================================
// ระบบบันทึกข้อมูลสนับสนุนการจัดทำดอกไม้จันทน์
// MAIN APPS SCRIPT + GitHub Pages API
// ============================================================

const SPREADSHEET_ID = '1kO0kBEW4Yy6J2Qb3spXP4pF_QupRrMIsYShGLmpdHts';
const MAIN_SHEET_NAME = 'Sandalwood_Web';

const EXPECTED_HEADERS = [
  'ที่','หน่วยงาน','ประเภทการสนับสนุน','เป้าหมาย (ดอก)',
  'อุปกรณ์ที่สนับสนุน','วัน/เดือน/ปี ดำเนินการ',
  'วันที่คาดว่าจะส่งมอบ','ผู้ประสาน','หมายเลขโทรศัพท์','วันที่บันทึก'
];

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'save') {
      const data = dataFromParams_(e.parameter);
      const result = saveToSandalwood_(data);
      const callback = e.parameter.callback;
      if (callback && /^[A-Za-z_$][A-Za-z0-9_$\.]*$/.test(callback)) {
        return ContentService
          .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return jsonResponse_(result);
    }

    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบบันทึกข้อมูลสนับสนุนการจัดทำดอกไม้จันทน์')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);

  } catch (err) {
    return jsonResponse_({ok:false,saved:false,error:String(err.message || err)});
  }
}

// รับ POST แบบ application/x-www-form-urlencoded จาก GitHub Pages
// แล้วตอบกลับเป็น HTML ใน iframe เพื่อส่งผลกลับหน้าเว็บด้วย postMessage
function doPost(e) {
  let result;
  try {
    const data = dataFromParams_((e && e.parameter) || {});
    result = saveToSandalwood_(data);
  } catch (err) {
    result = {ok:false,saved:false,error:String(err.message || err)};
  }
  return postMessageResponse_(result);
}

function dataFromParams_(p) {
  return {
    agency: p.agency || '',
    supportType: p.supportType || '',
    target: p.target || '-',
    equipment: p.equipment || '-',
    startDate: p.startDate || '',
    deliveryDate: p.deliveryDate || '',
    coordinator: p.coordinator || '',
    phone: p.phone || ''
  };
}

function postMessageResponse_(result) {
  const payload = JSON.stringify({type:'sheetSaveResult', result:result})
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');
  const html = '<!doctype html><html><body>' +
    '<script>window.parent.postMessage(' + payload + ', "*");</script>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 กรอกข้อมูล')
    .addItem('เพิ่มรายการใหม่','showForm')
    .addToUi();
}

function showForm() {
  const html=HtmlService.createHtmlOutputFromFile('index').setWidth(450).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html,'ระบบบันทึกข้อมูลสนับสนุนการจัดทำดอกไม้จันทน์');
}

function saveRow(data) {
  return saveToSandalwood_(data);
}

function saveToSandalwood_(data) {
  validateData_(data);

  const lock=LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet=ss.getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) sheet=ss.insertSheet(MAIN_SHEET_NAME);

    ensureHeaders_(sheet);

    const no=Math.max(1,sheet.getLastRow());

    sheet.getRange(sheet.getLastRow()+1,1,1,EXPECTED_HEADERS.length).setValues([[
      no,
      clean_(data.agency),
      clean_(data.supportType),
      data.target || '-',
      clean_(data.equipment || '-'),
      clean_(data.startDate),
      clean_(data.deliveryDate),
      clean_(data.coordinator),
      String(data.phone || ''),
      new Date()
    ]]);

    SpreadsheetApp.flush();

    return {
      ok:true,
      saved:true,
      no:no,
      sheet:MAIN_SHEET_NAME,
      message:'บันทึกข้อมูลเรียบร้อยแล้ว'
    };

  } catch (err) {
    return {
      ok:false,
      saved:false,
      error:String(err.message || err)
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow()===0) {
    sheet.getRange(1,1,1,EXPECTED_HEADERS.length).setValues([EXPECTED_HEADERS]);
    styleHeader_(sheet);
    return;
  }

  const lastColumn=sheet.getLastColumn();

  if (lastColumn<EXPECTED_HEADERS.length) {
    const firstMissing=lastColumn+1;
    const missing=EXPECTED_HEADERS.slice(firstMissing-1);
    sheet.getRange(1,firstMissing,1,missing.length).setValues([missing]);
    sheet.getRange(1,firstMissing,1,missing.length)
      .setBackground('#262626')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }

  sheet.setFrozenRows(1);
}

function styleHeader_(sheet) {
  sheet.getRange(1,1,1,EXPECTED_HEADERS.length)
    .setBackground('#262626')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function validateData_(data) {
  ['agency','supportType','startDate','deliveryDate','coordinator','phone'].forEach(function(key){
    if (!String(data[key] || '').trim()) throw new Error('ข้อมูลไม่ครบ: '+key);
  });

  if (data.supportType!=='สนับสนุนดอกไม้จันทน์' &&
      data.supportType!=='สนับสนุนอุปกรณ์') {
    throw new Error('ประเภทการสนับสนุนไม่ถูกต้อง');
  }

  if (data.supportType==='สนับสนุนดอกไม้จันทน์' &&
      (!data.target || Number(data.target)<1)) {
    throw new Error('กรุณาระบุเป้าหมายดอกไม้จันทน์');
  }

  if (data.supportType==='สนับสนุนอุปกรณ์' &&
      (!String(data.equipment || '').trim() || String(data.equipment).trim()==='-')) {
    throw new Error('กรุณาระบุอุปกรณ์ที่สนับสนุน');
  }

  if (!/^[0-9]{9,10}$/.test(String(data.phone))) {
    throw new Error('หมายเลขโทรศัพท์ไม่ถูกต้อง');
  }
}

function clean_(value) {
  return String(value===undefined || value===null ? '' : value).trim();
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
