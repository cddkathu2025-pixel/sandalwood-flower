// SAFE MODE: Backend แยกจาก Apps Script ระบบหลัก
// ใช้เฉพาะกับ Web App ของโปรเจกต์นี้
const SPREADSHEET_ID = '1kO0kBEW4Yy6J2Qb3spXP4pF_QupRrMIsYShGLmpdHts';
const TARGET_SHEET_NAME = 'Sandalwood_Web';

const HEADERS = [
  'ที่',
  'หน่วยงาน',
  'ประเภทการสนับสนุน',
  'เป้าหมาย (ดอก)',
  'อุปกรณ์ที่สนับสนุน',
  'วัน/เดือน/ปี ดำเนินการ',
  'วันที่คาดว่าจะส่งมอบ',
  'ผู้ประสาน',
  'หมายเลขโทรศัพท์'
];

function doGet() {
  return json_({
    ok: true,
    message: 'API ทำงานแล้ว',
    sheet: TARGET_SHEET_NAME
  });
}

function doPost(e) {
  try {
    // รองรับทั้ง JSON POST และ form POST จาก GitHub Pages
    let raw = '';
    if (e && e.postData && e.postData.contents) {
      raw = e.postData.contents;
    }

    // เมื่อส่งผ่าน HTML form จะเข้าทาง e.parameter.payload
    if ((!raw || raw.trim() === '') && e && e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    }

    if (!raw) throw new Error('ไม่พบข้อมูล POST');

    const data = JSON.parse(raw);
    validate_(data);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(TARGET_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(TARGET_SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setBackground('#262626')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      // แถวที่ 1 คือหัวตาราง ดังนั้นรายการแรกต้องเป็น 1
      const no = Math.max(1, sheet.getLastRow());

      sheet.appendRow([
        no,
        data.agency,
        data.supportType,
        data.target,
        data.equipment,
        data.startDate,
        data.deliveryDate,
        data.coordinator,
        data.phone
      ]);

      SpreadsheetApp.flush();

      return json_({
        ok: true,
        no: no,
        message: 'บันทึกข้อมูลเรียบร้อยแล้ว',
        sheet: TARGET_SHEET_NAME
      });
    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    return json_({
      ok: false,
      error: String(err.message || err)
    });
  }
}

function validate_(d) {
  ['agency', 'supportType', 'startDate', 'deliveryDate', 'coordinator', 'phone']
    .forEach(function(k) {
      if (!String(d[k] || '').trim()) {
        throw new Error('ข้อมูลไม่ครบ: ' + k);
      }
    });

  if (
    d.supportType !== 'สนับสนุนดอกไม้จันทน์' &&
    d.supportType !== 'สนับสนุนอุปกรณ์'
  ) {
    throw new Error('ประเภทการสนับสนุนไม่ถูกต้อง');
  }

  if (d.supportType === 'สนับสนุนดอกไม้จันทน์') {
    if (!d.target || Number(d.target) < 1) {
      throw new Error('กรุณาระบุเป้าหมายดอกไม้จันทน์');
    }
  }

  if (d.supportType === 'สนับสนุนอุปกรณ์') {
    if (!String(d.equipment || '').trim() || d.equipment === '-') {
      throw new Error('กรุณาระบุอุปกรณ์ที่สนับสนุน');
    }
  }

  if (!/^[0-9]{9,10}$/.test(String(d.phone))) {
    throw new Error('หมายเลขโทรศัพท์ไม่ถูกต้อง');
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
