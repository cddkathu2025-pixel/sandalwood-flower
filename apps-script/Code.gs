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
    service: 'sandalwood-flower',
    mode: 'safe',
    sheet: TARGET_SHEET_NAME
  });
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw) throw new Error('ไม่พบข้อมูล POST');

    const data = JSON.parse(raw);
    validate_(data);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ปลอดภัย: แตะเฉพาะแท็บของระบบนี้เท่านั้น
    // ถ้ายังไม่มี จะสร้างเฉพาะ Sandalwood_Web
    let sheet = ss.getSheetByName(TARGET_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(TARGET_SHEET_NAME);
    }

    // ถ้าแท็บว่าง ให้สร้างหัวตาราง
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setBackground('#262626')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // ใช้ Lock ป้องกันเลขซ้ำเมื่อมีคนบันทึกพร้อมกัน
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
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
        message: 'บันทึกข้อมูลเรียบร้อยแล้ว'
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

  if (!/^[0-9]{9,10}$/.test(String(d.phone))) {
    throw new Error('หมายเลขโทรศัพท์ไม่ถูกต้อง');
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
