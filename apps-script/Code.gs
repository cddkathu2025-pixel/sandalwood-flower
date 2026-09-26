// SAFE MODE: Backend แยกจาก Apps Script ระบบหลัก
// โปรเจกต์ sandalwood-flower เท่านั้น
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
  'หมายเลขโทรศัพท์',
  'วันที่บันทึก'
];

function doGet(e) {
  try {
    // GET ธรรมดา = ตรวจ API
    if (!e || !e.parameter || e.parameter.action !== 'save') {
      return json_({
        ok: true,
        message: 'API ทำงานแล้ว',
        sheet: TARGET_SHEET_NAME,
        mode: 'GET-SAVE'
      });
    }

    // GET action=save = บันทึกข้อมูลโดยตรง
    const data = {
      agency: e.parameter.agency || '',
      supportType: e.parameter.supportType || '',
      target: e.parameter.target || '-',
      equipment: e.parameter.equipment || '-',
      startDate: e.parameter.startDate || '',
      deliveryDate: e.parameter.deliveryDate || '',
      coordinator: e.parameter.coordinator || '',
      phone: e.parameter.phone || ''
    };

    return save_(data);

  } catch (err) {
    return json_({
      ok: false,
      error: String(err.message || err)
    });
  }
}

// รองรับ POST เดิมไว้ด้วย เพื่อไม่ทำให้วิธีเดิมเสีย
function doPost(e) {
  try {
    let raw = e && e.postData && e.postData.contents;
    if ((!raw || !raw.trim()) && e && e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    }
    if (!raw) throw new Error('ไม่พบข้อมูล POST');

    return save_(JSON.parse(raw));
  } catch (err) {
    return json_({
      ok: false,
      error: String(err.message || err)
    });
  }
}

function save_(data) {
  validate_(data);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(TARGET_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(TARGET_SHEET_NAME);
    }

    // สร้าง/ปรับหัวตารางโดยไม่ลบข้อมูลเดิม
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      styleHeaders_(sheet);
    } else if (sheet.getLastColumn() < HEADERS.length) {
      const firstMissing = sheet.getLastColumn() + 1;
      const missing = HEADERS.slice(firstMissing - 1);
      sheet.getRange(1, firstMissing, 1, missing.length).setValues([missing]);
      sheet.getRange(1, firstMissing, 1, missing.length)
        .setBackground('#262626')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }

    // header อยู่แถว 1: รายการแรก = 1
    const no = Math.max(1, sheet.getLastRow());

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, HEADERS.length).setValues([[
      no,
      data.agency,
      data.supportType,
      data.target,
      data.equipment,
      data.startDate,
      data.deliveryDate,
      data.coordinator,
      String(data.phone),
      new Date()
    ]]);

    SpreadsheetApp.flush();

    return json_({
      ok: true,
      saved: true,
      no: no,
      sheet: TARGET_SHEET_NAME,
      message: 'บันทึกข้อมูลเรียบร้อยแล้ว'
    });

  } finally {
    lock.releaseLock();
  }
}

function styleHeaders_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setBackground('#262626')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
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
