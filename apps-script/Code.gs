// ============================================================
// ระบบบันทึกข้อมูลสนับสนุนการจัดทำดอกไม้จันทน์
// MAIN APPS SCRIPT + GitHub Pages API
// ============================================================

const SPREADSHEET_ID = '1kO0kBEW4Yy6J2Qb3spXP4pF_QupRrMIsYShGLmpdHts';
const MAIN_SHEET_NAME = 'Sandalwood_Web';

const EXPECTED_HEADERS = [
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

// ------------------------------------------------------------
// เดิม: เปิดหน้า index ใน Apps Script
// ใหม่: ถ้ามี ?action=save จะบันทึกจาก GitHub Pages
// ------------------------------------------------------------
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'save') {
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

      return saveToSandalwood_(data);
    }

    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบบันทึกข้อมูลสนับสนุนการจัดทำดอกไม้จันทน์')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);

  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: String(err.message || err)
    });
  }
}

// ------------------------------------------------------------
// รองรับ POST จากระบบเดิม/ระบบภายนอก
// ------------------------------------------------------------
function doPost(e) {
  try {
    let raw = '';

    if (e && e.postData && e.postData.contents) {
      raw = e.postData.contents;
    }

    if ((!raw || !raw.trim()) && e && e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    }

    if (!raw) {
      throw new Error('ไม่พบข้อมูลที่ส่งมา');
    }

    return saveToSandalwood_(JSON.parse(raw));

  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: String(err.message || err)
    });
  }
}

// ------------------------------------------------------------
// เดิม: เมนูใน Google Sheet
// ------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 กรอกข้อมูล')
    .addItem('เพิ่มรายการใหม่', 'showForm')
    .addToUi();
}

// ------------------------------------------------------------
// เดิม: เปิดฟอร์ม
// ------------------------------------------------------------
function showForm() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(450)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'ระบบบันทึกข้อมูลสนับสนุนการจัดทำดอกไม้จันทน์'
  );
}

// ------------------------------------------------------------
// เดิม: saveRow
// ปรับให้เขียน Sandalwood_Web โดยตรง
// ไม่ใช้ getActiveSheet()
// ไม่ clear ข้อมูล
// ------------------------------------------------------------
function saveRow(data) {
  return saveToSandalwood_(data);
}

// ------------------------------------------------------------
// ฟังก์ชันบันทึกกลาง ใช้ทั้ง saveRow / GET / POST
// ------------------------------------------------------------
function saveToSandalwood_(data) {

  validateData_(data);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName(MAIN_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(MAIN_SHEET_NAME);
    }

    ensureHeaders_(sheet);

    // แถวแรกเป็นหัวตาราง
    // ข้อมูลแถว 2 = เลข 1
    const no = Math.max(1, sheet.getLastRow());

    const row = [
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
    ];

    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      EXPECTED_HEADERS.length
    ).setValues([row]);

    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      saved: true,
      no: no,
      sheet: MAIN_SHEET_NAME,
      message: 'บันทึกข้อมูลเรียบร้อยแล้ว'
    });

  } catch (err) {

    return jsonResponse_({
      ok: false,
      saved: false,
      error: String(err.message || err)
    });

  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// ตรวจ/สร้างหัวตาราง โดยไม่ลบข้อมูลเดิม
// ------------------------------------------------------------
function ensureHeaders_(sheet) {

  if (sheet.getLastRow() === 0) {

    sheet
      .getRange(1, 1, 1, EXPECTED_HEADERS.length)
      .setValues([EXPECTED_HEADERS]);

    styleHeader_(sheet);
    return;
  }

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < EXPECTED_HEADERS.length) {

    const firstMissingColumn = lastColumn + 1;

    const missingHeaders =
      EXPECTED_HEADERS.slice(firstMissingColumn - 1);

    sheet
      .getRange(
        1,
        firstMissingColumn,
        1,
        missingHeaders.length
      )
      .setValues([missingHeaders]);

    sheet
      .getRange(
        1,
        firstMissingColumn,
        1,
        missingHeaders.length
      )
      .setBackground('#262626')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

  }

  sheet.setFrozenRows(1);
}

// ------------------------------------------------------------
// รูปแบบหัวตาราง
// ------------------------------------------------------------
function styleHeader_(sheet) {

  sheet
    .getRange(
      1,
      1,
      1,
      EXPECTED_HEADERS.length
    )
    .setBackground('#262626')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
}

// ------------------------------------------------------------
// Validation
// ------------------------------------------------------------
function validateData_(data) {

  const required = [
    'agency',
    'supportType',
    'startDate',
    'deliveryDate',
    'coordinator',
    'phone'
  ];

  required.forEach(function(key) {

    if (!String(data[key] || '').trim()) {
      throw new Error('ข้อมูลไม่ครบ: ' + key);
    }

  });

  if (
    data.supportType !== 'สนับสนุนดอกไม้จันทน์' &&
    data.supportType !== 'สนับสนุนอุปกรณ์'
  ) {
    throw new Error('ประเภทการสนับสนุนไม่ถูกต้อง');
  }

  if (data.supportType === 'สนับสนุนดอกไม้จันทน์') {

    if (
      data.target === undefined ||
      data.target === null ||
      Number(data.target) < 1
    ) {
      throw new Error('กรุณาระบุเป้าหมายดอกไม้จันทน์');
    }

  }

  if (data.supportType === 'สนับสนุนอุปกรณ์') {

    if (
      !String(data.equipment || '').trim() ||
      String(data.equipment).trim() === '-'
    ) {
      throw new Error('กรุณาระบุอุปกรณ์ที่สนับสนุน');
    }

  }

  if (!/^[0-9]{9,10}$/.test(String(data.phone))) {
    throw new Error('หมายเลขโทรศัพท์ไม่ถูกต้อง');
  }
}

// ------------------------------------------------------------
// ป้องกัน null / undefined
// ------------------------------------------------------------
function clean_(value) {
  return String(value === undefined || value === null ? '' : value).trim();
}

// ------------------------------------------------------------
// JSON response
// ------------------------------------------------------------
function jsonResponse_(obj) {

  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

}
