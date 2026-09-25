// SAFE MODE: Backend แยกจาก Apps Script ระบบหลัก
// ใช้เฉพาะกับ Web App ของโปรเจกต์นี้
const SPREADSHEET_ID = '1kO0kBEW4Yy6J2Qb3spXP4pF_QupRrMIsYShGLmpdHts';
const TARGET_SHEET_NAME = 'Sandalwood_Web';

function doGet() {
  return json_({ok:true,service:'sandalwood-flower',mode:'safe'});
}

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    validate_(data);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบชีต Sandalwood_Web: เพื่อความปลอดภัย กรุณาสร้างแท็บนี้ก่อน โดยระบบจะไม่สร้าง/ลบ/ล้างแท็บอัตโนมัติ');

    const headers = ['ที่','หน่วยงาน','ประเภทการสนับสนุน','เป้าหมาย (ดอก)','อุปกรณ์ที่สนับสนุน','วัน/เดือน/ปี ดำเนินการ','วันที่คาดว่าจะส่งมอบ','ผู้ประสาน','หมายเลขโทรศัพท์'];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }

    const no = Math.max(1, sheet.getLastRow());
    sheet.appendRow([no,data.agency,data.supportType,data.target,data.equipment,data.startDate,data.deliveryDate,data.coordinator,data.phone]);
    return json_({ok:true,no:no});
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}

function validate_(d) {
  ['agency','supportType','startDate','deliveryDate','coordinator','phone'].forEach(k=>{
    if (!String(d[k] || '').trim()) throw new Error('ข้อมูลไม่ครบ: '+k);
  });
  if (!['สนับสนุนดอกไม้จันทน์','สนับสนุนอุปกรณ์'].includes(d.supportType)) throw new Error('ประเภทการสนับสนุนไม่ถูกต้อง');
  if (!/^[0-9]{9,10}$/.test(String(d.phone))) throw new Error('หมายเลขโทรศัพท์ไม่ถูกต้อง');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}