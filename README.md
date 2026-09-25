# sandalwood-flower

ระบบหน้าเว็บสำหรับบันทึกข้อมูลสนับสนุนการจัดทำดอกไม้จันทน์ อำเภอกะทู้ จังหวัดภูเก็ต

## Safe Mode
- Repository นี้แยกจากระบบหลัก
- หน้าเว็บไม่ลบหรือแก้ข้อมูล Google Sheet โดยตรง
- Backend จะต้องเป็น Google Apps Script Web App ที่แยกจาก Apps Script เดิม
- ห้ามใช้ `sheet.clear()` กับข้อมูลระบบหลัก
- `index.html` จะยังไม่ส่งข้อมูลจนกว่าจะกำหนด `APPS_SCRIPT_URL`

## การเชื่อมต่อ
เปิด `index.html` แล้วกำหนด URL ของ Google Apps Script Web App ที่สร้างแยกสำหรับระบบนี้ในตัวแปร `APPS_SCRIPT_URL`

Google Sheet เดิม: 1kO0kBEW4Yy6J2Qb3spXP4pF_QupRrMIsYShGLmpdHts
