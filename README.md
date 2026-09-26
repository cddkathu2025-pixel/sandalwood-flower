# ระบบบันทึกข้อมูลสนับสนุนการจัดทำดอกไม้จันทน์

เว็บไซต์สำหรับบันทึกข้อมูลการสนับสนุนการจัดทำดอกไม้จันทน์ อำเภอกะทู้ จังหวัดภูเก็ต

## โครงสร้างระบบปัจจุบัน

- Frontend: GitHub Pages
- Database/API: Supabase REST Data API
- ตาราง: `public.sandalwood_web`
- ผู้กรอกข้อมูลไม่ต้อง Login
- ไม่ใช้ Google Apps Script
- ไม่ใช้ Google OAuth
- ไม่ใช้ Google Sheets
- ฝั่ง Browser ใช้เฉพาะ Supabase Publishable Key
- ห้ามนำ Secret/Service Role Key มาใส่ใน `index.html`

## ความปลอดภัย

ตารางเปิด RLS และให้ `anon` ทำได้เฉพาะ INSERT เท่านั้น ส่วน SELECT/UPDATE/DELETE ถูกปิดสำหรับผู้ใช้สาธารณะ

ข้อมูลที่ส่งจะถูกตรวจสอบซ้ำด้วย RLS ได้แก่ ประเภทการสนับสนุน จำนวน/อุปกรณ์ ชื่อหน่วยงาน ผู้ประสาน และรูปแบบโทรศัพท์

## หน้าใช้งาน

GitHub Pages:
https://cddkathu2025-pixel.github.io/sandalwood-flower/

ระบบเวอร์ชัน: 2026.09.26.2
