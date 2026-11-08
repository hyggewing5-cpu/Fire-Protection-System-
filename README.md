# ระบบตรวจป้องกันอัคคีภัยก่อนปิดสำนักงาน

- `index.html` — หน้าเว็บ (ใช้ได้ทั้งมือถือ/iPad/คอมพิวเตอร์) โฮสต์บน GitHub Pages ได้
- `Code.js` — โค้ดฝั่งเซิร์ฟเวอร์ สำหรับวางใน Google Apps Script (ไม่ทำงานบน GitHub Pages)

## วิธีใช้
1. วาง `Code.js` ใน Apps Script ของ Google Sheet แล้ว Deploy เป็น Web app (Execute as: Me, Access: Anyone)
2. ใส่ URL Web app (ลงท้าย `/exec`) ที่ `var API_URL = '';` ใน `index.html`
3. เปิด GitHub Pages: Settings → Pages → Branch main / root

รหัสผ่านผู้ดูแลตั้งที่ชีต Config เซลล์ A3
