# ระบบแจ้งซ่อมออนไลน์ — คณะสัตวแพทยศาสตร์ ม.อ.

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/panithank-lab/vet-repair-system)

**Facility Repair Management System v3.0**

## 🔗 Live URL
> จะได้หลัง Deploy: `https://vet-repair-system.vercel.app`

## 📋 ฟีเจอร์
- แจ้งซ่อมผ่าน QR Code ประจำอาคาร
- วิศวกร/ช่างอนุมัติและมอบหมายงาน (1–3 คนพร้อมกัน)
- ติดตามสถานะ 7 ระดับ
- Dashboard สถิติ งบประมาณ Hotspot ภาระงาน
- ฐานข้อมูล Google Sheets

## 🏗️ อาคารที่รองรับ
- อาคารจุฬาภรณการุณยรักษ์
- อาคารสหเวชศาสตร์
- โรงพยาบาลปศุสัตว์
- อาคารวิจัยสัตวน้ำและสัตว์ปีก

## 🚀 วิธีตั้งค่า
1. Clone repo นี้
2. เปิด Google Sheet → Extensions → Apps Script → วาง `Code.gs`
3. Deploy as Web App → Copy URL
4. เปิดเว็บ → Login → ⚙ ตั้งค่า → วาง URL → กด เริ่มใช้งาน

## 👤 บัญชีผู้ใช้ (เปลี่ยนรหัสก่อนใช้งานจริง)
| Username | Role |
|----------|------|
| admin | วิศวกร (นายปณิธาน) |
| tech1 | ช่างเทคนิค (นายธนกร) |
| tech2 | ช่างเทคนิค (นายเอกรัตน์) |

## 🛠️ Tech Stack
- Frontend: HTML5 / CSS3 / Vanilla JavaScript
- Backend: Google Apps Script (REST API)
- Database: Google Sheets
- Hosting: Vercel

---
พัฒนาโดย งานโครงสร้างกายภาพและสารสนเทศ คณะสัตวแพทยศาสตร์ มหาวิทยาลัยสงขลานครินทร์
