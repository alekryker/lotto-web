# หน้าแทง + หลังบ้าน

- `index.html` — หน้าแทง (เลือกหวย → แทง → ส่งโพย)
- `admin.html` — หลังบ้าน: ยอดซื้อแต่ละหวย และเปิดดูแต่ละโพยว่าซื้ออะไร
- `config.js` — ใส่ URL ของ Apps Script ที่เดียว
- `apps-script/Code.gs` — โค้ดหลังบ้าน (Google Apps Script + Google Sheet)

## ติดตั้งหลังบ้าน (Google Sheet ใหม่)

1. สร้าง Google Sheet ใหม่ → เมนู **ส่วนขยาย > Apps Script**
2. ลบโค้ดเดิม วางโค้ดจาก `apps-script/Code.gs` ทั้งไฟล์ แล้วกดบันทึก
3. เลือกฟังก์ชัน `setup` ด้านบน → กด **Run** (ครั้งแรกจะขออนุญาต กดยอมรับ)
   จะได้ชีต `หวย` `ตั้งค่า` `โพย`
4. **Project Settings (รูปเฟือง) > Script properties > Add** → ชื่อ `ADMIN_KEY` ค่า = รหัสแอดมิน
5. **Deploy > New deployment > Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   กด Deploy แล้วคัดลอก URL ที่ลงท้าย `/exec`
6. เปิด `config.js` ใส่ URL ใน `API_URL`

## ใช้งานประจำวัน (แก้ในชีต ไม่ต้อง deploy ใหม่)

- ชีต **หวย**: ชื่อหวย / เวลาปิด (เช่น `18:10`) / ติ๊ก "เปิดรับ"
  หวยที่ปิด 00:00–05:59 จะนับเป็นงวดของวันก่อนหน้า (ปิดหลังเที่ยงคืน)
- ชีต **ตั้งค่า**: ธนาคาร / เลขบัญชี / ชื่อบัญชี (แสดงหลังส่งโพย)
- ชีต **โพย**: ทุกโพยที่ส่งเข้ามา (คอลัมน์ "รายการ" เก็บรายละเอียดเป็น JSON)

ถ้าแก้ `Code.gs` ภายหลัง ต้อง **Deploy > Manage deployments > แก้ไข > New version** (URL เดิม)

## เปิดบนเว็บ (GitHub Pages)

Settings > Pages > Source: **Deploy from a branch** → `main` / `root`
- หน้าแทง: `https://<user>.github.io/<repo>/`
- หลังบ้าน: `https://<user>.github.io/<repo>/admin.html`

หน้าหลังบ้านเปิดได้ทุกคน แต่ข้อมูลโหลดได้เฉพาะคนที่มีรหัส `ADMIN_KEY`
