# แดชบอร์ดคัดกรองสุขภาพช่องปากนักเรียน (Oral-Health Screening Dashboard)

โปรแกรมเดสก์ท็อป (`.exe`) สำหรับรันในเครื่องผู้ใช้ อ่านข้อมูลจากไฟล์ **`.xlsm`/`.xlsx`** ในโฟลเดอร์เดียวกัน แล้วสร้างแดชบอร์ดวิเคราะห์สุขภาพช่องปาก/ความเสี่ยงฟันผุ เปิดในเบราว์เซอร์ พร้อมระบบแจ้งเตือนเวอร์ชันใหม่ผ่าน GitHub

---

## สำหรับผู้ใช้ (ดาวน์โหลดไปใช้งาน)

1. เข้าหน้า **Releases** ของ repo นี้ → ดาวน์โหลด **`OralHealthDashboard.exe`** เวอร์ชันล่าสุด
2. วางไฟล์ `.exe` ไว้ใน **โฟลเดอร์เดียวกับไฟล์ข้อมูล** (ไฟล์ `.xlsm` หรือ `.xlsx` ที่มีชีตชื่อ `Records`)
3. ดับเบิลคลิก `OralHealthDashboard.exe`
4. โปรแกรมจะอ่านข้อมูล สร้างหน้าแดชบอร์ด แล้วเปิดในเบราว์เซอร์อัตโนมัติ (ไฟล์ `oral_health_dashboard.html` จะถูกสร้างไว้ในโฟลเดอร์ เปิดซ้ำได้)

> **อัปเดตเวอร์ชัน:** ทุกครั้งที่เปิดโปรแกรม จะตรวจสอบเวอร์ชันล่าสุดจาก GitHub ให้อัตโนมัติ ถ้ามีของใหม่จะขึ้นแถบสีน้ำเงินด้านบนของแดชบอร์ดพร้อมลิงก์ดาวน์โหลด — เพียงโหลด `.exe` ใหม่มาวางทับตัวเดิม

### รูปแบบไฟล์ข้อมูลที่ต้องมี
- ต้องมีชีตชื่อ **`Records`** โครงหัวตาราง **3 แถว**: แถวที่ 1 = หมวดหมู่ · แถวที่ 2 = ชื่อฟิลด์ไทย · แถวที่ 3 = โค้ดฟิลด์ (เช่น `sex`, `d`, `t16`, `riskgroup`) · แถวที่ 4 เป็นต้นไป = ข้อมูลนักเรียน
- ถ้ามีชีต **`Codebook`** (คอลัมน์ `column_name … alert`) โปรแกรมจะใช้คอลัมน์ `alert` มากำหนดค่าที่ "เฝ้าระวัง" ให้อัตโนมัติ (ถ้าไม่มีก็ใช้ค่าเริ่มต้นในโปรแกรม)
- คำอธิบายความหมายของค่าตัวเลือก (เช่น `sweet_drink_freq` 0–3) ถูกฝังไว้ในโปรแกรมแล้ว

---

## สำหรับผู้ดูแล (ตั้งค่า repo + ออกเวอร์ชัน)

### 1) สร้าง repo และ push โค้ด
```bash
git init
git add .
git commit -m "initial: oral-health dashboard app"
git branch -M main
git remote add origin https://github.com/<username>/oral-health-dashboard.git
git push -u origin main
```
> ไม่ต้อง commit ไฟล์ข้อมูลนักเรียนจริงขึ้น GitHub สาธารณะ (มีใน `.gitignore` ให้แล้ว)

### 2) ออกเวอร์ชันใหม่ → บิลด์ `.exe` อัตโนมัติ
```bash
git tag v1.0.0
git push origin v1.0.0
```
เมื่อ push tag ที่ขึ้นต้นด้วย `v` → GitHub Actions (`.github/workflows/release.yml`) จะ:
1. บิลด์ `.exe` บนเครื่อง **Windows** ด้วย PyInstaller
2. ฝังเลขเวอร์ชัน (จาก tag) และชื่อ repo ลงในโปรแกรม (ใช้สำหรับระบบตรวจอัปเดต)
3. สร้าง **Release** และแนบ `OralHealthDashboard.exe` ให้อัตโนมัติ

ออกเวอร์ชันถัดไปก็แค่เพิ่ม tag ใหม่ เช่น `v1.1.0`, `v1.2.0` …

> กดบิลด์เองได้จากแท็บ **Actions → Build & Release EXE → Run workflow** ก็ได้

### 3) รัน/พัฒนาในเครื่อง (ไม่ต้องบิลด์ exe)
```bash
pip install -r requirements.txt
python app.py      # วางไฟล์ .xlsm ไว้โฟลเดอร์เดียวกันก่อน
```

### บิลด์ `.exe` เองบน Windows (ถ้าต้องการ)
```powershell
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --name OralHealthDashboard `
  --add-data "template.html;." `
  --add-data "logo_datauri.txt;." app.py
# ได้ไฟล์ที่ dist\OralHealthDashboard.exe
```

---

## โครงสร้างโปรเจกต์
```
oral-health-dashboard/
├─ app.py                     # โปรแกรมหลัก (อ่าน xlsm → สร้าง HTML → เปิดเบราว์เซอร์ → เช็คอัปเดต)
├─ template.html              # เทมเพลตแดชบอร์ด (กราฟ/ตัวกรอง/ตารางรายชื่อ/ส่งออก CSV)
├─ logo_datauri.txt           # โลโก้ (ฝังเป็น data URI)
├─ requirements.txt
├─ .github/workflows/release.yml   # CI: บิลด์ exe + ออก Release
├─ .gitignore
└─ README.md
```

## หมายเหตุเรื่อง Antivirus
ไฟล์ `.exe` ที่สร้างจาก PyInstaller บางครั้งถูก SmartScreen/แอนติไวรัสเตือน (false positive) เพราะยังไม่มี code-signing certificate หากต้องการลดปัญหานี้ควรทำ code signing เพิ่มเติม (มีค่าใช้จ่ายใบรับรอง)
