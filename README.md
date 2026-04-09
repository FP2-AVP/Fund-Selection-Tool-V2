# Fund Selection Tool For FP2

ระบบเลือกกองทุนสำหรับ Financial Planner รุ่นที่ 2
เชื่อมต่อกับ Google Sheets ขององค์กรผ่าน **Google OAuth 2.0**

---

## โครงสร้างไฟล์

```
Fund List Tool Project V2/
├── index.html          ← หน้าหลัก (entry point)
├── css/
│   └── style.css       ← stylesheet ทั้งหมด
├── js/
│   ├── config.js       ← ⚠️ ตั้งค่า CLIENT_ID และชื่อ Tab ที่นี่
│   ├── sheets.js       ← Google Sheets API handler
│   └── app.js          ← main application logic
├── fonts/
│   ├── THSarabunNew.woff
│   ├── THSarabunNew-Bold.woff
│   ├── THSarabunNew-Italic.woff
│   └── THSarabunNew-BoldItalic.woff
├── Data/
│   ├── AVP Master Fund ID.gsheet
│   ├── AVP Thai Fund for Quality.gsheet
│   ├── Percentrank Freestyle.gsheet
│   ├── Raw For Sec.gsheet
│   └── iShare Index Passive Return.gsheet
└── README.md           ← ไฟล์นี้
```

---

## การตั้งค่าก่อนใช้งาน (สำคัญ!)

### ขั้นตอนที่ 1 – สร้าง OAuth 2.0 Client ID

1. ไปที่ [Google Cloud Console](https://console.cloud.google.com/)
2. สร้าง Project ใหม่ หรือเลือก Project ที่มีอยู่
3. ไปที่ **APIs & Services → Library** → ค้นหา `Google Sheets API` → Enable
4. ไปที่ **APIs & Services → Credentials → + CREATE CREDENTIALS → OAuth client ID**
5. เลือก Application type: **Web application**
6. ตั้งชื่อ เช่น `Fund Selection Tool`
7. เพิ่ม **Authorized JavaScript origins**:
   - ระหว่าง Development: `http://localhost` หรือ `http://127.0.0.1`
   - Production: `https://yourdomain.com`
8. กด **Create** → Copy **Client ID** ที่ได้

> หมายเหตุ: ต้องตั้งค่า OAuth Consent Screen ก่อน หากยังไม่เคยทำ
> ไปที่ **OAuth consent screen** → เลือก Internal (สำหรับ Google Workspace) → กรอกข้อมูลให้ครบ

### ขั้นตอนที่ 2 – ใส่ Client ID ในไฟล์ config.js

เปิดไฟล์ `js/config.js` แล้วแก้ไขบรรทัดนี้:

```javascript
CLIENT_ID: 'YOUR_CLIENT_ID.apps.googleusercontent.com',
// เปลี่ยนเป็น Client ID จริง เช่น:
CLIENT_ID: '123456789-abcdefg.apps.googleusercontent.com',
```

### ขั้นตอนที่ 3 – ตรวจสอบชื่อ Tab ของ Google Sheets

เปิดแต่ละ Google Sheet แล้วดูชื่อ Tab ที่แถบล่างสุด
จากนั้นแก้ไขในไฟล์ `js/config.js` ส่วน `PAGES`:

```javascript
'thai-annualized': {
  tabName: 'Sheet1',   // ← เปลี่ยนเป็นชื่อ Tab จริง เช่น 'Annualized'
},
```

---

## วิธีเปิดใช้งาน

> **⚠️ ไม่สามารถเปิดโดยตรงจาก File Explorer ได้** (เนื่องจาก Google OAuth ต้องใช้ HTTP/HTTPS)

### วิธีที่ 1 – ใช้ VS Code Live Server (แนะนำสำหรับพัฒนา)
1. ติดตั้ง Extension "Live Server" ใน VS Code
2. คลิกขวาที่ `index.html` → **Open with Live Server**
3. เพิ่ม `http://127.0.0.1:5500` ใน Authorized JavaScript origins

### วิธีที่ 2 – Python Simple Server
```bash
cd "Fund List Tool Project V2"
python -m http.server 8080
```
แล้วเปิด `http://localhost:8080`

### วิธีที่ 3 – Deploy บน GitHub Pages หรือ Firebase Hosting
สำหรับการใช้งานจริงในองค์กร

#### GitHub Pages

โปรเจกต์นี้มี workflow สำหรับ deploy ขึ้น GitHub Pages แล้วที่
`.github/workflows/deploy-github-pages.yml`

ขั้นตอน:

1. สร้าง GitHub repository แล้ว push โค้ดขึ้น branch `main`
2. ไปที่ **Settings → Pages**
3. ตรง **Source** ให้เลือก **GitHub Actions**
4. push ขึ้น `main` อีกครั้ง หรือกดรัน workflow `Deploy static site to GitHub Pages`
5. เว็บจะได้ URL ประมาณ `https://<github-username>.github.io/<repo-name>/`

> สำคัญ: ต้องเพิ่ม origin `https://<github-username>.github.io` ใน Google Cloud Console ที่
> **APIs & Services → Credentials → OAuth 2.0 Client ID → Authorized JavaScript origins**
> มิฉะนั้น Google Sign-in จะใช้งานไม่ได้บน GitHub Pages

### โหมดแหล่งข้อมูล

ตั้งค่าได้ในไฟล์ `js/config.override.js`

- `google_first` — อ่านจาก Google Sheets ก่อน และ fallback เป็น JSON ถ้าอ่านไม่สำเร็จ
- `google_only` — อ่านจาก Google Sheets เท่านั้น
- `local_first` — อ่าน JSON ก่อน และ fallback เป็น Google Sheets
- `local_only` — อ่านจาก JSON เท่านั้น

ตอนนี้ค่า default ของโปรเจกต์ถูกตั้งเป็น `google_first`

---

## Google Sheets ที่ใช้ในระบบ

| ชื่อไฟล์ | Sheet ID | ใช้ใน |
|---------|----------|-------|
| AVP Master Fund ID | `1EYrIfINFg-WePEPr88PuDgsEW2OAZxkWtiq1Bml_1RA` | เลือกกองทุน, Master Fund Annualized/Calendar |
| AVP Thai Fund for Quality | `1UhD2XgEmRUZXasGTxH-E0TMl5dyuu2sRxXRVGzXTCzE` | กองทุนไทย Annualized/Calendar |
| Percentrank Freestyle | `1xt5YTBJoFKYc3gsLqhoCRsbrZHI8CSq5lRzBBqJwrdE` | (พร้อมขยายในอนาคต) |
| Raw For Sec | `1qNzpxP5D9RAaxhwAVe5Wf8rBhrGNhXn7EYQgBD86cBs` | (พร้อมขยายในอนาคต) |
| iShare Index Passive Return | `1WTuIL2UtRg6cacUWZo5nsbinyRWPbMuPLNXjAStliUY` | (พร้อมขยายในอนาคต) |

---

## ฟีเจอร์หลัก

| ฟีเจอร์ | รายละเอียด |
|--------|-----------|
| 🔒 Google OAuth 2.0 | เข้าถึงได้เฉพาะบัญชีที่มีสิทธิ์ใน Sheets ขององค์กร |
| 📊 แดชบอร์ด | สรุปจำนวนข้อมูลจากทุก Sheet + Quick Links |
| ☑ เลือกกองทุน | Checkbox selection + Search/Filter |
| ⚖ เปรียบเทียบกองทุน | เลือกหลายรายการแล้วเปรียบเทียบแบบ side-by-side |
| 🔍 ค้นหา/กรอง | Real-time search ทุกคอลัมน์ |
| ↕ Sort ตาราง | คลิกหัวตารางเพื่อเรียงลำดับ |
| ⬇ Export Excel | ดาวน์โหลด .xlsx ได้ทุกหน้า |
| ⚡ Cache | แคชข้อมูล 5 นาที เพื่อลด API calls |

---

## การแก้ไขปัญหาที่พบบ่อย

### "CLIENT_ID ยังไม่ได้ตั้งค่า"
→ เปิด `js/config.js` แล้วใส่ Client ID จริง

### "Token expired / Session หมดอายุ"
→ กดปุ่ม **เข้าสู่ระบบ** อีกครั้ง (Token มีอายุ 1 ชั่วโมง)

### "ไม่มีสิทธิ์เข้าถึง Sheet (403)"
→ ตรวจสอบว่าบัญชี Google ที่ใช้ Sign-in มีสิทธิ์อ่าน Google Sheet นั้น

### "ไม่พบ Sheet หรือ Tab (404)"
→ ตรวจสอบ Sheet ID และชื่อ Tab ใน `js/config.js`

### ข้อมูลไม่อัพเดต
→ กดปุ่ม **↻ รีเฟรช** เพื่อล้างแคชและดึงข้อมูลใหม่

---

## เทคโนโลยีที่ใช้

- **HTML5 / CSS3 / Vanilla JavaScript** — ไม่มี Framework
- **Google Identity Services** (GIS) — OAuth 2.0 Token-based auth
- **Google Sheets API v4** — ดึงข้อมูลผ่าน REST API
- **SheetJS (xlsx.js)** — Export ไฟล์ Excel
- **THSarabunNew** — ฟอนต์ภาษาไทย

---

*Fund Selection Tool For FP2 · Avenger Planner · v1.0*
