# Fund Selection Tool For FP2

ระบบเลือกและวิเคราะห์กองทุนสำหรับ Financial Planner รุ่นที่ 2  
เชื่อมต่อกับ Google Sheets ขององค์กรผ่าน **Google OAuth 2.0**

---

## เมนูทั้งหมด

| เมนู | หน้า | แหล่งข้อมูล |
|------|------|------------|
| คู่มือการใช้งาน | คู่มือ README | — |
| แดชบอร์ด | ภาพรวมสถิติ | ทุก Sheet |
| เลือกกองทุน | คัดเลือก + Highlight | Percentrank Freestyle |
| กองทุนไทย Annualized Return | ผลตอบแทนทุก Period | AVP Thai Fund for Quality |
| กองทุนไทย Calendar Year | ผลตอบแทนรายปี | AVP Thai Fund for Quality |
| Master Fund Annualized Return | Master Fund ทุก Period | AVP Master Fund ID |
| Master Fund Calendar Year | Master Fund รายปี | AVP Master Fund ID |
| ค่าธรรมเนียม | TER เทียบ Master vs กองไทย | Raw For Sec + AVP Master Fund ID |
| ปัจจัยประกอบอื่นๆ | Scatter + ตาราง Sharpe/Sortino/IR/Treynor/Max DD | AVP Master Fund ID |
| Top 10 Holding | การถือครองสูงสุด 10 อันดับ | AVP Master Fund ID |
| Avenger Studio | Presentation Designer | — |

---

## โครงสร้างไฟล์

```
Fund List Tool Project V2/
├── index.html              ← หน้าหลัก (entry point)
├── css/
│   ├── style.css           ← stylesheet หลัก
│   └── other-factors.css   ← override สัดส่วนหน้า ปัจจัยประกอบอื่นๆ
├── js/
│   ├── config.js           ← ⚠️ ตั้งค่า CLIENT_ID และ Sheet IDs ที่นี่
│   ├── config.override.js  ← override แหล่งข้อมูลและ paths ของ JSON
│   ├── sheets.js           ← Google Sheets API handler
│   ├── app.js              ← main application logic (ทุกหน้า)
│   └── avenger-studio.js   ← Presentation Designer
├── fonts/
│   ├── THSarabunNew.woff
│   ├── THSarabunNew-Bold.woff
│   ├── THSarabunNew-Italic.woff
│   └── THSarabunNew-BoldItalic.woff
├── Data/
│   ├── AVP Master Fund ID - 2026-Q1.json      ← ข้อมูล Master Fund
│   ├── AVP Thai Fund for Quality - 2026-Q1.json ← ข้อมูลกองทุนไทย
│   ├── Percentrank Freestyle - 2026-Q1.json   ← ข้อมูลเลือกกองทุน
│   ├── Raw For Sec - 2026-Q1.json             ← ข้อมูลค่าธรรมเนียม
│   └── iShare Index Passive Return - 2026-Q1.json
└── README.md
```

---

## การตั้งค่าก่อนใช้งาน (สำคัญ!)

### ขั้นตอนที่ 1 – สร้าง OAuth 2.0 Client ID

1. ไปที่ [Google Cloud Console](https://console.cloud.google.com/)
2. สร้าง Project ใหม่ หรือเลือก Project ที่มีอยู่
3. ไปที่ **APIs & Services → Library** → ค้นหา `Google Sheets API` → Enable
4. ไปที่ **APIs & Services → Credentials → + CREATE CREDENTIALS → OAuth client ID**
5. เลือก Application type: **Web application**
6. เพิ่ม **Authorized JavaScript origins**:
   - Development: `http://localhost:8080`
   - Production: `https://<github-username>.github.io`
7. กด **Create** → Copy **Client ID**

### ขั้นตอนที่ 2 – ตั้งค่าใน config.js

```javascript
CLIENT_ID: 'YOUR_CLIENT_ID.apps.googleusercontent.com',
```

### ขั้นตอนที่ 3 – อัปเดต Quarter ใหม่

เปิด `js/config.override.js` แก้ path ของไฟล์ JSON:

```javascript
THAI:   'Data/AVP Thai Fund for Quality - 2026-Q2.json',
MASTER: 'Data/AVP Master Fund ID - 2026-Q2.json',
```

---

## วิธีเปิดใช้งาน

> **⚠️ ต้องเปิดผ่าน HTTP Server เท่านั้น** (Google OAuth ไม่ทำงานบน `file://`)

### Mac

```bash
cd "Fund List Tool Project V2"
python3 -m http.server 8080
```
แล้วเปิด `http://localhost:8080`

### Windows

ดับเบิลคลิก `start-server.bat`

### GitHub Pages

push ขึ้น `main` → workflow `Deploy static site to GitHub Pages` จะรันอัตโนมัติ

---

## Google Sheets ที่ใช้ในระบบ

| Sheet | Sheet ID | ใช้ใน |
|-------|----------|-------|
| AVP Master Fund ID | `10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig` | Master Fund Annualized/Calendar, ค่าธรรมเนียม, ปัจจัยประกอบอื่นๆ, Top 10 Holding |
| AVP Thai Fund for Quality | `1m1rSyJAel9atGMrmeRSwgYWa9wgc4gi7-3cp4Yvc8GM` | กองทุนไทย Annualized/Calendar |
| Percentrank Freestyle | `1s-0ciSOB2Tj0C9azeMXyd1zZxljOg8I5QilI0FgjdW4` | เลือกกองทุน |
| Raw For Sec | `16agx9pl9adtMh-U7MCbgnIncBxpciCvFgsdurH6Ob8w` | ค่าธรรมเนียม |

### โหมดแหล่งข้อมูล (`config.override.js`)

| โหมด | พฤติกรรม |
|------|---------|
| `google_first` | อ่าน Google Sheets ก่อน → fallback เป็น JSON (default) |
| `local_first` | อ่าน JSON ก่อน → fallback เป็น Google Sheets |
| `local_only` | อ่าน JSON เท่านั้น (เหมาะสำหรับ offline) |
| `google_only` | อ่าน Google Sheets เท่านั้น |

---

## ฟีเจอร์หน้า ปัจจัยประกอบอื่นๆ

| ฟีเจอร์ | รายละเอียด |
|--------|-----------|
| Scatter Chart | เลือก Metric แกน X / Y ได้อิสระจาก 12 ตัวชี้วัด |
| Toggle Annualized/Calendar | สลับระหว่าง YTD/1Y/3Y/5Y/10Y กับ ปี 2016–2025 |
| ตาราง Side Panel | แสดงค่า X, Y ของทุกกองที่ plot บนกราฟ |
| Checkbox | เลือก/ซ่อนกองทุนบนกราฟแบบ real-time |
| State Persist | จำ Period / Axis / Checkbox ไว้เมื่อเปลี่ยนเมนูแล้วกลับมา |
| ตัวชี้วัด | Return, Sharpe Ratio (3 แบบ), Information Ratio (2 แบบ), Sortino Ratio (3 แบบ), Treynor Ratio (2 แบบ), Max Drawdown |

> หมายเหตุ: Calendar mode ไม่มีข้อมูล Sharpe Ratio รายปี ระบบจะซ่อนให้อัตโนมัติ

---

## Presentation Table Presets

กำหนดใน `PRESENTATION_TABLE_PRESETS` ใน `js/app.js`

| Preset | หน้า | rowsPerSlide |
|--------|------|-------------|
| `thaiAnnualizedV2` | กองทุนไทย Annualized Return | 30 |
| `thaiCalendar` | กองทุนไทย Calendar Year | 30 |
| `masterAnnualizedV2` | Master Fund Annualized Return | 30 |
| `masterCalendar` | Master Fund Calendar Year | 30 |
| `masterFees2` | ค่าธรรมเนียม | 30 |
| `default` | fallback | 30 |

---

## วิธีเพิ่ม Quarter ใหม่

ระบบตรวจ Tab จาก Google Sheet อัตโนมัติหลัง Login

1. เพิ่ม Tab ใหม่ใน Google Sheet (รูปแบบ `YYYY-Q1` ถึง `YYYY-Q4`)
2. วางไฟล์ JSON ใหม่ใน `Data/`
3. แก้ 2 บรรทัดใน `js/config.override.js`
4. Refresh — Quarter ใหม่ปรากฏใน Sidebar ทันที

---

## การแก้ไขปัญหาที่พบบ่อย

| ปัญหา | วิธีแก้ |
|-------|--------|
| "CLIENT_ID ยังไม่ได้ตั้งค่า" | เปิด `js/config.js` ใส่ Client ID จริง |
| "Token expired" | กดปุ่ม **เข้าสู่ระบบ** อีกครั้ง |
| "ไม่มีสิทธิ์ (403)" | ตรวจสอบว่าบัญชีมีสิทธิ์อ่าน Google Sheet |
| ข้อมูลไม่อัพเดต | กดปุ่ม **↻ รีเฟรช** เพื่อล้าง cache |
| Quarter Selector ไม่แสดง Tab | ชื่อ Tab ต้องเป็น `YYYY-Q1` เท่านั้น |
| หน้า ปัจจัยประกอบอื่นๆ ไม่มีข้อมูล | ต้องเลือกกองทุนจากเมนู "เลือกกองทุน" ก่อน และกองต้องผูก Master Fund ID ไว้ |

---

## เทคโนโลยีที่ใช้

- **HTML5 / CSS3 / Vanilla JavaScript** — ไม่มี Framework
- **Google Identity Services** (GIS) — OAuth 2.0
- **Google Sheets API v4** — REST API
- **SheetJS (xlsx.js)** — Export Excel
- **Fabric.js + jsPDF** — Presentation / PDF rendering
- **THSarabunNew** — ฟอนต์ภาษาไทย

---

*Fund Selection Tool For FP2 · Avenger Planner · v2.1 · อัปเดต เมษายน 2569*
