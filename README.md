# 🚀 Fund Selection Tool For FP2
**ระบบเลือกและวิเคราะห์กองทุนสำหรับ Financial Planner รุ่นที่ 2**
*เชื่อมต่อกับ Google Sheets ขององค์กรผ่าน Google OAuth 2.0 และระบบสำรองข้อมูล Local JSON*

---

## 🎨 Avenger Studio & Analysis
ส่วนงานวิเคราะห์เชิงลึกและการจัดการข้อมูลเพื่อการนำเสนอ:

### 🏆 Top 10 Holding (การวิเคราะห์การถือครอง)
* **การเปรียบเทียบ:** รองรับการเปรียบเทียบสูงสุด 8 ช่อง (รวมข้อมูลจาก iShare Index Passive Return)
* **⚠️ หมายเหตุสำคัญ (Performance):** เนื่องจากข้อมูลมีปริมาณมหาศาล **ผู้ใช้จำเป็นต้องเลือกรายชื่อกองทุนที่ต้องการเปรียบเทียบก่อนทุกครั้ง** ระบบไม่สามารถโหลดอัตโนมัติได้ทันทีเพื่อป้องกันหน้าเพจโหลดช้า

### 📄 Planner Designer V7 (สถานะระบบ)
* **ระบบส่งข้อมูลไปทำ Presentation:** ⚠️ **[Status: In Development]** ขณะนี้ฟังก์ชันการส่งข้อมูลอัตโนมัติไปยังหน้า Presentation ยังไม่สมบูรณ์ 100% 
* **ข้อแนะนำ:** ในระหว่างการพัฒนาเวอร์ชันสมบูรณ์ ผู้ใช้งานอาจต้องดำเนินการตรวจสอบหรือปรับแต่งข้อมูลเพิ่มเติมก่อนการส่งออกเป็นไฟล์ PDF

---

## 📑 รายละเอียดเมนูการใช้งาน (System Menu)
| เมนู | รายละเอียดข้อมูล | แหล่งข้อมูล (Google Sheet ID) |
| :--- | :--- | :--- |
| **🔍 เลือกกองทุน** | คัดเลือก / Highlight / Percentrank | `1s-0ciSOB2Tj0C9azeMXyd1zZxljOg8I5QilI0FgjdW4` |
| **📈 กองทุนไทย** | Annualized & Calendar Year (AVP Thai Fund) | `1m1rSyJAel9atGMrmeRSwgYWa9wgc4gi7-3cp4Yvc8GM` |
| **🌍 Master Fund** | Annualized & Calendar Year (AVP Master ID) | `1XpU2fK6_tqYpM7o5o0eXz_m_8fTSE_ixVl13rWBig` |
| **💰 ค่าธรรมเนียม** | TER Comparison (Raw For Sec) | `16agx9pl9adtMh-U7MCbgnIncBxpciCvFgsdurH6Ob8w` |
| **📊 Top 10** | **ต้องเลือกกองทุนก่อนโหลดข้อมูล** | `1miHQVkwEq7k4S0upoYsRg8Z0KsbZirbwtn7BKCk_Edw` |

---

## 💾 ระบบบันทึกข้อมูล (Upcoming)
* **ระบบ Draft งาน:** บันทึกรายชื่อกองทุนที่เลือกไว้ในแต่ละ Session ไม่ต้องเริ่มต้นใหม่
* **Resume Session:** เรียกคืนข้อมูลที่เคยวิเคราะห์ค้างไว้กลับมาทำงานต่อได้ทันที
* **Cloud Storage:** พื้นที่จัดเก็บข้อมูลออนไลน์เพื่อความสะดวกในการเข้าถึงจากหลายอุปกรณ์

---

## 📁 โครงสร้างโปรเจกต์ (Project Structure)
* `index.html` : หน้าหลัก (Entry Point)
* `js/config.js` : ⚠️ **จุดตั้งค่า CLIENT_ID และ Sheet IDs หลัก**
* `js/config.override.js` : ตั้งค่าแหล่งข้อมูล (Sheets/JSON) และ Paths
* `js/avenger-studio.js` : ระบบควบคุมการสร้าง PDF และตัวจัดการ Presentation
* `Data/` : แหล่งเก็บไฟล์สำรองข้อมูล JSON (เช่น `2026-Q1.json`)

---

## 🛠 วิธีเพิ่ม Quarter ใหม่
1. ระบบจะตรวจหา Tab จาก Google Sheets โดยอัตโนมัติหลัง Login
2. ให้เพิ่ม Tab ใหม่ใน Google Sheet โดยใช้ชื่อรูปแบบ `YYYY-Q1` ถึง `YYYY-Q4`
3. หากใช้งานแบบ Offline ให้วางไฟล์ JSON ใหม่ลงในโฟลเดอร์ `Data/`
4. แก้ไข Path ข้อมูลในไฟล์ `js/config.override.js` และ Refresh หน้าเว็บ

---

## ❓ การแก้ไขปัญหา (Troubleshooting)
* **"CLIENT_ID ยังไม่ได้ตั้งค่า":** ตรวจสอบและระบุค่าในไฟล์ `js/config.js`
* **"Token expired":** กดปุ่ม "เข้าสู่ระบบ" ใหม่ที่มุมบนขวา
* **หน้าเพจโหลดช้า:** ตรวจสอบว่าไม่ได้เลือกโหลดข้อมูล Top 10 จำนวนมากเกินไปในครั้งเดียว

---
> **Avenger Planner** | *Fund Selection Tool v2.1 (Update: April 2026)*