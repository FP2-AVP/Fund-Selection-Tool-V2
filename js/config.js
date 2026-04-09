/* ============================================================
   Fund Selection Tool – FP2
   Configuration File
   ============================================================
   ⚠️  สิ่งที่ต้องแก้ไขก่อนใช้งาน:
   1. ใส่ CLIENT_ID จาก Google Cloud Console (ดูวิธีใน README.md)
   2. ตรวจสอบ / แก้ไขชื่อ Tab ของแต่ละ Google Sheet ให้ถูกต้อง
   ============================================================ */

const CONFIG = {

  /* ----------------------------------------------------------
     Google OAuth 2.0 Client ID
     สร้างได้จาก: https://console.cloud.google.com/
     APIs & Services → Credentials → Create Credentials → OAuth client ID
     Application type: Web application
     ---------------------------------------------------------- */
  CLIENT_ID: '1024922664942-958fev7r5815ucu2u4rkqpahagj6ish9.apps.googleusercontent.com',

  /* ----------------------------------------------------------
     OAuth Scopes
     ---------------------------------------------------------- */
  SCOPES: [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
  ].join(' '),

  /* ----------------------------------------------------------
     Google Sheet IDs (ดึงมาจากไฟล์ .gsheet อัตโนมัติ)
     ---------------------------------------------------------- */
  SHEETS: {
    MASTER_FUND_ID:        '1EYrIfINFg-WePEPr88PuDgsEW2OAZxkWtiq1Bml_1RA',
    THAI_FUND_QUALITY:     '1UhD2XgEmRUZXasGTxH-E0TMl5dyuu2sRxXRVGzXTCzE',
    PERCENTRANK_FREESTYLE: '1xt5YTBJoFKYc3gsLqhoCRsbrZHI8CSq5lRzBBqJwrdE',
    RAW_FOR_SEC:           '1qNzpxP5D9RAaxhwAVe5Wf8rBhrGNhXn7EYQgBD86cBs',
    ISHARE_INDEX_PASSIVE:  '1WTuIL2UtRg6cacUWZo5nsbinyRWPbMuPLNXjAStliUY',
  },

  /* ----------------------------------------------------------
     Page Configuration
     ⚠️  tabName = ชื่อ Tab (Sheet) ใน Google Sheet
         ตรวจสอบชื่อจาก URL ของ Google Sheet หรือแถบล่างสุด
         เช่น  tabName: 'Annualized'  หรือ  tabName: 'Sheet1'
     ---------------------------------------------------------- */
  PAGES: {

    'select-fund': {
      sheetId:  '1EYrIfINFg-WePEPr88PuDgsEW2OAZxkWtiq1Bml_1RA',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'เลือกกองทุน',
      source:   'AVP Master Fund ID',
    },

    'thai-annualized': {
      sheetId:  '1UhD2XgEmRUZXasGTxH-E0TMl5dyuu2sRxXRVGzXTCzE',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'กองทุนไทย Annualized Return',
      source:   'AVP Thai Fund for Quality',
    },

    'thai-annualized-v2': {
      sheetId:  '1UhD2XgEmRUZXasGTxH-E0TMl5dyuu2sRxXRVGzXTCzE',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'กองทุนไทย Annualized Return V2',
      source:   'AVP Thai Fund for Quality',
    },

    'thai-calendar': {
      sheetId:  '1UhD2XgEmRUZXasGTxH-E0TMl5dyuu2sRxXRVGzXTCzE',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'กองทุนไทย Calendar Year',
      source:   'AVP Thai Fund for Quality',
    },

    'master-annualized': {
      sheetId:  '1EYrIfINFg-WePEPr88PuDgsEW2OAZxkWtiq1Bml_1RA',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'Master Fund Annualized Return',
      source:   'AVP Master Fund ID',
    },

    'master-annualized-v2': {
      sheetId:  '1EYrIfINFg-WePEPr88PuDgsEW2OAZxkWtiq1Bml_1RA',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'Master Fund Annualized Return V2',
      source:   'AVP Master Fund ID',
    },

    'master-calendar': {
      sheetId:  '1EYrIfINFg-WePEPr88PuDgsEW2OAZxkWtiq1Bml_1RA',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'Master Fund Calendar Year',
      source:   'AVP Master Fund ID',
    },

    'master-placeholder-1': {
      sheetId:  '1qNzpxP5D9RAaxhwAVe5Wf8rBhrGNhXn7EYQgBD86cBs',
      tabName:  '2026-Q1',
      title:    'ค่าธรรมเนียม',
      source:   'Raw For Sec + AVP Master Fund ID',
    },

    'master-placeholder-2': {
      sheetId:  '1EYrIfINFg-WePEPr88PuDgsEW2OAZxkWtiq1Bml_1RA',
      tabName:  '2026-Q1',
      title:    'Top 10 Holding',
      source:   'AVP Master Fund ID',
    },

    'master-placeholder-3': {
      sheetId:  '1EYrIfINFg-WePEPr88PuDgsEW2OAZxkWtiq1Bml_1RA',
      tabName:  '2026-Q1',
      title:    'Cost Efficiency Master Fund 5Y',
      source:   'AVP Master Fund ID',
    },

    'master-placeholder-4': {
      sheetId:  '1qNzpxP5D9RAaxhwAVe5Wf8rBhrGNhXn7EYQgBD86cBs',
      tabName:  '2026-Q1',
      title:    'ค่าธรรมเนียม V2',
      source:   'Raw For Sec + AVP Master Fund ID',
    },
  },

  /* ----------------------------------------------------------
     Cache TTL (milliseconds)  default: 5 นาที
     ---------------------------------------------------------- */
  CACHE_TTL: 5 * 60 * 1000,

  /* ----------------------------------------------------------
     DEV MODE – ข้ามหน้า Login (true = ข้าม / false = ต้อง Login)
     ⚠️ ตั้งเป็น false ก่อน Deploy จริง
     ---------------------------------------------------------- */
  AUTH_MODE: 'token_popup',
  BYPASS_LOGIN: false,
};
