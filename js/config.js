/* ============================================================
   Fund Selection Tool – FP2
   Configuration File
   ============================================================
   ⚠️  สิ่งที่ต้องแก้ไขก่อนใช้งาน:
   1. ใส่ CLIENT_ID จาก Google Cloud Console (ดูวิธีใน README.md)
   2. ตรวจสอบ / แก้ไขชื่อ Tab ของแต่ละ Google Sheet ให้ถูกต้อง
   ============================================================ */

const CONFIG = {

  /* Keep the app session across browser/app restarts. Google access tokens are
     deliberately not stored; the app requests a fresh token silently. */
  AUTH_SESSION_HOURS: 24,
  AUTH_BACKEND_URL: '',

  /* ----------------------------------------------------------
     Google OAuth 2.0 Client ID
     สร้างได้จาก: https://console.cloud.google.com/
     APIs & Services → Credentials → Create Credentials → OAuth client ID
     Application type: Web application
     ---------------------------------------------------------- */
  CLIENT_ID: '847033645933-m3do0gub1cdkesproseahvvrudavgi7p.apps.googleusercontent.com',

  /* ----------------------------------------------------------
     OAuth Scopes
     ---------------------------------------------------------- */
	  SCOPES: [
	    'https://www.googleapis.com/auth/spreadsheets',
	    'https://www.googleapis.com/auth/drive.readonly',
	    'https://www.googleapis.com/auth/drive.file',
	    'https://www.googleapis.com/auth/userinfo.email',
	    'https://www.googleapis.com/auth/userinfo.profile',
	  ].join(' '),

  /* ----------------------------------------------------------
     Google Sheet IDs (ดึงมาจากไฟล์ .gsheet อัตโนมัติ)
     ---------------------------------------------------------- */
  SHEETS: {
    MASTER_FUND_ID:        '10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig',
    THAI_FUND_QUALITY:     '1m1rSyJAel9atGMrmeRSwgYWa9wgc4gi7-3cp4Yvc8GM',
    FUND_KEY_PERFORMANCE:  '1s-0ciSOB2Tj0C9azeMXyd1zZxljOg8I5QilI0FgjdW4',
    PERCENTRANK_FREESTYLE: '1s-0ciSOB2Tj0C9azeMXyd1zZxljOg8I5QilI0FgjdW4',
    RAW_FOR_SEC:           '16agx9pl9adtMh-U7MCbgnIncBxpciCvFgsdurH6Ob8w',
    ISHARE_INDEX_PASSIVE:  '1miHQVkwEq7k4S0upoYsRg8Z0KsbZirbwtn7BKCk_Edw',
  },

  /* ----------------------------------------------------------
     SEC GitHub Actions Proxy
     ใช้ Google Apps Script เป็นตัวกลางรัน workflow แทนการใส่ GitHub token ในหน้าเว็บ
     ---------------------------------------------------------- */
  SEC_GITHUB_PROXY_URL: 'https://script.google.com/macros/s/AKfycbx64UW8yjrgrVBXAX7K5reiFHLldPRpNuXVXdmcxPBvznNGdNT5Jid176hN5d-cs7sHcQ/exec',
  SEC_GITHUB_PROXY_SECRET_KEY: '',

  /* ----------------------------------------------------------
     Page Configuration
     ⚠️  tabName = ชื่อ Tab (Sheet) ใน Google Sheet
         ตรวจสอบชื่อจาก URL ของ Google Sheet หรือแถบล่างสุด
         เช่น  tabName: 'Annualized'  หรือ  tabName: 'Sheet1'
     ---------------------------------------------------------- */
  PAGES: {

    'select-fund': {
      sheetId:  '1s-0ciSOB2Tj0C9azeMXyd1zZxljOg8I5QilI0FgjdW4',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'เลือกกองทุน',
      source:   'Fund Key Performance AVP',
    },

    'thai-annualized': {
      sheetId:  '1m1rSyJAel9atGMrmeRSwgYWa9wgc4gi7-3cp4Yvc8GM',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'กองทุนไทย Annualized Return',
      source:   'AVP Thai Fund for Quality',
    },

    'thai-annualized-v2': {
      sheetId:  '1m1rSyJAel9atGMrmeRSwgYWa9wgc4gi7-3cp4Yvc8GM',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'กองทุนไทย Annualized Return',
      source:   'AVP Thai Fund for Quality',
    },

    'thai-calendar': {
      sheetId:  '1m1rSyJAel9atGMrmeRSwgYWa9wgc4gi7-3cp4Yvc8GM',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'กองทุนไทย Calendar Year',
      source:   'AVP Thai Fund for Quality',
    },

    'master-annualized': {
      sheetId:  '10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'Master Fund Annualized Return',
      source:   'AVP Master Fund ID',
    },

    'master-annualized-v2': {
      sheetId:  '10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'Master Fund Annualized Return',
      source:   'AVP Master Fund ID',
    },

    'master-calendar': {
      sheetId:  '10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig',
      tabName:  '2026-Q1',   // ← แก้ชื่อ Tab ตามจริง
      title:    'Master Fund Calendar Year',
      source:   'AVP Master Fund ID',
    },

    'master-placeholder-1': {
      sheetId:  '16agx9pl9adtMh-U7MCbgnIncBxpciCvFgsdurH6Ob8w',
      tabName:  '2026-Q1',
      title:    'ค่าธรรมเนียม',
      source:   'Data For SEC API + AVP Master Fund ID',
    },

    'master-placeholder-2': {
      sheetId:  '10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig',
      tabName:  '2026-Q1',
      title:    'Top 10 Holding',
      source:   'AVP Master Fund ID',
    },

    'master-placeholder-3': {
      sheetId:  '10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig',
      tabName:  '2026-Q1',
      title:    'Cost Efficiency Master Fund 5Y',
      source:   'AVP Master Fund ID',
    },

    'master-placeholder-4': {
      sheetId:  '16agx9pl9adtMh-U7MCbgnIncBxpciCvFgsdurH6Ob8w',
      tabName:  '2026-Q1',
      title:    'ค่าธรรมเนียม V2',
      source:   'Data For SEC API + AVP Master Fund ID',
    },

    'master-placeholder-7': {
      sheetId:  '10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig',
      tabName:  '2026-Q1',
      title:    'Top 10 Holding V2',
      source:   'AVP Master Fund ID',
    },

    'master-placeholder-8': {
      title:  'Top 10 Holding',
      source: 'Multi-Fund Compare API',
    },

    'income-fund-1': {
      title:  'Income Fund',
      source: 'Fund Key Performance AVP',
    },

    'income-fund-2': {
      title:  'Income Fund 2',
      source: 'Fund Key Performance AVP',
    },

    'robustness-ft-import': {
      title:  'เตรียมข้อมูลจาก FT.com',
      source: 'FT Markets historical prices',
    },

    'ft-top10-holding': {
      title:  'Top 10 Holding',
      source: 'FT Markets qualitative local snapshot',
    },

    'upside-downside-capture': {
      title:  'Upside Downside Capture',
      source: 'FT Markets + Fund performance data',
    },

    'fund-list-update': {
      sheetId: '1akokyOCLUAYF7zCrpubOMl1UZHRv8cOeDsNTJzahbqU',
      tabName: 'fund_list_changes',
      title: 'Fund List Update',
      source: 'Google Sheet: fund_list_changes',
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
