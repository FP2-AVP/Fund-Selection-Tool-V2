/* ============================================================
   Local Data Override
   ============================================================
   DATA_SOURCE options:
     'local_only'  – อ่านจาก JSON เท่านั้น (ไม่ต้อง login Google)
     'local_first' – อ่าน JSON ก่อน, fallback Google ถ้าหาไม่เจอ
     'google_first' – อ่าน Google Sheets ก่อน, fallback JSON ถ้าอ่านไม่สำเร็จ
     'google_only' – อ่านจาก Google Sheets เท่านั้น

   วิธีอัปเดตข้อมูลรายไตรมาส:
     1. Export CSV จาก Google Sheets แต่ละ Sheet
     2. ใส่ไฟล์ใหม่ในโฟลเดอร์ Data/
     3. เปลี่ยนชื่อไฟล์ด้านล่างให้ตรงกับไฟล์ที่ export มา
   ============================================================ */

if (typeof CONFIG !== 'undefined') {

  CONFIG.DATA_SOURCE = 'google_first';
  CONFIG.GITHUB_TOKEN = '';
  CONFIG.MASTER_ALLOCATIONS_API_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbylmcVwHgAuybiO7YiGftQQc7nfsylF01T7Nvni2BDkY9z10QFKaieISEM7cc8HHybKyg/exec';
  CONFIG.MASTER_ALLOCATIONS_API_SECRET_KEY = 'change-this-master-allocations-api-key';
  CONFIG.FIXED_INCOME_FACTORS_API_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzHr34Gg48W9awG2Hde3P_XOxGUVLf2k_W-ySRWt-2IaIa7VOD6IsYfm-LmfhaYSwRn/exec';
  CONFIG.FIXED_INCOME_FACTORS_API_SECRET_KEY = 'change-this-fixed-income-factors-api-key';
  CONFIG.INCOME_FUND_DIVIDEND_DB_URL = '';
  CONFIG.INCOME_FUND_DIVIDEND_API_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxaj7IgZ5h3dbL56ekA848PYvUnVhSh9u6OYcYhIr-PChoKxKgCD51OWQuGov4yWqJAjA/exec';
  CONFIG.INCOME_FUND_DIVIDEND_API_SECRET_KEY = 'change-this-income-dividend-api-key';
  CONFIG.FT_HISTORICAL_API_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbz7MBn8-2BfYCtWdh-uuiy__8phjGnoEQAz7fKcqOcEAspmjG6amOYYpTaEoQNKufBV/exec';
  CONFIG.FT_HISTORICAL_API_SECRET_KEY = 'change-this-ft-historical-api-key';

  /* ── ชื่อไฟล์ข้อมูล (แก้ตรงนี้เมื่อเปลี่ยน Quarter) ── */
  const FILE = {
    THAI:   'Data/AVP Thai Fund for Quality - 2026-Q1.json',
    MASTER: 'Data/AVP Master Fund ID - 2026-Q1.json',
    FUND_KEY_PERFORMANCE: 'Data/Fund Key Performance AVP - 2026-Q1.json',
    SEC_API: 'Data/Data For SEC API - 2026-Q1.json',
  };

  /* ── Mapping หน้าเว็บ → ไฟล์ข้อมูล ── */
  if (CONFIG.PAGES?.['select-fund']) {
    CONFIG.PAGES['select-fund'].localFile = FILE.FUND_KEY_PERFORMANCE;
    CONFIG.PAGES['select-fund'].source    = 'Fund Key Performance AVP';
  }
  if (CONFIG.PAGES?.['thai-annualized']) {
    CONFIG.PAGES['thai-annualized'].localFile = FILE.THAI;
  }
  if (CONFIG.PAGES?.['thai-annualized-v2']) {
    CONFIG.PAGES['thai-annualized-v2'].localFile = FILE.THAI;
  }
  if (CONFIG.PAGES?.['thai-calendar']) {
    CONFIG.PAGES['thai-calendar'].localFile = FILE.THAI;
  }
  if (CONFIG.PAGES?.['master-annualized']) {
    CONFIG.PAGES['master-annualized'].localFile = FILE.MASTER;
  }
  if (CONFIG.PAGES?.['master-annualized-v2']) {
    CONFIG.PAGES['master-annualized-v2'].localFile = FILE.MASTER;
  }
  if (CONFIG.PAGES?.['master-calendar']) {
    CONFIG.PAGES['master-calendar'].localFile = FILE.MASTER;
  }
  if (CONFIG.PAGES?.['master-placeholder-1']) {
    CONFIG.PAGES['master-placeholder-1'].localFile = FILE.SEC_API;
    CONFIG.PAGES['master-placeholder-1'].source    = 'Data For SEC API + AVP Master Fund ID';
  }
  if (CONFIG.PAGES?.['master-placeholder-2']) {
    CONFIG.PAGES['master-placeholder-2'].localFile = FILE.MASTER;
  }
  if (CONFIG.PAGES?.['master-placeholder-3']) {
    CONFIG.PAGES['master-placeholder-3'].localFile = FILE.MASTER;
  }
  if (CONFIG.PAGES?.['master-placeholder-4']) {
    CONFIG.PAGES['master-placeholder-4'].localFile = FILE.SEC_API;
    CONFIG.PAGES['master-placeholder-4'].source    = 'Data For SEC API + AVP Master Fund ID';
  }
  if (CONFIG.PAGES?.['master-placeholder-7']) {
    CONFIG.PAGES['master-placeholder-7'].localFile = FILE.MASTER;
  }
}
