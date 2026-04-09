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

  /* ── ชื่อไฟล์ข้อมูล (แก้ตรงนี้เมื่อเปลี่ยน Quarter) ── */
  const FILE = {
    THAI:   'Data/AVP Thai Fund for Quality - 2026-Q1.json',
    MASTER: 'Data/AVP Master Fund ID - 2026-Q1.json',
    PERCENTRANK: 'Data/Percentrank Freestyle - 2026-Q1.json',
    RAW_SEC: 'Data/Raw For Sec - 2026-Q1.json',
  };

  /* ── Mapping หน้าเว็บ → ไฟล์ข้อมูล ── */
  if (CONFIG.PAGES?.['select-fund']) {
    CONFIG.PAGES['select-fund'].localFile = FILE.PERCENTRANK;
    CONFIG.PAGES['select-fund'].source    = 'Percentrank Freestyle';
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
    CONFIG.PAGES['master-placeholder-1'].localFile = FILE.RAW_SEC;
  }
  if (CONFIG.PAGES?.['master-placeholder-2']) {
    CONFIG.PAGES['master-placeholder-2'].localFile = FILE.MASTER;
  }
  if (CONFIG.PAGES?.['master-placeholder-3']) {
    CONFIG.PAGES['master-placeholder-3'].localFile = FILE.MASTER;
  }
  if (CONFIG.PAGES?.['master-placeholder-4']) {
    CONFIG.PAGES['master-placeholder-4'].localFile = FILE.RAW_SEC;
  }
}
