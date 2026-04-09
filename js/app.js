/* ============================================================
   Fund Selection Tool – FP2
   Main Application Logic
   ============================================================ */
'use strict';

/* ── DOM helpers ── */
const $  = (sel, ctx = document) => ctx.querySelector(sel);
const $$ = (sel, ctx = document) => [...ctx.querySelectorAll(sel)];

/* ── Pagination ── */
const PAGE_SIZE_OPTIONS = [25, 50, 100, 200];

/* ── Application State ── */
const State = {
  page:         'dashboard',
  sortCol:      null,
  sortDir:      'asc',
  tablePage:    1,            // current pagination page
  pageSize:     25,
  selectFundFilters: {
    category: '',
    type: '',
    style: '',
    dividend: '',
    query: '',
    pageSize: 25,
  },
  selectFundSort: {
    key: '',
    dir: 'asc',
  },
  reportSorts: {
    'thai-annualized': { key: '', dir: 'asc' },
    'thai-annualized-rank': { key: '', dir: 'asc' },
    'thai-annualized-v2': { key: '', dir: 'asc' },
    'thai-calendar': { key: '', dir: 'asc' },
    'master-annualized': { key: 'r5y', dir: 'desc' },
    'master-annualized-v2': { key: 'r5y', dir: 'desc' },
    'master-calendar': { key: 'ret-2025', dir: 'desc' },
  },
  reportOptions: {
    'thai-annualized-v2-view': 'return',
    'thai-annualized-v2-left': 'return',
    'thai-annualized-v2-right': 'rank',
    'thai-calendar-left': 'return',
    'thai-calendar-right': 'rank',
    'thai-calendar-years': ['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
  },
  selectedKeys: new Set(),   // keys of selected rows in select-fund page
  selectedFunds: {},         // fundCode -> metadata for cross-page filtering
  highlights:   {},           // fundCode → colorIndex (0-4), persists across pages
  _cache:       {},
  _pageDataSource: {},
  _compareRows: null,
};

/* ── Highlight color palette (5 สี) ── */
const HL_COLORS = [
  { name: 'เหลือง', bg: '#FFFDE7', dot: '#F9A825' },
  { name: 'เขียว',  bg: '#E8F5E9', dot: '#2E7D32' },
  { name: 'ฟ้า',    bg: '#E3F2FD', dot: '#1565C0' },
  { name: 'ส้ม',    bg: '#FFF3E0', dot: '#E65100' },
  { name: 'ชมพู',   bg: '#FCE4EC', dot: '#AD1457' },
];

const PRESENTATION_TABLE_PRESETS = {
  thaiAnnualizedV2: {
    rowsPerSlide: 30,
    titleFontSizePx: 34,
    headerFontSizePx: 12,
    bodyFontSizePx: 10,
    groupHeightPx: 26,
    headerHeightPx: 26,
    columnWidthsPx: {
      code: 140,
      type: 65,
      dividend: 85,
      metric: 50,
    },
  },
  thaiCalendar: {
    rowsPerSlide: 30,
    titleFontSizePx: 34,
    headerFontSizePx: 12,
    bodyFontSizePx: 10,
    groupHeightPx: 26,
    headerHeightPx: 26,
    columnWidthsPx: {
      code: 140,
      type: 65,
      dividend: 85,
      metric: 50,
    },
  },
  masterAnnualizedV2: {
    rowsPerSlide: 30,
    titleFontSizePx: 34,
    headerFontSizePx: 12,
    bodyFontSizePx: 10,
    groupHeightPx: 26,
    headerHeightPx: 26,
    columnWidthsPx: {
      code: 140,
      type: 65,
      dividend: 85,
      metric: 50,
    },
  },
  masterCalendar: {
    rowsPerSlide: 30,
    titleFontSizePx: 34,
    headerFontSizePx: 12,
    bodyFontSizePx: 10,
    groupHeightPx: 26,
    headerHeightPx: 26,
    columnWidthsPx: {
      name: 190,
      currency: 72,
      thai: 120,
      metric: 46,
    },
  },
  masterFees2: {
    rowsPerSlide: 30,
    titleFontSizePx: 34,
    headerFontSizePx: 12,
    bodyFontSizePx: 10,
    groupHeightPx: 26,
    headerHeightPx: 26,
    columnWidthsPx: {
      master: 260,
      thai: 120,
      masterTer: 70,
      thaiTer: 80,
      date: 80,
      combined: 90,
    },
  },
  default: {
    rowsPerSlide: 30,
  },
};

const TOP_10_HOLDING_API_URL = 'https://script.google.com/macros/s/AKfycbyUUX7mbDdjfRvPJllUoGYk-GGDZdFw8Im_OMwrlo6YYE3x-BcLoD_7y4a9gS9GJNH3Ow/exec';

function getPresentationTablePreset(presetKey) {
  return PRESENTATION_TABLE_PRESETS[presetKey]
    || PRESENTATION_TABLE_PRESETS.default;
}

/* ── Fund classification helpers (derived from fund code) ── */
function deriveFundType(code) {
  if (/SSF/i.test(code))             return 'SSF';
  if (/RMF/i.test(code))             return 'RMF';
  if (/LTF/i.test(code))             return 'LTF';
  if (/TESGX|THAIESG/i.test(code))   return 'TESGX';
  return 'General';
}
function deriveDividend(code) {
  return (/-ID$|-RD$|-D$|-DIV$/i.test(code)) ? 'Dividend' : 'No Dividend';
}
function deriveStyle(code, masterFundName) {
  const s = (code + ' ' + (masterFundName || '')).toUpperCase();
  if (/INDEX|PASSIVE|ETF|SET50|SET100|SETCLMV|SETESG|SETTHSI/.test(s)) return 'Passive';
  return 'Active';
}

/* ============================================================
   UTILITIES
   ============================================================ */

/* Thai Buddhist-era date string */
function thaiDate() {
  const d   = new Date();
  const day = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'][d.getDay()];
  const mon = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
               'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'][d.getMonth()];
  return `วัน${day}ที่ ${d.getDate()} ${mon} ${d.getFullYear() + 543}`;
}

/* HTML escape */
function esc(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function findColumnIndex(headers, candidates) {
  const normalize = (value) => String(value ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  const lowerHeaders = headers.map(normalize);
  for (const candidate of candidates) {
    const wanted = normalize(candidate);
    const idx = lowerHeaders.findIndex(h => h === wanted);
    if (idx !== -1) return idx;
  }
  return -1;
}

function normalizeFundKey(value) {
  return String(value ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function getSelectedMasterIds() {
  return new Set(
    Object.values(State.selectedFunds)
      .map(f => String(f.masterId || '').trim())
      .filter(Boolean)
  );
}

function getFundHighlightIndex(row, headerMeta) {
  const selected = State.selectedFunds;
  const get = (idx) => idx >= 0 ? String(row[idx] ?? '').trim() : '';
  const rowCode = normalizeFundKey(get(headerMeta.codeIdx));
  const rowMasterId = get(headerMeta.masterIdIdx);

  if (rowCode && State.selectedKeys.has(rowCode) && State.highlights[rowCode] !== undefined) {
    return State.highlights[rowCode];
  }

  if (rowMasterId) {
    const matched = Object.entries(selected).find(([code, f]) =>
      State.selectedKeys.has(code) && String(f.masterId || '').trim() === rowMasterId
    )?.[1];
    if (matched && State.highlights[matched.code] !== undefined) return State.highlights[matched.code];
  }

  return undefined;
}

function buildHighlightSelect(code, currentValue) {
  const selectedColor = currentValue !== undefined && HL_COLORS[currentValue]
    ? HL_COLORS[currentValue]
    : null;
  const swatch = ['🟨', '🟩', '🟦', '🟧', '🩷'];
  const options = [
    `<option value="">ไม่เลือกสี</option>`,
    ...HL_COLORS.map((color, index) =>
      `<option value="${index}" ${String(currentValue) === String(index) ? 'selected' : ''}>${esc((swatch[index] || '■') + ' ' + color.name)}</option>`
    ),
  ];
  return `
    <select class="hl-select${selectedColor ? ' has-color' : ''}" data-fund="${esc(code)}" aria-label="เลือกสีไฮไลต์ของ ${esc(code)}"${selectedColor ? ` style="background:${selectedColor.bg};border-color:${selectedColor.dot};color:${selectedColor.dot}"` : ''}>
      ${options.join('')}
    </select>`;
}

function pageToolActions(pageKey, sourceLabel = '', extraActions = '') {
  const exportable = new Set([
    'thai-annualized',
    'thai-annualized-rank',
    'thai-annualized-v2',
    'thai-calendar',
    'master-annualized',
    'master-annualized-v2',
    'master-calendar',
    'master-placeholder-1',
    'master-placeholder-3',
    'master-placeholder-4',
  ]);
  if (!exportable.has(pageKey)) return '';
  const sourceBadge = getPageDataSourceBadge(pageKey);
  return `
    <div class="page-tools">
      <div class="page-tools-meta">
        ${sourceLabel ? `<span class="badge badge-source">แหล่งข้อมูลจาก: ${esc(sourceLabel)}</span>` : ''}
        ${sourceBadge ? `<span class="badge badge-data-origin">${esc(sourceBadge)}</span>` : ''}
        ${extraActions}
      </div>
      <div class="page-tools-actions">
        <button class="btn btn-primary" id="btn-send-table" title="ส่งข้อมูลไปทำ Presentation">
          ส่งข้อมูลไปทำ presentation
        </button>
      </div>
    </div>`;
}

function getPageDataSourceBadge(pageKey) {
  return State._pageDataSource?.[pageKey] || '';
}

async function elementToImageBlob(el) {
  const rect = el.getBoundingClientRect();
  const cloned = el.cloneNode(true);
  const styles = [...document.styleSheets].map(sheet => {
    try {
      return [...sheet.cssRules].map(rule => rule.cssText).join('\n');
    } catch {
      return '';
    }
  }).join('\n');

  const html = `
    <html xmlns="http://www.w3.org/1999/xhtml">
      <head><style>${styles}</style></head>
      <body style="margin:0;background:#ffffff;">${cloned.outerHTML}</body>
    </html>`;

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${Math.ceil(rect.width)}" height="${Math.ceil(rect.height)}">
      <foreignObject width="100%" height="100%">${html}</foreignObject>
    </svg>`;

  const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(blob);

  try {
    const img = await new Promise((resolve, reject) => {
      const node = new Image();
      node.onload = () => resolve(node);
      node.onerror = reject;
      node.src = url;
    });
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(rect.width);
    canvas.height = Math.ceil(rect.height);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0);
    return await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
  } finally {
    URL.revokeObjectURL(url);
  }
}

function elementToImageSvgDataURL(el) {
  const rect = el.getBoundingClientRect();
  const cloned = el.cloneNode(true);
  const styles = [...document.styleSheets].map(sheet => {
    try {
      return [...sheet.cssRules].map(rule => rule.cssText).join('\n');
    } catch {
      return '';
    }
  }).join('\n');

  const html = `
    <html xmlns="http://www.w3.org/1999/xhtml">
      <head><style>${styles}</style></head>
      <body style="margin:0;background:#ffffff;">${cloned.outerHTML}</body>
    </html>`;

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${Math.ceil(rect.width)}" height="${Math.ceil(rect.height)}">
      <foreignObject width="100%" height="100%">${html}</foreignObject>
    </svg>`;

  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

async function blobToDataURL(blob) {
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function createCaptureTarget(area, card) {
  const reportFrame = $('#report-frame', area);
  const reportStage = $('#report-stage', area);
  if (!reportFrame || !reportStage) {
    return { node: card, cleanup: () => {} };
  }

  const naturalWidth = card.scrollWidth;
  const naturalHeight = card.scrollHeight;
  const targetWidth = 1600;
  const targetHeight = 900;
  const padX = 36;
  const padY = 28;
  const availableWidth = targetWidth - (padX * 2);
  const availableHeight = targetHeight - (padY * 2);
  const scale = Math.min(availableWidth / naturalWidth, availableHeight / naturalHeight, 1);

  const wrapper = document.createElement('div');
  wrapper.style.position = 'fixed';
  wrapper.style.left = '-100000px';
  wrapper.style.top = '0';
  wrapper.style.width = `${targetWidth}px`;
  wrapper.style.height = `${targetHeight}px`;
  wrapper.style.display = 'flex';
  wrapper.style.alignItems = 'center';
  wrapper.style.justifyContent = 'center';
  wrapper.style.padding = `${padY}px ${padX}px`;
  wrapper.style.boxSizing = 'border-box';
  wrapper.style.background = '#ffffff';
  wrapper.style.overflow = 'hidden';

  const stage = document.createElement('div');
  stage.style.width = `${naturalWidth}px`;
  stage.style.height = `${naturalHeight}px`;
  stage.style.transformOrigin = 'top left';
  stage.style.transform = `scale(${scale})`;
  stage.style.flex = '0 0 auto';

  const clonedCard = card.cloneNode(true);
  if (clonedCard.classList) clonedCard.classList.add('report-card-presentation');
  clonedCard.style.margin = '0';
  clonedCard.style.boxShadow = 'none';

  stage.appendChild(clonedCard);
  wrapper.appendChild(stage);
  document.body.appendChild(wrapper);

  return {
    node: wrapper,
    cleanup: () => wrapper.remove(),
  };
}

const STUDIO_DB_NAME = 'avenger-studio-db';
const STUDIO_DB_VERSION = 2;
const STUDIO_QUEUE_STORE = 'presentationQueue';
const STUDIO_SLIDES_STORE = 'slides';

function openStudioDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(STUDIO_DB_NAME, STUDIO_DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STUDIO_QUEUE_STORE)) {
        db.createObjectStore(STUDIO_QUEUE_STORE, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STUDIO_SLIDES_STORE)) {
        db.createObjectStore(STUDIO_SLIDES_STORE, { keyPath: 'id' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function studioDbGetAll(storeName) {
  const db = await openStudioDb();
  return await new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readonly');
    const req = tx.objectStore(storeName).getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

async function studioDbPut(storeName, value) {
  const db = await openStudioDb();
  return await new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readwrite');
    tx.objectStore(storeName).put(value);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function studioDbDeleteMany(storeName, ids) {
  const db = await openStudioDb();
  return await new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readwrite');
    const store = tx.objectStore(storeName);
    ids.forEach(id => store.delete(id));
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

function bindPageImageActions(area, cardId, filename) {
  const MAX_PRESENTATION_ROWS = 90;
  const card = $(`#${cardId}`, area);
  if (!card) return;
  const tableBtn = $('#btn-send-table', area);
  if (tableBtn) {
    const supported = typeof App._currentTableExport === 'function';
    tableBtn.disabled = !supported;
    tableBtn.title = supported
      ? 'ส่งข้อมูลไปทำ Presentation'
      : 'หน้านี้ยังไม่รองรับการส่งเป็นตาราง';
    tableBtn.style.opacity = supported ? '' : '.55';
    tableBtn.style.cursor = supported ? '' : 'not-allowed';
  }

  $('#btn-send-table', area)?.addEventListener('click', async () => {
    try {
      if (typeof App._currentTableExport !== 'function') {
        toast('หน้านี้ยังไม่รองรับการส่งเป็นตาราง', 'warning');
        return;
      }
      const queued = await App.readPresentationQueue();
      const staleImageIds = queued
        .filter(item => item && item.kind !== 'table')
        .map(item => item.id)
        .filter(Boolean);
      if (staleImageIds.length) {
        await studioDbDeleteMany(STUDIO_QUEUE_STORE, staleImageIds);
      }
      const payload = App._currentTableExport();
      if (!payload?.rows?.length) {
        toast('ไม่พบข้อมูลตารางสำหรับส่งเข้า Presentation', 'warning');
        return;
      }
      const totalRows = payload.rows.length;
      const clippedRows = totalRows > MAX_PRESENTATION_ROWS
        ? payload.rows.slice(0, MAX_PRESENTATION_ROWS)
        : payload.rows;
      const finalPayload = clippedRows.length === totalRows
        ? payload
        : { ...payload, rows: clippedRows };
      const item = {
        id: `table-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        kind: 'table',
        replaceImageSlides: true,
        filename,
        page: State.page,
        createdAt: new Date().toISOString(),
        payload: finalPayload,
      };
      await App.queuePresentationSlide(item);
      if (totalRows > MAX_PRESENTATION_ROWS) {
        toast(`ส่งข้อมูลไปทำ Presentation แล้ว (จำกัด ${MAX_PRESENTATION_ROWS} rows จาก ${totalRows} rows)`, 'warning', 4200);
      } else {
        toast('ส่งข้อมูลไปทำ Presentation แล้ว', 'success');
      }
    } catch (err) {
      toast(`ส่งตารางไม่สำเร็จ: ${err.message || err}`, 'error');
    }
  });
}

/* Show toast notification */
function toast(msg, type = 'info', dur = 3200) {
  const el = $('#toast');
  el.textContent = msg;
  el.className   = `toast ${type}`;
  el.classList.remove('hidden');
  clearTimeout(App._toastTimer);
  App._toastTimer = setTimeout(() => el.classList.add('hidden'), dur);
}

/* Set loading state in an area */
function setLoading(area, msg = 'กำลังโหลดข้อมูล...') {
  area.innerHTML = `
    <div class="card">
      <div class="state-box">
        <div class="spinner"></div>
        <span>${esc(msg)}</span>
      </div>
    </div>`;
}

/* Set error state */
function setError(area, msg, retryPage) {
  area.innerHTML = `
    <div class="card">
      <div class="err-box">
        <div class="state-icon">⚠️</div>
        <strong>เกิดข้อผิดพลาด</strong>
        <p style="white-space:pre-wrap;max-width:520px;font-size:.88rem">${esc(msg)}</p>
        ${retryPage ? `<button class="btn btn-ghost btn-sm" onclick="App.navigate('${retryPage}')">↻ ลองใหม่</button>` : ''}
      </div>
    </div>`;
}

/* Detect if a value looks like a number */
function isNum(v) {
  return v !== '' && v !== null && v !== undefined &&
    !isNaN(parseFloat(String(v).replace(/,/g, '')));
}

/* Parse numeric (handle comma thousands) */
function parseNum(v) {
  return parseFloat(String(v).replace(/,/g, ''));
}

function isLocalMode() {
  return CONFIG.DATA_SOURCE === 'local_first' || CONFIG.DATA_SOURCE === 'local_only';
}

function ensureExtendedPageConfigs() {
  if (!CONFIG.PAGES) CONFIG.PAGES = {};

  const defaults = {
    'master-placeholder-1': {
      sheetId: CONFIG.SHEETS?.RAW_FOR_SEC || '',
      tabName: '2026-Q1',
      title: 'ค่าธรรมเนียม',
      source: 'Raw For Sec + AVP Master Fund ID',
      localFile: 'Data/Raw For Sec - 2026-Q1.json',
    },
    'master-placeholder-2': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'Top 10 Holding',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
    },
    'master-placeholder-3': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'Cost Efficiency Master Fund 5Y',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
    },
    'master-placeholder-4': {
      sheetId: CONFIG.SHEETS?.RAW_FOR_SEC || '',
      tabName: '2026-Q1',
      title: 'ค่าธรรมเนียม 2',
      source: 'Raw For Sec + AVP Master Fund ID',
      localFile: 'Data/Raw For Sec - 2026-Q1.json',
    },
  };

  Object.entries(defaults).forEach(([pageKey, fallback]) => {
    CONFIG.PAGES[pageKey] = {
      ...fallback,
      ...(CONFIG.PAGES[pageKey] || {}),
    };
  });
}

async function fetchLocalRows(localFile) {
  if (!localFile) {
    throw new Error('Local data file is not configured');
  }

  const resp = await fetch(localFile, { cache: 'no-store' });
  if (!resp.ok) {
    throw new Error(`Local data not found (${resp.status})`);
  }

  const payload = await resp.json();
  const rows = Array.isArray(payload) ? payload : payload?.values;
  if (!Array.isArray(rows)) {
    throw new Error(`Invalid local data format in ${localFile}`);
  }

  return rows;
}

async function fetchPageData(pageKey) {
  ensureExtendedPageConfigs();
  const cfg = CONFIG.PAGES[pageKey];
  if (!cfg) throw new Error(`Unknown page: ${pageKey}`);

  const mode = CONFIG.DATA_SOURCE || 'google_first';

  if (mode === 'local_only') {
    const rows = await fetchLocalRows(cfg.localFile);
    State._pageDataSource[pageKey] = 'Source: Local JSON';
    return rows;
  }

  if (mode === 'local_first') {
    try {
      const rows = await fetchLocalRows(cfg.localFile);
      State._pageDataSource[pageKey] = 'Source: Local JSON';
      return rows;
    } catch (err) {
      const rows = await SheetsAPI.fetchSheetData(cfg.sheetId, cfg.tabName);
      State._pageDataSource[pageKey] = 'Source: Google Sheets fallback';
      return rows;
    }
  }

  if (mode === 'google_first') {
    try {
      const rows = await SheetsAPI.fetchSheetData(cfg.sheetId, cfg.tabName);
      State._pageDataSource[pageKey] = 'Source: Google Sheets';
      return rows;
    } catch (err) {
      if (!cfg.localFile) throw err;
      const rows = await fetchLocalRows(cfg.localFile);
      State._pageDataSource[pageKey] = 'Source: Local JSON fallback';
      return rows;
    }
  }

  const rows = await SheetsAPI.fetchSheetData(cfg.sheetId, cfg.tabName);
  State._pageDataSource[pageKey] = 'Source: Google Sheets';
  return rows;
}

/* ============================================================
   DATA CACHE
   ============================================================ */
async function fetchCached(pageKey) {
  const key = `page::${pageKey}`;
  const now = Date.now();
  if (State._cache[key] && now - State._cache[key].ts < CONFIG.CACHE_TTL) {
    return State._cache[key].data;
  }
  const data = await fetchPageData(pageKey);
  State._cache[key] = { data, ts: now };
  return data;
}

function clearCache() {
  State._cache = {};
  State._pageDataSource = {};
}

const PERCENTILE_HEAT_RANGES = [
  { min: 0, max: 5, color: '#7ABC81', text: '#103c1c' },
  { min: 5.01, max: 25, color: '#A8D086', text: '#26411d' },
  { min: 25.01, max: 50, color: '#CBDFB8', text: '#2d4724' },
  { min: 50.01, max: 75, color: '#FCEC92', text: '#5f4a08' },
  { min: 75.01, max: 95, color: '#EDB392', text: '#6c3518' },
  { min: 95.01, max: 100, color: '#E7726F', text: '#5d1111' },
];

function normalizePercentileValue(value) {
  const n = parseNum(value);
  if (Number.isNaN(n)) return NaN;
  return Math.abs(n) <= 1 ? n * 100 : n;
}

function percentileHeatStyle(value) {
  const n = normalizePercentileValue(value);
  if (Number.isNaN(n)) return '';
  const match = PERCENTILE_HEAT_RANGES.find(range => n >= range.min && n <= range.max);
  if (!match) return '';
  return `background:${match.color};color:${match.text};`;
}

function formatPercentileDisplay(value) {
  const n = normalizePercentileValue(value);
  if (Number.isNaN(n)) return '';
  return `${n.toFixed(0)}%`;
}

function formatReturnDisplay(value) {
  const n = parseNum(value);
  if (Number.isNaN(n)) return '';
  return n.toFixed(2);
}

function rankHeatStyle(value) {
  const n = parseNum(value);
  if (Number.isNaN(n)) return '';
  if (n <= 2) return 'background:#7cc47f;color:#103c1c;';
  if (n <= 4) return 'background:#c8dfb3;color:#26411d;';
  if (n <= 6) return 'background:#fde68a;color:#5f4a08;';
  return '';
}

function buildPercentrankFunds(rows) {
  const headers = rows[0] || [];
  const idx = {
    name: findColumnIndex(headers, ['Name']),
    code: findColumnIndex(headers, ['Fund Code']),
    type: findColumnIndex(headers, ['Fund Type']),
    dividend: findColumnIndex(headers, ['Dividend']),
    style: findColumnIndex(headers, ['Style']),
  };

  const get = (row, i) => i >= 0 ? String(row[i] ?? '').trim() : '';
  return rows.slice(1).map(row => ({
    row,
    code: get(row, idx.code),
    key: normalizeFundKey(get(row, idx.code)),
    name: get(row, idx.name),
    type: get(row, idx.type),
    dividend: get(row, idx.dividend),
    style: get(row, idx.style),
  })).filter(f => f.code);
}

function buildSelectedFundsCatalog(rows) {
  const headers = rows[0] || [];
  const CI = {
    CATEGORY: findColumnIndex(headers, ['AVP® Category', 'AVP®  Category', 'AVP Category']),
    CODE: findColumnIndex(headers, ['Fund Code', 'FundId']),
    MASTER: findColumnIndex(headers, ['Master Fund']),
    ISIN: findColumnIndex(headers, ['ISIN']),
    TYPE: findColumnIndex(headers, ['Fund Type', 'Type']),
    DIVIDEND: findColumnIndex(headers, ['Dividend']),
    STYLE: findColumnIndex(headers, ['Style']),
  };

  return rows.slice(1).map(r => ({
    category:   r[CI.CATEGORY]    || '',
    code:       r[CI.CODE]        || '',
    key:        normalizeFundKey(r[CI.CODE] || ''),
    masterId:   r[CI.ISIN]        || '',
    masterName: r[CI.MASTER]      || '',
    type:       r[CI.TYPE]        || deriveFundType(r[CI.CODE] || ''),
    dividend:   r[CI.DIVIDEND]    || deriveDividend(r[CI.CODE] || ''),
    style:      r[CI.STYLE]       || deriveStyle(r[CI.CODE] || '', r[CI.MASTER] || ''),
  })).filter(f => f.code);
}

async function ensureSelectedFundsCatalog() {
  if (Object.keys(State.selectedFunds || {}).length) {
    return State.selectedFunds;
  }
  const rows = await fetchCached('select-fund');
  const allFunds = buildSelectedFundsCatalog(rows);
  State.selectedFunds = Object.fromEntries(allFunds.map(f => [f.key, f]));
  return State.selectedFunds;
}

function sortIndicator(active, dir) {
  const text = !active ? '↕' : (dir === 'asc' ? '↑' : '↓');
  return `<span class="sort-indicator" aria-hidden="true">${text}</span>`;
}

function renderSortLabel(label, active, dir, escapeLabel = true) {
  const safeLabel = escapeLabel ? esc(label) : label;
  return `<span class="sort-label ${active ? 'is-active' : ''}"><span class="sort-text">${safeLabel}</span>${sortIndicator(active, dir)}</span>`;
}

function isMissingValue(value) {
  const s = String(value ?? '').trim();
  return s === '' || s === '-' || s === '–';
}

function compareValues(a, b, dir = 'asc') {
  const av = a ?? '';
  const bv = b ?? '';
  const aMissing = isMissingValue(av);
  const bMissing = isMissingValue(bv);
  if (aMissing && bMissing) return 0;
  if (aMissing) return 1;
  if (bMissing) return -1;
  const an = parseNum(av);
  const bn = parseNum(bv);
  let result;
  if (!Number.isNaN(an) && !Number.isNaN(bn)) result = an - bn;
  else result = String(av).localeCompare(String(bv), 'th');
  return dir === 'asc' ? result : -result;
}

function toggleNamedSort(target, key) {
  if (target.key === key) {
    if (target.dir === 'asc') target.dir = 'desc';
    else {
      target.key = '';
      target.dir = 'asc';
    }
    return;
  }
  target.key = key;
  target.dir = 'asc';
}

function buildMetricRanks(items, keys, getValue) {
  const ranks = {};
  const totals = {};
  keys.forEach(key => {
    const values = items
      .map(item => ({ code: item.code, value: parseNum(getValue(item, key)) }))
      .filter(entry => !Number.isNaN(entry.value))
      .sort((a, b) => b.value - a.value);
    totals[key] = values.length;
    let currentRank = 0;
    let lastValue = null;
    values.forEach((entry, index) => {
      if (lastValue === null || entry.value !== lastValue) {
        currentRank = index + 1;
        lastValue = entry.value;
      }
      if (!ranks[entry.code]) ranks[entry.code] = {};
      ranks[entry.code][key] = currentRank;
    });
  });
  return { ranks, totals };
}

function rankCellStyle(rank, total) {
  const r = parseNum(rank);
  if (Number.isNaN(r) || !total) return '';
  if (total === 1) return 'background:#7ABC81;color:#24364f;';
  const ratio = Math.max(0, Math.min(1, (r - 1) / (total - 1)));
  const start = { r: 122, g: 188, b: 129 };
  const end = { r: 255, g: 255, b: 255 };
  const mix = (a, b) => Math.round(a + (b - a) * ratio);
  const bg = `rgb(${mix(start.r, end.r)}, ${mix(start.g, end.g)}, ${mix(start.b, end.b)})`;
  return `background:${bg};color:#24364f;`;
}

function extractInlineColors(styleText = '') {
  const bg = /background:\s*([^;]+)/i.exec(styleText)?.[1]?.trim() || '';
  const color = /color:\s*([^;]+)/i.exec(styleText)?.[1]?.trim() || '';
  return { bg, color };
}

function applyPresentationFit(area, frameId, stageId, cardId) {
  const frame = $(`#${frameId}`, area);
  const stage = $(`#${stageId}`, area);
  const card = $(`#${cardId}`, area);
  if (!frame || !stage || !card) return;

  const naturalWidth = card.scrollWidth;
  const naturalHeight = card.scrollHeight;
  const availableWidth = Math.max(240, frame.clientWidth - 24);
  const availableHeight = Math.max(180, frame.clientHeight - 24);
  const scale = Math.min(availableWidth / naturalWidth, availableHeight / naturalHeight, 1);

  stage.style.width = `${naturalWidth}px`;
  stage.style.height = `${naturalHeight}px`;
  stage.style.transform = `scale(${scale})`;
}

// Generic table payload builder — converts flat header+rows arrays into a Studio table payload.
// Used by pages that don't have custom formatters (master-annualized, thai-calendar, master-calendar, etc.)
function buildSimpleTablePayload(title, source, headers, dataRows) {
  const cols = headers.map((h, i) => ({
    key: `col-${i}`,
    label: String(h || ''),
    weight: i === 0 ? 1.8 : 0.9,
    align: i === 0 ? 'left' : 'center',
    bg: '#3f5d8c',
    color: '#ffffff',
  }));
  const rows = dataRows.map((row, rowIndex) => ({
    cells: headers.map((_, i) => ({
      text: String(row[i] ?? ''),
      bg: rowIndex % 2 === 0 ? '#f8fbff' : '#eef4fb',
      color: '#334155',
      weight: i === 0 ? 1.8 : 0.9,
      align: i === 0 ? 'left' : 'center',
    })),
  }));
  return {
    kind: 'table',
    title: title || '',
    subtitle: '',
    source: source || '',
    headerGroups: [{ label: '', span: cols.length, bg: '#dbe4f0', color: '#334155' }],
    columns: cols,
    rows,
  };
}

function annualizedGroupTheme(cfg) {
  switch (cfg?.mode) {
    case 'pct':
      return { bg: '#c92c1b', color: '#ffffff' };
    case 'rank':
      return { bg: '#2f537f', color: '#ffffff' };
    default:
      return { bg: '#4ba3dc', color: '#ffffff' };
  }
}

function presentationSortTheme(isActive, dir, fallbackBg = '#365a88', fallbackColor = '#ffffff') {
  if (!isActive) {
    return {
      bg: fallbackBg,
      color: fallbackColor,
      suffix: '',
    };
  }
  return {
    bg: '#f7d774',
    color: '#4e3500',
    suffix: dir === 'desc' ? ' ↓' : ' ↑',
  };
}

function buildPresentationSortColumn(sortState, sortKey, label, weight, options = {}) {
  const {
    align = 'center',
    bg = '#365a88',
    color = '#ffffff',
    widthPx,
  } = options;
  const theme = presentationSortTheme(sortState?.key === sortKey, sortState?.dir, bg, color);
  return {
    key: sortKey,
    label: `${label}${theme.suffix}`,
    weight,
    widthPx,
    align,
    bg: theme.bg,
    color: theme.color,
  };
}

function buildPresentationTablePayload({
  presetKey = 'default',
  title,
  source,
  rowsPerSlide,
  headerGroups,
  columns,
  rows,
}) {
  const preset = getPresentationTablePreset(presetKey);
  return {
    kind: 'table',
    rowsPerSlide: rowsPerSlide ?? preset.rowsPerSlide ?? 18,
    titleFontSizePx: preset.titleFontSizePx,
    headerFontSizePx: preset.headerFontSizePx,
    bodyFontSizePx: preset.bodyFontSizePx,
    groupHeightPx: preset.groupHeightPx,
    headerHeightPx: preset.headerHeightPx,
    title,
    subtitle: '',
    source,
    headerGroups,
    columns,
    rows,
  };
}

function buildThaiAnnualizedExportPayload(pageKey, sorted, metricKeys, leftCfg, rightCfg, helpers, sortState) {
  const preset = getPresentationTablePreset('thaiAnnualizedV2');
  const leftTheme = annualizedGroupTheme(leftCfg);
  const rightTheme = annualizedGroupTheme(rightCfg);
  const metricCell = (cfg, fund, key, baseRowBg) => {
    if (cfg.mode === 'return') {
      return {
        text: cfg.tableValue(fund, key, helpers),
        bg: baseRowBg,
        color: '#475569',
        weight: 0.78,
        strong: false,
      };
    }

    const styleText = cfg.mode === 'rank'
      ? rankCellStyle(helpers.sortable(fund, `rank${key.slice(1)}`), helpers.rankTotals[key])
      : percentileHeatStyle(helpers.get(fund.row, helpers.col[`p${key.slice(1)}`]));
    const colors = extractInlineColors(styleText);
    return {
      text: cfg.tableValue(fund, key, helpers),
      bg: colors.bg || baseRowBg,
      color: colors.color || '#334155',
      weight: 0.78,
      strong: true,
    };
  };
  const rows = sorted.map((fund, rowIndex) => {
    const baseRowBg = rowIndex % 2 === 0 ? '#f8fbff' : '#eef4fb';
    const highlightIdx = State.highlights[fund.key];
    const highlightBg = highlightIdx !== undefined ? (HL_COLORS[highlightIdx]?.bg || baseRowBg) : baseRowBg;
    return {
      cells: [
        { text: fund.code, bg: highlightBg, color: '#35507a', weight: 2.2, align: 'center', strong: true },
        { text: fund.type || '-', bg: baseRowBg, color: '#475569', weight: 1.05 },
        { text: fund.dividend || '-', bg: baseRowBg, color: '#475569', weight: 1.3 },
        ...metricKeys.map(key => metricCell(leftCfg, fund, key, baseRowBg)),
        ...metricKeys.map(key => metricCell(rightCfg, fund, key, baseRowBg)),
      ],
    };
  });

  const columns = [
    buildPresentationSortColumn(sortState, 'code', 'ชื่อกอง', 2.2, { align: 'center', widthPx: preset.columnWidthsPx?.code }),
    buildPresentationSortColumn(sortState, 'type', 'ประเภท', 1.05, { widthPx: preset.columnWidthsPx?.type }),
    buildPresentationSortColumn(sortState, 'dividend', 'Dividend', 1.3, { widthPx: preset.columnWidthsPx?.dividend }),
    ...metricKeys.map(key => {
      const label = key === 'rytd' ? 'YTD' : key.slice(1).toUpperCase();
      return buildPresentationSortColumn(sortState, leftCfg.sortKeyForMetric(key), label, 0.78, {
        ...leftTheme,
        widthPx: preset.columnWidthsPx?.metric,
      });
    }),
    ...metricKeys.map(key => {
      const label = key === 'rytd' ? 'YTD' : key.slice(1).toUpperCase();
      return buildPresentationSortColumn(sortState, rightCfg.sortKeyForMetric(key), label, 0.78, {
        ...rightTheme,
        widthPx: preset.columnWidthsPx?.metric,
      });
    }),
  ];

  return buildPresentationTablePayload({
    presetKey: 'thaiAnnualizedV2',
    title: CONFIG.PAGES[pageKey]?.title || 'Thai Annualized Report',
    source: CONFIG.PAGES['select-fund']?.source || 'Percentrank Freestyle',
    headerGroups: [
      { label: '', span: 3, bg: '#dbe4f0', color: '#334155' },
      { label: leftCfg.groupTitle, span: 6, ...leftTheme },
      { label: rightCfg.groupTitle, span: 6, ...rightTheme },
    ],
    columns,
    rows,
  });
}

function buildMasterAnnualizedExportPayload(sorted, metricKeys, masterLinks, rankMap, rankTotals, CI, get, sortState) {
  const returnTheme = { bg: '#4ba3dc', color: '#ffffff' };
  const rankTheme = { bg: '#2f537f', color: '#ffffff' };
  const rows = sorted.map((item, rowIndex) => {
    const baseRowBg = rowIndex % 2 === 0 ? '#f8fbff' : '#eef4fb';
    const thaiCodes = masterLinks[item.key] || [];
    const highlightedThai = thaiCodes.find(f => State.highlights[f.key] !== undefined);
    const thaiBg = highlightedThai ? (HL_COLORS[State.highlights[highlightedThai.key]]?.bg || baseRowBg) : baseRowBg;
    return {
      cells: [
        { text: item.name, bg: baseRowBg, color: '#35507a', weight: 2.4, align: 'left', strong: true },
        { text: thaiCodes.length ? thaiCodes.map(f => f.code).join(', ') : '-', bg: thaiBg, color: '#475569', weight: 1.9, align: 'left' },
        ...metricKeys.map(key => ({
          text: formatReturnDisplay(get(item.row, CI[key])) || '-',
          bg: baseRowBg,
          color: '#475569',
          weight: 0.82,
        })),
        ...metricKeys.map(key => {
          const rankValue = rankMap[item.key]?.[key] ?? '';
          const colors = extractInlineColors(rankCellStyle(rankValue, rankTotals[key]));
          return {
            text: String(rankValue || '-'),
            bg: colors.bg || baseRowBg,
            color: colors.color || '#334155',
            weight: 0.82,
            strong: true,
          };
        }),
      ],
    };
  });

  const metricLabel = key => key === 'rytd' ? 'YTD' : key.slice(1).toUpperCase();

  return {
    kind: 'table',
    title: CONFIG.PAGES['master-annualized']?.title || 'Master Fund Annualized',
    subtitle: '',
    source: CONFIG.PAGES['master-annualized']?.source || 'AVP Master Fund ID',
    headerGroups: [
      { label: '', span: 2, bg: '#dbe4f0', color: '#334155' },
      { label: 'ผลตอบแทน (%)', span: metricKeys.length, ...returnTheme },
      { label: 'อันดับในกลุ่มที่แสดง', span: metricKeys.length, ...rankTheme },
    ],
    columns: [
      baseSort('name', 'Master Fund', 2.4, 'left'),
      baseSort('thai', 'กองทุนในไทย', 1.9, 'left'),
      ...metricKeys.map(key => {
        const label = metricLabel(key);
        const theme = presentationSortTheme(sortState?.key === key, sortState?.dir, returnTheme.bg, returnTheme.color);
        return { key, label: `${label}${theme.suffix}`, weight: 0.82, bg: theme.bg, color: theme.color };
      }),
      ...metricKeys.map(key => {
        const sortKey = `rank${key.slice(1)}`;
        const label = metricLabel(key);
        const theme = presentationSortTheme(sortState?.key === sortKey, sortState?.dir, rankTheme.bg, rankTheme.color);
        return { key: sortKey, label: `${label}${theme.suffix}`, weight: 0.82, bg: theme.bg, color: theme.color };
      }),
    ],
    rows,
  };
}

function buildMasterAnnualizedV2ExportPayload(sorted, metricKeys, linksByRowKey, rankMap, rankTotals, CI, get, sortState) {
  const returnTheme = { bg: '#4ba3dc', color: '#ffffff' };
  const rankTheme = { bg: '#2f537f', color: '#ffffff' };
  const rows = sorted.map((item, rowIndex) => {
    const baseRowBg = rowIndex % 2 === 0 ? '#f8fbff' : '#eef4fb';
    const thaiFundsForRow = linksByRowKey[item.key] || [];
    const uniqueThaiFunds = [...new Map(thaiFundsForRow.map(f => [f.key, f])).values()];
    const thaiFragments = uniqueThaiFunds.length
      ? uniqueThaiFunds.map(fund => {
          const colorIdx = State.highlights[fund.key];
          return {
            text: fund.code,
            strong: true,
            color: '#334155',
            bg: colorIdx !== undefined ? (HL_COLORS[colorIdx]?.bg || '') : '',
          };
        })
      : [{ text: '-', color: '#475569' }];
    return {
      cells: [
        { text: item.name, bg: baseRowBg, color: '#35507a', weight: 2.4, align: 'left', strong: true },
        { text: item.currency || '-', bg: baseRowBg, color: '#475569', weight: 1.15 },
        {
          text: uniqueThaiFunds.length ? uniqueThaiFunds.map(f => f.code).join(', ') : '-',
          bg: baseRowBg,
          color: '#475569',
          weight: 1.9,
          align: 'left',
          fragments: thaiFragments,
        },
        ...metricKeys.map(key => ({
          text: formatReturnDisplay(get(item.row, CI[key])) || '-',
          bg: baseRowBg,
          color: '#475569',
          weight: 0.82,
        })),
        ...metricKeys.map(key => {
          const rankValue = rankMap[item.key]?.[key] ?? '';
          const colors = extractInlineColors(rankCellStyle(rankValue, rankTotals[key]));
          return {
            text: String(rankValue || '-'),
            bg: colors.bg || baseRowBg,
            color: colors.color || '#334155',
            weight: 0.82,
            strong: true,
          };
        }),
      ],
    };
  });

  const baseSort = (key, label, weight, align = 'center') => {
    const theme = presentationSortTheme(sortState?.key === key, sortState?.dir, '#365a88', '#ffffff');
    return { key, label: `${label}${theme.suffix}`, weight, align, bg: theme.bg, color: theme.color };
  };
  const metricLabel = key => key === 'rytd' ? 'YTD' : key.slice(1).toUpperCase();

  const columns = [
    buildPresentationSortColumn(sortState, 'name', 'Master Fund', 2.4, { align: 'left' }),
    buildPresentationSortColumn(sortState, 'currency', 'Base Currency', 1.15),
    buildPresentationSortColumn(sortState, 'thai', 'กองทุนในไทย', 1.9, { align: 'left' }),
    ...metricKeys.map(key => {
      const label = metricLabel(key);
      return buildPresentationSortColumn(sortState, key, label, 0.82, returnTheme);
    }),
    ...metricKeys.map(key => {
      const sortKey = `rank${key.slice(1)}`;
      const label = metricLabel(key);
      return buildPresentationSortColumn(sortState, sortKey, label, 0.82, rankTheme);
    }),
  ];

  return buildPresentationTablePayload({
    presetKey: 'masterAnnualizedV2',
    title: CONFIG.PAGES['master-annualized-v2']?.title || 'Master Fund Annualized V2',
    source: CONFIG.PAGES['master-annualized-v2']?.source || 'AVP Master Fund ID',
    headerGroups: [
      { label: '', span: 3, bg: '#dbe4f0', color: '#334155' },
      { label: 'ผลตอบแทน (%)', span: metricKeys.length, ...returnTheme },
      { label: 'อันดับในกลุ่มที่แสดง', span: metricKeys.length, ...rankTheme },
    ],
    columns,
    rows,
  });
}

function buildMasterCalendarExportPayload(sorted, yearKeys, linksByRowKey, CI, get, sortState) {
  const returnTheme = { bg: '#4ba3dc', color: '#ffffff' };
  const preset = getPresentationTablePreset('masterCalendar');
  const rows = sorted.map((item, rowIndex) => {
    const baseRowBg = rowIndex % 2 === 0 ? '#f8fbff' : '#eef4fb';
    const thaiFundsForRow = linksByRowKey[item.key] || [];
    const uniqueThaiFunds = [...new Map(thaiFundsForRow.map(f => [f.key, f])).values()];
    const thaiFragments = uniqueThaiFunds.length
      ? uniqueThaiFunds.map(fund => {
          const colorIdx = State.highlights[fund.key];
          return {
            text: fund.code,
            strong: true,
            color: '#334155',
            bg: colorIdx !== undefined ? (HL_COLORS[colorIdx]?.bg || '') : '',
          };
        })
      : [{ text: '-', color: '#475569' }];
    return {
      cells: [
        { text: item.name, bg: baseRowBg, color: '#35507a', weight: 2.4, align: 'left', strong: true },
        { text: item.currency || '-', bg: baseRowBg, color: '#475569', weight: 1.15 },
        {
          text: uniqueThaiFunds.length ? uniqueThaiFunds.map(f => f.code).join(', ') : '-',
          bg: baseRowBg,
          color: '#475569',
          weight: 1.9,
          align: 'left',
          fragments: thaiFragments,
        },
        ...yearKeys.map(year => ({
          text: formatReturnDisplay(get(item.row, CI[`ret${year}`])) || '-',
          bg: baseRowBg,
          color: '#475569',
          weight: 0.82,
        })),
      ],
    };
  });

  const columns = [
    buildPresentationSortColumn(sortState, 'name', 'Master Fund', 2.4, { align: 'left', widthPx: preset.columnWidthsPx?.name }),
    buildPresentationSortColumn(sortState, 'currency', 'Base Currency', 1.15, { widthPx: preset.columnWidthsPx?.currency }),
    buildPresentationSortColumn(sortState, 'thai', 'กองทุนในไทย', 1.9, { align: 'left', widthPx: preset.columnWidthsPx?.thai }),
    ...yearKeys.map(year => buildPresentationSortColumn(sortState, `ret-${year}`, year, 0.82, {
      ...returnTheme,
      widthPx: preset.columnWidthsPx?.metric,
    })),
  ];

  return buildPresentationTablePayload({
    presetKey: 'masterCalendar',
    title: CONFIG.PAGES['master-calendar']?.title || 'Master Fund Calendar Year',
    source: CONFIG.PAGES['master-calendar']?.source || 'AVP Master Fund ID',
    headerGroups: [
      { label: '', span: 3, bg: '#dbe4f0', color: '#334155' },
      { label: 'Return (Cumulative) (%)', span: yearKeys.length, ...returnTheme },
    ],
    columns,
    rows,
  });
}

function feeCombinedStyle(value, maxValue) {
  const n = parseNum(value);
  if (Number.isNaN(n) || !maxValue) return { bg: '', color: '' };
  const ratio = Math.max(0, Math.min(1, 1 - (n / maxValue)));
  const light = { r: 232, g: 247, b: 236 };
  const dark = { r: 126, g: 193, b: 133 };
  const mix = (a, b) => Math.round(a + (b - a) * ratio);
  return {
    bg: `rgb(${mix(light.r, dark.r)}, ${mix(light.g, dark.g)}, ${mix(light.b, dark.b)})`,
    color: '#183b22',
  };
}

function buildFeeV2ExportPayload(title, source, feeRows) {
  const maxCombined = Math.max(...feeRows.map(row => row.combined || 0), 0);
  const preset = getPresentationTablePreset('masterFees2');
  return buildPresentationTablePayload({
    presetKey: 'masterFees2',
    title,
    source,
    headerGroups: [
      { label: '', span: 2, bg: '#1f3f74', color: '#ffffff' },
      { label: 'TER (%) Q1-2026', span: 4, bg: '#1f3f74', color: '#ffffff' },
    ],
    columns: [
      { key: 'master', label: 'Master Fund', weight: 2.9, widthPx: preset.columnWidthsPx?.master, align: 'left', bg: '#1f3f74', color: '#ffffff' },
      { key: 'thai', label: 'กองไทย', weight: 1.55, widthPx: preset.columnWidthsPx?.thai, align: 'center', bg: '#1f3f74', color: '#ffffff' },
      { key: 'masterTer', label: 'Master Fund', weight: 0.95, widthPx: preset.columnWidthsPx?.masterTer, bg: '#1f3f74', color: '#ffffff' },
      { key: 'thaiTer', label: 'กองไทย (Sec)', weight: 1.02, widthPx: preset.columnWidthsPx?.thaiTer, bg: '#1f3f74', color: '#ffffff' },
      { key: 'date', label: 'Date', weight: 1.05, widthPx: preset.columnWidthsPx?.date, bg: '#1f3f74', color: '#ffffff' },
      { key: 'combined', label: 'Combined TER', weight: 1.08, widthPx: preset.columnWidthsPx?.combined, bg: '#1f3f74', color: '#ffffff' },
    ],
    rows: feeRows.map((row, rowIndex) => {
      const baseRowBg = rowIndex % 2 === 0 ? '#f4f6fa' : '#eceff4';
      const combinedStyle = feeCombinedStyle(row.combined, maxCombined);
      const highlightStyle = row.highlightColor ? { bg: row.highlightColor, color: '#173566' } : { bg: baseRowBg, color: '#334155' };
      return {
        cells: [
          { text: row.masterName, bg: baseRowBg, color: '#334155', weight: 2.9, align: 'left' },
          { text: row.thaiCode, bg: highlightStyle.bg, color: highlightStyle.color, weight: 1.55, strong: true },
          { text: row.masterTerText || '-', bg: baseRowBg, color: '#334155', weight: 0.95 },
          { text: row.thaiTerText || '-', bg: baseRowBg, color: '#334155', weight: 1.02 },
          { text: row.feeDate || '-', bg: baseRowBg, color: '#334155', weight: 1.05 },
          { text: row.combinedText || '-', bg: combinedStyle.bg || baseRowBg, color: combinedStyle.color || '#183b22', weight: 1.08, strong: true },
        ],
      };
    }),
  });
}

function buildThaiAnnualizedTablePayload(pageKey, view, viewCfg, sorted, metricKeys, helpers) {
  const metricLabels = metricKeys.map(key => key === 'rytd' ? 'YTD' : key.slice(1).toUpperCase());
  const rows = sorted.map((fund, rowIndex) => {
    const baseRowBg = rowIndex % 2 === 0 ? '#f8fbff' : '#eef4fb';
    const highlightIdx = State.highlights[fund.key];
    const highlightBg = highlightIdx !== undefined ? (HL_COLORS[highlightIdx]?.bg || baseRowBg) : baseRowBg;
    const cells = [
      { text: fund.code, bg: highlightBg, color: '#35507a', weight: 2.2, align: 'left', strong: true },
      { text: fund.type || '-', bg: baseRowBg, color: '#475569', weight: 1.05 },
      { text: fund.dividend || '-', bg: baseRowBg, color: '#475569', weight: 1.3 },
      ...metricKeys.map(key => ({
        text: helpers.formatReturn(helpers.get(fund.row, helpers.col[key])) || '-',
        bg: baseRowBg,
        color: '#475569',
        weight: 0.78,
      })),
      ...metricKeys.map(key => {
        const styleText = view === 'rank'
          ? rankCellStyle(helpers.sortable(fund, `rank${key.slice(1)}`), helpers.rankTotals[key])
          : percentileHeatStyle(helpers.get(fund.row, helpers.col[`p${key.slice(1)}`]));
        const colors = extractInlineColors(styleText);
        const text = view === 'rank'
          ? (helpers.sortable(fund, `rank${key.slice(1)}`) || '-')
          : (formatPercentileDisplay(helpers.get(fund.row, helpers.col[`p${key.slice(1)}`])) || '-');
        return {
          text,
          bg: colors.bg || '#ffffff',
          color: colors.color || '#334155',
          weight: 0.78,
          strong: true,
        };
      }),
    ];
    return { cells };
  });

  return {
    kind: 'annualized-table',
    rowsPerSlide: 18,
    title: CONFIG.PAGES[pageKey]?.title || 'Thai Annualized Report',
    subtitle: viewCfg.groupTitle,
    source: CONFIG.PAGES['select-fund']?.source || 'Percentrank Freestyle',
    headerGroups: [
      { label: '', span: 3, bg: '#dbe4f0', color: '#334155' },
      { label: 'ผลตอบแทน (%)', span: 6, bg: '#6198cb', color: '#ffffff' },
      { label: viewCfg.groupTitle, span: 6, bg: '#3f5d8c', color: '#ffffff' },
    ],
    columns: [
      { key: 'code', label: 'ชื่อกอง', weight: 2.2, align: 'left' },
      { key: 'type', label: 'ประเภท', weight: 1.05 },
      { key: 'dividend', label: 'Dividend', weight: 1.3 },
      ...metricLabels.map(label => ({ key: `return-${label}`, label, weight: 0.78 })),
      ...metricLabels.map(label => ({ key: `metric-${label}`, label, weight: 0.78 })),
    ],
    rows,
  };
}

function getThaiAnnualizedViewConfig(view) {
  if (view === 'rank') {
    return {
      loading: 'กำลังโหลดรายงาน Annualized Rank...',
      groupTitle: 'อันดับในกลุ่มที่แสดง',
      sortKeyForMetric: key => `rank${key.slice(1)}`,
      renderMetricCell: (fund, key, helpers) => {
        const rankKey = `rank${key.slice(1)}`;
        const value = helpers.sortable(fund, rankKey);
        return `<td class="report-num report-rank-cell" style="${rankCellStyle(value, helpers.rankTotals[key])}">${esc(value || '-')}</td>`;
      },
    };
  }

  return {
    loading: 'กำลังโหลดรายงาน Annualized...',
    groupTitle: 'Percentile Rank (%)',
    sortKeyForMetric: key => `p${key.slice(1)}`,
    renderMetricCell: (fund, key, helpers) => {
      const value = helpers.get(fund.row, helpers.col[`p${key.slice(1)}`]);
      return `<td class="report-num report-heat" style="${percentileHeatStyle(value)}">${esc(formatPercentileDisplay(value) || '-')}</td>`;
    },
  };
}

function getThaiAnnualizedMetricConfig(mode) {
  if (mode === 'rank') {
    return {
      mode: 'rank',
      groupTitle: 'อันดับในกลุ่มที่แสดง',
      groupClass: 'group-navy',
      sortKeyForMetric: key => `rank${key.slice(1)}`,
      renderMetricCell: (fund, key, helpers) => {
        const rankKey = `rank${key.slice(1)}`;
        const value = helpers.sortable(fund, rankKey);
        return `<td class="report-num report-rank-cell" style="${rankCellStyle(value, helpers.rankTotals[key])}">${esc(value || '-')}</td>`;
      },
      tableValue: (fund, key, helpers) => helpers.sortable(fund, `rank${key.slice(1)}`) || '-',
    };
  }

  if (mode === 'pct') {
    return {
      mode: 'pct',
      groupTitle: 'Percentile Rank (%)',
      groupClass: 'group-red',
      sortKeyForMetric: key => `p${key.slice(1)}`,
      renderMetricCell: (fund, key, helpers) => {
        const value = helpers.get(fund.row, helpers.col[`p${key.slice(1)}`]);
        return `<td class="report-num report-heat" style="${percentileHeatStyle(value)}">${esc(formatPercentileDisplay(value) || '-')}</td>`;
      },
      tableValue: (fund, key, helpers) => formatPercentileDisplay(helpers.get(fund.row, helpers.col[`p${key.slice(1)}`])) || '-',
    };
  }

  return {
    mode: 'return',
    groupTitle: 'ผลตอบแทน (%)',
    groupClass: 'group-blue',
    sortKeyForMetric: key => key,
    renderMetricCell: (fund, key, helpers) => {
      const value = formatReturnDisplay(helpers.get(fund.row, helpers.col[key]));
      return `<td class="report-num">${esc(value || '-')}</td>`;
    },
    tableValue: (fund, key, helpers) => formatReturnDisplay(helpers.get(fund.row, helpers.col[key])) || '-',
  };
}

function syncThaiAnnualizedViewSort(pageKey, nextView) {
  const sortState = State.reportSorts[pageKey];
  if (!sortState?.key) return;

  const keyMap = {
    p3m: 'rank3m',
    p6m: 'rank6m',
    pytd: 'rankytd',
    p1y: 'rank1y',
    p3y: 'rank3y',
    p5y: 'rank5y',
    rank3m: 'p3m',
    rank6m: 'p6m',
    rankytd: 'pytd',
    rank1y: 'p1y',
    rank3y: 'p3y',
    rank5y: 'p5y',
  };

  if (nextView === 'rank' && /^p/.test(sortState.key)) {
    sortState.key = keyMap[sortState.key] || '';
  } else if (nextView === 'return' && /^rank/.test(sortState.key)) {
    sortState.key = keyMap[sortState.key] || '';
  }
}

async function renderThaiAnnualizedReport(area, pageKey, view = 'return', showToggle = false) {
  const viewCfg = getThaiAnnualizedViewConfig(view);
  setLoading(area, viewCfg.loading);

  let rawRows;
  try {
    rawRows = await fetchCached('select-fund');
  } catch (e) {
    setError(area, e.message, pageKey);
    return;
  }

  const headers = rawRows[0] || [];
  const funds = buildPercentrankFunds(rawRows);
  const selected = State.selectedKeys.size > 0
    ? funds.filter(f => State.selectedKeys.has(f.key))
    : funds;

  const col = {
    r3m: findColumnIndex(headers, ['3 Month Return %']),
    r6m: findColumnIndex(headers, ['6 Month Return %']),
    rytd: findColumnIndex(headers, ['YTD Return %']),
    r1y: findColumnIndex(headers, ['1 Yr Anlsd %']),
    r3y: findColumnIndex(headers, ['3 Yr Anlsd %']),
    r5y: findColumnIndex(headers, ['5 Yr Anlsd %']),
    p3m: findColumnIndex(headers, ['3M']),
    p6m: findColumnIndex(headers, ['6M']),
    pytd: findColumnIndex(headers, ['YTD']),
    p1y: findColumnIndex(headers, ['1Y']),
    p3y: findColumnIndex(headers, ['3Y']),
    p5y: findColumnIndex(headers, ['5Y']),
  };
  const metricKeys = ['r3m', 'r6m', 'rytd', 'r1y', 'r3y', 'r5y'];
  const get = (row, i) => i >= 0 ? String(row[i] ?? '').trim() : '';
  const { ranks: rankMap, totals: rankTotals } = buildMetricRanks(selected, metricKeys, (fund, key) => get(fund.row, col[key]));
  const sortState = State.reportSorts[pageKey] || (State.reportSorts[pageKey] = { key: '', dir: 'asc' });
  const sortable = (fund, key) => {
    const mapping = {
      code: fund.code,
      type: fund.type,
      dividend: fund.dividend,
      r3m: get(fund.row, col.r3m),
      r6m: get(fund.row, col.r6m),
      rytd: get(fund.row, col.rytd),
      r1y: get(fund.row, col.r1y),
      r3y: get(fund.row, col.r3y),
      r5y: get(fund.row, col.r5y),
      p3m: get(fund.row, col.p3m),
      p6m: get(fund.row, col.p6m),
      pytd: get(fund.row, col.pytd),
      p1y: get(fund.row, col.p1y),
      p3y: get(fund.row, col.p3y),
      p5y: get(fund.row, col.p5y),
      rank3m: rankMap[fund.code]?.r3m ?? '',
      rank6m: rankMap[fund.code]?.r6m ?? '',
      rankytd: rankMap[fund.code]?.rytd ?? '',
      rank1y: rankMap[fund.code]?.r1y ?? '',
      rank3y: rankMap[fund.code]?.r3y ?? '',
      rank5y: rankMap[fund.code]?.r5y ?? '',
    };
    return mapping[key];
  };
  const sorted = sortState.key
    ? [...selected].sort((a, b) => compareValues(sortable(a, sortState.key), sortable(b, sortState.key), sortState.dir))
    : selected;
  const fitEnabled = false;

  const isDualMetricV2 = showToggle && pageKey === 'thai-annualized-v2';
  const leftMode = isDualMetricV2
    ? (State.reportOptions['thai-annualized-v2-left'] || 'return')
    : 'return';
  const rightMode = isDualMetricV2
    ? (State.reportOptions['thai-annualized-v2-right'] || 'rank')
    : (view === 'rank' ? 'rank' : 'pct');
  const leftCfg = getThaiAnnualizedMetricConfig(leftMode);
  const rightCfg = getThaiAnnualizedMetricConfig(rightMode);

  const toggleActions = isDualMetricV2 ? `
    <div class="metric-toggle-stack">
      <div class="metric-toggle-group">
        <span class="metric-toggle-label">ฝั่งซ้าย</span>
        <div class="view-toggle" role="tablist" aria-label="เลือกข้อมูลฝั่งซ้าย">
          <button class="btn btn-ghost view-toggle-btn ${leftMode === 'return' ? 'is-active' : ''}" type="button" data-annualized-side="left" data-annualized-mode="return">Return</button>
          <button class="btn btn-ghost view-toggle-btn ${leftMode === 'pct' ? 'is-active' : ''}" type="button" data-annualized-side="left" data-annualized-mode="pct">Percentile</button>
          <button class="btn btn-ghost view-toggle-btn ${leftMode === 'rank' ? 'is-active' : ''}" type="button" data-annualized-side="left" data-annualized-mode="rank">Rank</button>
        </div>
      </div>
      <div class="metric-toggle-group">
        <span class="metric-toggle-label">ฝั่งขวา</span>
        <div class="view-toggle" role="tablist" aria-label="เลือกข้อมูลฝั่งขวา">
          <button class="btn btn-ghost view-toggle-btn ${rightMode === 'return' ? 'is-active' : ''}" type="button" data-annualized-side="right" data-annualized-mode="return">Return</button>
          <button class="btn btn-ghost view-toggle-btn ${rightMode === 'pct' ? 'is-active' : ''}" type="button" data-annualized-side="right" data-annualized-mode="pct">Percentile</button>
          <button class="btn btn-ghost view-toggle-btn ${rightMode === 'rank' ? 'is-active' : ''}" type="button" data-annualized-side="right" data-annualized-mode="rank">Rank</button>
        </div>
      </div>
    </div>` : (showToggle ? `
    <div class="metric-toggle-stack">
      <div class="view-toggle" role="tablist" aria-label="เลือกมุมมอง Annualized">
        <button class="btn btn-ghost view-toggle-btn ${view === 'return' ? 'is-active' : ''}" type="button" data-annualized-view="return">Return</button>
        <button class="btn btn-ghost view-toggle-btn ${view === 'rank' ? 'is-active' : ''}" type="button" data-annualized-view="rank">Rank</button>
      </div>
    </div>` : '');

  const body = sorted.map(f => {
    const highlight = State.highlights[f.key];
    const codeStyle = highlight !== undefined ? ` style="background:${HL_COLORS[highlight].bg};"` : '';
    return `
      <tr>
        <td class="report-name"${codeStyle}>${esc(f.code)}</td>
        <td>${esc(f.type || '-')}</td>
        <td>${esc(f.dividend || '-')}</td>
        ${metricKeys.map(key => leftCfg.renderMetricCell(f, key, { col, get, rankTotals, sortable })).join('')}
        ${metricKeys.map(key => rightCfg.renderMetricCell(f, key, { col, get, rankTotals, sortable })).join('')}
      </tr>`;
  }).join('');

  area.innerHTML = `
    ${pageToolActions(pageKey, CONFIG.PAGES['select-fund']?.source || 'Percentrank Freestyle', toggleActions)}
    <div class="card report-card report-card-annualized" id="report-card">
      <table class="annualized-report">
        <thead>
          <tr class="report-group-row">
            <th colspan="3" class="group-blank"></th>
            <th colspan="6" class="${leftCfg.groupClass}">${leftCfg.groupTitle}</th>
            <th colspan="6" class="${rightCfg.groupClass}">${rightCfg.groupTitle}</th>
          </tr>
          <tr>
            <th class="report-sort ${sortState.key === 'code' ? 'is-active' : ''}" data-report-sort="code">${renderSortLabel('ชื่อกอง', sortState.key === 'code', sortState.dir)}</th>
            <th class="report-sort ${sortState.key === 'type' ? 'is-active' : ''}" data-report-sort="type">${renderSortLabel('ประเภท', sortState.key === 'type', sortState.dir)}</th>
            <th class="report-sort ${sortState.key === 'dividend' ? 'is-active' : ''}" data-report-sort="dividend">${renderSortLabel('Dividend', sortState.key === 'dividend', sortState.dir)}</th>
            ${metricKeys.map(key => {
              const label = key === 'rytd' ? 'YTD' : key.slice(1).toUpperCase();
              const sortKey = leftCfg.sortKeyForMetric(key);
              return `<th class="report-sort ${sortState.key === sortKey ? 'is-active' : ''}" data-report-sort="${sortKey}">${renderSortLabel(label, sortState.key === sortKey, sortState.dir)}</th>`;
            }).join('')}
            ${metricKeys.map(key => {
              const sortKey = rightCfg.sortKeyForMetric(key);
              const label = key === 'rytd' ? 'YTD' : key.slice(1).toUpperCase();
              return `<th class="report-sort ${sortState.key === sortKey ? 'is-active' : ''}" data-report-sort="${sortKey}">${renderSortLabel(label, sortState.key === sortKey, sortState.dir)}</th>`;
            }).join('')}
          </tr>
        </thead>
        <tbody>${body}</tbody>
      </table>
    </div>`;

  $$('[data-annualized-view]', area).forEach(el => {
    el.addEventListener('click', () => {
      const nextView = el.dataset.annualizedView;
      if (!nextView || nextView === State.reportOptions['thai-annualized-v2-view']) return;
      syncThaiAnnualizedViewSort(pageKey, nextView);
      State.reportOptions['thai-annualized-v2-view'] = nextView;
      renderThaiAnnualizedReport(area, pageKey, nextView, true);
    });
  });

  $$('[data-annualized-mode]', area).forEach(el => {
    el.addEventListener('click', () => {
      const side = el.dataset.annualizedSide;
      const nextMode = el.dataset.annualizedMode;
      if (!side || !nextMode) return;
      const stateKey = side === 'left' ? 'thai-annualized-v2-left' : 'thai-annualized-v2-right';
      if (State.reportOptions[stateKey] === nextMode) return;
      State.reportOptions[stateKey] = nextMode;
      sortState.key = '';
      sortState.dir = 'asc';
      renderThaiAnnualizedReport(area, pageKey, view, showToggle);
    });
  });

  $$('.report-sort', area).forEach(el => {
    el.addEventListener('click', () => {
      toggleNamedSort(sortState, el.dataset.reportSort);
      renderThaiAnnualizedReport(area, pageKey, view, showToggle);
    });
  });
  App._currentTableExport = () => {
    return buildThaiAnnualizedExportPayload(pageKey, sorted, metricKeys, leftCfg, rightCfg, {
      col,
      get,
      rankTotals,
      sortable,
    }, sortState);
  };
  bindPageImageActions(area, 'report-card', pageKey);
  App._currentExport = null;
}

function calendarRankNoStyle(rank) {
  const n = parseNum(rank);
  if (Number.isNaN(n)) return '';
  if (n === 1) return 'background:#7ABC81;color:#24364f;';
  if (n === 2) return 'background:#A8D086;color:#24364f;';
  if (n === 3) return 'background:#FCEC92;color:#5f4a08;';
  return '';
}

function compressYearRanges(years) {
  const nums = [...years]
    .map(y => parseInt(y, 10))
    .filter(n => !Number.isNaN(n))
    .sort((a, b) => a - b);
  if (!nums.length) return '';

  const ranges = [];
  let start = nums[0];
  let prev = nums[0];

  for (let i = 1; i < nums.length; i += 1) {
    const cur = nums[i];
    if (cur === prev + 1) {
      prev = cur;
      continue;
    }
    ranges.push(start === prev ? `${start}` : `${start}-${prev}`);
    start = cur;
    prev = cur;
  }
  ranges.push(start === prev ? `${start}` : `${start}-${prev}`);
  return ranges.join(', ');
}

function getThaiCalendarMetricConfig(mode) {
  if (mode === 'pct') {
    return {
      mode: 'pct',
      groupTitle: 'Percentile Rank (%)',
      groupClass: 'group-red',
      sortKeyForYear: year => `pct-${year}`,
      renderCell: (fund, year, helpers) => {
        const value = helpers.get(fund.row, helpers.rankPct[year]);
        return `<td class="report-num report-heat" style="${percentileHeatStyle(value)}">${esc(formatPercentileDisplay(value) || '-')}</td>`;
      },
      tableValue: (fund, year, helpers) => formatPercentileDisplay(helpers.get(fund.row, helpers.rankPct[year])) || '-',
      cellStyle: (fund, year, helpers) => extractInlineColors(percentileHeatStyle(helpers.get(fund.row, helpers.rankPct[year]))),
    };
  }

  if (mode === 'rank') {
    return {
      mode: 'rank',
      groupTitle: 'Rank No.',
      groupClass: 'group-navy',
      sortKeyForYear: year => `no-${year}`,
      renderCell: (fund, year, helpers) => {
        const value = helpers.getRankNo(fund, year);
        return `<td class="report-num report-rank-cell" style="${rankCellStyle(value, helpers.calendarRankTotals?.[year])}">${esc(value || '-')}</td>`;
      },
      tableValue: (fund, year, helpers) => String(helpers.getRankNo(fund, year) || '-'),
      cellStyle: (fund, year, helpers) => extractInlineColors(rankCellStyle(helpers.getRankNo(fund, year), helpers.calendarRankTotals?.[year])),
    };
  }

  return {
    mode: 'return',
    groupTitle: 'Calendar Year Return (%)',
    groupClass: 'group-blue',
    sortKeyForYear: year => `ret-${year}`,
    renderCell: (fund, year, helpers) => {
      const value = helpers.get(fund.row, helpers.returnCols[year]);
      return `<td class="report-num">${esc(formatReturnDisplay(value) || '-')}</td>`;
    },
    tableValue: (fund, year, helpers) => formatReturnDisplay(helpers.get(fund.row, helpers.returnCols[year])) || '-',
    cellStyle: () => ({ bg: '', color: '' }),
  };
}

function buildThaiCalendarExportPayload(sorted, visibleYears, leftCfg, rightCfg, helpers, sortState) {
  const leftTheme = annualizedGroupTheme(leftCfg);
  const rightTheme = annualizedGroupTheme(rightCfg);
  const rows = sorted.map((fund, rowIndex) => {
    const baseRowBg = rowIndex % 2 === 0 ? '#f8fbff' : '#eef4fb';
    const highlightIdx = State.highlights[fund.key];
    const highlightBg = highlightIdx !== undefined ? (HL_COLORS[highlightIdx]?.bg || baseRowBg) : baseRowBg;
    return {
      cells: [
        ...visibleYears.map(year => {
          const colors = leftCfg.cellStyle(fund, year, helpers);
          return {
            text: leftCfg.tableValue(fund, year, helpers),
            bg: colors.bg || baseRowBg,
            color: colors.color || '#475569',
            weight: 0.82,
            strong: leftCfg.mode !== 'return',
          };
        }),
        { text: fund.code, bg: highlightBg, color: '#35507a', weight: 2.2, align: 'center', strong: true },
        ...visibleYears.map(year => {
          const colors = rightCfg.cellStyle(fund, year, helpers);
          return {
            text: rightCfg.tableValue(fund, year, helpers),
            bg: colors.bg || baseRowBg,
            color: colors.color || '#475569',
            weight: 0.82,
            strong: rightCfg.mode !== 'return',
          };
        }),
      ],
    };
  });

  const columns = [
    ...visibleYears.map(year => buildPresentationSortColumn(
      sortState,
      leftCfg.sortKeyForYear(year),
      year,
      0.82,
      leftTheme
    )),
    buildPresentationSortColumn(sortState, 'code', 'Fund Code', 2.2, { align: 'center' }),
    ...visibleYears.map(year => buildPresentationSortColumn(
      sortState,
      rightCfg.sortKeyForYear(year),
      year,
      0.82,
      rightTheme
    )),
  ];

  return buildPresentationTablePayload({
    presetKey: 'thaiCalendar',
    title: CONFIG.PAGES['thai-calendar']?.title || 'Thai Calendar Year',
    source: CONFIG.PAGES['select-fund']?.source || 'Percentrank Freestyle',
    headerGroups: [
      { label: leftCfg.groupTitle, span: visibleYears.length, ...leftTheme },
      { label: '', span: 1, bg: '#dbe4f0', color: '#334155' },
      { label: rightCfg.groupTitle, span: visibleYears.length, ...rightTheme },
    ],
    columns,
    rows,
  });
}

function normalizeMasterMatchText(value) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/\bhlthcr\b/g, 'healthcare')
    .replace(/\bheath?lthcare\b/g, 'healthcare')
    .replace(/\bghc\b/g, 'global health care')
    .replace(/\bglb\b/g, 'global')
    .replace(/\bscn\b/g, 'science')
    .replace(/\bsciences\b/g, 'science')
    .replace(/\beq\b/g, 'equity')
    .replace(/\bportfolios?\b/g, 'portfolio')
    .replace(/\bintl\b/g, 'international')
    .replace(/\binnovtr\b/g, 'innovation')
    .replace(/\binnovt?r\b/g, 'innovation')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function masterMatchTokens(value) {
  const stop = new Set([
    'fund', 'class', 'usd', 'eur', 'acc', 'ac', 'portfolio', 'management',
    'port', 's', 'i', 'i2', 'h2', 'a', 'c', 'share', 'shares', 'the',
  ]);
  return normalizeMasterMatchText(value)
    .split(/\s+/)
    .filter(Boolean)
    .filter(token => !stop.has(token));
}

function scoreMasterMatch(masterName, candidateName) {
  const a = masterMatchTokens(masterName);
  const b = masterMatchTokens(candidateName);
  if (!a.length || !b.length) return 0;
  const bSet = new Set(b);
  const overlap = a.filter(token => bSet.has(token)).length;
  return overlap / Math.max(a.length, b.length);
}

function findBestMasterRow(masterRows, masterName) {
  let best = null;
  let bestScore = 0;
  masterRows.forEach(row => {
    const score = scoreMasterMatch(masterName, row.name);
    if (score > bestScore) {
      best = row;
      bestScore = score;
    }
  });
  return bestScore >= 0.45 ? best : null;
}

/* ============================================================
   TABLE BUILDER
   ============================================================ */

/*
 * buildTable(rows, opts)
 * rows[0] = header row
 * rows[1..] = data rows
 * opts: { selectable, selectedKeys }
 */
function buildTable(rows, opts = {}) {
  const {
    selectable = false,
    selectedKeys = new Set(),
    getRowMeta = null,
  } = opts;

  if (!rows || rows.length < 2) {
    return `
      <div class="state-box">
        <div class="state-icon">📭</div>
        <p>ไม่พบข้อมูล – ตรวจสอบชื่อ Tab ใน <code>js/config.js</code></p>
      </div>`;
  }

  const headers  = rows[0];
  const dataRows = rows.slice(1);

  /* Auto-detect numeric columns from first 30 rows */
  const numCols = new Set();
  dataRows.slice(0, 30).forEach(r =>
    r.forEach((v, i) => { if (isNum(v)) numCols.add(i); })
  );

  let html = '<div class="table-wrapper"><table>';

  /* ── Header ── */
  html += '<thead><tr>';
  if (selectable) html += '<th class="th-check"><input type="checkbox" id="chk-all" title="เลือกทั้งหมด"></th>';
  headers.forEach((h, i) => {
    html += `<th class="th-sortable ${State.sortCol === i ? (State.sortDir === 'asc' ? 'th-asc' : 'th-desc') : ''}" data-col="${i}">${renderSortLabel(h, State.sortCol === i, State.sortDir)}</th>`;
  });
  html += '</tr></thead>';

  /* ── Body ── */
  html += '<tbody>';
  dataRows.forEach((row, ri) => {
    const key    = esc(String(row[0] ?? ri));
    const meta   = getRowMeta?.(row, ri) || {};
    const selCls = (selectable && selectedKeys.has(String(row[0] ?? ri))) ? 'row-selected' : '';
    const rowCls = [selCls, meta.className || ''].filter(Boolean).join(' ');
    const rowStyle = meta.style ? ` style="${meta.style}"` : '';
    html += `<tr data-ri="${ri}" data-key="${key}" class="${rowCls}"${rowStyle}>`;
    if (selectable) {
      const chk = selectedKeys.has(String(row[0] ?? ri)) ? 'checked' : '';
      html += `<td class="td-check"><input type="checkbox" class="row-chk" data-key="${key}" ${chk}></td>`;
    }
    headers.forEach((_, ci) => {
      const v    = row[ci] ?? '';
      const isN  = numCols.has(ci) && isNum(v);
      const fv   = isN ? parseNum(v) : 0;
      const cls  = [
        isN ? 'td-num' : '',
        isN && fv > 0 ? 'td-positive' : '',
        isN && fv < 0 ? 'td-negative' : '',
      ].filter(Boolean).join(' ');
      html += `<td class="${cls}">${esc(v)}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table></div>';
  return html;
}

/* ── Sort rows (preserves header) ── */
function sortRows(rows, colIdx, dir) {
  const [hdr, ...body] = rows;
  body.sort((a, b) => {
    const av = a[colIdx] ?? '', bv = b[colIdx] ?? '';
    const an = parseNum(av), bn = parseNum(bv);
    if (!isNaN(an) && !isNaN(bn)) return dir === 'asc' ? an - bn : bn - an;
    return dir === 'asc'
      ? String(av).localeCompare(String(bv), 'th')
      : String(bv).localeCompare(String(av), 'th');
  });
  return [hdr, ...body];
}

/* ── Filter rows by search query ── */
function filterRows(rows, query) {
  if (!query) return rows;
  const q   = query.toLowerCase();
  const hdr = rows[0];
  const bod = rows.slice(1).filter(r =>
    r.some(v => String(v ?? '').toLowerCase().includes(q))
  );
  return [hdr, ...bod];
}

/* ── Bind table sort / checkbox interactions ── */
function bindTable(area, getRows, opts = {}) {
  const { selectable = false, onSelChange, getRowMeta = null } = opts;

  /* Sort headers */
  $$('thead th[data-col]', area).forEach(th => {
    th.addEventListener('click', () => {
      const ci = parseInt(th.dataset.col);
      if (State.sortCol === ci) {
        State.sortDir = State.sortDir === 'asc' ? 'desc' : 'asc';
      } else {
        State.sortCol = ci;
        State.sortDir = 'asc';
      }
      $$('thead th', area).forEach(t =>
        t.classList.remove('th-asc', 'th-desc')
      );
      th.classList.add(State.sortDir === 'asc' ? 'th-asc' : 'th-desc');

      /* Re-render table area */
      const rows    = getRows();
      const sorted  = sortRows(rows, State.sortCol, State.sortDir);
      const tArea   = $('#tbl-area', area);
      if (tArea) {
        tArea.innerHTML = buildTable(sorted, {
          selectable,
          selectedKeys: State.selectedKeys,
          getRowMeta,
        });
        bindCheckboxes(area, onSelChange);
      }
    });
  });

  if (selectable) bindCheckboxes(area, onSelChange);
}

function bindCheckboxes(area, onSelChange) {
  const chkAll = $('#chk-all', area);
  if (chkAll) {
    chkAll.addEventListener('change', () => {
      $$('.row-chk', area).forEach(c => {
        c.checked = chkAll.checked;
        const k = c.dataset.key;
        if (chkAll.checked) State.selectedKeys.add(k);
        else State.selectedKeys.delete(k);
        c.closest('tr').classList.toggle('row-selected', chkAll.checked);
      });
      onSelChange?.();
    });
  }
  $$('.row-chk', area).forEach(c => {
    c.addEventListener('change', () => {
      const k = c.dataset.key;
      if (c.checked) State.selectedKeys.add(k);
      else State.selectedKeys.delete(k);
      c.closest('tr').classList.toggle('row-selected', c.checked);
      onSelChange?.();
    });
  });
}

/* ============================================================
   EXPORT TO EXCEL
   ============================================================ */
function exportExcel(rows, filename = 'fund-data') {
  if (typeof XLSX === 'undefined') {
    toast('ไม่พบไลบรารี xlsx กรุณารอให้โหลดเสร็จแล้วลองใหม่', 'error');
    return;
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data');
  const date = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `${filename}_${date}.xlsx`);
  toast('ดาวน์โหลด Excel สำเร็จ', 'success');
}

function parseDmyDate(value) {
  const text = String(value || '').trim();
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (!m) return 0;
  const [, dd, mm, yyyy] = m;
  return Number(`${yyyy}${mm.padStart(2, '0')}${dd.padStart(2, '0')}`);
}

function toFixedSafe(value, digits = 2) {
  const n = parseNum(value);
  if (Number.isNaN(n)) return '';
  return n.toFixed(digits);
}

function clampRowsBySelection(items, limit = 18) {
  const selected = State.selectedKeys.size > 0
    ? items.filter(item => State.selectedKeys.has(item.key))
    : items;
  return selected.slice(0, limit);
}

function buildMasterRecords(rows) {
  const headers = rows[0] || [];
  const ci = {
    name: findColumnIndex(headers, ['Group/Investment']),
    fundId: findColumnIndex(headers, ['FundId']),
    isin: findColumnIndex(headers, ['ISIN']),
    currency: findColumnIndex(headers, ['Base Currency']),
    ongoingCost: findColumnIndex(headers, ['Ongoing Cost Actual']),
    ongoingCostDate: findColumnIndex(headers, ['Ongoing Cost Actual Date']),
    return3y: findColumnIndex(headers, ['Return(Annualized) 3Y']),
    sd3y: findColumnIndex(headers, ['Std Dev(Annualized) 3Y']),
    sharpe3y: findColumnIndex(headers, ['Sharpe Ratio(Annualized) 3Y']),
    drawdown3y: findColumnIndex(headers, ['Max Drawdown 3Y']),
    return5y: findColumnIndex(headers, ['Return(Annualized) 5Y']),
    sd5y: findColumnIndex(headers, ['Std Dev(Annualized) 5Y']),
    sharpe5y: findColumnIndex(headers, ['Sharpe Ratio(Annualized) 5Y']),
    drawdown5y: findColumnIndex(headers, ['Max Drawdown 5Y']),
  };
  const get = (row, index) => index >= 0 ? String(row[index] ?? '').trim() : '';
  return rows.slice(1).map(row => ({
    row,
    name: get(row, ci.name),
    fundId: get(row, ci.fundId),
    isin: get(row, ci.isin),
    currency: get(row, ci.currency),
    ongoingCost: get(row, ci.ongoingCost),
    ongoingCostDate: get(row, ci.ongoingCostDate),
    return3y: get(row, ci.return3y),
    sd3y: get(row, ci.sd3y),
    sharpe3y: get(row, ci.sharpe3y),
    drawdown3y: get(row, ci.drawdown3y),
    return5y: get(row, ci.return5y),
    sd5y: get(row, ci.sd5y),
    sharpe5y: get(row, ci.sharpe5y),
    drawdown5y: get(row, ci.drawdown5y),
  })).filter(item => item.name);
}

function pickBestMasterRecord(records) {
  return [...records].sort((a, b) => {
    const scoreA =
      (parseNum(a.return5y) === parseNum(a.return5y) ? 4 : 0) +
      (parseNum(a.sd5y) === parseNum(a.sd5y) ? 3 : 0) +
      (parseNum(a.ongoingCost) === parseNum(a.ongoingCost) ? 2 : 0) +
      (a.currency ? 1 : 0);
    const scoreB =
      (parseNum(b.return5y) === parseNum(b.return5y) ? 4 : 0) +
      (parseNum(b.sd5y) === parseNum(b.sd5y) ? 3 : 0) +
      (parseNum(b.ongoingCost) === parseNum(b.ongoingCost) ? 2 : 0) +
      (b.currency ? 1 : 0);
    return scoreB - scoreA;
  })[0] || null;
}

function buildRawSecLookup(rows) {
  const headers = rows[0] || [];
  const ci = {
    projName: findColumnIndex(headers, ['proj_abbr_name']),
    className: findColumnIndex(headers, ['fund_class_name']),
    fundName: findColumnIndex(headers, ['Fund Name']),
    ter: findColumnIndex(headers, ['TER']),
    front: findColumnIndex(headers, ['Front']),
    back: findColumnIndex(headers, ['Back']),
    date: findColumnIndex(headers, ['date']),
    asOfDate: findColumnIndex(headers, ['as_of_date']),
  };
  const get = (row, index) => index >= 0 ? String(row[index] ?? '').trim() : '';
  const map = new Map();
  rows.slice(1).forEach(row => {
    const ter = get(row, ci.ter);
    if (ter === '') return;
    const record = {
      projName: get(row, ci.projName),
      className: get(row, ci.className),
      fundName: get(row, ci.fundName),
      ter,
      front: get(row, ci.front),
      back: get(row, ci.back),
      date: get(row, ci.date),
      asOfDate: get(row, ci.asOfDate),
    };
    const score = parseDmyDate(record.asOfDate) || parseDmyDate(record.date);
    [record.className, record.fundName, record.projName]
      .map(v => String(v || '').trim().toUpperCase())
      .filter(Boolean)
      .forEach(key => {
        const existing = map.get(key);
        const existingScore = existing ? (parseDmyDate(existing.asOfDate) || parseDmyDate(existing.date)) : -1;
        if (!existing || score >= existingScore) {
          map.set(key, record);
        }
      });
  });
  return map;
}

async function buildSelectedMasterUniverse() {
  const selectRows = await fetchCached('select-fund');
  const catalog = buildSelectedFundsCatalog(selectRows);
  const allFunds = catalog
    .filter(f => f.code)
    .sort((a, b) => String(a.code).localeCompare(String(b.code), 'th'));
  const selectedFunds = State.selectedKeys.size > 0
    ? allFunds.filter(f => State.selectedKeys.has(f.key))
    : [];
  const masterRows = buildMasterRecords(await fetchCached('master-placeholder-2'));
  const byIsin = {};
  masterRows.forEach(record => {
    const key = String(record.isin || '').trim();
    if (!key) return;
    if (!byIsin[key]) byIsin[key] = [];
    byIsin[key].push(record);
  });

  const matchByIsin = (funds) => funds
    .filter(fund => {
      const isin = String(fund.masterId || '').trim();
      return isin && isin !== '-' && !!byIsin[isin]?.length;
    });

  const matchedSelectedFunds = matchByIsin(selectedFunds);
  const matchedAllFunds = matchByIsin(allFunds);
  const matchedFunds = matchedSelectedFunds.length
    ? matchedSelectedFunds
    : matchedAllFunds;

  const scopedFunds = State.selectedKeys.size > 0
    ? (matchedSelectedFunds.length ? matchedFunds : matchedFunds.slice(0, 18))
    : matchedFunds.slice(0, 18);

  return scopedFunds.map(fund => {
    const exact = byIsin[String(fund.masterId || '').trim()] || [];
    const master = exact.length ? pickBestMasterRecord(exact) : null;
    return { fund, master };
  }).filter(item => item.master);
}

function buildFeeComparisonRows(universe, rawLookup) {
  return universe.map(({ fund, master }) => {
    const raw = rawLookup.get(String(fund.code || '').trim().toUpperCase()) || null;
    const masterTer = parseNum(master.ongoingCost);
    const thaiTer = parseNum(raw?.ter);
    const combined = !Number.isNaN(masterTer) && !Number.isNaN(thaiTer) ? masterTer + thaiTer : NaN;
    return {
      thaiCode: fund.code,
      masterName: master.name,
      masterTer,
      thaiTer,
      combined,
      feeDate: raw?.date || master.ongoingCostDate || '',
      masterTerText: toFixedSafe(master.ongoingCost),
      thaiTerText: raw?.ter || '',
      combinedText: Number.isNaN(combined) ? '' : combined.toFixed(2),
      frontText: raw?.front || '',
      backText: raw?.back || '',
    };
  }).filter(item => !Number.isNaN(item.masterTer) && !Number.isNaN(item.thaiTer));
}

function buildScatterSvg(points, options = {}) {
  const width = options.width || 720;
  const height = options.height || 420;
  const padLeft = 72;
  const padRight = 26;
  const padTop = 28;
  const padBottom = 56;
  if (!points.length) return '<div class="state-box"><p>ไม่พบข้อมูลสำหรับสร้างกราฟ</p></div>';

  const xs = points.map(p => p.x);
  const ys = points.map(p => p.y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  const dx = maxX - minX || 1;
  const dy = maxY - minY || 1;
  const x0 = minX - (dx * 0.12);
  const x1 = maxX + (dx * 0.12);
  const y0 = minY - (dy * 0.18);
  const y1 = maxY + (dy * 0.12);
  const plotW = width - padLeft - padRight;
  const plotH = height - padTop - padBottom;
  const sx = value => padLeft + ((value - x0) / ((x1 - x0) || 1)) * plotW;
  const sy = value => padTop + plotH - ((value - y0) / ((y1 - y0) || 1)) * plotH;
  const ticks = 5;
  const xTicks = Array.from({ length: ticks }, (_, i) => x0 + (((x1 - x0) / (ticks - 1)) * i));
  const yTicks = Array.from({ length: ticks }, (_, i) => y0 + (((y1 - y0) / (ticks - 1)) * i));

  return `
    <svg viewBox="0 0 ${width} ${height}" class="insight-scatter-svg" role="img" aria-label="${esc(options.title || 'Scatter chart')}">
      <rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff" />
      ${yTicks.map(tick => `
        <line x1="${padLeft}" y1="${sy(tick)}" x2="${width - padRight}" y2="${sy(tick)}" stroke="#dbe4f0" stroke-width="1" />
        <text x="${padLeft - 10}" y="${sy(tick) + 4}" text-anchor="end" fill="#64748b" font-size="12">${tick.toFixed(2)}</text>
      `).join('')}
      ${xTicks.map(tick => `
        <line x1="${sx(tick)}" y1="${padTop}" x2="${sx(tick)}" y2="${height - padBottom}" stroke="#e7edf5" stroke-width="1" />
        <text x="${sx(tick)}" y="${height - padBottom + 20}" text-anchor="middle" fill="#64748b" font-size="12">${tick.toFixed(2)}</text>
      `).join('')}
      <line x1="${padLeft}" y1="${height - padBottom}" x2="${width - padRight}" y2="${height - padBottom}" stroke="#8ea3c2" stroke-width="1.2" />
      <line x1="${padLeft}" y1="${padTop}" x2="${padLeft}" y2="${height - padBottom}" stroke="#8ea3c2" stroke-width="1.2" />
      ${points.map(point => `
        <g>
          <circle cx="${sx(point.x)}" cy="${sy(point.y)}" r="${point.r || 5}" fill="${point.color || '#1a3c6e'}" opacity="0.95" />
          <line x1="${sx(point.x)}" y1="${sy(point.y)}" x2="${sx(point.x) + 20}" y2="${sy(point.y) - 16}" stroke="#94a3b8" stroke-width="1.2" />
          <text x="${sx(point.x) + 24}" y="${sy(point.y) - 18}" fill="#475569" font-size="12" font-weight="700">${esc(point.label)}</text>
        </g>
      `).join('')}
      <text x="${width / 2}" y="${height - 12}" text-anchor="middle" fill="#334155" font-size="13" font-weight="700">${esc(options.xLabel || '')}</text>
      <text x="18" y="${height / 2}" text-anchor="middle" fill="#334155" font-size="13" font-weight="700" transform="rotate(-90 18 ${height / 2})">${esc(options.yLabel || '')}</text>
    </svg>`;
}

function buildInsightSummaryCards(items) {
  return `
    <div class="insight-summary-grid">
      ${items.map(item => `
        <div class="insight-summary-card">
          <span class="insight-summary-label">${esc(item.label)}</span>
          <strong class="insight-summary-value">${esc(item.value)}</strong>
          <span class="insight-summary-note">${esc(item.note || '')}</span>
        </div>
      `).join('')}
    </div>`;
}

function buildInsightTable(rows, columns) {
  return `
    <div class="insight-table-wrap">
      <table class="insight-table">
        <thead>
          <tr>${columns.map(col => `<th>${esc(col.label)}</th>`).join('')}</tr>
        </thead>
        <tbody>
          ${rows.map(row => `
            <tr>${columns.map(col => `<td class="${col.className || ''}">${col.render ? col.render(row) : esc(String(row[col.key] ?? ''))}</td>`).join('')}</tr>
          `).join('')}
        </tbody>
      </table>
    </div>`;
}

/* ============================================================
   PAGES
   ============================================================ */
const Pages = {

  /* ── DASHBOARD ── */
  dashboard(area) {
    const labels  = ['กองทุนที่เลือกได้','กองทุนไทย Annualized Return V2','กองทุนไทย Calendar','Master Fund Annualized Return V2'];
    const classes = ['','c-accent','c-gold','c-success'];
    const pages   = ['select-fund','thai-annualized-v2','thai-calendar','master-annualized-v2'];

    area.innerHTML = `
      <div class="stats-grid" id="stats-grid">
        ${labels.map((lbl, i) => `
          <div class="stat-card ${classes[i]}" style="cursor:pointer" data-page="${pages[i]}">
            <div class="stat-label">${lbl}</div>
            <div class="stat-value" id="stat-${i}">–</div>
            <div class="stat-desc">รายการ</div>
          </div>`).join('')}
      </div>
      <div class="section-title">เข้าถึงข้อมูลได้เลย</div>
      <div class="quick-grid" id="quick-grid"></div>`;

    /* Stat card click → navigate */
    $$('.stat-card[data-page]', area).forEach(el =>
      el.addEventListener('click', () => App.navigate(el.dataset.page))
    );

    /* Quick links */
    const links = [
      { page: 'select-fund',       icon: quickIcon('check'), title: 'เลือกกองทุน',             sub: 'AVP Master Fund ID' },
      { page: 'thai-annualized-v2', icon: quickIcon('trend'), title: 'กองทุนไทย Annualized Return V2', sub: 'สลับดู Return และ Rank ได้' },
      { page: 'thai-calendar',     icon: quickIcon('cal'),    title: 'กองทุนไทย Calendar Year',  sub: 'AVP Thai Fund for Quality' },
      { page: 'master-annualized-v2', icon: quickIcon('globe'), title: 'Master Fund Annualized Return V2', sub: 'ISIN match + Base Currency' },
      { page: 'master-calendar',   icon: quickIcon('list'),   title: 'Master Fund Calendar Year',sub: 'AVP Master Fund ID' },
      { page: 'guide',             icon: quickIcon('book'),   title: 'คู่มือการใช้งาน',           sub: 'ขั้นตอนการตั้งค่าและใช้งาน' },
    ];
    $('#quick-grid', area).innerHTML = links.map(l => `
      <div class="quick-link" data-page="${l.page}">
        <div class="quick-link-icon">${l.icon}</div>
        <div class="quick-link-info">
          <h4>${l.title}</h4>
          <p>${l.sub}</p>
        </div>
      </div>`).join('');

    $$('.quick-link', area).forEach(el =>
      el.addEventListener('click', () => App.navigate(el.dataset.page))
    );

    /* Load counts lazily in background — one at a time to avoid memory spike */
    (async () => {
      const countPages = {
        'thai-annualized-v2': 'select-fund',
        'master-annualized-v2': 'master-annualized',
      };
      for (let i = 0; i < pages.length; i++) {
        try {
          const rows = await fetchCached(countPages[pages[i]] || pages[i]);
          const el   = $(`#stat-${i}`, area);
          if (el) el.textContent = Math.max(0, rows.length - 1).toLocaleString();
        } catch { /* ignore */ }
      }
    })();

    App._currentExport = null;
  },

  /* ── GENERIC TABLE ── */
  async genericTable(area, pageKey) {
    const cfg = CONFIG.PAGES[pageKey];
    setLoading(area, `กำลังโหลด ${cfg.title}...`);

    let rawRows;
    try {
      rawRows = await fetchCached(pageKey);
    } catch (e) {
      setError(area, e.message, pageKey);
      return;
    }

    State.sortCol = null;
    State.sortDir = 'asc';

    const render = (query = '', goPage = 1) => {
      State.tablePage = goPage;
      const headers = rawRows[0] || [];
      const codeIdx = findColumnIndex(headers, ['Fund Code', 'FundId']);
      const masterIdIdx = findColumnIndex(headers, ['Master FundId', 'FundId', 'ISIN']);
      const selectedMasterIds = getSelectedMasterIds();
      const shouldFilterBySelection = pageKey !== 'select-fund' && State.selectedKeys.size > 0 && (codeIdx !== -1 || masterIdIdx !== -1);

      const rowsAfterSelection = shouldFilterBySelection
        ? [
            headers,
            ...rawRows.slice(1).filter(row => {
              const rowCode = codeIdx >= 0 ? normalizeFundKey(row[codeIdx]) : '';
              const rowMasterId = masterIdIdx >= 0 ? String(row[masterIdIdx] ?? '').trim() : '';
              return State.selectedKeys.has(rowCode) || selectedMasterIds.has(rowMasterId);
            }),
          ]
        : rawRows;

      const filtered  = filterRows(rowsAfterSelection, query);
      const sorted    = State.sortCol !== null
        ? sortRows(filtered, State.sortCol, State.sortDir)
        : filtered;
      const totalData  = Math.max(0, sorted.length - 1);
      const totalPages = Math.max(1, Math.ceil(totalData / State.pageSize));
      const pg         = Math.min(Math.max(1, State.tablePage), totalPages);
      const startIdx   = (pg - 1) * State.pageSize + 1;
      const endIdx     = Math.min(startIdx + State.pageSize, sorted.length);
      const pageSlice  = [sorted[0], ...sorted.slice(startIdx, endIdx)];

      area.innerHTML = `
        ${pageToolActions(pageKey, cfg.source)}
        <div class="card" id="report-card">
          <div class="card-header">
            <span class="card-title">${esc(cfg.title)}</span>
            <div class="filter-bar">
              <div class="search-wrap">
                <span class="s-icon">${searchIcon()}</span>
                <input class="search-input" id="tbl-search" type="text"
                  placeholder="ค้นหา..." value="${esc(query)}" autocomplete="off">
              </div>
              ${pageKey !== 'select-fund' ? `
                <span class="row-count-badge ${State.selectedKeys.size > 0 ? 'is-info' : ''}">
                  ${State.selectedKeys.size > 0 ? `แสดงตามกองทุนที่เลือก ${State.selectedKeys.size} รายการ` : 'ยังไม่ได้จำกัดตามกองทุนที่เลือก'}
                </span>` : ''}
              <span class="row-count-badge">${totalData.toLocaleString()} รายการ</span>
              <span class="badge badge-primary">${esc(cfg.source)}</span>
              ${getPageDataSourceBadge(pageKey) ? `<span class="badge badge-data-origin">${esc(getPageDataSourceBadge(pageKey))}</span>` : ''}
            </div>
          </div>
          <div id="tbl-area">${buildTable(pageSlice, {
            getRowMeta: (row) => {
              const ci = getFundHighlightIndex(row, { codeIdx, masterIdIdx });
              if (ci === undefined) return {};
              return {
                className: 'row-highlighted',
                style: `background:${HL_COLORS[ci].bg}`,
              };
            },
          })}</div>
          ${totalPages > 1 ? `
          <div class="pagination-bar">
            <label class="page-size-wrap">แถวต่อหน้า :
              <select class="page-size-select" id="page-size">
                ${PAGE_SIZE_OPTIONS.map(size => `<option value="${size}" ${size === State.pageSize ? 'selected' : ''}>${size}</option>`).join('')}
              </select>
            </label>
            <button class="btn btn-ghost btn-sm" id="pg-prev" ${pg <= 1 ? 'disabled' : ''}>← ก่อนหน้า</button>
            <span class="pg-info">หน้า ${pg} / ${totalPages} &nbsp;(แสดง ${startIdx}–${Math.min(endIdx-1, totalData)} จาก ${totalData.toLocaleString()})</span>
            <button class="btn btn-ghost btn-sm" id="pg-next" ${pg >= totalPages ? 'disabled' : ''}>ถัดไป →</button>
          </div>` : ''}
        </div>`;

      const rowMeta = (row) => {
        const ci = getFundHighlightIndex(row, { codeIdx, masterIdIdx });
        if (ci === undefined) return {};
        return {
          className: 'row-highlighted',
          style: `background:${HL_COLORS[ci].bg}`,
        };
      };

      bindTable(area, () => {
        const q = $('#tbl-search', area)?.value.trim() ?? '';
        return filterRows(rowsAfterSelection, q);
      }, { getRowMeta: rowMeta });

      const inp = $('#tbl-search', area);
      let timer;
      inp.addEventListener('input', () => {
        clearTimeout(timer);
        timer = setTimeout(() => render(inp.value.trim(), 1), 280);
      });

      $('#pg-prev', area)?.addEventListener('click', () => render(inp?.value.trim() ?? '', pg - 1));
      $('#pg-next', area)?.addEventListener('click', () => render(inp?.value.trim() ?? '', pg + 1));
      $('#page-size', area)?.addEventListener('change', e => {
        State.pageSize = parseInt(e.target.value, 10) || 25;
        render(inp?.value.trim() ?? '', 1);
      });
      bindPageImageActions(area, 'report-card', pageKey);
    };

    render();
    App._currentExport = () => exportExcel(rawRows, cfg.title);
    App._currentTableExport = () => buildSimpleTablePayload(cfg.title, cfg.source || '', rawRows[0] || [], rawRows.slice(1));
  },

  /* ── SELECT FUND ── */
  async selectFund(area) {
    const cfg = CONFIG.PAGES['select-fund'];
    setLoading(area, 'กำลังโหลดรายการกองทุน...');

    let rawRows;
    try {
      rawRows = await fetchCached('select-fund');
    } catch (e) {
      setError(area, e.message, 'select-fund');
      return;
    }

    const headers = rawRows[0] || [];
    const CI = {
      CATEGORY: findColumnIndex(headers, ['AVP® Category', 'AVP®  Category', 'AVP Category']),
    };
    const allFunds = buildSelectedFundsCatalog(rawRows);

    State.selectedFunds = Object.fromEntries(allFunds.map(f => [f.key, f]));

    /* Unique dropdown options */
    const categories = [...new Set(allFunds.map(f => f.category).filter(Boolean))].sort();

    const opt = (val, label, cur) =>
      `<option value="${esc(val)}" ${cur === val ? 'selected' : ''}>${esc(label)}</option>`;

    const render = (goPage = 1, opts = {}) => {
      const preserveScroll = !!opts.preserveScroll;
      const prevScrollTop = preserveScroll ? area.scrollTop : 0;
      State.tablePage = goPage;
      State.pageSize = State.selectFundFilters.pageSize || PAGE_SIZE_OPTIONS[0];

      /* Filter */
      let visible = allFunds.filter(f => {
        if (State.selectFundFilters.query && !f.code.toLowerCase().includes(State.selectFundFilters.query) && !f.masterName.toLowerCase().includes(State.selectFundFilters.query)) return false;
        if (State.selectFundFilters.category && f.category !== State.selectFundFilters.category) return false;
        if (State.selectFundFilters.type  && f.type     !== State.selectFundFilters.type)  return false;
        if (State.selectFundFilters.style && f.style    !== State.selectFundFilters.style) return false;
        if (State.selectFundFilters.dividend   && f.dividend !== State.selectFundFilters.dividend)   return false;
        return true;
      });

      if (State.selectFundSort.key) {
        const sortableValue = (fund, key) => {
          if (key === 'highlight') {
            const idx = State.highlights[fund.key];
            return idx === undefined ? '' : HL_COLORS[idx]?.name || '';
          }
          return fund[key];
        };
        visible = [...visible].sort((a, b) =>
          compareValues(sortableValue(a, State.selectFundSort.key), sortableValue(b, State.selectFundSort.key), State.selectFundSort.dir)
        );
      }

      const total = visible.length;
      const totalPages = Math.max(1, Math.ceil(total / State.pageSize));
      const pg = Math.min(Math.max(1, State.tablePage), totalPages);
      const si = (pg - 1) * State.pageSize;
      const pageData = visible.slice(si, si + State.pageSize);

      /* Table rows */
      const allVisibleChecked = pageData.length > 0 && pageData.every(f => State.selectedKeys.has(f.key));
      const someVisibleChecked = pageData.some(f => State.selectedKeys.has(f.key));

      const tRows = pageData.map(f => {
        const isSelected = State.selectedKeys.has(f.key);
        return `
          <tr class="${isSelected ? 'row-selected' : ''}">
            <td class="td-check">
              <input type="checkbox" class="row-chk" data-key="${esc(f.key)}" ${isSelected ? 'checked' : ''}>
            </td>
            <td class="td-code">${esc(f.code)}</td>
            <td class="td-hl">${buildHighlightSelect(f.key, State.highlights[f.key])}</td>
            <td>${esc(f.dividend)}</td>
            <td>${esc(f.style)}</td>
            <td class="td-isin">${esc(f.masterId)}</td>
            <td>${esc(f.masterName)}</td>
          </tr>`;
      }).join('');

      area.innerHTML = `
        <div class="card sf-card">
          <div class="sf-filterbar">
            <div class="sf-search">
              <span class="s-icon">${searchIcon()}</span>
              <input class="search-input" id="sf-q" type="text"
                placeholder="ชื่อกองทุน / Fund Code..." value="${esc(State.selectFundFilters.query)}" autocomplete="off">
            </div>
            <div class="sf-drop">
              <div class="sf-droplabel">หมวดหมู่ AVP</div>
              <select class="sf-select" id="sf-category">
                ${opt('','ทั้งหมด',State.selectFundFilters.category)}
                ${categories.map(g => opt(g, g, State.selectFundFilters.category)).join('')}
              </select>
            </div>
            <div class="sf-drop">
              <div class="sf-droplabel">ประเภทกองทุน</div>
              <select class="sf-select" id="sf-type">
                ${['','General','SSF','RMF','LTF','TESGX'].map((v,i) => opt(v, i===0?'ทั้งหมด':v, State.selectFundFilters.type)).join('')}
              </select>
            </div>
            <div class="sf-drop">
              <div class="sf-droplabel">STYLE</div>
              <select class="sf-select" id="sf-style">
                ${['','Active','Passive'].map((v,i) => opt(v, i===0?'ทั้งหมด':v, State.selectFundFilters.style)).join('')}
              </select>
            </div>
            <div class="sf-drop">
              <div class="sf-droplabel">DIVIDEND</div>
              <select class="sf-select" id="sf-div">
                ${['','Dividend','No Dividend'].map((v,i) => opt(v, i===0?'ทั้งหมด':v, State.selectFundFilters.dividend)).join('')}
              </select>
            </div>
            <button class="btn btn-ghost btn-sm" id="sf-reset">↺ รีเซ็ต</button>
          </div>
          <div class="sf-meta">
            <span class="row-count-badge">${total.toLocaleString()} รายการ</span>
            <span class="row-count-badge is-info">เลือกแล้ว ${State.selectedKeys.size.toLocaleString()} กองทุน</span>
            <span class="badge badge-primary">${esc(cfg.source)}</span>
            ${getPageDataSourceBadge('select-fund') ? `<span class="badge badge-data-origin">${esc(getPageDataSourceBadge('select-fund'))}</span>` : ''}
            ${Object.keys(State.highlights).length > 0
              ? `<span class="badge badge-accent">ตั้งค่าสีไว้ ${Object.keys(State.highlights).length} กองทุน</span>`
              : ''}
            ${CI.CATEGORY === -1
              ? `<span class="badge badge-warning">ยังไม่พบคอลัมน์ AVP® Category</span>`
              : ''}
          </div>
          <div class="table-wrapper">
            <table>
              <thead><tr>
                <th class="th-check">
                  <input type="checkbox" id="sf-chk-all" title="เลือกทั้งหมดที่แสดง" ${allVisibleChecked ? 'checked' : ''} ${pageData.length === 0 ? 'disabled' : ''}>
                </th>
                <th class="sf-sort ${State.selectFundSort.key === 'code' ? 'is-active' : ''}" data-sort-key="code">${renderSortLabel('Fund Code', State.selectFundSort.key === 'code', State.selectFundSort.dir)}</th>
                <th class="sf-sort ${State.selectFundSort.key === 'highlight' ? 'is-active' : ''}" data-sort-key="highlight">${renderSortLabel('Highlight', State.selectFundSort.key === 'highlight', State.selectFundSort.dir)}</th>
                <th class="sf-sort ${State.selectFundSort.key === 'dividend' ? 'is-active' : ''}" data-sort-key="dividend">${renderSortLabel('Dividend', State.selectFundSort.key === 'dividend', State.selectFundSort.dir)}</th>
                <th class="sf-sort ${State.selectFundSort.key === 'style' ? 'is-active' : ''}" data-sort-key="style">${renderSortLabel('Style', State.selectFundSort.key === 'style', State.selectFundSort.dir)}</th>
                <th class="sf-sort ${State.selectFundSort.key === 'masterId' ? 'is-active' : ''}" data-sort-key="masterId">${renderSortLabel('ISIN', State.selectFundSort.key === 'masterId', State.selectFundSort.dir)}</th>
                <th class="sf-sort ${State.selectFundSort.key === 'masterName' ? 'is-active' : ''}" data-sort-key="masterName">${renderSortLabel('Master Fund', State.selectFundSort.key === 'masterName', State.selectFundSort.dir)}</th>
              </tr></thead>
              <tbody>${tRows}</tbody>
            </table>
          </div>
          ${totalPages > 1 ? `
          <div class="pagination-bar">
            <label class="page-size-wrap">แถวต่อหน้า :
              <select class="page-size-select" id="sf-page-size">
                ${PAGE_SIZE_OPTIONS.map(size => `<option value="${size}" ${size === State.pageSize ? 'selected' : ''}>${size}</option>`).join('')}
              </select>
            </label>
            <button class="btn btn-ghost btn-sm" id="pg-prev" ${pg<=1?'disabled':''}>← ก่อนหน้า</button>
            <span class="pg-info">หน้า ${pg} / ${totalPages} &nbsp;(${si+1}–${Math.min(si+State.pageSize,total)} จาก ${total.toLocaleString()})</span>
            <button class="btn btn-ghost btn-sm" id="pg-next" ${pg>=totalPages?'disabled':''}>ถัดไป →</button>
          </div>` : ''}
        </div>`;

      /* Bind search */
      const qEl = $('#sf-q', area);
      let timer;
      qEl.addEventListener('input', () => {
        clearTimeout(timer);
        qEl.value = qEl.value.toUpperCase();
        timer = setTimeout(() => {
          State.selectFundFilters.query = qEl.value.trim().toLowerCase();
          render(1);
        }, 280);
      });

      /* Bind dropdowns */
      $('#sf-category', area).addEventListener('change', e => { State.selectFundFilters.category = e.target.value; render(1); });
      $('#sf-type',  area).addEventListener('change', e => { State.selectFundFilters.type  = e.target.value; render(1); });
      $('#sf-style', area).addEventListener('change', e => { State.selectFundFilters.style = e.target.value; render(1); });
      $('#sf-div',   area).addEventListener('change', e => { State.selectFundFilters.dividend = e.target.value; render(1); });
      $$('.sf-sort', area).forEach(el => {
        el.addEventListener('click', () => {
          toggleNamedSort(State.selectFundSort, el.dataset.sortKey);
          render(1);
        });
      });

      /* Reset */
      $('#sf-reset', area).addEventListener('click', () => {
        State.selectFundFilters = {
          category: '',
          type: '',
          style: '',
          dividend: '',
          query: '',
          pageSize: 25,
        };
        State.selectFundSort = { key: '', dir: 'asc' };
        State.tablePage = 1;
        State.pageSize = 25;
        State.selectedKeys.clear();
        State.selectedFunds = {};
        State.highlights = {};
        render(1);
      });

      /* Selection */
      const chkAll = $('#sf-chk-all', area);
      if (chkAll) {
        chkAll.indeterminate = !allVisibleChecked && someVisibleChecked;
        chkAll.addEventListener('change', () => {
          pageData.forEach(f => {
            if (chkAll.checked) State.selectedKeys.add(f.key);
            else State.selectedKeys.delete(f.key);
          });
          render(pg, { preserveScroll: true });
        });
      }

      $$('.row-chk', area).forEach(el => {
        el.addEventListener('change', () => {
          const key = el.dataset.key;
          if (el.checked) State.selectedKeys.add(key);
          else State.selectedKeys.delete(key);
          render(pg, { preserveScroll: true });
        });
      });

      /* Highlight color selection */
      $$('.hl-select', area).forEach(el => {
        el.addEventListener('change', () => {
          const fund = el.dataset.fund;
          const rawValue = el.value;
          if (rawValue === '') delete State.highlights[fund];
          else State.highlights[fund] = parseInt(rawValue, 10);
          render(pg, { preserveScroll: true });
        });
      });

      /* Pagination */
      $('#pg-prev', area)?.addEventListener('click', () => render(pg - 1));
      $('#pg-next', area)?.addEventListener('click', () => render(pg + 1));
      $('#sf-page-size', area)?.addEventListener('change', e => {
        State.pageSize = parseInt(e.target.value, 10) || 25;
        State.selectFundFilters.pageSize = State.pageSize;
        render(1);
      });

      if (preserveScroll) {
        area.scrollTop = prevScrollTop;
      }
    };

    render();
    App._currentExport = () => {
      const selectedRows = allFunds
        .filter(f => State.selectedKeys.has(f.key))
        .map(f => [
          f.code,
          f.category,
          f.type,
          f.dividend,
          f.style,
          f.masterId,
          f.masterName,
          State.highlights[f.key] !== undefined ? HL_COLORS[State.highlights[f.key]].name : '',
        ]);

      if (selectedRows.length === 0) {
        toast('ยังไม่ได้เลือกกองทุนสำหรับ Export', 'warning');
        return;
      }

      exportExcel([
        ['Fund Code', 'AVP Category', 'Type', 'Dividend', 'Style', 'Master Fund ID', 'Master Fund Name', 'Highlight Color'],
        ...selectedRows,
      ], 'selected-funds');
    };
  },

  async thaiAnnualized(area) {
    return renderThaiAnnualizedReport(area, 'thai-annualized', 'return', false);
  },

  async thaiAnnualizedRank(area) {
    return renderThaiAnnualizedReport(area, 'thai-annualized-rank', 'rank', false);
  },

  async thaiAnnualizedV2(area) {
    const currentView = State.reportOptions['thai-annualized-v2-view'] === 'rank' ? 'rank' : 'return';
    return renderThaiAnnualizedReport(area, 'thai-annualized-v2', currentView, true);
  },

  async thaiCalendar(area) {
    setLoading(area, 'กำลังโหลดรายงาน Calendar Year...');

    let rawRows;
    try {
      rawRows = await fetchCached('select-fund');
    } catch (e) {
      setError(area, e.message, 'thai-calendar');
      return;
    }

    const headers = rawRows[0] || [];
    const funds = buildPercentrankFunds(rawRows);
    const selected = State.selectedKeys.size > 0
      ? funds.filter(f => State.selectedKeys.has(f.key))
      : funds;

    const allYears = ['2010','2011','2012','2013','2014','2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'];
    const returnCols = Object.fromEntries(allYears.map(y => [y, findColumnIndex(headers, [`Calendar Year Return ${y}`])]));
    const rankPct = Object.fromEntries(allYears.map(y => [y, findColumnIndex(headers, [`Rank % Calender Year ${y}`])]));
    const get = (row, i) => i >= 0 ? String(row[i] ?? '').trim() : '';
    const sortState = State.reportSorts['thai-calendar'];
    const selectedYears = (State.reportOptions['thai-calendar-years'] || []).filter(y => allYears.includes(y));
    const visibleYears = selectedYears.length ? selectedYears : ['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'];
    const { ranks: calendarRankMap, totals: calendarRankTotals } = buildMetricRanks(selected, visibleYears, (fund, year) => get(fund.row, returnCols[year]));
    const getRankNo = (fund, year) => calendarRankMap[fund.code]?.[year] ?? '';
    const leftMode = State.reportOptions['thai-calendar-left'] || 'return';
    const rightMode = State.reportOptions['thai-calendar-right'] || 'rank';
    const leftCfg = getThaiCalendarMetricConfig(leftMode);
    const rightCfg = getThaiCalendarMetricConfig(rightMode);
    const helpers = { get, getRankNo, rankPct, returnCols, calendarRankTotals };
    const sortableCalendar = (fund, key) => {
      if (key === 'code') return fund.code;
      if (key.startsWith('pct-')) return get(fund.row, rankPct[key.slice(4)]);
      if (key.startsWith('no-')) return getRankNo(fund, key.slice(3));
      if (key.startsWith('ret-')) return get(fund.row, returnCols[key.slice(4)]);
      return '';
    };
    const sorted = sortState.key
      ? [...selected].sort((a, b) => compareValues(sortableCalendar(a, sortState.key), sortableCalendar(b, sortState.key), sortState.dir))
      : selected;

    const yearChips = allYears.map(y => {
      const active = visibleYears.includes(y);
      return `<button class="btn btn-ghost year-chip ${active ? 'is-active' : ''}" data-calendar-year="${y}" title="แสดงหรือซ่อนปี ${y}">${y}</button>`;
    }).join('');

    const toggleActions = `
      <div class="metric-toggle-stack">
        <div class="metric-toggle-group">
          <span class="metric-toggle-label">ฝั่งซ้าย</span>
          <div class="view-toggle" role="tablist" aria-label="เลือกข้อมูลฝั่งซ้ายของ Calendar Year">
            <button class="btn btn-ghost view-toggle-btn ${leftMode === 'return' ? 'is-active' : ''}" type="button" data-calendar-side="left" data-calendar-mode="return">Return</button>
            <button class="btn btn-ghost view-toggle-btn ${leftMode === 'pct' ? 'is-active' : ''}" type="button" data-calendar-side="left" data-calendar-mode="pct">Percentile</button>
            <button class="btn btn-ghost view-toggle-btn ${leftMode === 'rank' ? 'is-active' : ''}" type="button" data-calendar-side="left" data-calendar-mode="rank">Rank No.</button>
          </div>
        </div>
        <div class="metric-toggle-group">
          <span class="metric-toggle-label">ฝั่งขวา</span>
          <div class="view-toggle" role="tablist" aria-label="เลือกข้อมูลฝั่งขวาของ Calendar Year">
            <button class="btn btn-ghost view-toggle-btn ${rightMode === 'return' ? 'is-active' : ''}" type="button" data-calendar-side="right" data-calendar-mode="return">Return</button>
            <button class="btn btn-ghost view-toggle-btn ${rightMode === 'pct' ? 'is-active' : ''}" type="button" data-calendar-side="right" data-calendar-mode="pct">Percentile</button>
            <button class="btn btn-ghost view-toggle-btn ${rightMode === 'rank' ? 'is-active' : ''}" type="button" data-calendar-side="right" data-calendar-mode="rank">Rank No.</button>
          </div>
        </div>
        <div class="year-chip-wrap">${yearChips}</div>
      </div>`;

    const body = sorted.map(f => {
      const highlight = State.highlights[f.key];
      const style = highlight !== undefined ? ` style="background:${HL_COLORS[highlight].bg};"` : '';
      return `
        <tr>
          ${visibleYears.map(year => leftCfg.renderCell(f, year, helpers)).join('')}
          <td class="calendar-code"${style}>${esc(f.code)}</td>
          ${visibleYears.map(year => rightCfg.renderCell(f, year, helpers)).join('')}
        </tr>`;
    }).join('');

    area.innerHTML = `
      ${pageToolActions(
        'thai-calendar',
        CONFIG.PAGES['select-fund']?.source || 'Percentrank Freestyle',
        toggleActions
      )}
      <div class="card report-card report-card-calendar" id="report-card">
        <table class="annualized-report calendar-v2-report">
          <thead>
            <tr class="report-group-row">
              <th colspan="${visibleYears.length}" class="${leftCfg.groupClass}">${leftCfg.groupTitle}</th>
              <th colspan="1" class="group-blank"></th>
              <th colspan="${visibleYears.length}" class="${rightCfg.groupClass}">${rightCfg.groupTitle}</th>
            </tr>
            <tr>
              ${visibleYears.map(y => `<th class="report-sort ${sortState.key === leftCfg.sortKeyForYear(y) ? 'is-active' : ''}" data-report-sort="${leftCfg.sortKeyForYear(y)}">${renderSortLabel(y, sortState.key === leftCfg.sortKeyForYear(y), sortState.dir)}</th>`).join('')}
              <th class="report-sort ${sortState.key === 'code' ? 'is-active' : ''}" data-report-sort="code">${renderSortLabel('Fund Code', sortState.key === 'code', sortState.dir)}</th>
              ${visibleYears.map(y => `<th class="report-sort ${sortState.key === rightCfg.sortKeyForYear(y) ? 'is-active' : ''}" data-report-sort="${rightCfg.sortKeyForYear(y)}">${renderSortLabel(y, sortState.key === rightCfg.sortKeyForYear(y), sortState.dir)}</th>`).join('')}
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>`;

    $$('[data-calendar-year]', area).forEach(el => {
      el.addEventListener('click', () => {
        const year = el.dataset.calendarYear;
        const current = new Set(State.reportOptions['thai-calendar-years'] || []);
        if (current.has(year)) current.delete(year);
        else current.add(year);
        const next = allYears.filter(y => current.has(y));
        State.reportOptions['thai-calendar-years'] = next.length ? next : ['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'];
        Pages.thaiCalendar(area);
      });
    });
    $$('[data-calendar-mode]', area).forEach(el => {
      el.addEventListener('click', () => {
        const side = el.dataset.calendarSide;
        const nextMode = el.dataset.calendarMode;
        if (!side || !nextMode) return;
        const stateKey = side === 'left' ? 'thai-calendar-left' : 'thai-calendar-right';
        if (State.reportOptions[stateKey] === nextMode) return;
        State.reportOptions[stateKey] = nextMode;
        sortState.key = '';
        sortState.dir = 'asc';
        Pages.thaiCalendar(area);
      });
    });
    $$('.report-sort', area).forEach(el => {
      el.addEventListener('click', () => {
        toggleNamedSort(sortState, el.dataset.reportSort);
        Pages.thaiCalendar(area);
      });
    });
    App._currentTableExport = () => {
      return buildThaiCalendarExportPayload(sorted, visibleYears, leftCfg, rightCfg, helpers, sortState);
    };
    bindPageImageActions(area, 'report-card', 'thai-calendar');
    App._currentExport = null;
  },

  async masterAnnualized(area) {
    setLoading(area, 'กำลังโหลดรายงาน Master Annualized...');

    let rawRows;
    try {
      await ensureSelectedFundsCatalog();
      rawRows = await fetchCached('master-annualized');
    } catch (e) {
      setError(area, e.message, 'master-annualized');
      return;
    }

    const headers = rawRows[0] || [];
    const CI = {
      name: findColumnIndex(headers, ['Group/Investment']),
      r3m: findColumnIndex(headers, ['Return(Cumulative) 3M']),
      r6m: findColumnIndex(headers, ['Return(Cumulative) 6M']),
      rytd: findColumnIndex(headers, ['Return(Cumulative) YTD']),
      r1y: findColumnIndex(headers, ['Return(Cumulative) 1Y']),
      r3y: findColumnIndex(headers, ['Return(Annualized) 3Y']),
      r5y: findColumnIndex(headers, ['Return(Annualized) 5Y']),
      r10y: findColumnIndex(headers, ['Return(Annualized) 10Y']),
    };
    const metricKeys = ['r3m','r6m','rytd','r1y','r3y','r5y','r10y'];
    const get = (row, i) => i >= 0 ? String(row[i] ?? '').trim() : '';

    const masterRows = rawRows.slice(1).map(row => ({
      row,
      name: get(row, CI.name),
      key: normalizeMasterMatchText(get(row, CI.name)),
      code: normalizeMasterMatchText(get(row, CI.name)),
    })).filter(item => item.name);

    const thaiSource = Object.values(State.selectedFunds).filter(f => f.masterName && f.masterName !== '0');
    const thaiFunds = State.selectedKeys.size > 0
      ? thaiSource.filter(f => State.selectedKeys.has(f.key))
      : thaiSource;

    const masterLinks = {};
    thaiFunds.forEach(fund => {
      const matched = findBestMasterRow(masterRows, fund.masterName);
      if (!matched) return;
      if (!masterLinks[matched.key]) masterLinks[matched.key] = [];
      masterLinks[matched.key].push(fund);
    });

    const displayRows = Object.keys(masterLinks).length > 0
      ? masterRows.filter(item => masterLinks[item.key]?.length)
      : masterRows;

    const { ranks: rankMap, totals: rankTotals } = buildMetricRanks(displayRows, metricKeys, (item, key) => get(item.row, CI[key]));
    const sortState = State.reportSorts['master-annualized'];
    const sortableMasterAnnualized = (item, key) => {
      const mapping = {
        name: item.name,
        thai: (masterLinks[item.key] || []).map(f => f.code).join(', '),
        r3m: get(item.row, CI.r3m),
        r6m: get(item.row, CI.r6m),
        rytd: get(item.row, CI.rytd),
        r1y: get(item.row, CI.r1y),
        r3y: get(item.row, CI.r3y),
        r5y: get(item.row, CI.r5y),
        r10y: get(item.row, CI.r10y),
        rank3m: rankMap[item.key]?.r3m ?? '',
        rank6m: rankMap[item.key]?.r6m ?? '',
        rankytd: rankMap[item.key]?.rytd ?? '',
        rank1y: rankMap[item.key]?.r1y ?? '',
        rank3y: rankMap[item.key]?.r3y ?? '',
        rank5y: rankMap[item.key]?.r5y ?? '',
        rank10y: rankMap[item.key]?.r10y ?? '',
      };
      return mapping[key];
    };
    const sorted = sortState.key
      ? [...displayRows].sort((a, b) => compareValues(sortableMasterAnnualized(a, sortState.key), sortableMasterAnnualized(b, sortState.key), sortState.dir))
      : displayRows;

    const body = sorted.map(item => {
      const thaiCodes = masterLinks[item.key] || [];
      const thaiHtml = thaiCodes.length
        ? thaiCodes.map(f => {
            const ci = State.highlights[f.key];
            const style = ci !== undefined ? ` style="background:${HL_COLORS[ci].bg};"` : '';
            return `<span class="linked-fund-chip"${style}>${esc(f.code)}</span>`;
          }).join(' ')
        : '-';
      return `
        <tr>
          <td class="master-name-cell">${esc(item.name)}</td>
          <td class="thai-link-cell">${thaiHtml}</td>
          ${metricKeys.map(k => `<td class="report-num">${esc(formatReturnDisplay(get(item.row, CI[k])) || '-')}</td>`).join('')}
          ${metricKeys.map(k => {
            const rankKey = `rank${k.slice(1)}`;
            const value = sortableMasterAnnualized(item, rankKey);
            return `<td class="report-num report-rank-cell" style="${rankCellStyle(value, rankTotals[k])}">${esc(value || '-')}</td>`;
          }).join('')}
        </tr>`;
    }).join('');

    area.innerHTML = `
      ${pageToolActions('master-annualized', CONFIG.PAGES['master-annualized']?.source || 'AVP Master Fund ID')}
      <div class="card report-card report-card-master" id="report-card">
        <table class="annualized-report master-annualized-report">
          <thead>
            <tr class="report-group-row">
              <th colspan="2" class="group-blank"></th>
              <th colspan="7" class="group-blue">ผลตอบแทน (%)</th>
              <th colspan="7" class="group-navy">อันดับในกลุ่มที่แสดง</th>
            </tr>
            <tr>
              <th class="report-sort ${sortState.key === 'name' ? 'is-active' : ''}" data-report-sort="name">${renderSortLabel('Master Fund', sortState.key === 'name', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'thai' ? 'is-active' : ''}" data-report-sort="thai">${renderSortLabel('กองทุนในไทย', sortState.key === 'thai', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r3m' ? 'is-active' : ''}" data-report-sort="r3m">${renderSortLabel('3M', sortState.key === 'r3m', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r6m' ? 'is-active' : ''}" data-report-sort="r6m">${renderSortLabel('6M', sortState.key === 'r6m', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rytd' ? 'is-active' : ''}" data-report-sort="rytd">${renderSortLabel('YTD', sortState.key === 'rytd', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r1y' ? 'is-active' : ''}" data-report-sort="r1y">${renderSortLabel('1Y', sortState.key === 'r1y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r3y' ? 'is-active' : ''}" data-report-sort="r3y">${renderSortLabel('3Y', sortState.key === 'r3y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r5y' ? 'is-active' : ''}" data-report-sort="r5y">${renderSortLabel('5Y', sortState.key === 'r5y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r10y' ? 'is-active' : ''}" data-report-sort="r10y">${renderSortLabel('10Y', sortState.key === 'r10y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank3m' ? 'is-active' : ''}" data-report-sort="rank3m">${renderSortLabel('3M', sortState.key === 'rank3m', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank6m' ? 'is-active' : ''}" data-report-sort="rank6m">${renderSortLabel('6M', sortState.key === 'rank6m', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rankytd' ? 'is-active' : ''}" data-report-sort="rankytd">${renderSortLabel('YTD', sortState.key === 'rankytd', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank1y' ? 'is-active' : ''}" data-report-sort="rank1y">${renderSortLabel('1Y', sortState.key === 'rank1y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank3y' ? 'is-active' : ''}" data-report-sort="rank3y">${renderSortLabel('3Y', sortState.key === 'rank3y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank5y' ? 'is-active' : ''}" data-report-sort="rank5y">${renderSortLabel('5Y', sortState.key === 'rank5y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank10y' ? 'is-active' : ''}" data-report-sort="rank10y">${renderSortLabel('10Y', sortState.key === 'rank10y', sortState.dir)}</th>
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>`;

    $$('.report-sort', area).forEach(el => {
      el.addEventListener('click', () => {
        toggleNamedSort(sortState, el.dataset.reportSort);
        Pages.masterAnnualized(area);
      });
    });
    bindPageImageActions(area, 'report-card', 'master-annualized');
    App._currentExport = null;
    App._currentTableExport = () => {
      return buildMasterAnnualizedExportPayload(sorted, metricKeys, masterLinks, rankMap, rankTotals, CI, get, sortState);
    };
  },

  async masterAnnualizedV2(area) {
    setLoading(area, 'กำลังโหลดรายงาน Master Annualized V2...');

    let rawRows;
    try {
      await ensureSelectedFundsCatalog();
      rawRows = await fetchCached('master-annualized-v2');
    } catch (e) {
      setError(area, e.message, 'master-annualized-v2');
      return;
    }

    const headers = rawRows[0] || [];
    const CI = {
      name: findColumnIndex(headers, ['Group/Investment']),
      isin: findColumnIndex(headers, ['ISIN']),
      currency: findColumnIndex(headers, ['Base Currency']),
      r3m: findColumnIndex(headers, ['Return(Cumulative) 3M']),
      r6m: findColumnIndex(headers, ['Return(Cumulative) 6M']),
      rytd: findColumnIndex(headers, ['Return(Cumulative) YTD']),
      r1y: findColumnIndex(headers, ['Return(Cumulative) 1Y']),
      r3y: findColumnIndex(headers, ['Return(Annualized) 3Y']),
      r5y: findColumnIndex(headers, ['Return(Annualized) 5Y']),
      r10y: findColumnIndex(headers, ['Return(Annualized) 10Y']),
    };
    const metricKeys = ['r3m','r6m','rytd','r1y','r3y','r5y','r10y'];
    const get = (row, i) => i >= 0 ? String(row[i] ?? '').trim() : '';

    const masterRows = rawRows.slice(1).map((row, index) => ({
      row,
      idx: index,
      name: get(row, CI.name),
      isin: get(row, CI.isin),
      currency: get(row, CI.currency),
      key: `${get(row, CI.isin)}::${get(row, CI.currency)}::${get(row, CI.name)}::${index}`,
      code: `${get(row, CI.isin)}::${get(row, CI.currency)}::${get(row, CI.name)}::${index}`,
    })).filter(item => item.name && item.isin);

    const thaiSource = Object.values(State.selectedFunds).filter(f => f.masterId && f.masterId !== '-');
    const thaiFunds = thaiSource.filter(f => State.selectedKeys.has(f.key));

    if (!thaiFunds.length) {
      setError(area, 'ยังไม่ได้เลือกกองทุนจากเมนูเลือกกองทุน หรือกองที่เลือกยังไม่ผูก Master Fund ID', 'master-annualized-v2');
      return;
    }

    const linksByRowKey = {};
    thaiFunds.forEach(fund => {
      const isin = String(fund.masterId || '').trim();
      if (!isin) return;
      const matches = masterRows.filter(item => item.isin === isin);
      if (!matches.length) return;
      matches.forEach(item => {
        if (!linksByRowKey[item.key]) linksByRowKey[item.key] = [];
        linksByRowKey[item.key].push(fund);
      });
    });

    const displayRows = masterRows.filter(item => linksByRowKey[item.key]?.length);
    const { ranks: rankMap, totals: rankTotals } = buildMetricRanks(displayRows, metricKeys, (item, key) => get(item.row, CI[key]));
    const sortState = State.reportSorts['master-annualized-v2'];
    const sortableMasterAnnualizedV2 = (item, key) => {
      const mapping = {
        name: item.name,
        currency: item.currency,
        thai: (linksByRowKey[item.key] || []).map(f => f.code).join(', '),
        r3m: get(item.row, CI.r3m),
        r6m: get(item.row, CI.r6m),
        rytd: get(item.row, CI.rytd),
        r1y: get(item.row, CI.r1y),
        r3y: get(item.row, CI.r3y),
        r5y: get(item.row, CI.r5y),
        r10y: get(item.row, CI.r10y),
        rank3m: rankMap[item.key]?.r3m ?? '',
        rank6m: rankMap[item.key]?.r6m ?? '',
        rankytd: rankMap[item.key]?.rytd ?? '',
        rank1y: rankMap[item.key]?.r1y ?? '',
        rank3y: rankMap[item.key]?.r3y ?? '',
        rank5y: rankMap[item.key]?.r5y ?? '',
        rank10y: rankMap[item.key]?.r10y ?? '',
      };
      return mapping[key];
    };
    const sorted = sortState.key
      ? [...displayRows].sort((a, b) => compareValues(sortableMasterAnnualizedV2(a, sortState.key), sortableMasterAnnualizedV2(b, sortState.key), sortState.dir))
      : displayRows;

    const body = sorted.map(item => {
      const thaiFundsForRow = linksByRowKey[item.key] || [];
      const uniqueThaiFunds = [...new Map(thaiFundsForRow.map(f => [f.key, f])).values()];
      const thaiHtml = uniqueThaiFunds.length
        ? uniqueThaiFunds.map(f => {
            const ci = State.highlights[f.key];
            const style = ci !== undefined ? ` style="background:${HL_COLORS[ci].bg};"` : '';
            return `<span class="linked-fund-chip"${style}>${esc(f.code)}</span>`;
          }).join(' ')
        : '-';
      return `
        <tr>
          <td class="master-name-cell">${esc(item.name)}</td>
          <td>${esc(item.currency || '-')}</td>
          <td class="thai-link-cell">${thaiHtml}</td>
          ${metricKeys.map(k => `<td class="report-num">${esc(formatReturnDisplay(get(item.row, CI[k])) || '-')}</td>`).join('')}
          ${metricKeys.map(k => {
            const rankKey = `rank${k.slice(1)}`;
            const value = sortableMasterAnnualizedV2(item, rankKey);
            return `<td class="report-num report-rank-cell" style="${rankCellStyle(value, rankTotals[k])}">${esc(value || '-')}</td>`;
          }).join('')}
        </tr>`;
    }).join('');

    area.innerHTML = `
      ${pageToolActions('master-annualized-v2', CONFIG.PAGES['master-annualized-v2']?.source || 'AVP Master Fund ID')}
      <div class="card report-card report-card-master" id="report-card">
        <table class="annualized-report master-annualized-report">
          <thead>
            <tr class="report-group-row">
              <th colspan="3" class="group-blank"></th>
              <th colspan="7" class="group-blue">ผลตอบแทน (%)</th>
              <th colspan="7" class="group-navy">อันดับในกลุ่มที่แสดง</th>
            </tr>
            <tr>
              <th class="report-sort ${sortState.key === 'name' ? 'is-active' : ''}" data-report-sort="name">${renderSortLabel('Master Fund', sortState.key === 'name', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'currency' ? 'is-active' : ''}" data-report-sort="currency">${renderSortLabel('Base Currency', sortState.key === 'currency', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'thai' ? 'is-active' : ''}" data-report-sort="thai">${renderSortLabel('กองทุนในไทย', sortState.key === 'thai', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r3m' ? 'is-active' : ''}" data-report-sort="r3m">${renderSortLabel('3M', sortState.key === 'r3m', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r6m' ? 'is-active' : ''}" data-report-sort="r6m">${renderSortLabel('6M', sortState.key === 'r6m', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rytd' ? 'is-active' : ''}" data-report-sort="rytd">${renderSortLabel('YTD', sortState.key === 'rytd', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r1y' ? 'is-active' : ''}" data-report-sort="r1y">${renderSortLabel('1Y', sortState.key === 'r1y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r3y' ? 'is-active' : ''}" data-report-sort="r3y">${renderSortLabel('3Y', sortState.key === 'r3y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r5y' ? 'is-active' : ''}" data-report-sort="r5y">${renderSortLabel('5Y', sortState.key === 'r5y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'r10y' ? 'is-active' : ''}" data-report-sort="r10y">${renderSortLabel('10Y', sortState.key === 'r10y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank3m' ? 'is-active' : ''}" data-report-sort="rank3m">${renderSortLabel('3M', sortState.key === 'rank3m', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank6m' ? 'is-active' : ''}" data-report-sort="rank6m">${renderSortLabel('6M', sortState.key === 'rank6m', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rankytd' ? 'is-active' : ''}" data-report-sort="rankytd">${renderSortLabel('YTD', sortState.key === 'rankytd', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank1y' ? 'is-active' : ''}" data-report-sort="rank1y">${renderSortLabel('1Y', sortState.key === 'rank1y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank3y' ? 'is-active' : ''}" data-report-sort="rank3y">${renderSortLabel('3Y', sortState.key === 'rank3y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank5y' ? 'is-active' : ''}" data-report-sort="rank5y">${renderSortLabel('5Y', sortState.key === 'rank5y', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'rank10y' ? 'is-active' : ''}" data-report-sort="rank10y">${renderSortLabel('10Y', sortState.key === 'rank10y', sortState.dir)}</th>
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>`;

    $$('.report-sort', area).forEach(el => {
      el.addEventListener('click', () => {
        toggleNamedSort(sortState, el.dataset.reportSort);
        Pages.masterAnnualizedV2(area);
      });
    });
    bindPageImageActions(area, 'report-card', 'master-annualized-v2');
    App._currentExport = null;
    App._currentTableExport = () => {
      return buildMasterAnnualizedV2ExportPayload(sorted, metricKeys, linksByRowKey, rankMap, rankTotals, CI, get, sortState);
    };
  },

  async masterCalendar(area) {
    setLoading(area, 'กำลังโหลดรายงาน Master Calendar Year...');

    let rawRows;
    try {
      await ensureSelectedFundsCatalog();
      rawRows = await fetchCached('master-calendar');
    } catch (e) {
      setError(area, e.message, 'master-calendar');
      return;
    }

    const headers = rawRows[0] || [];
    const yearKeys = ['2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'];
    const CI = {
      name: findColumnIndex(headers, ['Group/Investment']),
      isin: findColumnIndex(headers, ['ISIN']),
      currency: findColumnIndex(headers, ['Base Currency']),
      ...Object.fromEntries(yearKeys.map(year => [`ret${year}`, findColumnIndex(headers, [`Return(Cumulative) ${year}`])])),
    };
    const get = (row, i) => i >= 0 ? String(row[i] ?? '').trim() : '';

    const masterRows = rawRows.slice(1).map((row, index) => ({
      row,
      idx: index,
      name: get(row, CI.name),
      isin: get(row, CI.isin),
      currency: get(row, CI.currency),
      key: `${get(row, CI.isin)}::${get(row, CI.currency)}::${get(row, CI.name)}::${index}`,
    })).filter(item => item.name && item.isin);

    const thaiSource = Object.values(State.selectedFunds).filter(f => f.masterId && f.masterId !== '-');
    const thaiFunds = thaiSource.filter(f => State.selectedKeys.has(f.key));

    if (!thaiFunds.length) {
      setError(area, 'ยังไม่ได้เลือกกองทุนจากเมนูเลือกกองทุน หรือกองที่เลือกยังไม่ผูก Master Fund ID', 'master-calendar');
      return;
    }

    const linksByRowKey = {};
    thaiFunds.forEach(fund => {
      const isin = String(fund.masterId || '').trim();
      if (!isin) return;
      const matches = masterRows.filter(item => item.isin === isin);
      matches.forEach(item => {
        if (!linksByRowKey[item.key]) linksByRowKey[item.key] = [];
        linksByRowKey[item.key].push(fund);
      });
    });

    const displayRows = masterRows.filter(item => linksByRowKey[item.key]?.length);
    const sortState = State.reportSorts['master-calendar'];
    const sortableMasterCalendar = (item, key) => {
      const mapping = {
        name: item.name,
        currency: item.currency,
        thai: (linksByRowKey[item.key] || []).map(f => f.code).join(', '),
        ...Object.fromEntries(yearKeys.map(year => [`ret-${year}`, get(item.row, CI[`ret${year}`])])),
      };
      return mapping[key];
    };
    const sorted = sortState.key
      ? [...displayRows].sort((a, b) => compareValues(sortableMasterCalendar(a, sortState.key), sortableMasterCalendar(b, sortState.key), sortState.dir))
      : displayRows;

    const body = sorted.map(item => {
      const thaiFundsForRow = linksByRowKey[item.key] || [];
      const uniqueThaiFunds = [...new Map(thaiFundsForRow.map(f => [f.key, f])).values()];
      const thaiHtml = uniqueThaiFunds.length
        ? uniqueThaiFunds.map(f => {
            const ci = State.highlights[f.key];
            const style = ci !== undefined ? ` style="background:${HL_COLORS[ci].bg};"` : '';
            return `<span class="linked-fund-chip"${style}>${esc(f.code)}</span>`;
          }).join(' ')
        : '-';
      return `
        <tr>
          <td class="master-name-cell">${esc(item.name)}</td>
          <td>${esc(item.currency || '-')}</td>
          <td class="thai-link-cell">${thaiHtml}</td>
          ${yearKeys.map(year => `<td class="report-num">${esc(formatReturnDisplay(get(item.row, CI[`ret${year}`])) || '-')}</td>`).join('')}
        </tr>`;
    }).join('');

    area.innerHTML = `
      ${pageToolActions('master-calendar', CONFIG.PAGES['master-calendar']?.source || 'AVP Master Fund ID')}
      <div class="card report-card report-card-master" id="report-card">
        <table class="annualized-report master-annualized-report">
          <thead>
            <tr class="report-group-row">
              <th colspan="3" class="group-blank"></th>
              <th colspan="${yearKeys.length}" class="group-blue">Return (Cumulative) (%)</th>
            </tr>
            <tr>
              <th class="report-sort ${sortState.key === 'name' ? 'is-active' : ''}" data-report-sort="name">${renderSortLabel('Master Fund', sortState.key === 'name', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'currency' ? 'is-active' : ''}" data-report-sort="currency">${renderSortLabel('Base Currency', sortState.key === 'currency', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'thai' ? 'is-active' : ''}" data-report-sort="thai">${renderSortLabel('กองทุนในไทย', sortState.key === 'thai', sortState.dir)}</th>
              ${yearKeys.map(year => `<th class="report-sort ${sortState.key === `ret-${year}` ? 'is-active' : ''}" data-report-sort="ret-${year}">${renderSortLabel(year, sortState.key === `ret-${year}`, sortState.dir)}</th>`).join('')}
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>`;

    $$('.report-sort', area).forEach(el => {
      el.addEventListener('click', () => {
        toggleNamedSort(sortState, el.dataset.reportSort);
        Pages.masterCalendar(area);
      });
    });
    bindPageImageActions(area, 'report-card', 'master-calendar');
    App._currentExport = null;
    App._currentTableExport = () => buildMasterCalendarExportPayload(sorted, yearKeys, linksByRowKey, CI, get, sortState);
  },

  async masterFees(area) {
    setLoading(area, 'กำลังโหลดรายงานค่าธรรมเนียม...');

    try {
      const [rawSecRows, universe] = await Promise.all([
        fetchCached('master-placeholder-1'),
        buildSelectedMasterUniverse(),
      ]);
      const rawLookup = buildRawSecLookup(rawSecRows);
      const feeRows = buildFeeComparisonRows(universe, rawLookup)
        .sort((a, b) => compareValues(a.combined, b.combined, 'asc'));

      if (!feeRows.length) {
        setError(area, 'ไม่พบข้อมูลค่าธรรมเนียมที่จับคู่ได้จาก Raw For Sec และ AVP Master Fund ID', 'master-placeholder-1');
        return;
      }

      const totalAvg = feeRows.reduce((sum, row) => sum + (Number.isNaN(row.combined) ? 0 : row.combined), 0) / feeRows.length;
      const cheapest = feeRows[0];
      const priciest = [...feeRows].sort((a, b) => compareValues(a.combined, b.combined, 'desc'))[0];
      const avgThai = feeRows.reduce((sum, row) => sum + (Number.isNaN(row.thaiTer) ? 0 : row.thaiTer), 0) / feeRows.length;
      const avgMaster = feeRows.reduce((sum, row) => sum + (Number.isNaN(row.masterTer) ? 0 : row.masterTer), 0) / feeRows.length;
      const maxCombined = Math.max(...feeRows.map(row => row.combined || 0), 0);

      const barRows = feeRows.map(row => {
        const masterPct = maxCombined > 0 ? ((row.masterTer || 0) / maxCombined) * 100 : 0;
        const thaiPct = maxCombined > 0 ? ((row.thaiTer || 0) / maxCombined) * 100 : 0;
        return `
          <div class="insight-bar-row">
            <div class="insight-bar-meta">
              <strong>${esc(row.thaiCode)}</strong>
              <span>${esc(row.masterName)}</span>
            </div>
            <div class="insight-bar-track">
              <span class="insight-bar-segment is-master" style="width:${masterPct}%"></span>
              <span class="insight-bar-segment is-thai" style="width:${thaiPct}%"></span>
            </div>
            <div class="insight-bar-values">
              <span>Master ${esc(row.masterTerText || '-')}</span>
              <span>Thai ${esc(row.thaiTerText || '-')}</span>
              <strong>${esc(row.combinedText || '-')}</strong>
            </div>
          </div>`;
      }).join('');

      area.innerHTML = `
        ${pageToolActions('master-placeholder-1', CONFIG.PAGES['master-placeholder-1']?.source || 'Raw For Sec + AVP Master Fund ID')}
        <div id="report-card" class="insight-page">
        ${buildInsightSummaryCards([
          { label: 'กองทุนที่จับคู่ได้', value: `${feeRows.length} กอง`, note: 'คัดจากกองที่เลือกไว้ก่อน ถ้าไม่ได้เลือกจะใช้ชุดแรกของรายการ' },
          { label: 'Combined TER เฉลี่ย', value: `${toFixedSafe(totalAvg, 2) || '-'}%`, note: `Master เฉลี่ย ${toFixedSafe(avgMaster, 2) || '-'}% + ไทยเฉลี่ย ${toFixedSafe(avgThai, 2) || '-'}%` },
          { label: 'ต่ำสุด', value: `${esc(cheapest.thaiCode)} · ${esc(cheapest.combinedText || '-') }%`, note: cheapest.masterName },
          { label: 'สูงสุด', value: `${esc(priciest.thaiCode)} · ${esc(priciest.combinedText || '-') }%`, note: priciest.masterName },
        ])}
        <div class="insight-layout insight-layout-fee">
          <div class="card insight-panel">
            <div class="insight-panel-head">
              <h3>เปรียบเทียบค่าธรรมเนียมรวม</h3>
              <p>แบ่งสีให้เห็นชัดว่าค่าใช้จ่ายมาจากฝั่ง Master Fund และกองไทยส่วนไหนมากกว่า</p>
            </div>
            <div class="insight-bar-list">${barRows}</div>
          </div>
          <div class="card insight-panel">
            <div class="insight-panel-head">
              <h3>รายละเอียดรายกอง</h3>
              <p>อ้างอิง TER จาก Raw For Sec - 2026-Q1 และ Ongoing Cost จาก AVP Master Fund ID</p>
            </div>
            ${buildInsightTable(feeRows, [
              { key: 'masterName', label: 'Master Fund', className: 'td-left' },
              { key: 'thaiCode', label: 'กองไทย', className: 'td-chip', render: row => `<span class="linked-fund-chip">${esc(row.thaiCode)}</span>` },
              { key: 'masterTerText', label: 'Master Fund' },
              { key: 'thaiTerText', label: 'กองไทย' },
              { key: 'feeDate', label: 'Date' },
              { key: 'combinedText', label: 'Combined TER', className: 'td-strong td-accent' },
            ])}
          </div>
        </div>
        </div>`;

      bindPageImageActions(area, 'report-card', 'master-fees');
      App._currentExport = null;
      App._currentTableExport = () => buildSimpleTablePayload(
        CONFIG.PAGES['master-placeholder-1']?.title || 'ค่าธรรมเนียม',
        CONFIG.PAGES['master-placeholder-1']?.source || 'Raw For Sec + AVP Master Fund ID',
        ['Master Fund', 'กองไทย', 'Master Fund', 'กองไทย', 'Date', 'Combined TER'],
        feeRows.map(row => [
          row.masterName,
          row.thaiCode,
          row.masterTerText || '-',
          row.thaiTerText || '-',
          row.feeDate || '-',
          row.combinedText || '-',
        ])
      );
    } catch (e) {
      setError(area, e.message, 'master-placeholder-1');
    }
  },

  async masterFeesV2(area) {
    setLoading(area, 'กำลังโหลดรายงานค่าธรรมเนียม V2...');

    try {
      const [rawSecRows, universe] = await Promise.all([
        fetchCached('master-placeholder-4'),
        buildSelectedMasterUniverse(),
      ]);
      const rawLookup = buildRawSecLookup(rawSecRows);
      const feeRows = buildFeeComparisonRows(universe, rawLookup)
        .sort((a, b) => compareValues(a.combined, b.combined, 'asc'))
        .map(row => {
          const matchedFund = Object.values(State.selectedFunds).find(f => f.code === row.thaiCode);
          const colorIdx = matchedFund ? State.highlights[matchedFund.key] : undefined;
          return {
            ...row,
            highlightColor: colorIdx !== undefined ? HL_COLORS[colorIdx]?.bg || '' : '',
          };
        });

      if (!feeRows.length) {
        setError(area, 'ไม่พบข้อมูลค่าธรรมเนียมที่จับคู่ได้จาก Raw For Sec และ AVP Master Fund ID', 'master-placeholder-4');
        return;
      }

      const source = CONFIG.PAGES['master-placeholder-4']?.source || 'Raw For Sec + AVP Master Fund ID';
      const maxCombined = Math.max(...feeRows.map(item => item.combined || 0), 0);

      area.innerHTML = `
        ${pageToolActions('master-placeholder-4', source)}
        <div id="report-card" class="card report-card">
          <div class="fee-v2-table-wrap">
            <table class="fee-v2-table">
              <thead>
                <tr>
                  <th rowspan="2">Master Fund</th>
                  <th rowspan="2">กองไทย</th>
                  <th colspan="4">TER (%) Q1-2026</th>
                </tr>
                <tr>
                  <th>Master Fund</th>
                  <th>กองไทย (Sec)</th>
                  <th>Date</th>
                  <th>Combined TER</th>
                </tr>
              </thead>
              <tbody>
                ${feeRows.map(row => {
                  const combinedStyle = feeCombinedStyle(row.combined, maxCombined);
                  return `
                    <tr>
                      <td class="is-master">${esc(row.masterName)}</td>
                      <td class="is-thai"${row.highlightColor ? ` style="background:${row.highlightColor}"` : ''}>${esc(row.thaiCode)}</td>
                      <td>${esc(row.masterTerText || '-')}</td>
                      <td>${esc(row.thaiTerText || '-')}</td>
                      <td>${esc(row.feeDate || '-')}</td>
                      <td class="is-combined"${combinedStyle.bg ? ` style="background:${combinedStyle.bg};color:${combinedStyle.color}"` : ''}>${esc(row.combinedText || '-')}</td>
                    </tr>`;
                }).join('')}
              </tbody>
            </table>
          </div>
        </div>`;

      bindPageImageActions(area, 'report-card', 'master-fees-v2');
      App._currentExport = null;
      App._currentTableExport = () => buildFeeV2ExportPayload(
        CONFIG.PAGES['master-placeholder-4']?.title || 'ค่าธรรมเนียม 2',
        source,
        feeRows
      );
    } catch (e) {
      setError(area, e.message, 'master-placeholder-4');
    }
  },

  async masterMenu02(area) {
    setLoading(area, 'กำลังเตรียมหน้า Top 10 Holding...');

    try {
      const universe = await buildSelectedMasterUniverse();
      const suggestedIsin = String(universe[0]?.master?.isin || universe[0]?.fund?.masterId || 'IE00BFRSYJ83').trim();
      const selectedThaiCodes = universe
        .map(({ fund }) => String(fund.code || '').trim())
        .filter(Boolean)
        .slice(0, 12);
      const selectedChips = selectedThaiCodes.length
        ? selectedThaiCodes.map(code => `<span class="ft-viewer-chip">${esc(code)}</span>`).join('')
        : '<span class="ft-viewer-empty">ยังไม่ได้เลือกกองจากเมนูเลือกกองทุน ระบบจึงใส่ ISIN ตัวอย่างไว้ให้ก่อน</span>';

      area.innerHTML = `
        <div id="report-card" class="card ft-viewer-card">
          <div class="ft-viewer-wrap">
            <div class="ft-viewer-header">
              <div>
                <h2>FT Fund Data Viewer</h2>
                <p>ใช้ดึงข้อมูล Top 10 Holding และข้อมูลกองจาก FT API ผ่าน ISIN</p>
              </div>
              <div class="ft-viewer-source">API: FT Fund Data Viewer</div>
            </div>

            <div class="ft-viewer-selection">
              <div class="ft-viewer-selection-label">กองทุนไทยที่เลือกอยู่</div>
              <div class="ft-viewer-chip-row">${selectedChips}</div>
            </div>

            <div class="ft-viewer-controls">
              <div class="ft-viewer-row">
                <input id="ft-isin-input" type="text" placeholder="เช่น IE00BFRSYJ83" value="${esc(suggestedIsin)}" />
                <button id="ft-load-btn" class="btn btn-primary" type="button">ดึงข้อมูล</button>
              </div>

              <div class="ft-viewer-checklist">
                <label><input type="checkbox" value="summary" checked> summary</label>
                <label><input type="checkbox" value="sizes" checked> sizes</label>
                <label><input type="checkbox" value="fees" checked> fees</label>
                <label><input type="checkbox" value="manager" checked> manager</label>
                <label><input type="checkbox" value="performance" checked> performance</label>
                <label><input type="checkbox" value="allocation" checked> allocation</label>
                <label><input type="checkbox" value="objective" checked> objective</label>
                <label><input type="checkbox" value="risk"> risk</label>
                <label><input type="checkbox" value="ratings"> ratings</label>
                <label><input type="checkbox" value="historical"> historical</label>
                <label><input type="checkbox" value="holdings" checked> holdings</label>
              </div>
            </div>

            <div id="ft-status" class="ft-viewer-status"></div>
            <div id="ft-output" class="ft-viewer-output"></div>
          </div>
        </div>`;

      const isinInput = $('#ft-isin-input', area);
      const loadBtn = $('#ft-load-btn', area);
      const statusEl = $('#ft-status', area);
      const outputEl = $('#ft-output', area);

      const section = (title, content) => `<div class="ft-viewer-section"><h3>${esc(title)}</h3>${content}</div>`;
      const objectTable = (obj) => {
        const rows = Object.entries(obj || {}).map(([key, value]) =>
          `<tr><th>${esc(key)}</th><td>${esc(String(value ?? ''))}</td></tr>`
        ).join('');
        return `<table><tbody>${rows}</tbody></table>`;
      };
      const arrayTable = (arr) => {
        if (!Array.isArray(arr) || !arr.length) return '<p class="ft-viewer-empty">ไม่มีข้อมูล</p>';
        const headers = Object.keys(arr[0] || {});
        const thead = `<tr>${headers.map(key => `<th>${esc(key)}</th>`).join('')}</tr>`;
        const tbody = arr.map(row =>
          `<tr>${headers.map(key => `<td>${esc(String(row?.[key] ?? ''))}</td>`).join('')}</tr>`
        ).join('');
        return `<table><thead>${thead}</thead><tbody>${tbody}</tbody></table>`;
      };
      const renderData = (data) => {
        let html = '';
        if (data.summary) html += section('Summary', objectTable(data.summary));
        if (data.sizes) html += section('Sizes', objectTable(data.sizes));
        if (data.fees) html += section('Fees', objectTable(data.fees));
        if (data.manager) html += section('Manager', objectTable(data.manager));
        if (data.objective) html += section('Objective', `<p>${esc(String(data.objective || ''))}</p>`);
        if (data.performance) html += section('Performance', objectTable(data.performance));
        if (Array.isArray(data.holdings) && data.holdings.length) html += section('Holdings', arrayTable(data.holdings));
        if (data.allocation?.assetType?.length) html += section('Allocation: Asset Type', arrayTable(data.allocation.assetType));
        if (data.allocation?.sector?.length) html += section('Allocation: Sector', arrayTable(data.allocation.sector));
        if (data.allocation?.region?.length) html += section('Allocation: Region', arrayTable(data.allocation.region));
        if (data.risk && Object.keys(data.risk).length) html += section('Risk', objectTable(data.risk));
        if (data.ratings && Object.keys(data.ratings).length) html += section('Ratings', objectTable(data.ratings));
        if (Array.isArray(data.historical) && data.historical.length) html += section('Historical Prices', arrayTable(data.historical));
        html += section('Raw JSON', `<pre>${esc(JSON.stringify(data, null, 2))}</pre>`);
        outputEl.innerHTML = html;
      };
      const loadData = async () => {
        const isin = String(isinInput?.value || '').trim();
        const fields = $$('.ft-viewer-checklist input:checked', area).map(el => el.value);

        if (!isin) {
          statusEl.textContent = 'กรุณาใส่ ISIN';
          outputEl.innerHTML = '';
          return;
        }

        const url = `${TOP_10_HOLDING_API_URL}?isin=${encodeURIComponent(isin)}&fields=${encodeURIComponent(fields.join(','))}`;
        statusEl.textContent = 'กำลังโหลด...';
        outputEl.innerHTML = '';

        try {
          const res = await fetch(url);
          const data = await res.json();
          if (!data.ok) {
            statusEl.textContent = data.error || 'เกิดข้อผิดพลาด';
            return;
          }

          statusEl.textContent = `โหลดสำเร็จ: ${data.isin || isin}`;
          renderData(data);
        } catch (err) {
          statusEl.textContent = `เกิดข้อผิดพลาด: ${err.message}`;
        }
      };

      loadBtn?.addEventListener('click', loadData);
      isinInput?.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') loadData();
      });
      loadData();

      App._currentExport = null;
      App._currentTableExport = null;
    } catch (e) {
      setError(area, e.message, 'master-placeholder-2');
    }
  },

  async masterMenu03(area) {
    setLoading(area, 'กำลังโหลดกราฟค่าธรรมเนียมเทียบผลตอบแทน 5Y...');

    try {
      const universe = await buildSelectedMasterUniverse();
      const rows = universe
        .map(({ fund, master }) => ({
          thaiCode: fund.code,
          masterName: master.name,
          ongoingCost: parseNum(master.ongoingCost),
          return5y: parseNum(master.return5y),
          sharpe5y: parseNum(master.sharpe5y),
          drawdown5y: parseNum(master.drawdown5y),
        }))
        .filter(row => !Number.isNaN(row.ongoingCost) && !Number.isNaN(row.return5y))
        .sort((a, b) => compareValues(a.return5y, b.return5y, 'desc'));

      if (!rows.length) {
        setError(area, 'ไม่พบข้อมูล Ongoing Cost / Return 5Y เพียงพอสำหรับสร้างกราฟ', 'master-placeholder-3');
        return;
      }

      const efficient = [...rows].sort((a, b) => compareValues((a.return5y - a.ongoingCost), (b.return5y - b.ongoingCost), 'desc'))[0];
      const cheapest = [...rows].sort((a, b) => compareValues(a.ongoingCost, b.ongoingCost, 'asc'))[0];
      const deepestDrawdown = [...rows]
        .filter(row => !Number.isNaN(row.drawdown5y))
        .sort((a, b) => compareValues(a.drawdown5y, b.drawdown5y, 'asc'))[0] || null;

      const points = rows.map(row => ({
        x: row.ongoingCost,
        y: row.return5y,
        label: row.thaiCode,
        color: row === efficient ? '#76d943' : '#163b72',
        r: row === efficient ? 6.4 : 4.8,
      }));

      area.innerHTML = `
        ${pageToolActions('master-placeholder-3', CONFIG.PAGES['master-placeholder-3']?.source || 'AVP Master Fund ID')}
        <div id="report-card" class="insight-page">
        ${buildInsightSummaryCards([
          { label: 'คุ้มค่าที่สุด', value: `${esc(efficient.thaiCode)} · ${(efficient.return5y - efficient.ongoingCost).toFixed(2)}`, note: 'ดูจาก Return 5Y หัก Ongoing Cost' },
          { label: 'Ongoing Cost ต่ำสุด', value: `${esc(cheapest.thaiCode)} · ${toFixedSafe(cheapest.ongoingCost, 2)}%`, note: cheapest.masterName },
          { label: 'Max Drawdown 5Y ลึกสุด', value: deepestDrawdown ? `${esc(deepestDrawdown.thaiCode)} · ${toFixedSafe(deepestDrawdown.drawdown5y, 2)}%` : '-', note: deepestDrawdown?.masterName || 'ใช้เป็นมุมเสริมเรื่อง downside' },
          { label: 'จำนวนกองในกราฟ', value: `${rows.length} กอง`, note: 'กองที่อยู่ซ้ายและสูง จะเด่นเรื่อง cost efficiency มากกว่า' },
        ])}
        <div class="insight-layout">
          <div class="card insight-panel">
            <div class="insight-panel-head">
              <h3>Cost vs Return ของ Master Fund 5Y</h3>
              <p>ช่วยมองว่าค่าธรรมเนียมฝั่ง Master สูงแค่ไหนเมื่อเทียบกับผลตอบแทนย้อนหลัง 5 ปี</p>
            </div>
            <div class="insight-scatter-wrap">
              ${buildScatterSvg(points, {
                title: 'Master Fund Cost vs Return 5Y',
                xLabel: 'Ongoing Cost Actual',
                yLabel: 'Return 5Y',
              })}
            </div>
          </div>
          <div class="card insight-panel">
            <div class="insight-panel-head">
              <h3>ตารางประกอบ</h3>
              <p>มี Sharpe 5Y ให้ดูเพิ่มเพื่อชั่งน้ำหนักความคุ้มค่าต่อความเสี่ยง</p>
            </div>
            ${buildInsightTable(rows, [
              { key: 'masterName', label: 'Master Fund', className: 'td-left' },
              { key: 'thaiCode', label: 'กองไทย', className: 'td-chip', render: row => `<span class="linked-fund-chip">${esc(row.thaiCode)}</span>` },
              { key: 'ongoingCost', label: 'Ongoing Cost', render: row => esc(toFixedSafe(row.ongoingCost, 2)) },
              { key: 'return5y', label: 'Return (5Y)', className: 'td-strong', render: row => esc(toFixedSafe(row.return5y, 2)) },
              { key: 'sharpe5y', label: 'Sharpe (5Y)', render: row => esc(toFixedSafe(row.sharpe5y, 2) || '-') },
            ])}
          </div>
        </div>
        </div>`;

      bindPageImageActions(area, 'report-card', 'master-menu-03');
      App._currentExport = null;
      App._currentTableExport = () => buildSimpleTablePayload(
        CONFIG.PAGES['master-placeholder-3']?.title || 'Master Fund Menu 03',
        CONFIG.PAGES['master-placeholder-3']?.source || 'AVP Master Fund ID',
        ['Master Fund', 'กองไทย', 'Ongoing Cost', 'Return (5Y)', 'Sharpe (5Y)'],
        rows.map(row => [
          row.masterName,
          row.thaiCode,
          toFixedSafe(row.ongoingCost, 2),
          toFixedSafe(row.return5y, 2),
          toFixedSafe(row.sharpe5y, 2) || '-',
        ])
      );
    } catch (e) {
      setError(area, e.message, 'master-placeholder-3');
    }
  },

  /* ── GUIDE ── */
  guide(area) {
    area.innerHTML = `
      <div class="card guide-wrap">
        <div class="card-header">
          <span class="card-title">📖 คู่มือการใช้งาน</span>
          <span class="badge badge-primary">v1.0</span>
        </div>
        <div class="card-body">

          <h2>Fund Selection Tool For FP2</h2>
          <p>ระบบเลือกกองทุนสำหรับ Financial Planner รุ่นที่ 2 เชื่อมต่อกับ Google Sheets ขององค์กรโดยตรง ผ่าน Google OAuth 2.0 — เข้าถึงได้เฉพาะบัญชีที่มีสิทธิ์ใน Sheets เท่านั้น</p>

          <h2>⚙️ การตั้งค่าก่อนใช้งานครั้งแรก</h2>
          <div class="guide-callout">
            <strong>⚠️ ต้องทำก่อนเปิดใช้งาน</strong>
            <ol style="margin-top:8px">
              <li>สร้าง OAuth 2.0 Client ID ที่ <strong>Google Cloud Console</strong></li>
              <li>เปิดไฟล์ <code>js/config.js</code> แล้วใส่ <code>CLIENT_ID</code></li>
              <li>ตรวจสอบชื่อ Tab ของแต่ละ Google Sheet ให้ถูกต้อง</li>
              <li>Host ไฟล์บน Web Server (เช่น GitHub Pages หรือ Firebase Hosting)</li>
            </ol>
          </div>

          <h3>วิธีสร้าง Google OAuth 2.0 Client ID</h3>
          <ol>
            <li>ไปที่ <code>console.cloud.google.com</code></li>
            <li>สร้าง Project ใหม่ (หรือใช้ Project ที่มีอยู่)</li>
            <li>ไปที่ <strong>APIs &amp; Services → Enabled APIs</strong> → เปิดใช้ <em>Google Sheets API</em></li>
            <li>ไปที่ <strong>APIs &amp; Services → Credentials → Create Credentials → OAuth client ID</strong></li>
            <li>เลือก Application type: <strong>Web application</strong></li>
            <li>ใส่ <strong>Authorized JavaScript origins</strong> เช่น <code>https://yourdomain.com</code></li>
            <li>Copy Client ID ที่ได้ไปใส่ใน <code>js/config.js</code></li>
          </ol>

          <h2>📋 เมนูหลัก</h2>

          <h3>⊞ แดชบอร์ด</h3>
          <p>แสดงสรุปจำนวนข้อมูลจากทุก Google Sheet และ Quick Links ไปยังแต่ละหน้า</p>

          <h3>☑ เลือกกองทุน</h3>
          <p>ดึงข้อมูลจาก <strong>AVP Master Fund ID</strong> รองรับ:</p>
          <ul>
            <li>ติ๊กเลือกกองทุนหลายรายการพร้อมกัน</li>
            <li>ค้นหากองทุนแบบ Real-time</li>
            <li>⚖ เปรียบเทียบกองทุนที่เลือก (เลือกอย่างน้อย 2 รายการ)</li>
            <li>Export เฉพาะกองทุนที่เลือกเป็น Excel</li>
          </ul>

          <h3>📈 กองทุนไทย Annualized / 📅 Calendar Year</h3>
          <p>ดึงข้อมูลจาก <strong>AVP Thai Fund for Quality</strong> — แต่ละหน้าใช้คนละ Tab ของ Sheet เดียวกัน</p>

          <h3>🌐 Master Fund Annualized / 📆 Calendar Year</h3>
          <p>ดึงข้อมูลจาก <strong>AVP Master Fund ID</strong> — แต่ละหน้าใช้คนละ Tab ของ Sheet เดียวกัน</p>

          <h2>🗂 Google Sheets ที่ใช้ในระบบ</h2>
          <div class="guide-sheet-list">
            <div class="guide-sheet-item">
              <div class="sheet-name">AVP Master Fund ID</div>
              <div class="sheet-desc">รายการกองทุนหลัก, Annualized, Calendar Year</div>
            </div>
            <div class="guide-sheet-item">
              <div class="sheet-name">AVP Thai Fund for Quality</div>
              <div class="sheet-desc">กองทุนไทยคัดสรร Annualized &amp; Calendar</div>
            </div>
            <div class="guide-sheet-item">
              <div class="sheet-name">Percentrank Freestyle</div>
              <div class="sheet-desc">คะแนน Percentrank</div>
            </div>
            <div class="guide-sheet-item">
              <div class="sheet-name">Raw For Sec</div>
              <div class="sheet-desc">ข้อมูลดิบสำหรับ SEC</div>
            </div>
            <div class="guide-sheet-item">
              <div class="sheet-name">iShare Index Passive Return</div>
              <div class="sheet-desc">ดัชนีอ้างอิง Passive</div>
            </div>
          </div>

          <h2>🔧 การตั้งค่าชื่อ Tab</h2>
          <p>แก้ไขชื่อ Tab ของแต่ละหน้าได้ในไฟล์ <code>js/config.js</code> ที่ส่วน <code>PAGES</code>:</p>
          <div class="guide-callout">
            <code>tabName: 'Sheet1'</code> → เปลี่ยนเป็นชื่อ Tab จริงของ Google Sheet<br>
            ตัวอย่าง: <code>tabName: 'Annualized'</code>
          </div>

          <h2>📥 การ Export ข้อมูล</h2>
          <p>กดปุ่ม <strong>ดาวน์โหลด Excel</strong> ที่มุมบนขวาเพื่อ Export ข้อมูลทั้งหน้าเป็นไฟล์ <code>.xlsx</code></p>
          <p>ในหน้า "เลือกกองทุน" สามารถ Export เฉพาะกองทุนที่ติ๊กเลือกไว้ได้</p>

          <h2>↻ การรีเฟรชข้อมูล</h2>
          <p>ระบบแคชข้อมูลไว้ <strong>5 นาที</strong> กดปุ่ม <strong>รีเฟรช</strong> เพื่อดึงข้อมูลใหม่ทันที</p>

          <hr class="divider">
          <p class="text-muted">Fund Selection Tool For FP2 · Avenger Planner · v1.0 · ${new Date().getFullYear()}</p>
        </div>
      </div>`;

    App._currentExport = null;
  },

  placeholder(area, title) {
    area.innerHTML = `
      <div class="card">
        <div class="state-box">
          <div class="state-icon">🗂</div>
          <strong>${esc(title)}</strong>
          <p>เมนูนี้ถูกเปิดรอไว้แล้ว และยังอยู่ระหว่างเตรียมข้อมูล/หน้าจอ</p>
        </div>
      </div>`;
    App._currentExport = null;
  },

  /* ── Avenger Studio (PP Page embedded via iframe) ── */
  async avengerStudio(area) {
    area.style.padding = '0';
    area.style.overflow = 'hidden';
    area.style.display = 'flex';
    area.style.flexDirection = 'column';
    if (!window.AvengerStudio?.mount) {
      setError(area, 'ไม่พบ Avenger Studio module', 'avenger-studio');
      return;
    }
    setLoading(area, 'กำลังเปิด Avenger Studio...');
    try {
      await window.AvengerStudio.mount(area, {
        onStateChange(nextState) {
          App._studioReady = true;
          App._studioState = nextState || null;
        },
      });
      App._studioReady = true;
      App.readPresentationQueue().then(async (queued) => {
        for (const item of queued) {
          try {
            const imported = await window.AvengerStudio.importQueueItem(item);
            if (imported) await studioDbDeleteMany(STUDIO_QUEUE_STORE, [item.id]);
          } catch (err) {
            toast(`นำเข้า ${item.kind === 'table' ? 'ตาราง' : 'สไลด์'} ไม่สำเร็จ: ${err.message || err}`, 'error', 5000);
            throw err;
          }
        }
      }).catch(() => {});
    } catch (err) {
      setError(area, `เปิด Avenger Studio ไม่สำเร็จ\n${err.message || err}`, 'avenger-studio');
    }
    App._currentExport = null;
  },
};

/* ============================================================
   MODAL
   ============================================================ */
const Modal = {
  _rows: null,

  open(title, rows) {
    this._rows = rows;
    $('#modal-title').textContent = title;
    $('#modal-body').innerHTML = buildTable(rows);
    $('#modal-overlay').classList.remove('hidden');
  },

  close() {
    $('#modal-overlay').classList.add('hidden');
    this._rows = null;
  },

  exportCurrent() {
    if (!this._rows) return;
    exportExcel(this._rows, 'comparison');
  },
};

/* ============================================================
   ICON HELPERS (inline SVG)
   ============================================================ */
function searchIcon() {
  return `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>`;
}

function quickIcon(type) {
  const icons = {
    check: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>`,
    trend: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>`,
    cal:   `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`,
    globe: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>`,
    list:  `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>`,
    book:  `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>`,
  };
  return icons[type] || '';
}

/* ============================================================
   MAIN APP
   ============================================================ */
const App = {
  _currentExport: null,
  _currentTableExport: null,
  _toastTimer:    null,
  _studioReady:   false,
  _studioState:   null,

  /* Init after login */
  init() {
    $('#topbar-date').textContent = thaiDate();

    /* Nav */
    $$('.nav-item').forEach(el => {
      el.addEventListener('click', e => {
        e.preventDefault();
        App.navigate(el.dataset.page);
      });
    });

    /* Modal */
    $('#modal-close').addEventListener('click',  () => Modal.close());
    $('#btn-close-modal').addEventListener('click', () => Modal.close());
    $('#btn-export-modal').addEventListener('click', () => Modal.exportCurrent());
    $('#modal-overlay').addEventListener('click', e => {
      if (e.target === $('#modal-overlay')) Modal.close();
    });

    /* Logout */
    $('#btn-logout').addEventListener('click', () => {
      SheetsAPI.signOut();
      clearCache();
      State.selectedKeys.clear();
      State.selectedFunds = {};
      State.highlights = {};
      State.selectFundFilters = {
        category: '',
        type: '',
        style: '',
        dividend: '',
        query: '',
        pageSize: 25,
      };
      $('#app').classList.add('hidden');
      $('#login-screen').classList.remove('hidden');
    });

    /* Navigate to default page */
    App.navigate('dashboard');
  },

  navigate(page) {
    if (State.page === 'avenger-studio' && page !== 'avenger-studio' && window.AvengerStudio?.persistCurrent) {
      try { window.AvengerStudio.persistCurrent(); } catch { /* noop */ }
    }
    State.page        = page;
    State.sortCol     = null;
    State.sortDir     = 'asc';
    State.tablePage   = 1;
    State.pageSize    = page === 'select-fund'
      ? (State.selectFundFilters.pageSize || PAGE_SIZE_OPTIONS[0])
      : PAGE_SIZE_OPTIONS[0];
    App._currentExport = null;
    App._currentTableExport = null;
    if (page !== 'avenger-studio') App._studioReady = false;

    /* Update nav active state */
    $$('.nav-item').forEach(el =>
      el.classList.toggle('active', el.dataset.page === page)
    );

    /* Update page title */
    const titles = {
      'dashboard':         { title: 'แดชบอร์ด', subtitle: '' },
      'select-fund':       { title: 'เลือกกองทุน', subtitle: '' },
      'thai-annualized':   { title: 'กองทุนไทย Annualized Return', subtitle: '' },
      'thai-annualized-rank': { title: 'กองทุนไทย Annualized Rank', subtitle: '' },
      'thai-annualized-v2': { title: 'กองทุนไทย Annualized Return V2', subtitle: 'สลับมุมมอง Return และ Rank ได้' },
      'thai-calendar':     { title: 'กองทุนไทย Calendar Year', subtitle: '' },
      'master-annualized': { title: 'Master Fund Annualized Return V2', subtitle: 'จับคู่ด้วย ISIN และแยก Base Currency' },
      'master-annualized-v2': { title: 'Master Fund Annualized Return V2', subtitle: 'จับคู่ด้วย ISIN และแยก Base Currency' },
      'master-calendar':   { title: 'Master Fund Calendar Year', subtitle: '' },
      'master-placeholder-1': { title: 'ค่าธรรมเนียม', subtitle: 'เทียบ TER ของกองไทยกับ Ongoing Cost ของ Master Fund' },
      'master-placeholder-2': { title: 'Top 10 Holding', subtitle: '' },
      'master-placeholder-3': { title: 'Cost Efficiency Master Fund 5Y', subtitle: 'ดูค่าธรรมเนียมเทียบผลตอบแทนย้อนหลัง 5 ปี' },
      'master-placeholder-4': { title: 'ค่าธรรมเนียม 2', subtitle: '' },
      'master-placeholder-5': { title: 'Master Fund Menu 05', subtitle: 'เมนูสำรองรอใช้งาน' },
      'master-placeholder-6': { title: 'Master Fund Menu 06', subtitle: 'เมนูสำรองรอใช้งาน' },
      'guide':             { title: 'คู่มือการใช้งาน', subtitle: '' },
      'avenger-studio':   { title: 'Avenger Studio', subtitle: 'Planner Designer V7' },
    };
    const pageMeta = titles[page] || { title: page, subtitle: '' };
    $('#page-title').textContent = pageMeta.title;
    $('#page-subtitle').textContent = pageMeta.subtitle || '';
    $('#page-subtitle').classList.toggle('hidden', !pageMeta.subtitle);

    const area = $('#content-area');
    // Reset content-area styles (full-bleed pages อาจ override ไว้)
    area.style.padding       = '';
    area.style.overflow      = '';
    area.style.display       = '';
    area.style.flexDirection = '';

    switch (page) {
      case 'dashboard':         Pages.dashboard(area);                      break;
      case 'select-fund':       Pages.selectFund(area);                     break;
      case 'thai-annualized':   Pages.thaiAnnualized(area);                 break;
      case 'thai-annualized-rank': Pages.thaiAnnualizedRank(area);          break;
      case 'thai-annualized-v2': Pages.thaiAnnualizedV2(area);              break;
      case 'thai-calendar':     Pages.thaiCalendar(area);                   break;
      case 'master-annualized': Pages.masterAnnualizedV2(area);              break;
      case 'master-annualized-v2': Pages.masterAnnualizedV2(area);           break;
      case 'master-calendar':   Pages.masterCalendar(area);                  break;
      case 'master-placeholder-1': Pages.masterFees(area);                  break;
      case 'master-placeholder-2': Pages.masterMenu02(area);                break;
      case 'master-placeholder-3': Pages.masterMenu03(area);                break;
      case 'master-placeholder-4': Pages.masterFeesV2(area);                break;
      case 'master-placeholder-5': Pages.placeholder(area, 'Master Fund Menu 05'); break;
      case 'master-placeholder-6': Pages.placeholder(area, 'Master Fund Menu 06'); break;
      case 'guide':             Pages.guide(area);                          break;
      case 'avenger-studio':    Pages.avengerStudio(area);                  break;
      default:
        area.innerHTML = '<div class="card"><div class="state-box">ไม่พบหน้าที่ต้องการ</div></div>';
    }
  },

  async readPresentationQueue() {
    return await studioDbGetAll(STUDIO_QUEUE_STORE);
  },

  async queuePresentationSlide(item) {
    await studioDbPut(STUDIO_QUEUE_STORE, item);
    const queue = await this.readPresentationQueue();
    this._studioState = {
      ...(this._studioState || {}),
      pendingQueue: queue.length,
    };
    if (State.page === 'avenger-studio' && this._studioReady && window.AvengerStudio?.isMounted()) {
      const imported = await window.AvengerStudio.importQueueItem(item);
      if (imported) await studioDbDeleteMany(STUDIO_QUEUE_STORE, [item.id]);
    }
  },
};

/* ============================================================
   BOOT – Login & Auth
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => {

  if (isLocalMode()) {
    $('#user-name').textContent   = 'Local Data Mode';
    $('#user-email').textContent  = CONFIG.DATA_SOURCE === 'local_only'
      ? 'loading from /Data only'
      : 'loading from /Data first';
    $('#user-avatar').textContent = 'L';
    $('#login-screen').classList.add('hidden');
    $('#app').classList.remove('hidden');
    App.init();
    return;
  }

  /* ── DEV BYPASS ── */
  if (CONFIG.BYPASS_LOGIN) {
    $('#user-name').textContent   = 'Dev Mode';
    $('#user-email').textContent  = 'bypass login';
    $('#user-avatar').textContent = 'D';
    $('#login-screen').classList.add('hidden');
    $('#app').classList.remove('hidden');
    App.init();
    return;
  }

  const btnSignin = $('#btn-signin');

  const googleSvg = `<svg width="20" height="20" viewBox="0 0 48 48">
    <path fill="#4285F4" d="M45.12 24.5c0-1.56-.14-3.06-.4-4.5H24v8.51h11.84c-.51 2.75-2.06 5.08-4.39 6.64v5.52h7.11c4.16-3.83 6.56-9.47 6.56-16.17z"/>
    <path fill="#34A853" d="M24 46c5.94 0 10.92-1.97 14.56-5.33l-7.11-5.52c-1.97 1.32-4.49 2.1-7.45 2.1-5.73 0-10.58-3.87-12.32-9.07H4.34v5.7C7.96 41.07 15.4 46 24 46z"/>
    <path fill="#FBBC05" d="M11.68 28.18A13.9 13.9 0 0 1 10.9 24c0-1.45.25-2.86.78-4.18v-5.7H4.34A23.94 23.94 0 0 0 0 24c0 3.86.92 7.51 2.56 10.74l7.12-5.56z"/>
    <path fill="#EA4335" d="M24 9.75c3.23 0 6.13 1.11 8.41 3.29l6.31-6.31C34.91 3.19 29.93 1 24 1 15.4 1 7.96 5.93 4.34 13.26l7.34 5.7c1.74-5.2 6.59-9.21 12.32-9.21z"/>
  </svg> เข้าสู่ระบบด้วย Google`;

  function showLoginError(msg) {
    const el = $('#login-error');
    if (!el) return;
    el.innerHTML = '⚠️ ' + msg;
    el.classList.remove('hidden');
  }
  function hideLoginError() {
    const el = $('#login-error');
    if (el) el.classList.add('hidden');
  }

  btnSignin.addEventListener('click', async () => {
    hideLoginError();

    /* Guard: CLIENT_ID must be configured */
    if (!CONFIG.CLIENT_ID || CONFIG.CLIENT_ID.includes('YOUR_CLIENT_ID')) {
      showLoginError(
        'ยังไม่ได้ตั้งค่า <strong>CLIENT_ID</strong><br>' +
        'กรุณาเปิดไฟล์ <code>js/config.js</code> แล้วใส่ Client ID จาก Google Cloud Console<br>' +
        '<a href="README.md" target="_blank" style="color:#1a3c6e;font-weight:700">ดูขั้นตอนใน README.md →</a>'
      );
      return;
    }

    /* Guard: GIS must be loaded */
    if (typeof google === 'undefined' || !google.accounts) {
      showLoginError('Google Identity Services ยังไม่โหลดเสร็จ<br>กรุณารอสักครู่แล้วลองใหม่ หรือตรวจสอบการเชื่อมต่ออินเทอร์เน็ต');
      return;
    }

    /* Guard: must be served over HTTP/HTTPS, not file:// */
    if (window.location.protocol === 'file:') {
      showLoginError(
        'ไม่สามารถเปิดผ่าน File Explorer ได้<br>' +
        'กรุณาเปิดผ่าน Web Server เช่น<br>' +
        '• VS Code Live Server<br>' +
        '• <code>python -m http.server 8080</code> แล้วเปิด <code>http://localhost:8080</code>'
      );
      return;
    }

    btnSignin.disabled = true;
    btnSignin.innerHTML = `<span class="spin-sm"></span> กำลังเชื่อมต่อ...`;

    try {
      await SheetsAPI.requestToken();
      const info = await SheetsAPI.getUserInfo();

      const name    = info.name || info.email || 'User';
      const email   = info.email || '';
      const initial = (info.given_name || name).charAt(0).toUpperCase();

      $('#user-name').textContent   = name;
      $('#user-email').textContent  = email;
      $('#user-avatar').textContent = initial;

      $('#login-screen').classList.add('hidden');
      $('#app').classList.remove('hidden');

      App.init();

    } catch (e) {
      btnSignin.disabled = false;
      btnSignin.innerHTML = googleSvg;
      showLoginError('เข้าสู่ระบบไม่สำเร็จ: ' + e.message);
    }
  });
});
