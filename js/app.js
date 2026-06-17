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
const SELECT_FUND_PAGE_SIZE_OPTIONS = [100, 200];
const SELECT_FUND_DEFAULT_PAGE_SIZE = SELECT_FUND_PAGE_SIZE_OPTIONS[0];

/* ── Application State ── */
const State = {
  page:         'dashboard',
  sidebarCollapsed: false,
  sidebarOpen: false,
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
    pageSize: SELECT_FUND_DEFAULT_PAGE_SIZE,
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
  _pageDataMeta: {},
  _compareRows: null,
  top10HoldingV3: null,
  currentQuarter: null,      // ← active quarter tab (auto-detected from Sheets)
  availableQuarters: [],     // ← list of detected quarter tabs
  currentUser: { name: '', email: '' },
};

const SIDEBAR_PREF_KEY = 'fund-selection-sidebar-collapsed';

function isCompactViewport() {
  return window.matchMedia('(max-width: 640px)').matches;
}

function readSidebarPreference() {
  try {
    return localStorage.getItem(SIDEBAR_PREF_KEY) === '1';
  } catch {
    return false;
  }
}

function saveSidebarPreference(value) {
  try {
    localStorage.setItem(SIDEBAR_PREF_KEY, value ? '1' : '0');
  } catch {
    /* ignore preference persistence failures */
  }
}

/* ── Highlight color palette ── */
const HL_COLORS = [
  { name: 'เหลือง', bg: '#F9EC9D', dot: '#F9EC9D' },
  { name: 'เขียว',  bg: '#7ABC81', dot: '#7ABC81' },
  { name: 'ฟ้า',    bg: '#4AA3DC', dot: '#4AA3DC' },
  { name: 'ส้ม',    bg: '#EDB392', dot: '#EDB392' },
  { name: 'แดง',    bg: '#E7726F', dot: '#E7726F' },
];

const OF_POINT_COLORS = [
  { name: 'น้ำเงินเข้ม', dot: '#1a3c6e' },
  { name: 'ฟ้า',         dot: '#5aa2de' },
  { name: 'พีช',         dot: '#e9b48c' },
  { name: 'ทอง',         dot: '#e3a72f' },
  { name: 'ม่วง',        dot: '#8c3fe3' },
  { name: 'ฟ้าน้ำทะเล',  dot: '#2f99bf' },
  { name: 'เขียว',       dot: '#7dc182' },
  { name: 'ส้มเข้ม',     dot: '#dd6b20' },
  { name: 'ชมพูม่วง',    dot: '#d946ef' },
  { name: 'มิ้นต์',      dot: '#14b8a6' },
  { name: 'แดง',         dot: '#ef4444' },
  { name: 'คราม',        dot: '#6366f1' },
  { name: 'เขียวมะนาว',  dot: '#84cc16' },
  { name: 'อำพัน',       dot: '#f59e0b' },
  { name: 'ฟ้าใส',       dot: '#06b6d4' },
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
      master: 0,
      thai: 150,
      masterTer: 100,
      thaiTer: 100,
      combined: 100,
      spacer: 26,
      frontLoad: 100,
      backLoad: 100,
      initial: 100,
      subsequent: 100,
      fxHedging: 100,
      depositCurrency: 100,
      source: 100,
    },
  },
  default: {
    rowsPerSlide: 30,
  },
};

const TOP_10_HOLDING_API_URL = 'https://script.google.com/macros/s/AKfycbw6JSZPkHutKcBGDTpQyDZbcMFcKrU9VjJX5CV0jRdDtdxPCVzKJGIRrk3j9lzouAMO/exec';
const PRESENTATION_CLIPBOARD_TARGET_FONT_PT = 12;
const PRESENTATION_CLIPBOARD_POWERPOINT_SCALE_FIX = 2;
const PRESENTATION_CLIPBOARD_HTML_FONT_PT = PRESENTATION_CLIPBOARD_TARGET_FONT_PT * PRESENTATION_CLIPBOARD_POWERPOINT_SCALE_FIX;
const PRESENTATION_CLIPBOARD_FONT_FAMILY = "'TH Sarabun New','THSarabunNew','Sarabun',Arial,sans-serif";
const PRESENTATION_CLIPBOARD_TABLE_SIZE_SCALE = 1.15;
const PRESENTATION_CLIPBOARD_TABLE_MAX_WIDTH_PX = 0;

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

function firstNonEmptyColumnValue(rows, columnIndex) {
  if (!Array.isArray(rows) || columnIndex < 0) return '';
  for (const row of rows.slice(1)) {
    const value = String(row?.[columnIndex] ?? '').trim();
    if (value) return value;
  }
  return '';
}

function rememberPageDataMeta(pageKey, rows, cfg = {}) {
  const headers = Array.isArray(rows?.[0]) ? rows[0] : [];
  const fundSizeDateIdx = findColumnIndex(headers, [
    'Fund Size Date',
    'Fund size date',
    'FundSizeDate',
    'Fund SizeDate',
  ]);
  const fundSizeDate = firstNonEmptyColumnValue(rows, fundSizeDateIdx);
  State._pageDataMeta[pageKey] = {
    ...(State._pageDataMeta[pageKey] || {}),
    quarter: State.currentQuarter || cfg.tabName || '',
    fundSizeDate,
    hasFundSizeDate: fundSizeDateIdx >= 0,
  };
  return rows;
}

function getPercentrankFundSizeDate() {
  return State._pageDataMeta?.['select-fund']?.fundSizeDate || '';
}

function parseMonthIndexFromDateValue(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return null;

  const monthNames = {
    jan: 0, january: 0,
    feb: 1, february: 1,
    mar: 2, march: 2,
    apr: 3, april: 3,
    may: 4,
    jun: 5, june: 5,
    jul: 6, july: 6,
    aug: 7, august: 7,
    sep: 8, sept: 8, september: 8,
    oct: 9, october: 9,
    nov: 10, november: 10,
    dec: 11, december: 11,
  };

  const lower = raw.toLowerCase();
  for (const [name, index] of Object.entries(monthNames)) {
    if (new RegExp(`\\b${name}\\b`, 'i').test(lower)) return index;
  }

  const iso = raw.match(/\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b/);
  if (iso) {
    const month = Number(iso[2]);
    return month >= 1 && month <= 12 ? month - 1 : null;
  }

  const parts = raw.match(/\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b/);
  if (parts) {
    const first = Number(parts[1]);
    const second = Number(parts[2]);
    let month = second; // default to DD/MM/YYYY, common in Thai reporting.
    if (first <= 12 && second > 12) month = first;
    if (first > 12 && second <= 12) month = second;
    return month >= 1 && month <= 12 ? month - 1 : null;
  }

  return null;
}

function previousMonthLabelFromDateValue(value) {
  const monthIndex = parseMonthIndexFromDateValue(value);
  if (monthIndex === null) return '';
  const previousIndex = (monthIndex + 11) % 12;
  return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][previousIndex];
}

function sourceBadgeHtml(pageKey, sourceLabel = '') {
  if (!sourceLabel) return '';
  const isPercentrank = /percentrank freestyle/i.test(sourceLabel);
  const fundSizeDate = isPercentrank ? getPercentrankFundSizeDate() : '';
  const previousMonth = fundSizeDate ? previousMonthLabelFromDateValue(fundSizeDate) : '';
  const label = previousMonth ? `${sourceLabel} - ${previousMonth}` : sourceLabel;
  return `<span class="badge badge-source">แหล่งข้อมูลจาก: ${esc(label)}</span>`;
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
  const swatch = ['🟨', '🟩', '🟦', '🟧', '🟥'];
  const options = [
    `<option value="">ไม่เลือกสี</option>`,
    ...HL_COLORS.map((color, index) =>
      `<option value="${index}" ${String(currentValue) === String(index) ? 'selected' : ''}>${esc((swatch[index] || '■') + ' ' + color.name)}</option>`
    ),
  ];
  return `
    <select class="hl-select${selectedColor ? ' has-color' : ''}" data-fund="${esc(code)}" aria-label="เลือกสีไฮไลต์ของ ${esc(code)}"${selectedColor ? ` style="background:${selectedColor.bg};border-color:${selectedColor.dot};color:#1a3c6e"` : ''}>
      ${options.join('')}
    </select>`;
}
function buildHighlightDotPicker(code, currentValue, fallbackColor) {
  const selectedColor = currentValue !== undefined && OF_POINT_COLORS[currentValue]
    ? OF_POINT_COLORS[currentValue]
    : null;
  const currentColor = selectedColor?.dot || fallbackColor || '#cbd5e1';
  const swatches = [
    `<button class="of-palette-swatch is-clear" type="button" data-color-index="" title="ใช้สีอัตโนมัติ">×</button>`,
    ...OF_POINT_COLORS.map((color, index) =>
      `<button class="of-palette-swatch${String(currentValue) === String(index) ? ' is-active' : ''}" type="button" data-color-index="${index}" title="${esc(color.name)}" style="--swatch:${color.dot};background:${color.dot}"></button>`
    ),
  ];
  return `
    <div class="of-color-picker" data-fund="${esc(code)}">
      <button class="of-color-trigger${selectedColor ? ' is-custom' : ''}" type="button" aria-label="เลือกสีของ ${esc(code)}" style="--picker:${currentColor};background:${currentColor};border-color:${currentColor}"></button>
      <div class="of-color-palette" hidden>
        ${swatches.join('')}
      </div>
    </div>`;
}

function pageToolActions(pageKey, sourceLabel = '', extraActions = '') {
  const sourceBadge = getPageDataSourceBadge(pageKey);
  return `
    <div class="page-tools">
      <div class="page-tools-meta">
        ${sourceBadgeHtml(pageKey, sourceLabel)}
        ${sourceBadge ? `<span class="badge badge-data-origin">${esc(sourceBadge)}</span>` : ''}
        ${extraActions}
      </div>
      <div class="page-tools-actions">
        <button class="btn btn-ghost" id="btn-copy-table" title="คัดลอกข้อมูลแบบคงรูปแบบ">
          คัดลอกแบบมีฟอร์แมต
        </button>
        <button class="btn btn-primary" id="btn-send-table" title="ส่งข้อมูลไปทำ Presentation">
          ส่งข้อมูลไปทำ presentation
        </button>
      </div>
    </div>`;
}

function getPageDataSourceBadge(pageKey) {
  return State._pageDataSource?.[pageKey] || '';
}

function presentationClipboardColumnWidth(column) {
  if (column?.widthPx) return `${Math.max(24, Number(column.widthPx) || 0)}px`;
  const weight = Number(column?.weight || 1);
  return `${Math.max(42, Math.round(weight * 44))}px`;
}

function presentationClipboardTableWidth(columns = []) {
  const total = columns.reduce((sum, column) => {
    const width = column?.widthPx
      ? Math.max(24, Number(column.widthPx) || 0)
      : Math.max(42, Math.round(Number(column?.weight || 1) * 44));
    return sum + width;
  }, 0);
  return Math.max(980, total + 24);
}

function presentationCellText(cell) {
  if (!cell) return '';
  const fragments = Array.isArray(cell.fragments) ? cell.fragments : [];
  if (fragments.length) {
    return fragments.map(fragment => String(fragment?.text ?? '')).join(' ').trim();
  }
  return String(cell.text ?? '').trim();
}

function buildPresentationClipboardPlainText(payload = {}) {
  const lines = [];
  const title = String(payload.title || '').trim();
  const subtitle = String(payload.subtitle || '').trim();
  const source = String(payload.source || '').trim();
  const columns = Array.isArray(payload.columns) ? payload.columns : [];
  const rows = Array.isArray(payload.rows) ? payload.rows : [];

  if (title) lines.push(title);
  if (subtitle) lines.push(subtitle);
  if (source) lines.push(`ที่มา : ${source}`);
  if (columns.length) {
    lines.push(columns.map(column => String(column?.label || '')).join('\t'));
  }
  rows.forEach(row => {
    const cells = Array.isArray(row?.cells) ? row.cells : [];
    lines.push(cells.map(presentationCellText).join('\t'));
  });
  return lines.join('\n');
}

function presentationClipboardFontStyle() {
  return [
    `font-family:${PRESENTATION_CLIPBOARD_FONT_FAMILY}`,
    "mso-ascii-font-family:'TH Sarabun New'",
    "mso-hansi-font-family:'TH Sarabun New'",
    "mso-bidi-font-family:'TH Sarabun New'",
    "mso-fareast-font-family:'TH Sarabun New'",
    `font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}pt`,
    `mso-ansi-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
    `mso-fareast-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
    `mso-bidi-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
  ].join(';');
}

const RENDERED_CLIPBOARD_STYLE_PROPS = [
  'background-color',
  'border-bottom-color',
  'border-bottom-style',
  'border-bottom-width',
  'border-collapse',
  'border-left-color',
  'border-left-style',
  'border-left-width',
  'border-right-color',
  'border-right-style',
  'border-right-width',
  'border-spacing',
  'border-top-color',
  'border-top-style',
  'border-top-width',
  'box-sizing',
  'color',
  'font-style',
  'font-variant-numeric',
  'font-weight',
  'margin-bottom',
  'margin-left',
  'margin-right',
  'margin-top',
  'padding-bottom',
  'padding-left',
  'padding-right',
  'padding-top',
  'table-layout',
  'text-align',
  'vertical-align',
  'white-space',
  'width',
  'word-break',
];

function appendOfficeFontHints(el) {
  el.setAttribute('lang', 'TH');
  el.setAttribute('face', 'TH Sarabun New');
  el.style.setProperty('font-family', PRESENTATION_CLIPBOARD_FONT_FAMILY, 'important');
  el.style.setProperty('font-size', `${PRESENTATION_CLIPBOARD_HTML_FONT_PT}pt`, 'important');
  el.style.setProperty('mso-ascii-font-family', "'TH Sarabun New'");
  el.style.setProperty('mso-hansi-font-family', "'TH Sarabun New'");
  el.style.setProperty('mso-bidi-font-family', "'TH Sarabun New'");
  el.style.setProperty('mso-fareast-font-family', "'TH Sarabun New'");
  el.style.setProperty('mso-ansi-font-size', `${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`);
  el.style.setProperty('mso-fareast-font-size', `${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`);
  el.style.setProperty('mso-bidi-font-size', `${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`);
}

function isRenderedActiveSortCell(sourceEl) {
  return sourceEl.matches?.('.report-sort.is-active, .sf-sort.is-active, .th-asc, .th-desc')
    || !!sourceEl.querySelector?.('.sort-label.is-active');
}

function applyRenderedActiveSortStyle(cloneEl) {
  const activeBg = '#f7d774';
  const activeFg = '#4e3500';
  cloneEl.setAttribute('bgcolor', activeBg);
  cloneEl.style.setProperty('background-color', activeBg);
  cloneEl.style.setProperty('border-color', '#d79a12');
  cloneEl.style.setProperty('color', activeFg);
  cloneEl.style.setProperty('font-weight', '700');
  cloneEl.style.setProperty('mso-highlight', activeBg);
  cloneEl.querySelectorAll('.sort-label, .sort-text').forEach(child => {
    child.style.setProperty('background-color', activeBg);
    child.style.setProperty('box-shadow', 'none');
    child.style.setProperty('color', activeFg);
    child.style.setProperty('font-weight', '700');
    child.style.setProperty('mso-highlight', activeBg);
  });
}

function renderedClipboardCssColor(value, fallback = '') {
  if (!value || value === 'transparent' || value === 'rgba(0, 0, 0, 0)') return fallback;
  return value;
}

function renderedClipboardOfficeColor(value, fallback = '') {
  const color = renderedClipboardCssColor(value, fallback);
  const match = String(color).match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (!match) return color;
  return `#${match.slice(1, 4).map(part => Number(part).toString(16).padStart(2, '0')).join('')}`;
}

function renderedClipboardScaledSize(value, scale = PRESENTATION_CLIPBOARD_TABLE_SIZE_SCALE) {
  const size = Number(value) || 0;
  if (size <= 0) return 0;
  return Math.max(1, Math.round(size * scale));
}

function renderedClipboardTableScale(tableWidth) {
  const baseScale = PRESENTATION_CLIPBOARD_TABLE_SIZE_SCALE;
  const maxWidth = Number(PRESENTATION_CLIPBOARD_TABLE_MAX_WIDTH_PX) || 0;
  if (maxWidth > 0 && tableWidth > 0) return Math.min(baseScale, maxWidth / tableWidth);
  return baseScale;
}

function renderedClipboardBorderStyle(computed) {
  const color = renderedClipboardOfficeColor(computed.getPropertyValue('border-top-color'), '#dbe4f0');
  const width = computed.getPropertyValue('border-top-width') || '1px';
  const style = computed.getPropertyValue('border-top-style') || 'solid';
  return `${width} ${style === 'none' ? 'solid' : style} ${color}`;
}

function renderedClipboardCellText(cell) {
  const clone = cell.cloneNode(true);
  clone.querySelectorAll('.sort-indicator').forEach(el => el.remove());
  return (clone.textContent || '').replace(/\s+/g, ' ').trim();
}

function renderedClipboardInlineText(value = '') {
  return esc(String(value || '').replace(/\s+/g, ' '));
}

function renderedClipboardInlineNodeHtml(node, inheritedColor = '#0f172a', sizeScale = PRESENTATION_CLIPBOARD_TABLE_SIZE_SCALE) {
  if (!node) return '';
  if (node.nodeType === Node.TEXT_NODE) return renderedClipboardInlineText(node.textContent);
  if (node.nodeType !== Node.ELEMENT_NODE) return '';

  const el = node;
  if (el.classList?.contains('sort-indicator')) return '';
  const tagName = el.tagName?.toLowerCase();
  if (tagName === 'br') return '<br>';

  const computed = window.getComputedStyle(el);
  const bg = renderedClipboardOfficeColor(computed.getPropertyValue('background-color'), '');
  const color = renderedClipboardOfficeColor(computed.getPropertyValue('color'), inheritedColor);
  const fontWeight = Number.parseInt(computed.getPropertyValue('font-weight'), 10) || 500;
  const isChip = el.classList?.contains('linked-fund-chip');
  const hasHighlight = !!bg && bg !== '#ffffff';
  const inner = Array.from(el.childNodes || [])
    .map(child => renderedClipboardInlineNodeHtml(child, color, sizeScale))
    .join('')
    || renderedClipboardInlineText(el.textContent);

  if (tagName === 'a') {
    const href = String(el.getAttribute('href') || '').trim();
    const linkStyle = [
      `font-family:${PRESENTATION_CLIPBOARD_FONT_FAMILY}`,
      `font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}pt`,
      "mso-ascii-font-family:'TH Sarabun New'",
      "mso-hansi-font-family:'TH Sarabun New'",
      "mso-bidi-font-family:'TH Sarabun New'",
      "mso-fareast-font-family:'TH Sarabun New'",
      `mso-ansi-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
      `mso-fareast-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
      `mso-bidi-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
      `color:${color || '#3559d7'}`,
      'font-weight:700',
      'text-decoration:underline',
    ].join(';');
    return href
      ? `<a href="${esc(href)}" style="${linkStyle}">${inner || '&nbsp;'}</a>`
      : `<span style="${linkStyle}">${inner || '&nbsp;'}</span>`;
  }

  if (!isChip && !hasHighlight) return inner;

  const padY = renderedClipboardScaledSize(2, sizeScale);
  const padX = renderedClipboardScaledSize(6, sizeScale);
  const marginRight = renderedClipboardScaledSize(4, sizeScale);
  const style = [
    `font-family:${PRESENTATION_CLIPBOARD_FONT_FAMILY}`,
    `font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}pt`,
    "mso-ascii-font-family:'TH Sarabun New'",
    "mso-hansi-font-family:'TH Sarabun New'",
    "mso-bidi-font-family:'TH Sarabun New'",
    "mso-fareast-font-family:'TH Sarabun New'",
    `mso-ansi-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
    `mso-fareast-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
    `mso-bidi-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
    'display:inline-block',
    `padding:${padY}px ${padX}px`,
    `margin:1px ${marginRight}px 1px 0`,
    `background-color:${bg || 'transparent'}`,
    bg ? `mso-highlight:${bg}` : '',
    `color:${color || inheritedColor}`,
    `font-weight:${fontWeight >= 600 ? 700 : 500}`,
    'line-height:1.1',
    'white-space:nowrap',
  ].filter(Boolean).join(';');

  return `<span style="${style}">${inner || '&nbsp;'}</span>`;
}

function renderedClipboardCellContentHtml(cell, sizeScale = PRESENTATION_CLIPBOARD_TABLE_SIZE_SCALE) {
  const hasRichInlineContent = !!Array.from(cell.querySelectorAll?.('*') || []).some(el => {
    if (el.matches?.('a[href]')) return true;
    if (el.classList?.contains('linked-fund-chip')) return true;
    const bg = renderedClipboardOfficeColor(window.getComputedStyle(el).getPropertyValue('background-color'), '');
    return !!bg && bg !== '#ffffff';
  });
  if (!hasRichInlineContent) {
    return esc(renderedClipboardCellText(cell)).replace(/\n/g, '<br>') || '&nbsp;';
  }
  const computed = window.getComputedStyle(cell);
  const color = renderedClipboardOfficeColor(computed.getPropertyValue('color'), '#0f172a');
  const html = Array.from(cell.childNodes || [])
    .map(child => renderedClipboardInlineNodeHtml(child, color, sizeScale))
    .join('')
    .trim();
  return html || '&nbsp;';
}

function renderedClipboardCellHtml(cell, rowHeight = 0, sizeScale = PRESENTATION_CLIPBOARD_TABLE_SIZE_SCALE, preferredWidth = 0) {
  const tag = cell.tagName.toLowerCase() === 'th' ? 'th' : 'td';
  const computed = window.getComputedStyle(cell);
  const isActiveSort = isRenderedActiveSortCell(cell);
  const bg = isActiveSort
    ? '#f7d774'
    : renderedClipboardOfficeColor(computed.getPropertyValue('background-color'), '#ffffff');
  const color = isActiveSort
    ? '#4e3500'
    : renderedClipboardOfficeColor(computed.getPropertyValue('color'), '#0f172a');
  const textAlign = computed.getPropertyValue('text-align') || 'center';
  const align = ['left', 'right', 'center'].includes(textAlign) ? textAlign : 'center';
  const computedWeight = Number.parseInt(computed.getPropertyValue('font-weight'), 10);
  const weight = isActiveSort || tag === 'th' || computedWeight >= 600 ? 700 : 500;
  const colspan = Number(cell.getAttribute('colspan') || 1);
  const rowspan = Number(cell.getAttribute('rowspan') || 1);
  const attrs = [
    colspan > 1 ? `colspan="${colspan}"` : '',
    rowspan > 1 ? `rowspan="${rowspan}"` : '',
    `align="${align}"`,
    'valign="middle"',
    `bgcolor="${bg}"`,
    preferredWidth > 0 ? `width="${preferredWidth}"` : '',
    'lang="TH"',
  ].filter(Boolean).join(' ');
  const style = [
    `font-family:${PRESENTATION_CLIPBOARD_FONT_FAMILY}`,
    `font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}pt`,
    "mso-ascii-font-family:'TH Sarabun New'",
    "mso-hansi-font-family:'TH Sarabun New'",
    "mso-bidi-font-family:'TH Sarabun New'",
    "mso-fareast-font-family:'TH Sarabun New'",
    `mso-ansi-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
    `mso-fareast-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
    `mso-bidi-font-size:${PRESENTATION_CLIPBOARD_HTML_FONT_PT}.0pt`,
    `background-color:${bg}`,
    `color:${color}`,
    `font-weight:${weight}`,
    `border:${renderedClipboardBorderStyle(computed)}`,
    preferredWidth > 0 ? `width:${preferredWidth}px` : '',
    preferredWidth > 0 ? `min-width:${preferredWidth}px` : '',
    preferredWidth > 0 ? `max-width:${preferredWidth}px` : '',
    preferredWidth > 0 ? 'mso-width-source:userset' : '',
    rowHeight > 0 ? `height:${rowHeight}px` : '',
    rowHeight > 0 ? 'mso-height-source:userset' : '',
    `padding:${renderedClipboardScaledSize(5, sizeScale)}px ${renderedClipboardScaledSize(8, sizeScale)}px`,
    `text-align:${align}`,
    'vertical-align:middle',
    'line-height:1.15',
    'white-space:nowrap',
    'word-break:keep-all',
  ].filter(Boolean).join(';');
  const content = renderedClipboardCellContentHtml(cell, sizeScale);
  return `<${tag} ${attrs} style="${style}">${content}</${tag}>`;
}

function getRenderedClipboardColumnWidths(sourceTable) {
  const rows = Array.from(sourceTable.rows || []);
  if (!rows.length) return [];

  const candidates = rows
    .map(row => {
      const cells = Array.from(row.cells || []);
      if (!cells.length) return null;
      const parts = cells.map(cell => {
        const span = Math.max(1, Number(cell.getAttribute('colspan') || 1));
        const width = Math.ceil(cell.getBoundingClientRect().width || cell.offsetWidth || 0);
        return { span, width };
      });
      const colCount = parts.reduce((sum, part) => sum + part.span, 0);
      const mergedCount = parts.reduce((sum, part) => sum + (part.span > 1 ? 1 : 0), 0);
      return { parts, colCount, mergedCount };
    })
    .filter(Boolean);

  if (!candidates.length) return [];

  // Prefer rows that represent real columns (no colspan), then widest column count.
  const best = candidates
    .slice()
    .sort((a, b) => {
      if (a.mergedCount === 0 && b.mergedCount !== 0) return -1;
      if (a.mergedCount !== 0 && b.mergedCount === 0) return 1;
      if (b.colCount !== a.colCount) return b.colCount - a.colCount;
      return a.mergedCount - b.mergedCount;
    })[0];

  const widths = new Array(best.colCount).fill(0);
  let colIndex = 0;
  best.parts.forEach(part => {
    const perColumnWidth = part.width > 0 ? Math.max(18, Math.round(part.width / part.span)) : 0;
    for (let i = 0; i < part.span; i += 1) {
      widths[colIndex + i] = perColumnWidth;
    }
    colIndex += part.span;
  });

  // Fill any missing width slot from all rows as fallback.
  const fallback = new Array(best.colCount).fill(null).map(() => ({ sum: 0, weight: 0 }));
  candidates.forEach(candidate => {
    let idx = 0;
    candidate.parts.forEach(part => {
      const per = part.width > 0 ? Math.round(part.width / part.span) : 0;
      if (per > 0) {
        const weight = part.span === 1 ? 1 : 0.2;
        for (let i = 0; i < part.span && idx + i < fallback.length; i += 1) {
          fallback[idx + i].sum += per * weight;
          fallback[idx + i].weight += weight;
        }
      }
      idx += part.span;
    });
  });

  return widths.map((width, idx) => {
    if (width > 0) return width;
    const entry = fallback[idx];
    if (entry && entry.weight > 0) return Math.max(18, Math.round(entry.sum / entry.weight));
    return 18;
  });
}

function getRenderedClipboardColumnCount(sourceTable) {
  return Array.from(sourceTable.rows || []).reduce((maxCount, row) => {
    const count = Array.from(row.cells || []).reduce((sum, cell) => {
      return sum + Math.max(1, Number(cell.getAttribute('colspan') || 1));
    }, 0);
    return Math.max(maxCount, count);
  }, 0);
}

function getRenderedClipboardColumnWidthsFromPayload(payload, expectedCount = 0) {
  const columns = Array.isArray(payload?.columns) ? payload.columns : [];
  if (!columns.length) return [];
  const widths = columns.map(column => {
    const widthPx = Number(column?.widthPx || 0);
    if (widthPx > 0) return Math.max(18, Math.round(widthPx));
    const weight = Number(column?.weight || 0);
    if (weight > 0) return Math.max(18, Math.round(weight * 48));
    return 0;
  });
  if (expectedCount > 0 && widths.length !== expectedCount) return [];
  if (widths.some(width => width <= 0)) return [];
  return widths;
}

function buildCleanRenderedClipboardTable(sourceTable, payload = {}) {
  const rows = Array.from(sourceTable.rows || []);
  if (!rows.length) return '';
  const expectedColCount = getRenderedClipboardColumnCount(sourceTable);
  const payloadColumnWidths = getRenderedClipboardColumnWidthsFromPayload(payload, expectedColCount);
  const rawColumnWidths = payloadColumnWidths.length
    ? payloadColumnWidths
    : getRenderedClipboardColumnWidths(sourceTable);
  const rawTableWidth = rawColumnWidths.reduce((sum, width) => sum + width, 0);
  const sizeScale = renderedClipboardTableScale(rawTableWidth);
  const columnWidths = rawColumnWidths.map(width => renderedClipboardScaledSize(width, sizeScale));
  const tableWidth = columnWidths.reduce((sum, width) => sum + width, 0);
  const colgroup = columnWidths.length
    ? `<colgroup>${columnWidths.map(width => `<col style="width:${width}px;">`).join('')}</colgroup>`
    : '';
  const tableSizeStyle = tableWidth > 0
    ? `table-layout:fixed;width:${tableWidth}px;`
    : 'table-layout:auto;width:auto;';
  const rowSpanTrack = new Array(columnWidths.length).fill(0);
  return `<table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #dbe4f0;${tableSizeStyle}background:#ffffff;${presentationClipboardFontStyle()};">
    ${colgroup}
    ${rows.map(row => {
      for (let i = 0; i < rowSpanTrack.length; i += 1) {
        if (rowSpanTrack[i] > 0) rowSpanTrack[i] -= 1;
      }
      const rawRowHeight = Math.ceil(row.getBoundingClientRect().height || row.offsetHeight || 0);
      const rowHeight = renderedClipboardScaledSize(rawRowHeight, sizeScale);
      const rowStyle = rowHeight > 0 ? ` style="height:${rowHeight}px;mso-height-source:userset"` : '';
      let colIndex = 0;
      const cellsHtml = Array.from(row.cells || []).map(cell => {
        while (rowSpanTrack[colIndex] > 0) colIndex += 1;
        const colspan = Math.max(1, Number(cell.getAttribute('colspan') || 1));
        const rowspan = Math.max(1, Number(cell.getAttribute('rowspan') || 1));
        const preferredWidth = columnWidths
          .slice(colIndex, colIndex + colspan)
          .reduce((sum, width) => sum + (Number(width) || 0), 0);
        if (rowspan > 1) {
          for (let i = 0; i < colspan; i += 1) {
            rowSpanTrack[colIndex + i] = Math.max(rowSpanTrack[colIndex + i] || 0, rowspan - 1);
          }
        }
        colIndex += colspan;
        return renderedClipboardCellHtml(cell, rowHeight, sizeScale, preferredWidth);
      }).join('');
      return `<tr${rowStyle}>${cellsHtml}</tr>`;
    }).join('')}
  </table>`;
}

function applyRenderedTableOfficeAttributes(sourceEl, cloneEl) {
  const tag = sourceEl.tagName.toLowerCase();
  if (tag !== 'th' && tag !== 'td') return;
  const computed = window.getComputedStyle(sourceEl);
  const textAlign = computed.getPropertyValue('text-align') || 'center';
  const normalizedAlign = ['left', 'right', 'center'].includes(textAlign) ? textAlign : 'center';
  cloneEl.setAttribute('align', normalizedAlign);
  cloneEl.setAttribute('valign', 'middle');
  cloneEl.style.setProperty('text-align', normalizedAlign, 'important');
  cloneEl.style.setProperty('vertical-align', 'middle', 'important');
}

function inlineRenderedClipboardStyles(sourceEl, cloneEl) {
  if (!(sourceEl instanceof Element) || !(cloneEl instanceof Element)) return;
  const computed = window.getComputedStyle(sourceEl);
  RENDERED_CLIPBOARD_STYLE_PROPS.forEach(prop => {
    const value = computed.getPropertyValue(prop);
    if (value) cloneEl.style.setProperty(prop, value);
  });

  const tag = sourceEl.tagName.toLowerCase();
  if (tag === 'table') {
    const rect = sourceEl.getBoundingClientRect();
    if (rect.width > 0) {
      const width = `${Math.ceil(rect.width)}px`;
      cloneEl.style.setProperty('width', width);
      cloneEl.style.setProperty('min-width', width);
    }
  } else if (tag === 'th' || tag === 'td') {
    cloneEl.style.removeProperty('min-width');
  }

  appendOfficeFontHints(cloneEl);
  applyRenderedTableOfficeAttributes(sourceEl, cloneEl);

  Array.from(sourceEl.children).forEach((sourceChild, idx) => {
    const cloneChild = cloneEl.children[idx];
    if (cloneChild) inlineRenderedClipboardStyles(sourceChild, cloneChild);
  });

  if ((tag === 'th' || tag === 'td') && isRenderedActiveSortCell(sourceEl)) {
    applyRenderedActiveSortStyle(cloneEl);
  }
}

function buildRenderedTableClipboardHtml(card, payload = {}) {
  const sourceTable = card?.querySelector?.('table');
  if (!sourceTable) return '';

  const title = String(payload.title || '').trim();
  const subtitle = String(payload.subtitle || '').trim();
  const source = String(payload.source || '').trim();
  const fontFamily = "font-family:'TH Sarabun New','THSarabunNew','Sarabun',Arial,sans-serif";
  const clipboardFontStyle = presentationClipboardFontStyle();
  const tableHtml = buildCleanRenderedClipboardTable(sourceTable, payload);
  if (!tableHtml) return '';

  return `<!DOCTYPE html>
  <html lang="th">
    <head>
      <meta charset="utf-8">
    </head>
    <body lang="TH" style="margin:0;padding:0;background:#ffffff;${clipboardFontStyle};color:#334155;mso-fareast-language:TH;">
      <div style="display:inline-block;max-width:none;${clipboardFontStyle};">
        ${title ? `<div lang="TH" style="${fontFamily};${clipboardFontStyle};line-height:1.2;font-weight:700;color:#1a3c6e;margin:0 0 4px;mso-fareast-language:TH;">${esc(title)}</div>` : ''}
        ${subtitle ? `<div lang="TH" style="${fontFamily};${clipboardFontStyle};line-height:1.2;color:#64748b;margin:0 0 4px;mso-fareast-language:TH;">${esc(subtitle)}</div>` : ''}
        ${source ? `<div lang="TH" style="${fontFamily};${clipboardFontStyle};line-height:1.2;color:#64748b;margin:0 0 10px;mso-fareast-language:TH;">ที่มา : ${esc(source)}</div>` : ''}
        ${tableHtml}
      </div>
    </body>
  </html>`;
}

function buildPresentationClipboardHtml(payload = {}) {
  const title = String(payload.title || '').trim();
  const subtitle = String(payload.subtitle || '').trim();
  const source = String(payload.source || '').trim();
  const headerGroups = Array.isArray(payload.headerGroups) ? payload.headerGroups : [];
  const columns = Array.isArray(payload.columns) ? payload.columns : [];
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const colWidths = columns.map(presentationClipboardColumnWidth);
  const tableWidth = presentationClipboardTableWidth(columns);
  const fontFamily = "font-family:'TH Sarabun New','THSarabunNew','Sarabun',Arial,sans-serif";
  const clipboardFontStyle = presentationClipboardFontStyle();

  const renderCellContent = (cell) => {
    if (!cell) return '&nbsp;';
    const fragments = Array.isArray(cell.fragments) ? cell.fragments : [];
    if (fragments.length) {
      return fragments.map(fragment => {
        const fragBg = fragment?.bg || '';
        const fragColor = fragment?.color || cell.color || '#334155';
        const fragStrong = fragment?.strong || cell.strong;
        const fragText = esc(String(fragment?.text ?? ''));
        return `
          <span style="
            ${fontFamily};
            display:inline-block;
            margin:2px 4px 2px 0;
            padding:1px 7px;
            border-radius:999px;
            background:${fragBg || 'transparent'};
            color:${fragColor};
            font-weight:${fragStrong ? 700 : 500};
            ${clipboardFontStyle};
            line-height:1.2;
            white-space:nowrap;
          ">${fragText || '&nbsp;'}</span>`;
      }).join('');
    }
    const text = esc(String(cell.text ?? '')).replace(/\n/g, '<br>');
    if (cell.href) {
      return `<a href="${esc(cell.href)}" style="${fontFamily};${clipboardFontStyle};color:${cell.color || '#3559d7'};font-weight:700;text-decoration:underline;">${text || '&nbsp;'}</a>`;
    }
    return text || '&nbsp;';
  };

  const cellBaseStyle = (cell, align = 'center') => {
    const bg = cell?.bg || 'transparent';
    const color = cell?.color || '#334155';
    const weight = cell?.strong ? 700 : 500;
    return [
      fontFamily,
      'border:1px solid #dbe4f0',
      'padding:7px 9px',
      `background:${bg}`,
      `color:${color}`,
      `font-weight:${weight}`,
      clipboardFontStyle,
      `text-align:${cell?.align === 'left' || align === 'left' ? 'left' : 'center'}`,
      'vertical-align:middle',
      'line-height:1.25',
      'word-break:break-word',
      'white-space:normal',
    ].join(';');
  };

  let html = `<!DOCTYPE html>
  <html lang="th">
    <head>
      <meta charset="utf-8">
    </head>
    <body lang="TH" style="margin:0;padding:0;background:#ffffff;${fontFamily};${clipboardFontStyle};color:#334155;mso-fareast-language:TH;">
      <div style="display:inline-block;max-width:100%;">
        ${title ? `<div lang="TH" style="${fontFamily};${clipboardFontStyle};line-height:1.2;font-weight:700;color:#1a3c6e;margin:0 0 4px;mso-fareast-language:TH;">${esc(title)}</div>` : ''}
        ${subtitle ? `<div lang="TH" style="${fontFamily};${clipboardFontStyle};line-height:1.2;color:#64748b;margin:0 0 4px;mso-fareast-language:TH;">${esc(subtitle)}</div>` : ''}
        ${source ? `<div lang="TH" style="${fontFamily};${clipboardFontStyle};line-height:1.2;color:#64748b;margin:0 0 10px;mso-fareast-language:TH;">ที่มา : ${esc(source)}</div>` : ''}
        <table style="border-collapse:collapse;table-layout:fixed;width:${tableWidth}px;min-width:${tableWidth}px;border:1px solid #cbd5e1;${clipboardFontStyle};line-height:1.25;background:#ffffff;">
          ${colWidths.length ? `<colgroup>${colWidths.map(width => `<col style="width:${width};">`).join('')}</colgroup>` : ''}
          <thead>
            ${headerGroups.length ? `
              <tr>
                ${headerGroups.map(group => `
                  <th colspan="${group.span || 1}" style="
                    ${fontFamily};
                    border:1px solid #dbe4f0;
                    padding:9px 12px;
                    background:${group.bg || '#dbe4f0'};
                    color:${group.color || '#334155'};
                    ${clipboardFontStyle};
                    font-weight:700;
                    text-align:center;
                    vertical-align:middle;
                    line-height:1.15;
                  ">${esc(String(group.label || '')) || '&nbsp;'}</th>
                `).join('')}
              </tr>
            ` : ''}
            <tr>
              ${columns.map((column) => `
                <th style="
                  ${cellBaseStyle({ bg: column.bg || '#3f5d8c', color: column.color || '#ffffff', strong: true, align: column.align }, column.align)}
                  ${fontFamily};
                  ${clipboardFontStyle};
                  background:${column.bg || '#3f5d8c'};
                  color:${column.color || '#ffffff'};
                  font-weight:700;
                  padding:9px 12px;
                ">${esc(String(column?.label || '')) || '&nbsp;'}</th>
              `).join('')}
            </tr>
          </thead>
          <tbody>
            ${rows.map((row, rowIndex) => `
              <tr>
                ${(Array.isArray(row?.cells) ? row.cells : []).map((cell, cellIndex) => {
                  const column = columns[cellIndex] || {};
                  return `
                    <td style="${cellBaseStyle(cell, column.align)}${column.widthPx ? `;width:${presentationClipboardColumnWidth(column)}` : ''}">
                      ${renderCellContent(cell)}
                    </td>`;
                }).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </body>
  </html>`;

  return html;
}

async function copyPresentationPayloadToClipboard(payload, options = {}) {
  const plainText = buildPresentationClipboardPlainText(payload);
  const renderedHtml = buildRenderedTableClipboardHtml(options.card, payload);
  const html = renderedHtml || buildPresentationClipboardHtml(payload);

  return copyHtmlToClipboard(html, plainText, renderedHtml ? 'rendered-html' : 'html');
}

async function copyHtmlToClipboard(html, plainText = '', htmlMode = 'html', options = {}) {
  if (navigator.clipboard?.write && window.ClipboardItem) {
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/html': new Blob([html], { type: 'text/html' }),
          'text/plain': new Blob([plainText], { type: 'text/plain' }),
        }),
      ]);
      return htmlMode;
    } catch (err) {
      if (options.allowTextFallback === false) {
        throw err;
      }
      /* fall through to plain-text fallback */
    }
  }

  if (options.allowTextFallback === false) {
    throw new Error('ไม่สามารถเขียน HTML ลงคลิปบอร์ดได้');
  }

  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(plainText);
    return 'text';
  }

  throw new Error('ไม่สามารถคัดลอกลงคลิปบอร์ดได้');
}

function cloneNodeForCapture(el) {
  const cloned = el.cloneNode(true);
  const sourceCanvases = el.matches?.('canvas')
    ? [el]
    : Array.from(el.querySelectorAll('canvas'));
  const clonedCanvases = cloned.matches?.('canvas')
    ? [cloned]
    : Array.from(cloned.querySelectorAll('canvas'));

  sourceCanvases.forEach((canvas, idx) => {
    const clonedCanvas = clonedCanvases[idx];
    if (!clonedCanvas) return;
    try {
      const img = document.createElement('img');
      img.src = canvas.toDataURL('image/png');
      img.alt = '';
      img.style.width = `${canvas.width || canvas.clientWidth || clonedCanvas.clientWidth}px`;
      img.style.height = `${canvas.height || canvas.clientHeight || clonedCanvas.clientHeight}px`;
      img.style.display = 'block';
      img.style.maxWidth = '100%';
      img.width = canvas.width || canvas.clientWidth || 0;
      img.height = canvas.height || canvas.clientHeight || 0;
      clonedCanvas.replaceWith(img);
    } catch {
      /* ignore canvas snapshot errors and keep fallback clone */
    }
  });

  return cloned;
}

async function elementToImageBlob(el) {
  const rect = el.getBoundingClientRect();
  const cloned = cloneNodeForCapture(el);
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
  const cloned = cloneNodeForCapture(el);
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
  const copyBtn = $('#btn-copy-table', area);
  if (tableBtn) {
    const supportsTable = typeof App._currentTableExport === 'function';
    const supportsImage = typeof App._currentImageExport === 'function';
    const supported = supportsTable || supportsImage;
    tableBtn.disabled = !supported;
    tableBtn.title = supported
      ? (supportsImage ? 'ส่งกราฟหรือรายงานไปทำ Presentation' : 'ส่งข้อมูลไปทำ Presentation')
      : 'หน้านี้ยังไม่รองรับการส่งออกไป Presentation';
    tableBtn.style.opacity = supported ? '' : '.55';
    tableBtn.style.cursor = supported ? '' : 'not-allowed';
  }
  if (copyBtn) {
    const supportsTable = typeof App._currentTableExport === 'function';
    const supportsClipboard = typeof App._currentClipboardExport === 'function';
    const supported = supportsTable || supportsClipboard;
    copyBtn.disabled = !supported;
    copyBtn.title = supported
      ? 'คัดลอกข้อมูลพร้อมสีและรูปแบบลงคลิปบอร์ด'
      : 'หน้านี้ยังไม่รองรับการคัดลอกแบบมีฟอร์แมต';
    copyBtn.textContent = 'คัดลอกแบบมีฟอร์แมต';
    copyBtn.style.opacity = supported ? '' : '.55';
    copyBtn.style.cursor = supported ? '' : 'not-allowed';
  }

  $('#btn-send-table', area)?.addEventListener('click', async () => {
    try {
      if (typeof App._currentImageExport === 'function') {
        const queued = await App.readPresentationQueue();
        const staleTableIds = queued
          .filter(item => item && item.kind === 'table')
          .map(item => item.id)
          .filter(Boolean);
        if (staleTableIds.length) {
          await studioDbDeleteMany(STUDIO_QUEUE_STORE, staleTableIds);
        }
        const exported = await App._currentImageExport();
        if (!exported?.image) {
          toast('ไม่พบภาพกราฟสำหรับส่งเข้า Presentation', 'warning');
          return;
        }
        const item = {
          id: exported.id || `image-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          kind: 'image',
          filename: exported.filename || filename,
          page: exported.page || State.page,
          createdAt: exported.createdAt || new Date().toISOString(),
          image: exported.image,
        };
        await App.queuePresentationSlide(item);
        toast('ส่งกราฟไปทำ Presentation แล้ว', 'success');
        return;
      }
      if (typeof App._currentTableExport !== 'function') {
        toast('หน้านี้ยังไม่รองรับการส่งออกไป Presentation', 'warning');
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

  copyBtn?.addEventListener('click', async () => {
    try {
      if (typeof App._currentClipboardExport === 'function') {
        const mode = await App._currentClipboardExport();
        toast(mode === 'custom-html'
          ? 'คัดลอกกราฟและตารางพร้อมฟอร์แมตแล้ว'
          : 'คัดลอกข้อมูลพร้อมฟอร์แมตแล้ว', 'success');
        return;
      }
      if (typeof App._currentTableExport !== 'function') {
        toast('หน้านี้ยังไม่รองรับการคัดลอกแบบมีฟอร์แมต', 'warning');
        return;
      }
      const payload = App._currentTableExport();
      if (!payload?.rows?.length) {
        toast('ไม่พบข้อมูลตารางสำหรับคัดลอก', 'warning');
        return;
      }
      const mode = await copyPresentationPayloadToClipboard(payload, { card });
      toast(mode === 'rendered-html'
        ? 'คัดลอกจากตารางหน้าเว็บจริงพร้อมฟอร์แมตแล้ว'
        : mode === 'html'
          ? 'คัดลอกตารางพร้อมฟอร์แมตแล้ว'
          : 'คัดลอกข้อมูลตารางแบบข้อความสำรองแล้ว', 'success');
    } catch (err) {
      toast(`คัดลอกไม่สำเร็จ: ${err.message || err}`, 'error');
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
      title: 'ค่าธรรมเนียม',
      source: 'Raw For Sec + AVP Master Fund ID',
      localFile: 'Data/Raw For Sec - 2026-Q1.json',
    },
    'master-placeholder-5': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'ปัจจัยประกอบอื่นๆ',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
    },
    'master-placeholder-9': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'ปัจจัยประกอบอื่นๆ 2',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
    },
    'master-placeholder-10': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'ปัจจัยประกอบอื่นๆ 3',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
    },
    'master-placeholder-11': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'ปัจจัยประกอบอื่นๆ 4',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
    },
    'master-placeholder-12': {
      sheetId: CONFIG.SHEETS?.RAW_FOR_SEC || '',
      tabName: '2026-Q1',
      title: 'เปรียบเทียบค่าธรรมเนียม',
      source: 'Raw For Sec + AVP Master Fund ID',
      localFile: 'Data/Raw For Sec - 2026-Q1.json',
    },
    'master-placeholder-7': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'Top 10 Holding V2',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
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

  // ถ้า user เลือก Quarter จาก dropdown ให้ override tabName ทุก page
  const tabName = State.currentQuarter || cfg.tabName;
  const cfgWithTab = { ...cfg, tabName };

  const mode = CONFIG.DATA_SOURCE || 'google_first';

  if (mode === 'local_only') {
    const rows = await fetchLocalRows(cfgWithTab.localFile);
    State._pageDataSource[pageKey] = 'Source: Local JSON';
    return rememberPageDataMeta(pageKey, rows, cfgWithTab);
  }

  if (mode === 'local_first') {
    try {
      const rows = await fetchLocalRows(cfgWithTab.localFile);
      State._pageDataSource[pageKey] = 'Source: Local JSON';
      return rememberPageDataMeta(pageKey, rows, cfgWithTab);
    } catch (err) {
      const rows = await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
      State._pageDataSource[pageKey] = 'Source: Google Sheets fallback';
      return rememberPageDataMeta(pageKey, rows, cfgWithTab);
    }
  }

  if (mode === 'google_first') {
    try {
      const rows = await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
      State._pageDataSource[pageKey] = `Source: Google Sheets (${cfgWithTab.tabName})`;
      return rememberPageDataMeta(pageKey, rows, cfgWithTab);
    } catch (err) {
      if (!cfgWithTab.localFile) throw err;
      const rows = await fetchLocalRows(cfgWithTab.localFile);
      State._pageDataSource[pageKey] = 'Source: Local JSON fallback';
      return rememberPageDataMeta(pageKey, rows, cfgWithTab);
    }
  }

  const rows = await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
  State._pageDataSource[pageKey] = `Source: Google Sheets (${cfgWithTab.tabName})`;
  return rememberPageDataMeta(pageKey, rows, cfgWithTab);
}

function resolveQuarterLocalFile(localFile, baseTabName, requestedTabName) {
  if (!localFile || !baseTabName || !requestedTabName || baseTabName === requestedTabName) return localFile;
  return String(localFile).replace(baseTabName, requestedTabName);
}

function quarterSortValue(label) {
  const match = String(label || '').trim().match(/^(\d{4})-Q([1-4])$/i);
  if (!match) return Number.NEGATIVE_INFINITY;
  return (Number(match[1]) * 10) + Number(match[2]);
}

function getFeeComparisonQuarterPair() {
  const pageQuarter = CONFIG.PAGES?.['master-placeholder-12']?.tabName || '';
  const currentQuarter = String(State.currentQuarter || pageQuarter || '').trim();
  const detectedQuarters = (State.availableQuarters || [])
    .map(q => String(q || '').trim())
    .filter(q => /^\d{4}-Q[1-4]$/i.test(q));

  const quarters = [...new Set([
    ...detectedQuarters,
    currentQuarter,
  ].filter(Boolean))].sort((a, b) => quarterSortValue(b) - quarterSortValue(a));

  if (!quarters.length) {
    return { currentQuarter: '', previousQuarter: '' };
  }

  const resolvedCurrent = quarters.includes(currentQuarter) ? currentQuarter : quarters[0];
  const currentIndex = quarters.indexOf(resolvedCurrent);
  const previousQuarter = currentIndex >= 0 && currentIndex < quarters.length - 1
    ? quarters[currentIndex + 1]
    : '';

  return {
    currentQuarter: resolvedCurrent,
    previousQuarter,
  };
}

async function fetchPageDataForTab(pageKey, tabName) {
  ensureExtendedPageConfigs();
  const cfg = CONFIG.PAGES[pageKey];
  if (!cfg) throw new Error(`Unknown page: ${pageKey}`);

  const cfgWithTab = {
    ...cfg,
    tabName,
    localFile: resolveQuarterLocalFile(cfg.localFile, cfg.tabName, tabName),
  };

  const mode = CONFIG.DATA_SOURCE || 'google_first';

  if (mode === 'local_only') {
    return await fetchLocalRows(cfgWithTab.localFile);
  }

  if (mode === 'local_first') {
    try {
      return await fetchLocalRows(cfgWithTab.localFile);
    } catch (err) {
      return await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
    }
  }

  if (mode === 'google_first') {
    try {
      return await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
    } catch (err) {
      if (!cfgWithTab.localFile) throw err;
      return await fetchLocalRows(cfgWithTab.localFile);
    }
  }

  return await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
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

async function fetchCachedForTab(pageKey, tabName) {
  const key = `page::${pageKey}::${tabName}`;
  const now = Date.now();
  if (State._cache[key] && now - State._cache[key].ts < CONFIG.CACHE_TTL) {
    return State._cache[key].data;
  }
  const data = await fetchPageDataForTab(pageKey, tabName);
  State._cache[key] = { data, ts: now };
  return data;
}

function clearCache() {
  State._cache = {};
  State._pageDataSource = {};
  State._pageDataMeta = {};
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
    CODE:     findColumnIndex(headers, ['Fund Code', 'FundId', 'SecId', 'Sec ID', 'sec_id', 'Code']),
    NAME:     findColumnIndex(headers, ['Name', 'Fund Name', 'FundName']),
    MASTER:   findColumnIndex(headers, ['Master Fund', 'Master Fund Name', 'MasterFund']),
    ISIN:     findColumnIndex(headers, ['ISIN', 'Master Fund ID', 'MasterFundId', 'Master Fund Id']),
    TYPE:     findColumnIndex(headers, ['Fund Type', 'Type', 'FundType']),
    DIVIDEND: findColumnIndex(headers, ['Dividend', 'Div']),
    STYLE:    findColumnIndex(headers, ['Style']),
    ASSET_HOUSE: findColumnIndex(headers, ['Asset House', 'AssetHouse', 'AMC']),
  };

  return rows.slice(1).map(r => {
    const code = (CI.CODE >= 0 ? r[CI.CODE] : '') || '';
    const name = (CI.NAME >= 0 ? r[CI.NAME] : '') || '';
    return {
      category:   (CI.CATEGORY >= 0 ? r[CI.CATEGORY] : '')  || '',
      code,
      key:        normalizeFundKey(code),
      name,
      masterId:   (CI.ISIN >= 0   ? r[CI.ISIN]   : '') || '',
      masterName: (CI.MASTER >= 0 ? r[CI.MASTER] : '') || name,
      type:       (CI.TYPE >= 0   ? r[CI.TYPE]   : '') || deriveFundType(code),
      dividend:   (CI.DIVIDEND >= 0 ? r[CI.DIVIDEND] : '') || deriveDividend(code),
      style:      (CI.STYLE >= 0  ? r[CI.STYLE]  : '') || deriveStyle(code, (CI.MASTER >= 0 ? r[CI.MASTER] : '') || ''),
      assetHouse: (CI.ASSET_HOUSE >= 0 ? r[CI.ASSET_HOUSE] : '') || '',
    };
  }).filter(f => f.code);
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

function buildMasterCalendarExportPayload(sorted, yearKeys, linksByRowKey, rankMap, rankTotals, CI, get, sortState) {
  const returnTheme = { bg: '#4ba3dc', color: '#ffffff' };
  const rankTheme = { bg: '#2f537f', color: '#ffffff' };
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
        ...yearKeys.map(year => {
          const rankValue = rankMap[item.key]?.[year] ?? '';
          const colors = extractInlineColors(rankCellStyle(rankValue, rankTotals[year]));
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

  const columns = [
    buildPresentationSortColumn(sortState, 'name', 'Master Fund', 2.4, { align: 'left', widthPx: preset.columnWidthsPx?.name }),
    buildPresentationSortColumn(sortState, 'currency', 'Base Currency', 1.15, { widthPx: preset.columnWidthsPx?.currency }),
    buildPresentationSortColumn(sortState, 'thai', 'กองทุนในไทย', 1.9, { align: 'left', widthPx: preset.columnWidthsPx?.thai }),
    ...yearKeys.map(year => buildPresentationSortColumn(sortState, `ret-${year}`, year, 0.82, {
      ...returnTheme,
      widthPx: preset.columnWidthsPx?.metric,
    })),
    ...yearKeys.map(year => buildPresentationSortColumn(sortState, `rank-${year}`, year, 0.82, {
      ...rankTheme,
      widthPx: preset.columnWidthsPx?.metric,
    })),
  ];

  return buildPresentationTablePayload({
    presetKey: 'masterCalendar',
    title: CONFIG.PAGES['master-calendar']?.title || 'Master Fund Calendar Year',
    source: CONFIG.PAGES['master-calendar']?.source || 'AVP Master Fund ID',
    headerGroups: [
      { label: '', span: 3, bg: '#dbe4f0', color: '#334155' },
      { label: 'Calendar Year Return (%)', span: yearKeys.length, ...returnTheme },
      { label: 'อันดับในกลุ่มที่แสดง', span: yearKeys.length, ...rankTheme },
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

function renderFeeCompareMiniTable(rows, options = {}) {
  const title = options.title || 'ตารางเปรียบเทียบค่าธรรมเนียม';
  const maxCombined = Math.max(...rows.map(item => item.combined || 0), 0);
  return `
    <div class="card fee-compare-card">
      <div class="fee-compare-card-head">
        <h3>${esc(title)}</h3>
      </div>
      <div class="fee-compare-table-wrap">
        <table class="fee-compare-table">
          <colgroup>
            <col class="fee-compare-col-master">
            <col class="fee-compare-col-thai">
            <col class="fee-compare-col-ter">
            <col class="fee-compare-col-ter">
            <col class="fee-compare-col-ter">
          </colgroup>
          <thead>
            <tr>
              <th rowspan="2">Master Fund</th>
              <th rowspan="2">กองไทย</th>
              <th colspan="3">TER (%)</th>
            </tr>
            <tr>
              <th>Master Fund</th>
              <th>กองไทย</th>
              <th>COMBINED TER</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(row => {
              const combinedStyle = feeCombinedStyle(row.combined, maxCombined);
              return `
                <tr>
                  <td class="fee-compare-master">${row.masterNameHtml || esc(row.masterName)}</td>
                  <td class="fee-compare-thai"${row.highlightColor ? ` style="background:${row.highlightColor}"` : ''}>${esc(row.thaiCode)}</td>
                  <td class="fee-compare-num">${esc(row.masterTerText || '-')}</td>
                  <td class="fee-compare-num">${esc(row.thaiTerText || '-')}</td>
                  <td class="fee-compare-num fee-compare-combined"${combinedStyle.bg ? ` style="background:${combinedStyle.bg};color:${combinedStyle.color}"` : ''}>${esc(row.combinedText || '-')}</td>
                </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
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
      { label: 'TER (%) Q1-2026', span: 3, bg: '#1f3f74', color: '#ffffff' },
      { label: '', span: 1, bg: '#eef2f7', color: '#eef2f7' },
      { label: 'การซื้อ-ขาย (%)', span: 2, bg: '#2d5a33', color: '#ffffff' },
      { label: '', span: 4, bg: '#1f3f74', color: '#ffffff' },
    ],
    columns: [
      { key: 'master', label: 'Master Fund', weight: 1.55, align: 'left', bg: '#1f3f74', color: '#ffffff' },
      { key: 'thai', label: 'กองไทย', weight: 1.2, widthPx: preset.columnWidthsPx?.thai, align: 'center', bg: '#1f3f74', color: '#ffffff' },
      { key: 'masterTer', label: 'Master Fund', weight: 0.9, widthPx: preset.columnWidthsPx?.masterTer, bg: '#1f3f74', color: '#ffffff' },
      { key: 'thaiTer', label: 'กองไทย (Sec)', weight: 0.9, widthPx: preset.columnWidthsPx?.thaiTer, bg: '#1f3f74', color: '#ffffff' },
      { key: 'combined', label: 'Combined TER', weight: 0.9, widthPx: preset.columnWidthsPx?.combined, bg: '#1f3f74', color: '#ffffff' },
      { key: 'spacer', label: '', weight: 0.28, widthPx: preset.columnWidthsPx?.spacer, bg: '#eef2f7', color: '#eef2f7' },
      { key: 'frontLoad', label: 'IN (ซื้อ)', weight: 0.9, widthPx: preset.columnWidthsPx?.frontLoad, bg: '#2d5a33', color: '#ffffff' },
      { key: 'backLoad', label: 'OUT (ขาย)', weight: 0.9, widthPx: preset.columnWidthsPx?.backLoad, bg: '#2d5a33', color: '#ffffff' },
      { key: 'initial', label: 'การซื้อครั้งแรกขั้นตํ่า', weight: 0.9, widthPx: preset.columnWidthsPx?.initial, bg: '#1f3f74', color: '#ffffff' },
      { key: 'subsequent', label: 'การซื้อครั้งถัดไปขั้นตํ่า', weight: 0.9, widthPx: preset.columnWidthsPx?.subsequent, bg: '#1f3f74', color: '#ffffff' },
      { key: 'fxHedging', label: 'FX Hedging', weight: 0.9, widthPx: preset.columnWidthsPx?.fxHedging, bg: '#1f3f74', color: '#ffffff' },
      { key: 'depositCurrency', label: 'Base Currency', weight: 0.9, widthPx: preset.columnWidthsPx?.depositCurrency, bg: '#1f3f74', color: '#ffffff' },
      { key: 'sourceLink', label: 'FACTSHEET', weight: 0.9, widthPx: preset.columnWidthsPx?.source, bg: '#1f3f74', color: '#ffffff' },
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
          { text: row.combinedText || '-', bg: combinedStyle.bg || baseRowBg, color: combinedStyle.color || '#183b22', weight: 1.08, strong: true },
          { text: '', bg: '#eef2f7', color: '#eef2f7', weight: 0.28 },
          { text: row.frontText || '-', bg: baseRowBg, color: '#334155', weight: 0.86 },
          { text: row.backText || '-', bg: baseRowBg, color: '#334155', weight: 0.86 },
          { text: row.initialText || '-', bg: baseRowBg, color: '#334155', weight: 0.94 },
          { text: row.subsequentText || '-', bg: baseRowBg, color: '#334155', weight: 0.94 },
          { text: row.fxHedgingText || '-', bg: baseRowBg, color: '#334155', weight: 0.88 },
          { text: row.depositCurrencyText || '-', bg: baseRowBg, color: '#334155', weight: 0.96 },
          { text: row.sourceLink ? 'LINK' : '-', href: row.sourceLink ? buildFactsheetViewerUrl(row.sourceLink) : '', bg: baseRowBg, color: row.sourceLink ? '#3559d7' : '#64748b', weight: 0.7, strong: !!row.sourceLink },
        ],
      };
    }),
  });
}

function buildFactsheetViewerUrl(url) {
  const normalized = String(url || '').trim();
  if (!normalized) return '';
  return `https://docs.google.com/viewer?url=${encodeURIComponent(normalized)}&embedded=true`;
}

function top10Text(value) {
  const text = String(value ?? '').trim();
  return text || '—';
}

function top10ArrayItemValue(item) {
  if (!item || typeof item !== 'object') return top10Text(item);
  const name = top10Text(item.name || item.companyName || item.field || item.label || item.key || '');
  const value = top10Text(item.text || item.weightText || item.value || item.percent || item.weightPercent || item.weight || '');
  if (name === '—' && value === '—') return '—';
  if (name === '—') return value;
  if (value === '—') return name;
  return `${name} (${value})`;
}

function buildTop10SingleExportPayload(title, source, isin, data = {}) {
  const headers = ['Section', 'Field', 'Value'];
  const rows = [];
  const pushObject = (section, obj) => {
    Object.entries(obj || {}).forEach(([key, value]) => {
      rows.push([section, key, top10Text(value)]);
    });
  };
  const pushArray = (section, arr) => {
    (Array.isArray(arr) ? arr : []).forEach((item, index) => {
      rows.push([section, String(index + 1), top10ArrayItemValue(item)]);
    });
  };

  if (isin) rows.push(['Meta', 'ISIN', top10Text(isin)]);
  if (data.summary) pushObject('Summary', data.summary);
  if (data.sizes) pushObject('Sizes', data.sizes);
  if (data.fees) pushObject('Fees', data.fees);
  if (data.manager) pushObject('Manager', data.manager);
  if (data.performance) pushObject('Performance', data.performance);
  if (data.risk) pushObject('Risk', data.risk);
  if (data.ratings) pushObject('Ratings', data.ratings);
  if (data.objective) rows.push(['Objective', 'Text', top10Text(data.objective)]);
  pushArray('Holdings', data.holdings);
  pushArray('Allocation: Asset Type', data.allocation?.assetType);
  pushArray('Allocation: Sector', data.allocation?.sector);
  pushArray('Allocation: Region', data.allocation?.region);
  pushArray('Historical', data.historical);

  return buildSimpleTablePayload(title, source, headers, rows);
}

function buildTop10CompareExportPayload(title, source, entries = []) {
  const funds = Array.isArray(entries) ? entries : [];
  const headers = ['Section', 'Metric', ...funds.map(entry => top10Text(entry.label || entry.isin))];
  const rows = [];
  const addMetricRows = (section, metrics, valueGetter) => {
    metrics.forEach(metric => {
      rows.push([
        section,
        metric.label,
        ...funds.map(entry => top10Text(valueGetter(entry, metric))),
      ]);
    });
  };
  const metrics = {
    fees: [
      { key: 'ongoingCharge', label: 'ongoingCharge' },
      { key: 'initialCharge', label: 'initialCharge' },
      { key: 'maxAnnualCharge', label: 'maxAnnualCharge' },
      { key: 'exitCharge', label: 'exitCharge' },
    ],
    manager: [
      { key: 'name', label: 'name' },
      { key: 'startDate', label: 'startDate' },
    ],
    performance: [
      { key: '1M', label: '1M' },
      { key: '3M', label: '3M' },
      { key: '6M', label: '6M' },
      { key: '1Y', label: '1Y' },
      { key: '3Y', label: '3Y' },
      { key: '5Y', label: '5Y' },
    ],
    risk: ['riskLevel', 'volatility', 'category'],
    holding: Array.from({ length: 10 }, (_, i) => i + 1),
  };

  addMetricRows('Fees', metrics.fees, (entry, metric) => entry.data?.fees?.[metric.key]);
  addMetricRows('Manager', metrics.manager, (entry, metric) => entry.data?.manager?.[metric.key]);
  addMetricRows('Performance', metrics.performance, (entry, metric) => entry.data?.performance?.[metric.key]);
  addMetricRows('Risk', metrics.risk.map(key => ({ key, label: key })), (entry, metric) => entry.data?.risk?.[metric.key]);

  metrics.holding.forEach(rank => {
    rows.push([
      'Top 10 Holding',
      `#${rank}`,
      ...funds.map(entry => {
        const item = Array.isArray(entry.data?.holdings) ? entry.data.holdings[rank - 1] : null;
        return top10ArrayItemValue(item);
      }),
    ]);
  });

  return buildSimpleTablePayload(title, source, headers, rows);
}

function buildTop10V3ExportPayload(title, source, funds = []) {
  const rows = Array.isArray(funds) ? funds : [];
  const headers = [
    'Fund',
    'ISIN',
    'Ongoing Charge',
    'Initial Charge',
    'Max Annual Charge',
    'Exit Charge',
    ...Array.from({ length: 10 }, (_, i) => `Holding #${i + 1}`),
  ];
  const dataRows = rows.map(fund => {
    const holdings = Array.isArray(fund?._topHoldings) ? fund._topHoldings : [];
    return [
      top10Text((fund?._selectorName || fund?.name || '').split('–')[0].trim()),
      top10Text(fund?._isin),
      top10Text(fund?._feesRaw),
      top10Text(fund?._feesInitial),
      top10Text(fund?._feesMaxAnnual),
      top10Text(fund?._feesExit),
      ...Array.from({ length: 10 }, (_, i) => top10ArrayItemValue(holdings[i])),
    ];
  });

  return buildSimpleTablePayload(title, source, headers, dataRows);
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

function formatMinimumPurchaseDisplay(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  const n = parseNum(text);
  if (Number.isNaN(n)) return text;
  const decimals = text.includes('.') ? text.split('.')[1].length : 0;
  return n.toLocaleString('en-US', {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals,
  });
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
    fundSize: findColumnIndex(headers, ['Fund Size', 'AUM', 'Net Assets', 'Total Net Assets']),
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
    fundSize: get(row, ci.fundSize),
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
    initial: findColumnIndex(headers, ['Initial']),
    subsequent: findColumnIndex(headers, ['Subsequent']),
    fxHedging: findColumnIndex(headers, ['fx_hedging']),
    pdfFactsheet: findColumnIndex(headers, ['pdf_factsheet']),
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
      initial: get(row, ci.initial),
      subsequent: get(row, ci.subsequent),
      fxHedging: get(row, ci.fxHedging),
      pdfFactsheet: get(row, ci.pdfFactsheet),
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
  const masterRows = buildMasterRecords(await fetchCached('master-placeholder-2'));
  return buildSelectedMasterUniverseFromRows(selectRows, masterRows);
}

function buildSelectedMasterUniverseFromRows(selectRows, masterRows) {
  const catalog = buildSelectedFundsCatalog(selectRows);
  const allFunds = catalog
    .filter(f => f.code)
    .sort((a, b) => String(a.code).localeCompare(String(b.code), 'th'));
  const selectedFunds = State.selectedKeys.size > 0
    ? allFunds.filter(f => State.selectedKeys.has(f.key))
    : [];
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

async function buildSelectedMasterUniverseForQuarter(tabName, selectRows = null) {
  const resolvedSelectRows = selectRows || await fetchCached('select-fund');
  const masterRows = buildMasterRecords(await fetchCachedForTab('master-placeholder-2', tabName));
  return buildSelectedMasterUniverseFromRows(resolvedSelectRows, masterRows);
}

function buildFeeComparisonRows(universe, rawLookup) {
  const rows = universe.map(({ fund, master }) => {
    const raw = rawLookup.get(String(fund.code || '').trim().toUpperCase()) || null;
    const masterTer = parseNum(master.ongoingCost);
    const thaiTer = parseNum(raw?.ter);
    const combined = !Number.isNaN(masterTer) && !Number.isNaN(thaiTer) ? masterTer + thaiTer : NaN;
    const front = toFixedSafe(raw?.front, 4);
    const back = toFixedSafe(raw?.back, 4);
    const fxHedging = toFixedSafe(raw?.fxHedging, 2);
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
      frontText: front || '',
      backText: back || '',
      initialText: formatMinimumPurchaseDisplay(raw?.initial),
      subsequentText: formatMinimumPurchaseDisplay(raw?.subsequent),
      fxHedgingText: fxHedging || '',
      depositCurrencyText: master?.currency || '',
      sourceLink: raw?.pdfFactsheet || '',
    };
  }).filter(item => !Number.isNaN(item.masterTer) && !Number.isNaN(item.thaiTer));
  return annotateFeeComparisonNameDiffs(rows);
}

function tokenizeComparableText(text) {
  return String(text || '').trim().split(/\s+/).filter(Boolean);
}

function normalizeComparableToken(token) {
  return String(token || '').toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '');
}

function compareTextTokens(left, right) {
  const leftTokens = tokenizeComparableText(left);
  const rightTokens = tokenizeComparableText(right);
  const len = Math.max(leftTokens.length, rightTokens.length);
  const differences = [];
  let matchingCount = 0;

  for (let i = 0; i < len; i += 1) {
    const leftToken = leftTokens[i] || '';
    const rightToken = rightTokens[i] || '';
    if (normalizeComparableToken(leftToken) === normalizeComparableToken(rightToken)) {
      if (leftToken || rightToken) matchingCount += 1;
      continue;
    }
    differences.push({ index: i, left: leftToken, right: rightToken });
  }

  return { leftTokens, rightTokens, differences, matchingCount };
}

function buildDiffHighlightedText(text, diffIndexes) {
  const tokens = tokenizeComparableText(text);
  return tokens.map((token, index) => {
    const safe = esc(token);
    return diffIndexes.has(index)
      ? `<span class="fee-v2-name-diff">${safe}</span>`
      : safe;
  }).join(' ');
}

function annotateComparableNameDiffs(rows, {
  nameKey = 'masterName',
  htmlKey = 'masterNameHtml',
  diffKey = 'masterNameDiffs',
} = {}) {
  return rows.map((row, index) => {
    const rowName = row?.[nameKey];
    let bestMatch = null;

    rows.forEach((candidate, candidateIndex) => {
      if (candidateIndex === index) return;
      const candidateName = candidate?.[nameKey];
      const result = compareTextTokens(rowName, candidateName);
      const tokenCount = Math.max(result.leftTokens.length, result.rightTokens.length);
      if (tokenCount < 2) return;
      if (!result.differences.length || result.differences.length > 2) return;
      if (result.matchingCount < tokenCount - result.differences.length) return;

      if (!bestMatch
        || result.matchingCount > bestMatch.matchingCount
        || (result.matchingCount === bestMatch.matchingCount && result.differences.length < bestMatch.differences.length)) {
        bestMatch = result;
      }
    });

    if (!bestMatch) return { ...row, [htmlKey]: esc(rowName) };

    const diffIndexes = new Set(bestMatch.differences.map(item => item.index));
    return {
      ...row,
      [diffKey]: bestMatch.differences,
      [htmlKey]: buildDiffHighlightedText(rowName, diffIndexes),
    };
  });
}

function annotateFeeComparisonNameDiffs(rows) {
  return annotateComparableNameDiffs(rows, {
    nameKey: 'masterName',
    htmlKey: 'masterNameHtml',
    diffKey: 'masterNameDiffs',
  });
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
    const labels  = ['กองทุนที่เลือกได้','กองทุนไทย Annualized Return','กองทุนไทย Calendar','Master Fund Annualized Return'];
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
      { page: 'thai-annualized-v2', icon: quickIcon('trend'), title: 'กองทุนไทย Annualized Return', sub: 'สลับดู Return และ Rank ได้' },
      { page: 'thai-calendar',     icon: quickIcon('cal'),    title: 'กองทุนไทย Calendar Year',  sub: 'AVP Thai Fund for Quality' },
      { page: 'master-annualized-v2', icon: quickIcon('globe'), title: 'Master Fund Annualized Return', sub: 'ISIN match + Base Currency' },
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

  async feeComparisonPlaceholder(area) {
    const pageKey = 'master-placeholder-12';
    setLoading(area, 'กำลังโหลดตารางเปรียบเทียบค่าธรรมเนียม...');
    try {
      const { currentQuarter, previousQuarter } = getFeeComparisonQuarterPair();
      if (!currentQuarter || !previousQuarter) {
        setError(area, 'เมนูนี้ต้องมีข้อมูลอย่างน้อย 2 Quarter ใน Google Sheets เช่น 2026-Q3 และ 2026-Q1', pageKey);
        return;
      }

      const selectRows = await fetchCached('select-fund');
      const [
        rawSecRowsCurrent,
        rawSecRowsPrevious,
        universeCurrent,
        universePrevious,
      ] = await Promise.all([
        fetchCachedForTab(pageKey, currentQuarter),
        fetchCachedForTab(pageKey, previousQuarter),
        buildSelectedMasterUniverseForQuarter(currentQuarter, selectRows),
        buildSelectedMasterUniverseForQuarter(previousQuarter, selectRows),
      ]);

      const decorateRows = (rows) => rows
        .sort((a, b) => compareValues(a.combined, b.combined, 'asc'))
        .map(row => {
          const matchedFund = Object.values(State.selectedFunds).find(f => f.code === row.thaiCode);
          const colorIdx = matchedFund ? State.highlights[matchedFund.key] : undefined;
          return {
            ...row,
            highlightColor: colorIdx !== undefined ? HL_COLORS[colorIdx]?.bg || '' : '',
          };
        });

      const feeRowsCurrent = decorateRows(buildFeeComparisonRows(universeCurrent, buildRawSecLookup(rawSecRowsCurrent)));
      const feeRowsPrevious = decorateRows(buildFeeComparisonRows(universePrevious, buildRawSecLookup(rawSecRowsPrevious)));

      if (!feeRowsCurrent.length && !feeRowsPrevious.length) {
        setError(area, 'ไม่พบข้อมูลค่าธรรมเนียมที่จับคู่ได้จาก Raw For Sec และ AVP Master Fund ID', pageKey);
        return;
      }

      const source = `${CONFIG.PAGES[pageKey]?.source || 'Raw For Sec + AVP Master Fund ID'} · ${currentQuarter} vs ${previousQuarter}`;
      area.innerHTML = `
        ${pageToolActions(pageKey, source)}
        <div id="report-card" class="fee-compare-page">
          <div class="fee-compare-grid">
            ${renderFeeCompareMiniTable(feeRowsCurrent, { title: `ตารางค่าธรรมเนียม · ${currentQuarter}` })}
            ${renderFeeCompareMiniTable(feeRowsPrevious, { title: `ตารางค่าธรรมเนียม · ${previousQuarter}` })}
          </div>
        </div>`;
    } catch (e) {
      setError(area, e.message, pageKey);
      return;
    }
    App._currentExport = null;
    App._currentTableExport = null;
    App._currentClipboardExport = null;
    App._currentImageExport = null;
    bindPageImageActions(area, 'report-card', 'master-fee-compare');
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
      App._currentExport = () => exportExcel(rawRows, cfg.title);
      App._currentTableExport = () => buildSimpleTablePayload(cfg.title, cfg.source || '', rawRows[0] || [], rawRows.slice(1));
      bindPageImageActions(area, 'report-card', pageKey);
    };

    render();
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
      State.pageSize = State.selectFundFilters.pageSize || SELECT_FUND_DEFAULT_PAGE_SIZE;

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
      const tRows = pageData.map(f => {
        const isSelected = State.selectedKeys.has(f.key);
        return `
          <tr data-key="${esc(f.key)}" class="${isSelected ? 'row-selected' : ''}">
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
            <span class="row-count-badge is-info" id="sf-selected-count">เลือกแล้ว ${State.selectedKeys.size.toLocaleString()} กองทุน</span>
            ${sourceBadgeHtml('select-fund', cfg.source)}
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
                ${SELECT_FUND_PAGE_SIZE_OPTIONS.map(size => `<option value="${size}" ${size === State.pageSize ? 'selected' : ''}>${size}</option>`).join('')}
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
          pageSize: SELECT_FUND_DEFAULT_PAGE_SIZE,
        };
        State.selectFundSort = { key: '', dir: 'asc' };
        State.tablePage = 1;
        State.pageSize = SELECT_FUND_DEFAULT_PAGE_SIZE;
        State.selectedKeys.clear();
        State.selectedFunds = {};
        State.highlights = {};
        render(1);
      });

      /* Selection */
      const syncSelectionUi = () => {
        const selectedCountEl = $('#sf-selected-count', area);
        if (selectedCountEl) {
          selectedCountEl.textContent = `เลือกแล้ว ${State.selectedKeys.size.toLocaleString()} กองทุน`;
        }
        const rowCheckboxes = $$('.row-chk', area);
        const checkedCount = rowCheckboxes.filter(cb => cb.checked).length;
        const selectAllCheckbox = $('#sf-chk-all', area);
        if (selectAllCheckbox) {
          selectAllCheckbox.checked = rowCheckboxes.length > 0 && checkedCount === rowCheckboxes.length;
          selectAllCheckbox.indeterminate = checkedCount > 0 && checkedCount < rowCheckboxes.length;
        }
      };

      const chkAll = $('#sf-chk-all', area);
      if (chkAll) {
        chkAll.addEventListener('change', () => {
          $$('.row-chk', area).forEach(cb => {
            const key = cb.dataset.key;
            cb.checked = chkAll.checked;
            if (chkAll.checked) State.selectedKeys.add(key);
            else State.selectedKeys.delete(key);
            cb.closest('tr')?.classList.toggle('row-selected', chkAll.checked);
          });
          syncSelectionUi();
        });
      }

      $$('.row-chk', area).forEach(el => {
        el.addEventListener('change', () => {
          const key = el.dataset.key;
          if (el.checked) State.selectedKeys.add(key);
          else State.selectedKeys.delete(key);
          el.closest('tr')?.classList.toggle('row-selected', el.checked);
          syncSelectionUi();
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
        State.pageSize = parseInt(e.target.value, 10) || SELECT_FUND_DEFAULT_PAGE_SIZE;
        State.selectFundFilters.pageSize = State.pageSize;
        render(1);
      });

      syncSelectionUi();

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
        <div class="year-chip-wrap">
          <span class="metric-toggle-label">ปีที่ต้องการแสดง</span>
          ${yearChips}
        </div>
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

    const displayRowsRaw = Object.keys(masterLinks).length > 0
      ? masterRows.filter(item => masterLinks[item.key]?.length)
      : masterRows;
    const displayRows = annotateComparableNameDiffs(displayRowsRaw, {
      nameKey: 'name',
      htmlKey: 'nameHtml',
      diffKey: 'nameDiffs',
    });

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
          <td class="master-name-cell">${item.nameHtml || esc(item.name)}</td>
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
    App._currentExport = null;
    App._currentTableExport = () => {
      return buildMasterAnnualizedExportPayload(sorted, metricKeys, masterLinks, rankMap, rankTotals, CI, get, sortState);
    };
    bindPageImageActions(area, 'report-card', 'master-annualized');
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

    const displayRows = annotateComparableNameDiffs(
      masterRows.filter(item => linksByRowKey[item.key]?.length),
      {
        nameKey: 'name',
        htmlKey: 'nameHtml',
        diffKey: 'nameDiffs',
      }
    );
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
          <td class="master-name-cell">${item.nameHtml || esc(item.name)}</td>
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
    App._currentExport = null;
    App._currentTableExport = () => {
      return buildMasterAnnualizedV2ExportPayload(sorted, metricKeys, linksByRowKey, rankMap, rankTotals, CI, get, sortState);
    };
    bindPageImageActions(area, 'report-card', 'master-annualized-v2');
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
    const allYears = ['2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'];
    const CI = {
      name: findColumnIndex(headers, ['Group/Investment']),
      isin: findColumnIndex(headers, ['ISIN']),
      currency: findColumnIndex(headers, ['Base Currency']),
      ...Object.fromEntries(allYears.map(year => [`ret${year}`, findColumnIndex(headers, [`Return(Cumulative) ${year}`])])),
    };
    const get = (row, i) => i >= 0 ? String(row[i] ?? '').trim() : '';
    const selectedYears = (State.reportOptions['master-calendar-years'] || []).filter(y => allYears.includes(y));
    const visibleYears = selectedYears.length ? selectedYears : allYears;

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

    const displayRows = annotateComparableNameDiffs(
      masterRows.filter(item => linksByRowKey[item.key]?.length),
      {
        nameKey: 'name',
        htmlKey: 'nameHtml',
        diffKey: 'nameDiffs',
      }
    );
    const { ranks: rankMap, totals: rankTotals } = buildMetricRanks(displayRows, visibleYears, (item, year) => get(item.row, CI[`ret${year}`]));
    const sortState = State.reportSorts['master-calendar'];
    const sortableMasterCalendar = (item, key) => {
      const mapping = {
        name: item.name,
        currency: item.currency,
        thai: (linksByRowKey[item.key] || []).map(f => f.code).join(', '),
        ...Object.fromEntries(visibleYears.map(year => [`ret-${year}`, get(item.row, CI[`ret${year}`])])),
        ...Object.fromEntries(visibleYears.map(year => [`rank-${year}`, rankMap[item.key]?.[year] ?? ''])),
      };
      return mapping[key];
    };
    const sorted = sortState.key
      ? [...displayRows].sort((a, b) => compareValues(sortableMasterCalendar(a, sortState.key), sortableMasterCalendar(b, sortState.key), sortState.dir))
      : displayRows;

    const yearChips = allYears.map(year => {
      const active = visibleYears.includes(year);
      return `<button class="btn btn-ghost year-chip ${active ? 'is-active' : ''}" data-master-calendar-year="${year}" title="แสดงหรือซ่อนปี ${year}">${year}</button>`;
    }).join('');
    const toggleActions = `
      <div class="metric-toggle-stack master-calendar-toggle-stack master-calendar-toolbar-secondary">
        <div class="metric-toggle-group master-calendar-year-row">
          <span class="metric-toggle-label">ปีที่ต้องการแสดง</span>
          <div class="year-chip-wrap master-calendar-year-chip-wrap">
            ${yearChips}
          </div>
        </div>
      </div>`;

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
          <td class="master-name-cell">${item.nameHtml || esc(item.name)}</td>
          <td>${esc(item.currency || '-')}</td>
          <td class="thai-link-cell">${thaiHtml}</td>
          ${visibleYears.map(year => `<td class="report-num">${esc(formatReturnDisplay(get(item.row, CI[`ret${year}`])) || '-')}</td>`).join('')}
          ${visibleYears.map(year => {
            const value = rankMap[item.key]?.[year] ?? '';
            return `<td class="report-num report-rank-cell" style="${rankCellStyle(value, rankTotals[year])}">${esc(value || '-')}</td>`;
          }).join('')}
        </tr>`;
    }).join('');

    area.innerHTML = `
      ${pageToolActions('master-calendar', CONFIG.PAGES['master-calendar']?.source || 'AVP Master Fund ID', toggleActions)}
      <div class="card report-card report-card-master" id="report-card">
        <table class="annualized-report master-annualized-report">
          <thead>
            <tr class="report-group-row">
              <th colspan="3" class="group-blank"></th>
              <th colspan="${visibleYears.length}" class="group-blue">Calendar Year Return (%)</th>
              <th colspan="${visibleYears.length}" class="group-navy">อันดับในกลุ่มที่แสดง</th>
            </tr>
            <tr>
              <th class="report-sort ${sortState.key === 'name' ? 'is-active' : ''}" data-report-sort="name">${renderSortLabel('Master Fund', sortState.key === 'name', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'currency' ? 'is-active' : ''}" data-report-sort="currency">${renderSortLabel('Base Currency', sortState.key === 'currency', sortState.dir)}</th>
              <th class="report-sort ${sortState.key === 'thai' ? 'is-active' : ''}" data-report-sort="thai">${renderSortLabel('กองทุนในไทย', sortState.key === 'thai', sortState.dir)}</th>
              ${visibleYears.map(year => `<th class="report-sort ${sortState.key === `ret-${year}` ? 'is-active' : ''}" data-report-sort="ret-${year}">${renderSortLabel(year, sortState.key === `ret-${year}`, sortState.dir)}</th>`).join('')}
              ${visibleYears.map(year => `<th class="report-sort ${sortState.key === `rank-${year}` ? 'is-active' : ''}" data-report-sort="rank-${year}">${renderSortLabel(year, sortState.key === `rank-${year}`, sortState.dir)}</th>`).join('')}
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>`;

    $$('[data-master-calendar-year]', area).forEach(el => {
      el.addEventListener('click', () => {
        const year = el.dataset.masterCalendarYear;
        const current = new Set(State.reportOptions['master-calendar-years'] || []);
        if (current.has(year)) current.delete(year);
        else current.add(year);
        const next = allYears.filter(y => current.has(y));
        State.reportOptions['master-calendar-years'] = next.length ? next : allYears;
        Pages.masterCalendar(area);
      });
    });
    $$('.report-sort', area).forEach(el => {
      el.addEventListener('click', () => {
        toggleNamedSort(sortState, el.dataset.reportSort);
        Pages.masterCalendar(area);
      });
    });
    App._currentExport = null;
    App._currentTableExport = () => buildMasterCalendarExportPayload(sorted, visibleYears, linksByRowKey, rankMap, rankTotals, CI, get, sortState);
    bindPageImageActions(area, 'report-card', 'master-calendar');
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
      bindPageImageActions(area, 'report-card', 'master-fees');
    } catch (e) {
      setError(area, e.message, 'master-placeholder-1');
    }
  },

  async masterFeesV2(area) {
    setLoading(area, 'กำลังโหลดรายงานค่าธรรมเนียม...');

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
              <colgroup>
                <col class="fee-v2-col-master">
                <col class="fee-v2-col-thai">
                <col class="fee-v2-col-ter-1">
                <col class="fee-v2-col-ter-2">
                <col class="fee-v2-col-ter-3">
                <col class="fee-v2-col-spacer">
                <col class="fee-v2-col-fixed">
                <col class="fee-v2-col-fixed">
                <col class="fee-v2-col-fixed">
                <col class="fee-v2-col-fixed">
                <col class="fee-v2-col-fixed">
                <col class="fee-v2-col-fixed">
                <col class="fee-v2-col-fixed">
              </colgroup>
              <thead>
                <tr>
                  <th rowspan="2">Master Fund</th>
                  <th rowspan="2" class="fee-v2-main-thai-head">กองไทย</th>
                  <th colspan="3">TER (%)</th>
                  <th rowspan="2" class="fee-v2-th-spacer"></th>
                  <th colspan="2" class="fee-v2-th-group fee-v2-th-group--trade">การซื้อ-ขาย (%)</th>
                  <th rowspan="2" class="fee-v2-th-two-line"><span>การซื้อ</span><span>ครั้งแรกขั้นตํ่า</span></th>
                  <th rowspan="2" class="fee-v2-th-two-line"><span>การซื้อ</span><span>ครั้งถัดไปขั้นตํ่า</span></th>
                  <th rowspan="2">FX Hedging</th>
                  <th rowspan="2">Base Currency</th>
                  <th rowspan="2">FACTSHEET</th>
                </tr>
                <tr>
                  <th class="fee-v2-ter-head fee-v2-ter-head-1">Master Fund</th>
                  <th class="fee-v2-ter-head fee-v2-ter-head-2">กองไทย</th>
                  <th class="fee-v2-ter-head fee-v2-ter-head-3">COMBINED TER</th>
                  <th class="fee-v2-th-group--trade">IN (ซื้อ)</th>
                  <th class="fee-v2-th-group--trade">OUT (ขาย)</th>
                </tr>
              </thead>
              <tbody>
                ${feeRows.map(row => {
                  const combinedStyle = feeCombinedStyle(row.combined, maxCombined);
                  return `
                    <tr>
                      <td class="is-master">${row.masterNameHtml || esc(row.masterName)}</td>
                      <td class="is-thai"${row.highlightColor ? ` style="background:${row.highlightColor}"` : ''}>${esc(row.thaiCode)}</td>
                      <td class="fee-v2-ter-cell fee-v2-ter-cell-1">${esc(row.masterTerText || '-')}</td>
                      <td class="fee-v2-ter-cell fee-v2-ter-cell-2">${esc(row.thaiTerText || '-')}</td>
                      <td class="fee-v2-ter-cell fee-v2-ter-cell-3 is-combined"${combinedStyle.bg ? ` style="background:${combinedStyle.bg};color:${combinedStyle.color}"` : ''}>${esc(row.combinedText || '-')}</td>
                      <td class="fee-v2-td-spacer"></td>
                      <td>${esc(row.frontText || '-')}</td>
                      <td>${esc(row.backText || '-')}</td>
                      <td>${esc(row.initialText || '-')}</td>
                      <td>${esc(row.subsequentText || '-')}</td>
                      <td>${esc(row.fxHedgingText || '-')}</td>
                      <td>${esc(row.depositCurrencyText || '-')}</td>
                      <td>${row.sourceLink ? `<a class="fee-v2-source-link" href="${esc(buildFactsheetViewerUrl(row.sourceLink))}" target="_blank" rel="noopener noreferrer" aria-label="เปิด factsheet ของ ${esc(row.thaiCode)}">LINK</a>` : '<span class="fee-v2-muted">-</span>'}</td>
                    </tr>`;
                }).join('')}
              </tbody>
            </table>
          </div>
        </div>`;

      App._currentExport = null;
      App._currentTableExport = () => buildFeeV2ExportPayload(
        CONFIG.PAGES['master-placeholder-4']?.title || 'ค่าธรรมเนียม',
        source,
        feeRows
      );
      bindPageImageActions(area, 'report-card', 'master-fees-v2');
    } catch (e) {
      setError(area, e.message, 'master-placeholder-4');
    }
  },

  async masterOtherFactors(area) {
    setLoading(area, 'กำลังโหลดปัจจัยประกอบอื่นๆ...');

    let rawRows;
    try {
      await ensureSelectedFundsCatalog();
      rawRows = await fetchCached('master-placeholder-5');
    } catch (e) {
      setError(area, e.message, 'master-placeholder-5');
      return;
    }

    const headers = rawRows[0] || [];
    const get = (row, i) => i >= 0 ? String(row[i] ?? '').trim() : '';

    const masterRows = rawRows.slice(1).map((row, index) => {
      const name = get(row, findColumnIndex(headers, ['Group/Investment']));
      const isin = get(row, findColumnIndex(headers, ['ISIN']));
      const currency = get(row, findColumnIndex(headers, ['Base Currency']));
      return { row, name, isin, currency, key: `${isin}::${currency}::${name}::${index}` };
    }).filter(item => item.name && item.isin);

    const thaiSource = Object.values(State.selectedFunds).filter(f => f.masterId && f.masterId !== '-');
    const thaiFunds  = thaiSource.filter(f => State.selectedKeys.has(f.key));

    if (!thaiFunds.length) {
      setError(area, 'ยังไม่ได้เลือกกองทุนจากเมนูเลือกกองทุน หรือกองที่เลือกยังไม่ผูก Master Fund ID', 'master-placeholder-5');
      return;
    }

    const linksByRowKey = {};
    thaiFunds.forEach(fund => {
      const isin = String(fund.masterId || '').trim();
      if (!isin) return;
      masterRows.filter(item => item.isin === isin).forEach(item => {
        if (!linksByRowKey[item.key]) linksByRowKey[item.key] = [];
        linksByRowKey[item.key].push(fund);
      });
    });
    const displayRows = masterRows.filter(item => linksByRowKey[item.key]?.length);
    if (!displayRows.length) {
      setError(area, 'ไม่พบข้อมูล Master Fund ID ที่ตรงกับกองทุนที่เลือก', 'master-placeholder-5');
      return;
    }

    /* ── FIX 2: Persist state across navigation using State._of ── */
    if (!State._of) State._of = {};
    const S = State._of;
    if (!S.mode)        S.mode        = 'annualized';
    if (!S.period)      S.period      = '3Y';
    if (!S.xKey)        S.xKey        = 'maxdd';
    if (!S.yKey)        S.yKey        = 'return';
    if (!S.visibleKeys) S.visibleKeys = new Set(displayRows.map(r => r.key));
    if (!S.sideSort)    S.sideSort    = { key: null, dir: 0 };
    if (!S.labelOffsets) S.labelOffsets = {};
    if (!S.pointColors) S.pointColors = {};
    // sync any new rows that weren't in previous visit
    displayRows.forEach(r => { if (!S.visibleKeys.has) S.visibleKeys = new Set(displayRows.map(r2 => r2.key)); });

    /* ── Metric definitions ── */
    const ANNUALIZED_METRICS = [
      { key: 'return',        label: 'Return',                    getCol: (p, mode) => {
          if (mode === 'calendar') return `Return(Cumulative) ${p}`;
          const map = {
            '3M':'Return(Cumulative) 3M',
            '6M':'Return(Cumulative) 6M',
            YTD:'Return(Cumulative) YTD',
            '1Y':'Return(Cumulative) 1Y',
            '3Y':'Return(Annualized) 3Y',
            '5Y':'Return(Annualized) 5Y',
            '10Y':'Return(Annualized) 10Y'
          };
          return map[p] || '';
        }
      },
      { key: 'stddev',        label: 'Std Dev',                   prefix: 'Std Dev(Annualized)' },
      { key: 'sharpe',        label: 'Sharpe Ratio',              prefix: 'Sharpe Ratio(Annualized)' },
      { key: 'sharpe_arith',  label: 'Sharpe Ratio (arith)',      prefix: 'Sharpe Ratio (arith)(Annualized)' },
      { key: 'sharpe_geo',    label: 'Sharpe Ratio (geo)',        prefix: 'Sharpe Ratio (geo)(Annualized)' },
      { key: 'ir_arith',      label: 'Information Ratio (arith)', prefix: 'Information Ratio (arith)(Annualized)' },
      { key: 'ir_geo',        label: 'Information Ratio (geo)',   prefix: 'Information Ratio (geo)(Annualized)' },
      { key: 'sortino',       label: 'Sortino Ratio',             prefix: 'Sortino Ratio(Annualized)' },
      { key: 'sortino_arith', label: 'Sortino Ratio (arith)',     prefix: 'Sortino Ratio (arith)(Annualized)' },
      { key: 'sortino_geo',   label: 'Sortino Ratio (geo)',       prefix: 'Sortino Ratio (geo)(Annualized)' },
      { key: 'treynor_arith', label: 'Treynor Ratio (arith)',     prefix: 'Treynor Ratio (arith)(Annualized)' },
      { key: 'treynor_geo',   label: 'Treynor Ratio (geo)',       prefix: 'Treynor Ratio (geo)(Annualized)' },
      { key: 'maxdd',         label: 'Max Drawdown',              prefix: 'Max Drawdown' },
    ];
    const CALENDAR_METRICS   = ANNUALIZED_METRICS.filter(m => !m.key.startsWith('sharpe'));
    const ANNUALIZED_PERIODS = ['3M','6M','YTD','1Y','3Y','5Y','10Y'];
    const CALENDAR_YEARS     = ['2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'];
    const DOT_COLORS = [
      '#1a3c6e','#5aa2de','#e9b48c','#e3a72f','#8c3fe3',
      '#2f99bf','#7dc182','#dd6b20','#d946ef','#14b8a6',
      '#ef4444','#6366f1','#84cc16','#f59e0b','#06b6d4'
    ];

    function getColName(metric, period, mode) {
      if (metric.getCol) return metric.getCol(period, mode);
      return `${metric.prefix} ${period}`;
    }
    function getColIdx(metric, period, mode) {
      const col = getColName(metric, period, mode);
      return col ? findColumnIndex(headers, [col]) : -1;
    }
    function getBaseMetrics(mode) {
      return mode === 'calendar' ? CALENDAR_METRICS : ANNUALIZED_METRICS;
    }
    function getMetricsForPeriod(mode, period) {
      return getBaseMetrics(mode).filter(metric => getColIdx(metric, period, mode) >= 0);
    }
    function getAvailablePeriods(mode, xKey, yKey) {
      const periods = mode === 'calendar' ? CALENDAR_YEARS : ANNUALIZED_PERIODS;
      const metrics = getBaseMetrics(mode);
      const xMeta = metrics.find(metric => metric.key === xKey);
      const yMeta = metrics.find(metric => metric.key === yKey);
      return periods.filter(period => {
        const xOk = !xMeta || getColIdx(xMeta, period, mode) >= 0;
        const yOk = !yMeta || getColIdx(yMeta, period, mode) >= 0;
        return xOk && yOk;
      });
    }
    function normalizeOtherFactorsSelection() {
      const baseMetrics = getBaseMetrics(S.mode);
      if (!baseMetrics.find(metric => metric.key === S.xKey)) S.xKey = baseMetrics[0]?.key || 'return';
      if (!baseMetrics.find(metric => metric.key === S.yKey)) {
        S.yKey = baseMetrics.find(metric => metric.key !== S.xKey)?.key || S.xKey;
      }

      let availablePeriods = getAvailablePeriods(S.mode, S.xKey, S.yKey);
      if (!availablePeriods.length) {
        availablePeriods = S.mode === 'calendar' ? CALENDAR_YEARS.slice() : ANNUALIZED_PERIODS.slice();
      }
      if (!availablePeriods.includes(S.period)) S.period = availablePeriods[0];

      let periodMetrics = getMetricsForPeriod(S.mode, S.period);
      if (!periodMetrics.length) periodMetrics = baseMetrics;
      if (!periodMetrics.find(metric => metric.key === S.xKey)) S.xKey = periodMetrics[0]?.key || S.xKey;
      if (!periodMetrics.find(metric => metric.key === S.yKey)) {
        S.yKey = periodMetrics.find(metric => metric.key !== S.xKey)?.key || S.xKey;
      }

      availablePeriods = getAvailablePeriods(S.mode, S.xKey, S.yKey);
      if (availablePeriods.length && !availablePeriods.includes(S.period)) S.period = availablePeriods[0];
      periodMetrics = getMetricsForPeriod(S.mode, S.period);
      if (!periodMetrics.length) periodMetrics = baseMetrics;

      return { availablePeriods, periodMetrics };
    }
    function scatterLabelStateKey() {
      return `${S.mode}::${S.xKey}::${S.yKey}`;
    }
    function getScatterLabelOffsets() {
      const key = scatterLabelStateKey();
      if (!S.labelOffsets[key]) S.labelOffsets[key] = {};
      return S.labelOffsets[key];
    }
    function persistScatterLabelOffset(pointKey, dx, dy) {
      const offsets = getScatterLabelOffsets();
      offsets[pointKey] = { dx, dy };
    }
    function resetScatterLabelOffsets() {
      delete S.labelOffsets[scatterLabelStateKey()];
    }

    /* ── FIX 1: Smart label placement to avoid overlap ── */
    function placeLabels(points, W, H, pL, pR, pT, pB) {
      const plotW = W - pL - pR, plotH = H - pT - pB;
      const manualOffsets = getScatterLabelOffsets();
      const labeled = points.map((p, idx) => ({
        ...p,
        cx: pL + ((p.x - p._x0) / (p._xRange || 1)) * plotW,
        cy: pT + plotH - ((p.y - p._y0) / (p._yRange || 1)) * plotH,
        lx: 0, ly: 0,
        labelW: Math.max(18, Math.min(220, p.label.length * 7.4)),
        labelH: 24,
      }));

      // For each point, try 8 candidate positions and pick the one with least overlap
      const CANDIDATES = [
        [18, -14], [-18-60, -14], [18, 14], [-18-60, 14],
        [10, -24], [10, 28], [-10-60, -24], [-10-60, 28],
      ];
      labeled.forEach((p, i) => {
        let bestScore = Infinity, bestCx = 18, bestCy = -14;
        CANDIDATES.forEach(([dx, dy]) => {
          const lx = p.cx + dx, ly = p.cy + dy;
          const LW = p.labelW, LH = p.labelH;
          // Boundary penalty
          let score = 0;
          if (lx < pL) score += 100;
          if (lx + LW > W - pR) score += 100;
          if (ly - LH < pT) score += 100;
          if (ly > H - pB) score += 100;
          // Overlap penalty with other already-placed labels
          labeled.slice(0, i).forEach(q => {
            const qlx = q.lx, qly = q.ly;
            const ox = Math.max(0, Math.min(lx+LW, qlx+q.labelW) - Math.max(lx, qlx));
            const oy = Math.max(0, Math.min(ly, qly) - Math.max(ly-LH, qly-q.labelH));
            score += ox * oy * 2;
          });
          // Overlap with dots
          labeled.forEach((q, j) => {
            if (j === i) return;
            const dist = Math.hypot(lx - q.cx, ly - q.cy);
            if (dist < 20) score += (20 - dist) * 3;
          });
          if (score < bestScore) { bestScore = score; bestCx = dx; bestCy = dy; }
        });
        p.lx = p.cx + bestCx;
        p.ly = p.cy + bestCy;
        const manual = manualOffsets[p.key];
        if (manual && Number.isFinite(manual.dx) && Number.isFinite(manual.dy)) {
          p.lx = p.cx + manual.dx;
          p.ly = p.cy + manual.dy;
        }
      });
      return labeled;
    }

    function captureScatterPointState(container) {
      const points = new Map();
      if (!container) return points;
      container.querySelectorAll('.of-point[data-point-key]').forEach(node => {
        const key = node.dataset.pointKey;
        const dot = node.querySelector('.of-point-dot');
        const link = node.querySelector('.of-point-link');
        const label = node.querySelector('.of-point-label');
        const box = node.querySelector('.of-point-box');
        if (!key || !dot || !link || !label) return;
        points.set(key, {
          dot: {
            cx: parseFloat(dot.getAttribute('cx') || '0'),
            cy: parseFloat(dot.getAttribute('cy') || '0'),
            r: parseFloat(dot.getAttribute('r') || '0'),
          },
          link: {
            x1: parseFloat(link.getAttribute('x1') || '0'),
            y1: parseFloat(link.getAttribute('y1') || '0'),
            x2: parseFloat(link.getAttribute('x2') || '0'),
            y2: parseFloat(link.getAttribute('y2') || '0'),
          },
          label: {
            x: parseFloat(label.getAttribute('x') || '0'),
            y: parseFloat(label.getAttribute('y') || '0'),
          },
          box: box ? {
            x: parseFloat(box.getAttribute('x') || '0'),
            y: parseFloat(box.getAttribute('y') || '0'),
          } : null,
        });
      });
      return points;
    }

    function animateScatterTransition(prevState) {
      if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
      const scatterArea = area.querySelector('#of-scatter-area');
      if (!scatterArea) return;

      const duration = 420;
      const easeOutCubic = t => 1 - Math.pow(1 - t, 3);
      const active = [];

      scatterArea.querySelectorAll('.of-point[data-point-key]').forEach(node => {
        const key = node.dataset.pointKey;
        const dot = node.querySelector('.of-point-dot');
        const link = node.querySelector('.of-point-link');
        const label = node.querySelector('.of-point-label');
        const box = node.querySelector('.of-point-box');
        if (!key || !dot || !link || !label) return;

        const previous = prevState?.get(key);
        if (!previous) {
          node.animate(
            [
              { opacity: 0, transform: 'translateY(8px) scale(0.96)' },
              { opacity: 1, transform: 'translateY(0) scale(1)' },
            ],
            { duration: 260, easing: 'ease-out', fill: 'forwards' }
          );
          return;
        }

        const finalState = {
          dot: {
            cx: parseFloat(dot.getAttribute('cx') || '0'),
            cy: parseFloat(dot.getAttribute('cy') || '0'),
            r: parseFloat(dot.getAttribute('r') || '0'),
          },
          link: {
            x1: parseFloat(link.getAttribute('x1') || '0'),
            y1: parseFloat(link.getAttribute('y1') || '0'),
            x2: parseFloat(link.getAttribute('x2') || '0'),
            y2: parseFloat(link.getAttribute('y2') || '0'),
          },
          label: {
            x: parseFloat(label.getAttribute('x') || '0'),
            y: parseFloat(label.getAttribute('y') || '0'),
          },
          box: box ? {
            x: parseFloat(box.getAttribute('x') || '0'),
            y: parseFloat(box.getAttribute('y') || '0'),
          } : null,
        };

        dot.setAttribute('cx', previous.dot.cx);
        dot.setAttribute('cy', previous.dot.cy);
        dot.setAttribute('r', previous.dot.r);
        link.setAttribute('x1', previous.link.x1);
        link.setAttribute('y1', previous.link.y1);
        link.setAttribute('x2', previous.link.x2);
        link.setAttribute('y2', previous.link.y2);
        label.setAttribute('x', previous.label.x);
        label.setAttribute('y', previous.label.y);
        if (box && previous.box) {
          box.setAttribute('x', previous.box.x);
          box.setAttribute('y', previous.box.y);
        }

        active.push({ dot, link, label, box, previous, finalState });
      });

      if (!active.length) return;

      const start = performance.now();
      function frame(now) {
        const progress = Math.min(1, (now - start) / duration);
        const eased = easeOutCubic(progress);

        active.forEach(item => {
          const lerp = (from, to) => from + ((to - from) * eased);
          item.dot.setAttribute('cx', lerp(item.previous.dot.cx, item.finalState.dot.cx).toFixed(1));
          item.dot.setAttribute('cy', lerp(item.previous.dot.cy, item.finalState.dot.cy).toFixed(1));
          item.dot.setAttribute('r', lerp(item.previous.dot.r, item.finalState.dot.r).toFixed(2));
          item.link.setAttribute('x1', lerp(item.previous.link.x1, item.finalState.link.x1).toFixed(1));
          item.link.setAttribute('y1', lerp(item.previous.link.y1, item.finalState.link.y1).toFixed(1));
          item.link.setAttribute('x2', lerp(item.previous.link.x2, item.finalState.link.x2).toFixed(1));
          item.link.setAttribute('y2', lerp(item.previous.link.y2, item.finalState.link.y2).toFixed(1));
          item.label.setAttribute('x', lerp(item.previous.label.x, item.finalState.label.x).toFixed(1));
          item.label.setAttribute('y', lerp(item.previous.label.y, item.finalState.label.y).toFixed(1));
          if (item.box && item.previous.box && item.finalState.box) {
            item.box.setAttribute('x', lerp(item.previous.box.x, item.finalState.box.x).toFixed(1));
            item.box.setAttribute('y', lerp(item.previous.box.y, item.finalState.box.y).toFixed(1));
          }
        });

        if (progress < 1) {
          requestAnimationFrame(frame);
        }
      }

      requestAnimationFrame(frame);
    }
    function syncScatterLabelBoxes() {
      const scatterArea = area.querySelector('#of-scatter-area');
      if (!scatterArea) return;
      scatterArea.querySelectorAll('.of-point[data-point-key]').forEach(node => {
        const label = node.querySelector('.of-point-label');
        const box = node.querySelector('.of-point-box');
        if (!label || !box || typeof label.getBBox !== 'function') return;
        const bbox = label.getBBox();
        const padX = 8;
        const padY = 5;
        box.setAttribute('x', (bbox.x - padX).toFixed(1));
        box.setAttribute('y', (bbox.y - padY).toFixed(1));
        box.setAttribute('width', (bbox.width + padX * 2).toFixed(1));
        box.setAttribute('height', (bbox.height + padY * 2).toFixed(1));
      });
    }
    function bindScatterLabelDrag() {
      const scatterArea = area.querySelector('#of-scatter-area');
      const svg = scatterArea?.querySelector('svg.of-svg');
      if (!scatterArea || !svg) return;

      const toSvgPoint = (clientX, clientY) => {
        const pt = svg.createSVGPoint();
        pt.x = clientX;
        pt.y = clientY;
        const ctm = svg.getScreenCTM();
        return ctm ? pt.matrixTransform(ctm.inverse()) : { x: 0, y: 0 };
      };

      scatterArea.querySelectorAll('.of-point[data-point-key]').forEach(node => {
        const label = node.querySelector('.of-point-label');
        const box = node.querySelector('.of-point-box');
        const dot = node.querySelector('.of-point-dot');
        const link = node.querySelector('.of-point-link');
        if (!label || !dot || !link) return;

        const startDrag = event => {
          event.preventDefault();
          const start = toSvgPoint(event.clientX, event.clientY);
          const startX = parseFloat(label.getAttribute('x') || '0');
          const startY = parseFloat(label.getAttribute('y') || '0');
          const boxStartX = box ? parseFloat(box.getAttribute('x') || '0') : 0;
          const boxStartY = box ? parseFloat(box.getAttribute('y') || '0') : 0;
          label.classList.add('is-dragging');
          box?.classList.add('is-dragging');
          (box || label).setPointerCapture?.(event.pointerId);

          const onMove = moveEvent => {
            const current = toSvgPoint(moveEvent.clientX, moveEvent.clientY);
            const nextX = startX + (current.x - start.x);
            const nextY = startY + (current.y - start.y);
            label.setAttribute('x', nextX.toFixed(1));
            label.setAttribute('y', nextY.toFixed(1));
            if (box) {
              box.setAttribute('x', (boxStartX + (current.x - start.x)).toFixed(1));
              box.setAttribute('y', (boxStartY + (current.y - start.y)).toFixed(1));
            }
            link.setAttribute('x2', (nextX + 2).toFixed(1));
            link.setAttribute('y2', nextY.toFixed(1));
            syncScatterLabelBoxes();
          };

          const onUp = () => {
            label.classList.remove('is-dragging');
            box?.classList.remove('is-dragging');
            (box || label).releasePointerCapture?.(event.pointerId);
            (box || label).removeEventListener('pointermove', onMove);
            (box || label).removeEventListener('pointerup', onUp);
            (box || label).removeEventListener('pointercancel', onUp);
            const finalX = parseFloat(label.getAttribute('x') || '0');
            const finalY = parseFloat(label.getAttribute('y') || '0');
            const dotX = parseFloat(dot.getAttribute('cx') || '0');
            const dotY = parseFloat(dot.getAttribute('cy') || '0');
            persistScatterLabelOffset(node.dataset.pointKey, finalX - dotX, finalY - dotY);
            syncScatterLabelBoxes();
          };

          (box || label).addEventListener('pointermove', onMove);
          (box || label).addEventListener('pointerup', onUp);
          (box || label).addEventListener('pointercancel', onUp);
        };

        label.addEventListener('pointerdown', startDrag);
        box?.addEventListener('pointerdown', startDrag);
      });
    }
    function niceNumber(range, round) {
      const exponent = Math.floor(Math.log10(range || 1));
      const fraction = (range || 1) / Math.pow(10, exponent);
      let niceFraction;
      if (round) {
        if (fraction < 1.5) niceFraction = 1;
        else if (fraction < 3) niceFraction = 2;
        else if (fraction < 7) niceFraction = 5;
        else niceFraction = 10;
      } else {
        if (fraction <= 1) niceFraction = 1;
        else if (fraction <= 2) niceFraction = 2;
        else if (fraction <= 5) niceFraction = 5;
        else niceFraction = 10;
      }
      return niceFraction * Math.pow(10, exponent);
    }
    function buildNiceTicks(min, max, tickCount = 6) {
      const safeMin = Number.isFinite(min) ? min : 0;
      const safeMax = Number.isFinite(max) ? max : 1;
      if (safeMin === safeMax) {
        const pad = Math.abs(safeMin || 1) * 0.1 || 1;
        min = safeMin - pad;
        max = safeMax + pad;
      } else {
        min = safeMin;
        max = safeMax;
      }
      const range = niceNumber(max - min, false);
      const step = niceNumber(range / Math.max(1, tickCount - 1), true);
      const niceMin = Math.floor(min / step) * step;
      const niceMax = Math.ceil(max / step) * step;
      const ticks = [];
      for (let value = niceMin, i = 0; value <= niceMax + step * 0.5 && i < 12; value += step, i += 1) {
        ticks.push(Number(value.toFixed(10)));
      }
      return { ticks, min: niceMin, max: niceMax, step };
    }
    function formatAxisTick(value, step) {
      const absStep = Math.abs(step || 1);
      if (absStep >= 1) return `${Math.round(value)}`;
      if (absStep >= 0.1) return value.toFixed(1);
      return value.toFixed(2);
    }

    /* ── Render scatter chart + side table ── */
    function renderScatterWithTable(xKey, yKey, period, mode, visibleKeys) {
      const metrics = getMetricsForPeriod(mode, period);
      const xMeta = metrics.find(m => m.key === xKey) || metrics[0];
      const yMeta = metrics.find(m => m.key === yKey) || (metrics[1] || metrics[0]);
      const xIdx  = getColIdx(xMeta, period, mode);
      const yIdx  = getColIdx(yMeta, period, mode);

      const rawPoints = displayRows
        .filter(item => !visibleKeys || visibleKeys.has(item.key))
        .map((item, i) => {
          const xRaw = get(item.row, xIdx), yRaw = get(item.row, yIdx);
          const xVal = parseFloat(xRaw.replace(/,/g,'')), yVal = parseFloat(yRaw.replace(/,/g,''));
          if (isNaN(xVal) || isNaN(yVal)) return null;
          const links = linksByRowKey[item.key] || [];
          const colorKey = links[0]?.code || item.key;
          const colorIdx = S.pointColors[colorKey] !== undefined ? S.pointColors[colorKey] : undefined;
          const color = colorIdx !== undefined ? OF_POINT_COLORS[colorIdx]?.dot : DOT_COLORS[i % DOT_COLORS.length];
          const label = links.map(f => f.code).join(',') || item.name.slice(0,14);
          return { x: xVal, y: yVal, color, label, r: 8, xRaw, yRaw, name: item.name, key: encodeURIComponent(item.key) };
        }).filter(Boolean);

      const plotCount  = rawPoints.length;
      const totalCount = displayRows.filter(item => !visibleKeys || visibleKeys.has(item.key)).length;
      let scatterHtml  = '<div class="of-no-data">ไม่พบข้อมูลสำหรับสร้างกราฟ</div>';

      if (rawPoints.length >= 1) {
        const W=780, H=680, pL=74, pR=28, pT=28, pB=70;
        const xs = rawPoints.map(p=>p.x), ys = rawPoints.map(p=>p.y);
        const minX=Math.min(...xs), maxX=Math.max(...xs);
        const minY=Math.min(...ys), maxY=Math.max(...ys);
        const dx=maxX-minX||1, dy=maxY-minY||1;
        const xNice = buildNiceTicks(minX-dx*0.15, maxX+dx*0.15, 6);
        const yNice = buildNiceTicks(minY-dy*0.20, maxY+dy*0.15, 6);
        const x0=xNice.min, x1=xNice.max;
        const y0=yNice.min, y1=yNice.max;
        const plotW=W-pL-pR, plotH=H-pT-pB;
        const sx = v => pL + ((v-x0)/(x1-x0||1))*plotW;
        const sy = v => pT + plotH - ((v-y0)/(y1-y0||1))*plotH;
        const xTicks = xNice.ticks;
        const yTicks = yNice.ticks;

        // Attach scale info for label placer
        rawPoints.forEach(p => { p._x0=x0; p._xRange=x1-x0||1; p._y0=y0; p._yRange=y1-y0||1; });
        const placed = placeLabels(rawPoints, W, H, pL, pR, pT, pB);

        scatterHtml = `<svg viewBox="0 0 ${W} ${H}" class="of-svg" xmlns="http://www.w3.org/2000/svg">
          <rect width="${W}" height="${H}" fill="#fff"/>
          ${yTicks.map(t=>`<line x1="${pL}" y1="${sy(t).toFixed(1)}" x2="${W-pR}" y2="${sy(t).toFixed(1)}" stroke="#e2e8f0" stroke-width="1"/>
            <text x="${pL-8}" y="${(sy(t)+4).toFixed(1)}" text-anchor="end" fill="#64748b" font-size="16" font-weight="500">${formatAxisTick(t, yNice.step)}</text>`).join('')}
          ${xTicks.map(t=>`<line x1="${sx(t).toFixed(1)}" y1="${pT}" x2="${sx(t).toFixed(1)}" y2="${H-pB}" stroke="#e2e8f0" stroke-width="1"/>
            <text x="${sx(t).toFixed(1)}" y="${H-pB+18}" text-anchor="middle" fill="#64748b" font-size="16" font-weight="500">${formatAxisTick(t, xNice.step)}</text>`).join('')}
          <line x1="${pL}" y1="${H-pB}" x2="${W-pR}" y2="${H-pB}" stroke="#94a3b8" stroke-width="1.5"/>
          <line x1="${pL}" y1="${pT}"   x2="${pL}"   y2="${H-pB}" stroke="#94a3b8" stroke-width="1.5"/>
          ${placed.map(p=>`
            <g class="of-point" data-point-key="${p.key}">
              <line class="of-point-link" x1="${p.cx.toFixed(1)}" y1="${p.cy.toFixed(1)}" x2="${(p.lx+2).toFixed(1)}" y2="${(p.ly).toFixed(1)}" stroke="#b0bec5" stroke-width="0.9" stroke-dasharray="3,2"/>
              <circle class="of-point-dot" cx="${p.cx.toFixed(1)}" cy="${p.cy.toFixed(1)}" r="8" fill="${p.color}" opacity="0.92"/>
              <rect class="of-point-box" x="${(p.lx-3).toFixed(1)}" y="${(p.ly-18).toFixed(1)}" width="${(p.labelW+6).toFixed(1)}" height="24" rx="6" fill="rgba(255,255,255,0.94)" stroke="${p.color}" stroke-width="1.4"/>
              <text class="of-point-label" x="${p.lx.toFixed(1)}" y="${p.ly.toFixed(1)}" fill="#1e293b" font-size="16" font-weight="700" paint-order="stroke" stroke="#fff" stroke-width="3">${esc(p.label)}</text>
            </g>
          `).join('')}
          <text x="${W/2}" y="${H-6}" text-anchor="middle" fill="#334155" font-size="20" font-weight="700">${esc(xMeta.label)} ${esc(period)}</text>
          <text x="16" y="${H/2}" text-anchor="middle" fill="#334155" font-size="20" font-weight="700" transform="rotate(-90,16,${H/2})">${esc(yMeta.label)} ${esc(period)}</text>
        </svg>`;
      }

      const sidePoints = [...rawPoints];
      if (S.sideSort?.key && S.sideSort?.dir) {
        const sortKey = S.sideSort.key;
        const sortDir = S.sideSort.dir;
        sidePoints.sort((a, b) => {
          const aNum = parseFloat(sortKey === 'x' ? a.xRaw : a.yRaw);
          const bNum = parseFloat(sortKey === 'x' ? b.xRaw : b.yRaw);
          if (isNaN(aNum) && isNaN(bNum)) return a.label.localeCompare(b.label);
          if (isNaN(aNum)) return 1;
          if (isNaN(bNum)) return -1;
          if (aNum === bNum) return a.label.localeCompare(b.label);
          return sortDir === 1 ? aNum - bNum : bNum - aNum;
        });
      }

      const sideSortIcon = (key) => {
        if (S.sideSort?.key !== key || !S.sideSort?.dir) return ' ↕';
        return S.sideSort.dir === -1 ? ' ↓' : ' ↑';
      };

      const sidePeriodLabel = period ? ` (${period})` : '';
      const sideRows = sidePoints.map(p=>{
        const xNum = parseFloat(p.xRaw), yNum = parseFloat(p.yRaw);
        return `<tr>
          <td class="of-side-fund-td"><div class="of-side-fund-inner"><span class="of-dot" style="background:${p.color}"></span><span>${esc(p.label)}</span></div></td>
          <td class="of-side-val of-side-val-mid">${isNaN(xNum)?'—':xNum.toFixed(2)}</td>
          <td class="of-side-val of-side-val-mid">${isNaN(yNum)?'—':yNum.toFixed(2)}</td>
        </tr>`;
      }).join('');

      return {
        html: `<div class="of-scatter-section">
          <div class="of-scatter-left">${scatterHtml}</div>
          <div class="of-scatter-right">
            <table class="of-side-table">
              <colgroup>
                <col class="of-side-col-fund">
                <col class="of-side-col-metric">
                <col class="of-side-col-metric">
              </colgroup>
              <thead><tr>
                <th class="of-side-th">กองทุน</th>
                <th class="of-side-th of-side-th-sort of-side-val-mid" data-side-sort="x" role="button" tabindex="0" aria-sort="${S.sideSort?.key==='x'?(S.sideSort.dir===1?'ascending':S.sideSort.dir===-1?'descending':'none'):'none'}">${esc(xMeta.label)}${esc(sidePeriodLabel)}${sideSortIcon('x')}</th>
                <th class="of-side-th of-side-th-sort of-side-val-mid" data-side-sort="y" role="button" tabindex="0" aria-sort="${S.sideSort?.key==='y'?(S.sideSort.dir===1?'ascending':S.sideSort.dir===-1?'descending':'none'):'none'}">${esc(yMeta.label)}${esc(sidePeriodLabel)}${sideSortIcon('y')}</th>
              </tr></thead>
              <tbody>${sideRows||'<tr><td colspan="3" class="of-no-data-row">ไม่มีข้อมูล</td></tr>'}</tbody>
            </table>
          </div>
        </div>`,
        plotCount, totalCount, xLabel: xMeta.label, yLabel: yMeta.label,
      };
    }

    /* ── Bottom table ── */
    function renderTable(period, mode, visibleKeys) {
      const metrics = getMetricsForPeriod(mode, period);
      function fmtVal(v) {
        if (!v||v==='-') return '<span class="of-na">—</span>';
        const n = parseFloat(v.replace(/,/g,''));
        if (isNaN(n)) return `<span class="of-na">${esc(v)}</span>`;
        return `<span class="${n<0?'of-neg':n>0?'of-pos':'of-zero'}">${n.toFixed(2)}</span>`;
      }
      const rows = displayRows.map((item,i) => {
        const links = linksByRowKey[item.key]||[];
        const colorKey = links[0]?.code || item.key;
        const colorIdx = S.pointColors[colorKey] !== undefined ? S.pointColors[colorKey] : undefined;
        const dotColor = colorIdx!==undefined ? OF_POINT_COLORS[colorIdx].dot : DOT_COLORS[i%DOT_COLORS.length];
        const isVis = !visibleKeys||visibleKeys.has(item.key);
        const primaryThaiCode = links[0]?.code || '';
        const pointPicker = primaryThaiCode ? buildHighlightDotPicker(primaryThaiCode, S.pointColors[primaryThaiCode], dotColor) : '';
        const cells = metrics.map(m=>`<td class="of-td">${fmtVal(get(item.row,getColIdx(m,period,mode)))}</td>`).join('');
        return `<tr class="${isVis?'':'of-row-hidden'}" data-row-key="${esc(item.key)}">
          <td class="of-td of-td-cb"><input type="checkbox" class="of-cb" data-row-key="${esc(item.key)}" ${isVis?'checked':''}></td>
          <td class="of-td of-td-hl">${pointPicker}</td>
          <td class="of-td of-td-name"><span class="of-name-text" title="${esc(item.name)}">${esc(item.name.length>32?item.name.slice(0,30)+'…':item.name)}</span></td>
          <td class="of-td of-td-thai">${esc(links.map(f=>f.code).join(', '))}</td>
          ${cells}
        </tr>`;
      }).join('');
      const metricHeaders = metrics.map(m=>`<th class="of-th">${esc(m.label)}</th>`).join('');
      return `<div class="of-table-wrap"><table class="of-table" id="of-bottom-table">
        <thead><tr>
          <th class="of-th of-th-cb"><input type="checkbox" id="of-cb-all" ${[...visibleKeys].length===displayRows.length?'checked':''} title="เลือก/ยกเลิกทั้งหมด"></th>
          <th class="of-th of-th-hl">สี</th>
          <th class="of-th of-th-name">Master Fund</th>
          <th class="of-th of-th-thai">กองทุนไทย</th>
          ${metricHeaders}
        </tr></thead>
        <tbody>${rows}</tbody>
      </table></div>`;
    }

    /* ── Helpers ── */
    const source = CONFIG.PAGES['master-placeholder-5']?.source||'AVP Master Fund ID';
    function buildPeriodBtns(mode, active, periods) {
      return (periods || (mode==='calendar'?CALENDAR_YEARS:ANNUALIZED_PERIODS))
        .map(p=>`<button class="of-period-btn${p===active?' is-active':''}" data-of-period="${p}" type="button">${p}</button>`).join('');
    }
    function buildMetricOpts(mode, sel, period) {
      return getMetricsForPeriod(mode, period)
        .map(m=>`<option value="${m.key}"${m.key===sel?' selected':''}>${esc(m.label)}</option>`).join('');
    }
    function buildScatterTitle(xLabel, yLabel, period) {
      return `กราฟเปรียบเทียบ ${xLabel} และ ${yLabel} (${period})`;
    }
    function buildOtherFactorsClipboardPlainText() {
      const title = area.querySelector('#of-scatter-title')?.textContent?.trim() || CONFIG.PAGES['master-placeholder-5']?.title || 'ปัจจัยประกอบอื่นๆ';
      const table = area.querySelector('.of-side-table');
      const lines = [title, `ที่มา : ${source}`];
      if (table) {
        Array.from(table.rows || []).forEach(row => {
          lines.push(Array.from(row.cells || []).map(cell => (cell.textContent || '').replace(/\s+/g, ' ').trim()).join('\t'));
        });
      }
      return lines.join('\n');
    }
    async function svgToPngDataUrl(svg, width = 860, height = 500, scale = 2) {
      const svgClone = svg.cloneNode(true);
      svgClone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      svgClone.setAttribute('width', String(width));
      svgClone.setAttribute('height', String(height));
      const exportFont = 'TH Sarabun New, THSarabunNew, Sarabun, Arial, sans-serif';
      svgClone.querySelectorAll('text').forEach(text => {
        text.setAttribute('font-family', exportFont);
        text.style.fontFamily = exportFont;
      });
      svgClone.querySelectorAll('.of-point').forEach(point => {
        const label = point.querySelector('.of-point-label');
        const box = point.querySelector('.of-point-box');
        if (!label || !box) return;
        const text = (label.textContent || '').trim();
        const fontSize = parseFloat(label.getAttribute('font-size') || '16') || 16;
        const estimatedWidth = Math.max(34, Math.ceil(text.length * fontSize * 0.46) + 22);
        const labelX = parseFloat(label.getAttribute('x') || '0');
        const labelY = parseFloat(label.getAttribute('y') || '0');
        box.setAttribute('x', (labelX - 9).toFixed(1));
        box.setAttribute('y', (labelY - 19).toFixed(1));
        box.setAttribute('width', String(estimatedWidth));
        box.setAttribute('height', '27');
      });
      const serialized = new XMLSerializer().serializeToString(svgClone);
      const blob = new Blob([serialized], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      try {
        const img = await new Promise((resolve, reject) => {
          const node = new Image();
          node.onload = () => resolve(node);
          node.onerror = reject;
          node.src = url;
        });
        const canvas = document.createElement('canvas');
        canvas.width = Math.round(width * scale);
        canvas.height = Math.round(height * scale);
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        return canvas.toDataURL('image/png');
      } finally {
        URL.revokeObjectURL(url);
      }
    }
    function buildOtherFactorsSvgExport() {
      const svg = area.querySelector('#of-scatter-area .of-scatter-left svg.of-svg');
      if (!svg) return null;
      const svgClone = svg.cloneNode(true);
      const viewBox = (svg.getAttribute('viewBox') || '').split(/\s+/).map(Number);
      const exportWidth = Number.isFinite(viewBox[2]) ? viewBox[2] : (svg.clientWidth || 780);
      const exportHeight = Number.isFinite(viewBox[3]) ? viewBox[3] : (svg.clientHeight || 680);
      const exportFont = 'TH Sarabun New, THSarabunNew, Sarabun, Arial, sans-serif';
      svgClone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      svgClone.setAttribute('width', String(exportWidth));
      svgClone.setAttribute('height', String(exportHeight));
      svgClone.querySelectorAll('text').forEach(text => {
        text.setAttribute('font-family', exportFont);
        text.style.fontFamily = exportFont;
      });
      const markup = `<?xml version="1.0" encoding="UTF-8"?>\n${new XMLSerializer().serializeToString(svgClone)}`;
      const rawTitle = area.querySelector('#of-scatter-title')?.textContent?.trim() || 'other-factors-chart';
      const safeFilename = rawTitle
        .replace(/[\\/:*?"<>|]/g, '-')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 120) || 'other-factors-chart';
      return { filename: `${safeFilename}.svg`, markup };
    }
    function downloadOtherFactorsSvg() {
      const exported = buildOtherFactorsSvgExport();
      if (!exported?.markup) {
        toast('ไม่พบกราฟสำหรับดาวน์โหลด SVG', 'warning');
        return;
      }
      const blob = new Blob([exported.markup], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = exported.filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
      toast('ดาวน์โหลดกราฟเป็น SVG แล้ว', 'success');
    }
    function cleanOtherFactorsHeaderText(text) {
      return String(text || '')
        .replace(/[↕↑↓]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    }
    function buildOtherFactorsSideTableClipboardHtml(sideTable) {
      const headers = Array.from(sideTable.tHead?.rows?.[0]?.cells || [])
        .map(cell => cleanOtherFactorsHeaderText(cell.textContent));
      const bodyRows = Array.from(sideTable.tBodies?.[0]?.rows || []).map(row =>
        Array.from(row.cells || []).map(cell => (cell.textContent || '').replace(/\s+/g, ' ').trim())
      );
      const widths = [260, 115, 115];
      const tableWidth = widths.reduce((sum, width) => sum + width, 0);
      const clipboardFontStyle = presentationClipboardFontStyle();
      const fontFamily = "font-family:'TH Sarabun New','THSarabunNew','Sarabun',Arial,sans-serif";
      const headerStyle = [
        fontFamily,
        clipboardFontStyle,
        'background:#1a2744',
        'color:#ffffff',
        'font-weight:700',
        'text-align:center',
        'vertical-align:middle',
        'border:1px solid #dbe4f0',
        'padding:9px 12px',
        'line-height:1.15',
        'white-space:normal',
      ].join(';');
      const cellStyle = (rowIndex, cellIndex) => [
        fontFamily,
        clipboardFontStyle,
        `background:${rowIndex % 2 === 0 ? '#ffffff' : '#f8f8f8'}`,
        'color:#334155',
        'font-weight:500',
        `text-align:${cellIndex === 0 ? 'center' : 'center'}`,
        'vertical-align:middle',
        'border:1px solid #dbe4f0',
        'padding:7px 9px',
        'line-height:1.2',
        'white-space:normal',
        'word-break:break-word',
      ].join(';');

      return `<table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;table-layout:fixed;width:${tableWidth}px;min-width:${tableWidth}px;background:#ffffff;${clipboardFontStyle};">
        <colgroup>${widths.map(width => `<col style="width:${width}px;">`).join('')}</colgroup>
        <thead>
          <tr>${headers.map(label => `<th style="${headerStyle}">${esc(label) || '&nbsp;'}</th>`).join('')}</tr>
        </thead>
        <tbody>
          ${bodyRows.map((row, rowIndex) => `
            <tr>${row.map((text, cellIndex) => `<td style="${cellStyle(rowIndex, cellIndex)}">${esc(text) || '&nbsp;'}</td>`).join('')}</tr>
          `).join('')}
        </tbody>
      </table>`;
    }
    async function buildOtherFactorsClipboardHtml() {
      const title = area.querySelector('#of-scatter-title')?.textContent?.trim() || CONFIG.PAGES['master-placeholder-5']?.title || 'ปัจจัยประกอบอื่นๆ';
      const svg = area.querySelector('#of-scatter-area .of-scatter-left svg');
      const sideTable = area.querySelector('#of-scatter-area .of-side-table');
      if (!svg || !sideTable) return '';

      const graphImage = await svgToPngDataUrl(svg, 860, 500, 2);
      const sideTableHtml = buildOtherFactorsSideTableClipboardHtml(sideTable);
      const clipboardFontStyle = presentationClipboardFontStyle();
      const fontFamily = "font-family:'TH Sarabun New','THSarabunNew','Sarabun',Arial,sans-serif";

      return `<!DOCTYPE html>
      <html lang="th">
        <head><meta charset="utf-8"></head>
        <body lang="TH" style="margin:0;padding:0;background:#ffffff;${fontFamily};${clipboardFontStyle};color:#334155;mso-fareast-language:TH;">
          <div style="display:inline-block;background:#ffffff;${fontFamily};${clipboardFontStyle};">
            <div lang="TH" style="${fontFamily};${clipboardFontStyle};line-height:1.2;font-weight:700;color:#1a3c6e;margin:0 0 4px;mso-fareast-language:TH;">${esc(CONFIG.PAGES['master-placeholder-5']?.title || 'ปัจจัยประกอบอื่นๆ')}</div>
            <div lang="TH" style="${fontFamily};${clipboardFontStyle};line-height:1.2;color:#1a2744;font-weight:700;margin:0 0 4px;mso-fareast-language:TH;">${esc(title)}</div>
            <div lang="TH" style="${fontFamily};${clipboardFontStyle};line-height:1.2;color:#64748b;margin:0 0 10px;mso-fareast-language:TH;">ที่มา : ${esc(source)}</div>
            <div style="display:block;white-space:nowrap;background:#ffffff;">
              <div style="display:inline-block;vertical-align:top;margin:0 18px 0 0;padding:0;background:#ffffff;">
                <img src="${graphImage}" width="860" height="500" style="display:block;width:860px;height:500px;border:0;background:#ffffff;" alt="${esc(title)}">
              </div>
              <div style="display:inline-block;vertical-align:top;margin:0;padding:0;background:#ffffff;white-space:normal;">
                ${sideTableHtml}
              </div>
            </div>
          </div>
        </body>
      </html>`;
    }
    async function copyOtherFactorsClipboard() {
      const html = await buildOtherFactorsClipboardHtml();
      if (!html) throw new Error('ไม่พบกราฟหรือตารางด้านขวาสำหรับคัดลอก');
      return copyHtmlToClipboard(html, buildOtherFactorsClipboardPlainText(), 'custom-html', { allowTextFallback: false });
    }

    const exportSvgAction = `<button class="btn btn-ghost of-export-svg-btn" id="of-download-svg" type="button" title="ดาวน์โหลดกราฟเป็นไฟล์ SVG">ดาวน์โหลด SVG</button>`;

    /* ── Initial render ── */
    const initialSelection = normalizeOtherFactorsSelection();
    const {html:initScatter, plotCount, totalCount, xLabel, yLabel} = renderScatterWithTable(S.xKey, S.yKey, S.period, S.mode, S.visibleKeys);

    area.innerHTML = `
      ${pageToolActions('master-placeholder-5', source, exportSvgAction)}
      <div class="card report-card" id="report-card">
        <div class="of-topbar">
          <div class="of-topbar-left">
            <div class="of-period-group" id="of-period-btns">${buildPeriodBtns(S.mode, S.period, initialSelection.availablePeriods)}</div>
            <span class="of-axis-label">แกน X</span>
            <select class="of-axis-select" id="of-select-x">${buildMetricOpts(S.mode, S.xKey, S.period)}</select>
            <span class="of-axis-label">แกน Y</span>
            <select class="of-axis-select" id="of-select-y">${buildMetricOpts(S.mode, S.yKey, S.period)}</select>
          </div>
          <div class="of-topbar-right">
            <div class="view-toggle" role="tablist">
              <button class="btn btn-ghost view-toggle-btn${S.mode==='annualized'?' is-active':''}" data-of-mode="annualized" type="button">Annualized</button>
              <button class="btn btn-ghost view-toggle-btn${S.mode==='calendar'?' is-active':''}" data-of-mode="calendar" type="button">Calendar</button>
            </div>
            <button class="btn btn-ghost of-reset-label-btn" id="of-reset-labels" type="button">รีเซ็ตตำแหน่งชื่อ</button>
            <span class="of-count" id="of-count">${plotCount} / ${totalCount} กองบนกราฟ</span>
          </div>
        </div>
        <div class="of-scatter-title" id="of-scatter-title">${esc(buildScatterTitle(xLabel, yLabel, S.period))}</div>
        <div id="of-scatter-area">${initScatter}</div>
        <div id="of-table-container">${renderTable(S.period, S.mode, S.visibleKeys)}</div>
      </div>
      <style>
        .of-page-tools{padding:8px 0 10px;display:flex;gap:8px;align-items:center}
        .of-topbar{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;padding:10px 16px;border-bottom:1px solid #e2e8f0;background:#f8fafc}
        .of-topbar-left,.of-topbar-right{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
        .of-period-group{display:flex;gap:3px;flex-wrap:wrap}
        .of-period-btn{padding:4px 13px;border:1px solid #d0d9e8;border-radius:20px;background:#fff;color:#334155;font-size:0.82rem;cursor:pointer;font-family:inherit;transition:all .15s}
        .of-period-btn.is-active{background:#1a2744;color:#fff;border-color:#1a2744;font-weight:700}
        .of-period-btn:hover:not(.is-active){background:#eef2f7}
        .of-axis-label{font-size:0.8rem;color:#64748b;font-weight:600;white-space:nowrap}
        .of-axis-select{padding:4px 8px;border:1px solid #d0d9e8;border-radius:6px;background:#fff;color:#1e293b;font-size:0.8rem;font-family:inherit;cursor:pointer;max-width:200px}
        .of-count{font-size:0.79rem;color:#475569;padding:4px 12px;background:#e8edf5;border-radius:20px;white-space:nowrap}
        .of-reset-label-btn{padding:6px 12px;font-size:0.8rem}
        .of-scatter-title{padding:12px 16px;font-size:1.375rem;font-weight:700;color:#1a2744;background:#eef2fa;border-bottom:1px solid #dbe4f0;letter-spacing:.01em;line-height:1.25}
        .of-scatter-section{display:flex;align-items:flex-start;gap:24px;padding:0 16px 0 0;border-bottom:2px solid #e2e8f0;background:#fff}
        .of-scatter-left{flex:4;min-width:0;padding:12px 0 12px 12px}
        .of-scatter-right{flex:3;min-width:260px;overflow-y:auto;max-height:530px;padding-left:4px}
        .of-svg{width:100%;display:block}
        .of-no-data{padding:60px 20px;text-align:center;color:#94a3b8;font-size:0.9rem}
        .of-side-table{width:100%;border-collapse:collapse;font-size:0.8rem;font-family:'Sarabun','THSarabunNew',sans-serif;table-layout:fixed}
        .of-side-col-fund{width:52%}
        .of-side-col-metric{width:24%}
        .of-side-th{padding:8px 10px;background:#1a2744;color:#fff;font-weight:600;font-size:0.75rem;white-space:nowrap;position:sticky;top:0;z-index:1}
        .of-side-th-sort{cursor:pointer;user-select:none}
        .of-side-th-sort:hover{background:#223257}
        .of-side-val{text-align:right}
        .of-side-val-mid{text-align:center!important}
        .of-side-table tbody tr{border-bottom:1px solid #eef2f7}
        .of-side-table tbody tr:nth-child(even){background:#f8fafc}
        .of-side-table tbody tr:hover td{background:#e8f0fb!important}
        .of-side-fund-td{padding:5px 10px;vertical-align:middle}
        .of-side-fund-inner{display:flex;align-items:center;gap:6px;font-size:0.78rem;font-weight:600;color:#1e293b}
        .of-side-table td.of-side-val{padding:5px 10px;text-align:right;color:#334155;font-weight:500;font-size:0.8rem}
        #of-table-container{margin-top:18px}
        .of-table-wrap{overflow-x:auto}
        .of-table{width:100%;border-collapse:collapse;font-size:0.79rem;font-family:'Sarabun','THSarabunNew',sans-serif}
        .of-th{padding:8px 10px;background:#1a2744;color:#fff;font-weight:600;font-size:0.73rem;white-space:nowrap;text-align:center;border-right:1px solid #2d3f6b}
        .of-th-cb{width:36px;text-align:center;background:#0f1f3d}
        .of-th-hl{width:72px;background:#0f1f3d}
        .of-th-name{text-align:left;min-width:180px;background:#0f1f3d}
        .of-th-thai{text-align:left;min-width:110px}
        .of-td{padding:6px 10px;border-bottom:1px solid #eef2f7;border-right:1px solid #f1f5fb;text-align:right;vertical-align:middle}
        .of-td-cb{text-align:center;background:#f8fafc;width:36px}
        .of-td-hl{text-align:center;background:#f8fafc;position:relative;overflow:visible}
        .of-td-name{text-align:left;background:#f8fafc;display:flex;align-items:center;gap:6px}
        .of-td-thai{text-align:left;color:#475569;font-size:0.76rem}
        .of-row-hidden td{opacity:.35}
        .of-table tbody tr:hover td{background:#f1f5fb!important}
        .of-dot{display:inline-block;width:9px;height:9px;border-radius:50%;flex-shrink:0}
        .of-point-box{cursor:grab}
        .of-point-box.is-dragging{cursor:grabbing}
        .of-point-label{cursor:grab;user-select:none}
        .of-point-label.is-dragging{cursor:grabbing}
        .of-name-text{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:200px;display:inline-block;vertical-align:middle}
        .of-color-picker{position:relative;display:inline-flex;align-items:center;justify-content:center}
        .of-color-trigger{width:18px;height:18px;border-radius:999px;border:2px solid var(--picker);box-shadow:0 0 0 2px #fff inset;cursor:pointer;display:inline-block}
        .of-color-trigger.is-custom{box-shadow:0 0 0 1px #fff inset,0 0 0 1px rgba(15,23,42,0.06)}
        .of-color-palette{position:absolute;top:24px;left:50%;transform:translateX(-50%);display:grid;grid-template-columns:repeat(4,18px);gap:8px;padding:10px;background:#fff;border:1px solid #dbe4f0;border-radius:12px;box-shadow:0 10px 24px rgba(15,23,42,0.14);z-index:12}
        .of-color-palette[hidden]{display:none!important}
        .of-palette-swatch{width:18px;height:18px;border-radius:999px;border:2px solid #fff;box-shadow:0 0 0 1px rgba(15,23,42,0.14);cursor:pointer;padding:0}
        .of-palette-swatch.is-active{box-shadow:0 0 0 2px #1a2744}
        .of-palette-swatch.is-clear{width:18px;height:18px;border-radius:999px;background:#f8fafc;color:#64748b;border:1px solid #cbd5e1;font-size:12px;line-height:1}
        .of-pos{color:#16a34a;font-weight:600}.of-neg{color:#dc2626;font-weight:600}.of-zero{color:#64748b}.of-na{color:#94a3b8;font-size:0.75rem}
        .of-cb{cursor:pointer;width:14px;height:14px}
        .of-no-data-row{padding:12px;text-align:center;color:#94a3b8}
        @media(max-width:900px){.of-scatter-section{flex-direction:column;gap:12px;padding:0}.of-scatter-left{padding:12px}.of-scatter-right{flex:none;width:100%;max-height:240px;padding-left:0;border-top:1px solid #e2e8f0}}
      </style>`;

    /* ── Event helpers ── */
    function syncOtherFactorsControls() {
      const { availablePeriods } = normalizeOtherFactorsSelection();
      const periodWrap = area.querySelector('#of-period-btns');
      const xSelect = area.querySelector('#of-select-x');
      const ySelect = area.querySelector('#of-select-y');
      if (periodWrap) periodWrap.innerHTML = buildPeriodBtns(S.mode, S.period, availablePeriods);
      if (xSelect) xSelect.innerHTML = buildMetricOpts(S.mode, S.xKey, S.period);
      if (ySelect) ySelect.innerHTML = buildMetricOpts(S.mode, S.yKey, S.period);
      if (xSelect) xSelect.value = S.xKey;
      if (ySelect) ySelect.value = S.yKey;
    }
    function refreshScatter() {
      syncOtherFactorsControls();
      const prevScatterState = captureScatterPointState(area.querySelector('#of-scatter-area'));
      const {html,plotCount:pc,totalCount:tc,xLabel:xl,yLabel:yl} = renderScatterWithTable(S.xKey,S.yKey,S.period,S.mode,S.visibleKeys);
      area.querySelector('#of-scatter-area').innerHTML = html;
      area.querySelector('#of-scatter-title').textContent = buildScatterTitle(xl, yl, S.period);
      area.querySelector('#of-count').textContent = `${pc} / ${tc} กองบนกราฟ`;
      animateScatterTransition(prevScatterState);
      syncScatterLabelBoxes();
      bindSideTableEvents();
      bindScatterLabelDrag();
    }
    function fullRefresh() {
      refreshScatter();
      area.querySelector('#of-table-container').innerHTML = renderTable(S.period, S.mode, S.visibleKeys);
      bindBottomTableEvents();
    }
    function bindSideTableEvents() {
      area.querySelectorAll('[data-side-sort]').forEach(th => {
        const cycleSort = () => {
          const key = th.dataset.sideSort;
          const currentKey = S.sideSort?.key || null;
          const currentDir = S.sideSort?.dir || 0;
          if (currentKey !== key || currentDir === 0) S.sideSort = { key, dir: -1 };
          else if (currentDir === -1) S.sideSort = { key, dir: 1 };
          else S.sideSort = { key: null, dir: 0 };
          refreshScatter();
        };
        th.addEventListener('click', cycleSort);
        th.addEventListener('keydown', e => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            cycleSort();
          }
        });
      });
    }
    function bindBottomTableEvents() {
      const closeAllColorPalettes = () => {
        area.querySelectorAll('.of-color-palette').forEach(palette => { palette.hidden = true; });
      };
      area.querySelectorAll('.of-color-trigger').forEach(btn => {
        btn.onclick = event => {
          event.preventDefault();
          event.stopPropagation();
          const picker = btn.closest('.of-color-picker');
          const palette = picker?.querySelector('.of-color-palette');
          if (!palette) return;
          const willOpen = palette.hidden;
          closeAllColorPalettes();
          palette.hidden = !willOpen;
        };
      });
      area.querySelectorAll('.of-palette-swatch').forEach(btn => {
        btn.onclick = event => {
          event.preventDefault();
          event.stopPropagation();
          const picker = btn.closest('.of-color-picker');
          const fund = picker?.dataset.fund;
          const rawValue = btn.dataset.colorIndex ?? '';
          if (!fund) return;
          if (rawValue === '') delete S.pointColors[fund];
          else S.pointColors[fund] = parseInt(rawValue, 10);
          closeAllColorPalettes();
          fullRefresh();
        };
      });
      area.onclick = event => {
        if (!event.target.closest('.of-color-picker')) closeAllColorPalettes();
      };
      area.querySelectorAll('.of-cb[data-row-key]').forEach(cb => {
        cb.addEventListener('change', function() {
          const k = this.dataset.rowKey;
          if (this.checked) S.visibleKeys.add(k); else S.visibleKeys.delete(k);
          const row = area.querySelector(`tr[data-row-key="${CSS.escape(k)}"]`);
          if (row) row.classList.toggle('of-row-hidden', !this.checked);
          refreshScatter();
          const all = area.querySelectorAll('.of-cb[data-row-key]');
          const checked = [...all].filter(c=>c.checked).length;
          const cbAll = area.querySelector('#of-cb-all');
          if (cbAll) { cbAll.checked = checked===all.length; cbAll.indeterminate = checked>0&&checked<all.length; }
        });
      });
      const cbAll = area.querySelector('#of-cb-all');
      if (cbAll) cbAll.addEventListener('change', function() {
        if (this.checked) S.visibleKeys = new Set(displayRows.map(r=>r.key));
        else S.visibleKeys = new Set();
        fullRefresh();
      });
    }

    /* ── Wire events ── */
    area.querySelector('#of-period-btns').addEventListener('click', e => {
      const btn = e.target.closest('[data-of-period]');
      if (!btn) return;
      S.period = btn.dataset.ofPeriod;
      area.querySelectorAll('.of-period-btn').forEach(b=>b.classList.toggle('is-active',b.dataset.ofPeriod===S.period));
      fullRefresh();
    });
    area.querySelector('#of-select-x').addEventListener('change', e => { S.xKey=e.target.value; refreshScatter(); });
    area.querySelector('#of-select-y').addEventListener('change', e => { S.yKey=e.target.value; refreshScatter(); });
    area.querySelector('#of-reset-labels').addEventListener('click', () => {
      resetScatterLabelOffsets();
      refreshScatter();
    });
    area.querySelector('#of-download-svg')?.addEventListener('click', downloadOtherFactorsSvg);
    area.querySelectorAll('[data-of-mode]').forEach(btn => {
      btn.addEventListener('click', function() {
        S.mode = this.dataset.ofMode;
        area.querySelectorAll('[data-of-mode]').forEach(b=>b.classList.toggle('is-active',b.dataset.ofMode===S.mode));
        fullRefresh();
      });
    });

    syncScatterLabelBoxes();
    bindSideTableEvents();
    bindScatterLabelDrag();
    bindBottomTableEvents();
    App._currentExport = null;
    App._currentTableExport = null;
    App._currentClipboardExport = copyOtherFactorsClipboard;
    bindPageImageActions(area, 'report-card', 'master-other-factors');
  },

  /* ── Coming Soon placeholder ── */
  comingSoon(area) {
    area.innerHTML = `
      <div class="card report-card" id="report-card" style="min-height:300px;display:flex;align-items:center;justify-content:center">
        <div style="text-align:center;padding:60px 40px;max-width:480px">
          <div style="width:64px;height:64px;background:#e8f0fb;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 20px">
            <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#1a3c6e" stroke-width="1.8">
              <circle cx="12" cy="12" r="10"/>
              <polyline points="12 6 12 12 16 14"/>
            </svg>
          </div>
          <h3 style="font-size:1.1rem;font-weight:700;color:#1a2744;margin-bottom:10px">กำลังเตรียมหน้านี้</h3>
          <p style="font-size:0.88rem;color:#64748b;line-height:1.7;margin:0">
            เมนูนี้ถูกเปิดรอไว้แล้ว<br>
            และยังอยู่ระหว่างเตรียมข้อมูล / หน้าจอ เท่านั้น
          </p>
        </div>
      </div>`;
    App._currentExport = null;
    App._currentTableExport = null;
  },

  /* ── Variants of masterOtherFactors with different default metrics ── */
  async masterOtherFactorsVariant(area, stateKey, defaults) {
    // Temporarily set defaults for new visits, then delegate
    if (!State[stateKey]) {
      State[stateKey] = {
        mode: 'annualized',
        period: defaults.period || '3Y',
        xKey:   defaults.xKey  || 'maxdd',
        yKey:   defaults.yKey  || 'return',
        visibleKeys: null,
      };
    }
    // Patch State._of to point at variant state, run masterOtherFactors, restore
    const saved = State._of;
    State._of = State[stateKey];
    await Pages.masterOtherFactors(area);
    // After render, bind back the right store
    State._of = saved;
    // Re-point the live state key so events update the right store
    State._of = State[stateKey];
  },

  /* ── บันทึกข้อมูล (Drafts / Notes) ── */
  async notesPage(area) {
    const DRAFTS_KEY = 'avp-fund-drafts';
    let currentDrafts = [];
    let storageMode = 'local';
    let draftCategoryFilter = '';
    let draftTypeFilter = '';
    let draftQuarterFilter = '';
    let draftAuthorFilter = '';
    let overwriteDraftId = '';

    function loadLocalDrafts() {
      try { return JSON.parse(localStorage.getItem(DRAFTS_KEY) || '[]'); }
      catch { return []; }
    }

    function saveLocalDrafts(arr) {
      localStorage.setItem(DRAFTS_KEY, JSON.stringify(arr));
    }

    async function loadDrafts() {
      try {
        const res = await fetch('/api/drafts', { cache: 'no-store' });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        storageMode = 'file';
        const fileDrafts = Array.isArray(data.drafts) ? data.drafts : [];
        const localDrafts = loadLocalDrafts();
        const fileIds = new Set(fileDrafts.map(d => String(d.id || '')));
        const draftsToMigrate = localDrafts.filter(d => d?.id && !fileIds.has(String(d.id)));
        for (const draft of draftsToMigrate) {
          try {
            await fetch('/api/drafts', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                ...draft,
                highlights: draft.highlights || {},
                filters: draft.filters || {},
                migratedFromLocalStorage: true,
              }),
            });
          } catch {
            /* keep local copy if migration fails */
          }
        }
        if (draftsToMigrate.length) {
          localStorage.removeItem(DRAFTS_KEY);
          const migrated = await fetch('/api/drafts', { cache: 'no-store' });
          if (migrated.ok) {
            const migratedData = await migrated.json();
            return Array.isArray(migratedData.drafts) ? migratedData.drafts : fileDrafts;
          }
        }
        return fileDrafts;
      } catch {
        storageMode = 'local';
        return loadLocalDrafts();
      }
    }

    async function saveDraft(draft) {
      if (storageMode === 'file') {
        const res = await fetch('/api/drafts', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(draft),
        });
        if (!res.ok) throw new Error(`บันทึกไฟล์ไม่สำเร็จ (${res.status})`);
        return res.json();
      }
      const existingIdx = currentDrafts.findIndex(item => String(item?.id || '') === String(draft?.id || ''));
      if (existingIdx >= 0) currentDrafts.splice(existingIdx, 1);
      currentDrafts.unshift(draft);
      saveLocalDrafts(currentDrafts);
      return { ok: true, draft };
    }

    async function deleteDraft(draft, idx) {
      if (storageMode === 'file' && draft?.id) {
        const res = await fetch(`/api/drafts/${encodeURIComponent(draft.id)}`, { method: 'DELETE' });
        if (!res.ok) throw new Error(`ลบไฟล์ไม่สำเร็จ (${res.status})`);
        return;
      }
      currentDrafts.splice(idx, 1);
      saveLocalDrafts(currentDrafts);
    }

    function restoreDraft(d) {
      State.selectedFunds = d.selectedFunds || {};
      State.selectedKeys  = new Set(d.selectedKeys || []);
      State.highlights    = d.highlights || {};
      State.selectFundFilters = {
        ...State.selectFundFilters,
        ...(d.filters?.selectFundFilters || {}),
      };
      State.selectFundSort = {
        ...State.selectFundSort,
        ...(d.filters?.selectFundSort || {}),
      };
      if (d.currentQuarter) State.currentQuarter = d.currentQuarter;
    }

    function buildDraftName(category, type, dateText) {
      const parts = [
        category || 'ทุกหมวด AVP',
        type || 'ทุกประเภท',
        dateText || new Date().toISOString().slice(0, 10),
      ];
      return parts.filter(Boolean).join(' · ');
    }

    function getDraftSelectedFundsSnapshot() {
      const selected = {};
      [...State.selectedKeys].forEach(key => {
        if (State.selectedFunds?.[key]) selected[key] = State.selectedFunds[key];
      });
      return selected;
    }

    function draftQuarterOf(d) {
      return d.currentQuarter || d.dataQuarter || d.filters?.currentQuarter || d.quarter || '';
    }

    function buildDraftPayload(existingDraft = null) {
      const userDate = area.querySelector('#notes-date')?.value || '';
      const avpCategory = State.selectFundFilters.category || '';
      const fundType = State.selectFundFilters.type || '';
      const authorEmail = State.currentUser?.email || area.querySelector('#notes-author')?.value.trim() || '';
      const notes = area.querySelector('#notes-notes')?.value.trim() || '';
      const nowIso = new Date().toISOString();
      return {
        id: existingDraft?.id || Date.now().toString(),
        name: buildDraftName(avpCategory, fundType, userDate),
        avpCategory,
        fundType,
        userDate,
        author: authorEmail,
        authorEmail,
        authorName: State.currentUser?.name || existingDraft?.authorName || '',
        notes,
        createdAt: existingDraft?.createdAt || nowIso,
        selectedFunds: getDraftSelectedFundsSnapshot(),
        selectedKeys: [...State.selectedKeys],
        highlights: { ...State.highlights },
        filters: {
          selectFundFilters: { ...State.selectFundFilters },
          selectFundSort: { ...State.selectFundSort },
        },
        currentQuarter: State.currentQuarter || '',
      };
    }

    function renderPage() {
      const drafts = currentDrafts;
      const selectedCount = State.selectedKeys.size;
      const highlightedCount = Object.keys(State.highlights || {}).length;
      const currentCategory = State.selectFundFilters.category || 'ทั้งหมด';
      const currentType = State.selectFundFilters.type || 'ทั้งหมด';
      const currentAuthor = State.currentUser?.email || $('#user-email')?.textContent?.trim() || '';
      const draftCategories = [...new Set(drafts
        .map(d => d.avpCategory || d.filters?.selectFundFilters?.category || d.asset || '')
        .filter(Boolean))]
        .sort((a, b) => String(a).localeCompare(String(b), 'th'));
      const draftTypes = [...new Set(drafts
        .map(d => d.fundType || d.filters?.selectFundFilters?.type || '')
        .filter(Boolean))]
        .sort((a, b) => String(a).localeCompare(String(b), 'th'));
      const draftQuarters = [...new Set([
        ...(State.availableQuarters || []),
        ...drafts.map(draftQuarterOf),
      ].filter(Boolean))]
        .sort()
        .reverse();
      const draftAuthors = [...new Set(drafts
        .map(d => d.author || d.authorEmail || '')
        .filter(Boolean))]
        .sort((a, b) => String(a).localeCompare(String(b), 'th'));
      const shownDrafts = drafts.map((d, i) => ({ d, i })).filter(({ d }) => {
        const category = d.avpCategory || d.filters?.selectFundFilters?.category || d.asset || '';
        const type = d.fundType || d.filters?.selectFundFilters?.type || '';
        const quarter = draftQuarterOf(d);
        const author = d.author || d.authorEmail || '';
        if (draftCategoryFilter && category !== draftCategoryFilter) return false;
        if (draftTypeFilter && type !== draftTypeFilter) return false;
        if (draftQuarterFilter && quarter !== draftQuarterFilter) return false;
        if (draftAuthorFilter && author !== draftAuthorFilter) return false;
        return true;
      });
      const storageText = storageMode === 'file'
        ? 'บันทึกเป็นไฟล์กลางใน Drafts/'
        : 'บันทึกใน Browser นี้ชั่วคราว';

      const draftCards = shownDrafts.length === 0
        ? `<div class="notes-empty"><svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg><p>ยังไม่มีบันทึกที่บันทึกไว้</p></div>`
        : shownDrafts.map(({ d, i }) => {
            const fundCount = Object.keys(d.selectedFunds || {}).length;
            const hlCount = Object.keys(d.highlights || {}).length;
            const dateStr = d.createdAt ? new Date(d.createdAt).toLocaleDateString('th-TH', { year:'numeric', month:'short', day:'numeric' }) : '—';
            const categoryText = d.avpCategory || d.filters?.selectFundFilters?.category || d.asset || 'ทั้งหมด';
            const typeText = d.fundType || d.filters?.selectFundFilters?.type || 'ทั้งหมด';
            const quarterText = draftQuarterOf(d) || 'ไม่ระบุ';
            const isOverwriteTarget = overwriteDraftId && String(overwriteDraftId) === String(d.id || '');
            return `<div class="notes-card${isOverwriteTarget ? ' notes-card-overwrite-target' : ''}" data-idx="${i}">
              <div class="notes-card-icon"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg></div>
              <div class="notes-card-body">
                <div class="notes-card-title">${esc(d.name || 'ไม่มีชื่อ')}</div>
                <div class="notes-card-meta">
                  <span class="notes-tag notes-tag-asset">หมวด AVP: ${esc(categoryText || 'ทั้งหมด')}</span>
                  <span class="notes-tag notes-tag-type">ประเภท: ${esc(typeText || 'ทั้งหมด')}</span>
                  <span class="notes-tag notes-tag-quarter">ชุดข้อมูล: ${esc(quarterText)}</span>
                  <span class="notes-tag">วันที่ ${esc(d.userDate || dateStr)}</span>
                  ${d.author ? `<span class="notes-tag">โดย ${esc(d.author)}</span>` : ''}
                  <span class="notes-tag notes-tag-count">${fundCount} กองทุน</span>
                  ${hlCount ? `<span class="notes-tag notes-tag-highlight">${hlCount} สีไฮไลท์</span>` : ''}
                </div>
                ${d.notes ? `<div class="notes-card-desc">${esc(d.notes)}</div>` : ''}
                <div class="notes-card-saved">บันทึกเมื่อ ${dateStr}</div>
              </div>
              <div class="notes-card-actions">
                <button class="btn btn-primary notes-btn-load" data-idx="${i}">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                  โหลด
                </button>
                <button class="btn notes-btn-overwrite" data-idx="${i}" title="บันทึกทับรายการนี้">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 21h18"/><path d="M5 21V7l4-4h7l3 3v15"/><path d="M9 21v-6h6v6"/><path d="M9 3v4h6"/></svg>
                  บันทึกทับ
                </button>
                <button class="btn btn-ghost notes-btn-del" data-idx="${i}" title="ลบ">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>
                  ลบ
                </button>
              </div>
            </div>`;
          }).join('');

      area.innerHTML = `
        <div class="card report-card" id="report-card">
          <!-- ── Form ── -->
          <div class="notes-form-wrap">
            <div class="notes-form-head">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>
              <span>บันทึกดราฟใหม่</span>
              ${selectedCount > 0 ? `<span class="notes-fund-badge">${selectedCount} กองที่เลือกไว้</span>` : '<span class="notes-fund-badge notes-fund-badge-warn">ยังไม่ได้เลือกกองทุน</span>'}
            </div>
            <div class="notes-storage-mode">${esc(storageText)}${highlightedCount ? ` · มีไฮไลท์ ${highlightedCount} กอง` : ''}</div>
            <div class="notes-form-grid">
              <div class="notes-field">
                <label class="notes-label">หมวดหมู่ AVP ที่เลือกอยู่</label>
                <input class="notes-input notes-readonly" id="notes-category" type="text" value="${esc(currentCategory)}" readonly>
              </div>
              <div class="notes-field">
                <label class="notes-label">ประเภทกองทุนที่เลือกอยู่</label>
                <input class="notes-input notes-readonly" id="notes-type" type="text" value="${esc(currentType)}" readonly>
              </div>
              <div class="notes-field">
                <label class="notes-label">วันที่</label>
                <input class="notes-input" id="notes-date" type="date" value="${new Date().toISOString().slice(0,10)}">
              </div>
              <div class="notes-field">
                <label class="notes-label">บันทึกโดย</label>
                <input class="notes-input notes-readonly" id="notes-author" type="email" value="${esc(currentAuthor || 'ไม่พบอีเมลผู้ใช้งาน')}" readonly>
              </div>
              <div class="notes-field notes-field-full">
                <label class="notes-label">หมายเหตุ</label>
                <textarea class="notes-input notes-textarea" id="notes-notes" rows="2" placeholder="รายละเอียดเพิ่มเติม..."></textarea>
              </div>
            </div>
            <div class="notes-form-footer">
              <button class="btn btn-primary" id="notes-save-btn">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>
                บันทึกดราฟ
              </button>
              <span class="notes-save-hint">${overwriteDraftId ? `กำลังเลือกดราฟสำหรับบันทึกทับอยู่ 1 รายการ โดยข้อมูลปัจจุบันมี ${selectedCount} กองทุนที่เลือกไว้ พร้อมสีไฮไลท์และตัวกรองปัจจุบัน` : `ระบบจะบันทึก ${selectedCount} กองทุนที่เลือกไว้ พร้อมสีไฮไลท์และตัวกรองปัจจุบัน`}</span>
            </div>
          </div>

          <!-- ── Draft list ── -->
          <div class="notes-list-head">
            <div class="notes-list-title">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
              บันทึกที่บันทึกไว้ทั้งหมด (${shownDrafts.length}/${drafts.length})
            </div>
            <label class="notes-filter-wrap">Filter หมวดหมู่ AVP
              <select class="notes-filter-select" id="notes-category-filter">
                <option value="">ทั้งหมด</option>
                ${draftCategories.map(cat => `<option value="${esc(cat)}"${cat === draftCategoryFilter ? ' selected' : ''}>${esc(cat)}</option>`).join('')}
              </select>
            </label>
            <label class="notes-filter-wrap">Filter ประเภท
              <select class="notes-filter-select" id="notes-type-filter">
                <option value="">ทั้งหมด</option>
                ${draftTypes.map(type => `<option value="${esc(type)}"${type === draftTypeFilter ? ' selected' : ''}>${esc(type)}</option>`).join('')}
              </select>
            </label>
            <label class="notes-filter-wrap">Filter ชุดข้อมูล
              <select class="notes-filter-select" id="notes-quarter-filter">
                <option value="">ทั้งหมด</option>
                ${draftQuarters.map(q => `<option value="${esc(q)}"${q === draftQuarterFilter ? ' selected' : ''}>${esc(q)}</option>`).join('')}
              </select>
            </label>
            <label class="notes-filter-wrap">บันทึกโดย
              <select class="notes-filter-select" id="notes-author-filter">
                <option value="">ทั้งหมด</option>
                ${draftAuthors.map(author => `<option value="${esc(author)}"${author === draftAuthorFilter ? ' selected' : ''}>${esc(author)}</option>`).join('')}
              </select>
            </label>
          </div>
          <div class="notes-list" id="notes-list">${draftCards}</div>
        </div>

        <style>
          .notes-form-wrap{padding:20px 20px 16px;border-bottom:2px solid #e2e8f0}
          .notes-form-head{display:flex;align-items:center;gap:8px;font-size:1.05rem;font-weight:700;color:#1a2744;margin-bottom:14px}
          .notes-storage-mode{font-size:0.92rem;color:#64748b;background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:8px 11px;margin:-4px 0 14px}
          .notes-fund-badge{margin-left:auto;padding:4px 11px;border-radius:20px;font-size:0.9rem;font-weight:600;background:#d1fae5;color:#065f46}
          .notes-fund-badge-warn{background:#fef3c7;color:#92400e}
          .notes-form-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
          .notes-field{display:flex;flex-direction:column;gap:4px}
          .notes-field-full{grid-column:1/-1}
          .notes-label{font-size:0.92rem;font-weight:600;color:#475569}
          .notes-input{padding:8px 11px;border:1px solid #d0d9e8;border-radius:6px;font-size:1rem;font-family:inherit;color:#1e293b;transition:border .15s}
          .notes-readonly{background:#f8fafc;color:#475569}
          .notes-input:focus{outline:none;border-color:#1a3c6e;box-shadow:0 0 0 2px rgba(26,60,110,.1)}
          .notes-textarea{resize:vertical;min-height:72px}
          .notes-form-footer{display:flex;align-items:center;gap:12px;margin-top:14px}
          .notes-save-hint{font-size:0.92rem;color:#64748b}
          .notes-list-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 20px 8px;font-size:0.98rem;font-weight:700;color:#475569;border-bottom:1px solid #f1f5fb;flex-wrap:wrap}
          .notes-list-title{display:flex;align-items:center;gap:8px}
          .notes-filter-wrap{display:flex;align-items:center;gap:8px;font-size:0.92rem;color:#64748b;font-weight:600}
          .notes-filter-select{padding:6px 10px;border:1px solid #d0d9e8;border-radius:6px;background:#fff;color:#1e293b;font-family:inherit;font-size:0.96rem;min-width:180px}
          .notes-list{padding:14px 16px;display:flex;flex-direction:column;gap:12px}
          .notes-empty{padding:40px 20px;text-align:center;color:#94a3b8;display:flex;flex-direction:column;align-items:center;gap:10px;font-size:1rem}
          .notes-card{display:flex;align-items:flex-start;gap:14px;padding:14px 16px;background:#fff;border:1px solid #e2e8f0;border-radius:10px;transition:box-shadow .15s}
          .notes-card:hover{box-shadow:0 2px 12px rgba(26,60,110,.08)}
          .notes-card-overwrite-target{border-color:#f59e0b;box-shadow:0 0 0 2px rgba(245,158,11,.15)}
          .notes-card-icon{width:36px;height:36px;background:#e8f0fb;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0;color:#1a3c6e}
          .notes-card-body{flex:1;min-width:0}
          .notes-card-title{font-weight:700;font-size:1.02rem;color:#1e293b;margin-bottom:6px}
          .notes-card-meta{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:4px}
          .notes-tag{padding:3px 9px;border-radius:12px;font-size:0.88rem;font-weight:600;background:#f1f5fb;color:#475569}
          .notes-tag-asset{background:#dbeafe;color:#1d4ed8}
          .notes-tag-type{background:#ede9fe;color:#5b21b6}
          .notes-tag-quarter{background:#e0f2fe;color:#0369a1}
          .notes-tag-count{background:#d1fae5;color:#065f46}
          .notes-tag-highlight{background:#fef3c7;color:#92400e}
          .notes-card-desc{font-size:0.94rem;color:#64748b;margin-top:4px}
          .notes-card-saved{font-size:0.86rem;color:#94a3b8;margin-top:6px}
          .notes-card-actions{display:flex;flex-direction:column;gap:6px;flex-shrink:0}
          .notes-btn-load{font-size:1rem;padding:6px 12px;display:flex;align-items:center;gap:6px}
          .notes-btn-overwrite{font-size:1rem;padding:6px 12px;display:flex;align-items:center;gap:6px;border:1px solid #fcd34d;color:#b45309;background:#fffbeb}
          .notes-btn-overwrite:hover{background:#fef3c7}
          .notes-btn-del{font-size:1rem;color:#ef4444 !important;border-color:#fecaca;padding:6px 12px;display:flex;align-items:center;gap:6px}
          .notes-btn-del svg,.notes-btn-del span{color:#ef4444}
          .notes-btn-del:hover{background:#fef2f2}
          @media(max-width:768px){.notes-form-grid{grid-template-columns:1fr}.notes-filter-wrap{width:100%;justify-content:space-between}.notes-filter-select{min-width:0;flex:1}}
        </style>`;

      // ── Events ──
      area.querySelector('#notes-save-btn')?.addEventListener('click', async () => {
        const draft = buildDraftPayload();
        try {
          await saveDraft(draft);
          overwriteDraftId = '';
          currentDrafts = await loadDrafts();
          renderPage();
        } catch (err) {
          toast(`บันทึกไม่สำเร็จ: ${err.message || err}`, 'error', 5000);
        }
      });

      area.querySelector('#notes-category-filter')?.addEventListener('change', e => {
        draftCategoryFilter = e.target.value;
        renderPage();
      });

      area.querySelector('#notes-type-filter')?.addEventListener('change', e => {
        draftTypeFilter = e.target.value;
        renderPage();
      });

      area.querySelector('#notes-quarter-filter')?.addEventListener('change', e => {
        draftQuarterFilter = e.target.value;
        renderPage();
      });

      area.querySelector('#notes-author-filter')?.addEventListener('change', e => {
        draftAuthorFilter = e.target.value;
        renderPage();
      });

      area.querySelector('#notes-list')?.addEventListener('click', async e => {
        const loadBtn = e.target.closest('.notes-btn-load');
        const overwriteBtn = e.target.closest('.notes-btn-overwrite');
        const delBtn  = e.target.closest('.notes-btn-del');
        if (loadBtn) {
          const idx = parseInt(loadBtn.dataset.idx);
          const d = currentDrafts[idx];
          if (!d) return;
          restoreDraft(d);
          App.navigate('select-fund');
        }
        if (overwriteBtn) {
          const idx = parseInt(overwriteBtn.dataset.idx);
          const existingDraft = currentDrafts[idx];
          if (!existingDraft) return;
          if (!confirm(`บันทึกทับ "${existingDraft.name || 'ดราฟนี้'}" ?`)) return;
          try {
            overwriteDraftId = existingDraft.id || '';
            await saveDraft(buildDraftPayload(existingDraft));
            currentDrafts = await loadDrafts();
            renderPage();
            toast('บันทึกทับดราฟเรียบร้อยแล้ว', 'success');
          } catch (err) {
            toast(`บันทึกทับไม่สำเร็จ: ${err.message || err}`, 'error', 5000);
          }
        }
        if (delBtn) {
          if (!confirm('ลบบันทึกนี้?')) return;
          const idx = parseInt(delBtn.dataset.idx);
          try {
            await deleteDraft(currentDrafts[idx], idx);
            if (String(overwriteDraftId) === String(currentDrafts[idx]?.id || '')) overwriteDraftId = '';
            currentDrafts = await loadDrafts();
            renderPage();
          } catch (err) {
            toast(`ลบไม่สำเร็จ: ${err.message || err}`, 'error', 5000);
          }
        }
      });
    }

    setLoading(area, 'กำลังโหลดบันทึกข้อมูล...');
    currentDrafts = await loadDrafts();
    renderPage();
    App._currentExport = null;
    App._currentTableExport = null;
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
      const selectedCountText = selectedThaiCodes.length
        ? `${selectedThaiCodes.length} กองทุนที่พร้อมใช้เป็นจุดเริ่มต้น`
        : 'ยังไม่มีกองทุนที่เลือกจากเมนูเลือกกองทุน';
      const selectedChips = selectedThaiCodes.length
        ? selectedThaiCodes.map(code => `<span class="ft-viewer-chip">${esc(code)}</span>`).join('')
        : '<span class="ft-viewer-empty">ยังไม่ได้เลือกกองจากเมนูเลือกกองทุน ระบบจึงใส่ ISIN ตัวอย่างไว้ให้ก่อน</span>';

      area.innerHTML = `
        ${pageToolActions('master-placeholder-2', CONFIG.PAGES['master-placeholder-2']?.source || 'FT Fund Data Viewer')}
        <div id="report-card" class="card ft-viewer-card">
          <div class="ft-viewer-wrap">
            <section class="ft-viewer-hero">
              <div class="ft-viewer-hero-copy">
                <div class="ft-viewer-eyebrow">Master Fund Research</div>
                <h2>Top 10 Holding</h2>
                <p>ดึงข้อมูลจาก FT ผ่าน ISIN แล้วแสดงทุกส่วนที่ระบบอ่านได้ก่อน เพื่อช่วยดูภาพรวมให้ครบก่อนตัดสินใจว่าจะเก็บ field ไหนไว้ใช้งานจริง</p>
              </div>
              <div class="ft-viewer-hero-side">
                <div class="ft-viewer-source">API: FT Fund Data Viewer</div>
                <div class="ft-viewer-source-note">โหมดนี้ตั้งใจให้กดโหลดเอง เพื่อเลือกดูข้อมูลทีละกองอย่างแม่นยำ</div>
              </div>
            </section>

            <section class="ft-viewer-summary-grid">
              <div class="ft-viewer-summary-card">
                <span class="ft-viewer-summary-label">สถานะการทำงาน</span>
                <strong class="ft-viewer-summary-value">Manual Load</strong>
                <span class="ft-viewer-summary-note">ยังไม่ยิง API จนกว่าจะกดปุ่มดึงข้อมูล</span>
              </div>
              <div class="ft-viewer-summary-card">
                <span class="ft-viewer-summary-label">ISIN เริ่มต้น</span>
                <strong class="ft-viewer-summary-value ft-viewer-summary-code">${esc(suggestedIsin)}</strong>
                <span class="ft-viewer-summary-note">ใช้จากกองทุนที่เลือก หรือ fallback เป็นตัวอย่าง</span>
              </div>
              <div class="ft-viewer-summary-card">
                <span class="ft-viewer-summary-label">กองทุนที่เลือกอยู่</span>
                <strong class="ft-viewer-summary-value">${esc(selectedCountText)}</strong>
                <span class="ft-viewer-summary-note">ระบบแสดง chip เพื่อช่วยเช็กก่อนเริ่มโหลด</span>
              </div>
            </section>

            <section class="ft-viewer-panel ft-viewer-selection">
              <div class="ft-viewer-panel-head">
                <div>
                  <h3>กองทุนไทยที่เลือกอยู่</h3>
                  <p>ใช้ดูบริบทของกองที่กำลังตรวจสอบ ก่อนเลือก ISIN ที่จะดึงข้อมูลจริง</p>
                </div>
              </div>
              <div class="ft-viewer-chip-row">${selectedChips}</div>
            </section>

            <section class="ft-viewer-panel ft-viewer-controls">
              <div class="ft-viewer-panel-head">
                <div>
                  <h3>เครื่องมือดึงข้อมูล</h3>
                  <p>ระบบจะโหลดเมื่อคุณกดปุ่มเท่านั้น เพื่อช่วยลดการยิง request ซ้ำระหว่างไล่ดูกองทุนหลายตัว</p>
                </div>
              </div>
              <div class="ft-viewer-row">
                <div class="ft-viewer-input-wrap">
                  <label class="ft-viewer-input-label" for="ft-isin-input">ISIN</label>
                  <input id="ft-isin-input" type="text" placeholder="เช่น IE00BFRSYJ83" value="${esc(suggestedIsin)}" />
                </div>
                <button id="ft-load-btn" class="btn btn-primary ft-viewer-load-btn" type="button">ดึงข้อมูล</button>
              </div>

              <div class="ft-viewer-checklist-wrap">
                <div class="ft-viewer-selection-label">เลือกชุดข้อมูลที่ต้องการดึง</div>
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
            </section>

            <section class="ft-viewer-panel">
              <div class="ft-viewer-panel-head">
                <div>
                  <h3>สถานะการโหลด</h3>
                  <p>จะแสดงสถานะล่าสุดของ request และผลลัพธ์ที่ถูกโหลดเข้ามา</p>
                </div>
              </div>
              <div id="ft-status" class="ft-viewer-status">พร้อมให้กดดึงข้อมูล</div>
            </section>

            <section class="ft-viewer-panel">
              <div class="ft-viewer-panel-head">
                <div>
                  <h3>ผลลัพธ์ทั้งหมดจาก FT</h3>
                  <p>แสดงข้อมูลครบทุกส่วนที่โหลดได้ก่อน ทั้งตารางสรุป ตารางย่อย และ Raw JSON สำหรับใช้ตัดสินใจว่าจะเก็บ field ใดไว้ต่อ</p>
                </div>
              </div>
              <div id="ft-output" class="ft-viewer-output">
                <div class="ft-viewer-placeholder">
                  <div class="ft-viewer-placeholder-icon">FT</div>
                  <div>
                    <strong>ยังไม่มีข้อมูลที่โหลดเข้ามา</strong>
                    <p>เลือก ISIN และชุดข้อมูลที่ต้องการ แล้วกดปุ่มดึงข้อมูลเพื่อเริ่มแสดงผล</p>
                  </div>
                </div>
              </div>
            </section>
          </div>
        </div>`;

      const isinInput = $('#ft-isin-input', area);
      const loadBtn = $('#ft-load-btn', area);
      const statusEl = $('#ft-status', area);
      const outputEl = $('#ft-output', area);
      let latestExportPayload = null;

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
          statusEl.textContent = 'กรุณาใส่ ISIN ก่อนกดดึงข้อมูล';
          outputEl.innerHTML = '';
          latestExportPayload = null;
          return;
        }

        const url = `${TOP_10_HOLDING_API_URL}?isin=${encodeURIComponent(isin)}&fields=${encodeURIComponent(fields.join(','))}`;
        statusEl.textContent = `กำลังโหลดข้อมูลสำหรับ ${isin}...`;
        outputEl.innerHTML = '<div class="state-box"><div class="spinner"></div><span>กำลังดึงข้อมูลจาก FT และจัดรูปแบบผลลัพธ์...</span></div>';
        loadBtn.disabled = true;
        loadBtn.innerHTML = '<span class="spin-sm"></span> กำลังโหลด';

        try {
          const res = await fetch(url);
          const data = await res.json();
          if (!data.ok) {
            statusEl.textContent = data.error || 'เกิดข้อผิดพลาด';
            outputEl.innerHTML = '<p class="ft-viewer-empty">ไม่สามารถโหลดข้อมูลได้ กรุณาตรวจสอบ ISIN หรือเลือก field ใหม่แล้วลองอีกครั้ง</p>';
            latestExportPayload = null;
            return;
          }

          statusEl.textContent = `โหลดสำเร็จ: ${data.isin || isin} • ${fields.length || 0} fields`;
          renderData(data);
          latestExportPayload = buildTop10SingleExportPayload(
            CONFIG.PAGES['master-placeholder-2']?.title || 'Top 10 Holding',
            CONFIG.PAGES['master-placeholder-2']?.source || 'FT Fund Data Viewer',
            data.isin || isin,
            data,
          );
        } catch (err) {
          statusEl.textContent = `เกิดข้อผิดพลาด: ${err.message}`;
          outputEl.innerHTML = '<p class="ft-viewer-empty">เกิดข้อผิดพลาดระหว่างเชื่อมต่อกับ API</p>';
          latestExportPayload = null;
        } finally {
          loadBtn.disabled = false;
          loadBtn.textContent = 'ดึงข้อมูล';
        }
      };

      loadBtn?.addEventListener('click', loadData);
      isinInput?.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') loadData();
      });

      App._currentExport = null;
      App._currentTableExport = () => latestExportPayload;
      bindPageImageActions(area, 'report-card', 'master-top10-holding');
    } catch (e) {
      setError(area, e.message, 'master-placeholder-2');
    }
  },

  async masterMenu02V2(area) {
    setLoading(area, 'กำลังเตรียมหน้า Top 10 Holding V2...');

    try {
      const universe = await buildSelectedMasterUniverse();
      const compareCandidates = universe
        .map(({ fund, master }) => ({
          label: String(fund.code || master.isin || '').trim(),
          isin: String(master.isin || fund.masterId || '').trim(),
          masterName: String(master.name || '').trim(),
        }))
        .filter(item => item.label && item.isin);
      const defaultTokens = compareCandidates.length
        ? compareCandidates.slice(0, Math.min(compareCandidates.length, 3)).map(item => item.label)
        : ['IE00BFRSYJ83'];
      const selectedChips = compareCandidates.length
        ? compareCandidates.slice(0, 12).map(item => `<span class="thv2-chip">${esc(item.label)}</span>`).join('')
        : '<span class="ft-viewer-empty">ยังไม่ได้เลือกกองจากเมนูเลือกกองทุน สามารถพิมพ์ ISIN ตรง ๆ ได้</span>';

      area.innerHTML = `
        ${pageToolActions('master-placeholder-7', CONFIG.PAGES['master-placeholder-7']?.source || 'FT Fund Data Viewer')}
        <div id="report-card" class="card thv2-card">
          <div class="thv2-wrap">
            <section class="thv2-hero">
              <div class="thv2-hero-copy">
                <div class="thv2-eyebrow">Master Fund Research</div>
                <h2>Top 10 Holding V2</h2>
                <p>เปรียบเทียบหลายกองในมุมที่ต้องใช้ตัดสินใจจริง โดยแสดงเฉพาะ Fees, Manager, Performance, Allocation, Risk และ Top 10 Holding ใน layout เดียว</p>
              </div>
              <div class="thv2-badge">Manual Compare Load</div>
            </section>

            <section class="thv2-panel">
              <div class="thv2-panel-head">
                <div>
                  <h3>กองทุนที่พร้อมใช้</h3>
                  <p>ระบบจะใช้กองที่เลือกไว้เป็นตัวช่วย map ชื่อกองไทยกับ ISIN โดยคุณสามารถพิมพ์เป็นชื่อกองไทยหรือ ISIN ก็ได้</p>
                </div>
              </div>
              <div class="thv2-chip-row">${selectedChips}</div>
            </section>

            <section class="thv2-panel">
              <div class="thv2-panel-head">
                <div>
                  <h3>เลือกกองเพื่อเปรียบเทียบ</h3>
                  <p>พิมพ์ชื่อกองไทยหรือ ISIN หลายรายการ คั่นด้วย comma หรือขึ้นบรรทัดใหม่ แล้วกดดึงข้อมูล</p>
                </div>
              </div>
              <div class="thv2-control-row">
                <div class="thv2-input-wrap">
                  <label class="thv2-input-label" for="thv2-input">กองทุน / ISIN</label>
                  <textarea id="thv2-input" class="thv2-textarea" placeholder="เช่น KF-SINCOME, KFDIVERS-I หรือ IE00BFRSYJ83">${esc(defaultTokens.join(', '))}</textarea>
                </div>
                <button id="thv2-load-btn" class="btn btn-primary thv2-load-btn" type="button">ดึงข้อมูล</button>
              </div>
            </section>

            <section class="thv2-panel">
              <div class="thv2-panel-head">
                <div>
                  <h3>สถานะการโหลด</h3>
                  <p>โหลดทีเดียวหลายกอง และสรุปผลที่เปรียบเทียบได้ในหน้าเดียว</p>
                </div>
              </div>
              <div id="thv2-status" class="thv2-status">พร้อมให้กดดึงข้อมูล</div>
            </section>

            <section class="thv2-panel">
              <div class="thv2-panel-head">
                <div>
                  <h3>ผลลัพธ์เปรียบเทียบ</h3>
                  <p>แสดงเฉพาะ field ที่ต้องใช้ พร้อมเลย์เอาต์เปรียบเทียบหลายกองแบบอ่านเร็ว</p>
                </div>
              </div>
              <div id="thv2-output" class="thv2-output">
                <div class="thv2-placeholder">
                  <div class="thv2-placeholder-icon">V2</div>
                  <div>
                    <strong>ยังไม่มีข้อมูลเปรียบเทียบ</strong>
                    <p>กรอกกองทุนด้านบนแล้วกดดึงข้อมูล ระบบจะสร้าง compare board ให้ทันที</p>
                  </div>
                </div>
              </div>
            </section>
          </div>
        </div>`;

      const inputEl = $('#thv2-input', area);
      const loadBtn = $('#thv2-load-btn', area);
      const statusEl = $('#thv2-status', area);
      const outputEl = $('#thv2-output', area);
      let latestExportPayload = null;

      const compareLookup = new Map();
      compareCandidates.forEach(item => {
        compareLookup.set(String(item.label || '').trim().toUpperCase(), item);
        compareLookup.set(String(item.isin || '').trim().toUpperCase(), item);
      });
      const requestedFields = ['fees', 'manager', 'performance', 'allocation', 'risk', 'holdings'];

      const parseTokens = (value) => String(value || '')
        .split(/[\n,]+/)
        .map(token => token.trim())
        .filter(Boolean);

      const resolveEntries = (tokens) => {
        const used = new Set();
        return tokens.map(token => {
          const raw = String(token || '').trim();
          const key = raw.toUpperCase();
          const matched = compareLookup.get(key);
          const resolved = matched ? {
            label: matched.label,
            isin: matched.isin,
            masterName: matched.masterName,
          } : {
            label: raw,
            isin: raw,
            masterName: '',
          };
          const dedupeKey = `${resolved.label}__${resolved.isin}`.toUpperCase();
          if (used.has(dedupeKey)) return null;
          used.add(dedupeKey);
          return resolved;
        }).filter(Boolean);
      };

      const fieldRows = {
        fees: [
          { key: 'ongoingCharge', label: 'ongoingCharge' },
          { key: 'initialCharge', label: 'initialCharge' },
          { key: 'maxAnnualCharge', label: 'maxAnnualCharge' },
          { key: 'exitCharge', label: 'exitCharge' },
        ],
        manager: [
          { key: 'name', label: 'name' },
          { key: 'startDate', label: 'startDate' },
        ],
        performance: [
          { key: '1M', label: '1M' },
          { key: '3M', label: '3M' },
          { key: '6M', label: '6M' },
          { key: '1Y', label: '1Y' },
          { key: '3Y', label: '3Y' },
          { key: '5Y', label: '5Y' },
        ],
      };

      const formatValue = (value) => {
        const text = String(value ?? '').trim();
        return text || '—';
      };

      const renderMetricCard = (title, entry, rows) => `
        <article class="thv2-compare-card">
          <div class="thv2-card-head">
            <div>
              <h4>${esc(entry.label)}</h4>
              <p>${esc(entry.masterName || entry.isin)}</p>
            </div>
          </div>
          <div class="thv2-kv-list">
            ${rows.map(row => `
              <div class="thv2-kv-row">
                <span class="thv2-kv-key">${esc(row.label)}</span>
                <span class="thv2-kv-value">${esc(formatValue(entry.data?.[title]?.[row.key]))}</span>
              </div>
            `).join('')}
          </div>
        </article>`;

      const renderDistributionCard = (title, entry, items) => `
        <article class="thv2-compare-card">
          <div class="thv2-card-head">
            <div>
              <h4>${esc(entry.label)}</h4>
              <p>${esc(entry.masterName || entry.isin)}</p>
            </div>
          </div>
          <div class="thv2-list-stack">
            ${Array.isArray(items) && items.length ? items.map(item => `
              <div class="thv2-list-row">
                <span class="thv2-list-name">${esc(formatValue(item.name || item.companyName))}</span>
                <strong class="thv2-list-value">${esc(formatValue(item.text || item.weightText || (item.percent ?? item.weightPercent)))}</strong>
              </div>
            `).join('') : '<p class="ft-viewer-empty">ไม่มีข้อมูล</p>'}
          </div>
        </article>`;

      const renderRiskCard = (entry) => {
        const rows = Object.entries(entry.data?.risk || {});
        return `
          <article class="thv2-compare-card">
            <div class="thv2-card-head">
              <div>
                <h4>${esc(entry.label)}</h4>
                <p>${esc(entry.masterName || entry.isin)}</p>
              </div>
            </div>
            <div class="thv2-kv-list">
              ${rows.length ? rows.map(([key, value]) => `
                <div class="thv2-kv-row">
                  <span class="thv2-kv-key">${esc(key)}</span>
                  <span class="thv2-kv-value">${esc(formatValue(value))}</span>
                </div>
              `).join('') : '<p class="ft-viewer-empty">ไม่มีข้อมูล</p>'}
            </div>
          </article>`;
      };

      const sectionWrap = (title, note, cards) => `
        <section class="thv2-section">
          <div class="thv2-section-label">
            <h3>${esc(title)}</h3>
            <p>${esc(note)}</p>
          </div>
          <div class="thv2-cards-scroll">
            <div class="thv2-cards-track">
              ${cards.join('')}
            </div>
          </div>
        </section>`;

      const renderCompareBoard = (entries) => {
        const sections = [
          sectionWrap('Fees', 'แสดง 4 ค่าใช้จ่ายหลักของแต่ละกอง', entries.map(entry => renderMetricCard('fees', entry, fieldRows.fees))),
          sectionWrap('Manager', 'ชื่อผู้จัดการและวันที่เริ่มต้น', entries.map(entry => renderMetricCard('manager', entry, fieldRows.manager))),
          sectionWrap('Performance', 'ผลตอบแทนย้อนหลังตามช่วงเวลา', entries.map(entry => renderMetricCard('performance', entry, fieldRows.performance))),
          sectionWrap('Allocation: Sector', 'เปรียบเทียบสัดส่วนราย sector', entries.map(entry => renderDistributionCard('sector', entry, entry.data?.allocation?.sector))),
          sectionWrap('Allocation: Region', 'เปรียบเทียบสัดส่วนรายภูมิภาค', entries.map(entry => renderDistributionCard('region', entry, entry.data?.allocation?.region))),
          sectionWrap('Risk', 'ข้อมูล risk ที่ FT ส่งกลับมา', entries.map(entry => renderRiskCard(entry))),
          sectionWrap('Top 10 Holding', '10 อันดับหลักทรัพย์ที่ถืออยู่มากที่สุด', entries.map(entry => renderDistributionCard('holdings', entry, entry.data?.holdings))),
        ];
        outputEl.innerHTML = sections.join('');
      };

      const loadData = async () => {
        const tokens = parseTokens(inputEl?.value || '');
        const entries = resolveEntries(tokens);

        if (!entries.length) {
          statusEl.textContent = 'กรุณากรอกกองทุนหรือ ISIN อย่างน้อย 1 รายการ';
          outputEl.innerHTML = '<p class="ft-viewer-empty">ยังไม่มีรายการสำหรับเปรียบเทียบ</p>';
          latestExportPayload = null;
          return;
        }

        statusEl.textContent = `กำลังโหลดข้อมูล ${entries.length} กอง...`;
        outputEl.innerHTML = '<div class="state-box"><div class="spinner"></div><span>กำลังดึงข้อมูลและจัดวาง compare board...</span></div>';
        loadBtn.disabled = true;
        loadBtn.innerHTML = '<span class="spin-sm"></span> กำลังโหลด';

        try {
          const results = await Promise.all(entries.map(async (entry) => {
            const url = `${TOP_10_HOLDING_API_URL}?isin=${encodeURIComponent(entry.isin)}&fields=${encodeURIComponent(requestedFields.join(','))}`;
            const res = await fetch(url);
            const data = await res.json();
            return { ...entry, data, ok: !!data?.ok };
          }));

          const okResults = results.filter(item => item.ok);
          const failed = results.filter(item => !item.ok);

          if (!okResults.length) {
            statusEl.textContent = failed[0]?.data?.error || 'ไม่สามารถโหลดข้อมูลได้';
            outputEl.innerHTML = '<p class="ft-viewer-empty">ยังไม่พบข้อมูลที่ใช้เปรียบเทียบได้จากรายการที่กรอก</p>';
            latestExportPayload = null;
            return;
          }

          renderCompareBoard(okResults);
          latestExportPayload = buildTop10CompareExportPayload(
            CONFIG.PAGES['master-placeholder-7']?.title || 'Top 10 Holding V2',
            CONFIG.PAGES['master-placeholder-7']?.source || 'FT Fund Data Viewer',
            okResults,
          );
          statusEl.textContent = failed.length
            ? `โหลดสำเร็จ ${okResults.length} กอง และมี ${failed.length} กองที่โหลดไม่สำเร็จ`
            : `โหลดสำเร็จ ${okResults.length} กอง`;
        } catch (err) {
          statusEl.textContent = `เกิดข้อผิดพลาด: ${err.message}`;
          outputEl.innerHTML = '<p class="ft-viewer-empty">เกิดข้อผิดพลาดระหว่างเชื่อมต่อกับ API</p>';
          latestExportPayload = null;
        } finally {
          loadBtn.disabled = false;
          loadBtn.textContent = 'ดึงข้อมูล';
        }
      };

      loadBtn?.addEventListener('click', loadData);
      inputEl?.addEventListener('keydown', (event) => {
        if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') loadData();
      });

      App._currentExport = null;
      App._currentTableExport = () => latestExportPayload;
      bindPageImageActions(area, 'report-card', 'master-top10-holding-v2');
    } catch (e) {
      setError(area, e.message, 'master-placeholder-7');
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
      bindPageImageActions(area, 'report-card', 'master-menu-03');
    } catch (e) {
      setError(area, e.message, 'master-placeholder-3');
    }
  },

  /* ── GUIDE ── */
  async guide(area) {
    // แสดง loading ก่อน
    area.innerHTML = `
      <div class="card guide-wrap">
        <div class="card-header">
          <span class="card-title">📖 คู่มือการใช้งาน</span>
          <span class="badge badge-primary">v2.0</span>
        </div>
        <div class="card-body" id="guide-body">
          <div class="loading-wrap"><span class="spin"></span> กำลังโหลดคู่มือ...</div>
        </div>
      </div>`;

    const body = $('#guide-body', area);

    try {
      // ดึง README.md จาก server
      const resp = await fetch('README.md', { cache: 'no-cache' });
      if (!resp.ok) throw new Error(`โหลด README.md ไม่สำเร็จ (${resp.status})`);
      const md = await resp.text();

      // render markdown → HTML ด้วย marked.js
      if (typeof marked === 'undefined') throw new Error('marked.js ยังไม่โหลด');
      marked.setOptions({ breaks: true, gfm: true });
      body.innerHTML = `<div class="guide-readme">${marked.parse(md)}</div>`;

    } catch (err) {
      // fallback: แสดง error พร้อมลิงก์เปิด README โดยตรง
      body.innerHTML = `
        <div class="guide-callout" style="color:var(--danger)">
          ⚠️ โหลดคู่มือไม่สำเร็จ: ${esc(err.message)}
        </div>
        <p style="margin-top:12px">
          <a href="README.md" target="_blank" style="color:var(--primary);font-weight:700">
            เปิด README.md โดยตรง →
          </a>
        </p>`;
    }

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

  /* ─────────────────────────────────────────────────────────
     TOP 10 HOLDING V3  –  Multi-Fund Comparison Dashboard
     ───────────────────────────────────────────────────────── */
  async masterMenu02V3(area) {
    const ISHARE_TICKERS = ['ACWI','AAXJ','IXJ','IXN','IGF','IVV','CNYA','MCHI','INDA','IEV','AGG','REET'];
    const fundColors  = ['#1a3c6e','#e8a317','#2d9e6b','#d63b3b','#4a90d9','#8b5cf6','#f59e0b','#10b981'];
    const holdingsPalette = ['#1a3c6e','#2d9e6b','#8b5cf6','#e8a317','#4a90d9','#d63b3b','#5bb98c','#f3a93b','#6d8fd8','#c84a42','#7c5ce0','#5ca475'];

    setLoading(area, 'กำลังเตรียม Multi-Fund Compare Dashboard...');

    // ── โหลด Chart.js on-demand ──
    try {
      await new Promise((resolve, reject) => {
        if (typeof Chart !== 'undefined') { resolve(); return; }
        const s   = document.createElement('script');
        s.src     = 'https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js';
        s.onload  = resolve;
        s.onerror = () => reject(new Error('โหลด Chart.js ไม่สำเร็จ'));
        document.head.appendChild(s);
      });
    } catch (e) {
      setError(area, e.message, 'master-placeholder-8');
      return;
    }

    // ── ดึงรายชื่อกองทุนไทยจาก universe ──
    let thaiCodes = [];
    let thaiLookup = {}; // code -> { name, isin }
    try {
      const universe = await buildSelectedMasterUniverse();
      thaiCodes = universe
        .map(({ fund, master }) => ({
          code: String(fund.code || '').trim(),
          name: String(fund.name || master?.name || '').trim(),
          isin: String(master?.isin || fund.masterId || '').trim(),
        }))
        .filter(item => item.code);
      thaiCodes.forEach(item => { thaiLookup[item.code] = item; });
    } catch (_) {}

    const savedV3State = State.top10HoldingV3 || {};
    const saveV3State = (patch) => {
      State.top10HoldingV3 = {
        ...(State.top10HoldingV3 || {}),
        ...patch,
      };
    };

    // ── Demo data builder ──
    const rnd = (mn, mx, d = 2) => parseFloat((Math.random() * (mx - mn) + mn).toFixed(d));
    const buildDemoData = (names) => names.map((name, i) => {
      const yearlyDD = [rnd(-45,-20), rnd(-15,-5), rnd(-35,-15), rnd(-20,-8), rnd(-12,-3), rnd(-8,-2)];
      const perf = Array.from({ length: 12 }, (_, j) => j === 0 ? 0 : rnd(-3, 5))
        .reduce((acc, v) => { acc.push(parseFloat(((acc.at(-1) || 0) + v).toFixed(2))); return acc; }, []);
      return {
        name: name || '–', color: fundColors[i],
        nav: rnd(10, 100, 4), ytd: rnd(-10, 30), risk: Math.floor(Math.random() * 8 + 1),
        fee: rnd(0.5, 2.5), sharpe: rnd(0.2, 1.8), sd: rnd(5, 25),
        dividend: Math.random() > 0.6 ? 'มี' : 'ไม่มี',
        yearlyDD, perf,
        avgDD: parseFloat((yearlyDD.reduce((a, b) => a + b, 0) / yearlyDD.length).toFixed(2)),
        country: [],
      };
    });

    // ── Helper: get current selections ──
    const getSelections = () => {
      const thai = [];
      const iShareEl = area.querySelector('#v3-sel-0');
      for (let i = 1; i <= 7; i++) {
        const el = area.querySelector(`#v3-sel-${i}`);
        thai.push(el ? el.value.trim() : '');
      }
      return {
        iShare: iShareEl ? iShareEl.value.trim() : '',
        thai,
      };
    };

    // ── datalist options for Thai funds ──
    const thaiOptions = thaiCodes.length
      ? thaiCodes.map(item => `<option value="${esc(item.code)}">${esc(item.code)}${item.name ? ' – ' + item.name : ''}</option>`).join('')
      : '<option value="">ยังไม่มีข้อมูล</option>';

    // ── iShare dropdown options ──
    // ── Default selections from existing universe ──
    const defaultThaiCodes = Array.isArray(savedV3State?.selections?.thai) && savedV3State.selections.thai.length
      ? savedV3State.selections.thai.slice(0, 7)
      : thaiCodes.slice(0, 7).map(item => item.code);
    while (defaultThaiCodes.length < 7) defaultThaiCodes.push('');

    // ── Selector Card HTML builder ──
    const selectorCard = (idx, color, label, inputHtml) => `
      <div style="background:#fff;border-radius:12px;border:1px solid var(--border);border-top:3px solid ${color};padding:12px 14px;display:flex;flex-direction:column;gap:6px;">
        <div style="font-size:0.78rem;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.04em;">${label}</div>
        ${inputHtml}
        <div id="v3-card-info-${idx}" style="font-size:0.84rem;color:var(--text-muted);min-height:18px;"></div>
      </div>`;

    const iShareCard = selectorCard(0, fundColors[0], 'iShare Index',
      `<select id="v3-sel-0" style="width:100%;padding:8px 10px;border-radius:6px;border:1px solid var(--border);font-size:0.92rem;font-weight:600;color:var(--primary-dark);background:#fff;cursor:pointer;">
        <option value="">-- เลือก iShare --</option>
        ${ISHARE_TICKERS.map(t => `<option value="${t}"${t === (savedV3State?.selections?.iShare || '') ? ' selected' : ''}>${t}</option>`).join('')}
      </select>`
    );

    const thaiCards = defaultThaiCodes.map((def, i) => selectorCard(
      i + 1, fundColors[i + 1], `กองทุนไทย #${i + 1}`,
      `<div style="position:relative;">
        <input id="v3-sel-${i + 1}" list="v3-thai-datalist" value="${esc(def)}"
          placeholder="${thaiCodes.length ? 'เลือกหรือพิมพ์ชื่อกองทุน...' : 'พิมพ์รหัสกองทุน...'}"
          autocomplete="off"
          style="width:100%;padding:8px 10px;border-radius:6px;border:1px solid var(--border);font-size:0.92rem;font-weight:600;color:var(--primary-dark);background:#fff;box-sizing:border-box;cursor:pointer;" />
      </div>`
    )).join('');

    // ── Main HTML ──
    area.innerHTML = `
      <datalist id="v3-thai-datalist">${thaiOptions}</datalist>

      ${pageToolActions('master-placeholder-8', CONFIG.PAGES['master-placeholder-8']?.source || 'Multi-Fund Compare API')}

      <div class="card thv2-card" id="report-card">
        <div class="thv2-wrap">

          <!-- Fund Selectors (2×4 grid) -->
          <section class="thv2-panel">
            <div class="thv2-panel-head">
              <div>
                <h3>เลือกกองทุนที่ต้องการเปรียบเทียบ</h3>
                <p>ช่องที่ 1: iShare Index (Dropdown) &nbsp;|&nbsp; ช่องที่ 2-8: กองทุนไทย (พิมพ์หรือเลือกจาก Dropdown)</p>
              </div>
              <div style="display:flex;flex-direction:column;align-items:flex-end;gap:8px;">
                <button id="v3-load-btn" class="btn btn-primary" type="button" style="white-space:nowrap;">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
                  โหลดข้อมูลเปรียบเทียบ
                </button>
                <div id="v3-api-status-wrap" style="display:none;">
                  <div id="v3-api-status" style="padding:7px 14px;border-radius:8px;background:var(--primary-faint);color:var(--primary);font-size:0.92rem;font-weight:600;white-space:nowrap;"></div>
                </div>
              </div>
            </div>
            <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;">
              ${iShareCard}
              ${thaiCards}
            </div>
          </section>

          <div id="v3-presentation-card" style="display:flex;flex-direction:column;gap:18px;">
            <!-- 1. Master Fund Info Table -->
            <section class="thv2-panel" id="v3-master-info-section" style="display:none;">
              <div class="thv2-panel-head">
                <div>
                  <h3>ข้อมูล Master Fund</h3>
                  <p>แสดง ISIN และค่าธรรมเนียมหลักของแต่ละกองทุนที่เลือก</p>
                </div>
              </div>
              <div style="overflow-x:auto;">
                <table style="width:100%;border-collapse:collapse;font-size:0.92rem;">
                  <thead id="v3-master-thead"></thead>
                  <tbody id="v3-master-tbody"></tbody>
                </table>
              </div>
            </section>

            <!-- 2. Top 10 Holdings Comparison -->
            <section class="thv2-panel" id="v3-holdings-section" style="display:none;">
              <div class="thv2-panel-head">
                <div>
                  <h3>Top 10 Holdings เปรียบเทียบ</h3>
                  <p>การถือครองหลัก 10 อันดับแรกของแต่ละกองทุน (% น้ำหนัก)</p>
                </div>
              </div>
              <div style="overflow-x:auto;">
                <table style="width:100%;border-collapse:collapse;font-size:0.92rem;">
                  <thead id="v3-holdings-thead"></thead>
                  <tbody id="v3-holdings-tbody"></tbody>
                </table>
              </div>
            </section>

            <section class="thv2-panel" id="v3-holdings-structure-section" style="display:none;">
              <div class="thv2-panel-head">
                <div>
                  <h3>โครงสร้างน้ำหนัก Top 10 Holdings รายกองทุน</h3>
                  <p>กราฟเปรียบเทียบชื่อสินทรัพย์จริงใน Top 10 ของแต่ละกองทุน</p>
                </div>
              </div>
              <div style="position:relative;height:420px;"><canvas id="v3-top10-asset-chart"></canvas></div>
              <div id="v3-top10-asset-legend" style="margin-top:10px;display:flex;flex-wrap:wrap;gap:6px 14px;border-top:1px solid var(--border-light);padding-top:10px;"></div>
            </section>
          </div>

          <!-- 6. API Data Explorer -->
          <section class="thv2-panel" id="v3-explorer-section" style="display:none;">
            <div class="thv2-panel-head">
              <div><h3>API Data Explorer</h3><p>ข้อมูลดิบจาก Multi-Fund Compare API – รอกำหนด field mapping สำหรับ chart แต่ละส่วน</p></div>
              <button id="v3-toggle-raw" class="btn btn-ghost btn-sm">แสดง / ซ่อน JSON</button>
            </div>
            <div id="v3-raw" style="display:none;background:#1e293b;color:#94a3b8;border-radius:8px;padding:16px;font-family:monospace;font-size:0.84rem;max-height:440px;overflow:auto;white-space:pre-wrap;word-break:break-all;"></div>
          </section>

        </div>
      </div>`;

    for (let si = 1; si <= 7; si++) {
      const inp = area.querySelector(`#v3-sel-${si}`);
      if (!inp) continue;
      inp.addEventListener('change', () => { inp._v3prev = inp.value; });
      inp.addEventListener('input', () => { saveV3State({ selections: getSelections() }); });
      inp.addEventListener('blur', () => {
        setTimeout(() => {
          if (!inp.value && inp._v3prev !== undefined) {
            inp.value = inp._v3prev;
            inp._v3prev = undefined;
          }
          saveV3State({ selections: getSelections() });
        }, 150);
      });
    }
    area.querySelector('#v3-sel-0')?.addEventListener('change', () => { saveV3State({ selections: getSelections() }); });

    // ── Chart registry ──
    const chartReg = {};
    const mkChart = (id, config) => {
      if (chartReg[id]) { chartReg[id].destroy(); }
      const cv = document.getElementById(id);
      if (!cv) return null;
      return (chartReg[id] = new Chart(cv, config));
    };

    const common = {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false } },
    };
    let latestExportPayload = null;

      // ── Render dashboard ──
      const render = (data) => {
      // Show all sections
      ['v3-master-info-section','v3-holdings-section','v3-holdings-structure-section'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = '';
      });

      // Master Fund Info Table — transposed (funds = columns, fields = rows)
      const masterThead = document.getElementById('v3-master-thead');
      const masterTbody = document.getElementById('v3-master-tbody');
      const dash = '<span style="color:#cbd5e1;">—</span>';
      const fmtPct = v => { if (v == null) return dash; const n = parseFloat(String(v).replace(/%/g,'')); return isNaN(n) ? dash : n.toFixed(2) + '%'; };
      const fmtStr = v => v ? esc(v) : dash;
      const fmtIsin = v => v ? `<span style="font-family:monospace;font-size:0.86rem;color:#1a3c6e;">${esc(v)}</span>` : dash;

      if (masterThead) masterThead.innerHTML = `
        <tr style="background:#1a3c6e;color:#fff;font-size:0.86rem;">
          <th style="padding:10px 14px;min-width:130px;border-right:1px solid rgba(255,255,255,0.15);background:transparent;"></th>
          ${data.map(f => `<th style="padding:10px 10px;text-align:center;font-weight:700;min-width:90px;">
            <span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${f.color};margin-right:5px;vertical-align:middle;opacity:0.9;"></span>${esc((f._selectorName || f.name).split('–')[0].trim())}
          </th>`).join('')}
        </tr>`;

      const masterRows = [
        { label: 'ISIN Code',        fmt: f => fmtIsin(f._isin) },
        { label: 'Ongoing Charge',   fmt: f => fmtPct(f._feesRaw) },
        { label: 'Initial Charge',   fmt: f => fmtPct(f._feesInitial) },
        { label: 'Max Annual Charge',fmt: f => fmtPct(f._feesMaxAnnual) },
      ];
      if (masterTbody) masterTbody.innerHTML = masterRows.map((row, ri) => `
        <tr style="border-bottom:1px solid var(--border-light);${ri % 2 === 0 ? '' : 'background:#fafbfd;'}">
          <td style="padding:9px 14px;font-weight:700;font-size:0.9rem;color:var(--text-muted);background:#f8fafc;white-space:nowrap;border-right:1px solid var(--border-light);">${row.label}</td>
          ${data.map(f => `<td style="padding:9px 10px;text-align:center;font-size:0.9rem;color:var(--text);">${row.fmt(f)}</td>`).join('')}
        </tr>`).join('');

      // Top 10 Holdings comparison table
      const holdThead = document.getElementById('v3-holdings-thead');
      const holdTbody = document.getElementById('v3-holdings-tbody');
      const MAX_HOLD = 10;
      if (holdThead) holdThead.innerHTML = `
        <tr style="background:#1a3c6e;color:#fff;font-size:0.86rem;">
          <th style="padding:10px 10px;text-align:center;min-width:36px;border-right:1px solid rgba(255,255,255,0.15);background:transparent;font-weight:700;">#</th>
          ${data.map(f => `<th style="padding:10px 10px;text-align:center;font-weight:700;min-width:120px;">
            <span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${f.color};margin-right:5px;vertical-align:middle;opacity:0.9;"></span>${esc((f._selectorName || f.name).split('–')[0].trim())}
          </th>`).join('')}
        </tr>`;
      if (holdTbody) {
        const rows = Array.from({ length: MAX_HOLD }, (_, ri) => {
          const cells = data.map(f => {
            const h = (f._topHoldings || [])[ri];
            if (!h) return `<td style="padding:8px 10px;text-align:center;font-size:0.86rem;color:#cbd5e1;">—</td>`;
            const name = h.companyName || h.name || '—';
            const wt   = h.weightText || (h.weight != null ? h.weight + '%' : '');
            const combinedText = wt ? `${name} ${wt}` : name;
            return `<td style="padding:8px 10px;font-size:0.86rem;color:var(--text);">
              <div style="font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px;" title="${esc(combinedText)}">
                ${esc(name)}${wt ? ` <span style="font-size:0.8rem;color:var(--text-muted);font-weight:500;">${esc(wt)}</span>` : ''}
              </div>
            </td>`;
          }).join('');
          return `<tr style="border-bottom:1px solid var(--border-light);${ri % 2 === 0 ? '' : 'background:#fafbfd;'}">
            <td style="padding:8px 10px;text-align:center;font-size:0.86rem;font-weight:700;color:var(--text-muted);background:#f8fafc;border-right:1px solid var(--border-light);">${ri + 1}</td>
            ${cells}
          </tr>`;
        });
        holdTbody.innerHTML = rows.join('');
      }

      // Top 10 Holdings compare matrix by asset name
      const normalizedFunds = data.map(f => {
        const holdings = (f._topHoldings || []).slice(0, 10).map(item => {
          const weight = parseFloat(String(item?.weight ?? '').replace(/%/g, '').trim());
          return {
            name: String(item?.companyName || item?.name || '').trim(),
            weight: Number.isNaN(weight) ? 0 : weight,
          };
        }).filter(item => item.name);

        const byName = new Map();
        holdings.forEach(item => byName.set(item.name, item.weight));

        return {
          label: (f._selectorName || f.name).split('–')[0].trim(),
          top10Sum: holdings.reduce((sum, item) => sum + item.weight, 0),
          holdings,
          byName,
        };
      });

      const rankedAssets = Array.from(new Set(
        normalizedFunds.flatMap(f => f.holdings.map(item => item.name))
      )).map(name => ({
        name,
        total: normalizedFunds.reduce((sum, fund) => sum + (fund.byName.get(name) || 0), 0),
      })).sort((a, b) => b.total - a.total);

      const assetNames = rankedAssets.map(item => item.name);
      const xMaxRaw = Math.max(1, ...normalizedFunds.map(fund => fund.top10Sum));
      const xMax = Math.ceil(xMaxRaw / 2) * 2;

      const holdingsDatasets = assetNames.map((name, idx) => ({
        label: name,
        data: normalizedFunds.map(fund => fund.byName.get(name) || 0),
        backgroundColor: holdingsPalette[idx % holdingsPalette.length],
        borderWidth: 0,
        barThickness: 18,
      }));

      mkChart('v3-top10-asset-chart', {
        type: 'bar',
        data: {
          labels: normalizedFunds.map(fund => fund.label),
          datasets: holdingsDatasets,
        },
        options: {
          ...common,
          indexAxis: 'y',
          plugins: {
            legend: { display: false },
            tooltip: {
              mode: 'index',
              intersect: false,
              callbacks: {
                label(ctx) {
                  const value = typeof ctx.raw === 'number' ? ctx.raw : 0;
                  return `${ctx.dataset.label}: ${value.toFixed(2)}%`;
                },
              },
            },
          },
          scales: {
            x: {
              stacked: true,
              beginAtZero: true,
              max: xMax,
              ticks: { callback: v => v + '%' },
              grid: { color: '#dbe4f0' },
            },
            y: {
              stacked: true,
              grid: { display: false },
              ticks: { font: { size: 12, weight: '600' } },
            },
          },
        },
      });

      const top10Legend = document.getElementById('v3-top10-asset-legend');
      if (top10Legend) top10Legend.innerHTML = holdingsDatasets.map(ds => `
        <div style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:var(--text-muted);max-width:320px;">
          <span style="width:12px;height:12px;border-radius:3px;background:${ds.backgroundColor};display:inline-block;flex-shrink:0;"></span>
          <span style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${esc(ds.label)}</span>
        </div>`).join('');

      latestExportPayload = buildTop10V3ExportPayload(
        CONFIG.PAGES['master-placeholder-8']?.title || 'Top 10 Holding',
        CONFIG.PAGES['master-placeholder-8']?.source || 'Multi-Fund Compare API',
        data,
      );
    };

    const normalizeTopHoldings = (holdings) => {
      const rawRows = Array.isArray(holdings)
        ? holdings
        : Array.isArray(holdings?.rows)
          ? holdings.rows
          : [];

      return rawRows
        .map((item) => {
          const name = item?.companyName || item?.name || item?.company || item?.holding || '';
          const weightNum = (() => {
            const candidates = [item?.weight, item?.percent, item?.weightPercent, item?.portfolioWeight];
            for (const candidate of candidates) {
              const parsed = parseFloat(String(candidate ?? '').replace(/%/g, '').trim());
              if (!Number.isNaN(parsed)) return parsed;
            }
            return null;
          })();
          const weightText = item?.weightText
            || (weightNum != null ? `${weightNum}%` : '');

          return {
            ...item,
            companyName: name,
            name,
            weight: weightNum,
            weightText,
          };
        })
        .filter(item => item.name);
    };

    // ── Load button click ──
    const loadBtn  = document.getElementById('v3-load-btn');
    const statusWrap = document.getElementById('v3-api-status-wrap');
    const statusEl   = document.getElementById('v3-api-status');
    const explorerSec = document.getElementById('v3-explorer-section');
    const rawEl     = document.getElementById('v3-raw');
    const toggleBtn = document.getElementById('v3-toggle-raw');
    const updateCardInfo = (fundList, liveData) => {
      (fundList || []).forEach((f, i) => {
        const infoEl = area.querySelector(`#v3-card-info-${f.idx}`);
        if (!infoEl) return;
        const d = liveData?.[i];
        if (d?._apiOk) {
          infoEl.textContent = d._isin || f.isin || '—';
          infoEl.style.color = '#1a3c6e';
          infoEl.style.fontWeight = '600';
        } else {
          infoEl.textContent = '(ข้อมูลจำลอง)';
          infoEl.style.color = '#94a3b8';
          infoEl.style.fontWeight = 'normal';
        }
      });
    };

    if (toggleBtn && rawEl) {
      toggleBtn.addEventListener('click', () => {
        rawEl.style.display = rawEl.style.display === 'none' ? 'block' : 'none';
        saveV3State({ rawVisible: rawEl.style.display !== 'none' });
      });
    }

    if (loadBtn) {
      loadBtn.addEventListener('click', async () => {

        // ── Build fund list from selectors ──
        const fundList = [];
        saveV3State({ selections: getSelections() });
        const iShareTicker = area.querySelector('#v3-sel-0')?.value?.trim() || '';
        if (iShareTicker) {
          fundList.push({ idx: 0, name: iShareTicker, isin: iShareTicker, color: fundColors[0] });
        }
        for (let fi = 1; fi <= 7; fi++) {
          const inp = area.querySelector(`#v3-sel-${fi}`);
          const code = inp?.value?.trim() || '';
          if (code && !code.startsWith('กองทุนไทย #')) {
            const lu = thaiLookup[code];
            const isin = lu?.isin || code;
            const shortName = lu?.name ? code + ' – ' + lu.name.substring(0, 18) : code;
            fundList.push({ idx: fi, name: shortName, isin, color: fundColors[fi] });
          }
        }

        if (!fundList.length) {
          if (statusWrap) statusWrap.style.display = '';
          if (statusEl) {
            statusEl.style.background = '#fef3c7'; statusEl.style.color = '#92400e';
            statusEl.textContent = '⚠️ กรุณาเลือกกองทุนอย่างน้อย 1 กองก่อนโหลดข้อมูล';
          }
          saveV3State({
            statusText: '⚠️ กรุณาเลือกกองทุนอย่างน้อย 1 กองก่อนโหลดข้อมูล',
            statusTone: 'warning',
          });
          return;
        }

        // ── Render demo immediately (visual feedback) ──
        const demoData = buildDemoData(fundList.map(f => f.name));
        render(demoData);

        if (statusWrap) statusWrap.style.display = '';
        if (statusEl) {
          statusEl.style.background = 'var(--primary-faint)';
          statusEl.style.color = 'var(--primary)';
          statusEl.textContent = `⏳ กำลังดึงข้อมูลจาก Code.gs API (${fundList.length} กองทุน)...`;
        }
        saveV3State({
          statusText: `⏳ กำลังดึงข้อมูลจาก Code.gs API (${fundList.length} กองทุน)...`,
          statusTone: 'loading',
        });
        loadBtn.disabled = true;
        loadBtn.textContent = 'กำลังโหลด...';

        const FIELDS = 'summary,fees,performance,holdings,sizes';

        // Helper: parse "1.23%" or "1.23" → number
        const pct = s => {
          const n = parseFloat(String(s ?? '').replace(/%/g, '').trim());
          return isNaN(n) ? null : n;
        };

        try {
          // ── Parallel calls: one per fund with correct ?isin=X&fields=Y ──
          const apiCalls = await Promise.allSettled(
            fundList.map(f =>
              fetch(`${TOP_10_HOLDING_API_URL}?isin=${encodeURIComponent(f.isin)}&fields=${FIELDS}`)
                .then(r => r.json())
                .catch(() => ({ ok: false, error: 'network error' }))
            )
          );

          // ── Map API responses → dashboard format (fallback to demoData) ──
          const liveData = fundList.map((f, i) => {
            const base = demoData[i];
            const raw  = apiCalls[i];
            const hasUsefulPayload = (payload) => !!(
              payload
              && (
                payload.ok
                || payload.holdings
                || payload.summary
                || payload.summaryHtml
                || payload.performance
                || payload.fees
              )
            );
            const api  = (raw.status === 'fulfilled' && hasUsefulPayload(raw.value)) ? raw.value : null;
            if (!api) return { ...base, _selectorName: f.name, _isin: f.isin, _feesRaw: null, _feesInitial: null, _feesMaxAnnual: null, _feesExit: null, _manager: null };

            // Name
            const fundName = api.summary?.fundName || f.name;

            // YTD / performance
            const ytd = pct(api.performance?.['YTD'])
                     ?? pct(api.performance?.['1Y'])
                     ?? pct(api.performance?.['1M'])
                     ?? base.ytd;

            // Fee
            const fee = pct(api.fees?.ongoingCharge) ?? base.fee;

            // Dividend
            const inc = (api.summary?.incomeTreatment || '').toLowerCase();
            const dividend = inc.includes('accum') ? 'ไม่มี' : (inc ? 'มี' : base.dividend);

            return { ...base, name: fundName, ytd, fee, dividend,
                     _topHoldings: normalizeTopHoldings(api.holdings), _apiOk: true,
                     _selectorName: f.name,
                     _isin: f.isin,
                     _feesRaw:       api.fees?.ongoingCharge    ?? null,
                     _feesInitial:   api.fees?.initialCharge    ?? null,
                     _feesMaxAnnual: api.fees?.maxAnnualCharge  ?? null,
                     _feesExit:      api.fees?.exitCharge       ?? null,
                     _manager:       api.manager?.name || null };
          });

          render(liveData);

          // ── Update card-info labels below each selector ──
          updateCardInfo(fundList, liveData);

          // Raw explorer
          const allRaw = Object.fromEntries(
            fundList.map((f, i) => {
              const r = apiCalls[i];
              return [f.name, r.status === 'fulfilled' ? r.value : { ok: false }];
            })
          );
          if (rawEl) rawEl.textContent = JSON.stringify(allRaw, null, 2);
          if (explorerSec) explorerSec.style.display = '';
          if (rawEl) rawEl.style.display = 'block';

          const okCount = apiCalls.filter(r =>
            r.status === 'fulfilled'
            && (
              r.value?.ok
              || r.value?.holdings
              || r.value?.summary
              || r.value?.summaryHtml
              || r.value?.performance
              || r.value?.fees
            )
          ).length;
          if (statusEl) {
            const allOk = okCount === fundList.length;
            statusEl.style.background = allOk ? '#f0fdf4' : '#fef3c7';
            statusEl.style.color      = allOk ? '#15803d' : '#92400e';
            statusEl.textContent      = allOk
              ? `✓ ดึงข้อมูลสำเร็จครบทุกกองทุน (${okCount}/${fundList.length})`
              : `⚠️ ดึงข้อมูลสำเร็จ ${okCount}/${fundList.length} กองทุน — กองที่เหลือแสดงข้อมูลจำลองแทน`;
          }
          saveV3State({
            selections: getSelections(),
            fundList,
            liveData,
            rawText: JSON.stringify(allRaw, null, 2),
            showExplorer: true,
            rawVisible: true,
            statusText: statusEl?.textContent || '',
            statusTone: okCount === fundList.length ? 'success' : 'warning',
          });

        } catch (err) {
          if (statusEl) {
            statusEl.style.background = '#fef2f2'; statusEl.style.color = '#dc2626';
            statusEl.textContent = '✕ เกิดข้อผิดพลาด: ' + err.message + ' — แสดงข้อมูลจำลองแทน';
          }
          saveV3State({
            statusText: '✕ เกิดข้อผิดพลาด: ' + err.message + ' — แสดงข้อมูลจำลองแทน',
            statusTone: 'error',
          });
        } finally {
          loadBtn.disabled = false;
          loadBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg> โหลดข้อมูลเปรียบเทียบ`;
        }
      });
    }

    const restoreSavedState = () => {
      const persisted = State.top10HoldingV3;
      if (!persisted) return;

      if (persisted.statusText && statusEl) {
        statusWrap.style.display = '';
        statusEl.textContent = persisted.statusText;
        if (persisted.statusTone === 'success') {
          statusEl.style.background = '#f0fdf4';
          statusEl.style.color = '#15803d';
        } else if (persisted.statusTone === 'warning') {
          statusEl.style.background = '#fef3c7';
          statusEl.style.color = '#92400e';
        } else if (persisted.statusTone === 'error') {
          statusEl.style.background = '#fef2f2';
          statusEl.style.color = '#dc2626';
        } else {
          statusEl.style.background = 'var(--primary-faint)';
          statusEl.style.color = 'var(--primary)';
        }
      }

      if (Array.isArray(persisted.liveData) && persisted.liveData.length) {
        render(persisted.liveData);
        updateCardInfo(persisted.fundList, persisted.liveData);
      }

      if (persisted.showExplorer && explorerSec) {
        explorerSec.style.display = '';
      }
      if (typeof persisted.rawText === 'string' && rawEl) {
        rawEl.textContent = persisted.rawText;
        rawEl.style.display = persisted.rawVisible ? 'block' : 'none';
      }
    };
    restoreSavedState();

    App._currentTableExport = () => latestExportPayload;
    App._currentImageExport = async () => {
      const target = $('#report-card', area);
      if (!target) throw new Error('ไม่พบส่วนกราฟสำหรับส่งออก');
      const visiblePanels = Array.from(target.querySelectorAll('.thv2-panel'))
        .filter(panel => panel.style.display !== 'none' && panel.id !== 'v3-explorer-section');
      if (!visiblePanels.length) throw new Error('ยังไม่มีข้อมูลกราฟสำหรับส่งออก');

      const exportShell = document.createElement('div');
      exportShell.style.display = 'flex';
      exportShell.style.flexDirection = 'column';
      exportShell.style.gap = '18px';
      exportShell.style.width = `${target.clientWidth || target.scrollWidth || 1200}px`;
      visiblePanels.forEach(panel => exportShell.appendChild(cloneNodeForCapture(panel)));

      document.body.appendChild(exportShell);
      exportShell.style.position = 'fixed';
      exportShell.style.left = '-100000px';
      exportShell.style.top = '0';
      exportShell.style.background = '#ffffff';
      exportShell.style.padding = '0';

      const { node, cleanup } = createCaptureTarget(area, exportShell);
      try {
        const blob = await elementToImageBlob(node);
        if (!blob) throw new Error('สร้างภาพไม่สำเร็จ');
        const image = await blobToDataURL(blob);
        return {
          filename: 'top10-holding-report',
          image,
        };
      } finally {
        cleanup();
        exportShell.remove();
      }
    };
    bindPageImageActions(area, 'report-card', 'top10-holding-report');
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
  _currentClipboardExport: null,
  _currentImageExport: null,
  _toastTimer:    null,
  _studioReady:   false,
  _studioState:   null,

  syncSidebarUI() {
    const app = $('#app');
    const toggles = $$('.sidebar-toggle');
    const backdrop = $('#sidebar-backdrop');
    if (!app || toggles.length === 0) return;

    const compact = isCompactViewport();
    const isOpen = compact ? State.sidebarOpen : !State.sidebarCollapsed;

    app.classList.toggle('sidebar-collapsed', !compact && State.sidebarCollapsed);
    app.classList.toggle('sidebar-open', compact && State.sidebarOpen);
    if (backdrop) {
      backdrop.classList.toggle('hidden', !(compact && State.sidebarOpen));
    }

    toggles.forEach(toggle => {
      const label = compact
        ? (isOpen ? 'ซ่อนเมนูด้านข้าง' : 'แสดงเมนูด้านข้าง')
        : (State.sidebarCollapsed ? 'ขยายเมนูด้านข้าง' : 'ยุบเมนูด้านข้าง');
      toggle.setAttribute('aria-expanded', String(isOpen));
      toggle.setAttribute('aria-label', label);
      toggle.title = label;
    });
  },

  toggleSidebar(forceOpen) {
    if (isCompactViewport()) {
      State.sidebarOpen = typeof forceOpen === 'boolean'
        ? forceOpen
        : !State.sidebarOpen;
    } else {
      const nextCollapsed = typeof forceOpen === 'boolean'
        ? !forceOpen
        : !State.sidebarCollapsed;
      State.sidebarCollapsed = nextCollapsed;
      State.sidebarOpen = false;
      saveSidebarPreference(nextCollapsed);
    }
    App.syncSidebarUI();
  },

  /* Init after login */
  init() {
    State.sidebarCollapsed = readSidebarPreference();
    State.sidebarOpen = false;
    $('#topbar-date').textContent = thaiDate();

    /* Nav */
    $$('.nav-item').forEach(el => {
      el.title = el.textContent.trim();
      el.addEventListener('click', e => {
        e.preventDefault();
        App.navigate(el.dataset.page);
        if (isCompactViewport()) App.toggleSidebar(false);
      });
    });

    $$('.sidebar-toggle').forEach(toggle => {
      toggle.addEventListener('click', () => App.toggleSidebar());
    });
    $('#sidebar-backdrop').addEventListener('click', () => App.toggleSidebar(false));
    window.addEventListener('resize', () => {
      if (!isCompactViewport()) State.sidebarOpen = false;
      App.syncSidebarUI();
    });
    window.addEventListener('keydown', e => {
      if (e.key === 'Escape' && isCompactViewport() && State.sidebarOpen) {
        App.toggleSidebar(false);
      }
    });
    App.syncSidebarUI();

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
        pageSize: SELECT_FUND_DEFAULT_PAGE_SIZE,
      };
      State.top10HoldingV3 = null;
      State.currentUser = { name: '', email: '' };
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
      ? (State.selectFundFilters.pageSize || SELECT_FUND_DEFAULT_PAGE_SIZE)
      : PAGE_SIZE_OPTIONS[0];
    App._currentExport = null;
    App._currentTableExport = null;
    App._currentClipboardExport = null;
    App._currentImageExport = null;
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
      'thai-annualized-v2': { title: 'กองทุนไทย Annualized Return', subtitle: 'สลับมุมมอง Return และ Rank ได้' },
      'thai-calendar':     { title: 'กองทุนไทย Calendar Year', subtitle: '' },
      'master-annualized': { title: 'Master Fund Annualized Return', subtitle: 'จับคู่ด้วย ISIN และแยก Base Currency' },
      'master-annualized-v2': { title: 'Master Fund Annualized Return', subtitle: 'จับคู่ด้วย ISIN และแยก Base Currency' },
      'master-calendar':   { title: 'Master Fund Calendar Year', subtitle: '' },
      'master-placeholder-1': { title: 'ค่าธรรมเนียม', subtitle: 'เทียบ TER ของกองไทยกับ Ongoing Cost ของ Master Fund' },
      'master-placeholder-2': { title: 'Top 10 Holding', subtitle: '' },
      'master-placeholder-3': { title: 'Cost Efficiency Master Fund 5Y', subtitle: 'ดูค่าธรรมเนียมเทียบผลตอบแทนย้อนหลัง 5 ปี' },
      'master-placeholder-4': { title: 'ค่าธรรมเนียม', subtitle: '' },
      'master-placeholder-12': { title: 'เปรียบเทียบค่าธรรมเนียม', subtitle: 'เตรียมเมนูไว้สำหรับต่อยอดรายงานเปรียบเทียบค่าธรรมเนียม' },
      'master-placeholder-7': { title: 'Top 10 Holding V2', subtitle: 'เปรียบเทียบหลายกองในหน้าเดียว' },
      'master-placeholder-8': { title: 'Top 10 Holding', subtitle: 'วิเคราะห์เปรียบเทียบกองทุน 8 ช่องพร้อมกัน — iShare Index 1 กอง + กองทุนไทย 7 กอง' },
      'master-placeholder-5': { title: 'ปัจจัยประกอบอื่นๆ', subtitle: 'Sharpe, Sortino, Information, Treynor Ratio ทุก Period' },
      'master-placeholder-6': { title: 'Income Fund', subtitle: '' },
      'master-placeholder-9': { title: 'ปัจจัยประกอบอื่นๆ 2', subtitle: 'Sharpe vs Return' },
      'master-placeholder-10': { title: 'ปัจจัยประกอบอื่นๆ 3', subtitle: 'Sortino vs Return' },
      'master-placeholder-11': { title: 'ปัจจัยประกอบอื่นๆ 4', subtitle: 'Max DD vs Sortino' },
      'notes': { title: 'บันทึกข้อมูล', subtitle: 'ดราฟงานค้างและโหลดกลับมาทำต่อ' },
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
      case 'master-placeholder-12': Pages.feeComparisonPlaceholder(area);   break;
      case 'master-placeholder-7': Pages.masterMenu02V2(area);              break;
      case 'master-placeholder-8': Pages.masterMenu02V3(area);              break;
      case 'master-placeholder-5': Pages.masterOtherFactors(area);             break;
      case 'master-placeholder-6': Pages.placeholder(area, 'Income Fund');           break;
      case 'master-placeholder-9':
      case 'master-placeholder-10':
      case 'master-placeholder-11': Pages.comingSoon(area);                 break;
      case 'notes':             Pages.notesPage(area);                      break;
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
/* ============================================================
   QUARTER SELECTOR – Auto-detect tabs from Google Sheets
   ============================================================ */

const QuarterSelector = {

  /* ดึงรายชื่อ Tab จาก Sheet แล้วกรองเฉพาะรูปแบบ YYYY-QN */
  async detect() {
    const sel    = $('#quarter-select');
    const status = $('#quarter-status');
    if (!sel) return;

    status.className = 'quarter-status loading';
    sel.disabled = true;

    try {
      const sheetId = CONFIG.SHEETS?.MASTER_FUND_ID;
      if (!sheetId) throw new Error('ไม่พบ MASTER_FUND_ID');

      const { tabs } = await SheetsAPI.getSheetTabs(sheetId);

      /* กรองเฉพาะ Tab ที่เป็น Quarter เช่น 2025-Q1, 2026-Q3 */
      const quarters = tabs
        .filter(t => /^\d{4}-Q[1-4]$/i.test(t.trim()))
        .sort()
        .reverse(); // ล่าสุดขึ้นก่อน

      if (quarters.length === 0) {
        /* ถ้าไม่เจอ Tab รูปแบบ Quarter ให้แสดงทุก Tab */
        quarters.push(...tabs);
      }

      State.availableQuarters = quarters;
      State.currentQuarter    = quarters[0] || null;

      /* Populate dropdown */
      sel.innerHTML = quarters
        .map(q => `<option value="${q}"${q === State.currentQuarter ? ' selected' : ''}>${q}</option>`)
        .join('');
      sel.disabled = false;

      status.className  = 'quarter-status ok';
      status.textContent = '✓';

      /* เมื่อ user เลือก Quarter ใหม่ */
      sel.addEventListener('change', () => {
        const newQ = sel.value;
        if (newQ === State.currentQuarter) return;
        State.currentQuarter = newQ;
        clearCache();
        State.top10HoldingV3 = null;
        /* Re-render หน้าปัจจุบัน */
        App.navigate(State.page);
        toast(`เปลี่ยนเป็นข้อมูล ${newQ} แล้ว`, 'info');
      });

    } catch (err) {
      status.className  = 'quarter-status error';
      status.textContent = '✕';
      sel.innerHTML = `<option value="${CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1'}">
        ${CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1'} (default)</option>`;
      sel.disabled = false;
      State.currentQuarter = CONFIG.PAGES?.['select-fund']?.tabName || null;
      console.warn('Quarter auto-detect failed:', err.message);
    }
  },

  /* ถ้าเป็น local mode ให้ซ่อน selector */
  hide() {
    const el = $('#quarter-selector');
    if (el) el.style.display = 'none';
  },
};

document.addEventListener('DOMContentLoaded', () => {

  if (isLocalMode()) {
    State.currentUser = { name: 'Local Data Mode', email: 'loading from /Data only' };
    $('#user-name').textContent   = 'Local Data Mode';
    $('#user-email').textContent  = CONFIG.DATA_SOURCE === 'local_only'
      ? 'loading from /Data only'
      : 'loading from /Data first';
    $('#user-avatar').textContent = 'L';
    $('#login-screen').classList.add('hidden');
    $('#app').classList.remove('hidden');
    QuarterSelector.hide(); // local mode ไม่ต้องการ Quarter selector
    App.init();
    return;
  }

  /* ── DEV BYPASS ── */
  if (CONFIG.BYPASS_LOGIN) {
    State.currentUser = { name: 'Dev Mode', email: 'bypass login' };
    $('#user-name').textContent   = 'Dev Mode';
    $('#user-email').textContent  = 'bypass login';
    $('#user-avatar').textContent = 'D';
    $('#login-screen').classList.add('hidden');
    $('#app').classList.remove('hidden');
    QuarterSelector.hide(); // bypass mode ไม่มี token จริง
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

      State.currentUser = { name, email };
      $('#user-name').textContent   = name;
      $('#user-email').textContent  = email;
      $('#user-avatar').textContent = initial;

      $('#login-screen').classList.add('hidden');
      $('#app').classList.remove('hidden');

      App.init();
      QuarterSelector.detect(); // ← ตรวจ Quarter จาก Sheets หลัง login สำเร็จ

    } catch (e) {
      btnSignin.disabled = false;
      btnSignin.innerHTML = googleSvg;
      showLoginError('เข้าสู่ระบบไม่สำเร็จ: ' + e.message);
    }
  });
});
