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
const INCOME_FUND_SELECTION_SHEET_ID = '1Qe3B4nEz9brLRVSMZcTxilmNmfGLs3K3F5DAZ5meKVg';
const INCOME_FUND_SELECTION_TAB = 'Income Fund Selection';
const INCOME_FUND_SELECTION_HEADERS = ['Quarter', 'Fund Code', 'proj_id', 'fund_class_name', 'proj_abbr_name', 'Selected', 'Updated At', 'Updated By'];
const INCOME_FUND_DIVIDEND_DB_FILE = 'Dividend%20Database/exports/dividend_history_database.json';
const INCOME_FUND_DIVIDEND_DB_FALLBACK_FILES = [
  INCOME_FUND_DIVIDEND_DB_FILE,
  'Dividend Database/exports/dividend_history_database.json',
];
const INCOME_FUND_DIVIDEND_DRIVE_FOLDER_ID = '1SBJshoukb5O9hdTawTtexVQzo4oLfW3u';
const INCOME_FUND_DIVIDEND_DB_FILE_NAME = 'dividend_history_database.json';
const FT_HISTORICAL_DRIVE_FOLDER_ID = '1Locig3aW7hVs0SxCoeFpa76OJ0XcF1rg';
const FT_HISTORICAL_DB_FILE_NAME = 'ft_historical_prices_database.json';

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
  incomeFundFilters: {
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
  incomeFundSort: {
    key: '',
    dir: 'asc',
  },
  incomeFund2Filters: {
    query: '',
    source: '',
  },
  incomeFund2Sort: {
    key: 'latestDate',
    dir: 'desc',
  },
  reportSorts: {
    'thai-annualized': { key: '', dir: 'asc' },
    'thai-annualized-rank': { key: '', dir: 'asc' },
    'thai-annualized-v2': { key: '', dir: 'asc' },
    'thai-calendar': { key: '', dir: 'asc' },
    'master-annualized': { key: 'r5y', dir: 'desc' },
    'master-annualized-v2': { key: 'r5y', dir: 'desc' },
    'master-calendar': { key: 'ret-2025', dir: 'desc' },
    'master-placeholder-9': { key: '', dir: 'asc' },
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
  incomeFundSelectedKeys: new Set(),
  incomeFundSelectionRows: [],
  incomeFundSelectionLoaded: false,
  incomeFundSelectionSaving: false,
  incomeFundDividendDatabase: null,
  incomeFundDividendGithubToken: '',
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
  fundOverrides: { items: {}, updatedAt: null },
  fixedIncomeFactorsOverrides: { items: {}, updatedAt: null },
  fixedIncomeFactorsEditMode: false,
  fundDataManager: {
    query: '',
    selectedKey: '',
    showOnlyMapped: false,
    draftMapping: null,
  },
  fundSelectionLogs: {
    selectedItemId: '',
    query: '',
    roleFilter: '',
    baseRevision: null,
  },
	  dataImport: {
	    selectedJobKey: 'percentrank',
	    rawTab: '',
	    targetTab: '',
    preview: null,
    columnTypes: {},
    rawFiles: [],
    rawFilesError: '',
    isImporting: false,
	    jsonReadiness: null,
	    isCheckingJson: false,
	    isExportingJson: false,
	    previewScrollLeft: 0,
	  },
	  secDataImport: {
	    preview: null,
	    columnTypes: {},
	    targetTab: 'data_preparation',
	    targetExists: false,
	    isImporting: false,
	    isLoadingPreview: false,
	    jsonStatus: null,
	    isExportingJson: false,
	  },
  ftHistoricalImport: {
    url: 'https://markets.ft.com/data/etfs/tearsheet/summary?s=IXN:PCQ:USD',
    startDate: '',
    endDate: '',
    symbol: '',
  },
  ftHistoricalDatabase: null,
	  masterAllocations: { items: {}, updatedAt: null },
	};

const RAW_FILES_FOLDER_ID = '1_9lmPWFmack0DqaKuen5kbY57RJCFtmM';
const RAW_FILES_FOLDER_URL = 'https://drive.google.com/drive/u/2/folders/1_9lmPWFmack0DqaKuen5kbY57RJCFtmM';
const JSON_DRIVE_ROOT_FOLDER_ID = '1vUWAU5qP0qiIHPa5C4TZUybVmEwqfl6W';
const JSON_DRIVE_ROOT_FOLDER_URL = 'https://drive.google.com/drive/u/2/folders/1vUWAU5qP0qiIHPa5C4TZUybVmEwqfl6W';
const JSON_EXPORT_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwV-OUqOXxv1GAkdXJP5ICbXAa4gzEAxWjU6Yd9Z_KKkO2y1dhm4aWrjHik7_4gJNHx/exec';
const JSON_EXPORT_SECRET_KEY = 'sheets-to-drive-json';
const DRAFT_API_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzcOc5s3SYM2gkw3trpO43Jp0wwTasOIU8Mns4LJ-YQveE2cq5LXX98G0NVW7qHohGFKA/exec';
const DRAFT_API_SECRET_KEY = 'change-this-draft-api-key';
const MASTER_ALLOCATIONS_API_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxzNvGzKo8YGBPyDx8XOqb73hXTx_NetuBLTB4npMae9Jg1KM2HYZmaccds4e0koPMxqA/exec';
const MASTER_ALLOCATIONS_API_SECRET_KEY = 'change-this-master-allocations-api-key';
const FIXED_INCOME_FACTORS_API_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzHr34Gg48W9awG2Hde3P_XOxGUVLf2k_W-ySRWt-2IaIa7VOD6IsYfm-LmfhaYSwRn/exec';
const FIXED_INCOME_FACTORS_API_SECRET_KEY = 'change-this-fixed-income-factors-api-key';
const FT_HISTORICAL_API_WEB_APP_URL = '';
const FT_HISTORICAL_API_SECRET_KEY = 'change-this-ft-historical-api-key';

const JSON_STORE = {
  rootName: 'JSON Files',
  baseFiles: {
    secApi: 'Data For SEC API.json',
    percentrank: 'Fund Key Performance AVP.json',
    thaiQuality: 'AVP Thai Fund for Quality.json',
    masterFund: 'AVP Master Fund ID.json',
  },
  overrideFiles: {
    fundOverrides: 'fund_overrides.json',
    masterAllocations: 'fund_master_allocations.json',
    fixedIncomeFactorsOverrides: 'fixed_income_factors_overrides.json',
  },
};

const DATASET_REGISTRY = {
  secApi: {
    key: 'secApi',
    label: 'Data For SEC API',
    sheetId: '16agx9pl9adtMh-U7MCbgnIncBxpciCvFgsdurH6Ob8w',
    jsonFileName: JSON_STORE.baseFiles.secApi,
    localFile: 'Data/Data For SEC API - 2026-Q1.json',
  },
  fundKeyPerformance: {
    key: 'fundKeyPerformance',
    label: 'Fund Key Performance AVP',
    sheetId: '1s-0ciSOB2Tj0C9azeMXyd1zZxljOg8I5QilI0FgjdW4',
    jsonFileName: JSON_STORE.baseFiles.percentrank,
    localFile: 'Data/Fund Key Performance AVP - 2026-Q1.json',
  },
  thaiQuality: {
    key: 'thaiQuality',
    label: 'AVP Thai Fund for Quality',
    sheetId: '1m1rSyJAel9atGMrmeRSwgYWa9wgc4gi7-3cp4Yvc8GM',
    jsonFileName: JSON_STORE.baseFiles.thaiQuality,
    localFile: 'Data/AVP Thai Fund for Quality - 2026-Q1.json',
  },
  masterFund: {
    key: 'masterFund',
    label: 'AVP Master Fund ID',
    sheetId: '10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig',
    jsonFileName: JSON_STORE.baseFiles.masterFund,
    localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
  },
};

const REQUIRED_QUARTER_DATASET_KEYS = ['secApi', 'fundKeyPerformance', 'thaiQuality', 'masterFund'];

const PAGE_DATASET_KEYS = {
  'select-fund': 'fundKeyPerformance',
  'thai-annualized': 'thaiQuality',
  'thai-annualized-v2': 'thaiQuality',
  'thai-calendar': 'thaiQuality',
  'master-annualized': 'masterFund',
  'master-annualized-v2': 'masterFund',
  'master-calendar': 'masterFund',
  'master-placeholder-1': 'secApi',
  'master-placeholder-2': 'masterFund',
  'master-placeholder-3': 'masterFund',
  'master-placeholder-4': 'secApi',
  'master-placeholder-5': 'masterFund',
  'master-placeholder-7': 'masterFund',
  'master-placeholder-9': 'masterFund',
  'master-placeholder-10': 'masterFund',
  'master-placeholder-11': 'masterFund',
  'master-placeholder-12': 'secApi',
  'robustness-ft-import': 'masterFund',
  'ft-top10-holding': 'masterFund',
  'upside-downside-capture': 'masterFund',
};

const FUND_LIST_UPDATE_SAMPLE_ROWS = [
  {
    id: 'FLU-2026Q1-001',
    list_from: '2025-Q3',
    list_to: '2026-Q1',
    type: 'Core - General',
    asset_class: 'Money Market',
    fund_list_old: '',
    fund_list_current: 'KTSV-A',
    change_type: 'เพิ่มเข้ามาใหม่',
    note: 'เพิ่มเข้ามาใหม่',
  },
  {
    id: 'FLU-2026Q1-002',
    list_from: '2025-Q3',
    list_to: '2026-Q1',
    type: 'Core - General',
    asset_class: 'Global Bond (Income)',
    fund_list_old: 'SCBPINA',
    fund_list_current: 'B-IR-FOF',
    change_type: 'สลับตำแหน่ง',
    note: 'สลับตำแหน่ง',
  },
  {
    id: 'FLU-2026Q1-003',
    list_from: '2025-Q3',
    list_to: '2026-Q1',
    type: 'Core - General',
    asset_class: 'Global EQ',
    fund_list_old: 'KT-GESG-A',
    fund_list_current: '',
    change_type: 'นำออก',
    note: 'นำ KT-GESG-A ออก',
  },
  {
    id: 'FLU-2026Q1-004',
    list_from: '2025-Q3',
    list_to: '2026-Q1',
    type: 'Core - RMF',
    asset_class: 'Global Technology EQ',
    fund_list_old: 'K-GTECHRMF',
    fund_list_current: 'ES-TECHRMF',
    change_type: 'สลับตำแหน่ง',
    note: 'สลับตำแหน่ง',
  },
  {
    id: 'FLU-2026Q1-005',
    list_from: '2025-Q3',
    list_to: '2026-Q1',
    type: 'Core - SSF',
    asset_class: 'Thai Bond',
    fund_list_old: '',
    fund_list_current: 'K-FIXEDPLUS-SSF',
    change_type: 'เพิ่มเข้ามาใหม่',
    note: 'เพิ่มเข้ามาใหม่',
  },
];

const FUND_LIST_UPDATE_COLUMNS = [
  'id',
  'list_from',
  'list_to',
  'type',
  'asset_class',
  'fund_list_old',
  'fund_list_current',
  'change_type',
  'note',
  'updated_at',
  'updated_by',
];

const FUND_LIST_UPDATE_STATUS_OPTIONS = [
  'เหมือนเดิม',
  'เพิ่มเข้ามาใหม่',
  'นำออก',
  'สลับตำแหน่ง',
  'เปลี่ยนแปลง',
];

const FUND_LIST_UPDATE_ASSET_OPTIONS = [
  'Money Market',
  'Short-term Bond',
  'Mid-Term Bond',
  'Thai Bond',
  'Foreign Income Bond',
  'Global Bond (Income)',
  'Thai Mid/Long-Term Bond',
  'Thai Long-Term Bond',
  'Thai Property (Mixed)',
  'Mixed (EQ & Others)',
  'Global Multi-Asset',
  'Global EQ',
  'Asia EQ',
  'Vietnam EQ',
  'China A-Shares EQ',
  'China A-Share EQ',
  'China All-shares EQ',
  'India EQ',
  'Japan EQ',
  'US EQ',
  'US EQ (Nasdaq)',
  'Europe EQ',
  'Thai EQ',
  'Thai EQ (Passive)',
  'Global Technology EQ',
  'Tech EQ',
  'Global Healthcare EQ',
  'Global Healthcare',
  'Global Infrastructure EQ',
  'Global REITs',
  'REITs (Thai & Foreign)',
  'Gold',
  'Gold (Hedged)',
  'Gold (Unhedged)',
  'Covered Call Strategy (Global)',
];

const FUND_LIST_UPDATE_DRAFT_KEY = 'fund-list-update-draft-v1';
const FUND_LIST_UPDATE_SEED_FILE = 'Data/fund_list_update_q3_2025_seed.json';
const FUND_SELECTION_LOGS_DRIVE_FOLDER_ID = '12ciJQq-dpBr-DpdnzXCOXqtW_ijctJN6';
const FUND_SELECTION_LOGS_API_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycby3cNUpeF8IZrQHFnPbAtd1zJXkA9qu-QtCTgzyYO96O47ymuE-pXFRltxrKRYUg8GpXw/exec';
const FUND_SELECTION_LOGS_API_SECRET_KEY = 'change-this-fund-selection-logs-api-key';
const FUND_SELECTION_LOGS_SHEET_ID = '1fOdq3JSKTjRZLE8sQ62jn3OmmuKGuhDo2tUCLjy1zIg';
const FUND_SELECTION_LOGS_SHEET_HEADERS = [
  'quarter',
  'item_order',
  'item_id',
  'asset_class',
  'fund_type',
  'category',
  'role',
  'fund_code',
  'status',
  'reason',
  'tags',
  'data_as_of',
  'item_revision',
  'updated_by',
  'updated_at',
  'mention_id',
  'row_revision',
  'deleted',
];
const FUND_SELECTION_LOG_ROLES = [
  { value: 'mainChoice', label: 'ตัวเลือกหลัก' },
  { value: 'secondaryChoice', label: 'ตัวเลือกรอง' },
  { value: 'additionalNote', label: 'ความเห็นเพิ่มเติม' },
  { value: 'notSelected', label: 'ไม่ถูกคัดเลือก' },
];
const FUND_SELECTION_LOG_SENTIMENTS = [
  { value: 'positive', label: 'บวก' },
  { value: 'neutral', label: 'กลาง' },
  { value: 'negative', label: 'ลบ/ข้อควรระวัง' },
];
const FUND_SELECTION_LOG_STATUS_OPTIONS = [
  'กองทุนเดิม',
  'กองทุนเดิม (สลับตำแหน่ง)',
  'กองทุนเดิม (Passive)',
  'กองทุนเดิม (Active)',
  'กองทุนใหม่',
];
const FUND_SELECTION_TAG_RULES = [
  { tag: 'fee-low', keywords: ['ค่าธรรมเนียมต่ำ', 'fee ต่ำ', 'fee low', 'ter ต่ำ', 'ค่าธรรมเนียมถูก', 'ไม่แพง'] },
  { tag: 'fee-lowest', keywords: ['ค่าธรรมเนียมต่ำที่สุด', 'fee ต่ำที่สุด', 'ถูกที่สุด'] },
  { tag: 'fee-high', keywords: ['ค่าธรรมเนียมสูง', 'fee สูง', 'ter สูง', 'แพงกว่า'] },
  { tag: 'large-size', keywords: ['ขนาดกองใหญ่', 'กองใหญ่', 'fund size ใหญ่', 'ขนาดกองทุนใหญ่'] },
  { tag: 'small-size', keywords: ['ขนาดกองเล็ก', 'กองเล็ก', 'ขนาดกองทุนเล็ก', 'ขนาดกองทุนน้อย'] },
  { tag: 'performance', keywords: ['ผลตอบแทนดี', 'ผลตอบแทนโดดเด่น', 'ผลการดำเนินงานโดดเด่น', 'performance ดี'] },
  { tag: 'consistent', keywords: ['สม่ำเสมอ', 'สม่ําเสมอ', 'consistent'] },
  { tag: 'underperform', keywords: ['underperform', 'ผลตอบแทนต่ำกว่า', 'แย่กว่ากลุ่ม', 'ทำได้ไม่ดี'] },
  { tag: 'volatility-low', keywords: ['ผันผวนต่ำ', 'ความผันผวนต่ำ', 'volatility ต่ำ', 'sd ต่ำ'] },
  { tag: 'volatility-high', keywords: ['ผันผวนสูง', 'ความผันผวนสูง', 'volatility สูง', 'ขึ้นลงแรง'] },
  { tag: 'duration-low', keywords: ['duration ต่ำ', 'duration สั้น'] },
  { tag: 'duration-long', keywords: ['duration สูง', 'duration ยาว'] },
  { tag: 'diversified', keywords: ['กระจายตัวดี', 'กระจายการลงทุนดี', 'กระจายลงทุนดี', 'diversified'] },
  { tag: 'concentrated', keywords: ['กระจุก', 'กระจุกตัว', 'concentration', 'top 10'] },
  { tag: 'passive', keywords: ['passive', 'กองดัชนี', 'ดัชนี'] },
  { tag: 'active', keywords: ['active', 'ผู้จัดการกองทุน', 'เลือกหุ้นเอง'] },
  { tag: 'hedged', keywords: ['hedged', 'ป้องกันความเสี่ยงค่าเงิน', 'ป้องกันค่าเงิน'] },
  { tag: 'unhedged', keywords: ['unhedged', 'ไม่ป้องกันความเสี่ยงค่าเงิน', 'ไม่ป้องกันค่าเงิน', 'uh'] },
  { tag: 'new', keywords: ['กองใหม่', 'ใหม่', 'new'] },
  { tag: 'not-selected', keywords: ['ไม่ถูกคัดเลือก', 'ตัดออก', 'ไม่นำมาพิจารณา', 'ไม่แนะนำ', 'not selected'] },
  { tag: 'credit-risk', keywords: ['ความเสี่ยงเครดิต', 'credit risk', 'หุ้นกู้เอกชน', 'เครดิต'] },
  { tag: 'derivatives', keywords: ['derivatives', 'สัญญาซื้อขายล่วงหน้า'] },
  { tag: 'foreign-investment', keywords: ['ต่างประเทศ', 'ลงทุนต่างประเทศ', 'foreign'] },
  { tag: 'trading', keywords: ['trade', 'trading', 'ลงทุนระยะสั้น', 'ซื้อขายระยะสั้น'] },
  { tag: 'monitor', keywords: ['เฝ้าติดตาม', 'ติดตาม', 'track record สั้น', 'กองออกใหม่'] },
];

const DATA_IMPORT_JOBS = {
  percentrank: {
    key: 'percentrank',
    sourceLabel: 'Fund Key Performance AVP _ Morningstar (Use)',
    sourceType: 'google-sheet',
    sourceFileName: 'Fund Key Performance AVP _ Morningstar (Use)',
    targetLabel: 'Fund Key Performance AVP',
    label: 'Fund Key Performance AVP',
    rawSheetId: '1B8xjgXN9D1-hYR7uVn61y-LyBKbdy9eLX3kFZdAFCAU',
    rawGid: 1255878181,
    rawFolderId: RAW_FILES_FOLDER_ID,
    rawFolderUrl: RAW_FILES_FOLDER_URL,
    targetSheetId: '1s-0ciSOB2Tj0C9azeMXyd1zZxljOg8I5QilI0FgjdW4',
    targetPageKey: 'select-fund',
    defaultTab: '2026-Q1',
  },
  thaiQuality: {
    key: 'thaiQuality',
    sourceLabel: 'AVP Thai Fund for Quality.xlsx',
    sourceType: 'xlsx',
    sourceFileName: 'AVP Thai Fund for Quality.xlsx',
    targetLabel: 'AVP Thai Fund for Quality',
    label: 'AVP Thai Fund for Quality',
    rawFolderId: RAW_FILES_FOLDER_ID,
    rawFolderUrl: RAW_FILES_FOLDER_URL,
    targetSheetId: '1m1rSyJAel9atGMrmeRSwgYWa9wgc4gi7-3cp4Yvc8GM',
    targetPageKey: 'thai-annualized-v2',
    cleaner: 'masterFundRaw',
    defaultTab: '2026-Q1',
  },
  masterFund: {
    key: 'masterFund',
    sourceLabel: 'AVP Master Fund ID.xlsx',
    sourceType: 'xlsx',
    sourceFileName: 'AVP Master Fund ID.xlsx',
    cleaner: 'masterFundRaw',
    targetLabel: 'AVP Master Fund ID',
    label: 'AVP Master Fund ID',
    rawFolderId: RAW_FILES_FOLDER_ID,
    rawFolderUrl: RAW_FILES_FOLDER_URL,
    targetSheetId: '10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig',
    targetPageKey: 'master-annualized-v2',
    defaultTab: '2026-Q1',
  },
};

const DATA_IMPORT_JSON_EXPORTS = [
  { jobKey: 'percentrank', dataset: 'fund-key-performance', fileName: JSON_STORE.baseFiles.percentrank },
  { jobKey: 'thaiQuality', dataset: 'thai-quality', fileName: JSON_STORE.baseFiles.thaiQuality },
  { jobKey: 'masterFund', dataset: 'master-fund', fileName: JSON_STORE.baseFiles.masterFund },
];

const SEC_DATA_PREPARATION_TARGETS = [
  {
    key: 'data_preparation',
    label: 'data_preparation',
    spreadsheetId: '1SsL8fXFmKsAfnakrIBTtglZqhErrYz8SfiiZb4wKGDs',
    url: 'https://docs.google.com/spreadsheets/d/1SsL8fXFmKsAfnakrIBTtglZqhErrYz8SfiiZb4wKGDs/edit?gid=1854000650#gid=1854000650',
  },
];

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

function googleSheetUrl(spreadsheetId, gid = '') {
  const id = String(spreadsheetId || '').trim();
  if (!id) return '';
  const gidPart = gid !== '' && gid !== null && gid !== undefined
    ? `#gid=${encodeURIComponent(String(gid))}`
    : '';
  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(id)}/edit${gidPart}`;
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

function pageToolActions(pageKey, sourceLabel = '', extraActions = '', extraButtons = '') {
  const sourceBadge = getPageDataSourceBadge(pageKey);
  return `
    <div class="page-tools">
      <div class="page-tools-meta">
        ${sourceBadgeHtml(pageKey, sourceLabel)}
        ${sourceBadge ? `<span class="badge badge-data-origin">${esc(sourceBadge)}</span>` : ''}
        ${extraActions}
      </div>
      <div class="page-tools-actions">
        ${extraButtons}
        <button class="btn btn-ghost" id="btn-copy-table" title="คัดลอกข้อมูลแบบคงรูปแบบ">
          คัดลอกแบบมีฟอร์แมต
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

function bindPageImageActions(area, cardId, filename) {
  const card = $(`#${cardId}`, area);
  if (!card) return;
  const copyBtn = $('#btn-copy-table', area);
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

function applyDatasetConfig(pageKey, pageConfig) {
  const datasetKey = pageConfig.datasetKey || PAGE_DATASET_KEYS[pageKey];
  const dataset = DATASET_REGISTRY[datasetKey];
  if (!dataset) return pageConfig;
  return {
    ...pageConfig,
    datasetKey,
    sheetId: pageConfig.sheetId || dataset.sheetId,
    source: pageConfig.source || dataset.label,
    localFile: pageConfig.localFile || dataset.localFile,
    preferDriveJson: pageConfig.preferDriveJson !== false,
    driveJsonRootFolderId: pageConfig.driveJsonRootFolderId || JSON_DRIVE_ROOT_FOLDER_ID,
    driveJsonFileName: pageConfig.driveJsonFileName || dataset.jsonFileName,
    driveJsonFolderSegments: pageConfig.driveJsonFolderSegments || ['{year}', '{quarter}', 'base'],
  };
}

function ensureExtendedPageConfigs() {
  if (!CONFIG.PAGES) CONFIG.PAGES = {};

  const defaults = {
    'select-fund': {
      sheetId: CONFIG.SHEETS?.FUND_KEY_PERFORMANCE || CONFIG.SHEETS?.PERCENTRANK_FREESTYLE || '',
      tabName: '2026-Q1',
      title: 'เลือกกองทุน',
      source: 'Fund Key Performance AVP',
      localFile: 'Data/Fund Key Performance AVP - 2026-Q1.json',
      datasetKey: 'fundKeyPerformance',
    },
    'thai-annualized': {
      sheetId: CONFIG.SHEETS?.THAI_FUND_QUALITY || '',
      tabName: '2026-Q1',
      title: 'กองทุนไทย Annualized Return',
      source: 'AVP Thai Fund for Quality',
      localFile: 'Data/AVP Thai Fund for Quality - 2026-Q1.json',
      datasetKey: 'thaiQuality',
    },
    'thai-annualized-v2': {
      sheetId: CONFIG.SHEETS?.THAI_FUND_QUALITY || '',
      tabName: '2026-Q1',
      title: 'กองทุนไทย Annualized Return',
      source: 'AVP Thai Fund for Quality',
      localFile: 'Data/AVP Thai Fund for Quality - 2026-Q1.json',
      datasetKey: 'thaiQuality',
    },
    'thai-calendar': {
      sheetId: CONFIG.SHEETS?.THAI_FUND_QUALITY || '',
      tabName: '2026-Q1',
      title: 'กองทุนไทย Calendar Year',
      source: 'AVP Thai Fund for Quality',
      localFile: 'Data/AVP Thai Fund for Quality - 2026-Q1.json',
      datasetKey: 'thaiQuality',
    },
    'master-annualized': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'Master Fund Annualized Return',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'master-annualized-v2': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'Master Fund Annualized Return',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'master-calendar': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'Master Fund Calendar Year',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'master-placeholder-1': {
      sheetId: CONFIG.SHEETS?.RAW_FOR_SEC || '',
      tabName: '2026-Q1',
      title: 'ค่าธรรมเนียม',
      source: 'Data For SEC API + AVP Master Fund ID',
      localFile: 'Data/Data For SEC API - 2026-Q1.json',
      datasetKey: 'secApi',
    },
    'master-placeholder-2': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'Top 10 Holding',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'master-placeholder-3': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'Cost Efficiency Master Fund 5Y',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'master-placeholder-4': {
      sheetId: CONFIG.SHEETS?.RAW_FOR_SEC || '',
      tabName: '2026-Q1',
      title: 'ค่าธรรมเนียม',
      source: 'Data For SEC API + AVP Master Fund ID',
      localFile: 'Data/Data For SEC API - 2026-Q1.json',
      datasetKey: 'secApi',
    },
    'master-placeholder-5': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'ปัจจัยประกอบอื่นๆ',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'master-placeholder-9': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'ปัจจัยประกอบ กองทุนตราสารหนี้',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'master-placeholder-10': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'ปัจจัยประกอบ กองทุนตราสารหนี้',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'master-placeholder-11': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'ปัจจัยประกอบอื่นๆ 4',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'master-placeholder-12': {
      sheetId: CONFIG.SHEETS?.RAW_FOR_SEC || '',
      tabName: '2026-Q1',
      title: 'เปรียบเทียบค่าธรรมเนียม',
      source: 'Data For SEC API + AVP Master Fund ID',
      localFile: 'Data/Data For SEC API - 2026-Q1.json',
      datasetKey: 'secApi',
    },
    'master-placeholder-7': {
      sheetId: CONFIG.SHEETS?.MASTER_FUND_ID || '',
      tabName: '2026-Q1',
      title: 'Top 10 Holding V2',
      source: 'AVP Master Fund ID',
      localFile: 'Data/AVP Master Fund ID - 2026-Q1.json',
      datasetKey: 'masterFund',
    },
    'income-fund-1': {
      title: 'Income Fund',
      source: 'Data For SEC API + Fund Key Performance AVP',
    },
    'income-fund-2': {
      title: 'Income Fund 2',
      source: 'Data For SEC API + Fund Key Performance AVP',
    },
    'robustness-ft-import': {
      title: 'เตรียมข้อมูลจาก FT.com',
      source: 'FT Markets historical prices',
    },
    'upside-downside-capture': {
      title: 'Upside Downside Capture',
      source: 'FT Markets + Fund performance data',
    },
    'fund-list-update': {
      sheetId: CONFIG.SHEETS?.FUND_LIST_UPDATE || '',
      tabName: 'fund_list_changes',
      title: 'Fund List Update',
      source: 'Google Sheet: fund_list_changes',
    },
    'fund-selection-logs': {
      tabName: '2026-Q1',
      title: 'Fund Selection Logs',
      source: 'Google Drive JSON: Fund Selection Logs',
    },
  };

  Object.entries(defaults).forEach(([pageKey, fallback]) => {
    CONFIG.PAGES[pageKey] = applyDatasetConfig(pageKey, {
      ...fallback,
      ...(CONFIG.PAGES[pageKey] || {}),
    });
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

async function fetchIncomeFundSourceRows(pageKey, fallbackPageKey) {
  ensureExtendedPageConfigs();
  const cfg = CONFIG.PAGES[fallbackPageKey];
  const quarter = State.currentQuarter || cfg?.tabName || '2026-Q1';
  const localFile = resolveQuarterLocalFile(cfg?.localFile, cfg?.tabName, quarter);
  const rows = await fetchLocalRows(localFile);
  const sourceName = DATASET_REGISTRY[cfg?.datasetKey]?.label || cfg?.source || fallbackPageKey;
  const sourceBadge = `Local JSON: ${sourceName} (${quarter})`;
  const current = State._pageDataSource[pageKey]
    ? State._pageDataSource[pageKey].split(' + ')
    : [];
  if (!current.includes(sourceBadge)) current.push(sourceBadge);
  State._pageDataSource[pageKey] = current.join(' + ');
  return rows;
}

async function fetchIncomeFundListRows(pageKey) {
  ensureExtendedPageConfigs();
  const quarter = State.currentQuarter || CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1';
  const localFile = `Data/Income Fund - ${quarter}.json`;
  const resp = await fetch(localFile, { cache: 'no-store' });
  if (!resp.ok) {
    throw new Error(`Local income fund list not found (${resp.status})`);
  }
  const payload = await resp.json();
  const items = Array.isArray(payload?.items) ? payload.items : [];
  if (!items.length) {
    throw new Error(`Invalid income fund list in ${localFile}`);
  }
  State._pageDataSource[pageKey] = `Local JSON: Income Fund (${payload.quarter || quarter})`;
  return items.map(item => ({
    name: String(item.name || '').trim(),
    policy: String(item.policy || '').trim(),
    source: String(item.source || '').trim(),
  })).filter(item => item.name && item.policy);
}

async function fetchIncomeFundUniverseRows(pageKey) {
  ensureExtendedPageConfigs();
  const rows = await fetchCached('select-fund');
  State._pageDataSource[pageKey] = State._pageDataSource['select-fund'] || 'Fund Key Performance AVP';
  return rows;
}

async function loadIncomeFundSecMetadata(force = false) {
  if (!force && State.incomeFundSecMetadataByCode) return State.incomeFundSecMetadataByCode;
  const map = new Map();
  const remember = (key, meta) => {
    const normalized = normalizeFundKey(key);
    if (!normalized || map.has(normalized)) return;
    map.set(normalized, {
      projId: String(meta?.projId || '').trim(),
      fundClassName: String(meta?.fundClassName || '').trim(),
      projAbbrName: String(meta?.projAbbrName || '').trim(),
      matchSource: String(meta?.matchSource || '').trim(),
    });
  };

  try {
    const resp = await fetch('outputs/dividend_by_proj_id_all.json', { cache: 'no-store' });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const payload = await resp.json();
    const items = Array.isArray(payload?.items) ? payload.items : [];
    items.forEach(item => {
      remember(item?.fund_code, {
        projId: String(item?.proj_id || '').trim(),
        fundClassName: String(item?.sec_fund_class_name || '').trim(),
        projAbbrName: String(item?.sec_proj_abbr_name || '').trim(),
        matchSource: String(item?.match_source || '').trim(),
      });
    });
  } catch (err) {
    console.warn('Income Fund SEC metadata output file unavailable:', err);
  }

  try {
    const secRows = await fetchCached('master-placeholder-4');
    const headers = secRows[0] || [];
    const idx = {
      projId: findColumnIndex(headers, ['proj_id', 'Project ID']),
      fundClassName: findColumnIndex(headers, ['fund_class_name', 'class_abbr_name', 'Fund Class Name']),
      projAbbrName: findColumnIndex(headers, ['proj_abbr_name', 'Project Abbr Name']),
    };

    secRows.slice(1).forEach(row => {
      const meta = {
        projId: rowValue(row, idx.projId),
        fundClassName: rowValue(row, idx.fundClassName),
        projAbbrName: rowValue(row, idx.projAbbrName),
        matchSource: 'Data For SEC API',
      };
      remember(meta.fundClassName, meta);
      remember(meta.projAbbrName, meta);
    });
  } catch (err) {
    console.warn('Income Fund SEC metadata Drive JSON unavailable:', err);
  }

  State.incomeFundSecMetadataByCode = map;
  return State.incomeFundSecMetadataByCode;
}

function incomeFundSecMetadata(code) {
  const key = normalizeFundKey(code);
  return State.incomeFundSecMetadataByCode?.get(key) || {
    projId: '',
    fundClassName: '',
    projAbbrName: '',
    matchSource: '',
  };
}

function incomeFundSelectionQuarter() {
  return State.currentQuarter || CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1';
}

function incomeFundPolicyLabel(value) {
  return value === 'Redemption' ? 'Auto Redeem' : value;
}

function incomeFundDividendApiUrl() {
  return String(CONFIG.INCOME_FUND_DIVIDEND_API_WEB_APP_URL || '').trim();
}

function incomeFundDividendApiKey() {
  return String(CONFIG.INCOME_FUND_DIVIDEND_API_SECRET_KEY || '').trim();
}

function truthySheetValue(value) {
  return /^(true|yes|y|1|selected|ติ๊ก|เลือก)$/i.test(String(value || '').trim());
}

async function loadIncomeFundDividendDatabase(force = false) {
  if (!force && State.incomeFundDividendDatabase) return State.incomeFundDividendDatabase;
  const apiUrl = incomeFundDividendApiUrl();
  const apiKey = incomeFundDividendApiKey();
  const configuredUrl = String(CONFIG.INCOME_FUND_DIVIDEND_DB_URL || '').trim();
  const candidates = [...new Set([
    configuredUrl,
    apiUrl ? `${apiUrl}?action=database&fileName=${encodeURIComponent(INCOME_FUND_DIVIDEND_DB_FILE_NAME)}&folderId=${encodeURIComponent(INCOME_FUND_DIVIDEND_DRIVE_FOLDER_ID)}${apiKey ? `&key=${encodeURIComponent(apiKey)}` : ''}` : '',
    ...INCOME_FUND_DIVIDEND_DB_FALLBACK_FILES,
  ].filter(Boolean))];
  let lastError = null;

  for (const url of candidates) {
    try {
      const resp = await fetch(url, { cache: 'no-store' });
      if (!resp.ok) throw new Error(`${url} (${resp.status})`);
      const payload = await resp.json();
      if (!Array.isArray(payload?.dividend_history)) {
        throw new Error(`${url} format ไม่ถูกต้อง`);
      }
      State.incomeFundDividendDatabase = payload;
      return payload;
    } catch (err) {
      lastError = err;
      console.warn('Income Fund dividend database unavailable:', err);
    }
  }

  throw new Error(`โหลด Dividend History Database ไม่ได้: ${lastError?.message || 'ไม่พบไฟล์'}`);
}

async function triggerIncomeFundDividendSync(funds = [], options = {}) {
  const apiUrl = incomeFundDividendApiUrl();
  if (!apiUrl) {
    throw new Error('ยังไม่ได้ตั้งค่า CONFIG.INCOME_FUND_DIVIDEND_API_WEB_APP_URL');
  }
  const url = new URL(apiUrl);
  url.searchParams.set('action', 'sync');
  url.searchParams.set('quarter', incomeFundSelectionQuarter());
  url.searchParams.set('folderId', INCOME_FUND_DIVIDEND_DRIVE_FOLDER_ID);
  url.searchParams.set('fileName', INCOME_FUND_DIVIDEND_DB_FILE_NAME);
  if (incomeFundDividendApiKey()) url.searchParams.set('key', incomeFundDividendApiKey());

  const payload = {
    quarter: incomeFundSelectionQuarter(),
    folderId: INCOME_FUND_DIVIDEND_DRIVE_FOLDER_ID,
    fileName: INCOME_FUND_DIVIDEND_DB_FILE_NAME,
    githubToken: String(options.githubToken || '').trim(),
    funds: funds.map(fund => ({
      code: fund.code,
      projId: fund.projId,
      policy: fund.policy,
    })).filter(fund => fund.code || fund.projId),
  };

  const resp = await fetch(url.toString(), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload),
    cache: 'no-store',
    redirect: 'follow',
  });
  const text = await resp.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch {
    throw new Error(text || `Dividend sync API HTTP ${resp.status}`);
  }
  if (!resp.ok || data.ok === false) {
    throw new Error(data.error || data.message || `Dividend sync API HTTP ${resp.status}`);
  }
  return data;
}

function ftHistoricalApiUrl() {
  return String(CONFIG.FT_HISTORICAL_API_WEB_APP_URL || FT_HISTORICAL_API_WEB_APP_URL || '').trim();
}

function ftHistoricalApiKey() {
  return String(CONFIG.FT_HISTORICAL_API_SECRET_KEY || FT_HISTORICAL_API_SECRET_KEY || '').trim();
}

async function triggerFtHistoricalSync(payload = {}) {
  const apiUrl = ftHistoricalApiUrl();
  if (!apiUrl) {
    throw new Error('ยังไม่ได้ตั้งค่า CONFIG.FT_HISTORICAL_API_WEB_APP_URL');
  }
  const url = new URL(apiUrl);
  url.searchParams.set('action', 'sync');
  url.searchParams.set('folderId', FT_HISTORICAL_DRIVE_FOLDER_ID);
  url.searchParams.set('fileName', FT_HISTORICAL_DB_FILE_NAME);
  if (ftHistoricalApiKey()) url.searchParams.set('key', ftHistoricalApiKey());

  const body = {
    action: 'sync',
    key: ftHistoricalApiKey(),
    quarter: State.currentQuarter || '2026-Q1',
    folderId: FT_HISTORICAL_DRIVE_FOLDER_ID,
    fileName: FT_HISTORICAL_DB_FILE_NAME,
    url: String(payload.url || '').trim(),
    symbol: String(payload.symbol || '').trim(),
    startDate: String(payload.startDate || '').trim(),
    endDate: String(payload.endDate || '').trim(),
    runPrices: payload.runPrices === false ? 'false' : 'true',
    runQualitative: payload.runQualitative === false ? 'false' : 'true',
    githubToken: String(payload.githubToken || '').trim(),
  };

  const resp = await fetch(url.toString(), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body),
    cache: 'no-store',
    redirect: 'follow',
  });
  const text = await resp.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch {
    throw new Error(text || `FT sync API HTTP ${resp.status}`);
  }
  if (!resp.ok || data.ok === false) {
    throw new Error(data.error || data.message || `FT sync API HTTP ${resp.status}`);
  }
  return data;
}

function ftSymbolBase(value) {
  return String(value || '').trim().split(':')[0].toUpperCase();
}

function ftSymbolSlug(value) {
  return String(value || '').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '').toUpperCase();
}

function normalizeFtDatabasePayload(payload = {}) {
  const prices = Array.isArray(payload.prices) ? payload.prices : [];
  const selectedRows = Array.isArray(payload.rows) ? payload.rows : [];
  return {
    ...payload,
    prices: prices.length ? prices : selectedRows,
    rows: selectedRows.length ? selectedRows : prices,
    profile: Array.isArray(payload.profile) ? payload.profile : [],
    risk: Array.isArray(payload.risk) ? payload.risk : [],
    holdings: Array.isArray(payload.holdings) ? payload.holdings : [],
    symbols: Array.isArray(payload.symbols) ? payload.symbols : [],
    source: payload.source || `Google Drive: ${FT_HISTORICAL_DB_FILE_NAME}`,
  };
}

async function loadFtHistoricalDatabase(force = false) {
  if (!force && State.ftHistoricalDatabase) return State.ftHistoricalDatabase;
  const apiUrl = ftHistoricalApiUrl();
  if (!apiUrl) throw new Error('ยังไม่ได้ตั้งค่า CONFIG.FT_HISTORICAL_API_WEB_APP_URL');
  const url = new URL(apiUrl);
  url.searchParams.set('action', 'database');
  url.searchParams.set('folderId', FT_HISTORICAL_DRIVE_FOLDER_ID);
  url.searchParams.set('fileName', FT_HISTORICAL_DB_FILE_NAME);
  if (ftHistoricalApiKey()) url.searchParams.set('key', ftHistoricalApiKey());
  const resp = await fetch(url.toString(), { cache: 'no-store', redirect: 'follow' });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok || data.ok === false) {
    throw new Error(data.error || data.message || `โหลด FT database จาก Drive ไม่สำเร็จ (${resp.status})`);
  }
  const payload = normalizeFtDatabasePayload(data);
  State.ftHistoricalDatabase = payload;
  return payload;
}

async function loadFtHistoricalPayloadForPage(limit = 5000) {
  if (ftHistoricalApiUrl()) {
    return loadFtHistoricalDatabase();
  }
  const resp = await fetch(`/api/ft-historical-prices?limit=${encodeURIComponent(limit)}`, { cache: 'no-store' });
  const payload = await resp.json().catch(() => ({}));
  if (resp.status === 404) {
    throw new Error('ไม่พบ API /api/ft-historical-prices และยังไม่ได้ตั้งค่า Apps Script สำหรับ FT');
  }
  if (!resp.ok || payload.ok === false) throw new Error(payload.error || `โหลดข้อมูลไม่สำเร็จ (${resp.status})`);
  return normalizeFtDatabasePayload(payload);
}

function findFtSymbolMatch(payload, lookupValue) {
  const value = String(lookupValue || '').trim();
  if (!value) return '';
  const upper = value.toUpperCase();
  const base = ftSymbolBase(value);
  const slug = ftSymbolSlug(value);
  const symbols = payload?.symbols || [];
  const match = symbols.find(item => {
    const symbol = String(item.symbol || '').trim();
    return symbol.toUpperCase() === upper
      || ftSymbolBase(symbol) === base
      || ftSymbolSlug(symbol) === slug;
  });
  return match?.symbol || '';
}

function getFtQualitativeFromPayload(payload, lookupValue) {
  const symbol = findFtSymbolMatch(payload, lookupValue);
  if (!symbol) return null;
  const profile = (payload.profile || []).filter(row => String(row.symbol || '').trim().toUpperCase() === symbol.toUpperCase());
  const risk = (payload.risk || []).filter(row => String(row.symbol || '').trim().toUpperCase() === symbol.toUpperCase());
  const holdings = (payload.holdings || []).filter(row => String(row.symbol || '').trim().toUpperCase() === symbol.toUpperCase());
  const profileMap = Object.fromEntries(profile.map(row => [row.field, row.value]));
  const meta = (payload.symbols || []).find(item => String(item.symbol || '').trim().toUpperCase() === symbol.toUpperCase()) || {};
  return {
    symbol,
    displayName: meta.displayName || profileMap['FT display name'] || '',
    profile,
    profileMap,
    risk,
    holdings,
  };
}

function compactIsoDate(value) {
  const text = String(value || '').trim();
  return text ? text.slice(0, 10) : '';
}

function formatDividendNumber(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return '-';
  return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 4 });
}

function sourcePriority(source) {
  return String(source || '').toLowerCase() === 'sec' ? 1 : 2;
}

function dedupeDividendHistory(rows = []) {
  const byKey = new Map();
  rows.forEach(row => {
    const key = [
      compactIsoDate(row.book_close_date),
      compactIsoDate(row.dividend_date),
      String(Number(row.dividend_value || 0)),
    ].join('|');
    const existing = byKey.get(key);
    if (!existing || sourcePriority(row.source) < sourcePriority(existing.source)) {
      byKey.set(key, {
        ...row,
        book_close_date: compactIsoDate(row.book_close_date),
        dividend_date: compactIsoDate(row.dividend_date),
        dividend_value: Number(row.dividend_value || 0),
        sources: existing?.sources || new Set(),
      });
    }
    const target = byKey.get(key);
    target.sources = target.sources || new Set();
    if (row.source) target.sources.add(row.source);
  });
  return [...byKey.values()]
    .map(row => ({ ...row, sources: [...(row.sources || [])].sort().join(' + ') || row.source || '-' }))
    .sort((a, b) => String(b.dividend_date).localeCompare(String(a.dividend_date)));
}

function buildIncomeDividendIndex(database) {
  const map = new Map();
  (database?.dividend_history || []).forEach(row => {
    const code = normalizeFundKey(row.fund_code || row.class_abbr_name);
    if (!code) return;
    if (!map.has(code)) map.set(code, []);
    map.get(code).push(row);
  });
  return map;
}

function summarizeIncomeDividendRows(rows = []) {
  const history = dedupeDividendHistory(rows);
  const latest = history[0] || null;
  const total = history.reduce((sum, row) => sum + (Number(row.dividend_value) || 0), 0);
  const sourceSet = new Set();
  rows.forEach(row => { if (row.source) sourceSet.add(row.source); });
  return {
    history,
    latest,
    count: history.length,
    total,
    average: history.length ? total / history.length : 0,
    sources: [...sourceSet].sort().join(' + ') || '-',
    secRows: rows.filter(row => String(row.source || '').toLowerCase() === 'sec').length,
    finnomenaRows: rows.filter(row => String(row.source || '').toLowerCase() === 'finnomena').length,
  };
}

function buildIncomeDividendDetailHtml(row) {
  const history = Array.isArray(row?.history) ? row.history : [];
  const byYear = new Map();
  history.forEach(item => {
    const year = String(item.dividend_date || '').slice(0, 4) || '-';
    if (!byYear.has(year)) byYear.set(year, { year, count: 0, total: 0, latestDate: '' });
    const bucket = byYear.get(year);
    bucket.count += 1;
    bucket.total += Number(item.dividend_value || 0);
    if (!bucket.latestDate || String(item.dividend_date || '') > bucket.latestDate) bucket.latestDate = item.dividend_date || '';
  });
  const yearRows = [...byYear.values()]
    .sort((a, b) => String(b.year).localeCompare(String(a.year)))
    .map(item => `
      <tr>
        <td>${esc(item.year)}</td>
        <td class="report-num">${item.count.toLocaleString()}</td>
        <td class="report-num">${esc(formatDividendNumber(item.total))}</td>
        <td class="report-num">${esc(formatDividendNumber(item.count ? item.total / item.count : 0))}</td>
        <td class="report-num">${esc(item.latestDate || '-')}</td>
      </tr>
    `).join('');
  const detailRows = history.map(item => `
    <tr>
      <td class="report-num">${esc(item.book_close_date || '-')}</td>
      <td class="report-num">${esc(item.dividend_date || '-')}</td>
      <td class="report-num">${esc(formatDividendNumber(item.dividend_value))}</td>
      <td>${esc(item.sources || item.source || '-')}</td>
    </tr>
  `).join('');
  return `
    <div class="income-detail-modal">
      <div class="sf-meta" style="margin-bottom:14px">
        <span class="row-count-badge is-info">${esc(row.code || '-')}</span>
        <span class="row-count-badge">${esc(row.policy || '-')}</span>
        <span class="row-count-badge">${history.length.toLocaleString()} ครั้ง</span>
        <span class="row-count-badge">รวม ${esc(formatDividendNumber(row.total))}</span>
        <span class="badge badge-data-origin">${esc(row.sources || '-')}</span>
      </div>
      <div style="margin-bottom:14px;color:var(--muted);font-size:0.92rem">
        ${esc(row.name || '-')} ${row.projId ? `· proj_id ${esc(row.projId)}` : ''}
      </div>
      <h4 style="margin:0 0 8px;color:var(--primary);font-size:1rem">สรุปรายปี</h4>
      <div class="table-wrapper" style="margin-bottom:16px">
        <table>
          <thead><tr>
            <th>ปี</th>
            <th>จำนวนครั้ง</th>
            <th>รวม</th>
            <th>เฉลี่ย</th>
            <th>วันที่ล่าสุดในปี</th>
          </tr></thead>
          <tbody>${yearRows || '<tr><td colspan="5" class="empty-cell">ยังไม่มีประวัติ</td></tr>'}</tbody>
        </table>
      </div>
      <h4 style="margin:0 0 8px;color:var(--primary);font-size:1rem">ประวัติการปันผลทั้งหมด</h4>
      <div class="table-wrapper">
        <table>
          <thead><tr>
            <th>วันที่ปิดสมุด/XD</th>
            <th>วันที่จ่ายปันผล</th>
            <th>จำนวนเงิน</th>
            <th>Source</th>
          </tr></thead>
          <tbody>${detailRows || '<tr><td colspan="4" class="empty-cell">ยังไม่มีประวัติ</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `;
}

async function ensureIncomeFundSelectionSheet() {
  const meta = await SheetsAPI.getSheetTabs(INCOME_FUND_SELECTION_SHEET_ID);
  if (!meta.tabs.includes(INCOME_FUND_SELECTION_TAB)) {
    await SheetsAPI.addSheetTab(INCOME_FUND_SELECTION_SHEET_ID, INCOME_FUND_SELECTION_TAB);
    await SheetsAPI.updateSheetValues(INCOME_FUND_SELECTION_SHEET_ID, INCOME_FUND_SELECTION_TAB, [INCOME_FUND_SELECTION_HEADERS]);
    return [INCOME_FUND_SELECTION_HEADERS];
  }
  const rows = await SheetsAPI.fetchSheetData(INCOME_FUND_SELECTION_SHEET_ID, INCOME_FUND_SELECTION_TAB);
  if (!rows.length) {
    await SheetsAPI.updateSheetValues(INCOME_FUND_SELECTION_SHEET_ID, INCOME_FUND_SELECTION_TAB, [INCOME_FUND_SELECTION_HEADERS]);
    return [INCOME_FUND_SELECTION_HEADERS];
  }
  return rows;
}

async function loadIncomeFundSelection(force = false) {
  const quarter = incomeFundSelectionQuarter();
  if (!force && State.incomeFundSelectionLoaded && State.incomeFundSelectionQuarter === quarter) {
    return State.incomeFundSelectedKeys;
  }
  const rows = await ensureIncomeFundSelectionSheet();
  const headers = rows[0] || INCOME_FUND_SELECTION_HEADERS;
  const ci = {
    quarter: findColumnIndex(headers, ['Quarter']),
    code: findColumnIndex(headers, ['Fund Code', 'Code']),
    selected: findColumnIndex(headers, ['Selected']),
  };
  const selected = new Set();
  rows.slice(1).forEach(row => {
    const rowQuarter = rowValue(row, ci.quarter);
    const code = normalizeFundKey(rowValue(row, ci.code));
    if (rowQuarter === quarter && code && truthySheetValue(rowValue(row, ci.selected))) {
      selected.add(code);
    }
  });
  State.incomeFundSelectedKeys = selected;
  State.incomeFundSelectionRows = rows;
  State.incomeFundSelectionLoaded = true;
  State.incomeFundSelectionQuarter = quarter;
  return selected;
}

async function saveIncomeFundSelection() {
  const quarter = incomeFundSelectionQuarter();
  const updatedAt = new Date().toISOString();
  const updatedBy = State.currentUser?.email || '';
  const existingRows = State.incomeFundSelectionRows?.length
    ? State.incomeFundSelectionRows
    : [INCOME_FUND_SELECTION_HEADERS];
  const headers = existingRows[0] || INCOME_FUND_SELECTION_HEADERS;
  const quarterIdx = findColumnIndex(headers, ['Quarter']);
  const preservedRows = existingRows.slice(1).filter(row => rowValue(row, quarterIdx) !== quarter);
  const selectedRows = [...State.incomeFundSelectedKeys].sort().map(code => [
    quarter,
    code,
    incomeFundSecMetadata(code).projId,
    incomeFundSecMetadata(code).fundClassName,
    incomeFundSecMetadata(code).projAbbrName,
    'TRUE',
    updatedAt,
    updatedBy,
  ]);
  const nextRows = [
    INCOME_FUND_SELECTION_HEADERS,
    ...preservedRows,
    ...selectedRows,
  ];
  State.incomeFundSelectionSaving = true;
  await SheetsAPI.clearSheetValues(INCOME_FUND_SELECTION_SHEET_ID, INCOME_FUND_SELECTION_TAB);
  await SheetsAPI.updateSheetValues(INCOME_FUND_SELECTION_SHEET_ID, INCOME_FUND_SELECTION_TAB, nextRows);
  State.incomeFundSelectionRows = nextRows;
  State.incomeFundSelectionLoaded = true;
  State.incomeFundSelectionSaving = false;
  return nextRows;
}

function debounceIncomeFundSelectionSave(renderStatus) {
  clearTimeout(State.incomeFundSelectionSaveTimer);
  State.incomeFundSelectionSaveTimer = setTimeout(async () => {
    try {
      State.incomeFundSelectionSaving = true;
      renderStatus?.();
      await saveIncomeFundSelection();
      toast('บันทึก Income Fund Selection แล้ว', 'success', 2200);
    } catch (err) {
      State.incomeFundSelectionSaving = false;
      toast(err.message || 'บันทึก Income Fund Selection ไม่สำเร็จ', 'error', 6000);
    } finally {
      State.incomeFundSelectionSaving = false;
      renderStatus?.();
    }
  }, 450);
}

function driveJsonYearFromQuarter(tabName) {
  const match = String(tabName || '').trim().match(/^(\d{4})-Q[1-4]$/i);
  return match?.[1] || '';
}

function buildDriveJsonFolderSegments(cfg) {
  const quarter = String(cfg.tabName || '').trim().toUpperCase();
  const year = driveJsonYearFromQuarter(quarter);
  return (cfg.driveJsonFolderSegments || ['{year}', '{quarter}', 'base']).map(segment => (
    String(segment)
      .replace(/\{year\}/g, year)
      .replace(/\{quarter\}/g, quarter)
  ));
}

async function fetchDriveJsonRows(cfg) {
  if (!cfg.driveJsonRootFolderId || !cfg.driveJsonFileName) {
    throw new Error('Drive JSON source is not configured');
  }
  const rows = await SheetsAPI.fetchDriveJsonRows(
    cfg.driveJsonRootFolderId,
    buildDriveJsonFolderSegments(cfg),
    cfg.driveJsonFileName
  );
  if (!Array.isArray(rows)) {
    throw new Error(`Invalid Drive JSON format in ${cfg.driveJsonFileName}`);
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
      if (cfgWithTab.preferDriveJson) {
        try {
          const rows = await fetchDriveJsonRows(cfgWithTab);
          State._pageDataSource[pageKey] = `Source: Google Drive JSON (${cfgWithTab.tabName})`;
          return rememberPageDataMeta(pageKey, rows, cfgWithTab);
        } catch {
          /* fall through to Google Sheets */
        }
      }
      const rows = await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
      State._pageDataSource[pageKey] = 'Source: Google Sheets fallback';
      return rememberPageDataMeta(pageKey, rows, cfgWithTab);
    }
  }

  if (mode === 'google_first') {
    try {
      if (cfgWithTab.preferDriveJson) {
        const rows = await fetchDriveJsonRows(cfgWithTab);
        State._pageDataSource[pageKey] = `Source: Google Drive JSON (${cfgWithTab.tabName})`;
        return rememberPageDataMeta(pageKey, rows, cfgWithTab);
      }
      const rows = await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
      State._pageDataSource[pageKey] = `Source: Google Sheets (${cfgWithTab.tabName})`;
      return rememberPageDataMeta(pageKey, rows, cfgWithTab);
    } catch (err) {
      if (cfgWithTab.preferDriveJson && cfgWithTab.sheetId) {
        try {
          const rows = await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
          State._pageDataSource[pageKey] = `Source: Google Sheets fallback (${cfgWithTab.tabName})`;
          return rememberPageDataMeta(pageKey, rows, cfgWithTab);
        } catch {
          /* fall through to Local JSON */
        }
      }
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

function isQuarterLabel(value) {
  return /^\d{4}-Q[1-4]$/i.test(String(value || '').trim());
}

function sortQuartersDesc(quarters) {
  return [...new Set((quarters || []).filter(isQuarterLabel))]
    .sort((a, b) => quarterSortValue(b) - quarterSortValue(a));
}

function intersectQuarterLists(lists) {
  const normalized = (lists || []).map(list => new Set(sortQuartersDesc(list)));
  if (!normalized.length) return [];
  const [first, ...rest] = normalized;
  return sortQuartersDesc([...first].filter(q => rest.every(set => set.has(q))));
}

async function detectSheetReadyQuarters() {
  const metas = await Promise.all(REQUIRED_QUARTER_DATASET_KEYS.map(async key => {
    const dataset = DATASET_REGISTRY[key];
    const meta = await SheetsAPI.getSheetTabs(dataset.sheetId);
    return (meta.tabs || []).filter(isQuarterLabel);
  }));
  return intersectQuarterLists(metas);
}

async function detectDriveJsonReadyQuarters() {
  const requiredFiles = REQUIRED_QUARTER_DATASET_KEYS
    .map(key => DATASET_REGISTRY[key]?.jsonFileName)
    .filter(Boolean);
  const yearFolders = (await SheetsAPI.listDriveFolderFiles(JSON_DRIVE_ROOT_FOLDER_ID))
    .filter(file => file.mimeType === 'application/vnd.google-apps.folder' && /^\d{4}$/.test(String(file.name || '')));

  const ready = [];
  for (const yearFolder of yearFolders) {
    const quarterFolders = (await SheetsAPI.listDriveFolderFiles(yearFolder.id))
      .filter(file => file.mimeType === 'application/vnd.google-apps.folder' && isQuarterLabel(file.name));
    for (const quarterFolder of quarterFolders) {
      const baseFolder = (await SheetsAPI.listDriveFolderFiles(quarterFolder.id))
        .find(file => file.mimeType === 'application/vnd.google-apps.folder' && String(file.name || '').trim() === 'base');
      if (!baseFolder) continue;
      const baseFiles = await SheetsAPI.listDriveFolderFiles(baseFolder.id);
      const fileNames = new Set(baseFiles.map(file => String(file.name || '').trim()));
      if (requiredFiles.every(name => fileNames.has(name))) {
        ready.push(String(quarterFolder.name || '').trim().toUpperCase());
      }
    }
  }
  return sortQuartersDesc(ready);
}

async function detectReadyQuarters() {
  const [sheetQuarters, driveQuarters] = await Promise.all([
    detectSheetReadyQuarters(),
    detectDriveJsonReadyQuarters(),
  ]);
  return intersectQuarterLists([sheetQuarters, driveQuarters]);
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
      if (cfgWithTab.preferDriveJson) {
        try {
          return await fetchDriveJsonRows(cfgWithTab);
        } catch {
          /* fall through to Google Sheets */
        }
      }
      return await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
    }
  }

  if (mode === 'google_first') {
    try {
      if (cfgWithTab.preferDriveJson) {
        return await fetchDriveJsonRows(cfgWithTab);
      }
      return await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
    } catch (err) {
      if (cfgWithTab.preferDriveJson && cfgWithTab.sheetId) {
        try {
          return await SheetsAPI.fetchSheetData(cfgWithTab.sheetId, cfgWithTab.tabName);
        } catch {
          /* fall through to Local JSON */
        }
      }
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

function rowValue(row, index) {
  return index >= 0 ? String(row?.[index] ?? '').trim() : '';
}

function buildIncomeFundRows(secRows = [], keyPerformanceRows = []) {
  const secHeaders = secRows[0] || [];
  const secIdx = {
    nameTh: findColumnIndex(secHeaders, ['proj_name_th']),
    nameEn: findColumnIndex(secHeaders, ['proj_name_en']),
    className: findColumnIndex(secHeaders, ['fund_class_name']),
    abbr: findColumnIndex(secHeaders, ['proj_abbr_name']),
    policy: findColumnIndex(secHeaders, ['dividend_policy']),
  };

  const keyHeaders = keyPerformanceRows[0] || [];
  const keyIdx = {
    name: findColumnIndex(keyHeaders, ['Name', 'Fund Name', 'FundName']),
    code: findColumnIndex(keyHeaders, ['Fund Code', 'Code']),
    policy: findColumnIndex(keyHeaders, ['Dividend', 'Div']),
  };

  const items = [];
  secRows.slice(1).forEach(row => {
    if (rowValue(row, secIdx.policy).toUpperCase() !== 'Y') return;
    const baseName = rowValue(row, secIdx.nameTh)
      || rowValue(row, secIdx.nameEn)
      || rowValue(row, secIdx.abbr)
      || rowValue(row, secIdx.className);
    if (!baseName) return;
    const className = rowValue(row, secIdx.className);
    const abbr = rowValue(row, secIdx.abbr);
    const classSuffix = className && className.toLowerCase() !== 'main' && className !== abbr
      ? ` (${className})`
      : '';
    items.push({
      name: `${baseName}${classSuffix}`,
      policy: 'Dividend',
      source: 'Data For SEC API',
    });
  });

  keyPerformanceRows.slice(1).forEach(row => {
    const policy = rowValue(row, keyIdx.policy);
    if (policy !== 'Dividend' && policy !== 'Redemption') return;
    const name = rowValue(row, keyIdx.name) || rowValue(row, keyIdx.code);
    if (!name) return;
    items.push({
      name,
      policy,
      source: 'Fund Key Performance AVP',
    });
  });

  const seen = new Set();
  return items
    .sort((a, b) => a.name.localeCompare(b.name, 'th') || a.policy.localeCompare(b.policy, 'en'))
    .filter(item => {
      const key = `${normalizeFundKey(item.name)}::${item.policy}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function normalizeFundOverrideItem(item) {
  const code = String(item?.code || item?.key || '').trim().toUpperCase();
  if (!code) return null;
  return {
    key: code,
    code,
    name: String(item.name || '').trim(),
    category: String(item.category || '').trim(),
    type: String(item.type || '').trim(),
    dividend: String(item.dividend || '').trim(),
    style: String(item.style || '').trim(),
    masterId: String(item.masterId || '').trim(),
    masterName: String(item.masterName || '').trim(),
    assetHouse: String(item.assetHouse || '').trim(),
    note: String(item.note || '').trim(),
    status: String(item.status || 'Active Override').trim(),
    mode: String(item.mode || 'edited').trim(),
    createdAt: item.createdAt || '',
    updatedAt: item.updatedAt || '',
  };
}

function applyFundOverrides(catalog, overrides = State.fundOverrides?.items || {}) {
  const map = new Map((catalog || []).map(fund => [fund.key, { ...fund, isOverride: false }]));
  Object.values(overrides || {}).forEach(raw => {
    const override = normalizeFundOverrideItem(raw);
    if (!override) return;
    const base = map.get(override.key) || {
      key: override.key,
      code: override.code,
      name: '',
      category: '',
      type: deriveFundType(override.code),
      dividend: deriveDividend(override.code),
      style: 'Active',
      masterId: '',
      masterName: '',
      assetHouse: '',
    };
    map.set(override.key, {
      ...base,
      ...Object.fromEntries(Object.entries(override).filter(([, value]) => value !== '')),
      isOverride: true,
      override,
    });
  });
  return [...map.values()];
}

async function loadFundOverrides(force = false) {
  if (!force && State.fundOverridesLoaded) return State.fundOverrides;
  try {
    const resp = await fetch('/api/fund-overrides', { cache: 'no-store' });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    State.fundOverrides = {
      items: data.items || {},
      updatedAt: data.updatedAt || null,
      source: data.source || 'Data/fund_overrides.json',
    };
    State.fundOverridesLoaded = true;
  } catch (err) {
    State.fundOverrides = State.fundOverrides || { items: {}, updatedAt: null };
    State.fundOverridesLoadError = err.message || String(err);
  }
  return State.fundOverrides;
}

async function saveFundOverride(fund) {
  const resp = await fetch('/api/fund-overrides', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fund }),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok || data.ok === false) {
    throw new Error(data.error || `บันทึกไม่สำเร็จ (${resp.status})`);
  }
  State.fundOverridesLoaded = false;
  await loadFundOverrides(true);
  return data.fund;
}

async function deleteFundOverride(key) {
  const resp = await fetch(`/api/fund-overrides/${encodeURIComponent(key)}`, { method: 'DELETE' });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok || data.ok === false) {
    throw new Error(data.error || `ลบไม่สำเร็จ (${resp.status})`);
  }
  State.fundOverridesLoaded = false;
  await loadFundOverrides(true);
  return data;
}

async function loadMasterAllocations(force = false) {
  if (!force && State.masterAllocationsLoaded) return State.masterAllocations;
  const quarter = State.currentQuarter || CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1';
  try {
    const resp = await fetch(`/api/master-allocations?quarter=${encodeURIComponent(quarter)}`, { cache: 'no-store' });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    State.masterAllocations = {
      items: data.items || {},
      updatedAt: data.updatedAt || null,
      source: data.source || 'Data/fund_master_allocations.json',
    };
    State.masterAllocationsLoaded = true;
    State.masterAllocationsLoadError = '';
  } catch (err) {
    try {
      State.masterAllocations = await loadMasterAllocationsStaticOrRemote(quarter);
      State.masterAllocationsLoaded = true;
      State.masterAllocationsLoadError = '';
    } catch (fallbackErr) {
      State.masterAllocations = State.masterAllocations || { items: {}, updatedAt: null };
      State.masterAllocationsLoadError = fallbackErr.message || err.message || String(err);
    }
  }
  return State.masterAllocations;
}

async function saveMasterAllocation(item) {
  let data;
  try {
    const resp = await fetch('/api/master-allocations', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(item),
    });
    data = await resp.json().catch(() => ({}));
    if (!resp.ok || data.ok === false) {
      throw new Error(data.error || `บันทึก mapping ไม่สำเร็จ (${resp.status})`);
    }
  } catch (err) {
    if (!hasMasterAllocationsApi()) {
      throw new Error(`${err.message || err} · GitHub Pages ต้องตั้งค่า MASTER_ALLOCATIONS_API_WEB_APP_URL ก่อนบันทึก Mapping`);
    }
    data = await masterAllocationsApiRequest('save', { item });
  }
  State.masterAllocationsLoaded = false;
  await loadMasterAllocations(true);
  return data;
}

async function deleteMasterAllocation(key) {
  const quarter = State.currentQuarter || CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1';
  let data;
  try {
    const resp = await fetch(`/api/master-allocations/${encodeURIComponent(key)}?quarter=${encodeURIComponent(quarter)}`, { method: 'DELETE' });
    data = await resp.json().catch(() => ({}));
    if (!resp.ok || data.ok === false) {
      throw new Error(data.error || `ลบ mapping ไม่สำเร็จ (${resp.status})`);
    }
  } catch (err) {
    if (!hasMasterAllocationsApi()) {
      throw new Error(`${err.message || err} · GitHub Pages ต้องตั้งค่า MASTER_ALLOCATIONS_API_WEB_APP_URL ก่อนลบ Mapping`);
    }
    data = await masterAllocationsApiRequest('delete', { quarter, key });
  }
  State.masterAllocationsLoaded = false;
  await loadMasterAllocations(true);
  return data;
}

function hasMasterAllocationsApi() {
  return Boolean(masterAllocationsApiUrl());
}

function masterAllocationsApiUrl() {
  return String(CONFIG.MASTER_ALLOCATIONS_API_WEB_APP_URL || MASTER_ALLOCATIONS_API_WEB_APP_URL || '').trim();
}

function masterAllocationsApiKey() {
  return String(CONFIG.MASTER_ALLOCATIONS_API_SECRET_KEY || MASTER_ALLOCATIONS_API_SECRET_KEY || '').trim();
}

function masterAllocationsFileForQuarter(quarter) {
  const normalized = String(quarter || '2026-Q1').trim().toUpperCase();
  const year = /^\d{4}-Q[1-4]$/.test(normalized) ? normalized.slice(0, 4) : '2026';
  return `Data/${year}/${normalized}/overrides/fund_master_allocations.json`;
}

async function loadMasterAllocationsStaticOrRemote(quarter) {
  if (hasMasterAllocationsApi()) {
    const data = await masterAllocationsApiRequest('get', { quarter });
    return {
      items: data.items || data.data?.items || {},
      updatedAt: data.updatedAt || data.data?.updatedAt || null,
      source: data.source || data.drive?.fileName || 'Google Drive',
    };
  }

  const path = masterAllocationsFileForQuarter(quarter);
  const resp = await fetch(path, { cache: 'no-store' });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const data = await resp.json();
  return {
    items: data.items || {},
    updatedAt: data.updatedAt || null,
    source: path,
  };
}

function masterAllocationsApiJsonp(params) {
  return new Promise((resolve, reject) => {
    const callbackName = `__masterAllocApiCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement('script');
    const cleanup = () => {
      delete window[callbackName];
      script.remove();
    };
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('Master Allocations API request timeout'));
    }, 30000);
    window[callbackName] = (data) => {
      clearTimeout(timer);
      cleanup();
      resolve(data || {});
    };
    const url = new URL(masterAllocationsApiUrl());
    Object.entries({
      key: masterAllocationsApiKey(),
      callback: callbackName,
      ...params,
    }).forEach(([paramKey, value]) => {
      if (value !== undefined && value !== null && value !== '') {
        url.searchParams.set(paramKey, value);
      }
    });
    script.onerror = () => {
      clearTimeout(timer);
      cleanup();
      reject(new Error('Master Allocations API script failed to load'));
    };
    script.src = url.toString();
    document.head.appendChild(script);
  });
}

async function masterAllocationsApiFetch(params) {
  const url = new URL(masterAllocationsApiUrl());
  Object.entries({
    key: masterAllocationsApiKey(),
    ...params,
  }).forEach(([paramKey, value]) => {
    if (value !== undefined && value !== null && value !== '') {
      url.searchParams.set(paramKey, value);
    }
  });
  const res = await fetch(url.toString(), { cache: 'no-store', redirect: 'follow' });
  const text = await res.text();
  if (!res.ok) throw new Error(`Master Allocations API HTTP ${res.status}`);
  return JSON.parse(text);
}

async function masterAllocationsCompressedPayload(payload) {
  const text = JSON.stringify(payload || {});
  if (!('CompressionStream' in window)) return { payload: text };
  const stream = new Blob([text], { type: 'application/json' }).stream().pipeThrough(new CompressionStream('gzip'));
  const buffer = await new Response(stream).arrayBuffer();
  let binary = '';
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return {
    payloadGzipB64: btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, ''),
  };
}

async function masterAllocationsApiRequest(action, payload = {}) {
  const quarter = payload.quarter || State.currentQuarter || CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1';
  const params = { action, quarter };
  if (action === 'save') {
    Object.assign(params, await masterAllocationsCompressedPayload({
      action,
      key: masterAllocationsApiKey(),
      item: payload.item,
    }));
  } else if (action === 'delete') {
    params.keyToDelete = payload.key;
  }

  let data;
  try {
    data = await masterAllocationsApiFetch(params);
  } catch {
    data = await masterAllocationsApiJsonp(params);
  }
  if (data.ok === false) {
    throw new Error(data.error || `Master Allocations API ${action} failed`);
  }
  return data;
}

async function loadFixedIncomeFactorsOverrides(force = false) {
  if (!force && State.fixedIncomeFactorsOverridesLoaded) return State.fixedIncomeFactorsOverrides;
  const quarter = State.currentQuarter || CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1';
  try {
    const resp = await fetch('/api/fixed-income-factors-overrides', { cache: 'no-store' });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    State.fixedIncomeFactorsOverrides = {
      items: data.items || {},
      updatedAt: data.updatedAt || null,
      source: data.source || 'Data/fixed_income_factors_overrides.json',
    };
    State.fixedIncomeFactorsOverridesLoaded = true;
    State.fixedIncomeFactorsOverridesLoadError = '';
  } catch (err) {
    try {
      State.fixedIncomeFactorsOverrides = await loadFixedIncomeFactorsOverridesStaticOrRemote(quarter);
      State.fixedIncomeFactorsOverridesLoaded = true;
      State.fixedIncomeFactorsOverridesLoadError = '';
    } catch (fallbackErr) {
      State.fixedIncomeFactorsOverrides = State.fixedIncomeFactorsOverrides || { items: {}, updatedAt: null };
      State.fixedIncomeFactorsOverridesLoadError = fallbackErr.message || err.message || String(err);
    }
  }
  return State.fixedIncomeFactorsOverrides;
}

async function saveFixedIncomeFactorsOverrides(items) {
  let data;
  try {
    const resp = await fetch('/api/fixed-income-factors-overrides', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ items, replace: true }),
    });
    data = await resp.json().catch(() => ({}));
    if (!resp.ok || data.ok === false) {
      throw new Error(data.error || `บันทึก override ไม่สำเร็จ (${resp.status})`);
    }
  } catch (err) {
    if (!hasFixedIncomeFactorsApi()) {
      throw new Error(`${err.message || err} · GitHub Pages ต้องตั้งค่า FIXED_INCOME_FACTORS_API_WEB_APP_URL ก่อนบันทึก Override`);
    }
    data = await fixedIncomeFactorsApiRequest('save', { items, replace: true });
  }
  State.fixedIncomeFactorsOverridesLoaded = false;
  await loadFixedIncomeFactorsOverrides(true);
  return data.items;
}

function hasFixedIncomeFactorsApi() {
  return Boolean(fixedIncomeFactorsApiUrl());
}

function fixedIncomeFactorsApiUrl() {
  return String(CONFIG.FIXED_INCOME_FACTORS_API_WEB_APP_URL || FIXED_INCOME_FACTORS_API_WEB_APP_URL || '').trim();
}

function fixedIncomeFactorsApiKey() {
  return String(CONFIG.FIXED_INCOME_FACTORS_API_SECRET_KEY || FIXED_INCOME_FACTORS_API_SECRET_KEY || '').trim();
}

function fixedIncomeFactorsFileForQuarter(quarter) {
  const normalized = String(quarter || '2026-Q1').trim().toUpperCase();
  const year = /^\d{4}-Q[1-4]$/.test(normalized) ? normalized.slice(0, 4) : '2026';
  return `Data/${year}/${normalized}/overrides/fixed_income_factors_overrides.json`;
}

async function loadFixedIncomeFactorsOverridesStaticOrRemote(quarter) {
  if (hasFixedIncomeFactorsApi()) {
    const data = await fixedIncomeFactorsApiRequest('get', { quarter });
    return {
      items: data.items || data.data?.items || {},
      updatedAt: data.updatedAt || data.data?.updatedAt || null,
      source: data.source || data.drive?.fileName || 'Google Drive',
    };
  }

  const path = fixedIncomeFactorsFileForQuarter(quarter);
  const resp = await fetch(path, { cache: 'no-store' });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const data = await resp.json();
  return {
    items: data.items || {},
    updatedAt: data.updatedAt || null,
    source: path,
  };
}

function fixedIncomeFactorsApiJsonp(params) {
  return new Promise((resolve, reject) => {
    const callbackName = `__fixedIncomeFactorsApiCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement('script');
    const cleanup = () => {
      delete window[callbackName];
      script.remove();
    };
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('Fixed Income Factors API request timeout'));
    }, 30000);
    window[callbackName] = (data) => {
      clearTimeout(timer);
      cleanup();
      resolve(data || {});
    };
    const url = new URL(fixedIncomeFactorsApiUrl());
    Object.entries({
      key: fixedIncomeFactorsApiKey(),
      callback: callbackName,
      ...params,
    }).forEach(([paramKey, value]) => {
      if (value !== undefined && value !== null && value !== '') {
        url.searchParams.set(paramKey, value);
      }
    });
    script.onerror = () => {
      clearTimeout(timer);
      cleanup();
      reject(new Error('Fixed Income Factors API script failed to load'));
    };
    script.src = url.toString();
    document.head.appendChild(script);
  });
}

async function fixedIncomeFactorsApiFetch(params) {
  const url = new URL(fixedIncomeFactorsApiUrl());
  Object.entries({
    key: fixedIncomeFactorsApiKey(),
    ...params,
  }).forEach(([paramKey, value]) => {
    if (value !== undefined && value !== null && value !== '') {
      url.searchParams.set(paramKey, value);
    }
  });
  const res = await fetch(url.toString(), { cache: 'no-store', redirect: 'follow' });
  const text = await res.text();
  if (!res.ok) throw new Error(`Fixed Income Factors API HTTP ${res.status}`);
  return JSON.parse(text);
}

async function fixedIncomeFactorsCompressedPayload(payload) {
  const text = JSON.stringify(payload || {});
  if (!('CompressionStream' in window)) return { payload: text };
  const stream = new Blob([text], { type: 'application/json' }).stream().pipeThrough(new CompressionStream('gzip'));
  const buffer = await new Response(stream).arrayBuffer();
  let binary = '';
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return {
    payloadGzipB64: btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, ''),
  };
}

async function fixedIncomeFactorsApiRequest(action, payload = {}) {
  const quarter = payload.quarter || State.currentQuarter || CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1';
  const params = { action, quarter };
  if (action === 'save') {
    Object.assign(params, await fixedIncomeFactorsCompressedPayload({
      action,
      key: fixedIncomeFactorsApiKey(),
      items: payload.items,
      replace: payload.replace !== false,
    }));
  }

  let data;
  try {
    data = await fixedIncomeFactorsApiFetch(params);
  } catch {
    data = await fixedIncomeFactorsApiJsonp(params);
  }
  if (data.ok === false) {
    throw new Error(data.error || `Fixed Income Factors API ${action} failed`);
  }
  return data;
}

async function ensureSelectedFundsCatalog() {
  if (Object.keys(State.selectedFunds || {}).length) {
    return State.selectedFunds;
  }
  const rows = await fetchCached('select-fund');
  await loadFundOverrides();
  const allFunds = applyFundOverrides(buildSelectedFundsCatalog(rows));
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
    source: CONFIG.PAGES['select-fund']?.source || 'Fund Key Performance AVP',
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
    'FT Symbol',
    'FT Display Name',
    'Morningstar Category',
    'Investment Style',
    'Fund Size',
    'Share Class Size',
    'Ongoing Charge',
    'Initial Charge',
    'Max Annual Charge',
    'Exit Charge',
    'FT Sharpe 3Y',
    'FT Std Dev 3Y',
    'FT Alpha 3Y',
    'FT Beta 3Y',
    ...Array.from({ length: 10 }, (_, i) => `Holding #${i + 1}`),
  ];
  const dataRows = rows.map(fund => {
    const holdings = Array.isArray(fund?._topHoldings) ? fund._topHoldings : [];
    const profile = fund?._ftProfileMap || {};
    const firstProfile = (fields) => {
      for (const field of fields) {
        if (profile[field]) return profile[field];
      }
      return '';
    };
    const riskValue = (metric, period = '3 year') => {
      const found = (fund?._ftRisk || []).find(row => row.metric === metric && row.period === period);
      return found?.fundValue || '';
    };
    return [
      top10Text((fund?._selectorName || fund?.name || '').split('–')[0].trim()),
      top10Text(fund?._isin),
      top10Text(fund?._ftSymbol),
      top10Text(fund?._ftDisplayName),
      top10Text(profile['Morningstar category']),
      top10Text(profile['Investment style (stocks)']),
      top10Text(firstProfile(['Fund size', 'Total net assets'])),
      top10Text(firstProfile(['Share class size'])),
      top10Text(firstProfile(['Ongoing charge', 'Net expense ratio']) || fund?._feesRaw),
      top10Text(firstProfile(['Initial charge', 'Front end load']) || fund?._feesInitial),
      top10Text(firstProfile(['Max annual charge']) || fund?._feesMaxAnnual),
      top10Text(firstProfile(['Exit charge']) || fund?._feesExit),
      top10Text(riskValue('Sharpe ratio')),
      top10Text(riskValue('Standard deviation')),
      top10Text(riskValue('Alpha')),
      top10Text(riskValue('Beta')),
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
    source: CONFIG.PAGES['select-fund']?.source || 'Fund Key Performance AVP',
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
    ${pageToolActions(pageKey, CONFIG.PAGES['select-fund']?.source || 'Fund Key Performance AVP', toggleActions)}
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
    source: CONFIG.PAGES['select-fund']?.source || 'Fund Key Performance AVP',
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

function extractSecFeeValue(text, labels = []) {
  const raw = String(text || '').trim();
  if (!raw) return '';
  const parts = raw.split('|').map(part => part.trim()).filter(Boolean);
  const matched = parts.find(part => labels.some(label => part.toLowerCase().includes(String(label).toLowerCase())));
  const source = matched || raw;
  const afterColon = source.includes(':') ? source.split(':').slice(1).join(':') : source;
  const value = String(afterColon).replace(/,/g, '').match(/-?\d+(?:\.\d+)?/);
  return value ? value[0] : '';
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
    navFeesActual: findColumnIndex(headers, ['actual_value_fee_type_desc_nav']),
    buyFeesActual: findColumnIndex(headers, ['actual_value_fee_type_desc_buy']),
    feesLastUpdated: findColumnIndex(headers, ['fees_last_upd_date']),
    minimumSubIpo: findColumnIndex(headers, ['minimum_sub_ipo']),
    minimumSub: findColumnIndex(headers, ['minimum_sub']),
    recoveringPeriod: findColumnIndex(headers, ['recovering_period']),
    portfolioTurnoverRatio: findColumnIndex(headers, ['portfolio_turnover_ratio']),
    portfolioDurationPeriod: findColumnIndex(headers, ['portfolio_duration_period']),
    yieldToMaturity: findColumnIndex(headers, ['yield_to_maturity']),
  };
  const get = (row, index) => index >= 0 ? String(row[index] ?? '').trim() : '';
  const map = new Map();
  rows.slice(1).forEach(row => {
    const ter = get(row, ci.ter) || extractSecFeeValue(get(row, ci.navFeesActual), [
      'Total Fee and Expense',
      'ค่าใช้จ่ายรวมทั้งหมด',
    ]);
    const front = get(row, ci.front) || extractSecFeeValue(get(row, ci.buyFeesActual), [
      'Front-end Fee',
      'ค่าธรรมเนียมการขายหน่วยลงทุน',
    ]);
    const back = get(row, ci.back) || extractSecFeeValue(get(row, ci.buyFeesActual), [
      'Back-end Fee',
      'ค่าธรรมเนียมการรับซื้อคืนหน่วยลงทุน',
    ]);
    const record = {
      projName: get(row, ci.projName),
      className: get(row, ci.className),
      fundName: get(row, ci.fundName),
      ter,
      front,
      back,
      date: get(row, ci.date) || get(row, ci.feesLastUpdated),
      initial: get(row, ci.initial) || get(row, ci.minimumSubIpo),
      subsequent: get(row, ci.subsequent) || get(row, ci.minimumSub),
      fxHedging: get(row, ci.fxHedging),
      pdfFactsheet: get(row, ci.pdfFactsheet),
      asOfDate: get(row, ci.asOfDate) || get(row, ci.feesLastUpdated),
      recoveringPeriod: get(row, ci.recoveringPeriod),
      portfolioTurnoverRatio: get(row, ci.portfolioTurnoverRatio),
      portfolioDurationPeriod: get(row, ci.portfolioDurationPeriod),
      yieldToMaturity: get(row, ci.yieldToMaturity),
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

async function buildSelectedFeeUniverse() {
  const selectRows = await fetchCached('select-fund');
  await loadFundOverrides();
  const masterRows = buildMasterRecords(await fetchCached('master-placeholder-2'));
  return buildSelectedFeeUniverseFromRows(selectRows, masterRows);
}

function buildSelectedFeeUniverseFromRows(selectRows, masterRows) {
  const allFunds = applyFundOverrides(buildSelectedFundsCatalog(selectRows))
    .filter(f => f.code)
    .sort((a, b) => String(a.code).localeCompare(String(b.code), 'th'));

  if (!Object.keys(State.selectedFunds || {}).length) {
    State.selectedFunds = Object.fromEntries(allFunds.map(f => [f.key, f]));
  }

  const scopedFunds = State.selectedKeys.size > 0
    ? allFunds.filter(f => State.selectedKeys.has(f.key))
    : allFunds.slice(0, 18);

  const byIsin = {};
  masterRows.forEach(record => {
    const key = String(record.isin || '').trim();
    if (!key) return;
    if (!byIsin[key]) byIsin[key] = [];
    byIsin[key].push(record);
  });

  return scopedFunds.map(fund => {
    const exact = byIsin[String(fund.masterId || '').trim()] || [];
    const master = exact.length ? pickBestMasterRecord(exact) : null;
    return { fund, master };
  });
}

function buildFeeComparisonRows(universe, rawLookup, options = {}) {
  const includeThaiOnly = !!options.includeThaiOnly;
  const rows = universe.map(({ fund, master }) => {
    const raw = rawLookup.get(String(fund.code || '').trim().toUpperCase()) || null;
    const masterTer = parseNum(master?.ongoingCost);
    const thaiTer = parseNum(raw?.ter);
    const hasMasterTer = !Number.isNaN(masterTer);
    const hasThaiTer = !Number.isNaN(thaiTer);
    const combined = hasMasterTer && hasThaiTer
      ? masterTer + thaiTer
      : (includeThaiOnly && hasThaiTer ? thaiTer : NaN);
    const front = toFixedSafe(raw?.front, 4);
    const back = toFixedSafe(raw?.back, 4);
    const fxHedging = toFixedSafe(raw?.fxHedging, 2);
    return {
      thaiCode: fund.code,
      masterName: master?.name || fund.masterName || '-',
      masterTer,
      thaiTer,
      combined,
      feeDate: raw?.date || master?.ongoingCostDate || '',
      masterTerText: master ? toFixedSafe(master.ongoingCost) : '',
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
  }).filter(item => !Number.isNaN(item.thaiTer) && (includeThaiOnly || !Number.isNaN(item.masterTer)));
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

  async fundListUpdate(area) {
    const pageKey = 'fund-list-update';
    const cfg = CONFIG.PAGES?.[pageKey] || {};
    const sheetUrl = googleSheetUrl(cfg.sheetId);
    setLoading(area, 'กำลังโหลด Fund List Update...');

    const normalizeRows = (rows) => {
      if (!Array.isArray(rows) || !rows.length) return [];
      if (Array.isArray(rows[0])) {
        const headers = rows[0].map(value => String(value || '').trim());
        return rows.slice(1).filter(row => row.some(cell => String(cell || '').trim())).map((row, idx) => {
          const item = {};
          headers.forEach((header, colIdx) => {
            if (header) item[header] = row[colIdx] ?? '';
          });
          if (!item.id) item.id = `FLU-${idx + 1}`;
          return item;
        });
      }
      return rows;
    };

    const cleanRows = (items) => (Array.isArray(items) ? items : [])
      .map((row, idx) => ({
        id: String(row.id || `FLU-${Date.now()}-${idx + 1}`).trim(),
        list_from: String(row.list_from || row.quarter_from || '').trim(),
        list_to: String(row.list_to || row.quarter_to || '').trim(),
        type: String(row.type || '').trim(),
        asset_class: String(row.asset_class || '').trim(),
        fund_list_old: String(row.fund_list_old || row.old_fund || '').trim(),
        fund_list_current: String(row.fund_list_current || row.new_fund || '').trim(),
        change_type: String(row.change_type || '').trim(),
        note: String(row.note || '').trim(),
        updated_at: String(row.updated_at || '').trim(),
        updated_by: String(row.updated_by || State.currentUser?.email || '').trim(),
      }));

    const readDraftRows = () => {
      try {
        const payload = JSON.parse(localStorage.getItem(FUND_LIST_UPDATE_DRAFT_KEY) || 'null');
        return cleanRows(payload?.rows || []);
      } catch {
        return [];
      }
    };

    const writeDraftRows = () => {
      const now = new Date().toISOString();
      rows = cleanRows(rows).map(row => ({
        ...row,
        updated_at: row.updated_at || now,
        updated_by: row.updated_by || State.currentUser?.email || '',
      }));
      localStorage.setItem(FUND_LIST_UPDATE_DRAFT_KEY, JSON.stringify({
        updatedAt: now,
        rows,
      }));
      return now;
    };

    const rowsToSheetValues = () => [
      FUND_LIST_UPDATE_COLUMNS,
      ...cleanRows(rows).map(row => FUND_LIST_UPDATE_COLUMNS.map(column => row[column] || '')),
    ];

    const fetchSeedRows = async () => {
      const resp = await fetch(FUND_LIST_UPDATE_SEED_FILE, { cache: 'no-store' });
      if (!resp.ok) throw new Error(`Seed file not found (${resp.status})`);
      const payload = await resp.json();
      return cleanRows(payload?.rows || payload?.values || []);
    };

    const applyLoadedRows = (nextRows, nextSourceNote, nextSourceKey) => {
      rows = cleanRows(nextRows);
      listFrom = rows.find(row => row.list_from)?.list_from || '2025-Q3';
      listTo = rows.find(row => row.list_to)?.list_to || State.currentQuarter || '2026-Q1';
      const fromInput = $('#flu-list-from', area);
      const toInput = $('#flu-list-to', area);
      if (fromInput) fromInput.value = listFrom;
      if (toInput) toInput.value = listTo;
      sourceNote = nextSourceNote;
      loadError = '';
      State._pageDataSource[pageKey] = nextSourceKey;
      render();
    };

    let fundSuggestions = [];
    try {
      const selectRows = await fetchCached('select-fund');
      const seenCodes = new Set();
      fundSuggestions = buildSelectedFundsCatalog(selectRows)
        .filter(fund => {
          const key = normalizeFundKey(fund.code);
          if (!key || seenCodes.has(key)) return false;
          seenCodes.add(key);
          return true;
        })
        .sort((a, b) => String(a.code || '').localeCompare(String(b.code || ''), 'en'));
    } catch {
      fundSuggestions = [];
    }
    const fundMatches = (value = '') => {
      const q = String(value || '').trim().toLowerCase();
      if (!q) return fundSuggestions.slice(0, 10);
      const nq = normalizeFundKey(q);
      return fundSuggestions
        .filter(fund => (
          normalizeFundKey(fund.code).includes(nq)
          || String(fund.name || '').toLowerCase().includes(q)
          || String(fund.assetHouse || '').toLowerCase().includes(q)
        ))
        .slice(0, 10);
    };

    let rows = [];
    let sourceNote = 'Draft ใน Browser';
    let loadError = '';
    const draftRows = readDraftRows();
    let listFrom = '';
    let listTo = '';

    if (cfg.sheetId && window.SheetsAPI?.fetchSheetData) {
      try {
        rows = cleanRows(normalizeRows(await SheetsAPI.fetchSheetData(cfg.sheetId, cfg.tabName || 'fund_list_changes')));
        sourceNote = `${cfg.source || 'Google Sheet'} · ${cfg.tabName || 'fund_list_changes'}`;
        State._pageDataSource[pageKey] = 'Google Sheet';
      } catch (err) {
        loadError = err.message || String(err);
      }
    } else if (draftRows.length) {
      rows = draftRows;
      State._pageDataSource[pageKey] = 'Local draft';
    }

    if (!rows.length && draftRows.length) {
      rows = draftRows;
      sourceNote = 'Draft ใน Browser';
      State._pageDataSource[pageKey] = 'Local draft';
    }

    if (!rows.length) {
      try {
        rows = await fetchSeedRows();
        sourceNote = 'Seed จาก PDF: Fund List Q3-2025';
        State._pageDataSource[pageKey] = 'PDF seed';
      } catch { /* fallback to sample below */ }
    }

    if (!rows.length) {
      rows = cleanRows(FUND_LIST_UPDATE_SAMPLE_ROWS);
      sourceNote = 'ข้อมูลตัวอย่าง - กดบันทึก Draft เพื่อเริ่มใช้งาน';
      State._pageDataSource[pageKey] = cfg.sheetId ? 'Sample fallback' : 'Sample data';
    }
    listFrom = rows.find(row => row.list_from)?.list_from || '2025-Q3';
    listTo = rows.find(row => row.list_to)?.list_to || State.currentQuarter || '2026-Q1';

    const syncListMetaToRows = () => {
      listFrom = ($('#flu-list-from', area)?.value || listFrom || '').trim();
      listTo = ($('#flu-list-to', area)?.value || listTo || '').trim();
      rows = rows.map(row => ({
        ...row,
        list_from: listFrom,
        list_to: listTo,
      }));
    };

    const entryStatusOptions = () => [...new Set([
      ...FUND_LIST_UPDATE_STATUS_OPTIONS,
      ...rows.map(row => String(row.change_type || '').trim()).filter(Boolean),
    ])];
    const statusOptions = () => ['ทั้งหมด', ...entryStatusOptions()];
    const entryAssetOptions = () => [...new Set([
      ...FUND_LIST_UPDATE_ASSET_OPTIONS,
      ...rows.map(row => String(row.asset_class || '').trim()).filter(Boolean),
    ])].sort((a, b) => a.localeCompare(b, 'en'));

    const changeToneClass = (value = '') => {
      const text = String(value || '').trim();
      if (!text) return '';
      if (text.includes('เหมือนเดิม')) return 'flu-change-same';
      if (text.includes('เพิ่ม')) return 'flu-change-add';
      if (text.includes('สลับ')) return 'flu-change-swap';
      if (text.includes('นำ') && text.includes('ออก')) return 'flu-change-remove';
      if (text.includes('นํา') && text.includes('ออก')) return 'flu-change-remove';
      if (text.includes('ตัด') && text.includes('ออก')) return 'flu-change-remove';
      return '';
    };

    const statusCounts = () => rows.reduce((acc, row) => {
      const key = String(row.change_type || 'ไม่ระบุ').trim() || 'ไม่ระบุ';
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});

    const createBlankRow = () => {
      const now = new Date().toISOString();
      return {
        id: `FLU-${Date.now()}`,
        list_from: listFrom,
        list_to: listTo,
        type: '',
        asset_class: '',
        fund_list_old: '',
        fund_list_current: '',
        change_type: '',
        note: '',
        updated_at: now,
        updated_by: State.currentUser?.email || '',
      };
    };

    const appendBlankRow = () => {
      syncListMetaToRows();
      rows.push(createBlankRow());
      render();
    };

    const insertBlankRowAfter = (rowId) => {
      syncListMetaToRows();
      const index = rows.findIndex(row => row.id === rowId);
      if (index < 0) {
        rows.push(createBlankRow());
      } else {
        rows.splice(index + 1, 0, createBlankRow());
      }
      render();
    };

    const filteredRows = () => {
      const query = ($('#flu-search', area)?.value || '').trim().toLowerCase();
      const status = $('#flu-status', area)?.value || 'ทั้งหมด';
      return rows.filter(row => {
        const statusOk = status === 'ทั้งหมด' || String(row.change_type || '') === status;
        const queryOk = !query || FUND_LIST_UPDATE_COLUMNS.some(key => String(row[key] || '').toLowerCase().includes(query));
        return statusOk && queryOk;
      });
    };

    const renderA4Preview = (items) => {
      const pageSize = 42;
      const minLastPageRows = 6;
      const chunks = [];
      let currentType = null;
      let typeRows = [];
      const buildTypeChunks = (rowsForType, type) => {
        const typeChunks = [];
        for (let i = 0; i < rowsForType.length; i += pageSize) {
          typeChunks.push(rowsForType.slice(i, i + pageSize));
        }
        const last = typeChunks[typeChunks.length - 1];
        const prev = typeChunks[typeChunks.length - 2];
        if (typeChunks.length > 1 && last.length > 0 && last.length < minLastPageRows && prev.length > minLastPageRows) {
          const needed = Math.min(minLastPageRows - last.length, prev.length - minLastPageRows);
          last.unshift(...prev.splice(prev.length - needed, needed));
        }
        typeChunks.forEach(chunkRows => {
          chunks.push({
            type: type || '',
            rows: chunkRows,
          });
        });
      };
      const flushTypeRows = () => {
        if (!typeRows.length) return;
        buildTypeChunks(typeRows, currentType);
        typeRows = [];
      };
      items.forEach(row => {
        const rowType = String(row.type || '').trim();
        if (currentType !== null && rowType !== currentType) {
          flushTypeRows();
        }
        currentType = rowType;
        typeRows.push(row);
      });
      flushTypeRows();
      if (!chunks.length) {
        chunks.push({ type: '', rows: [] });
      }
      return chunks.map((pageChunk, pageIndex) => `
        <section class="flu-a4-page">
          <div class="flu-a4-head">
            <div>
              <h3>Fund List Update</h3>
              <p>${esc(listFrom || 'เดิม')} → ${esc(listTo || 'ปัจจุบัน')}${pageChunk.type ? ` · ${esc(pageChunk.type)}` : ''}</p>
            </div>
            <span>หน้า ${pageIndex + 1} / ${chunks.length}</span>
          </div>
          <table class="flu-a4-table">
            <thead>
              <tr>
                <th>Type</th>
                <th>สินทรัพย์</th>
                <th>Fund List ${esc(listFrom || 'เดิม')}</th>
                <th>Fund List ${esc(listTo || 'ปัจจุบัน')}</th>
                <th>สถานะ</th>
                <th>หมายเหตุ</th>
              </tr>
            </thead>
            <tbody>
              ${pageChunk.rows.map(row => `
                <tr>
                  <td>${esc(row.type || '')}</td>
                  <td>${esc(row.asset_class || '')}</td>
                  <td>${esc(row.fund_list_old || '')}</td>
                  <td>${esc(row.fund_list_current || '')}</td>
                  <td class="${changeToneClass(row.change_type)}">${esc(row.change_type || '')}</td>
                  <td class="${changeToneClass(row.note)}">${esc(row.note || '')}</td>
                </tr>
              `).join('') || '<tr><td colspan="6">ไม่มีข้อมูล</td></tr>'}
            </tbody>
          </table>
        </section>
      `).join('');
    };

    const render = () => {
      const filtered = filteredRows();
      const counts = statusCounts();
      const groupToneMap = new Map();
      let groupToneSeq = 0;
      const groupMeta = (row, prevRow) => {
        const key = `${row.type || ''}::${row.asset_class || ''}`;
        const prevKey = prevRow ? `${prevRow.type || ''}::${prevRow.asset_class || ''}` : '';
        if (!groupToneMap.has(key)) {
          groupToneSeq += 1;
          groupToneMap.set(key, ((groupToneSeq - 1) % 6) + 1);
        }
        return {
          key,
          tone: groupToneMap.get(key),
          isStart: key !== prevKey,
        };
      };

      const tbody = $('#flu-body', area);
      const count = $('#flu-count', area);
      const summary = $('#flu-summary', area);
      const sourceBadge = $('#flu-source-badge', area);
      const loadErrorBadge = $('#flu-load-error', area);
      const a4Preview = $('#flu-a4-preview', area);
      const oldHeader = $('#flu-old-list-header', area);
      const currentHeader = $('#flu-current-list-header', area);
      if (count) count.textContent = `${filtered.length.toLocaleString()} / ${rows.length.toLocaleString()} รายการ`;
      if (sourceBadge) sourceBadge.innerHTML = sourceBadgeHtml(pageKey, sourceNote);
      if (loadErrorBadge) loadErrorBadge.innerHTML = loadError ? `<span class="badge badge-warning">โหลด Sheet ไม่สำเร็จ: ${esc(loadError)}</span>` : '';
      if (oldHeader) oldHeader.textContent = `Fund List ${listFrom || 'เดิม'}`;
      if (currentHeader) currentHeader.textContent = `Fund List ${listTo || 'ปัจจุบัน'}`;
      if (a4Preview) a4Preview.innerHTML = renderA4Preview(filtered);
      if (summary) {
        summary.innerHTML = `
          <div class="flu-metric">
            <span>รายการทั้งหมด</span>
            <strong>${rows.length.toLocaleString()}</strong>
          </div>
          ${Object.entries(counts).map(([label, value]) => `
            <div class="flu-metric">
              <span>${esc(label)}</span>
              <strong>${Number(value).toLocaleString()}</strong>
            </div>
          `).join('')}`;
      }
      if (!tbody) return;
      const editableRowsHtml = filtered.length ? filtered.map((row, rowIndex) => {
        const meta = groupMeta(row, filtered[rowIndex - 1]);
        return `
          <tr data-id="${esc(row.id)}" class="flu-group-row flu-group-tone-${meta.tone}${meta.isStart ? ' is-group-start' : ''}">
            <td class="flu-drag-cell"><button class="flu-drag-handle" type="button" title="ลากเพื่อย้ายแถว" aria-label="ลากเพื่อย้ายแถว">↕</button></td>
            <td><input class="flu-cell-input" data-field="type" value="${esc(row.type || '')}" placeholder="Core - General"></td>
            <td>
              <select class="flu-cell-input" data-field="asset_class">
                <option value=""></option>
                ${entryAssetOptions().map(option => `<option value="${esc(option)}" ${row.asset_class === option ? 'selected' : ''}>${esc(option)}</option>`).join('')}
              </select>
            </td>
            <td><input class="flu-cell-input flu-code-input flu-fund-autocomplete" data-field="fund_list_old" value="${esc(row.fund_list_old || '')}" placeholder="พิมพ์รหัส เช่น KKP" autocomplete="off"></td>
            <td><input class="flu-cell-input flu-code-input flu-fund-autocomplete" data-field="fund_list_current" value="${esc(row.fund_list_current || '')}" placeholder="พิมพ์รหัส เช่น KKP" autocomplete="off"></td>
            <td>
              <select class="flu-cell-input flu-change-select ${changeToneClass(row.change_type)}" data-field="change_type">
                <option value=""></option>
                ${entryStatusOptions().map(option => `<option value="${esc(option)}" ${row.change_type === option ? 'selected' : ''}>${esc(option)}</option>`).join('')}
              </select>
            </td>
            <td><input class="flu-cell-input flu-note-input ${changeToneClass(row.note)}" data-field="note" value="${esc(row.note || '')}" placeholder="รายละเอียด"></td>
            <td><button class="btn btn-xs btn-danger flu-delete-row" type="button" data-id="${esc(row.id)}">ลบ</button></td>
          </tr>
          <tr class="flu-inline-add-row flu-group-tone-${meta.tone}" data-after-id="${esc(row.id)}">
            <td colspan="8"><button class="flu-inline-add" type="button" data-after-id="${esc(row.id)}">+ เพิ่มแถว</button></td>
          </tr>`;
      }).join('') : `
        <tr>
          <td colspan="8" class="flu-empty">ไม่พบรายการตามเงื่อนไขที่เลือก</td>
        </tr>`;
      tbody.innerHTML = `${editableRowsHtml}
        <tr class="flu-add-row-tr">
          <td colspan="8">
            <button class="flu-add-bottom" id="flu-add-row-bottom" type="button">+ เพิ่มแถวด้านล่าง</button>
          </td>
        </tr>`;
    };

    area.innerHTML = `
      <div class="flu-layout">
        <div class="page-tools flu-entry-tools">
          <div class="page-tools-meta">
            <span id="flu-source-badge">${sourceBadgeHtml(pageKey, sourceNote)}</span>
            <span class="badge badge-data-origin" id="flu-count">${rows.length.toLocaleString()} รายการ</span>
            <span id="flu-load-error">${loadError ? `<span class="badge badge-warning">โหลด Sheet ไม่สำเร็จ: ${esc(loadError)}</span>` : ''}</span>
          </div>
          <div class="page-tools-actions">
            ${sheetUrl ? `<a class="btn btn-ghost" id="flu-open-sheet" href="${sheetUrl}" target="_blank" rel="noopener noreferrer">เปิด Google Sheet</a>` : ''}
            <button class="btn btn-ghost" id="flu-load-sheet" type="button" ${cfg.sheetId ? '' : 'disabled'}>โหลดจาก Google Sheet</button>
            <button class="btn btn-ghost" id="flu-load-seed" type="button">โหลด Q3-2025 จาก PDF</button>
            <button class="btn btn-primary" id="flu-save-draft" type="button">บันทึก Draft</button>
            <button class="btn btn-success" id="flu-save-sheet" type="button" ${cfg.sheetId ? '' : 'disabled'}>บันทึกลง Google Sheet</button>
            <button class="btn btn-ghost" id="flu-export-csv" type="button">Export CSV</button>
          </div>
        </div>

        <div class="flu-summary" id="flu-summary"></div>

        <div class="card flu-card">
          <div class="card-header">
            <div>
              <div class="card-title">รอบข้อมูล Fund List</div>
              <div class="flu-subtitle">กำหนดครั้งเดียว แล้วระบบจะบันทึกไปกับทุกแถว</div>
            </div>
          </div>
          <div class="card-body">
            <div class="flu-period-grid">
              <label class="flu-period-field">
                <span>Fund List เดิม</span>
                <input class="flu-cell-input" id="flu-list-from" value="${esc(listFrom)}" placeholder="2025-Q3">
              </label>
              <label class="flu-period-field">
                <span>Fund List ปัจจุบัน</span>
                <input class="flu-cell-input" id="flu-list-to" value="${esc(listTo)}" placeholder="2026-Q1">
              </label>
            </div>
          </div>
        </div>

        <div class="card flu-card" id="report-card">
          <div class="card-header">
            <div>
              <div class="card-title">Fund List Update Entry</div>
              <div class="flu-subtitle">กรอกเป็น row-based data เพื่อบันทึกลง tab <code>fund_list_changes</code></div>
            </div>
          </div>
          <div class="card-body">
            <div class="filter-bar flu-filter-bar">
              <input class="search-input" id="flu-search" type="text" placeholder="ค้นหา fund code, asset class, note..." autocomplete="off">
              <select class="filter-select" id="flu-status">
                ${statusOptions().map(option => `<option value="${esc(option)}">${esc(option)}</option>`).join('')}
              </select>
            </div>
            <div class="table-wrapper flu-table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th class="flu-drag-head"></th>
                    <th>Type</th>
                    <th>สินทรัพย์</th>
                    <th id="flu-old-list-header">Fund List ${esc(listFrom || 'เดิม')}</th>
                    <th id="flu-current-list-header">Fund List ${esc(listTo || 'ปัจจุบัน')}</th>
                    <th>สถานะ</th>
                    <th>หมายเหตุ</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody id="flu-body"></tbody>
              </table>
            </div>
            <div class="flu-suggest-menu" id="flu-suggest-menu" hidden></div>
          </div>
        </div>

        <div class="card flu-card">
          <div class="card-header">
            <div class="card-title">Google Sheet Schema</div>
          </div>
          <div class="card-body">
            <div class="flu-schema">
              ${FUND_LIST_UPDATE_COLUMNS.map(column => `<code>${esc(column)}</code>`).join('')}
            </div>
          </div>
        </div>

        <div class="card flu-card">
          <div class="card-header">
            <div>
              <div class="card-title">รายชื่อกองที่เตรียมไว้</div>
              <div class="flu-subtitle">Preview ขนาดประมาณ A4 แบ่งหน้าตามจำนวนรายการที่แสดงอยู่</div>
            </div>
          </div>
          <div class="card-body flu-a4-stage" id="flu-a4-preview"></div>
        </div>
      </div>`;

    $('#flu-search', area)?.addEventListener('input', render);
    $('#flu-status', area)?.addEventListener('change', render);
    $('#flu-list-from', area)?.addEventListener('input', () => {
      syncListMetaToRows();
      render();
    });
    $('#flu-list-to', area)?.addEventListener('input', () => {
      syncListMetaToRows();
      render();
    });
    $('#flu-load-seed', area)?.addEventListener('click', async () => {
      try {
        applyLoadedRows(await fetchSeedRows(), 'Seed จาก PDF: Fund List Q3-2025', 'PDF seed');
        toast(`โหลด Q3-2025 จาก PDF แล้ว ${rows.length.toLocaleString()} รายการ`, 'success');
      } catch (err) {
        toast(`โหลด seed ไม่สำเร็จ: ${err.message || err}`, 'error');
      }
    });
    $('#flu-load-sheet', area)?.addEventListener('click', async () => {
      if (!cfg.sheetId) {
        toast('ยังไม่ได้ตั้งค่า Google Sheet ID สำหรับ Fund List Update', 'warning');
        return;
      }
      try {
        if (!SheetsAPI.accessToken) await SheetsAPI.requestToken(false);
        const sheetRows = normalizeRows(await SheetsAPI.fetchSheetData(cfg.sheetId, cfg.tabName || 'fund_list_changes'));
        applyLoadedRows(sheetRows, `${cfg.source || 'Google Sheet'} · ${cfg.tabName || 'fund_list_changes'}`, 'Google Sheet');
        toast(`โหลดจาก Google Sheet แล้ว ${rows.length.toLocaleString()} รายการ`, 'success');
      } catch (err) {
        loadError = err.message || String(err);
        render();
        toast(`โหลดจาก Google Sheet ไม่สำเร็จ: ${loadError}`, 'error', 6000);
      }
    });
    $('#flu-save-draft', area)?.addEventListener('click', () => {
      syncListMetaToRows();
      const savedAt = writeDraftRows();
      render();
      toast(`บันทึก Draft แล้ว (${savedAt.slice(0, 19).replace('T', ' ')})`, 'success');
    });
    $('#flu-save-sheet', area)?.addEventListener('click', async () => {
      if (!cfg.sheetId) {
        toast('ยังไม่ได้ตั้งค่า Google Sheet ID สำหรับ Fund List Update', 'warning');
        return;
      }
      try {
        syncListMetaToRows();
        writeDraftRows();
        await SheetsAPI.updateSheetValues(cfg.sheetId, cfg.tabName || 'fund_list_changes', rowsToSheetValues());
        State._pageDataSource[pageKey] = 'Google Sheet';
        render();
        toast('บันทึกลง Google Sheet แล้ว', 'success');
      } catch (err) {
        toast(`บันทึกลง Google Sheet ไม่สำเร็จ: ${err.message || err}`, 'error', 6000);
      }
    });
    $('#flu-export-csv', area)?.addEventListener('click', () => {
      syncListMetaToRows();
      const csv = rowsToSheetValues()
        .map(row => row.map(cell => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(','))
        .join('\n');
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `fund-list-update-${new Date().toISOString().slice(0, 10)}.csv`;
      a.click();
      URL.revokeObjectURL(url);
    });
    const autoSetChangeType = (row) => {
      const oldCode = normalizeFundKey(row.fund_list_old);
      const currentCode = normalizeFundKey(row.fund_list_current);
      if (!oldCode || !currentCode) return false;
      const nextType = oldCode === currentCode ? 'เหมือนเดิม' : 'เปลี่ยนแปลง';
      if (row.change_type === nextType) return false;
      row.change_type = nextType;
      return true;
    };
    $('#flu-body', area)?.addEventListener('input', (event) => {
      const input = event.target.closest('[data-field]');
      if (!input) return;
      const rowEl = input.closest('tr[data-id]');
      const row = rows.find(item => item.id === rowEl?.dataset.id);
      if (!row) return;
      row[input.dataset.field] = input.value;
      if (input.dataset.field === 'fund_list_old' || input.dataset.field === 'fund_list_current') {
        autoSetChangeType(row);
      }
      row.updated_at = new Date().toISOString();
      row.updated_by = State.currentUser?.email || row.updated_by || '';
      if (input.classList.contains('flu-fund-autocomplete')) {
        renderFundSuggestMenu(input);
      }
    });
    $('#flu-body', area)?.addEventListener('change', (event) => {
      const input = event.target.closest('[data-field]');
      if (!input) return;
      const rowEl = input.closest('tr[data-id]');
      const row = rows.find(item => item.id === rowEl?.dataset.id);
      if (!row) return;
      row[input.dataset.field] = input.value;
      if (input.dataset.field === 'fund_list_old' || input.dataset.field === 'fund_list_current') {
        autoSetChangeType(row);
      }
      row.updated_at = new Date().toISOString();
      row.updated_by = State.currentUser?.email || row.updated_by || '';
      render();
    });
    $('#flu-body', area)?.addEventListener('click', (event) => {
      const inlineAddBtn = event.target.closest('.flu-inline-add');
      if (inlineAddBtn) {
        insertBlankRowAfter(inlineAddBtn.dataset.afterId || '');
        return;
      }
      const addBtn = event.target.closest('.flu-add-bottom');
      if (addBtn) {
        appendBlankRow();
        return;
      }
      const btn = event.target.closest('.flu-delete-row');
      if (!btn) return;
      rows = rows.filter(row => row.id !== btn.dataset.id);
      render();
    });

    let dragRowId = '';
    let pointerDragId = '';
    let pointerOverId = '';
    const reorderRows = (fromId, toId) => {
      if (!fromId || !toId || fromId === toId) return false;
      const visibleIds = filteredRows().map(row => row.id);
      const fromVisibleIndex = visibleIds.indexOf(fromId);
      const toVisibleIndex = visibleIds.indexOf(toId);
      if (fromVisibleIndex < 0 || toVisibleIndex < 0) return false;
      const reorderedVisible = [...visibleIds];
      reorderedVisible.splice(fromVisibleIndex, 1);
      reorderedVisible.splice(toVisibleIndex, 0, fromId);
      const visibleSet = new Set(visibleIds);
      let cursor = 0;
      rows = rows.map(row => {
        if (!visibleSet.has(row.id)) return row;
        const nextId = reorderedVisible[cursor++];
        return rows.find(item => item.id === nextId) || row;
      });
      return true;
    };
    const clearPointerDragState = () => {
      pointerDragId = '';
      pointerOverId = '';
      document.body.classList.remove('flu-row-dragging');
      $$('.is-dragging, .is-drag-over', area).forEach(el => el.classList.remove('is-dragging', 'is-drag-over'));
    };
    const pointerMoveRow = (clientX, clientY) => {
      if (!pointerDragId) return;
      const el = document.elementFromPoint(clientX, clientY);
      const rowEl = el?.closest?.('#flu-body tr[data-id]');
      $$('.is-drag-over', area).forEach(row => {
        if (row !== rowEl) row.classList.remove('is-drag-over');
      });
      if (!rowEl || rowEl.dataset.id === pointerDragId) {
        pointerOverId = '';
        return;
      }
      pointerOverId = rowEl.dataset.id || '';
      rowEl.classList.add('is-drag-over');
    };
    $('#flu-body', area)?.addEventListener('pointerdown', (event) => {
      const handle = event.target.closest('.flu-drag-handle');
      if (!handle) return;
      const rowEl = handle.closest('tr[data-id]');
      if (!rowEl) return;
      event.preventDefault();
      hideFundSuggestMenu();
      pointerDragId = rowEl.dataset.id || '';
      pointerOverId = '';
      rowEl.classList.add('is-dragging');
      document.body.classList.add('flu-row-dragging');
      try { handle.setPointerCapture?.(event.pointerId); } catch { /* noop */ }
    });
    $('#flu-body', area)?.addEventListener('pointermove', (event) => {
      if (!pointerDragId) return;
      event.preventDefault();
      pointerMoveRow(event.clientX, event.clientY);
    });
    $('#flu-body', area)?.addEventListener('pointerup', (event) => {
      if (!pointerDragId) return;
      event.preventDefault();
      pointerMoveRow(event.clientX, event.clientY);
      const fromId = pointerDragId;
      const toId = pointerOverId;
      const moved = reorderRows(fromId, toId);
      clearPointerDragState();
      if (moved) render();
    });
    $('#flu-body', area)?.addEventListener('pointercancel', clearPointerDragState);
    $('#flu-body', area)?.addEventListener('dragstart', (event) => {
      const rowEl = event.target.closest('tr[data-id]');
      if (!rowEl || !event.target.closest('.flu-drag-handle')) {
        event.preventDefault();
        return;
      }
      dragRowId = rowEl.dataset.id || '';
      rowEl.classList.add('is-dragging');
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', dragRowId);
      hideFundSuggestMenu();
    });
    $('#flu-body', area)?.addEventListener('dragover', (event) => {
      const rowEl = event.target.closest('tr[data-id]');
      if (!dragRowId || !rowEl || rowEl.dataset.id === dragRowId) return;
      event.preventDefault();
      rowEl.classList.add('is-drag-over');
      event.dataTransfer.dropEffect = 'move';
    });
    $('#flu-body', area)?.addEventListener('dragleave', (event) => {
      const rowEl = event.target.closest('tr[data-id]');
      rowEl?.classList.remove('is-drag-over');
    });
    $('#flu-body', area)?.addEventListener('drop', (event) => {
      const rowEl = event.target.closest('tr[data-id]');
      if (!dragRowId || !rowEl) return;
      event.preventDefault();
      const moved = reorderRows(dragRowId, rowEl.dataset.id || '');
      dragRowId = '';
      $$('.is-dragging, .is-drag-over', area).forEach(el => el.classList.remove('is-dragging', 'is-drag-over'));
      if (moved) render();
    });
    $('#flu-body', area)?.addEventListener('dragend', () => {
      dragRowId = '';
      $$('.is-dragging, .is-drag-over', area).forEach(el => el.classList.remove('is-dragging', 'is-drag-over'));
    });

    const suggestMenu = $('#flu-suggest-menu', area);
    const hideFundSuggestMenu = () => {
      if (!suggestMenu) return;
      suggestMenu.hidden = true;
      suggestMenu.innerHTML = '';
    };
    const applyFundSuggestion = (input, fund) => {
      const rowEl = input?.closest('tr[data-id]');
      const row = rows.find(item => item.id === rowEl?.dataset.id);
      if (!row) return;
      input.value = fund.code || '';
      row[input.dataset.field] = fund.code || '';
      if (input.dataset.field === 'fund_list_old' || input.dataset.field === 'fund_list_current') {
        autoSetChangeType(row);
      }
      row.updated_at = new Date().toISOString();
      row.updated_by = State.currentUser?.email || row.updated_by || '';
      hideFundSuggestMenu();
      render();
    };
    const renderFundSuggestMenu = (input) => {
      if (!suggestMenu || !input) return;
      const matches = fundMatches(input.value);
      if (!matches.length) {
        hideFundSuggestMenu();
        return;
      }
      const rect = input.getBoundingClientRect();
      const layoutEl = $('.flu-layout', area) || area;
      const layoutRect = layoutEl.getBoundingClientRect();
      suggestMenu.style.left = `${Math.max(8, rect.left - layoutRect.left + layoutEl.scrollLeft)}px`;
      suggestMenu.style.top = `${rect.bottom - layoutRect.top + layoutEl.scrollTop + 4}px`;
      suggestMenu.style.width = `${Math.max(260, rect.width)}px`;
      suggestMenu.innerHTML = matches.map(fund => `
        <button class="flu-suggest-option" type="button" data-code="${esc(fund.code)}">
          <strong>${esc(fund.code || '-')}</strong>
        </button>
      `).join('');
      suggestMenu.hidden = false;
      suggestMenu.dataset.targetRowId = input.closest('tr[data-id]')?.dataset.id || '';
      suggestMenu.dataset.targetField = input.dataset.field || '';
    };
    $('#flu-body', area)?.addEventListener('focusin', (event) => {
      const input = event.target.closest('.flu-fund-autocomplete');
      if (!input) return;
      renderFundSuggestMenu(input);
    });
    $('#flu-body', area)?.addEventListener('keydown', (event) => {
      const input = event.target.closest('.flu-fund-autocomplete');
      if (!input) return;
      if (event.key === 'Escape') hideFundSuggestMenu();
      if (event.key !== 'Enter' || suggestMenu?.hidden) return;
      const firstOption = suggestMenu.querySelector('.flu-suggest-option');
      if (!firstOption) return;
      const fund = fundSuggestions.find(item => item.code === firstOption.dataset.code);
      if (fund) {
        event.preventDefault();
        applyFundSuggestion(input, fund);
      }
    });
    suggestMenu?.addEventListener('mousedown', (event) => {
      event.preventDefault();
      const option = event.target.closest('.flu-suggest-option');
      if (!option) return;
      const input = $(`#flu-body tr[data-id="${CSS.escape(suggestMenu.dataset.targetRowId || '')}"] [data-field="${CSS.escape(suggestMenu.dataset.targetField || '')}"]`, area);
      const fund = fundSuggestions.find(item => item.code === option.dataset.code);
      if (input && fund) applyFundSuggestion(input, fund);
    });
    area.addEventListener('click', (event) => {
      if (event.target.closest('.flu-fund-autocomplete') || event.target.closest('#flu-suggest-menu')) return;
      hideFundSuggestMenu();
    });
    render();
    App._currentExport = null;
    App._currentTableExport = () => buildSimpleTablePayload(
      'Fund List Update',
      sourceNote,
      ['Fund List เดิม', 'Fund List ปัจจุบัน', 'Type', 'สินทรัพย์', `Fund List ${listFrom || 'เดิม'}`, `Fund List ${listTo || 'ปัจจุบัน'}`, 'สถานะ', 'หมายเหตุ'],
      filteredRows().map(row => [
        row.list_from || '',
        row.list_to || '',
        row.type || '',
        row.asset_class || '',
        row.fund_list_old || '',
        row.fund_list_current || '',
        row.change_type || '',
        row.note || '',
      ]),
    );
    App._currentClipboardExport = null;
    App._currentImageExport = null;
    bindPageImageActions(area, 'report-card', 'fund-list-update');
  },

  async fundSelectionLogs(area) {
    const pageKey = 'fund-selection-logs';
    const quarter = State.currentQuarter || CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1';
    const sheetUrl = googleSheetUrl(FUND_SELECTION_LOGS_SHEET_ID);
    setLoading(area, 'กำลังโหลด Fund Selection Logs...');

    const roleLabel = (value) => FUND_SELECTION_LOG_ROLES.find(item => item.value === value)?.label || value || 'ไม่ระบุ';
    const nowIso = () => new Date().toISOString();
    const statusLabel = (value) => ({
      existing: 'กองทุนเดิม',
      switched: 'กองทุนเดิม (สลับตำแหน่ง)',
      passive: 'กองทุนเดิม (Passive)',
      active: 'กองทุนเดิม (Active)',
      new: 'กองทุนใหม่',
    }[String(value || '').trim()] || String(value || '').trim());
    const statusOptionsFor = (value) => {
      const current = statusLabel(value);
      return [...new Set(['', ...FUND_SELECTION_LOG_STATUS_OPTIONS, current].filter(option => option !== undefined))];
    };
    const itemTitleValue = (item = {}) => [item.assetClass, item.fundType]
      .map(value => String(value || '').trim())
      .filter(Boolean)
      .join(' / ');
    const applyItemTitle = (item, value) => {
      const parts = String(value || '').split('/').map(part => part.trim());
      item.assetClass = parts.shift() || '';
      item.fundType = parts.join(' / ');
    };
    const suggestedTagsFor = (mention = {}) => {
      const haystack = [
        mention.reason,
        mention.fundCode,
        mention.status,
      ].join(' ').toLowerCase();
      const existing = new Set(mention.tags || []);
      return FUND_SELECTION_TAG_RULES
        .filter(rule => !existing.has(rule.tag))
        .filter(rule => rule.keywords.some(keyword => haystack.includes(String(keyword || '').toLowerCase())))
        .map(rule => rule.tag);
    };
    const addMentionTag = (mention, tag) => {
      const cleanTag = String(tag || '').trim().replace(/,$/, '');
      if (!mention || !cleanTag) return false;
      const tags = new Set(mention.tags || []);
      const before = tags.size;
      tags.add(cleanTag);
      mention.tags = [...tags];
      mention.updatedAt = nowIso();
      mention.updatedBy = State.currentUser?.email || mention.updatedBy || '';
      return tags.size !== before;
    };
    const slug = (value, fallback = 'item') => String(value || fallback)
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '-')
      .replace(/[^a-z0-9ก-๙._-]+/g, '')
      .replace(/^-+|-+$/g, '')
      .slice(0, 80) || fallback;
    const newId = (prefix) => `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    const shouldUseLocalLogsApi = () => ['localhost', '127.0.0.1', '::1'].includes(window.location.hostname);
    const hasLogsApi = () => Boolean(String(FUND_SELECTION_LOGS_API_WEB_APP_URL || '').trim());

    const logsApiJsonp = (params) => new Promise((resolve, reject) => {
      const callbackName = `__fslApiCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
      const script = document.createElement('script');
      const cleanup = () => {
        delete window[callbackName];
        script.remove();
      };
      const timer = setTimeout(() => {
        cleanup();
        reject(new Error('Fund Selection Logs API request timeout'));
      }, 30000);
      window[callbackName] = (data) => {
        clearTimeout(timer);
        cleanup();
        resolve(data || {});
      };
      const url = new URL(FUND_SELECTION_LOGS_API_WEB_APP_URL);
      Object.entries({
        key: FUND_SELECTION_LOGS_API_SECRET_KEY,
        callback: callbackName,
        ...params,
      }).forEach(([key, value]) => {
        if (value !== undefined && value !== null && value !== '') {
          url.searchParams.set(key, value);
        }
      });
      script.onerror = () => {
        clearTimeout(timer);
        cleanup();
        reject(new Error('Fund Selection Logs API script failed to load'));
      };
      script.src = url.toString();
      document.head.appendChild(script);
    });

    const logsApiFetch = async (params) => {
      const url = new URL(FUND_SELECTION_LOGS_API_WEB_APP_URL);
      Object.entries({
        key: FUND_SELECTION_LOGS_API_SECRET_KEY,
        ...params,
      }).forEach(([key, value]) => {
        if (value !== undefined && value !== null && value !== '') {
          url.searchParams.set(key, value);
        }
      });
      const res = await fetch(url.toString(), { cache: 'no-store', redirect: 'follow' });
      const text = await res.text();
      if (!res.ok) throw new Error(`Fund Selection Logs API HTTP ${res.status}`);
      return JSON.parse(text);
    };

    const logsApiCompressedPayload = async (payload) => {
      const text = JSON.stringify(payload || {});
      if (!('CompressionStream' in window)) return { payload: text };
      const stream = new Blob([text], { type: 'application/json' }).stream().pipeThrough(new CompressionStream('gzip'));
      const buffer = await new Response(stream).arrayBuffer();
      let binary = '';
      const bytes = new Uint8Array(buffer);
      for (let i = 0; i < bytes.length; i += 0x8000) {
        binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      }
      return {
        payloadGzipB64: btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, ''),
      };
    };

    const logsApiRequest = async (action, payload = {}) => {
      const params = { action };
      if (action === 'get' || action === 'delete') {
        params.quarter = payload.quarter || quarter;
      } else {
        Object.assign(params, await logsApiCompressedPayload({
          action,
          key: FUND_SELECTION_LOGS_API_SECRET_KEY,
          ...payload,
        }));
      }
      let data;
      try {
        data = await logsApiFetch(params);
      } catch {
        data = await logsApiJsonp(params);
      }
      if (data.ok === false && !data.conflict) {
        throw new Error(data.error || `Fund Selection Logs API ${action} failed`);
      }
      return data;
    };

    const defaultLog = () => ({
      schemaVersion: 1,
      quarter,
      title: `Fund Selection Logs ${quarter}`,
      revision: 0,
      dataAsOf: '',
      createdAt: nowIso(),
      updatedAt: nowIso(),
      updatedBy: State.currentUser?.email || '',
      driveFolderId: FUND_SELECTION_LOGS_DRIVE_FOLDER_ID,
      items: [],
    });

    const normalizeMention = (mention = {}, idx = 0) => ({
      id: String(mention.id || newId(`mention-${idx + 1}`)),
      fundCode: String(mention.fundCode || mention.code || '').trim().toUpperCase(),
      role: String(mention.role || 'mainChoice').trim(),
      status: statusLabel(mention.status),
      sentiment: String(mention.sentiment || 'neutral').trim(),
      reason: String(mention.reason || mention.note || '').trim(),
      tags: Array.isArray(mention.tags)
        ? mention.tags.map(tag => String(tag || '').trim()).filter(Boolean)
        : String(mention.tags || '').split(',').map(tag => tag.trim()).filter(Boolean),
      updatedAt: String(mention.updatedAt || '').trim(),
      updatedBy: String(mention.updatedBy || '').trim(),
      rowRevision: Number(mention.rowRevision || mention.row_revision || 1),
    });

    const normalizeItem = (item = {}, idx = 0) => {
      const assetClass = String(item.assetClass || item.asset_class || '').trim();
      const fundType = String(item.fundType || item.fund_type || '').trim();
      const id = String(item.id || `${quarter}-${slug(assetClass || `section-${idx + 1}`)}-${slug(fundType || 'general')}`);
      return {
        id,
        assetClass,
        fundType,
        category: String(item.category || '').trim(),
        itemRevision: Number(item.itemRevision || 1),
        updatedAt: String(item.updatedAt || '').trim(),
        updatedBy: String(item.updatedBy || '').trim(),
        mentions: (Array.isArray(item.mentions) ? item.mentions : [])
          .map(normalizeMention),
      };
    };

    const normalizeLog = (payload) => {
      const base = { ...defaultLog(), ...(payload && typeof payload === 'object' ? payload : {}) };
      base.quarter = String(base.quarter || quarter).trim().toUpperCase();
      base.revision = Number(base.revision || 0);
      base.items = (Array.isArray(base.items) ? base.items : []).map(normalizeItem);
      base.driveFolderId = base.driveFolderId || FUND_SELECTION_LOGS_DRIVE_FOLDER_ID;
      return base;
    };

    const cleanForSave = () => ({
      ...log,
      quarter,
      title: ($('#fsl-title', area)?.value || log.title || `Fund Selection Logs ${quarter}`).trim(),
      dataAsOf: ($('#fsl-data-as-of', area)?.value || '').trim(),
      updatedBy: State.currentUser?.email || log.updatedBy || '',
      items: log.items.map(item => ({
        ...item,
        mentions: item.mentions
          .map(mention => ({
            ...mention,
            fundCode: String(mention.fundCode || '').trim().toUpperCase(),
            tags: Array.isArray(mention.tags) ? mention.tags.map(tag => String(tag || '').trim()).filter(Boolean) : [],
          }))
          .filter(mention => mention.fundCode || mention.reason || mention.tags.length),
      })),
    });
    const sheetColumnName = (index) => {
      let name = '';
      let value = Number(index || 0);
      while (value > 0) {
        const rem = (value - 1) % 26;
        name = String.fromCharCode(65 + rem) + name;
        value = Math.floor((value - 1) / 26);
      }
      return name || 'A';
    };
    const fslHeaderIndex = (headers = FUND_SELECTION_LOGS_SHEET_HEADERS) => {
      const normalizeHeader = (value) => String(value || '').trim().toLowerCase().replace(/[^a-z0-9ก-๙]+/g, '');
      const aliases = {
        quarter: 'quarter',
        itemorder: 'item_order',
        order: 'item_order',
        itemid: 'item_id',
        sectionid: 'item_id',
        assetclass: 'asset_class',
        asset: 'asset_class',
        fundtype: 'fund_type',
        type: 'fund_type',
        category: 'category',
        role: 'role',
        fundcode: 'fund_code',
        code: 'fund_code',
        status: 'status',
        reason: 'reason',
        note: 'reason',
        notes: 'reason',
        tags: 'tags',
        tag: 'tags',
        dataasof: 'data_as_of',
        asof: 'data_as_of',
        itemrevision: 'item_revision',
        updatedby: 'updated_by',
        updater: 'updated_by',
        updatedat: 'updated_at',
        mentionid: 'mention_id',
        rowrevision: 'row_revision',
        deleted: 'deleted',
      };
      const index = {};
      headers.forEach((header, headerIndex) => {
        const raw = String(header || '').trim().toLowerCase();
        const normalized = normalizeHeader(header);
        const canonical = aliases[normalized];
        if (raw) index[raw] = headerIndex;
        if (normalized) index[normalized] = headerIndex;
        if (canonical) index[canonical] = headerIndex;
      });
      return index;
    };
    const normalizeFslRole = (value) => {
      const key = String(value || '').trim().replace(/\s+/g, '').toLowerCase();
      return ({
        mainchoice: 'mainChoice',
        'ตัวเลือกหลัก': 'mainChoice',
        secondarychoice: 'secondaryChoice',
        'ตัวเลือกรอง': 'secondaryChoice',
        additionalnote: 'additionalNote',
        'ความเห็นเพิ่มเติม': 'additionalNote',
        notselected: 'notSelected',
        'ไม่ถูกคัดเลือก': 'notSelected',
      })[key] || String(value || 'mainChoice').trim();
    };
    const fslSplitTags = (value) => String(value || '')
      .split(/[,|]/)
      .map(tag => tag.trim())
      .filter(Boolean);
    const fslNumber = (value, fallback = 1) => {
      const number = Number(value);
      return Number.isFinite(number) && number > 0 ? number : fallback;
    };
    const fundSelectionLogFromSheetRows = (rows = []) => {
      const headerKeys = ['asset_class', 'fund_type', 'fund_code', 'reason', 'role', 'item_id', 'mention_id'];
      const headerCandidates = rows.slice(0, 20).map((row, rowIndex) => {
        const candidateIndex = fslHeaderIndex(row || []);
        const score = headerKeys.filter(key => candidateIndex[key] !== undefined).length;
        return { rowIndex, score };
      });
      const headerRowIndex = headerCandidates
        .filter(candidate => candidate.score >= 2)
        .sort((a, b) => b.score - a.score || a.rowIndex - b.rowIndex)[0]?.rowIndex || 0;
      const headers = (rows[headerRowIndex] || []).map(header => String(header || '').trim());
      const index = fslHeaderIndex(headers.length ? headers : FUND_SELECTION_LOGS_SHEET_HEADERS);
      const get = (row, key) => String(row[index[key]] ?? '').trim();
      const generatedAt = nowIso();
      const grouped = new Map();
      let dataAsOf = '';
      let updatedBy = '';
      let latestUpdatedAt = '';

      rows.slice(headerRowIndex + 1).forEach((row, rowIndex) => {
        if (!row.some(cell => String(cell ?? '').trim())) return;
        if (/^true$/i.test(get(row, 'deleted'))) return;
        const rowQuarter = (get(row, 'quarter') || quarter).toUpperCase();
        if (rowQuarter !== quarter) return;

        const assetClass = get(row, 'asset_class');
        const fundType = get(row, 'fund_type');
        const category = get(row, 'category');
        if (!assetClass && !fundType) return;

        dataAsOf = dataAsOf || get(row, 'data_as_of');
        updatedBy = get(row, 'updated_by') || updatedBy;
        latestUpdatedAt = get(row, 'updated_at') || latestUpdatedAt;

        const itemId = get(row, 'item_id') || `${quarter}-${slug(assetClass || `section-${rowIndex + 1}`, 'section')}-${slug(fundType || 'general', 'general')}`;
        const itemOrder = fslNumber(get(row, 'item_order'), rowIndex + 1);
        const groupKey = itemId || `${assetClass}|${fundType}|${category}`;
        if (!grouped.has(groupKey)) {
          grouped.set(groupKey, {
            id: itemId,
            assetClass,
            fundType,
            category,
            itemRevision: fslNumber(get(row, 'item_revision'), 1),
            updatedAt: get(row, 'updated_at') || generatedAt,
            updatedBy: get(row, 'updated_by'),
            itemOrder,
            mentions: [],
          });
        }

        const fundCode = get(row, 'fund_code').toUpperCase();
        const reason = get(row, 'reason');
        const tags = fslSplitTags(get(row, 'tags'));
        if (!fundCode && !reason && !tags.length) return;

        grouped.get(groupKey).mentions.push({
          id: get(row, 'mention_id') || newId(`mention-${rowIndex + 1}`),
          fundCode,
          role: normalizeFslRole(get(row, 'role')),
          status: statusLabel(get(row, 'status')),
          sentiment: 'neutral',
          reason,
          tags,
          updatedAt: get(row, 'updated_at') || generatedAt,
          updatedBy: get(row, 'updated_by'),
          rowRevision: fslNumber(get(row, 'row_revision'), 1),
        });
      });

      const items = [...grouped.values()]
        .sort((a, b) => (a.itemOrder - b.itemOrder) || String(a.assetClass).localeCompare(String(b.assetClass)))
        .map(item => {
          const { itemOrder, ...cleanItem } = item;
          return cleanItem;
        });

      return normalizeLog({
        schemaVersion: 1,
        quarter,
        title: `Fund Selection Logs ${quarter}`,
        revision: 0,
        dataAsOf,
        createdAt: generatedAt,
        updatedAt: latestUpdatedAt || generatedAt,
        updatedBy,
        driveFolderId: FUND_SELECTION_LOGS_DRIVE_FOLDER_ID,
        items,
      });
    };
    const loadLogFromSheet = async () => {
      if (!SheetsAPI.accessToken) await SheetsAPI.requestToken(false);
      const meta = await SheetsAPI.getSheetTabs(FUND_SELECTION_LOGS_SHEET_ID);
      if (!meta.tabs.includes(quarter)) {
        return {
          log: normalizeLog(null),
          sourceNote: 'New file',
          warning: `ยังไม่พบแท็บ ${quarter} ใน Google Sheet`,
        };
      }
      const rows = await SheetsAPI.fetchSheetData(FUND_SELECTION_LOGS_SHEET_ID, quarter);
      const logFromSheet = fundSelectionLogFromSheetRows(rows);
      return {
        log: logFromSheet,
        sourceNote: `Google Sheet: ${quarter}`,
        warning: logFromSheet.items.length ? '' : `แท็บ ${quarter} ยังไม่มีรายการบันทึก`,
      };
    };
    const loadLogFromJsonStore = async () => {
      if (!shouldUseLocalLogsApi() && hasLogsApi()) {
        try {
          const data = await logsApiRequest('get', { quarter });
          return {
            log: normalizeLog(data.log || null),
            sourceNote: data.source || (data.notFound ? 'New file' : `Google Drive: Fund Selection Logs - ${quarter}.json`),
            warning: data.notFound ? `ยังไม่พบไฟล์ของ ${quarter}` : '',
          };
        } catch (err) {
          try {
            const localFile = `Data/Fund Selection Logs - ${quarter}.json`;
            const resp = await fetch(localFile, { cache: 'no-store' });
            if (!resp.ok) throw new Error(`Local ${resp.status}`);
            const payload = await resp.json();
            return {
              log: normalizeLog(payload),
              sourceNote: `Local JSON fallback: ${localFile}`,
              warning: `Fund Selection Logs API ใช้งานไม่ได้: ${err.message || err}`,
            };
          } catch {
            return {
              log: normalizeLog(null),
              sourceNote: 'New file',
              warning: `Fund Selection Logs API ใช้งานไม่ได้: ${err.message || err}`,
            };
          }
        }
      }
      try {
        const resp = await fetch(`/api/fund-selection-logs?quarter=${encodeURIComponent(quarter)}`, { cache: 'no-store' });
        const data = await resp.json().catch(() => ({}));
        if (!resp.ok || data.ok === false) throw new Error(data.error || `API ${resp.status}`);
        return {
          log: normalizeLog(data.log || null),
          sourceNote: data.source || 'New file',
          warning: data.warning || (data.notFound ? `ยังไม่พบไฟล์ของ ${quarter}` : ''),
        };
      } catch (err) {
        try {
          const localFile = `Data/Fund Selection Logs - ${quarter}.json`;
          const resp = await fetch(localFile, { cache: 'no-store' });
          if (!resp.ok) throw new Error(`Local ${resp.status}`);
          const payload = await resp.json();
          return {
            log: normalizeLog(payload),
            sourceNote: `Local JSON fallback: ${localFile}`,
            warning: `API ใช้งานไม่ได้: ${err.message || err}`,
          };
        } catch {
          return {
            log: normalizeLog(null),
            sourceNote: 'New file',
            warning: `ยังไม่พบไฟล์ของ ${quarter}`,
          };
        }
      }
    };
    const fslSheetRecordFromRow = (headers, row = [], rowNumber = 1) => {
      const index = fslHeaderIndex(headers);
      const get = (key) => String(row[index[key]] ?? '').trim();
      return {
        rowNumber,
        mentionId: get('mention_id'),
        itemId: get('item_id'),
        quarter: get('quarter'),
        rowRevision: Number(get('row_revision') || 1),
        deleted: /^true$/i.test(get('deleted')),
        row,
      };
    };
    const alignFslRowToHeaders = (canonicalRow, headers) => {
      const canonicalIndex = fslHeaderIndex(FUND_SELECTION_LOGS_SHEET_HEADERS);
      return headers.map(header => canonicalRow[canonicalIndex[String(header || '').trim().toLowerCase()]] ?? '');
    };
    const fslMentionRecordsForSave = (sheetLog) => {
      const dataAsOf = String(sheetLog.dataAsOf || '').trim();
      const fallbackUpdatedBy = State.currentUser?.email || sheetLog.updatedBy || '';
      const fallbackUpdatedAt = nowIso();
      const records = [];
      (sheetLog.items || []).forEach((item, itemIndex) => {
        const mentions = Array.isArray(item.mentions) ? item.mentions : [];
        mentions.forEach(mention => {
          if (!mention.id) mention.id = newId('mention');
          records.push({
            item,
            mention,
            row: [
            quarter,
            String(itemIndex + 1),
            item.id || '',
            item.assetClass || '',
            item.fundType || '',
            item.category || '',
            mention.role || '',
            String(mention.fundCode || '').trim().toUpperCase(),
            mention.status || '',
            mention.reason || '',
            Array.isArray(mention.tags) ? mention.tags.join(', ') : '',
            dataAsOf,
            String(item.itemRevision || 1),
            mention.updatedBy || item.updatedBy || fallbackUpdatedBy,
            mention.updatedAt || item.updatedAt || fallbackUpdatedAt,
            mention.id || '',
            String(Number(mention.rowRevision || 1)),
            'FALSE',
          ],
          });
        });
      });
      return records;
    };
    const ensureFslSheetTabAndHeaders = async () => {
      const meta = await SheetsAPI.getSheetTabs(FUND_SELECTION_LOGS_SHEET_ID);
      const createdTab = !meta.tabs.includes(quarter);
      if (createdTab) await SheetsAPI.addSheetTab(FUND_SELECTION_LOGS_SHEET_ID, quarter);
      let rows = createdTab ? [] : await SheetsAPI.fetchSheetData(FUND_SELECTION_LOGS_SHEET_ID, quarter);
      const currentHeaders = (rows[0] || []).map(value => String(value || '').trim()).filter(Boolean);
      const mergedHeaders = currentHeaders.length
        ? [...currentHeaders, ...FUND_SELECTION_LOGS_SHEET_HEADERS.filter(header => !currentHeaders.includes(header))]
        : [...FUND_SELECTION_LOGS_SHEET_HEADERS];
      if (!currentHeaders.length || mergedHeaders.length !== currentHeaders.length) {
        await SheetsAPI.updateSheetRange(
          FUND_SELECTION_LOGS_SHEET_ID,
          `${SheetsAPI._quoteSheetName(quarter)}!A1:${sheetColumnName(mergedHeaders.length)}1`,
          [mergedHeaders],
        );
        rows = await SheetsAPI.fetchSheetData(FUND_SELECTION_LOGS_SHEET_ID, quarter);
      }
      return {
        createdTab,
        headers: mergedHeaders,
        rows,
      };
    };
    const saveFundSelectionRowsToSheet = async (sheetLog) => {
      if (!SheetsAPI.accessToken) await SheetsAPI.requestToken(false);
      const { createdTab, headers, rows: latestRows } = await ensureFslSheetTabAndHeaders();
      const headerCount = headers.length;
      const endColumn = sheetColumnName(headerCount);
      const existingRecords = latestRows
        .slice(1)
        .map((row, index) => fslSheetRecordFromRow(headers, row, index + 2))
        .filter(record => record.quarter.toUpperCase() === quarter && record.mentionId);
      const existingByMentionId = new Map(existingRecords.map(record => [record.mentionId, record]));
      const activeRecords = fslMentionRecordsForSave(sheetLog);
      const activeMentionIds = new Set(activeRecords.map(record => record.mention.id));
      const updates = [];
      const appends = [];
      const conflicts = [];
      let updatedCount = 0;
      let appendedCount = 0;
      let deletedCount = 0;

      activeRecords.forEach(record => {
        const existing = existingByMentionId.get(record.mention.id);
        if (existing && !existing.deleted) {
          const localRevision = Number(record.mention.rowRevision || 1);
          const latestRevision = Number(existing.rowRevision || 1);
          if (latestRevision !== localRevision) {
            conflicts.push({
              mentionId: record.mention.id,
              fundCode: record.mention.fundCode || '',
              latestRevision,
              localRevision,
            });
            return;
          }
          record.mention.rowRevision = latestRevision + 1;
          record.row[FUND_SELECTION_LOGS_SHEET_HEADERS.indexOf('row_revision')] = String(record.mention.rowRevision);
          const alignedRow = alignFslRowToHeaders(record.row, headers);
          updates.push({
            range: `${SheetsAPI._quoteSheetName(quarter)}!A${existing.rowNumber}:${endColumn}${existing.rowNumber}`,
            values: [alignedRow],
          });
          updatedCount += 1;
        } else {
          record.mention.rowRevision = 1;
          record.row[FUND_SELECTION_LOGS_SHEET_HEADERS.indexOf('row_revision')] = '1';
          appends.push(alignFslRowToHeaders(record.row, headers));
          appendedCount += 1;
        }
      });

      if (conflicts.length) {
        const sample = conflicts.slice(0, 5).map(item => item.fundCode || item.mentionId).join(', ');
        throw new Error(`ข้อมูลบางแถวถูกแก้โดยคนอื่นแล้ว (${sample}) กรุณาโหลดใหม่และ review ก่อนบันทึก`);
      }

      existingRecords.forEach(existing => {
        if (existing.deleted || activeMentionIds.has(existing.mentionId)) return;
        const row = [...existing.row];
        while (row.length < headerCount) row.push('');
        row[fslHeaderIndex(headers).deleted] = 'TRUE';
        row[fslHeaderIndex(headers).row_revision] = String(Number(existing.rowRevision || 1) + 1);
        row[fslHeaderIndex(headers).updated_at] = nowIso();
        row[fslHeaderIndex(headers).updated_by] = State.currentUser?.email || sheetLog.updatedBy || '';
        updates.push({
          range: `${SheetsAPI._quoteSheetName(quarter)}!A${existing.rowNumber}:${endColumn}${existing.rowNumber}`,
          values: [row],
        });
        deletedCount += 1;
      });

      if (updates.length) {
        await SheetsAPI.batchUpdateSheetRanges(FUND_SELECTION_LOGS_SHEET_ID, updates);
      }
      if (appends.length) {
        await SheetsAPI.appendSheetValues(FUND_SELECTION_LOGS_SHEET_ID, quarter, appends);
      }

      return {
        createdTab,
        updatedCount,
        appendedCount,
        deletedCount,
      };
    };
    const rowsToFundSelectionSheetValues = (sheetLog) => {
      const rows = [FUND_SELECTION_LOGS_SHEET_HEADERS];
      fslMentionRecordsForSave(sheetLog).forEach(record => rows.push(record.row));
      return rows;
    };
    const loadLog = async () => {
      try {
        const sheetResult = await loadLogFromSheet();
        if (sheetResult.log?.items?.length) return sheetResult;

        const jsonResult = await loadLogFromJsonStore();
        if (jsonResult.log?.items?.length) {
          return {
            ...jsonResult,
            warning: [
              `Google Sheet ${quarter} อ่านได้แต่ยังแปลงเป็นรายการไม่ได้`,
              sheetResult.warning,
              jsonResult.warning,
            ].filter(Boolean).join(' | '),
          };
        }
        return sheetResult;
      } catch (sheetErr) {
        const sheetWarning = `โหลด Google Sheet ไม่สำเร็จ: ${sheetErr.message || sheetErr}`;
        try {
          const jsonResult = await loadLogFromJsonStore();
          return {
            ...jsonResult,
            warning: [sheetWarning, jsonResult.warning].filter(Boolean).join(' | '),
          };
        } catch {
          return {
            log: normalizeLog(null),
            sourceNote: 'New file',
            warning: sheetWarning,
          };
        }
      }
    };
    const driveWarningMessage = (message = '') => {
      const text = String(message || '');
      if (!text) return '';
      if (text.includes('GOOGLE_SERVICE_ACCOUNT_JSON') || text.includes('GOOGLE_APPLICATION_CREDENTIALS')) {
        return 'บันทึกลงไฟล์ในเครื่องแล้ว แต่ยัง sync เข้า Google Drive ไม่ได้: ต้องตั้งค่า GOOGLE_SERVICE_ACCOUNT_JSON_UPLOAD หรือ GOOGLE_APPLICATION_CREDENTIALS ให้ server ก่อน';
      }
      return text;
    };

    let fundSuggestions = [];
    try {
      const selectRows = await fetchCached('select-fund');
      const seen = new Set();
      fundSuggestions = buildSelectedFundsCatalog(selectRows)
        .filter(fund => {
          const key = normalizeFundKey(fund.code);
          if (!key || seen.has(key)) return false;
          seen.add(key);
          return true;
        })
        .sort((a, b) => String(a.code || '').localeCompare(String(b.code || ''), 'en'));
    } catch {
      fundSuggestions = [];
    }
    const fslFundMatches = (value = '') => {
      const q = String(value || '').trim().toLowerCase();
      if (!q) return fundSuggestions.slice(0, 10);
      const nq = normalizeFundKey(q);
      return fundSuggestions
        .filter(fund => (
          normalizeFundKey(fund.code).includes(nq)
          || String(fund.name || '').toLowerCase().includes(q)
          || String(fund.assetHouse || '').toLowerCase().includes(q)
        ))
        .slice(0, 10);
    };

    let { log, sourceNote, warning } = await loadLog();
    let tagSuggestionTimer = null;
    State.fundSelectionLogs.baseRevision = log.revision;
    if (!State.fundSelectionLogs.selectedItemId || !log.items.some(item => item.id === State.fundSelectionLogs.selectedItemId)) {
      State.fundSelectionLogs.selectedItemId = log.items[0]?.id || '';
    }

    const filteredItems = () => {
      const q = String(State.fundSelectionLogs.query || '').trim().toLowerCase();
      const role = State.fundSelectionLogs.roleFilter || '';
      return log.items.filter(item => {
        const mentions = item.mentions || [];
        const roleOk = !role || mentions.some(mention => mention.role === role);
        const qOk = !q || [
          item.assetClass,
          item.fundType,
          item.category,
          ...mentions.flatMap(mention => [mention.fundCode, mention.reason, ...(mention.tags || [])]),
        ].some(value => String(value || '').toLowerCase().includes(q));
        return roleOk && qOk;
      });
    };

    const currentItem = () => log.items.find(item => item.id === State.fundSelectionLogs.selectedItemId) || log.items[0] || null;

    const groupedMentions = (item) => FUND_SELECTION_LOG_ROLES.map(role => ({
      ...role,
      mentions: (item?.mentions || []).filter(mention => mention.role === role.value),
    }));

    const updateSummary = () => {
      const mentionCount = log.items.reduce((sum, item) => sum + (item.mentions || []).length, 0);
      const fundCount = new Set(log.items.flatMap(item => (item.mentions || []).map(m => m.fundCode).filter(Boolean))).size;
      const summary = $('#fsl-summary', area);
      if (!summary) return;
      summary.innerHTML = `
        <div class="fsl-metric"><span>Quarter</span><strong>${esc(quarter)}</strong></div>
        <div class="fsl-metric"><span>หมวด</span><strong>${log.items.length.toLocaleString()}</strong></div>
        <div class="fsl-metric"><span>Mentions</span><strong>${mentionCount.toLocaleString()}</strong></div>
        <div class="fsl-metric"><span>กองทุนที่พูดถึง</span><strong>${fundCount.toLocaleString()}</strong></div>
        <div class="fsl-metric"><span>Revision</span><strong>${Number(log.revision || 0).toLocaleString()}</strong></div>`;
    };

    const renderItemList = () => {
      const items = filteredItems();
      const list = $('#fsl-item-list', area);
      if (!list) return;
      list.innerHTML = items.length ? items.map(item => {
        const active = item.id === currentItem()?.id ? ' is-active' : '';
        const mentions = item.mentions || [];
        const codes = mentions.map(mention => mention.fundCode).filter(Boolean).slice(0, 4);
        return `
          <button class="fsl-item${active}" type="button" data-id="${esc(item.id)}">
            <span class="fsl-item-title">${esc(item.assetClass || 'ยังไม่ระบุหมวด')}</span>
            <span class="fsl-item-meta">${esc(item.fundType || 'กองทั่วไป')} · ${mentions.length.toLocaleString()} mentions</span>
            <span class="fsl-item-codes">${codes.map(code => `<code>${esc(code)}</code>`).join('')}</span>
          </button>`;
      }).join('') : `<div class="fsl-empty">ไม่พบหมวดตามเงื่อนไข</div>`;
    };

    const renderMention = (mention) => `
      <div class="fsl-mention" data-mention-id="${esc(mention.id)}">
        <div class="fsl-mention-grid">
          <label>
            <span>Fund Code</span>
            <input class="fsl-input fsl-code-input fsl-fund-autocomplete" data-mention-field="fundCode" value="${esc(mention.fundCode || '')}" placeholder="เช่น KKP MP" autocomplete="off">
          </label>
          <label>
            <span>Status</span>
            <select class="fsl-input" data-mention-field="status">
              ${statusOptionsFor(mention.status).map(option => `<option value="${esc(option)}" ${String(mention.status || '') === option ? 'selected' : ''}>${option ? esc(option) : 'ไม่ระบุ'}</option>`).join('')}
            </select>
          </label>
        </div>
        <label class="fsl-field-full">
          <span>เหตุผล / ข้อความที่พูดถึง</span>
          <textarea class="fsl-textarea" data-mention-field="reason" rows="1" placeholder="บันทึกเหตุผลเชิงคุณภาพ">${esc(mention.reason || '')}</textarea>
        </label>
        <div class="fsl-mention-foot">
          <label class="fsl-tags-field">
            <span>Tags</span>
            <div class="fsl-tag-editor" data-mention-field="tags">
              <div class="fsl-tag-list">
                ${(mention.tags || []).map(tag => `
                  <button class="fsl-tag-chip" type="button" data-tag="${esc(tag)}" title="ลบ tag ${esc(tag)}">
                    <span>${esc(tag)}</span>
                    <span class="fsl-tag-remove" aria-hidden="true">x</span>
                  </button>
                `).join('')}
              </div>
              <input class="fsl-tag-input" data-tag-input="1" value="" placeholder="พิมพ์ tag แล้วกด Enter">
            </div>
            ${suggestedTagsFor(mention).length ? `
              <div class="fsl-tag-suggestions">
                <span>Suggested</span>
                ${suggestedTagsFor(mention).map(tag => `
                  <button class="fsl-tag-suggestion" type="button" data-tag="${esc(tag)}">+ ${esc(tag)}</button>
                `).join('')}
              </div>
            ` : ''}
          </label>
          <button class="btn btn-xs btn-danger fsl-delete-mention" type="button">ลบ</button>
        </div>
      </div>`;

    const renderDetail = () => {
      const item = currentItem();
      const detail = $('#fsl-detail', area);
      if (!detail) return;
      if (!item) {
        detail.innerHTML = `
          <div class="fsl-slide">
            <div class="fsl-empty">ยังไม่มีหมวด กด “เพิ่มหมวด” เพื่อเริ่มบันทึก</div>
          </div>`;
        return;
      }
      detail.innerHTML = `
        <div class="fsl-slide" data-item-id="${esc(item.id)}">
          <div class="fsl-slide-title-row">
            <div>
              <div class="fsl-slide-kicker">สรุปผลการคัดเลือก</div>
              <input class="fsl-title-input" data-item-field="itemTitle" value="${esc(itemTitleValue(item))}" placeholder="Asset Class / ประเภทกอง เช่น China All Shares EQ / SSF">
            </div>
            <button class="btn btn-danger btn-sm" id="fsl-delete-item" type="button">ลบหมวด</button>
          </div>
          <div class="fsl-item-meta-grid">
            <label>
              <span>ประเภทกอง</span>
              <input class="fsl-input" data-item-field="fundType" value="${esc(item.fundType || '')}" placeholder="กองทั่วไป / RMF / SSF">
            </label>
            <label>
              <span>Category</span>
              <input class="fsl-input" data-item-field="category" value="${esc(item.category || '')}" placeholder="Core / Satellite">
            </label>
            <label>
              <span>Item Revision</span>
              <input class="fsl-input" value="${esc(item.itemRevision || 1)}" disabled>
            </label>
          </div>
          ${groupedMentions(item).map(group => `
            <section class="fsl-section fsl-section-${esc(group.value)}">
              <div class="fsl-section-label">${esc(group.label)}</div>
              <div class="fsl-section-body">
                ${group.mentions.length ? group.mentions.map(renderMention).join('') : `<div class="fsl-empty fsl-empty-inline">ยังไม่มีข้อมูล</div>`}
                <button class="btn btn-ghost btn-sm fsl-add-mention" type="button" data-role="${esc(group.value)}">เพิ่ม ${esc(group.label)}</button>
              </div>
            </section>
          `).join('')}
        </div>`;
    };

    const renderAll = () => {
      updateSummary();
      renderItemList();
      renderDetail();
      const sourceBadge = $('#fsl-source-badge', area);
      if (sourceBadge) sourceBadge.innerHTML = sourceBadgeHtml(pageKey, sourceNote);
      const warningBadge = $('#fsl-warning', area);
      if (warningBadge) {
        warningBadge.innerHTML = warning ? `<span class="badge badge-warning">${esc(warning)}</span>` : '';
      }
      const count = $('#fsl-count', area);
      if (count) count.textContent = `${filteredItems().length.toLocaleString()} / ${log.items.length.toLocaleString()} หมวด`;
    };

    area.innerHTML = `
      <div class="fsl-page">
        <div class="page-tools">
          <div class="page-tools-meta">
            <span id="fsl-source-badge">${sourceBadgeHtml(pageKey, sourceNote)}</span>
            <span class="badge badge-data-origin" id="fsl-count">${log.items.length.toLocaleString()} หมวด</span>
            <span id="fsl-warning">${warning ? `<span class="badge badge-warning">${esc(warning)}</span>` : ''}</span>
          </div>
          <div class="page-tools-actions">
            <a class="btn btn-ghost" id="fsl-open-sheet" href="${sheetUrl}" target="_blank" rel="noopener noreferrer">เปิด Google Sheet</a>
            <button class="btn btn-ghost" id="fsl-add-item" type="button">เพิ่มหมวด</button>
            <button class="btn btn-primary" id="fsl-save-sheet" type="button">บันทึกลง Sheet</button>
            <button class="btn btn-secondary" id="fsl-save" type="button">สำรอง JSON ลง Drive</button>
            <button class="btn btn-ghost" id="fsl-reload" type="button">โหลดใหม่</button>
          </div>
        </div>

        <div class="fsl-summary" id="fsl-summary"></div>

        <div class="card fsl-file-card">
          <div class="card-header">
            <div>
              <div class="card-title">ไฟล์ข้อมูลประจำ Quarter</div>
              <div class="flu-subtitle">บันทึกเป็นไฟล์เดียว: <code>Fund Selection Logs - ${esc(quarter)}.json</code></div>
            </div>
          </div>
          <div class="card-body">
            <div class="fsl-file-grid">
              <label>
                <span>Title</span>
                <input class="fsl-input" id="fsl-title" value="${esc(log.title || '')}">
              </label>
              <label>
                <span>ข้อมูล ณ วันที่</span>
                <input class="fsl-input" id="fsl-data-as-of" value="${esc(log.dataAsOf || '')}" placeholder="2026-03-24">
              </label>
            </div>
          </div>
        </div>

        <div class="fsl-workspace">
          <aside class="fsl-sidebar">
            <div class="fsl-filter">
              <input class="search-input" id="fsl-search" value="${esc(State.fundSelectionLogs.query || '')}" placeholder="ค้นหา fund code, เหตุผล, tag...">
              <select class="filter-select" id="fsl-role-filter">
                <option value="">ทุกหัวข้อ</option>
                ${FUND_SELECTION_LOG_ROLES.map(role => `<option value="${esc(role.value)}" ${State.fundSelectionLogs.roleFilter === role.value ? 'selected' : ''}>${esc(role.label)}</option>`).join('')}
              </select>
            </div>
            <div class="fsl-item-list" id="fsl-item-list"></div>
          </aside>
          <section class="fsl-detail" id="fsl-detail"></section>
        </div>

        <div class="fsl-fund-suggest-menu" id="fsl-fund-suggest-menu" hidden></div>
      </div>`;

    $('#fsl-search', area)?.addEventListener('input', (event) => {
      State.fundSelectionLogs.query = event.target.value;
      renderItemList();
      const count = $('#fsl-count', area);
      if (count) count.textContent = `${filteredItems().length.toLocaleString()} / ${log.items.length.toLocaleString()} หมวด`;
    });
    $('#fsl-role-filter', area)?.addEventListener('change', (event) => {
      State.fundSelectionLogs.roleFilter = event.target.value;
      renderItemList();
      const count = $('#fsl-count', area);
      if (count) count.textContent = `${filteredItems().length.toLocaleString()} / ${log.items.length.toLocaleString()} หมวด`;
    });
    $('#fsl-add-item', area)?.addEventListener('click', () => {
      const item = normalizeItem({
        id: newId(`${quarter.toLowerCase()}-section`),
        assetClass: '',
        fundType: 'กองทั่วไป',
        category: 'Core',
        updatedAt: nowIso(),
        updatedBy: State.currentUser?.email || '',
        mentions: [],
      }, log.items.length);
      log.items.unshift(item);
      State.fundSelectionLogs.selectedItemId = item.id;
      renderAll();
    });
    $('#fsl-reload', area)?.addEventListener('click', () => App.navigate(pageKey));
    $('#fsl-save-sheet', area)?.addEventListener('click', async () => {
      const saveSheetBtn = $('#fsl-save-sheet', area);
      try {
        if (saveSheetBtn) saveSheetBtn.disabled = true;
        log = cleanForSave();
        const result = await saveFundSelectionRowsToSheet(log);
        const createdTabText = result.createdTab ? ' สร้างแท็บใหม่แล้ว' : '';
        sourceNote = `Google Sheet: ${quarter}`;
        warning = '';
        renderAll();
        toast(
          `บันทึกลง Google Sheet แล้ว: update ${result.updatedCount.toLocaleString()}, append ${result.appendedCount.toLocaleString()}, deleted ${result.deletedCount.toLocaleString()}${createdTabText}`,
          'success',
          5200,
        );
      } catch (err) {
        toast(`บันทึกลง Google Sheet ไม่สำเร็จ: ${err.message || err}`, 'error', 7500);
      } finally {
        if (saveSheetBtn) saveSheetBtn.disabled = false;
      }
    });
    $('#fsl-save', area)?.addEventListener('click', async () => {
      try {
        const saveBtn = $('#fsl-save', area);
        if (saveBtn) saveBtn.disabled = true;
        log = cleanForSave();
        let data;
        if (!shouldUseLocalLogsApi() && hasLogsApi()) {
          data = await logsApiRequest('save', {
            quarter,
            updatedBy: State.currentUser?.email || '',
            log,
          });
        } else {
          const resp = await fetch('/api/fund-selection-logs', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              quarter,
              updatedBy: State.currentUser?.email || '',
              log,
            }),
          });
          data = await resp.json().catch(() => ({}));
          if (!resp.ok && resp.status !== 409) throw new Error(data.error || `API ${resp.status}`);
          if (resp.status === 409) data.conflict = true;
        }
        if (data.conflict) {
          toast('มีคนแก้ไขไฟล์นี้ก่อนหน้าแล้ว กรุณาโหลดใหม่เพื่อตรวจข้อมูลล่าสุด', 'warning', 6500);
          warning = `Revision conflict: ไฟล์ล่าสุดเป็น revision ${data.currentRevision}`;
          renderAll();
          return;
        }
        if (data.ok === false) throw new Error(data.error || 'Fund Selection Logs API save failed');
        log = normalizeLog(data.log);
        State.fundSelectionLogs.baseRevision = log.revision;
        sourceNote = data.driveUploaded ? `Google Drive: ${data.drive?.fileName || `Fund Selection Logs - ${quarter}.json`}` : `Local JSON: Fund Selection Logs - ${quarter}.json`;
        warning = driveWarningMessage(data.warning || '');
        renderAll();
        toast(warning || 'สำรอง Fund Selection Logs เป็น JSON บน Google Drive แล้ว', warning ? 'warning' : 'success', warning ? 7500 : 3200);
      } catch (err) {
        toast(`สำรอง JSON ไม่สำเร็จ: ${err.message || err}`, 'error', 6500);
      } finally {
        const saveBtn = $('#fsl-save', area);
        if (saveBtn) saveBtn.disabled = false;
      }
    });
    $('#fsl-item-list', area)?.addEventListener('click', (event) => {
      const itemBtn = event.target.closest('.fsl-item');
      if (!itemBtn) return;
      State.fundSelectionLogs.selectedItemId = itemBtn.dataset.id || '';
      renderAll();
    });
    $('#fsl-detail', area)?.addEventListener('input', (event) => {
      const item = currentItem();
      if (!item) return;
      const itemField = event.target.closest('[data-item-field]');
      if (itemField) {
        if (itemField.dataset.itemField === 'itemTitle') {
          applyItemTitle(item, itemField.value);
          const fundTypeInput = $('#fsl-detail [data-item-field="fundType"]', area);
          if (fundTypeInput) fundTypeInput.value = item.fundType || '';
        } else {
          item[itemField.dataset.itemField] = itemField.value;
          if (itemField.dataset.itemField === 'fundType') {
            const titleInput = $('#fsl-detail [data-item-field="itemTitle"]', area);
            if (titleInput) titleInput.value = itemTitleValue(item);
          }
        }
        item.updatedAt = nowIso();
        item.updatedBy = State.currentUser?.email || item.updatedBy || '';
        if (['assetClass', 'fundType', 'itemTitle'].includes(itemField.dataset.itemField)) renderItemList();
        return;
      }
      const mentionField = event.target.closest('[data-mention-field]');
      if (!mentionField) return;
      const mentionEl = mentionField.closest('[data-mention-id]');
      const mention = item.mentions.find(m => m.id === mentionEl?.dataset.mentionId);
      if (!mention) return;
      if (mentionField.dataset.mentionField === 'tags') {
        return;
      } else if (mentionField.dataset.mentionField === 'fundCode') {
        mention.fundCode = mentionField.value.toUpperCase();
        renderFslFundSuggestMenu(mentionField);
      } else {
        mention[mentionField.dataset.mentionField] = mentionField.value;
      }
      mention.updatedAt = nowIso();
      mention.updatedBy = State.currentUser?.email || mention.updatedBy || '';
      if (mentionField.dataset.mentionField === 'reason') {
        clearTimeout(tagSuggestionTimer);
        tagSuggestionTimer = setTimeout(() => {
          if (State.page === pageKey) renderDetail();
        }, 450);
      }
    });
    $('#fsl-detail', area)?.addEventListener('change', (event) => {
      const item = currentItem();
      if (!item) return;
      const mentionField = event.target.closest('[data-mention-field]');
      if (!mentionField) return;
      const mentionEl = mentionField.closest('[data-mention-id]');
      const mention = item.mentions.find(m => m.id === mentionEl?.dataset.mentionId);
      if (!mention) return;
      if (mentionField.dataset.mentionField === 'tags') {
        return;
      } else if (mentionField.dataset.mentionField === 'fundCode') {
        mention.fundCode = mentionField.value.toUpperCase();
      } else {
        mention[mentionField.dataset.mentionField] = mentionField.value;
      }
      mention.updatedAt = nowIso();
      mention.updatedBy = State.currentUser?.email || mention.updatedBy || '';
    });
    $('#fsl-detail', area)?.addEventListener('click', (event) => {
      const item = currentItem();
      if (!item) return;
      const addBtn = event.target.closest('.fsl-add-mention');
      if (addBtn) {
        item.mentions.push(normalizeMention({
          id: newId('mention'),
          role: addBtn.dataset.role || 'mainChoice',
          sentiment: addBtn.dataset.role === 'notSelected' ? 'negative' : 'neutral',
          updatedAt: nowIso(),
          updatedBy: State.currentUser?.email || '',
        }));
        item.itemRevision = Number(item.itemRevision || 0) + 1;
        renderAll();
        return;
      }
      if (event.target.closest('.fsl-delete-mention')) {
        const mentionEl = event.target.closest('[data-mention-id]');
        item.mentions = item.mentions.filter(mention => mention.id !== mentionEl?.dataset.mentionId);
        item.itemRevision = Number(item.itemRevision || 0) + 1;
        renderAll();
        return;
      }
      const tagChip = event.target.closest('.fsl-tag-chip');
      if (tagChip) {
        const mentionEl = tagChip.closest('[data-mention-id]');
        const mention = item.mentions.find(m => m.id === mentionEl?.dataset.mentionId);
        if (!mention) return;
        const tag = tagChip.dataset.tag || '';
        mention.tags = (mention.tags || []).filter(value => value !== tag);
        mention.updatedAt = nowIso();
        mention.updatedBy = State.currentUser?.email || mention.updatedBy || '';
        renderAll();
        return;
      }
      const tagSuggestion = event.target.closest('.fsl-tag-suggestion');
      if (tagSuggestion) {
        const mentionEl = tagSuggestion.closest('[data-mention-id]');
        const mention = item.mentions.find(m => m.id === mentionEl?.dataset.mentionId);
        if (!mention) return;
        addMentionTag(mention, tagSuggestion.dataset.tag || '');
        renderAll();
        return;
      }
      if (event.target.closest('#fsl-delete-item')) {
        log.items = log.items.filter(row => row.id !== item.id);
        State.fundSelectionLogs.selectedItemId = log.items[0]?.id || '';
        renderAll();
      }
    });
    $('#fsl-detail', area)?.addEventListener('keydown', (event) => {
      const input = event.target.closest('.fsl-tag-input');
      if (!input) return;
      if (!['Enter', ','].includes(event.key)) return;
      event.preventDefault();
      const item = currentItem();
      if (!item) return;
      const mentionEl = input.closest('[data-mention-id]');
      const mention = item.mentions.find(m => m.id === mentionEl?.dataset.mentionId);
      if (!mention) return;
      const tag = String(input.value || '').trim().replace(/,$/, '');
      if (!tag) return;
      addMentionTag(mention, tag);
      input.value = '';
      renderAll();
    });

    const suggestMenu = $('#fsl-fund-suggest-menu', area);
    const hideFslFundSuggestMenu = () => {
      if (!suggestMenu) return;
      suggestMenu.hidden = true;
      suggestMenu.innerHTML = '';
    };
    const applyFslFundSuggestion = (input, fund) => {
      const item = currentItem();
      if (!item || !input || !fund) return;
      const mentionEl = input.closest('[data-mention-id]');
      const mention = item.mentions.find(m => m.id === mentionEl?.dataset.mentionId);
      if (!mention) return;
      input.value = fund.code || '';
      mention.fundCode = fund.code || '';
      mention.updatedAt = nowIso();
      mention.updatedBy = State.currentUser?.email || mention.updatedBy || '';
      hideFslFundSuggestMenu();
      renderItemList();
    };
    const renderFslFundSuggestMenu = (input) => {
      if (!suggestMenu || !input) return;
      const matches = fslFundMatches(input.value);
      if (!matches.length) {
        hideFslFundSuggestMenu();
        return;
      }
      const rect = input.getBoundingClientRect();
      const pageRect = $('.fsl-page', area)?.getBoundingClientRect() || area.getBoundingClientRect();
      suggestMenu.style.left = `${Math.max(8, rect.left - pageRect.left + area.scrollLeft)}px`;
      suggestMenu.style.top = `${rect.bottom - pageRect.top + area.scrollTop + 4}px`;
      suggestMenu.style.width = `${Math.max(360, rect.width)}px`;
      suggestMenu.innerHTML = matches.map(fund => `
        <button class="fsl-fund-suggest-option" type="button" data-code="${esc(fund.code || '')}">
          <strong>${esc(fund.code || '-')}</strong>
          <span>${esc(fund.name || fund.assetHouse || '')}</span>
        </button>
      `).join('');
      suggestMenu.hidden = false;
      suggestMenu.dataset.targetMentionId = input.closest('[data-mention-id]')?.dataset.mentionId || '';
    };
    $('#fsl-detail', area)?.addEventListener('focusin', (event) => {
      const input = event.target.closest('.fsl-fund-autocomplete');
      if (!input) return;
      renderFslFundSuggestMenu(input);
    });
    $('#fsl-detail', area)?.addEventListener('keydown', (event) => {
      const input = event.target.closest('.fsl-fund-autocomplete');
      if (!input) return;
      if (event.key === 'Escape') hideFslFundSuggestMenu();
      if (event.key !== 'Enter' || suggestMenu?.hidden) return;
      const firstOption = suggestMenu.querySelector('.fsl-fund-suggest-option');
      if (!firstOption) return;
      const fund = fundSuggestions.find(item => item.code === firstOption.dataset.code);
      if (fund) {
        event.preventDefault();
        applyFslFundSuggestion(input, fund);
      }
    });
    suggestMenu?.addEventListener('mousedown', (event) => {
      event.preventDefault();
      const option = event.target.closest('.fsl-fund-suggest-option');
      if (!option) return;
      const input = $(`#fsl-detail [data-mention-id="${CSS.escape(suggestMenu.dataset.targetMentionId || '')}"] .fsl-fund-autocomplete`, area);
      const fund = fundSuggestions.find(item => item.code === option.dataset.code);
      if (input && fund) applyFslFundSuggestion(input, fund);
    });
    area.addEventListener('click', (event) => {
      if (event.target.closest('.fsl-fund-autocomplete') || event.target.closest('#fsl-fund-suggest-menu')) return;
      hideFslFundSuggestMenu();
    });

    renderAll();
    App._currentExport = null;
    App._currentTableExport = () => {
      const rows = log.items.flatMap(item => (item.mentions || []).map(mention => [
        log.quarter,
        item.assetClass || '',
        item.fundType || '',
        roleLabel(mention.role),
        mention.fundCode || '',
        mention.reason || '',
        (mention.tags || []).join(', '),
      ]));
      return buildSimpleTablePayload(
        'Fund Selection Logs',
        sourceNote,
        ['Quarter', 'Asset Class', 'Fund Type', 'Role', 'Fund Code', 'Reason', 'Tags'],
        rows,
      );
    };
    App._currentClipboardExport = null;
    App._currentImageExport = null;
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
      await loadFundOverrides();
      const [
        rawSecRowsCurrent,
        rawSecRowsPrevious,
        masterRowsCurrent,
        masterRowsPrevious,
      ] = await Promise.all([
        fetchCachedForTab(pageKey, currentQuarter),
        fetchCachedForTab(pageKey, previousQuarter),
        fetchCachedForTab('master-placeholder-2', currentQuarter),
        fetchCachedForTab('master-placeholder-2', previousQuarter),
      ]);
      const universeCurrent = buildSelectedFeeUniverseFromRows(selectRows, buildMasterRecords(masterRowsCurrent));
      const universePrevious = buildSelectedFeeUniverseFromRows(selectRows, buildMasterRecords(masterRowsPrevious));

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

      const feeRowsCurrent = decorateRows(buildFeeComparisonRows(universeCurrent, buildRawSecLookup(rawSecRowsCurrent), { includeThaiOnly: true }));
      const feeRowsPrevious = decorateRows(buildFeeComparisonRows(universePrevious, buildRawSecLookup(rawSecRowsPrevious), { includeThaiOnly: true }));

      if (!feeRowsCurrent.length && !feeRowsPrevious.length) {
        setError(area, 'ไม่พบข้อมูลค่าธรรมเนียมที่จับคู่ได้จาก Data For SEC API และ AVP Master Fund ID', pageKey);
        return;
      }

      const source = `${CONFIG.PAGES[pageKey]?.source || 'Data For SEC API + AVP Master Fund ID'} · ${currentQuarter} vs ${previousQuarter}`;
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

  async dataImport(area) {
    setLoading(area, 'กำลังเตรียมหน้าเตรียมข้อมูล...');

    const jobList = Object.values(DATA_IMPORT_JOBS);
    const workbookCache = {};

    const rememberImportPreviewScroll = () => {
      const wrap = area.querySelector('.import-preview-table');
      State.dataImport.previewScrollLeft = wrap?.scrollLeft || 0;
    };
    const restoreImportPreviewScroll = () => {
      const left = State.dataImport.previewScrollLeft || 0;
      if (!left) return;
      requestAnimationFrame(() => {
        const wrap = area.querySelector('.import-preview-table');
        if (wrap) wrap.scrollLeft = Math.min(left, wrap.scrollWidth);
      });
    };
    const normalizeFileName = (value) => String(value || '').trim().toLowerCase();
    const findRawFileForJob = (job) => {
      const expected = normalizeFileName(job.sourceFileName);
      return (State.dataImport.rawFiles || []).find(file => normalizeFileName(file.name) === expected)
        || (State.dataImport.rawFiles || []).find(file => normalizeFileName(file.name).includes(expected.replace(/\.xlsx$/i, '')));
    };
	    const isBlankCell = (value) => value === null || value === undefined || String(value).trim() === '';
	    const cleanImportText = (value) => String(value ?? '').replace(/[\u0000-\u001f\u007f-\u009f]/g, '').trim();
	    const withImportTimeout = (promise, message, timeoutMs = 20000) => {
	      let timer;
	      const timeout = new Promise((_, reject) => {
	        timer = setTimeout(() => reject(new Error(message)), timeoutMs);
	      });
	      return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
	    };
	    const transposeRows = (rows = []) => {
      const maxCols = rows.length ? Math.max(...rows.map(row => row.length)) : 0;
      return Array.from({ length: maxCols }, (_, colIdx) => rows.map(row => row[colIdx] ?? ''));
    };
    const rowsToObjects = (rows = []) => {
      const headers = rows[0] || [];
      return rows.slice(1).map(row => {
        const record = {};
        headers.forEach((header, idx) => {
          const key = String(header || `Column${idx + 1}`);
          record[key] = row[idx] ?? '';
        });
        return record;
      });
    };
    const objectsToRows = (records = []) => {
      if (!records.length) return [];
      const headers = [];
      records.forEach(record => {
        Object.keys(record).forEach(key => {
          if (!headers.includes(key)) headers.push(key);
        });
      });
      return [
        headers,
        ...records.map(record => headers.map(header => record[header] ?? '')),
      ];
    };
    const IMPORT_HEADER_TERMS = [
      'group/investment',
      'fund code',
      'fundid',
      'fund id',
      'isin',
      'master fund',
      'return(',
      'asset alloc',
    ];
    const IMPORT_CONTEXT_RE = /^(?:1M|3M|6M|YTD|1Y|3Y|5Y|10Y|20Y|\d{4})$/i;
    const isPerformanceImportHeader = (header) => (
      /return|drawdown|percentile|ratio|std dev|treynor|sortino|sharpe|information/i.test(header)
    );
    const fillMergedImportContext = (row = [], maxCols = row.length) => {
      let currentContext = '';
      return Array.from({ length: maxCols }, (_, idx) => {
        const value = cleanImportText(row[idx]);
        if (IMPORT_CONTEXT_RE.test(value)) currentContext = value.toUpperCase();
        else if (value && !IMPORT_CONTEXT_RE.test(value)) currentContext = '';
        return currentContext;
      });
    };
    const headerHasImportContext = (header) => (
      /\b(?:1M|3M|6M|YTD|1Y|3Y|5Y|10Y|20Y|\d{4})\b/i.test(header)
    );
    const appendContextToImportHeaders = (headers = [], contextRow = []) => {
      const contexts = fillMergedImportContext(contextRow, headers.length);
      return headers.map((header, idx) => {
        const cleanHeader = cleanImportText(header) || `Column ${idx + 1}`;
        const context = contexts[idx];
        if (!context || !isPerformanceImportHeader(cleanHeader)) return cleanHeader;
        if (headerHasImportContext(cleanHeader)) return cleanHeader;
        return `${cleanHeader} ${context}`;
      });
    };
    const contextualizeImportHeaders = (rows = [], externalContextRow = null) => {
      if (!rows.length) return { headers: [], dataStartIdx: 1 };

      const baseHeaders = rows[0] || [];
      const maxCols = Math.max(
        ...rows.map(row => row.length),
        externalContextRow?.length || 0
      );
      const cleanedBase = Array.from({ length: maxCols }, (_, idx) => (
        cleanImportText(baseHeaders[idx]) || `Column ${idx + 1}`
      ));
      const duplicates = cleanedBase.reduce((map, header) => {
        map.set(header, (map.get(header) || 0) + 1);
        return map;
      }, new Map());
      const duplicateCols = new Set(
        cleanedBase
          .map((header, idx) => (duplicates.get(header) > 1 ? idx : -1))
          .filter(idx => idx >= 0)
      );

      if (externalContextRow) {
        const headers = appendContextToImportHeaders(cleanedBase, externalContextRow);
        return { headers, dataStartIdx: 1 };
      }

      if (!duplicateCols.size) {
        return { headers: cleanedBase, dataStartIdx: 1 };
      }

      let bestContextIdx = -1;
      let bestScore = 0;
      rows.slice(1, 7).forEach((row, relIdx) => {
        let score = 0;
        duplicateCols.forEach(colIdx => {
          const context = cleanImportText(row[colIdx]);
          if (!context) return;
          if (IMPORT_CONTEXT_RE.test(context)) score += 3;
          else if (/return|drawdown|percentile|ratio|std dev|fee|asset|holding/i.test(context)) score += 2;
          else if (context.length <= 24) score += 1;
        });
        if (score > bestScore) {
          bestScore = score;
          bestContextIdx = relIdx + 1;
        }
      });

      if (bestContextIdx < 0 || bestScore < 3) {
        return { headers: cleanedBase, dataStartIdx: 1 };
      }

      const contextRow = rows[bestContextIdx] || [];
      const headers = cleanedBase.map((header, idx) => {
        const context = cleanImportText(contextRow[idx]);
        if (!duplicateCols.has(idx) || !context) return header;
        if (header.toLowerCase().includes(context.toLowerCase())) return header;
        return `${header} ${context}`;
      });

      return { headers, dataStartIdx: bestContextIdx + 1 };
    };
    const normalizeImportedTableRows = (rawRows = [], options = {}) => {
      if (!rawRows.length) return [];

      const rowHasValue = (row = []) => row.some(cell => !isBlankCell(cell));
      const trimmedRows = rawRows.filter(rowHasValue);
      if (!trimmedRows.length) return [];

      const scoreHeaderRow = (row = [], idx) => {
        const cells = row.map(cell => cleanImportText(cell)).filter(Boolean);
        const lower = cells.map(cell => cell.toLowerCase());
        const keywordScore = IMPORT_HEADER_TERMS.reduce((score, term) => (
          score + (lower.some(cell => cell.includes(term)) ? 20 : 0)
        ), 0);
        const textScore = cells.filter(cell => Number.isNaN(Number(cell.replace(/,/g, '')))).length;
        return (cells.length * 2) + keywordScore + textScore - (idx * 0.25);
      };
      const headerIdx = trimmedRows
        .slice(0, 30)
        .reduce((bestIdx, row, idx, sample) => (
          scoreHeaderRow(row, idx) > scoreHeaderRow(sample[bestIdx], bestIdx) ? idx : bestIdx
        ), 0);
      const dataRows = trimmedRows.slice(headerIdx);
      const contextRow = Number.isInteger(options.contextRowIdx)
        ? rawRows[options.contextRowIdx]
        : null;
      const { headers, dataStartIdx } = contextualizeImportHeaders(dataRows, contextRow);

      return [
        headers,
        ...dataRows.slice(dataStartIdx)
          .filter(rowHasValue)
          .map(row => headers.map((_, idx) => row[idx] ?? '')),
      ];
    };
    const hasRecognizableImportHeader = (rows = []) => {
      const headerText = (rows[0] || [])
        .map(cell => cleanImportText(cell).toLowerCase())
        .filter(Boolean);
      const matches = IMPORT_HEADER_TERMS.filter(term => headerText.some(cell => cell.includes(term)));
      return matches.length >= 2;
    };
    const cleanMorningstarExportRows = (rawRows = []) => {
      const periodIdx = rawRows.slice(0, 15).findIndex(row => {
        const periodCount = row.filter(value => IMPORT_CONTEXT_RE.test(cleanImportText(value))).length;
        return periodCount >= 3;
      });
      if (periodIdx < 0) return [];

      const headerOffset = rawRows.slice(periodIdx + 1, periodIdx + 8).findIndex(row => {
        const text = row.map(value => cleanImportText(value).toLowerCase()).join('|');
        return /group\/investment|fund code|fundid/.test(text)
          && /return|drawdown|percentile|ratio|std dev/i.test(text);
      });
      if (headerOffset < 0) return [];

      const headerIdx = periodIdx + 1 + headerOffset;
      const periodRow = rawRows[periodIdx] || [];
      const headerRow = rawRows[headerIdx] || [];
      const maxCols = Math.max(periodRow.length, headerRow.length, ...rawRows.slice(headerIdx + 1).map(row => row.length));

      const baseHeaders = Array.from({ length: maxCols }, (_, idx) => cleanImportText(headerRow[idx]));
      const headers = appendContextToImportHeaders(baseHeaders, periodRow);
      const dataRows = rawRows
        .slice(headerIdx + 1)
        .filter(row => row.some(cell => !isBlankCell(cell)))
        .filter(row => !isBlankCell(row[1]))
        .map(row => headers.map((_, idx) => row[idx] ?? ''));

      const keepColumn = (idx) => {
        if (cleanImportText(baseHeaders[idx])) return true;
        return dataRows.some(row => !isBlankCell(row[idx]));
      };
      const keepIndexes = headers
        .map((_, idx) => idx)
        .filter(keepColumn);
      if (!keepIndexes.length || !dataRows.length) return [];

      return [
        keepIndexes.map(idx => cleanImportText(headers[idx]) || `Column ${idx + 1}`),
        ...dataRows.map(row => keepIndexes.map(idx => row[idx] ?? '')),
      ];
    };
    const detectMorningstarRows = (rawRows = []) => {
      const periodIdx = rawRows.slice(0, 15).findIndex(row => {
        const periodCount = row.filter(value => IMPORT_CONTEXT_RE.test(cleanImportText(value))).length;
        return periodCount >= 3;
      });
      const headerOffset = periodIdx >= 0
        ? rawRows.slice(periodIdx + 1, periodIdx + 8).findIndex(row => {
            const text = row.map(value => cleanImportText(value).toLowerCase()).join('|');
            return /group\/investment|fund code|fundid/.test(text)
              && /return|drawdown|percentile|ratio|std dev/i.test(text);
          })
        : -1;
      return {
        periodIdx,
        headerIdx: periodIdx >= 0 && headerOffset >= 0 ? periodIdx + 1 + headerOffset : -1,
      };
    };
    const cleanMasterFundRawRows = (rawRows = []) => {
      if (!rawRows.length) return [];
      const morningstarRows = cleanMorningstarExportRows(rawRows);
      if (morningstarRows.length > 1 && hasRecognizableImportHeader(morningstarRows)) {
        return morningstarRows;
      }
      const fallbackRows = () => normalizeImportedTableRows(rawRows, { contextRowIdx: 6 });

      // Mirrors the workbook's Power Query: promote first row, skip the next 5 metadata rows.
      // In array terms this starts from original row 7 because row 1 became headers.
      const skipped = rawRows.slice(6);
      let transposed = transposeRows(skipped);
      transposed = transposed.filter(row => !isBlankCell(row[3]));
      transposed = transposed.map(row => row.filter((_, idx) => idx !== 1 && idx !== 2));

      let currentGroup = '';
      transposed = transposed.map(row => {
        const next = [...row];
        if (!isBlankCell(next[0])) currentGroup = next[0];
        next[0] = currentGroup;
        return next;
      });
      const preparedContextRow = transposed.map(row => row[0] ?? '');

      transposed = transposed.map(row => {
        // After removing Column2 and Column3, Power Query's Column4 is now array index 1.
        const column4 = cleanImportText(row[1]);
        const column1 = cleanImportText(row[0]);
        const merged = cleanImportText([column4, column1].filter(Boolean).join(' '));
        return [merged, ...row.slice(2)];
      });

      const seen = new Set();
      transposed = transposed.filter(row => {
        const key = String(row[0] ?? '');
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });

      let cleaned = transposeRows(transposed);
      cleaned = cleaned.filter(row => !isBlankCell(row[1]));
      if (!cleaned.length) return fallbackRows();

      const headers = cleaned[0].map((header, idx) => cleanImportText(header) || `Column ${idx + 1}`);
      const transformedRows = [
        headers,
        ...cleaned.slice(1).map(row => headers.map((_, idx) => row[idx] ?? '')),
      ];
      const contextualizedRows = [
        appendContextToImportHeaders(headers, preparedContextRow),
        ...transformedRows.slice(1),
      ];
      return contextualizedRows.length > 1 && hasRecognizableImportHeader(contextualizedRows)
        ? contextualizedRows
        : fallbackRows();
    };
    const cleanRowsForJob = (job, rows) => {
      if (job.cleaner === 'masterFundRaw') return cleanMasterFundRawRows(rows);
      return rows;
    };
    const readXlsxWorkbook = async (job) => {
      if (workbookCache[job.key]) return workbookCache[job.key];
      if (typeof XLSX === 'undefined') {
        throw new Error('ยังโหลดตัวอ่าน Excel ไม่สำเร็จ กรุณารีเฟรชหน้าเว็บแล้วลองอีกครั้ง');
      }
      const file = findRawFileForJob(job);
      if (!file?.id) {
        throw new Error(`ไม่พบไฟล์ ${job.sourceFileName} ใน Raw Files Folder`);
      }
      const buffer = await SheetsAPI.downloadDriveFile(file.id);
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
      workbookCache[job.key] = { file, workbook };
      return workbookCache[job.key];
    };
    const getRowsFromXlsx = async (job, tabName) => {
      const { workbook } = await readXlsxWorkbook(job);
      const sheetName = workbook.SheetNames.includes(tabName) ? tabName : workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) return [];
      return XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: '',
        blankrows: false,
      });
    };
    const loadJobMeta = async (job) => {
      const targetMeta = await SheetsAPI.getSheetTabs(job.targetSheetId);
      if (job.sourceType === 'xlsx') {
        const { file, workbook } = await readXlsxWorkbook(job);
        return {
          job,
          targetMeta,
          rawMeta: {
            title: file.name || job.sourceLabel,
            tabs: workbook.SheetNames,
            sheetTabs: workbook.SheetNames.map(title => ({ title, sheetId: null })),
          },
        };
      }
      const rawMeta = await SheetsAPI.getSheetTabs(job.rawSheetId);
      return { job, rawMeta, targetMeta };
    };

	    try {
	      State.dataImport.rawFiles = await withImportTimeout(
	        SheetsAPI.listDriveFolderFiles(RAW_FILES_FOLDER_ID),
	        'โหลดรายการไฟล์ Raw Files Folder นานเกินไป กรุณาเข้าสู่ระบบใหม่หรือรีเฟรชหน้า'
	      );
      State.dataImport.rawFilesError = '';
    } catch (e) {
      State.dataImport.rawFiles = [];
      State.dataImport.rawFilesError = e.message || String(e);
    }

    if (!DATA_IMPORT_JOBS[State.dataImport.selectedJobKey]) {
      State.dataImport.selectedJobKey = 'percentrank';
    }

    let active;
    try {
	      active = await withImportTimeout(
	        loadJobMeta(DATA_IMPORT_JOBS[State.dataImport.selectedJobKey]),
	        'โหลด metadata ของไฟล์ต้นทาง/ปลายทางนานเกินไป กรุณาลองใหม่อีกครั้ง'
	      );
    } catch (e) {
      setError(area, e.message, 'data-import');
      return;
    }

    const applyJobDefaults = (job, rawMeta) => {
      const rawTabs = rawMeta.sheetTabs || (rawMeta.tabs || []).map(title => ({ title, sheetId: null }));
      const defaultRawTab = rawTabs.find(tab => Number(tab.sheetId) === Number(job.rawGid))?.title
        || rawTabs[0]?.title
        || '';
      const validRawTab = rawTabs.some(tab => tab.title === State.dataImport.rawTab);
      const defaultTargetTab = State.currentQuarter
        || CONFIG.PAGES?.[job.targetPageKey]?.tabName
        || job.defaultTab;

      if (!validRawTab) State.dataImport.rawTab = defaultRawTab;
      if (!State.dataImport.targetTab) State.dataImport.targetTab = defaultTargetTab;
    };
    const getQuarterYear = (quarter) => {
      const match = String(quarter || '').trim().match(/^(\d{4})-Q[1-4]$/i);
      return match?.[1] || 'unknown-year';
    };
    const getJsonBasePath = (job, quarter) => {
      const cleanQuarter = String(quarter || '').trim() || job.defaultTab;
      const fileName = JSON_STORE.baseFiles[job.key] || `${job.targetLabel || job.label}.json`;
      return `${JSON_STORE.rootName}/${getQuarterYear(cleanQuarter)}/${cleanQuarter}/base/${fileName}`;
    };
    const getJsonOverridesPath = (quarter) => {
      const cleanQuarter = String(quarter || '').trim() || State.currentQuarter || 'unknown-quarter';
      return `${JSON_STORE.rootName}/${getQuarterYear(cleanQuarter)}/${cleanQuarter}/overrides/`;
    };
    const quoteImportSheetName = (tabName) => `'${String(tabName || '').replace(/'/g, "''")}'`;
    const fetchImportHeaderPreview = async (sheetId, tabName) => {
      const range = encodeURIComponent(`${quoteImportSheetName(tabName)}!A1:ZZ1`);
      const params = 'valueRenderOption=FORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING';
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}?${params}`;
      const data = await SheetsAPI._fetchGoogleJson(url);
      return data.values?.[0] || [];
    };
    const checkJsonReadiness = async (quarter) => {
      const cleanQuarter = String(quarter || '').trim().toUpperCase();
      const results = await Promise.all(DATA_IMPORT_JSON_EXPORTS.map(async item => {
        const importJob = DATA_IMPORT_JOBS[item.jobKey];
        try {
          const meta = await SheetsAPI.getSheetTabs(importJob.targetSheetId);
          const hasTab = (meta.tabs || []).includes(cleanQuarter);
          if (!hasTab) {
            return {
              ...item,
              label: importJob.targetLabel || importJob.label,
              sheetTitle: meta.title || importJob.targetLabel || importJob.label,
              ok: false,
              status: 'missing-tab',
              message: `ไม่พบ Tab ${cleanQuarter}`,
              columns: 0,
            };
          }
          const header = await fetchImportHeaderPreview(importJob.targetSheetId, cleanQuarter);
          const nonBlankColumns = header.filter(value => String(value ?? '').trim()).length;
          return {
            ...item,
            label: importJob.targetLabel || importJob.label,
            sheetTitle: meta.title || importJob.targetLabel || importJob.label,
            ok: nonBlankColumns > 0,
            status: nonBlankColumns > 0 ? 'ready' : 'empty',
            message: nonBlankColumns > 0 ? `พร้อมสร้าง JSON (${nonBlankColumns.toLocaleString()} columns)` : 'พบ Tab แล้ว แต่หัวตารางว่าง',
            columns: nonBlankColumns,
          };
        } catch (err) {
          return {
            ...item,
            label: importJob.targetLabel || importJob.label,
            sheetTitle: importJob.targetLabel || importJob.label,
            ok: false,
            status: 'error',
            message: err.message || String(err),
            columns: 0,
          };
        }
      }));
      return {
        quarter: cleanQuarter,
        checkedAt: new Date().toISOString(),
        results,
        allReady: results.every(item => item.ok),
      };
    };
    const buildJsonExportUrl = (quarter, dataset = 'all') => {
      const url = new URL(JSON_EXPORT_WEB_APP_URL);
      url.searchParams.set('quarter', String(quarter || '').trim().toUpperCase());
      url.searchParams.set('dataset', dataset);
      url.searchParams.set('key', JSON_EXPORT_SECRET_KEY);
      return url.toString();
    };
    const exportBaseJsonFromSheets = async (quarter) => {
      const cleanQuarter = String(quarter || '').trim().toUpperCase();
      const targetFolder = await SheetsAPI.resolveDriveFolderPath(JSON_DRIVE_ROOT_FOLDER_ID, [
        getQuarterYear(cleanQuarter),
        cleanQuarter,
        'base',
      ]);
      await SheetsAPI.resolveDriveFolderPath(JSON_DRIVE_ROOT_FOLDER_ID, [
        getQuarterYear(cleanQuarter),
        cleanQuarter,
        'overrides',
      ]);

      const files = [];
      for (const item of DATA_IMPORT_JSON_EXPORTS) {
        const importJob = DATA_IMPORT_JOBS[item.jobKey];
        const rows = await SheetsAPI.fetchSheetData(importJob.targetSheetId, cleanQuarter);
        if (!rows.length) throw new Error(`ไม่พบข้อมูลใน ${importJob.targetLabel || importJob.label} / ${cleanQuarter}`);
        const file = await SheetsAPI.uploadJsonToDriveFolder(targetFolder.id, item.fileName, rows);
        files.push({
          ...item,
          rows: rows.length,
          columns: Math.max(...rows.map(row => row.length), 0),
          fileId: file.id,
          fileName: file.name || item.fileName,
          webViewLink: file.webViewLink || '',
        });
      }
      return {
        ok: true,
        quarter: cleanQuarter,
        folderId: targetFolder.id,
        folderLink: targetFolder.webViewLink || `https://drive.google.com/drive/folders/${targetFolder.id}`,
        files,
      };
    };
    const renderJsonReadiness = (readiness) => {
      const quarter = String(State.dataImport.targetTab || '').trim().toUpperCase() || '-';
      const jsonBaseFolder = `${JSON_STORE.rootName}/${getQuarterYear(quarter)}/${quarter}/base/`;
      const jsonOverridesPath = getJsonOverridesPath(quarter);
      const results = readiness?.results || DATA_IMPORT_JSON_EXPORTS.map(item => {
        const importJob = DATA_IMPORT_JOBS[item.jobKey];
        return {
          ...item,
          label: importJob.targetLabel || importJob.label,
          sheetTitle: importJob.targetLabel || importJob.label,
          ok: false,
          status: 'pending',
          message: 'ยังไม่ได้ตรวจสอบ',
        };
      });
      return `
        <div class="json-readiness-box">
          <div class="json-readiness-copy">
            ตรวจสอบก่อนว่า Google Sheets ปลายทางมี Tab ${esc(quarter)} ครบทั้ง 3 ชุด แล้วค่อยสร้างไฟล์ JSON ลงโฟลเดอร์ base ของ Quarter นี้
          </div>
          <div class="json-path-grid">
            <div class="fund-field data-import-path-field">
              <span>Base JSON Folder</span>
              <input class="fund-input" value="${esc(jsonBaseFolder)}" readonly>
            </div>
            <div class="fund-field data-import-path-field">
              <span>Overrides ของ Quarter นี้</span>
              <input class="fund-input" value="${esc(jsonOverridesPath)}" readonly>
            </div>
          </div>
          <div class="json-readiness-list">
            ${results.map(item => `
              <div class="json-readiness-item ${item.ok ? 'is-ready' : `is-${esc(item.status || 'pending')}`}">
                <div>
                  <strong>${esc(item.fileName)}</strong>
                  <span>${esc(item.sheetTitle)} · dataset=${esc(item.dataset)}</span>
                </div>
                <span class="badge ${item.ok ? 'badge-success' : (item.status === 'pending' ? 'badge-muted' : 'badge-warning')}">${esc(item.message)}</span>
              </div>
            `).join('')}
          </div>
          <div class="data-import-actions json-export-actions">
            <button class="btn btn-ghost" id="json-readiness-check" type="button" ${State.dataImport.isCheckingJson ? 'disabled' : ''}>
              ${State.dataImport.isCheckingJson ? 'กำลังตรวจสอบ...' : 'ตรวจสอบความพร้อม'}
            </button>
            <button class="btn btn-primary import-commit-btn" id="json-export-open" type="button" ${readiness?.allReady && !State.dataImport.isExportingJson ? '' : 'disabled'}>
              ${State.dataImport.isExportingJson ? 'กำลังสร้าง JSON...' : 'สร้าง JSON ทั้ง 3 ไฟล์'}
            </button>
            ${readiness && !readiness.allReady ? '<span class="json-export-hint">ต้องพร้อมครบทั้ง 3 ชุดก่อนจึงจะสร้าง JSON ได้</span>' : ''}
          </div>
        </div>`;
    };

    applyJobDefaults(active.job, active.rawMeta);

    const detectImportValueType = (value) => {
      const text = String(value ?? '').trim();
      if (!text) return '';
      if (/^[-–—]+$/.test(text)) return '';
      if (/^(n\/a|na|null|none)$/i.test(text)) return '';
      if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(text)) return 'date';
      if (/^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(text)) return 'date';
      if (/^-?\d+(,\d{3})*(\.\d+)?%$/.test(text) || /^-?\d+(\.\d+)?%$/.test(text)) return 'percent';
      if (/^-?\d+(,\d{3})*(\.\d+)?$/.test(text) || /^-?\d+(\.\d+)?$/.test(text)) return 'number';
      return 'text';
    };
    const analyzeImportColumns = (rows = []) => {
      const headers = rows[0] || [];
      const maxCols = rows.length ? Math.max(...rows.map(row => row.length)) : headers.length;
      return Array.from({ length: maxCols }, (_, colIdx) => {
        const counts = { text: 0, number: 0, percent: 0, date: 0 };
        rows.slice(1, 101).forEach(row => {
          const type = detectImportValueType(row[colIdx]);
          if (type) counts[type] += 1;
        });
        const used = Object.entries(counts).filter(([, count]) => count > 0);
        const type = used.length === 0
          ? 'blank'
          : (used.length === 1 ? used[0][0] : 'mixed');
        const key = `${colIdx}:${headers[colIdx] || `Column ${colIdx + 1}`}`;
        return {
          key,
          index: colIdx + 1,
          name: headers[colIdx] || `Column ${colIdx + 1}`,
          detectedType: type,
          type: State.dataImport.columnTypes[key] || type,
          counts,
        };
      });
    };
    const normalizeImportHeader = (value) => cleanImportText(value)
      .replace(/\s+/g, ' ')
      .replace(/\s*\(\s*/g, '(')
      .replace(/\s*\)\s*/g, ')')
      .toLowerCase();
    const analyzeImportHeaders = (rows = []) => {
      const headers = rows[0] || [];
      const maxCols = rows.length ? Math.max(...rows.map(row => row.length)) : headers.length;
      const exact = new Map();
      const normalized = new Map();

      Array.from({ length: maxCols }, (_, colIdx) => {
        const fallback = `Column ${colIdx + 1}`;
        const name = cleanImportText(headers[colIdx]) || fallback;
        const exactKey = name;
        const normalizedKey = normalizeImportHeader(name);
        exact.set(exactKey, [...(exact.get(exactKey) || []), { name, index: colIdx + 1 }]);
        normalized.set(normalizedKey, [...(normalized.get(normalizedKey) || []), { name, index: colIdx + 1 }]);
      });

      const exactDuplicates = [...exact.values()]
        .filter(items => items.length > 1)
        .map(items => ({
          name: items[0].name,
          columns: items.map(item => item.index),
          suggestion: items.map((item, idx) => idx === 0 ? item.name : `${item.name}_${idx + 1}`),
        }));
      const normalizedDuplicates = [...normalized.values()]
        .filter(items => items.length > 1 && new Set(items.map(item => item.name)).size > 1)
        .map(items => ({
          normalizedName: normalizeImportHeader(items[0].name),
          variants: items.map(item => ({ name: item.name, index: item.index })),
        }));

      return {
        totalHeaders: maxCols,
        duplicateCount: exactDuplicates.length,
        similarCount: normalizedDuplicates.length,
        exactDuplicates,
        normalizedDuplicates,
      };
    };
    const typeBadgeClass = (type) => ({
      number: 'badge-success',
      percent: 'badge-accent',
      date: 'badge-primary',
      text: 'badge-data-origin',
      mixed: 'badge-warning',
      blank: 'badge-muted',
    }[type] || 'badge-data-origin');
    const typeLabel = (type) => {
      const labels = {
        number: 'Number',
        percent: 'Percent',
        date: 'Date',
        text: 'Text',
        mixed: 'Mixed',
        blank: 'Blank',
      };
      return labels[type] || String(type || '').replace(/^./, c => c.toUpperCase());
    };
    const buildImportColumnKey = (colIdx, headers = []) => `${colIdx}:${headers[colIdx] || `Column ${colIdx + 1}`}`;
    const refreshImportPreviewMeta = (preview, rows, patch = {}) => {
      const colCount = rows.length ? Math.max(...rows.map(row => row.length)) : 0;
      return {
        ...preview,
        ...patch,
        rows,
        rowCount: rows.length,
        colCount,
        headerText: (rows[0] || []).slice(0, 6).join(' | '),
        columns: analyzeImportColumns(rows),
        headerReport: analyzeImportHeaders(rows),
      };
    };
    const renameDuplicateImportHeaders = (rows = []) => {
      if (!rows.length) return rows;
      const headers = [...(rows[0] || [])];
      const seen = new Map();
      const renamed = headers.map((header, idx) => {
        const base = cleanImportText(header) || `Column ${idx + 1}`;
        const count = (seen.get(base) || 0) + 1;
        seen.set(base, count);
        return count === 1 ? base : `${base}_${count}`;
      });
      return [renamed, ...rows.slice(1)];
    };
    const removeImportColumns = (rows = [], columnIndexes = []) => {
      const removeSet = new Set(columnIndexes.map(Number).filter(Number.isInteger));
      if (!rows.length || !removeSet.size) return rows;
      return rows.map(row => row.filter((_, idx) => !removeSet.has(idx)));
    };
    const importTypeSelect = (col) => {
      const typeOptions = [
        ['auto', 'Auto'],
        ['text', 'Text'],
        ['number', 'Number'],
        ['percent', 'Percent'],
        ['date', 'Date'],
        ['blank', 'Blank'],
      ];
      return `
        <div class="import-header-type">
          <strong class="badge ${typeBadgeClass(col.type)}">${esc(typeLabel(col.type))}</strong>
          <select class="import-type-select" data-column-key="${esc(col.key)}" aria-label="เลือกชนิดข้อมูลของ ${esc(col.name)}">
            ${typeOptions.map(([value, label]) => `
              <option value="${value}" ${(State.dataImport.columnTypes[col.key] || 'auto') === value ? 'selected' : ''}>
                ${label}${value === 'auto' ? ` (${typeLabel(col.detectedType)})` : ''}
              </option>`).join('')}
          </select>
        </div>`;
    };

    const sampleTable = (rows = []) => {
      if (!rows.length) return '<div class="state-box">ไม่พบข้อมูลสำหรับ Preview</div>';
      const headers = rows[0] || [];
      const dataRows = rows.slice(1);
      const visibleRows = [headers, ...dataRows.slice(0, 5)];
      const maxCols = Math.max(...rows.map(row => row.length));
      const columns = State.dataImport.preview?.columns || analyzeImportColumns(rows);
      const duplicateHeaderColumns = new Set(
        (State.dataImport.preview?.headerReport?.exactDuplicates || [])
          .flatMap(item => item.columns || [])
          .map(col => Number(col) - 1)
      );
      const deleteHeaderButton = (colIdx) => duplicateHeaderColumns.has(colIdx)
        ? `<button class="import-delete-column-btn" type="button" data-delete-column="${colIdx}" title="ลบคอลัมน์นี้ออกจาก Preview">ลบคอลัมน์นี้</button>`
        : '';
      return `
        <div class="table-wrapper import-preview-table">
          <table>
            <tbody>
              ${visibleRows.map((row, rowIdx) => `
                <tr class="${rowIdx === 0 ? 'is-header-row' : ''}">
                  ${Array.from({ length: maxCols }, (_, i) => rowIdx === 0
                    ? `<td class="${duplicateHeaderColumns.has(i) ? 'is-duplicate-header' : ''}"><div class="import-header-cell"><span>${esc(row[i] ?? `Column ${i + 1}`)}</span>${importTypeSelect(columns[i] || { key: buildImportColumnKey(i, headers), name: `Column ${i + 1}`, type: 'blank', detectedType: 'blank' })}${deleteHeaderButton(i)}</div></td>`
                    : `<td>${esc(row[i] ?? '')}</td>`).join('')}
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
        <div class="pg-info import-preview-note">แสดงทุกคอลัมน์ และตัวอย่าง 5 แถวแรกจาก ${dataRows.length.toLocaleString()} rows</div>`;
    };
    const renderImportDebug = (debug) => {
      if (!debug) return '';
      const renderRow = (label, row = []) => `
        <div class="import-header-report-item">
          <span>${esc(label)}</span>
          <small>${esc((row || []).map((value, idx) => `${idx + 1}:${cleanImportText(value) || '-'}`).join(' | '))}</small>
        </div>`;
      return `
        <div class="import-header-report is-warning">
          <div class="import-header-report-head">
            <strong>Debug วิธีเตรียมข้อมูล</strong>
            <span class="badge badge-warning">${esc(debug.cleaner || 'raw')}</span>
          </div>
          <div class="import-header-report-list">
            ${debug.detected
              ? `<div class="import-header-report-item"><span>Detected rows</span><small>period = raw row ${debug.detected.periodIdx >= 0 ? debug.detected.periodIdx + 1 : '-'} | header = raw row ${debug.detected.headerIdx >= 0 ? debug.detected.headerIdx + 1 : '-'}</small></div>`
              : ''}
            ${renderRow('Raw row 1 ที่เว็บอ่านได้', debug.rawRow1)}
            ${renderRow('Raw row 6 ที่เว็บอ่านได้', debug.rawRow6)}
            ${renderRow('Raw row 7 ที่เว็บอ่านได้', debug.rawRow7)}
            ${renderRow('Raw row 8 ที่เว็บอ่านได้', debug.rawRow8)}
            ${renderRow('Header หลัง clean ที่ใช้จริง', debug.cleanedHeader)}
          </div>
        </div>`;
    };
    const renderHeaderReport = (report) => {
      if (!report) return '';
      const hasExact = report.exactDuplicates?.length > 0;
      const hasSimilar = report.normalizedDuplicates?.length > 0;
      const statusClass = hasExact ? 'is-danger' : (hasSimilar ? 'is-warning' : 'is-ok');
      const statusText = hasExact
        ? `พบหัวตารางซ้ำ ${report.duplicateCount.toLocaleString()} รายการ`
        : (hasSimilar ? `พบชื่อคล้ายกัน ${report.similarCount.toLocaleString()} รายการ` : 'ไม่พบหัวตารางซ้ำ');
      return `
        <div class="import-header-report ${statusClass}">
          <div class="import-header-report-head">
            <strong>ตรวจหัวตาราง</strong>
            <span class="badge ${hasExact ? 'badge-danger' : (hasSimilar ? 'badge-warning' : 'badge-success')}">${esc(statusText)}</span>
          </div>
          ${hasExact ? `
            <div class="import-header-actions">
              <span class="import-header-action-copy">
                ตอนนี้พบหัวตารางที่มีชื่อซ้ำ สามารถเลือกที่จะเปลี่ยนชื่อซ้ำอัตโนมัติ หรือลบคอลัมน์ซ้ำได้ โดยหัวตารางที่มีชื่อซ้ำกันจะถูกไฮไลท์ด้วยสีแดง
              </span>
              <button class="btn btn-primary btn-sm" id="import-rename-duplicates" type="button">เปลี่ยนชื่อซ้ำอัตโนมัติ</button>
              <button class="btn btn-danger btn-sm" id="import-remove-duplicate-columns" type="button">ลบคอลัมน์ซ้ำ</button>
              <span>ถ้าไม่แน่ใจ ให้เปลี่ยนชื่อก่อน เพราะจะไม่ทำข้อมูลหาย</span>
            </div>
            <div class="import-header-report-list">
              ${report.exactDuplicates.slice(0, 8).map(item => `
                <div class="import-header-report-item">
                  <span>โดยหัวตารางที่ซ้ำกัน คือ ${esc(item.name)}</span>
                  <small>ซ้ำที่คอลัมน์ ${esc(item.columns.join(', '))} · ชื่อแนะนำ: ${esc(item.suggestion.join(' | '))}</small>
                </div>
              `).join('')}
              ${report.exactDuplicates.length > 8 ? `<div class="import-header-report-more">และอีก ${(report.exactDuplicates.length - 8).toLocaleString()} รายการ</div>` : ''}
            </div>
          ` : ''}
          ${!hasExact && hasSimilar ? `
            <div class="import-header-report-list">
              ${report.normalizedDuplicates.slice(0, 6).map(item => `
                <div class="import-header-report-item">
                  <span>${esc(item.variants.map(variant => variant.name).join(' | '))}</span>
                  <small>ชื่อคล้ายกันหลังปรับช่องว่าง/วงเล็บ · คอลัมน์ ${esc(item.variants.map(variant => variant.index).join(', '))}</small>
                </div>
              `).join('')}
            </div>
          ` : ''}
        </div>`;
    };

      const render = () => {
      const job = active.job;
      const rawMeta = active.rawMeta;
      const targetMeta = active.targetMeta;
      const rawTabs = rawMeta.sheetTabs || (rawMeta.tabs || []).map(title => ({ title, sheetId: null }));
      const targetTabs = targetMeta.tabs || [];
      const preview = State.dataImport.preview;
      const targetExists = preview
        ? preview.targetExists
        : targetTabs.includes(State.dataImport.targetTab);
      const rawFiles = State.dataImport.rawFiles || [];
      const selectedFile = findRawFileForJob(job);
      area.innerHTML = `
        <div class="data-import-grid">
          <div class="card data-import-card">
            <div class="card-header">
              <span class="card-title data-import-card-title-step">
                <span class="data-import-step-no">1</span>
                อัพเดตข้อมูลผ่าน Raw Files Folder
                <a href="${esc(job.rawFolderUrl)}" target="_blank" rel="noopener noreferrer" class="data-import-title-link">เปิด Raw Files Folder</a>
              </span>
              <span class="badge badge-data-origin">Google Sheets</span>
            </div>
          </div>

          <div class="card data-import-card">
            <div class="card-header">
              <span class="card-title data-import-card-title-step">
                <span class="data-import-step-no">2</span>
                เลือกข้อมูลที่ต้องการจะดึง
                <span class="data-import-title-subtext">ไฟล์ใน Raw Files Folder</span>
              </span>
              ${State.dataImport.rawFilesError
                ? `<span class="badge badge-warning">อ่าน Drive ไม่ได้</span>`
                : `<span class="row-count-badge">${rawFiles.length.toLocaleString()} files</span>`}
            </div>
            <div class="card-body">
              <div class="data-import-raw-files data-import-raw-files-flat">
                ${State.dataImport.rawFilesError ? `
                  <div class="state-box import-drive-warning">${esc(State.dataImport.rawFilesError)}</div>
                ` : `
                  <div class="import-file-select-row">
                    <select class="fund-input" id="import-job-select" aria-label="เลือกไฟล์ต้นทาง">
                      ${jobList.map(item => `
                        <option value="${esc(item.key)}" ${item.key === job.key ? 'selected' : ''}>
                          ${esc(item.sourceFileName)} → ${esc(item.targetLabel)}
                        </option>
                      `).join('')}
                    </select>
                    <button class="btn btn-primary btn-sm" id="import-preview" type="button">ดึงข้อมูล</button>
                    ${selectedFile?.webViewLink ? `
                      <a href="${esc(selectedFile.webViewLink)}" target="_blank" rel="noopener noreferrer" class="btn btn-ghost btn-sm">เปิดไฟล์</a>
                    ` : ''}
                  </div>
                  <div class="import-file-selected">
                    <span>${esc(selectedFile?.name || job.sourceFileName)}</span>
                    <small>${esc(selectedFile?.modifiedTime ? new Date(selectedFile.modifiedTime).toLocaleString('th-TH') : 'ยังไม่พบไฟล์นี้ในโฟลเดอร์')}</small>
                  </div>
                `}
              </div>
            </div>
          </div>

          <div class="card data-import-card">
            <div class="card-header">
              <span class="card-title data-import-card-title-step"><span class="data-import-step-no">3</span>ยืนยันไฟล์ต้นทาง และ Tab ต้นทางที่ต้องการดึง</span>
            </div>
            <div class="card-body">
              <div class="data-import-section-grid">
                <div class="fund-field">
                  <span>ไฟล์ต้นทาง</span>
                  <input class="fund-input" value="${esc(rawMeta.title || job.rawSheetId)}" readonly>
                </div>
                <div class="fund-field">
                  <span>เลือก Tab ต้นทางที่ต้องการดึง</span>
                  <select class="fund-input" id="import-raw-tab">
                    ${rawTabs.map(tab => `<option value="${esc(tab.title)}" ${tab.title === State.dataImport.rawTab ? 'selected' : ''}>${esc(tab.title)}</option>`).join('')}
                  </select>
                </div>
              </div>
            </div>
          </div>

          ${preview ? `
          <div class="card data-import-card data-import-preview-card">
            <div class="card-header">
              <span class="card-title data-import-card-title-step"><span class="data-import-step-no">4</span>กำหนด Type ให้ถูกต้องก่อนบันทึกข้อมูล</span>
              <div class="page-tools-meta">
                <span class="row-count-badge">${preview.rowCount.toLocaleString()} rows</span>
                <span class="row-count-badge">${preview.colCount.toLocaleString()} columns</span>
              </div>
            </div>
            <div class="card-body">
              <div class="import-status-grid">
                <div><span>Raw Tab</span><strong>${esc(preview.rawTab)}</strong></div>
                <div><span>Target Tab</span><strong>${esc(preview.targetTab)}</strong></div>
                <div><span>Header</span><strong>${esc(preview.headerText || '-')}</strong></div>
                ${preview.cleaner ? `
                  <div><span>Cleaning</span><strong>${esc(preview.cleaner)}</strong></div>
                  <div><span>Raw Rows</span><strong>${Number(preview.sourceRowCount || 0).toLocaleString()}</strong></div>
                  <div><span>Output Rows</span><strong>${Number(preview.rowCount || 0).toLocaleString()}</strong></div>
                ` : ''}
              </div>
              ${renderImportDebug(preview.importDebug)}
              ${renderHeaderReport(preview.headerReport)}
              ${sampleTable(preview.rows)}
            </div>
          </div>` : ''}

          <div class="card data-import-card">
            <div class="card-header">
              <span class="card-title data-import-card-title-step"><span class="data-import-step-no">5</span>ฐานข้อมูลปลายทางที่ต้องการจะบันทึก</span>
              <div class="page-tools-meta">
                <span class="badge ${targetExists ? 'badge-warning' : 'badge-success'}">
                  ${targetExists ? 'พบ Tab แล้ว' : 'ยังไม่มี Tab นี้'}
                </span>
              </div>
            </div>
            <div class="card-body">
              <div class="data-import-form">
                <div class="fund-field">
                  <span>ฐานข้อมูลปลายทาง</span>
                  <input class="fund-input" value="${esc(targetMeta.title || job.targetLabel || job.label)}" readonly>
                </div>
                <div class="fund-field">
                  <span>ระบุข้อมูล โดยรูปแบบ : XXX-QX เช่น 2026-Q1 = ปี ค.ศ. 2026 ไตรมาส 1</span>
                  <input class="fund-input" id="import-target-tab" value="${esc(State.dataImport.targetTab)}" placeholder="2026-Q1">
                </div>
              </div>
            </div>
          </div>

          <div class="card data-import-card">
            <div class="card-header">
              <span class="card-title data-import-card-title-step"><span class="data-import-step-no">6</span>บันทึกข้อมูลปลายทางไปที่ Google Sheets</span>
              <div class="page-tools-meta">
                <span class="badge ${targetExists ? 'badge-warning' : 'badge-success'}">
                  ${targetExists ? 'พบ Tab แล้ว' : 'ยังไม่มี Tab นี้'}
                </span>
              </div>
            </div>
            <div class="card-body">
              <div class="import-status-grid">
                <div><span>Sheet</span><strong>${esc(job.targetLabel || job.label)}</strong></div>
                <div><span>Tab</span><strong>${esc(State.dataImport.targetTab || '-')}</strong></div>
                <div><span>Mode</span><strong>${targetExists ? 'เขียนทับเมื่อยืนยัน' : 'สร้าง Tab ใหม่'}</strong></div>
                <div class="import-status-action-cell">
                  <span>Action</span>
                  <div class="data-import-actions data-import-status-actions">
                    ${preview ? `<button class="btn btn-primary import-commit-btn" id="import-commit" type="button" ${State.dataImport.isImporting ? 'disabled' : ''}>บันทึกข้อมูล</button>` : ''}
                    ${preview?.targetExists ? `
                      <label class="fund-manager-check import-confirm">
                        <input type="checkbox" id="import-overwrite-confirm">
                        ติ๊กเพื่อยืนยันการเขียนทับข้อมูลใน ${esc(preview.targetTab)}
                      </label>` : ''}
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div class="card data-import-card">
            <div class="card-header">
              <span class="card-title data-import-card-title-step"><span class="data-import-step-no">7</span>บันทึกข้อมูลเป็นไฟล์ JSON</span>
              <div class="page-tools-meta">
                <span class="badge ${State.dataImport.jsonReadiness?.allReady ? 'badge-success' : 'badge-data-origin'}">
                  ${State.dataImport.jsonReadiness?.allReady ? 'พร้อมสร้าง JSON' : 'ตรวจไฟล์ทั้ง 3 ชุดก่อน'}
                </span>
              </div>
            </div>
            <div class="card-body">
              ${renderJsonReadiness(State.dataImport.jsonReadiness)}
            </div>
          </div>
        </div>`;
      restoreImportPreviewScroll();

      $('#import-job-select', area)?.addEventListener('change', async e => {
        const nextJob = DATA_IMPORT_JOBS[e.target.value] || DATA_IMPORT_JOBS.percentrank;
        State.dataImport.selectedJobKey = nextJob.key;
        State.dataImport.rawTab = '';
        State.dataImport.targetTab = State.currentQuarter
          || CONFIG.PAGES?.[nextJob.targetPageKey]?.tabName
          || nextJob.defaultTab;
        State.dataImport.preview = null;
        State.dataImport.columnTypes = {};
        setLoading(area, `กำลังโหลดข้อมูลจาก ${nextJob.sourceFileName}...`);
        try {
          active = await loadJobMeta(nextJob);
          applyJobDefaults(active.job, active.rawMeta);
          render();
        } catch (err) {
          setError(area, err.message || 'โหลดไฟล์ต้นทางไม่สำเร็จ', 'data-import');
        }
      });
      $('#import-raw-tab', area)?.addEventListener('change', e => {
        State.dataImport.rawTab = e.target.value;
        State.dataImport.preview = null;
        render();
      });
      $('#import-target-tab', area)?.addEventListener('input', e => {
        State.dataImport.targetTab = e.target.value.trim();
        State.dataImport.jsonReadiness = null;
        if (State.dataImport.preview) {
          State.dataImport.preview = {
            ...State.dataImport.preview,
            targetTab: State.dataImport.targetTab,
            targetExists: (targetMeta.tabs || []).includes(State.dataImport.targetTab),
          };
        }
      });
      $('#import-target-tab', area)?.addEventListener('change', e => {
        State.dataImport.targetTab = e.target.value.trim() || job.defaultTab;
        State.dataImport.jsonReadiness = null;
        if (State.dataImport.preview) {
          State.dataImport.preview = {
            ...State.dataImport.preview,
            targetTab: State.dataImport.targetTab,
            targetExists: (targetMeta.tabs || []).includes(State.dataImport.targetTab),
          };
        }
        render();
      });
      $$('.import-type-select', area).forEach(select => {
        select.addEventListener('change', e => {
          const key = e.target.dataset.columnKey;
          const value = e.target.value;
          if (!key) return;
          if (value === 'auto') {
            delete State.dataImport.columnTypes[key];
          } else {
            State.dataImport.columnTypes[key] = value;
          }
          if (State.dataImport.preview?.rows) {
            State.dataImport.preview = {
              ...State.dataImport.preview,
              columns: analyzeImportColumns(State.dataImport.preview.rows),
            };
          }
          render();
        });
      });
      $('#import-rename-duplicates', area)?.addEventListener('click', () => {
        const previewNow = State.dataImport.preview;
        if (!previewNow?.rows?.length) return;
        const rows = renameDuplicateImportHeaders(previewNow.rows);
        State.dataImport.preview = refreshImportPreviewMeta(previewNow, rows);
        toast('เปลี่ยนชื่อหัวตารางซ้ำอัตโนมัติแล้ว', 'success');
        render();
      });
      $('#import-remove-duplicate-columns', area)?.addEventListener('click', () => {
        const previewNow = State.dataImport.preview;
        const duplicates = previewNow?.headerReport?.exactDuplicates || [];
        const removeIndexes = duplicates.flatMap(item => (item.columns || []).slice(1).map(col => Number(col) - 1));
        if (!previewNow?.rows?.length || !removeIndexes.length) return;
        rememberImportPreviewScroll();
        const rows = removeImportColumns(previewNow.rows, removeIndexes);
        State.dataImport.preview = refreshImportPreviewMeta(previewNow, rows);
        toast(`ลบคอลัมน์ซ้ำ ${removeIndexes.length.toLocaleString()} คอลัมน์ออกจาก Preview แล้ว`, 'success', 4200);
        render();
      });
      $$('[data-delete-column]', area).forEach(button => {
        button.addEventListener('click', e => {
          const previewNow = State.dataImport.preview;
          const colIdx = Number(e.currentTarget.dataset.deleteColumn);
          if (!previewNow?.rows?.length || !Number.isInteger(colIdx)) return;
          const headerName = cleanImportText(previewNow.rows[0]?.[colIdx]) || `Column ${colIdx + 1}`;
          rememberImportPreviewScroll();
          const rows = removeImportColumns(previewNow.rows, [colIdx]);
          State.dataImport.preview = refreshImportPreviewMeta(previewNow, rows);
          toast(`ลบคอลัมน์ ${headerName} ออกจาก Preview แล้ว`, 'success', 4200);
          render();
        });
      });
      $('#import-preview', area)?.addEventListener('click', async () => {
        const rawTab = $('#import-raw-tab', area)?.value || State.dataImport.rawTab;
        const targetTab = $('#import-target-tab', area)?.value.trim() || job.defaultTab;
        State.dataImport.rawTab = rawTab;
        State.dataImport.targetTab = targetTab;
        State.dataImport.jsonReadiness = null;
        setLoading(area, 'กำลังอ่านข้อมูลจาก Raw...');
        try {
          const [sourceRows, latestTargetMeta] = await Promise.all([
            job.sourceType === 'xlsx'
              ? getRowsFromXlsx(job, rawTab)
              : SheetsAPI.fetchSheetData(job.rawSheetId, rawTab),
            SheetsAPI.getSheetTabs(job.targetSheetId),
          ]);
          const rows = cleanRowsForJob(job, sourceRows);
          State.dataImport.preview = refreshImportPreviewMeta({
            rawTab,
            targetTab,
            sourceRowCount: sourceRows.length,
            cleaner: job.cleaner || '',
            importDebug: {
              cleaner: job.cleaner || 'none',
              detected: job.cleaner === 'masterFundRaw' ? detectMorningstarRows(sourceRows) : null,
              rawRow1: (sourceRows[0] || []).slice(0, 40),
              rawRow6: (sourceRows[5] || []).slice(0, 40),
              rawRow7: (sourceRows[6] || []).slice(0, 40),
              rawRow8: (sourceRows[7] || []).slice(0, 40),
              cleanedHeader: (rows[0] || []).slice(0, 60),
            },
            targetExists: (latestTargetMeta.tabs || []).includes(targetTab),
          }, rows);
          render();
        } catch (err) {
          toast(err.message || 'ตรวจสอบข้อมูลไม่สำเร็จ', 'error', 5000);
          render();
        }
      });
      $('#import-commit', area)?.addEventListener('click', async () => {
        const previewNow = State.dataImport.preview;
        if (!previewNow?.rows?.length) {
          toast('ไม่มีข้อมูลสำหรับนำเข้า', 'warning');
          return;
        }
        if (previewNow.targetExists && !$('#import-overwrite-confirm', area)?.checked) {
          toast('กรุณายืนยันก่อนเขียนทับ Tab เดิม', 'warning');
          return;
        }
        State.dataImport.isImporting = true;
        setLoading(area, 'กำลังบันทึกข้อมูลไปยัง Google Sheets...');
        try {
          if (!previewNow.targetExists) {
            await SheetsAPI.addSheetTab(job.targetSheetId, previewNow.targetTab);
          } else {
            await SheetsAPI.clearSheetValues(job.targetSheetId, previewNow.targetTab);
          }
          await SheetsAPI.updateSheetValues(job.targetSheetId, previewNow.targetTab, previewNow.rows);
          clearCache();
          State.dataImport.isImporting = false;
          toast(`บันทึกข้อมูล ${previewNow.rowCount.toLocaleString()} rows ไปที่ ${previewNow.targetTab} แล้ว`, 'success', 5000);
          State.dataImport.jsonReadiness = null;
          State.dataImport.preview = {
            ...previewNow,
            targetExists: true,
          };
          render();
        } catch (err) {
          State.dataImport.isImporting = false;
          toast(err.message || 'บันทึกข้อมูลไม่สำเร็จ', 'error', 6000);
          render();
        }
      });
      $('#json-readiness-check', area)?.addEventListener('click', async () => {
        const targetTab = ($('#import-target-tab', area)?.value || State.dataImport.targetTab || job.defaultTab).trim().toUpperCase();
        if (!/^\d{4}-Q[1-4]$/.test(targetTab)) {
          toast('กรุณาระบุ Quarter รูปแบบ 2026-Q3 ก่อนตรวจ JSON', 'warning');
          return;
        }
        State.dataImport.targetTab = targetTab;
        State.dataImport.isCheckingJson = true;
        render();
        try {
          State.dataImport.jsonReadiness = await checkJsonReadiness(targetTab);
          State.dataImport.isCheckingJson = false;
          toast(State.dataImport.jsonReadiness.allReady ? 'ไฟล์ทั้ง 3 ชุดพร้อมสร้าง JSON แล้ว' : 'ยังมีบางชุดไม่พร้อมสร้าง JSON', State.dataImport.jsonReadiness.allReady ? 'success' : 'warning', 5000);
          render();
        } catch (err) {
          State.dataImport.isCheckingJson = false;
          toast(err.message || 'ตรวจสอบความพร้อม JSON ไม่สำเร็จ', 'error', 6000);
          render();
        }
      });
      $('#json-export-open', area)?.addEventListener('click', async () => {
        const readiness = State.dataImport.jsonReadiness;
        const targetTab = (readiness?.quarter || State.dataImport.targetTab || job.defaultTab).trim().toUpperCase();
        if (!readiness?.allReady) {
          toast('กรุณาตรวจสอบให้พร้อมครบทั้ง 3 ชุดก่อนสร้าง JSON', 'warning');
          return;
        }
        State.dataImport.isExportingJson = true;
        render();
        try {
          const result = await exportBaseJsonFromSheets(targetTab);
          const summary = result.files
            .map(file => `${file.fileName}: ${file.rows.toLocaleString()} rows`)
            .join(' · ');
          toast(`สร้าง JSON สำเร็จ: ${summary}`, 'success', 7000);
          if (result.folderLink) window.open(result.folderLink, '_blank', 'noopener,noreferrer');
        } catch (err) {
          toast(err.message || 'สร้าง JSON ไม่สำเร็จ', 'error', 8000);
        } finally {
          State.dataImport.isExportingJson = false;
          render();
        }
      });
    };

    render();
    App._currentExport = null;
  },

  secDataImport(area) {
    const secDatasets = [
      { id: '01_amcs', title: 'รายชื่อ บลจ.', file: '01_amcs.csv', endpoint: '/v2/fund/general-info/amcs', group: 'Master', cadence: 'เดือนละครั้ง', keys: ['unique_id'], columns: ['unique_id', 'comp_name_th', 'comp_name_en'] },
      { id: '02_profiles', title: 'ข้อมูลกองทุนรวม', file: '02_profiles.csv', endpoint: '/v2/fund/general-info/profiles', group: 'Master', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'proj_abbr_name', 'fund_status', 'last_upd_date'] },
      { id: '03_specifications', title: 'ประเภทกองทุนตามลักษณะพิเศษ', file: '03_specifications.csv', endpoint: '/v2/fund/general-info/specifications', group: 'Master', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'fund_class_name', 'spec_desc', 'last_upd_date'] },
      { id: '04_mutual_fund_fees', title: 'ค่าธรรมเนียมตามโครงการ', file: '04_mutual_fund_fees.csv', endpoint: '/v2/fund/general-info/mutual-fund-fees', group: 'Master', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'fund_class_name', 'fee_type_desc', 'rate', 'last_upd_date'] },
      { id: '05_involve_parties', title: 'ผู้เกี่ยวข้องกับกองทุน', file: '05_involve_parties.csv', endpoint: '/v2/fund/general-info/involve-parties', group: 'Master', cadence: 'เดือนละครั้ง', keys: ['proj_id'], columns: ['proj_id', 'entity_type', 'entity_name_th', 'last_upd_date'] },
      { id: '06_factsheet_urls', title: 'URL Fund Fact Sheet', file: '06_factsheet_urls.csv', endpoint: '/v2/fund/factsheet/urls', group: 'Master', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'fund_class_name', 'pdf_factsheet', 'as_of_date', 'last_upd_date'] },
      { id: '07_ipos', title: 'การเสนอขายกองทุนรวม', file: '07_ipos.csv', endpoint: '/v2/fund/factsheet/ipos', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id'], columns: ['proj_id', 'start_date', 'end_date', 'first_sell_start_date', 'last_upd_date'] },
      { id: '08_benchmarks', title: 'ดัชนีชี้วัด', file: '08_benchmarks.csv', endpoint: '/v2/fund/factsheet/benchmarks', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id'], columns: ['proj_id', 'benchmark', 'start_date', 'end_date', 'last_upd_date'] },
      { id: '09_subscription_redemption_minimums', title: 'ขั้นต่ำซื้อขายคงเหลือ', file: '09_subscription_redemption_minimums.csv', endpoint: '/v2/fund/factsheet/subscription-redemption-minimums', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'fund_class_name', 'minimum_sub', 'minimum_redempt', 'last_upd_date'] },
      { id: '10_subscription_redemption_periods', title: 'ระยะเวลาซื้อขายคืน', file: '10_subscription_redemption_periods.csv', endpoint: '/v2/fund/factsheet/subscription-redemption-periods', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'fund_class_name', 'type', 'period', 'last_upd_date'] },
      { id: '11_risk_spectrum', title: 'ระดับความเสี่ยง', file: '11_risk_spectrum.csv', endpoint: '/v2/fund/factsheet/risk-spectrum', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id'], columns: ['proj_id', 'risk_spectrum', 'risk_spectrum_desc', 'last_upd_date'] },
      { id: '12_statistics', title: 'ข้อมูลเชิงสถิติ', file: '12_statistics.csv', endpoint: '/v2/fund/factsheet/statistics', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'fund_class_name', 'maximum_drawdown', 'sharpe_ratio', 'last_upd_date'] },
      { id: '13_dividend_policy', title: 'นโยบายปันผล', file: '13_dividend_policy.csv', endpoint: '/v2/fund/factsheet/dividend-policy', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'fund_class_name', 'dividend_policy', 'last_upd_date'] },
      { id: '14_fees', title: 'ค่าธรรมเนียม Factsheet', file: '14_fees.csv', endpoint: '/v2/fund/factsheet/fees', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'fund_class_name', 'fee_type_desc', 'actual_value', 'last_upd_date'] },
      { id: '15_performance', title: 'ผลการดำเนินงานย้อนหลัง', file: '15_performance.csv', endpoint: '/v2/fund/factsheet/performance', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id', 'fund_class_name'], columns: ['proj_id', 'fund_class_name', 'reference_period', 'performance_value', 'last_upd_date'] },
      { id: '16_asset_allocation', title: 'สัดส่วนประเภททรัพย์สิน', file: '16_asset_allocation.csv', endpoint: '/v2/fund/factsheet/asset-allocation', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id'], columns: ['proj_id', 'asset_name', 'asset_ratio', 'last_upd_date'] },
      { id: '17_top5_holdings', title: 'ทรัพย์สินที่ลงทุน 5 อันดับแรก', file: '17_top5_holdings.csv', endpoint: '/v2/fund/factsheet/top5-holdings', group: 'Factsheet', cadence: 'เดือนละครั้ง', keys: ['proj_id'], columns: ['proj_id', 'asset_seq', 'asset_name', 'asset_ratio', 'last_upd_date'] },
      { id: '18_outstanding_portfolio', title: 'Portfolio รายไตรมาส', file: '18_outstanding_portfolio.csv', endpoint: '/v2/fund/outstanding/portfolio', group: 'Outstanding', cadence: 'รายไตรมาส', keys: ['proj_id', 'period'], columns: ['proj_id', 'period', 'isin_code', 'issuer', 'percent_nav', 'last_upd_date'] },
      { id: '19_portfolio_asset_type', title: 'Asset Type รายเดือน', file: '19_portfolio_asset_type.csv', endpoint: '/v2/fund/outstanding/portfolio-asset-type', group: 'Outstanding', cadence: 'รายเดือน', keys: ['proj_id', 'period'], columns: ['proj_id', 'period', 'assetliab_desc', 'percent_nav'] },
      { id: '20_nav_daily', title: 'NAV รายวัน', file: '20_nav_daily.csv', endpoint: '/v2/fund/daily-info/nav', group: 'Daily', cadence: 'ทุกวัน', keys: ['proj_id', 'fund_class_name', 'nav_date'], columns: ['proj_id', 'fund_class_name', 'nav_date', 'last_val', 'last_upd_date'] },
      { id: '21_dividend_history', title: 'ประวัติปันผล', file: '21_dividend_history.csv', endpoint: '/v2/fund/daily-info/dividend-history', group: 'Weekly', cadence: 'สัปดาห์ละครั้ง', keys: ['proj_id'], columns: ['proj_id', 'class_abbr_name', 'dividend_date', 'dividend_value', 'last_upd_date'] },
    ];

    const secImportPrefsKey = 'sec-data-import-prefs-v1';
    const defaultSelectedDatasetIds = secDatasets
      .filter(item => {
        const datasetNo = Number(String(item.id).slice(0, 2));
        return datasetNo >= 3 && datasetNo <= 19;
      })
      .map(item => item.id);
    const legacyDefaultSelectedDatasetIds = secDatasets
      .filter(item => !['01_amcs', '02_profiles'].includes(item.id))
      .map(item => item.id);
    const sameDatasetSet = (left, right) => (
      left.length === right.length && left.every(id => right.includes(id))
    );
    const configuredGithubToken = String(CONFIG?.GITHUB_TOKEN || window.APP_SECRETS?.githubToken || '').trim();
    const readSecImportPrefs = () => {
      try {
        return JSON.parse(localStorage.getItem(secImportPrefsKey) || '{}') || {};
      } catch {
        return {};
      }
    };
    const writeSecImportPrefs = prefs => {
      try {
        localStorage.setItem(secImportPrefsKey, JSON.stringify(prefs));
      } catch {
        /* ignore preference persistence failures */
      }
    };
    const secImportPrefs = readSecImportPrefs();
    const savedSelectedDatasetIds = Array.isArray(secImportPrefs.selectedDatasetIds)
      ? secImportPrefs.selectedDatasetIds.filter(id => secDatasets.some(item => item.id === id))
      : [];
    const useDefaultDatasetIds = !savedSelectedDatasetIds.length || sameDatasetSet(savedSelectedDatasetIds, legacyDefaultSelectedDatasetIds);
    const selectedDatasetIds = new Set(useDefaultDatasetIds ? defaultSelectedDatasetIds : savedSelectedDatasetIds);
    const secWorkflowUrl = 'https://github.com/FP2-AVP/Fund-Selection-Tool-V2/actions/workflows/sec-master-view-to-google-sheet.yml';
    const secSpreadsheetId = '1SsL8fXFmKsAfnakrIBTtglZqhErrYz8SfiiZb4wKGDs';
	    const secDataPreparationDefaultTarget = SEC_DATA_PREPARATION_TARGETS[0];
	    const secDataPreparationSpreadsheetId = secDataPreparationDefaultTarget.spreadsheetId;
	    const secDataPreparationSheetUrl = secDataPreparationDefaultTarget.url;
	    const secOutputSpreadsheetId = '16agx9pl9adtMh-U7MCbgnIncBxpciCvFgsdurH6Ob8w';
	    const secOutputSheetUrl = 'https://docs.google.com/spreadsheets/d/16agx9pl9adtMh-U7MCbgnIncBxpciCvFgsdurH6Ob8w/edit?gid=0#gid=0';
	    const secJsonDriveFolderId = JSON_DRIVE_ROOT_FOLDER_ID;
	    const secJsonDriveFolderUrl = JSON_DRIVE_ROOT_FOLDER_URL;
	    const secJsonFileName = 'Data For SEC API.json';
	    const secWorkflowDispatchUrl = 'https://api.github.com/repos/FP2-AVP/Fund-Selection-Tool-V2/actions/workflows/sec-master-view-to-google-sheet.yml/dispatches';
    const secWorkflowRunsUrl = 'https://api.github.com/repos/FP2-AVP/Fund-Selection-Tool-V2/actions/workflows/sec-master-view-to-google-sheet.yml/runs';
    const datasetKey = item => item.id.replace(/^\d+_/, '');
    const dataPreparationRequiredTabs = new Set([
      '02_profiles',
      '06_factsheet_urls',
      '07_ipos',
      '08_benchmarks',
      '09_subscription_redemption_minimums',
      '10_subscription_redemption_periods',
      '11_risk_spectrum',
      '12_statistics',
      '13_dividend_policy',
      '14_fees',
      '15_performance',
      '16_asset_allocation',
      '17_top5_holdings',
    ]);
    const dataPreparationRequirementBadge = item => dataPreparationRequiredTabs.has(item.id)
      ? '<span class="sec-dataset-status is-selected">จำเป็นต้องมี</span>'
      : '<span class="sec-dataset-status">เสริม</span>';
    const chip = (value, tone = '') => `<span class="sec-data-chip ${tone}">${esc(value)}</span>`;
    const secFactsheetLatestDatasetIds = new Set([
      '07_ipos',
      '08_benchmarks',
      '09_subscription_redemption_minimums',
      '10_subscription_redemption_periods',
      '11_risk_spectrum',
      '12_statistics',
      '13_dividend_policy',
      '14_fees',
      '15_performance',
      '16_asset_allocation',
      '17_top5_holdings',
    ]);
    const parseSecDateValue = value => {
      const text = String(value ?? '').trim();
      if (!text) return Number.NEGATIVE_INFINITY;
      const isoMatch = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
      if (isoMatch) {
        const year = Number(isoMatch[1]);
        const normalizedYear = year > 2400 ? year - 543 : year;
        return new Date(normalizedYear, Number(isoMatch[2]) - 1, Number(isoMatch[3])).getTime();
      }
      const thaiMatch = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
      if (thaiMatch) {
        const year = Number(thaiMatch[3]);
        const normalizedYear = year > 2400 ? year - 543 : year;
        return new Date(normalizedYear, Number(thaiMatch[2]) - 1, Number(thaiMatch[1])).getTime();
      }
      const parsed = Date.parse(text);
      return Number.isNaN(parsed) ? Number.NEGATIVE_INFINITY : parsed;
    };
    const secColumnIndex = (headers, candidates) => findColumnIndex(headers, candidates);
    const secSortableValue = (value, mode = 'date') => {
      const text = String(value ?? '').trim();
      if (!text) return Number.NEGATIVE_INFINITY;
      if (mode === 'period') {
        const numeric = Number(text.replace(/[^\d]/g, ''));
        return Number.isFinite(numeric) ? numeric : Number.NEGATIVE_INFINITY;
      }
      const parsed = parseSecDateValue(text);
      return parsed === Number.NEGATIVE_INFINITY ? text : parsed;
    };
    const secColumnRange = (values, headers, candidates, mode = 'date') => {
      const idx = secColumnIndex(headers, candidates);
      if (idx < 0) return null;
      let minValue = '';
      let maxValue = '';
      let minSort = Number.POSITIVE_INFINITY;
      let maxSort = Number.NEGATIVE_INFINITY;
      values.slice(1).forEach(row => {
        const value = String(row?.[idx] ?? '').trim();
        if (!value) return;
        const sortable = secSortableValue(value, mode);
        if (sortable < minSort || !minValue) {
          minSort = sortable;
          minValue = value;
        }
        if (sortable > maxSort || !maxValue) {
          maxSort = sortable;
          maxValue = value;
        }
      });
      if (!minValue && !maxValue) return null;
      return {
        min: minValue || maxValue,
        max: maxValue || minValue,
        same: (minValue || maxValue) === (maxValue || minValue),
      };
    };
    const secRangeText = (range) => {
      if (!range) return '';
      return range.same ? range.max : `${range.min} ถึง ${range.max}`;
    };
    const secDataCoverage = (values, item) => {
      const headers = values?.[0] || [];
      if (!headers.length) return { label: 'ข้อมูลจริง', value: '-', note: '' };
      if (item.id === '06_factsheet_urls') {
        const range = secColumnRange(values, headers, ['as_of_date']);
        return { label: 'as_of_date', value: secRangeText(range) || '-', note: '' };
      }
      if (['18_outstanding_portfolio', '19_portfolio_asset_type'].includes(item.id)) {
        const range = secColumnRange(values, headers, ['period'], 'period');
        return { label: 'period', value: secRangeText(range) || '-', note: '' };
      }
      if (item.id === '20_nav_daily') {
        const range = secColumnRange(values, headers, ['nav_date']);
        return { label: 'nav_date', value: secRangeText(range) || '-', note: '' };
      }
      if (item.id === '21_dividend_history') {
        const range = secColumnRange(values, headers, ['dividend_date']);
        return { label: 'dividend_date', value: secRangeText(range) || '-', note: '' };
      }
      if (secFactsheetLatestDatasetIds.has(item.id)) {
        const startRange = secColumnRange(values, headers, ['start_date']);
        const endRange = secColumnRange(values, headers, ['end_date']);
        const startText = secRangeText(startRange);
        const endText = secRangeText(endRange);
        const value = startText || endText
          ? `${startText || '-'} ถึง ${endText || '-'}`
          : '-';
        return {
          label: 'start/end ในข้อมูล',
          value,
          note: 'factsheet ใช้ latest=true จึงไม่ filter ตามวันที่ที่กรอก',
        };
      }
      const startRange = secColumnRange(values, headers, ['start_date']);
      const endRange = secColumnRange(values, headers, ['end_date']);
      const fallbackRange = secColumnRange(values, headers, ['as_of_date', 'nav_date', 'dividend_date', 'period']);
      const startText = secRangeText(startRange);
      const endText = secRangeText(endRange);
      if (startText || endText) return { label: 'start/end ในข้อมูล', value: `${startText || '-'} ถึง ${endText || '-'}`, note: '' };
      return { label: 'ข้อมูลจริง', value: secRangeText(fallbackRange) || '-', note: '' };
    };
    const secCoverageHtml = (status) => {
      if (!status?.exists) return '<span class="sec-dataset-status is-missing">-</span>';
      const coverage = status.coverage || { label: 'ข้อมูลจริง', value: '-' };
      return `
        <div class="sec-coverage-stack">
          <div><span>${esc(coverage.label || 'ข้อมูลจริง')}</span><strong>${esc(coverage.value || '-')}</strong></div>
          ${coverage.note ? `<small>${esc(coverage.note)}</small>` : ''}
        </div>`;
    };
	    const detectSecImportValueType = (value) => {
	      const text = String(value ?? '').trim();
	      if (!text) return '';
	      if (/^[-–—]+$/.test(text)) return '';
	      if (/^(n\/a|na|null|none)$/i.test(text)) return '';
	      if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(text)) return 'date';
	      if (/^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}/.test(text)) return 'date';
	      if (/^-?\d+(,\d{3})*(\.\d+)?%$/.test(text) || /^-?\d+(\.\d+)?%$/.test(text)) return 'percent';
	      if (/^-?\d+(,\d{3})*(\.\d+)?$/.test(text) || /^-?\d+(\.\d+)?$/.test(text)) return 'number';
	      return 'text';
	    };
	    const secTypeBadgeClass = (type) => ({
	      number: 'badge-success',
	      percent: 'badge-accent',
	      date: 'badge-primary',
	      text: 'badge-data-origin',
	      mixed: 'badge-warning',
	      blank: 'badge-muted',
	    }[type] || 'badge-data-origin');
	    const secTypeLabel = (type) => ({
	      number: 'Number',
	      percent: 'Percent',
	      date: 'Date',
	      text: 'Text',
	      mixed: 'Mixed',
	      blank: 'Blank',
	    }[type] || String(type || '').replace(/^./, c => c.toUpperCase()));
	    const analyzeSecImportColumns = (rows = []) => {
	      const headers = rows[0] || [];
	      const maxCols = rows.length ? Math.max(...rows.map(row => row.length)) : headers.length;
	      return Array.from({ length: maxCols }, (_, colIdx) => {
	        const counts = { text: 0, number: 0, percent: 0, date: 0 };
	        rows.slice(1, 101).forEach(row => {
	          const type = detectSecImportValueType(row[colIdx]);
	          if (type) counts[type] += 1;
	        });
	        const used = Object.entries(counts).filter(([, count]) => count > 0);
	        const detectedType = used.length === 0
	          ? 'blank'
	          : (used.length === 1 ? used[0][0] : 'mixed');
	        const key = `${colIdx}:${headers[colIdx] || `Column ${colIdx + 1}`}`;
	        return {
	          key,
	          index: colIdx + 1,
	          name: headers[colIdx] || `Column ${colIdx + 1}`,
	          detectedType,
	          type: State.secDataImport.columnTypes[key] || detectedType,
	          counts,
	        };
	      });
	    };
	    const refreshSecDataPreparationPreview = (rows = []) => {
	      const colCount = rows.length ? Math.max(...rows.map(row => row.length)) : 0;
	      return {
	        rows,
	        rowCount: Math.max(0, rows.length - 1),
	        colCount,
	        columns: analyzeSecImportColumns(rows),
	      };
	    };
	    const secImportTypeSelect = (col) => {
	      const typeOptions = [
	        ['auto', 'Auto'],
	        ['text', 'Text'],
	        ['number', 'Number'],
	        ['percent', 'Percent'],
	        ['date', 'Date'],
	        ['blank', 'Blank'],
	      ];
	      return `
	        <div class="import-header-type">
	          <strong class="badge ${secTypeBadgeClass(col.type)}">${esc(secTypeLabel(col.type))}</strong>
	          <select class="import-type-select sec-import-type-select" data-column-key="${esc(col.key)}" aria-label="เลือกชนิดข้อมูลของ ${esc(col.name)}">
	            ${typeOptions.map(([value, label]) => `
	              <option value="${value}" ${(State.secDataImport.columnTypes[col.key] || 'auto') === value ? 'selected' : ''}>
	                ${label}${value === 'auto' ? ` (${secTypeLabel(col.detectedType)})` : ''}
	              </option>`).join('')}
	          </select>
	        </div>`;
	    };
	    const secDataPreparationSampleTable = (preview) => {
	      const rows = preview?.rows || [];
	      if (!rows.length) return '<div class="state-box">ยังไม่ได้โหลด tab data_preparation</div>';
	      const headers = rows[0] || [];
	      const dataRows = rows.slice(1);
	      const visibleRows = [headers, ...dataRows.slice(0, 5)];
	      const maxCols = Math.max(...rows.map(row => row.length));
	      const columns = preview.columns || analyzeSecImportColumns(rows);
	      return `
	        <div class="table-wrapper import-preview-table">
	          <table>
	            <tbody>
	              ${visibleRows.map((row, rowIdx) => `
	                <tr class="${rowIdx === 0 ? 'is-header-row' : ''}">
	                  ${Array.from({ length: maxCols }, (_, i) => rowIdx === 0
	                    ? `<td><div class="import-header-cell"><span>${esc(row[i] ?? `Column ${i + 1}`)}</span>${secImportTypeSelect(columns[i] || { key: `${i}:Column ${i + 1}`, name: `Column ${i + 1}`, type: 'blank', detectedType: 'blank' })}</div></td>`
	                    : `<td>${esc(row[i] ?? '')}</td>`).join('')}
	                </tr>
	              `).join('')}
	            </tbody>
	          </table>
	        </div>
	        <div class="pg-info import-preview-note">แสดงทุกคอลัมน์ และตัวอย่าง 5 แถวแรกจาก ${dataRows.length.toLocaleString()} rows</div>`;
	    };
	    const secJsonStatusBadge = () => {
	      if (State.secDataImport.isExportingJson) return '<span class="badge badge-data-origin">กำลังสร้าง JSON</span>';
	      if (State.secDataImport.jsonStatus?.ok) return '<span class="badge badge-success">สร้าง JSON แล้ว</span>';
	      if (State.secDataImport.jsonStatus?.error) return '<span class="badge badge-warning">สร้าง JSON ไม่สำเร็จ</span>';
	      return '<span class="badge badge-data-origin">ยังไม่ได้สร้าง JSON</span>';
	    };
	    const getSecJsonQuarter = () => {
	      const targetQuarter = String(State.secDataImport.targetTab || '').trim().toUpperCase();
	      if (/^\d{4}-Q[1-4]$/.test(targetQuarter)) return targetQuarter;
	      const cleanQuarter = String(State.currentQuarter || '').trim().toUpperCase();
	      return /^\d{4}-Q[1-4]$/.test(cleanQuarter) ? cleanQuarter : 'unknown-quarter';
	    };
	    const getSecJsonYear = (quarter) => {
	      const match = String(quarter || '').trim().match(/^(\d{4})-Q[1-4]$/i);
	      return match?.[1] || 'unknown-year';
	    };
	    const getSecJsonPathParts = () => {
	      const quarter = getSecJsonQuarter();
	      return {
	        quarter,
	        year: getSecJsonYear(quarter),
	        folderSegments: [getSecJsonYear(quarter), quarter, 'base'],
	      };
	    };
	    const getSecJsonDisplayPath = () => {
	      const { year, quarter } = getSecJsonPathParts();
	      return `${JSON_STORE.rootName}/${year}/${quarter}/base/${secJsonFileName}`;
	    };
	    const presetOptions = [
      ['custom', 'เลือก endpoint raw tabs เอง'],
    ].map(([value, label]) => `<option value="${esc(value)}">${esc(label)}</option>`).join('');
    const renderChips = (values, tone = '') => values.map(value => chip(value, tone)).join('');
    const renderDatasetHeader = item => `
      <th class="sec-dataset-select-cell" title="${esc(item.title)}">
        <label class="sec-dataset-toggle">
          <input
            type="checkbox"
            class="sec-dataset-check"
            value="${esc(item.id)}"
            ${selectedDatasetIds.has(item.id) ? 'checked' : ''}
          >
          <span>${esc(datasetKey(item))}</span>
        </label>
      </th>`;
    const renderDatasetCells = () => secDatasets.map(item => `
      <td data-sec-dataset-cell="${esc(item.id)}">
        <span class="sec-dataset-status ${selectedDatasetIds.has(item.id) ? 'is-selected' : ''}">
          ${selectedDatasetIds.has(item.id) ? 'เลือก' : '-'}
        </span>
      </td>`).join('');
    const prefValue = (key, fallback = '') => secImportPrefs[key] ?? fallback;
    const savedProjectPairs = Array.isArray(secImportPrefs.projectPairs)
      ? secImportPrefs.projectPairs
      : String(secImportPrefs.projIds || '').split(/\n+/).filter(Boolean).map(projId => ({ projId, fundClassName: '' }));
    const normalizedProjectPairs = pairs => {
      const rows = (pairs || []).map(pair => ({
        projId: String(pair.projId || '').trim(),
        fundClassName: String(pair.fundClassName || '').trim(),
      }));
      return rows.length ? rows : [{ projId: '', fundClassName: '' }];
    };
    const renderProjectPairRows = pairs => normalizedProjectPairs(pairs).map((pair, index) => `
      <div class="sec-project-pair-row" data-sec-project-pair-row>
        <input
          class="fund-input sec-project-pair-proj"
          type="text"
          value="${esc(pair.projId)}"
          placeholder="${index === 0 ? 'เว้นว่างเพื่อใช้ Registered list' : 'proj_id'}"
        >
        <input
          class="fund-input sec-project-pair-class"
          type="text"
          value="${esc(pair.fundClassName)}"
          placeholder="fund_class_name เว้นว่างได้"
        >
        <button class="btn btn-secondary sec-project-pair-remove" type="button" title="ลบบรรทัด">ลบ</button>
      </div>
    `).join('');
    const spreadsheetIdFromValue = value => {
      const text = String(value || '').trim();
      const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
      return match ? match[1] : text;
    };
    const masterViewColumns = {
      '01_amcs': { label: 'amc_name', values: ['บริษัทหลักทรัพย์จัดการกองทุนตัวอย่าง', 'บริษัทหลักทรัพย์จัดการกองทุนตัวอย่าง'] },
      '02_profiles': { label: 'fund_class_name', values: ['main', 'R'] },
      '03_specifications': { label: 'specifications', values: ['Thai ESG | RMF', 'SSF'] },
      '04_mutual_fund_fees': { label: 'mutual_fund_fees', values: ['Management Fee 1.07% | Total Fee 1.38%', 'Management Fee 0.95% | Total Fee 1.22%'] },
      '05_involve_parties': { label: 'involve_parties', values: ['Registrar: TSD | Trustee: Bank', 'Registrar: TSD | Trustee: Bank'] },
      '06_factsheet_urls': { label: 'factsheet_url_latest', values: ['https://.../factsheet_M0000.pdf', 'https://.../factsheet_M0001.pdf'] },
      '07_ipos': { label: 'ipo_period', values: ['2009-01-05 - 2009-01-12', '2009-02-02 - 2009-02-09'] },
      '08_benchmarks': { label: 'benchmark_latest', values: ['SET TRI | MSCI ACWI', 'SET TRI'] },
      '09_subscription_redemption_minimums': { label: 'minimums', values: ['Sub 1,000 | Redeem 1,000', 'Sub 500 | Redeem 500'] },
      '10_subscription_redemption_periods': { label: 'subscription_redemption_periods', values: ['ทุกวันทำการ | T+3', 'ทุกวันทำการ | T+2'] },
      '11_risk_spectrum': { label: 'risk_spectrum', values: ['RS6', 'RS5'] },
      '12_statistics': { label: 'statistics', values: ['Max DD -12.4 | Sharpe 0.51', 'Max DD -9.8 | Sharpe 0.62'] },
      '13_dividend_policy': { label: 'dividend_policy', values: ['จ่ายเงินปันผล', 'ไม่จ่ายเงินปันผล'] },
      '14_fees': { label: 'factsheet_fees', values: ['Front-end 1.00 | Total 1.38', 'Front-end 0.50 | Total 1.22'] },
      '15_performance': { label: 'performance_latest', values: ['1Y 8.42 | 3Y 5.10 | 5Y 4.88', '1Y 6.15 | 3Y 4.22 | 5Y 3.90'] },
      '16_asset_allocation': { label: 'asset_allocation', values: ['Equity 82.5 | Cash 7.1', 'Bond 65.2 | Equity 21.3'] },
      '17_top5_holdings': { label: 'top5_holdings', values: ['AAPL 8.2%, MSFT 7.1%, NVDA 6.4%', 'PTT 6.1%, CPALL 5.4%, AOT 4.9%'] },
      '18_outstanding_portfolio': { label: 'outstanding_portfolio', values: ['202603 | 1,246 holdings', '202603 | 842 holdings'] },
      '19_portfolio_asset_type': { label: 'portfolio_asset_type', values: ['202605 | Equity 74.2 | Cash 4.3', '202605 | Bond 58.8 | Cash 10.2'] },
      '20_nav_daily': { label: 'nav_latest', values: ['2026-06-24 | NAV 12.4567', '2026-06-24 | NAV 10.2345'] },
      '21_dividend_history': { label: 'dividend_history', values: ['ล่าสุด 2026-05-15 | 0.18', 'ล่าสุด 2026-04-18 | 0.12'] },
    };
    const masterBaseRows = [
      { proj_id: 'M0000_2552', fund_name: 'กองทุนตัวอย่าง', status: 'Registered', last_upd_date: '2026-06-24' },
      { proj_id: 'M0001_2552', fund_name: 'กองทุนตัวอย่าง Class R', status: 'Registered', last_upd_date: '2026-06-24' },
    ];
    const renderMasterViewHead = ids => `
      <tr>
        <th>proj_id</th>
        <th>fund_name</th>
        <th>status</th>
        <th>last_upd_date</th>
        ${ids.map(id => `<th>${esc(masterViewColumns[id]?.label || datasetKey({ id }))}</th>`).join('')}
      </tr>`;
    const renderMasterViewBody = ids => masterBaseRows.map((row, rowIndex) => `
      <tr>
        <td>${esc(row.proj_id)}</td>
        <td>${esc(row.fund_name)}</td>
        <td>${esc(row.status)}</td>
        <td>${esc(row.last_upd_date)}</td>
        ${ids.map(id => `<td>${esc(masterViewColumns[id]?.values[rowIndex] || '-')}</td>`).join('')}
      </tr>`).join('');

    area.innerHTML = `
      <div class="sec-data-layout">
        <div class="card sec-data-card">
          <div class="card-header">
            <span class="card-title data-import-card-title-step"><span class="data-import-step-no">1</span>กำหนดเงื่อนไขการดึงข้อมูล SEC</span>
            <div class="page-tools-meta">
              <span class="badge badge-data-origin">21 datasets</span>
            </div>
          </div>
          <div class="card-body">
            <div class="sec-data-form-grid">
              <label class="fund-field sec-data-field-main">
                <span>รูปแบบการเลือกข้อมูล</span>
                <select class="fund-input sec-data-highlight" id="sec-dataset-preset">
                  ${presetOptions}
                </select>
              </label>
              <label class="fund-field">
                <span>โหมดการดึง</span>
                <select class="fund-input sec-data-highlight" id="sec-fund-status">
                  <option value="Registered" ${prefValue('fundStatus', 'Registered') === 'Registered' ? 'selected' : ''}>Registered only</option>
                  <option value="" ${prefValue('fundStatus', 'Registered') === '' ? 'selected' : ''}>ดึงทุกสถานะ</option>
                </select>
              </label>
              <div class="fund-field sec-data-field-full sec-project-pair-editor">
                <span>คู่ proj_id และ fund_class_name</span>
                <div class="sec-project-pair-head">
                  <span>proj_id</span>
                  <span>fund_class_name</span>
                  <span></span>
                </div>
                <div class="sec-project-pair-list" id="sec-project-pair-list">
                  ${renderProjectPairRows(savedProjectPairs)}
                </div>
                <button class="btn btn-secondary sec-project-pair-add" id="sec-add-project-pair" type="button">เพิ่ม proj_id และ fund_class_name ที่ต้องการค้นหา</button>
              </div>
              <label class="fund-field sec-data-field-half">
                <span>วันที่เริ่มต้น</span>
                <input class="fund-input" id="sec-start-date" type="date" value="${esc(prefValue('startDate', '2026-01-01'))}">
              </label>
              <label class="fund-field sec-data-field-half">
                <span>วันที่สิ้นสุด</span>
                <input class="fund-input" id="sec-end-date" type="date" value="${esc(prefValue('endDate', '2026-06-30'))}">
              </label>
              <label class="fund-field sec-data-field-half">
                <span>period เริ่มต้น</span>
                <input class="fund-input" id="sec-start-period" type="text" value="${esc(prefValue('startPeriod', '202601'))}" placeholder="YYYYMM">
              </label>
              <label class="fund-field sec-data-field-half">
                <span>period สิ้นสุด</span>
                <input class="fund-input" id="sec-end-period" type="text" value="${esc(prefValue('endPeriod', '202606'))}" placeholder="YYYYMM">
              </label>
              <label class="fund-field sec-data-field-half">
                <span>registered_max_funds (ระบุเป็น 0 หากต้องการจะทุกกอง)</span>
                <input class="fund-input sec-data-highlight" id="sec-registered-max-funds" type="number" min="0" step="1" value="${esc(prefValue('registeredMaxFunds', '300'))}">
              </label>
              <label class="fund-field sec-data-field-half">
                <span>max_pages</span>
                <input class="fund-input sec-data-highlight" id="sec-max-pages" type="number" min="0" step="1" value="${esc(prefValue('maxPages', '0'))}">
              </label>
              <label class="fund-field sec-data-field-main">
                <span>GitHub token</span>
                <input class="fund-input" id="sec-github-token" type="password" value="${esc(configuredGithubToken)}" placeholder="ใส่ token ตอนจะรัน หรือกำหนดไว้ใน config.override.js">
              </label>
              <div class="sec-data-actions">
                <button class="btn btn-secondary" id="sec-open-master-sheet" type="button">เปิด Google Sheet</button>
                <button class="btn sec-data-primary-btn" id="sec-run-workflow" type="button">รัน GitHub Actions</button>
              </div>
            </div>
            <div class="sec-workflow-panel">
              <div class="sec-workflow-copy">
                <strong>ระยะที่ 2:</strong>
                ปุ่มนี้จะส่งคำสั่งรัน workflow บน GitHub Actions จริง โดยค่าเริ่มต้นจะเขียนข้อมูลแยกเป็น tab raw ของแต่ละ endpoint
              </div>
              <div class="sec-workflow-input-grid">
                <div>
                  <span>workflow</span>
                  <strong>SEC Data Preparation to Google Sheet</strong>
                </div>
                <div>
                  <span>output_mode</span>
                  <strong id="sec-workflow-output-mode">raw_tabs</strong>
                </div>
                <div>
                  <span>dataset_preset</span>
                  <strong id="sec-workflow-preset">custom</strong>
                </div>
                <div>
                  <span>custom_datasets</span>
                  <strong id="sec-workflow-datasets">specifications, mutual_fund_fees, factsheet_urls, benchmarks</strong>
                </div>
                <div>
                  <span>registered_max_funds</span>
                  <strong id="sec-workflow-registered-max">300</strong>
                </div>
                <div>
                  <span>max_pages</span>
                  <strong id="sec-workflow-max-pages">0</strong>
                </div>
              </div>
              <div class="sec-workflow-warning">
                เมื่อ workflow เขียนสำเร็จ ระบบจะล้างข้อมูลเดิมใน tab endpoint ที่เลือก แล้วเขียนข้อมูลชุดล่าสุดใหม่ทั้งหมด โดย endpoint อื่นจะใช้ <strong>02_profiles เดิม</strong> เป็นฐาน เว้นแต่เลือกดึง 02_profiles เอง
              </div>
              <div class="sec-workflow-status" id="sec-workflow-status" hidden></div>
            </div>
          </div>
        </div>

        <div class="sec-data-summary-grid">
          <div class="sec-data-summary-card">
            <span>Registered proj_id</span>
            <strong>ใช้ 02_profiles เดิม</strong>
          </div>
          <div class="sec-data-summary-card">
            <span>Join key หลัก</span>
            <strong>proj_id + fund_class_name</strong>
          </div>
          <div class="sec-data-summary-card">
            <span>เลือกไว้ตอนนี้</span>
            <strong><span id="sec-selected-count">${selectedDatasetIds.size}</span> ชุดข้อมูล</strong>
          </div>
          <div class="sec-data-summary-card">
            <span>รูปแบบไฟล์หลังบ้าน</span>
            <strong>Raw endpoint tabs</strong>
          </div>
        </div>

        <div class="card sec-data-card">
          <div class="card-header">
            <span class="card-title data-import-card-title-step">
              <span class="data-import-step-no">2</span>
              เลือก endpoint และตรวจสถานะ raw tabs
              <button class="btn btn-secondary btn-sm" id="sec-check-tabs" type="button">ตรวจสถานะ Tabs</button>
            </span>
            <div class="page-tools-meta"><span class="badge badge-success">อ่านจาก tab จริง</span></div>
          </div>
          <div class="card-body">
            <div class="sec-command-preview">
              <span>ชุดข้อมูลที่จะดึง</span>
              <code id="sec-dataset-command">specifications, mutual_fund_fees, factsheet_urls, benchmarks</code>
            </div>
            <div class="sec-master-note">
              ใช้ checkbox ด้านหน้าเพื่อเลือกว่าจะดึง endpoint ไหน ส่วนสถานะจะแสดงจาก Google Sheet ปลายทาง เช่น <strong>02_profiles</strong>, <strong>06_factsheet_urls</strong>, <strong>20_nav_daily</strong>
            </div>
            <div class="table-scroll sec-table-scroll">
              <table class="sec-data-preview-table sec-tab-status-table">
                <thead>
                  <tr>
                    <th>เลือก</th>
                    <th>ความจำเป็นของข้อมูล</th>
                    <th>หลักฐานช่วงข้อมูล</th>
                    <th>สถานะ</th>
                    <th>Tab</th>
                    <th>ชุดข้อมูล</th>
                    <th>หัวคอลัมน์หลัก</th>
                  </tr>
                </thead>
                <tbody id="sec-tab-status-body">
                  ${secDatasets.map(item => `
                    <tr data-sec-tab-status-row="${esc(item.id)}" data-sec-dataset-row="${esc(item.id)}" class="${selectedDatasetIds.has(item.id) ? 'is-selected' : ''}">
                      <td>
                        <label class="sec-dataset-toggle sec-dataset-toggle-inline">
                          <input
                            type="checkbox"
                            class="sec-dataset-check"
                            value="${esc(item.id)}"
                            ${selectedDatasetIds.has(item.id) ? 'checked' : ''}
                          >
                          <span>${selectedDatasetIds.has(item.id) ? 'เลือก' : 'ข้าม'}</span>
                        </label>
                      </td>
                      <td>${dataPreparationRequirementBadge(item)}</td>
                      <td data-sec-update-date-cell><span class="sec-dataset-status">ยังไม่ได้ตรวจ</span></td>
                      <td data-sec-status-cell><span class="sec-dataset-status">ยังไม่ได้ตรวจ</span></td>
                      <td><code>${esc(item.id)}</code></td>
                      <td>${esc(item.title)}</td>
                      <td>${renderChips(item.columns)}</td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <div class="card sec-data-card">
          <div class="card-header">
            <span class="card-title data-import-card-title-step">
              <span class="data-import-step-no">3</span>
              รวบรวมข้อมูลเป็น data_preparation
            </span>
            <div class="page-tools-meta">
              <span class="badge badge-data-origin">master_view</span>
              <button class="btn sec-data-primary-btn btn-sm" id="sec-build-master-view" type="button">รวบรวมข้อมูล</button>
            </div>
          </div>
          <div class="card-body">
            <div class="sec-master-note">
              อ่านข้อมูลจาก raw tabs ที่ตรวจในขั้นตอนที่ 2 แล้วรวมเป็น tab <strong>data_preparation</strong> สำหรับใช้งานต่อ โดยใช้ workflow โหมด <strong>master_from_raw</strong>
            </div>
          </div>
        </div>

        <div class="card sec-data-card data-import-preview-card">
          <div class="card-header">
            <span class="card-title data-import-card-title-step">
              <span class="data-import-step-no">4</span>
              กำหนด Type ให้ถูกต้องก่อนบันทึกข้อมูล
            </span>
            <div class="page-tools-meta">
              ${State.secDataImport.preview ? `
                <span class="row-count-badge">${State.secDataImport.preview.rowCount.toLocaleString()} rows</span>
                <span class="row-count-badge">${State.secDataImport.preview.colCount.toLocaleString()} columns</span>
              ` : '<span class="badge badge-data-origin">data_preparation</span>'}
              <button class="btn btn-secondary btn-sm" id="sec-load-data-preparation-preview" type="button" ${State.secDataImport.isLoadingPreview ? 'disabled' : ''}>
                ${State.secDataImport.isLoadingPreview ? 'กำลังโหลด...' : 'โหลด data_preparation'}
              </button>
            </div>
          </div>
          <div class="card-body">
            ${State.secDataImport.preview ? `
              <div class="import-status-grid">
                <div><span>Source</span><strong>Raw For SEC</strong></div>
                <div><span>Tab</span><strong>data_preparation</strong></div>
                <div><span>Mode</span><strong>Preview Type</strong></div>
              </div>
            ` : `
              <div class="sec-master-note">
                โหลด tab <strong>data_preparation</strong> หลังจากรวบรวมข้อมูลแล้ว เพื่อตรวจชนิดข้อมูลของแต่ละคอลัมน์ก่อนนำไปใช้งานต่อ
              </div>
            `}
            ${secDataPreparationSampleTable(State.secDataImport.preview)}
          </div>
        </div>

        <div class="card sec-data-card">
          <div class="card-header">
            <span class="card-title data-import-card-title-step"><span class="data-import-step-no">5</span>ฐานข้อมูลปลายทางที่ต้องการจะบันทึก</span>
            <div class="page-tools-meta">
              <span class="badge ${State.secDataImport.targetExists ? 'badge-warning' : 'badge-success'}">
                ${State.secDataImport.targetExists ? 'พบ Tab แล้ว' : 'ยังไม่มี Tab นี้'}
              </span>
            </div>
          </div>
          <div class="card-body">
            <div class="data-import-form">
              <div class="fund-field">
                <span>ฐานข้อมูลปลายทาง</span>
                <input class="fund-input" value="SEC data_preparation" readonly>
              </div>
              <div class="fund-field">
                <span>ระบุชื่อ Tab ปลายทาง</span>
                <input class="fund-input" id="sec-output-target-tab" value="${esc(State.secDataImport.targetTab || 'data_preparation')}" placeholder="data_preparation">
              </div>
            </div>
            <div class="sec-master-note">
              ปลายทาง: <a href="${esc(secOutputSheetUrl)}" target="_blank" rel="noopener noreferrer">Google Sheet data_preparation</a>
            </div>
          </div>
        </div>

        <div class="card sec-data-card">
          <div class="card-header">
            <span class="card-title data-import-card-title-step"><span class="data-import-step-no">6</span>บันทึกข้อมูลปลายทางไปที่ Google Sheets</span>
            <div class="page-tools-meta">
              <span class="badge ${State.secDataImport.targetExists ? 'badge-warning' : 'badge-success'}">
                ${State.secDataImport.targetExists ? 'พบ Tab แล้ว' : 'ยังไม่มี Tab นี้'}
              </span>
            </div>
          </div>
          <div class="card-body">
            <div class="import-status-grid">
              <div><span>Sheet</span><strong>SEC data_preparation</strong></div>
              <div><span>Tab</span><strong>${esc(State.secDataImport.targetTab || 'data_preparation')}</strong></div>
              <div><span>Mode</span><strong>${State.secDataImport.targetExists ? 'เขียนทับเมื่อยืนยัน' : 'สร้าง Tab ใหม่'}</strong></div>
              <div class="import-status-action-cell">
                <span>Action</span>
                <div class="data-import-actions data-import-status-actions">
                  ${State.secDataImport.preview ? `<button class="btn btn-primary import-commit-btn" id="sec-output-save" type="button" ${State.secDataImport.isImporting ? 'disabled' : ''}>บันทึกข้อมูล</button>` : ''}
                  ${State.secDataImport.preview && State.secDataImport.targetExists ? `
                    <label class="fund-manager-check import-confirm">
                      <input type="checkbox" id="sec-output-overwrite-confirm">
                      ติ๊กเพื่อยืนยันการเขียนทับข้อมูลใน ${esc(State.secDataImport.targetTab || 'data_preparation')}
                    </label>` : ''}
                </div>
              </div>
            </div>
          </div>
        </div>

        <div class="card sec-data-card">
          <div class="card-header">
            <span class="card-title data-import-card-title-step"><span class="data-import-step-no">7</span>บันทึกข้อมูลเป็นไฟล์ JSON</span>
            <div class="page-tools-meta">${secJsonStatusBadge()}</div>
          </div>
	          <div class="card-body">
	            <div class="json-readiness-box">
	              <div class="json-readiness-copy">
	                สร้างไฟล์ <strong>${esc(secJsonFileName)}</strong> จาก preview ปัจจุบัน แล้วบันทึกตาม path ของ Quarter ที่เลือก
	              </div>
	              <div class="json-path-grid">
	                <div class="fund-field data-import-path-field">
	                  <span>Base JSON Folder</span>
	                  <input class="fund-input" value="${esc(secJsonDriveFolderUrl)}" readonly>
	                </div>
	                <div class="fund-field data-import-path-field">
	                  <span>JSON Path</span>
	                  <input class="fund-input" value="${esc(getSecJsonDisplayPath())}" readonly>
	                </div>
	              </div>
	              ${State.secDataImport.jsonStatus?.webViewLink ? `
	                <div class="json-readiness-list">
	                  <div class="json-readiness-item is-ready">
	                    <div>
	                      <strong>${esc(State.secDataImport.jsonStatus.fileName || secJsonFileName)}</strong>
	                      <span>อัปเดตล่าสุดแล้ว</span>
	                    </div>
                    <a class="badge badge-success" href="${esc(State.secDataImport.jsonStatus.webViewLink)}" target="_blank" rel="noopener noreferrer">เปิดไฟล์</a>
                  </div>
                </div>
              ` : ''}
              <div class="data-import-actions json-export-actions">
                <button class="btn btn-ghost" id="sec-open-json-folder" type="button">เปิดโฟลเดอร์ JSON</button>
                <button class="btn btn-primary import-commit-btn" id="sec-json-export" type="button" ${State.secDataImport.preview && !State.secDataImport.isExportingJson ? '' : 'disabled'}>
                  ${State.secDataImport.isExportingJson ? 'กำลังสร้าง JSON...' : 'สร้าง JSON'}
                </button>
              </div>
            </div>
          </div>
        </div>

      </div>`;

    const presetMap = {
      custom: null,
    };

    const syncSelectedState = () => {
      const selectedIds = $$('.sec-dataset-check:checked', area).map(input => input.value);
      const selectedSet = new Set(selectedIds);
      const selectedDatasetKeys = selectedIds.map(id => datasetKey(secDatasets.find(item => item.id === id) || { id }));
      const presetValue = 'custom';
      $('#sec-selected-count', area).textContent = String(selectedIds.length);
      $('#sec-dataset-command', area).textContent = selectedDatasetKeys.join(', ') || 'ยังไม่ได้เลือกหัวข้อมูล';
      $('#sec-workflow-preset', area).textContent = presetValue === 'custom' ? 'custom datasets' : presetValue;
      $('#sec-workflow-datasets', area).textContent = presetValue === 'custom' ? (selectedDatasetKeys.join(', ') || '-') : '-';
      $('#sec-workflow-output-mode', area).textContent = 'raw_tabs';
      $('#sec-workflow-registered-max', area).textContent = ($('#sec-registered-max-funds', area)?.value || '0').trim();
      $('#sec-workflow-max-pages', area).textContent = ($('#sec-max-pages', area)?.value || '0').trim();
      $$('[data-sec-dataset-row]', area).forEach(row => {
        row.classList.toggle('is-selected', selectedSet.has(row.dataset.secDatasetRow));
      });
      $$('[data-sec-dataset-cell]', area).forEach(cell => {
        const isSelected = selectedSet.has(cell.dataset.secDatasetCell);
        cell.innerHTML = `<span class="sec-dataset-status ${isSelected ? 'is-selected' : ''}">${isSelected ? 'เลือก' : '-'}</span>`;
      });
      $$('.sec-dataset-toggle-inline', area).forEach(label => {
        const input = $('input', label);
        const text = $('span', label);
        if (input && text) text.textContent = input.checked ? 'เลือก' : 'ข้าม';
      });
    };

    const readProjectPairs = () => $$('[data-sec-project-pair-row]', area).map(row => ({
      projId: $('.sec-project-pair-proj', row)?.value.trim() || '',
      fundClassName: $('.sec-project-pair-class', row)?.value.trim() || '',
    }));

    const activeProjectPairs = () => readProjectPairs().filter(pair => pair.projId);

    const projectPairPayload = () => activeProjectPairs()
      .map(pair => `${pair.projId}|${pair.fundClassName}`)
      .join('\n');

    const projectIdsPayload = () => {
      const ids = [];
      const seen = new Set();
      activeProjectPairs().forEach(pair => {
        if (!seen.has(pair.projId)) {
          seen.add(pair.projId);
          ids.push(pair.projId);
        }
      });
      return ids.join('\n');
    };

    const renderProjectPairEditorRows = pairs => {
      const list = $('#sec-project-pair-list', area);
      if (!list) return;
      list.innerHTML = renderProjectPairRows(pairs);
    };

    const saveSecImportPrefs = () => {
      writeSecImportPrefs({
        datasetPreset: $('#sec-dataset-preset', area)?.value || 'custom',
        selectedDatasetIds: $$('.sec-dataset-check:checked', area).map(input => input.value),
        fundStatus: $('#sec-fund-status', area)?.value || '',
        projectPairs: readProjectPairs(),
        projIds: projectIdsPayload(),
        fundClassName: '',
        startDate: $('#sec-start-date', area)?.value || '',
        endDate: $('#sec-end-date', area)?.value || '',
        startPeriod: $('#sec-start-period', area)?.value || '',
        endPeriod: $('#sec-end-period', area)?.value || '',
        registeredMaxFunds: $('#sec-registered-max-funds', area)?.value || '0',
        maxPages: $('#sec-max-pages', area)?.value || '0',
        outputMode: 'raw_tabs',
      });
    };

    const savedPreset = secImportPrefs.datasetPreset || 'custom';
    if ($('#sec-dataset-preset', area) && presetMap[savedPreset] !== undefined) {
      $('#sec-dataset-preset', area).value = savedPreset;
    }

    $$('.sec-dataset-check', area).forEach(input => {
      input.addEventListener('change', () => {
        $('#sec-dataset-preset', area).value = 'custom';
        syncSelectedState();
        saveSecImportPrefs();
      });
    });

    $('#sec-dataset-preset', area)?.addEventListener('change', e => {
      const values = presetMap[e.target.value];
      if (values) {
        const next = new Set(values);
        $$('.sec-dataset-check', area).forEach(input => {
          input.checked = next.has(input.value);
        });
      }
      syncSelectedState();
      saveSecImportPrefs();
    });

    let secWorkflowPollTimer = null;

    const updateWorkflowStatus = (message, tone = '') => {
      const statusEl = $('#sec-workflow-status', area);
      if (!statusEl) return;
      statusEl.hidden = !message;
      statusEl.className = `sec-workflow-status ${tone}`.trim();
      statusEl.innerHTML = message || '';
    };

    const workflowStatusLabel = run => {
      if (!run) return 'ยังไม่พบ run';
      if (run.status === 'completed') {
        if (run.conclusion === 'success') return 'สำเร็จ';
        if (run.conclusion === 'failure') return 'ล้มเหลว';
        if (run.conclusion === 'cancelled') return 'ถูกยกเลิก';
        return run.conclusion || 'เสร็จแล้ว';
      }
      if (run.status === 'in_progress') return 'กำลังรัน';
      if (run.status === 'queued') return 'รอคิว';
      return run.status || 'กำลังตรวจสอบ';
    };

    const workflowStatusTone = run => {
      if (!run) return 'is-running';
      if (run.status !== 'completed') return 'is-running';
      return run.conclusion === 'success' ? 'is-success' : 'is-error';
    };

    const workflowRunMessage = (run, prefix = 'สถานะล่าสุด') => {
      if (!run) return `${prefix}: ยังไม่พบ workflow run ล่าสุดจาก GitHub`;
      const label = workflowStatusLabel(run);
      const title = run.display_title || run.name || 'SEC Data Preparation to Google Sheet';
      const runNumber = run.run_number ? `#${run.run_number}` : '';
      const link = run.html_url ? ` <a href="${esc(run.html_url)}" target="_blank" rel="noopener noreferrer">เปิดรายละเอียด</a>` : '';
      return `${prefix}: <strong>${esc(label)}</strong> ${esc(runNumber)} - ${esc(title)}${link}`;
    };

    const fetchLatestWorkflowRun = async token => {
      const url = `${secWorkflowRunsUrl}?branch=main&event=workflow_dispatch&per_page=1`;
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          Accept: 'application/vnd.github+json',
          Authorization: `Bearer ${token}`,
          'X-GitHub-Api-Version': '2022-11-28',
        },
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(errorText || `GitHub API error ${response.status}`);
      }

      const payload = await response.json();
      return payload.workflow_runs?.[0] || null;
    };

    const stopWorkflowPolling = () => {
      if (!secWorkflowPollTimer) return;
      window.clearInterval(secWorkflowPollTimer);
      secWorkflowPollTimer = null;
    };

    const startWorkflowPolling = token => {
      stopWorkflowPolling();
      let pollCount = 0;
      const maxPolls = 240;

      const poll = async () => {
        pollCount += 1;
        try {
          const run = await fetchLatestWorkflowRun(token);
          updateWorkflowStatus(workflowRunMessage(run), workflowStatusTone(run));
          if (run?.status === 'completed' || pollCount >= maxPolls) {
            stopWorkflowPolling();
          }
        } catch (err) {
          stopWorkflowPolling();
          updateWorkflowStatus(`ส่งคำสั่งรันแล้ว แต่ตรวจสถานะไม่ได้: ${esc(err.message || 'GitHub API error')}`, 'is-warning');
        }
      };

      poll();
      secWorkflowPollTimer = window.setInterval(poll, 10000);
    };

	    const workflowInputs = () => {
	      const presetValue = 'custom';
	      const selectedIds = $$('.sec-dataset-check:checked', area).map(input => input.value);
	      const selectedDatasetKeys = selectedIds.map(id => datasetKey(secDatasets.find(item => item.id === id) || { id }));
	      const registeredMaxFunds = ($('#sec-registered-max-funds', area)?.value || '0').trim() || '0';
	      const maxPages = ($('#sec-max-pages', area)?.value || '0').trim() || '0';
	      return {
	        output_mode: 'raw_tabs',
	        tab_name: 'data_preparation',
	        output_spreadsheet_id: secSpreadsheetId,
	        dataset_preset: presetValue === 'custom' ? 'master_core' : presetValue,
	        custom_datasets: presetValue === 'custom' ? selectedDatasetKeys.join(',') : '',
	        proj_id: '',
	        proj_ids: projectIdsPayload(),
	        proj_class_pairs: projectPairPayload(),
	        fund_status: ($('#sec-fund-status', area)?.value || '').trim(),
	        use_existing_profiles: 'true',
	        registered_max_funds: registeredMaxFunds,
	        fund_class_name: '',
	        latest: 'true',
	        start_date: ($('#sec-start-date', area)?.value || '').trim(),
	        end_date: ($('#sec-end-date', area)?.value || '').trim(),
	        start_period: ($('#sec-start-period', area)?.value || '').trim(),
	        end_period: ($('#sec-end-period', area)?.value || '').trim(),
	        start_nav_date: ($('#sec-start-date', area)?.value || '').trim(),
	        end_nav_date: ($('#sec-end-date', area)?.value || '').trim(),
	        max_pages: maxPages,
	        continue_on_error: 'true',
	      };
	    };

    [
      'sec-fund-status',
      'sec-start-date',
      'sec-end-date',
      'sec-start-period',
      'sec-end-period',
      'sec-registered-max-funds',
      'sec-max-pages',
    ].forEach(id => {
      $(`#${id}`, area)?.addEventListener('input', () => {
        syncSelectedState();
        saveSecImportPrefs();
      });
      $(`#${id}`, area)?.addEventListener('change', () => {
        syncSelectedState();
        saveSecImportPrefs();
      });
    });

    $('#sec-add-project-pair', area)?.addEventListener('click', () => {
      renderProjectPairEditorRows([...readProjectPairs(), { projId: '', fundClassName: '' }]);
      saveSecImportPrefs();
    });

    $('#sec-project-pair-list', area)?.addEventListener('click', event => {
      const button = event.target.closest('.sec-project-pair-remove');
      if (!button) return;
      const rows = readProjectPairs();
      const rowEl = button.closest('[data-sec-project-pair-row]');
      const rowIndex = $$('[data-sec-project-pair-row]', area).indexOf(rowEl);
      const nextRows = rows.filter((_row, index) => index !== rowIndex);
      renderProjectPairEditorRows(nextRows.length ? nextRows : [{ projId: '', fundClassName: '' }]);
      saveSecImportPrefs();
    });

	    $('#sec-project-pair-list', area)?.addEventListener('input', saveSecImportPrefs);

	    const runSecWorkflow = async ({ button, inputs, runningMessage, successMessage }) => {
	      if (!button) return;
	      const token = (($('#sec-github-token', area)?.value || '').trim() || configuredGithubToken);
	      if (!token) {
	        updateWorkflowStatus('กรุณาใส่ GitHub token ก่อนรัน workflow', 'is-warning');
	        toast('กรุณาใส่ GitHub token ก่อนรัน workflow', 'warning');
	        return;
	      }

	      button.disabled = true;
	      stopWorkflowPolling();
	      updateWorkflowStatus(runningMessage || 'กำลังส่งคำสั่งรัน GitHub Actions...', 'is-running');

	      try {
	        const response = await fetch(secWorkflowDispatchUrl, {
	          method: 'POST',
	          headers: {
	            Accept: 'application/vnd.github+json',
	            Authorization: `Bearer ${token}`,
	            'Content-Type': 'application/json',
	            'X-GitHub-Api-Version': '2022-11-28',
	          },
	          body: JSON.stringify({
	            ref: 'main',
	            inputs,
	          }),
	        });

	        if (!response.ok) {
	          const errorText = await response.text();
	          if (response.status === 422 && errorText.includes('output_spreadsheet_id')) {
	            throw new Error('GitHub workflow บน repo ยังไม่รองรับ output_spreadsheet_id กรุณาอัปโหลดไฟล์ workflow เวอร์ชันล่าสุดก่อน แล้วค่อยรันใหม่');
	          }
	          throw new Error(errorText || `GitHub API error ${response.status}`);
	        }

	        updateWorkflowStatus('ส่งคำสั่งรันสำเร็จแล้ว กำลังตรวจสถานะจาก GitHub Actions...', 'is-running');
	        toast(successMessage || 'รัน GitHub Actions สำเร็จแล้ว', 'success', 5000);
	        startWorkflowPolling(token);
	      } catch (err) {
	        updateWorkflowStatus(err.message || 'รัน GitHub Actions ไม่สำเร็จ', 'is-error');
	        toast(err.message || 'รัน GitHub Actions ไม่สำเร็จ', 'error', 7000);
	      } finally {
	        button.disabled = false;
	      }
	    };

	    $('#sec-run-workflow', area)?.addEventListener('click', async () => {
	      const runButton = $('#sec-run-workflow', area);
	      const inputs = workflowInputs();
	      await runSecWorkflow({
	        button: runButton,
	        inputs,
	        runningMessage: 'กำลังส่งคำสั่งดึง raw tabs ไปที่ GitHub Actions...',
	        successMessage: 'รัน GitHub Actions สำเร็จแล้ว',
	      });
	    });

	    $('#sec-build-master-view', area)?.addEventListener('click', async () => {
	      const buildButton = $('#sec-build-master-view', area);
	      const inputs = {
	        ...workflowInputs(),
	        output_mode: 'master_from_raw',
	        tab_name: 'data_preparation',
	        dataset_preset: 'master_core',
	        custom_datasets: '',
	      };
	      await runSecWorkflow({
	        button: buildButton,
	        inputs,
	        runningMessage: 'กำลังส่งคำสั่งรวบรวม raw tabs เป็น data_preparation...',
	        successMessage: 'เริ่มรวบรวมข้อมูลเป็น data_preparation แล้ว',
	      });
	    });

    $('#sec-open-master-sheet', area)?.addEventListener('click', () => {
      window.open(`https://docs.google.com/spreadsheets/d/${secSpreadsheetId}/edit#gid=0`, '_blank', 'noopener,noreferrer');
    });

	    $('#sec-check-tabs', area)?.addEventListener('click', async () => {
	      const button = $('#sec-check-tabs', area);
	      button.disabled = true;
      try {
        const meta = await SheetsAPI.getSheetTabs(secSpreadsheetId);
        const tabSet = new Set(meta.tabs || []);
        const readRows = await Promise.all(secDatasets
          .map(async item => {
            if (!tabSet.has(item.id)) return { id: item.id, exists: false, rowCount: 0, coverage: null };
            try {
              const values = await SheetsAPI.fetchSheetData(secSpreadsheetId, item.id);
              return {
                id: item.id,
                exists: true,
                rowCount: Math.max(0, values.length - 1),
                coverage: secDataCoverage(values, item),
              };
            } catch {
              return { id: item.id, exists: true, rowCount: null, coverage: null };
            }
          }));
        const statusById = new Map(readRows.map(item => [item.id, item]));
        $$('[data-sec-tab-status-row]', area).forEach(row => {
          const status = statusById.get(row.dataset.secTabStatusRow);
          const dateCell = $('[data-sec-update-date-cell]', row);
          const statusCell = $('[data-sec-status-cell]', row);
          if (!status?.exists) {
            if (dateCell) dateCell.innerHTML = '<span class="sec-dataset-status is-missing">-</span>';
            if (statusCell) statusCell.innerHTML = '<span class="sec-dataset-status is-missing">ยังไม่มี</span>';
            return;
          }
          const suffix = Number.isFinite(status.rowCount) ? ` · ${status.rowCount} rows` : '';
          if (dateCell) {
            dateCell.innerHTML = secCoverageHtml(status);
          }
          if (statusCell) statusCell.innerHTML = `<span class="sec-dataset-status is-selected">พร้อม${esc(suffix)}</span>`;
        });
        toast('ตรวจสถานะ SEC tabs แล้ว', 'success');
      } catch (err) {
        toast(err.message || 'ตรวจสถานะ tab ไม่สำเร็จ', 'error', 7000);
      } finally {
	        button.disabled = false;
	      }
	    });

	    $('#sec-load-data-preparation-preview', area)?.addEventListener('click', async () => {
	      State.secDataImport.isLoadingPreview = true;
	      Pages.secDataImport(area);
	      try {
	        const [rows, targetMeta] = await Promise.all([
	          SheetsAPI.fetchSheetData(secSpreadsheetId, 'data_preparation'),
	          SheetsAPI.getSheetTabs(secOutputSpreadsheetId),
	        ]);
	        State.secDataImport.preview = refreshSecDataPreparationPreview(rows);
	        State.secDataImport.targetTab = State.secDataImport.targetTab || 'data_preparation';
	        State.secDataImport.targetExists = (targetMeta.tabs || []).includes(State.secDataImport.targetTab);
	        toast('โหลด data_preparation สำหรับกำหนด Type แล้ว', 'success');
	      } catch (err) {
	        toast(err.message || 'โหลด data_preparation ไม่สำเร็จ', 'error', 7000);
	      } finally {
	        State.secDataImport.isLoadingPreview = false;
	        Pages.secDataImport(area);
	      }
	    });

	    $('#sec-output-target-tab', area)?.addEventListener('input', e => {
	      State.secDataImport.targetTab = e.target.value.trim() || 'data_preparation';
	      State.secDataImport.jsonStatus = null;
	    });

	    $('#sec-output-target-tab', area)?.addEventListener('change', async e => {
	      State.secDataImport.targetTab = e.target.value.trim() || 'data_preparation';
	      State.secDataImport.jsonStatus = null;
	      try {
	        const meta = await SheetsAPI.getSheetTabs(secOutputSpreadsheetId);
	        State.secDataImport.targetExists = (meta.tabs || []).includes(State.secDataImport.targetTab);
	      } catch (err) {
	        toast(err.message || 'ตรวจสอบ Sheet ปลายทางไม่สำเร็จ', 'error', 6000);
	      }
	      Pages.secDataImport(area);
	    });

	    $('#sec-output-save', area)?.addEventListener('click', async () => {
	      const previewNow = State.secDataImport.preview;
	      const targetTab = ($('#sec-output-target-tab', area)?.value || State.secDataImport.targetTab || 'data_preparation').trim() || 'data_preparation';
	      if (!previewNow?.rows?.length) {
	        toast('ไม่มี data_preparation สำหรับบันทึก', 'warning');
	        return;
	      }
	      if (State.secDataImport.targetExists && !$('#sec-output-overwrite-confirm', area)?.checked) {
	        toast('กรุณายืนยันก่อนเขียนทับ Tab เดิม', 'warning');
	        return;
	      }
	      State.secDataImport.isImporting = true;
	      Pages.secDataImport(area);
	      try {
	        if (!State.secDataImport.targetExists) {
	          await SheetsAPI.addSheetTab(secOutputSpreadsheetId, targetTab);
	        } else {
	          await SheetsAPI.clearSheetValues(secOutputSpreadsheetId, targetTab);
	        }
	        await SheetsAPI.updateSheetValues(secOutputSpreadsheetId, targetTab, previewNow.rows);
	        State.secDataImport.targetTab = targetTab;
	        State.secDataImport.targetExists = true;
	        State.secDataImport.jsonStatus = null;
	        toast(`บันทึกข้อมูล ${previewNow.rowCount.toLocaleString()} rows ไปที่ ${targetTab} แล้ว`, 'success', 5000);
	      } catch (err) {
	        toast(err.message || 'บันทึกข้อมูล SEC ไม่สำเร็จ', 'error', 7000);
	      } finally {
	        State.secDataImport.isImporting = false;
	        Pages.secDataImport(area);
	      }
	    });

	    $('#sec-open-json-folder', area)?.addEventListener('click', () => {
	      window.open(secJsonDriveFolderUrl, '_blank', 'noopener,noreferrer');
	    });

	    $('#sec-json-export', area)?.addEventListener('click', async () => {
	      const previewNow = State.secDataImport.preview;
	      if (!previewNow?.rows?.length) {
	        toast('กรุณาโหลด data_preparation ก่อนสร้าง JSON', 'warning');
	        return;
	      }
	      const jsonPath = getSecJsonPathParts();
	      if (!/^\d{4}-Q[1-4]$/.test(jsonPath.quarter)) {
	        toast('กรุณาระบุชื่อ Tab ปลายทางเป็นรูปแบบ YYYY-Q1 ถึง YYYY-Q4 ก่อนสร้าง JSON', 'warning', 7000);
	        return;
	      }
	      State.secDataImport.isExportingJson = true;
	      State.secDataImport.jsonStatus = null;
	      Pages.secDataImport(area);
	      try {
	        const targetFolder = await SheetsAPI.resolveDriveFolderPath(
	          secJsonDriveFolderId,
	          jsonPath.folderSegments
	        );
	        const file = await SheetsAPI.uploadJsonToDriveFolder(
	          targetFolder.id,
	          secJsonFileName,
	          previewNow.rows
	        );
	        State.secDataImport.jsonStatus = {
	          ok: true,
	          fileId: file.id,
	          fileName: file.name || secJsonFileName,
	          path: `${JSON_STORE.rootName}/${jsonPath.year}/${jsonPath.quarter}/base/${secJsonFileName}`,
	          webViewLink: file.webViewLink || '',
	        };
	        toast(`สร้าง ${secJsonFileName} ใน ${jsonPath.quarter}/base แล้ว`, 'success', 5000);
	      } catch (err) {
	        State.secDataImport.jsonStatus = { ok: false, error: err.message || String(err) };
	        toast(err.message || 'สร้าง JSON ไม่สำเร็จ', 'error', 7000);
	      } finally {
	        State.secDataImport.isExportingJson = false;
	        Pages.secDataImport(area);
	      }
	    });

	    $$('.sec-import-type-select', area).forEach(select => {
	      select.addEventListener('change', e => {
	        const key = e.target.dataset.columnKey;
	        const value = e.target.value;
	        if (!key) return;
	        if (value === 'auto') {
	          delete State.secDataImport.columnTypes[key];
	        } else {
	          State.secDataImport.columnTypes[key] = value;
	        }
	        if (State.secDataImport.preview?.rows) {
	          State.secDataImport.preview = {
	            ...State.secDataImport.preview,
	            columns: analyzeSecImportColumns(State.secDataImport.preview.rows),
	          };
	        }
	        Pages.secDataImport(area);
	      });
	    });

	    syncSelectedState();
	  },

  /* ── MASTER MAPPING MANAGER ── */
  async fundDataManager(area) {
    setLoading(area, 'กำลังโหลดหน้าแก้ไขข้อมูลกองทุน...');

    let rawRows;
    let masterAnnualizedRows = [];
    let masterAnnualizedError = '';
    try {
      rawRows = await fetchCached('select-fund');
      await loadFundOverrides(true);
      await loadMasterAllocations(true);
    } catch (e) {
      setError(area, e.message, 'fund-data-manager');
      return;
    }
    try {
      masterAnnualizedRows = await fetchCached('master-annualized-v2');
    } catch (e) {
      masterAnnualizedError = e.message || String(e);
    }

    let catalog = applyFundOverrides(buildSelectedFundsCatalog(rawRows));
    const masterHeaders = masterAnnualizedRows[0] || [];
    const MASTER_CI = {
      name: findColumnIndex(masterHeaders, ['Group/Investment']),
      isin: findColumnIndex(masterHeaders, ['ISIN']),
      currency: findColumnIndex(masterHeaders, ['Base Currency']),
      r3m: findColumnIndex(masterHeaders, ['Return(Cumulative) 3M']),
      r6m: findColumnIndex(masterHeaders, ['Return(Cumulative) 6M']),
      rytd: findColumnIndex(masterHeaders, ['Return(Cumulative) YTD']),
      r1y: findColumnIndex(masterHeaders, ['Return(Cumulative) 1Y']),
      r3y: findColumnIndex(masterHeaders, ['Return(Annualized) 3Y']),
      r5y: findColumnIndex(masterHeaders, ['Return(Annualized) 5Y']),
      r10y: findColumnIndex(masterHeaders, ['Return(Annualized) 10Y']),
    };
    const masterGet = (row, idx) => idx >= 0 ? String(row[idx] ?? '').trim() : '';
    const masterAnnualizedByIsin = new Map();
    masterAnnualizedRows.slice(1).forEach(row => {
      const isin = masterGet(row, MASTER_CI.isin);
      if (isin && !masterAnnualizedByIsin.has(isin)) {
        masterAnnualizedByIsin.set(isin, row);
      }
    });
    const masterNameByIsin = new Map();
    const masterCurrencyByIsin = new Map();
    masterAnnualizedRows.slice(1).forEach(row => {
      const isin = masterGet(row, MASTER_CI.isin);
      if (!isin) return;
      if (!masterNameByIsin.has(isin)) masterNameByIsin.set(isin, masterGet(row, MASTER_CI.name));
      if (!masterCurrencyByIsin.has(isin)) masterCurrencyByIsin.set(isin, masterGet(row, MASTER_CI.currency));
    });
    const masterPerf = (masterId) => {
      const row = masterAnnualizedByIsin.get(String(masterId || '').trim());
      if (!row) return null;
      return {
        currency: masterGet(row, MASTER_CI.currency),
        r3m: masterGet(row, MASTER_CI.r3m),
        r6m: masterGet(row, MASTER_CI.r6m),
        rytd: masterGet(row, MASTER_CI.rytd),
        r1y: masterGet(row, MASTER_CI.r1y),
        r3y: masterGet(row, MASTER_CI.r3y),
        r5y: masterGet(row, MASTER_CI.r5y),
        r10y: masterGet(row, MASTER_CI.r10y),
      };
    };
    const masterOptions = [...new Map(catalog
      .filter(f => f.masterId || f.masterName)
      .map(f => [
        `${f.masterId || ''}__${f.masterName || ''}`,
        {
          masterId: f.masterId || '',
          masterName: f.masterName || '',
          label: `${f.masterId || '-'} · ${f.masterName || f.code}`,
        },
      ])).values()]
      .sort((a, b) => String(a.label).localeCompare(String(b.label), 'th'));
    masterAnnualizedByIsin.forEach((row, isin) => {
      if (masterOptions.some(option => option.masterId === isin)) return;
      masterOptions.push({
        masterId: isin,
        masterName: masterGet(row, MASTER_CI.name),
        label: `${isin} · ${masterGet(row, MASTER_CI.name)}`,
      });
    });
    masterOptions.sort((a, b) => String(a.label).localeCompare(String(b.label), 'th'));
    const opt = (val, label, cur) => `<option value="${esc(val)}" ${cur === val ? 'selected' : ''}>${esc(label)}</option>`;

    const findSelectedFund = () => {
      return catalog.find(f => f.key === State.fundDataManager.selectedKey) || null;
    };

    const getMapping = (fund) => {
      const draft = State.fundDataManager.draftMapping;
      if (draft?.key && draft.key === fund?.key) return draft;
      const saved = State.masterAllocations?.items?.[fund?.key || ''];
      if (saved?.allocations?.length) return saved;
      return {
        key: fund?.key || '',
        thaiFundCode: fund?.code || '',
        thaiFundName: fund?.name || '',
        allocations: [{
          masterId: fund?.masterId || '',
          masterName: fund?.masterName || '',
          weight: 100,
        }],
        note: '',
        sourceDate: '',
        status: 'Active',
      };
    };

    const renderImpactRows = () => {
      const impacts = [
        ['ค่าธรรมเนียม', 'สามารถนำ weight ไปคำนวณ Ongoing Cost / TER แบบถ่วงน้ำหนักได้'],
        ['Master Fund Report', 'ใช้แสดงว่ากองทุนไทย 1 กองอ้างอิง Master Fund มากกว่า 1 กอง'],
        ['Presentation / Export', 'ควรแสดง mapping และสัดส่วนเมื่อกองทุนมีหลาย Master'],
        ['Top 10 Holding', 'ไม่ใช้ mapping ชุดนี้ตาม scope ปัจจุบัน'],
      ];
      return impacts.map(([page, detail]) => `
        <tr>
          <td class="td-code">${esc(page)}</td>
          <td>${esc(detail)}</td>
          <td>${page === 'Top 10 Holding' ? '<span class="badge badge-primary">ไม่กระทบ</span>' : '<span class="badge badge-warning">ใช้เมื่อรองรับสูตร</span>'}</td>
        </tr>
      `).join('');
    };

    const fundSearchMatches = (searchValue = '') => {
      const searchQuery = String(searchValue || '').trim().toLowerCase();
      return catalog.filter(f => {
        const mapping = State.masterAllocations?.items?.[f.key];
        if (State.fundDataManager.showOnlyMapped && !mapping) return false;
        if (!searchQuery) return true;
        return [
          f.code,
          f.name,
          f.category,
          f.masterId,
          f.masterName,
        ].some(value => String(value || '').toLowerCase().includes(searchQuery));
      }).sort((a, b) => {
        const aMapped = !!State.masterAllocations?.items?.[a.key];
        const bMapped = !!State.masterAllocations?.items?.[b.key];
        if (aMapped !== bMapped) return aMapped ? -1 : 1;
        return String(a.code).localeCompare(String(b.code), 'th');
      });
    };

    const render = () => {
      let visible = fundSearchMatches(State.fundDataManager.query);

      if (!State.fundDataManager.selectedKey && visible[0]) {
        State.fundDataManager.selectedKey = visible[0].key;
      }
      const selected = findSelectedFund() || visible[0] || catalog[0] || null;
      if (selected?.key && State.fundDataManager.selectedKey !== selected.key) {
        State.fundDataManager.selectedKey = selected.key;
      }

      const mapping = selected ? getMapping(selected) : null;
      const allocationCount = Object.keys(State.masterAllocations?.items || {}).length;
      const totalWeight = (mapping?.allocations || []).reduce((sum, item) => sum + (parseFloat(item.weight) || 0), 0);
      const isBalanced = Math.abs(totalWeight - 100) <= 0.01;
      const weightedKeys = ['r3m','r6m','rytd','r1y','r3y','r5y','r10y'];
      const buildWeightedSummary = (allocations = []) => {
        const currencies = new Set();
        const values = {};
        weightedKeys.forEach(key => { values[key] = 0; });

        allocations.forEach(item => {
          const weight = parseFloat(item.weight) || 0;
          const perf = masterPerf(item.masterId);
          const currency = masterCurrencyByIsin.get(item.masterId) || item.baseCurrency || '';
          if (currency) currencies.add(currency);
          weightedKeys.forEach(key => {
            const value = parseNum(perf?.[key] ?? item.returns?.[key]);
            if (!Number.isNaN(value)) values[key] += value * weight / 100;
          });
        });

        return {
          masterId: allocations.map(item => item.masterId).filter(Boolean).join(', ') || '-',
          masterName: allocations.map(item => {
            const name = item.masterName || masterNameByIsin.get(item.masterId) || item.masterId || '-';
            const weight = parseFloat(item.weight) || 0;
            return `${name} (${weight.toFixed(2)}%)`;
          }).join(' + ') || '-',
          baseCurrency: currencies.size > 1 ? 'Mixed' : ([...currencies][0] || '-'),
          weight: allocations.reduce((sum, item) => sum + (parseFloat(item.weight) || 0), 0),
          values,
        };
      };
      const weightedSummary = buildWeightedSummary(mapping?.allocations || []);
      const allocationRows = (mapping?.allocations || []).map((item, idx) => {
        const perf = masterPerf(item.masterId);
        const hasSystemMaster = !!perf;
        const masterName = hasSystemMaster
          ? (masterNameByIsin.get(item.masterId) || item.masterName || '')
          : (item.masterName || '');
        const baseCurrency = hasSystemMaster
          ? (masterCurrencyByIsin.get(item.masterId) || item.baseCurrency || '')
          : (item.baseCurrency || '');
        return `
          <tr data-alloc-index="${idx}" class="${hasSystemMaster ? '' : 'is-manual-master'}">
            <td class="td-check">
              <button class="btn btn-danger btn-xs alloc-remove" type="button" ${mapping.allocations.length <= 1 ? 'disabled' : ''}>ลบ</button>
            </td>
            <td>
              <input class="fund-input fund-input-editable alloc-master-id" list="fdm-master-id-list" value="${esc(item.masterId || '')}" placeholder="ISIN / Master ID">
            </td>
            <td>
              <input class="fund-input ${hasSystemMaster ? '' : 'fund-input-editable'} alloc-master-name" value="${esc(masterName)}" placeholder="Master Fund name" ${hasSystemMaster ? 'readonly' : ''}>
            </td>
            <td>
              <input class="fund-input ${hasSystemMaster ? '' : 'fund-input-editable'} alloc-currency" value="${esc(baseCurrency)}" placeholder="Base Currency" ${hasSystemMaster ? 'readonly' : ''}>
            </td>
            <td class="td-num">
              <input class="fund-input fund-input-editable alloc-weight" type="number" min="0" max="100" step="0.01" value="${esc(item.weight)}">
            </td>
            ${weightedKeys.map(key => {
              const value = perf?.[key] ?? item.returns?.[key] ?? '';
              return hasSystemMaster
                ? `<td class="td-num alloc-return-${key}">${esc(formatReturnDisplay(value) || '-')}</td>`
                : `<td class="td-num"><input class="fund-input fund-input-editable alloc-return-input alloc-return-${key}" data-return-key="${key}" type="number" step="0.01" value="${esc(value)}" placeholder="-"></td>`;
            }).join('')}
          </tr>`;
      }).join('');
      const originalPerf = selected ? masterPerf(selected.masterId) : null;
      const originalSourceRow = selected ? `
        <tr>
          <td class="td-check"><span class="badge badge-primary">ข้อมูลเดิม</span></td>
          <td class="td-isin">${esc(selected.masterId || '-')}</td>
          <td>${esc(selected.masterName || '-')}</td>
          <td>${esc(originalPerf?.currency || '-')}</td>
          <td class="td-num">100.00</td>
          ${['r3m','r6m','rytd','r1y','r3y','r5y','r10y'].map(key => `<td class="td-num">${esc(formatReturnDisplay(originalPerf?.[key]) || '-')}</td>`).join('')}
        </tr>
      ` : '';
      const weightedSummaryRow = `
        <tr>
          <td class="td-check"><span class="badge badge-success">หลังแก้ไข</span></td>
          <td class="td-isin after-edit-master-id">${esc(weightedSummary.masterId)}</td>
          <td class="after-edit-master-name">${esc(weightedSummary.masterName)}</td>
          <td class="after-edit-currency">${esc(weightedSummary.baseCurrency)}</td>
          <td class="td-num after-edit-weight">${weightedSummary.weight.toFixed(2)}</td>
          ${weightedKeys.map(key => `<td class="td-num after-edit-${key}">${esc(formatReturnDisplay(weightedSummary.values[key]) || '-')}</td>`).join('')}
        </tr>
      `;
      const masterMappingColgroup = `
        <col class="mm-col-action">
        <col class="mm-col-id">
        <col class="mm-col-name">
        <col class="mm-col-currency">
        <col class="mm-col-weight">
        <col class="mm-col-return">
        <col class="mm-col-return">
        <col class="mm-col-return">
        <col class="mm-col-return">
        <col class="mm-col-return">
        <col class="mm-col-return">
        <col class="mm-col-return">
      `;

      area.innerHTML = `
        <div class="fund-manager-grid">
          <div class="card fund-manager-picker-card">
            <div class="sf-filterbar">
              <div class="sf-search fund-manager-search">
                <span class="s-icon">${searchIcon()}</span>
                <input class="search-input" id="fdm-q" type="text" placeholder="ค้นหากองทุนไทย / Fund Code / Master ID..." value="${esc(State.fundDataManager.query)}" autocomplete="off" aria-autocomplete="list" aria-controls="fdm-search-menu">
                <div class="fund-search-menu" id="fdm-search-menu" role="listbox"></div>
              </div>
              <label class="fund-manager-check">
                <input type="checkbox" id="fdm-only-mapped" ${State.fundDataManager.showOnlyMapped ? 'checked' : ''}>
                แสดงเฉพาะที่มี mapping
              </label>
              <span class="row-count-badge">${visible.length.toLocaleString()} รายการ</span>
              <span class="row-count-badge is-info">${allocationCount.toLocaleString()} mapping</span>
              <span class="badge badge-data-origin">Data/fund_master_allocations.json</span>
              ${State.masterAllocationsLoadError ? `<span class="badge badge-warning">ยังไม่ได้เชื่อมต่อ local server</span>` : ''}
            </div>
          </div>

          <div class="card fund-manager-editor-card">
            <div class="card-header">
              <span class="card-title">${selected ? `Master Mapping · ${esc(selected.code)}` : 'Master Mapping'}</span>
              <div class="page-tools-meta">
                <span class="badge ${isBalanced ? 'badge-success' : 'badge-danger'}">Weight รวม ${totalWeight.toFixed(2)}%</span>
                ${State.masterAllocations?.items?.[selected?.key] ? '<span class="badge badge-accent">Custom Mapping</span>' : '<span class="badge badge-primary">Default 100%</span>'}
              </div>
            </div>
            <div class="card-body">
              ${selected ? `
              <div class="fund-manager-summary">
                <div><span>กองทุนไทย</span><strong>${esc(selected.name || selected.code)}</strong></div>
                <div><span>Fund Code</span><strong>${esc(selected.code)}</strong></div>
                <div><span>Default Master</span><strong>${esc(selected.masterId || '-')}</strong></div>
              </div>
              <div class="fund-readonly-block">
                <div class="fund-subsection-head">
                  <span>ข้อมูลเดิมจากต้นทาง</span>
                  ${masterAnnualizedError ? '<span class="badge badge-warning">ยังโหลด Master Annualized ไม่ได้</span>' : '<span class="badge badge-data-origin">Master Fund Annualized Return</span>'}
                </div>
                <div class="table-wrapper fund-original-table">
                  <table class="annualized-report master-mapping-report">
                    <colgroup>${masterMappingColgroup}</colgroup>
                    <thead>
                      <tr class="report-group-row">
                        <th colspan="5" class="group-blank"></th>
                        <th colspan="7" class="group-blue">ผลตอบแทน (%)</th>
                      </tr>
                      <tr>
                        <th></th>
                        <th>Master ID / ISIN</th>
                        <th>Master Fund</th>
                        <th>Base Currency</th>
                        <th>Weight %</th>
                        <th>3M</th>
                        <th>6M</th>
                        <th>YTD</th>
                        <th>1Y</th>
                        <th>3Y</th>
                        <th>5Y</th>
                        <th>10Y</th>
                      </tr>
                    </thead>
                    <tbody>${originalSourceRow || '<tr><td colspan="12" class="text-muted">ไม่พบข้อมูลต้นทาง</td></tr>'}</tbody>
                  </table>
                </div>
              </div>
              <form id="fdm-form">
                <datalist id="fdm-master-id-list">${masterOptions.map(m => `<option value="${esc(m.masterId)}">${esc(m.label)}</option>`).join('')}</datalist>
                <div class="fund-edit-block">
                  <div class="fund-subsection-head">
                    <span>แก้ไข Master Mapping</span>
                    <div class="fund-subsection-actions">
                      <button class="btn btn-primary btn-sm" type="button" id="fdm-add-row">เพิ่ม Master Fund</button>
                      <span class="badge ${isBalanced ? 'badge-success' : 'badge-danger'}">Weight รวม ${totalWeight.toFixed(2)}%</span>
                    </div>
                  </div>
                  <div class="table-wrapper fund-allocation-table">
                    <table class="annualized-report master-mapping-report">
                      <colgroup>${masterMappingColgroup}</colgroup>
                      <thead>
                        <tr class="report-group-row">
                          <th colspan="5" class="group-blank"></th>
                          <th colspan="7" class="group-blue">ผลตอบแทน (%)</th>
                        </tr>
                        <tr>
                          <th></th>
                          <th>Master ID / ISIN</th>
                          <th>Master Fund</th>
                          <th>Base Currency</th>
                          <th>Weight %</th>
                          <th>3M</th>
                          <th>6M</th>
                          <th>YTD</th>
                          <th>1Y</th>
                          <th>3Y</th>
                          <th>5Y</th>
                          <th>10Y</th>
                        </tr>
                      </thead>
                      <tbody>${allocationRows}</tbody>
                    </table>
                  </div>
                </div>
                <div class="fund-after-edit-block">
                  <div class="fund-subsection-head">
                    <span>ข้อมูลหลังการแก้ไข</span>
                    <span class="badge badge-data-origin">คำนวณแบบถ่วงน้ำหนักตาม Weight %</span>
                  </div>
                  <div class="table-wrapper fund-after-edit-table">
                    <table class="annualized-report master-mapping-report">
                      <colgroup>${masterMappingColgroup}</colgroup>
                      <thead>
                        <tr class="report-group-row">
                          <th colspan="5" class="group-blank"></th>
                          <th colspan="7" class="group-blue">ผลตอบแทน (%)</th>
                        </tr>
                        <tr>
                          <th></th>
                          <th>Master ID / ISIN</th>
                          <th>Master Fund</th>
                          <th>Base Currency</th>
                          <th>Weight %</th>
                          <th>3M</th>
                          <th>6M</th>
                          <th>YTD</th>
                          <th>1Y</th>
                          <th>3Y</th>
                          <th>5Y</th>
                          <th>10Y</th>
                        </tr>
                      </thead>
                      <tbody>${weightedSummaryRow}</tbody>
                    </table>
                  </div>
                </div>
                <div class="fund-form-actions">
                  <button class="btn btn-primary" type="submit">บันทึก Mapping</button>
                  <button class="btn btn-danger" type="button" id="fdm-delete" ${State.masterAllocations?.items?.[selected.key] ? '' : 'disabled'}>ลบ mapping</button>
                </div>
              </form>` : '<div class="state-box">ไม่พบกองทุนให้จัดการ</div>'}
            </div>
          </div>

          <div class="card fund-manager-impact-card">
            <div class="card-header">
              <span class="card-title">ผลกระทบของข้อมูลที่แก้</span>
            </div>
            <div class="table-wrapper fund-impact-table">
              <table>
                <thead>
                  <tr>
                    <th>ส่วนของระบบ</th>
                    <th>ผลกระทบ</th>
                    <th>ระดับตรวจสอบ</th>
                  </tr>
                </thead>
                <tbody>${renderImpactRows()}</tbody>
              </table>
            </div>
          </div>
        </div>`;

      const qEl = $('#fdm-q', area);
      const menuEl = $('#fdm-search-menu', area);
      let activeSearchIndex = -1;
      const applyFundSelection = (fund) => {
        if (!fund) return;
        State.fundDataManager.draftMapping = null;
        State.fundDataManager.selectedKey = fund.key;
        State.fundDataManager.query = fund.code;
        render();
      };
      const setActiveSearchOption = (index) => {
        if (!menuEl) return;
        const options = $$('.fund-search-option', menuEl);
        activeSearchIndex = options.length ? Math.max(0, Math.min(index, options.length - 1)) : -1;
        options.forEach((option, idx) => {
          const isActive = idx === activeSearchIndex;
          option.classList.toggle('is-active', isActive);
          option.setAttribute('aria-selected', isActive ? 'true' : 'false');
        });
      };
      const renderFundSearchMenu = (forceOpen = false) => {
        if (!menuEl || !qEl) return [];
        const matches = fundSearchMatches(qEl.value).slice(0, 8);
        activeSearchIndex = -1;
        if (!forceOpen && document.activeElement !== qEl) {
          menuEl.classList.remove('is-open');
          return matches;
        }
        menuEl.innerHTML = matches.length ? matches.map((f, idx) => {
          const mapped = !!State.masterAllocations?.items?.[f.key];
          return `
            <button class="fund-search-option" type="button" role="option" data-index="${idx}" aria-selected="false">
              <span class="fund-search-code">${esc(f.code || '-')}</span>
              <span class="fund-search-name">${esc(f.name || f.masterName || '-')}</span>
              <span class="fund-search-meta">${esc(f.masterId || '-')} ${mapped ? '· mapping' : ''}</span>
            </button>
          `;
        }).join('') : '<div class="fund-search-empty">ไม่พบกองทุนที่ตรงกับคำค้น</div>';
        menuEl.classList.add('is-open');
        $$('.fund-search-option', menuEl).forEach(option => {
          option.addEventListener('mousedown', e => e.preventDefault());
          option.addEventListener('click', () => {
            const idx = parseInt(option.dataset.index || '-1', 10);
            applyFundSelection(matches[idx]);
          });
        });
        return matches;
      };
      const selectFundFromSearch = () => {
        State.fundDataManager.query = qEl?.value.trim() || '';
        const normalizedQuery = normalizeFundKey(State.fundDataManager.query);
        const exact = catalog.find(f => normalizeFundKey(f.code) === normalizedQuery);
        const currentMatches = fundSearchMatches(State.fundDataManager.query);
        const fallback = exact || currentMatches[0] || catalog.find(f => [
          f.code,
          f.name,
          f.category,
          f.masterId,
          f.masterName,
        ].some(value => String(value || '').toLowerCase().includes(State.fundDataManager.query.toLowerCase())));
        applyFundSelection(fallback);
      };
      qEl?.addEventListener('input', () => {
        State.fundDataManager.query = qEl.value.trim();
        renderFundSearchMenu(true);
      });
      qEl?.addEventListener('focus', () => {
        renderFundSearchMenu(true);
      });
      qEl?.addEventListener('blur', () => {
        window.setTimeout(() => menuEl?.classList.remove('is-open'), 120);
      });
      qEl?.addEventListener('keydown', e => {
        if (['ArrowDown', 'ArrowUp'].includes(e.key)) {
          e.preventDefault();
          const matches = fundSearchMatches(qEl.value).slice(0, 8);
          if (!menuEl?.classList.contains('is-open')) renderFundSearchMenu(true);
          if (matches.length) setActiveSearchOption(activeSearchIndex + (e.key === 'ArrowDown' ? 1 : -1));
        } else if (e.key === 'Enter') {
          e.preventDefault();
          const matches = fundSearchMatches(qEl.value).slice(0, 8);
          if (activeSearchIndex >= 0 && matches[activeSearchIndex]) {
            applyFundSelection(matches[activeSearchIndex]);
          } else {
            selectFundFromSearch();
          }
        } else if (e.key === 'Escape') {
          menuEl?.classList.remove('is-open');
        }
      });
      qEl?.addEventListener('change', selectFundFromSearch);
      $('#fdm-only-mapped', area)?.addEventListener('change', e => {
        State.fundDataManager.showOnlyMapped = e.target.checked;
        State.fundDataManager.selectedKey = '';
        render();
      });
      const readAllocationsFromForm = () => $$('.fund-allocation-table tbody tr', area)
        .map(row => {
          const masterId = $('.alloc-master-id', row)?.value.trim() || '';
          const returns = {};
          $$('.alloc-return-input', row).forEach(input => {
            const key = input.dataset.returnKey;
            if (key) returns[key] = input.value.trim();
          });
          return {
            masterId,
            masterName: masterNameByIsin.get(masterId) || $('.alloc-master-name', row)?.value.trim() || '',
            baseCurrency: masterCurrencyByIsin.get(masterId) || $('.alloc-currency', row)?.value.trim() || '',
            weight: parseFloat($('.alloc-weight', row)?.value || '0') || 0,
            returns,
          };
        })
        .filter(item => item.masterId || item.masterName || item.weight);
      const updateWeightedSummary = () => {
        const summary = buildWeightedSummary(readAllocationsFromForm());
        const setText = (selector, value) => {
          const el = $(selector, area);
          if (el) el.textContent = value;
        };
        setText('.after-edit-master-id', summary.masterId);
        setText('.after-edit-master-name', summary.masterName);
        setText('.after-edit-currency', summary.baseCurrency);
        setText('.after-edit-weight', summary.weight.toFixed(2));
        weightedKeys.forEach(key => {
          setText(`.after-edit-${key}`, formatReturnDisplay(summary.values[key]) || '-');
        });
      };
      const updateLocalMapping = (allocations) => {
        if (!selected) return;
        State.fundDataManager.draftMapping = {
          ...(State.masterAllocations?.items?.[selected.key] || {}),
          key: selected.key,
          thaiFundCode: selected.code,
          thaiFundName: selected.name || '',
          allocations,
          note: mapping.note || '',
          sourceDate: mapping.sourceDate || '',
          status: mapping.status || 'Active',
        };
      };
      $('#fdm-add-row', area)?.addEventListener('click', () => {
        const allocations = readAllocationsFromForm();
        allocations.push({ masterId: '', masterName: '', weight: Math.max(0, 100 - allocations.reduce((sum, item) => sum + item.weight, 0)) });
        updateLocalMapping(allocations);
        render();
      });
      $$('.alloc-remove', area).forEach(btn => {
        btn.addEventListener('click', () => {
          const idx = parseInt(btn.closest('tr')?.dataset.allocIndex || '-1', 10);
          const allocations = readAllocationsFromForm().filter((_, i) => i !== idx);
          updateLocalMapping(allocations.length ? allocations : [{ masterId: '', masterName: '', weight: 100 }]);
          render();
        });
      });
      $$('.alloc-master-id', area).forEach(input => {
        input.addEventListener('change', () => {
          updateLocalMapping(readAllocationsFromForm());
          render();
        });
      });
      $$('.alloc-weight, .alloc-return-input, .alloc-master-name, .alloc-currency', area).forEach(input => {
        input.addEventListener('input', updateWeightedSummary);
      });
      $('#fdm-delete', area)?.addEventListener('click', async () => {
        if (!selected.key) return;
        try {
          const result = await deleteMasterAllocation(selected.key);
          State.fundDataManager.draftMapping = null;
          State.fundDataManager.selectedKey = selected.key;
          toast(result.warning || 'ลบ mapping แล้ว', result.warning ? 'warning' : 'success');
          render();
        } catch (err) {
          toast(err.message || 'ลบไม่สำเร็จ', 'error');
        }
      });
      $('#fdm-form', area)?.addEventListener('submit', async e => {
        e.preventDefault();
        if (!selected) return;
        const payload = {
          quarter: State.currentQuarter || CONFIG.PAGES?.['select-fund']?.tabName || '2026-Q1',
          thaiFundCode: selected.code,
          thaiFundName: selected.name || '',
          allocations: readAllocationsFromForm(),
          sourceDate: mapping.sourceDate || '',
          status: mapping.status || 'Active',
          note: mapping.note || '',
        };
        try {
          const result = await saveMasterAllocation(payload);
          const saved = result.item || {};
          State.fundDataManager.draftMapping = null;
          State.fundDataManager.selectedKey = saved.key || selected.key;
          toast(result.warning || 'บันทึก Master Mapping แล้ว', result.warning ? 'warning' : 'success');
          render();
        } catch (err) {
          toast(err.message || 'บันทึกไม่สำเร็จ', 'error');
        }
      });
    };

    render();
    App._currentExport = null;
  },

  /* ── SELECT FUND ── */
  async selectFund(area) {
    const cfg = CONFIG.PAGES['select-fund'];
    setLoading(area, 'กำลังโหลดรายการกองทุน...');

    let rawRows;
    try {
      rawRows = await fetchCached('select-fund');
      await Promise.all([
        loadFundOverrides(),
        loadMasterAllocations(true),
      ]);
    } catch (e) {
      setError(area, e.message, 'select-fund');
      return;
    }

    const headers = rawRows[0] || [];
    const CI = {
      CATEGORY: findColumnIndex(headers, ['AVP® Category', 'AVP®  Category', 'AVP Category']),
    };
    const allFunds = applyFundOverrides(buildSelectedFundsCatalog(rawRows));

    State.selectedFunds = Object.fromEntries(allFunds.map(f => [f.key, f]));

    /* Unique dropdown options */
    const categories = [...new Set(allFunds.map(f => f.category).filter(Boolean))].sort();

    const opt = (val, label, cur) =>
      `<option value="${esc(val)}" ${cur === val ? 'selected' : ''}>${esc(label)}</option>`;

    const getSavedMasterAllocations = (fund) => {
      const saved = State.masterAllocations?.items?.[fund.key];
      const allocations = Array.isArray(saved?.allocations) ? saved.allocations : [];
      return allocations.filter(item => item?.masterId || item?.masterName);
    };
    const masterMappingDisplay = (fund) => {
      const allocations = getSavedMasterAllocations(fund);
      if (!allocations.length) {
        return {
          isMapped: false,
          masterId: fund.masterId,
          masterName: fund.masterName,
          searchText: `${fund.masterId || ''} ${fund.masterName || ''}`,
        };
      }
      const masterId = allocations.map(item => item.masterId).filter(Boolean).join(', ') || '-';
      const masterName = allocations.map(item => {
        const name = item.masterName || item.masterId || '-';
        const weight = parseFloat(item.weight);
        return Number.isNaN(weight) ? name : `${name} (${weight.toFixed(2)}%)`;
      }).join(' + ');
      return {
        isMapped: true,
        masterId,
        masterName,
        searchText: allocations.map(item => `${item.masterId || ''} ${item.masterName || ''}`).join(' '),
      };
    };

    const render = (goPage = 1, opts = {}) => {
      const preserveScroll = !!opts.preserveScroll;
      const prevScrollTop = preserveScroll ? area.scrollTop : 0;
      State.tablePage = goPage;
      State.pageSize = State.selectFundFilters.pageSize || SELECT_FUND_DEFAULT_PAGE_SIZE;

      /* Filter */
      let visible = allFunds.filter(f => {
        const mappingText = masterMappingDisplay(f).searchText.toLowerCase();
        if (State.selectFundFilters.query && !f.code.toLowerCase().includes(State.selectFundFilters.query) && !f.masterName.toLowerCase().includes(State.selectFundFilters.query) && !mappingText.includes(State.selectFundFilters.query)) return false;
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
          if (key === 'focusGroup') return State.incomeFundSelectedKeys.has(fund.key) ? 1 : 0;
          if (key === 'masterId' || key === 'masterName') {
            return masterMappingDisplay(fund)[key];
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
        const mapping = masterMappingDisplay(f);
        return `
          <tr data-key="${esc(f.key)}" class="${isSelected ? 'row-selected' : ''}">
            <td class="td-check">
              <input type="checkbox" class="row-chk" data-key="${esc(f.key)}" ${isSelected ? 'checked' : ''}>
            </td>
            <td class="td-code">${esc(f.code)}</td>
            <td class="td-hl">${buildHighlightSelect(f.key, State.highlights[f.key])}</td>
            <td>${esc(f.dividend)}</td>
            <td>${esc(f.style)}</td>
            <td class="td-isin">${esc(mapping.masterId)}</td>
            <td>${esc(mapping.masterName)}${mapping.isMapped ? ' <span class="badge badge-success fund-override-mini">Master Mapping</span>' : ''}${f.isOverride ? ' <span class="badge badge-accent fund-override-mini">แก้ไขแล้ว</span>' : ''}</td>
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
            ${Object.keys(State.fundOverrides?.items || {}).length > 0
              ? `<span class="badge badge-success">ข้อมูลแก้ไข ${Object.keys(State.fundOverrides.items).length} รายการ</span>`
              : ''}
            ${Object.keys(State.masterAllocations?.items || {}).length > 0
              ? `<span class="badge badge-success">Master Mapping ${Object.keys(State.masterAllocations.items).length} รายการ</span>`
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
        .map(f => {
          const mapping = masterMappingDisplay(f);
          return [
            f.code,
            f.category,
            f.type,
            f.dividend,
            f.style,
            mapping.masterId,
            mapping.masterName,
            State.highlights[f.key] !== undefined ? HL_COLORS[State.highlights[f.key]].name : '',
          ];
        });

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
        CONFIG.PAGES['select-fund']?.source || 'Fund Key Performance AVP',
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
        buildSelectedFeeUniverse(),
      ]);
      const rawLookup = buildRawSecLookup(rawSecRows);
      const feeRows = buildFeeComparisonRows(universe, rawLookup, { includeThaiOnly: true })
        .sort((a, b) => compareValues(a.combined, b.combined, 'asc'));

      if (!feeRows.length) {
        setError(area, 'ไม่พบข้อมูลค่าธรรมเนียมที่จับคู่ได้จาก Data For SEC API และ AVP Master Fund ID', 'master-placeholder-1');
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
        ${pageToolActions('master-placeholder-1', CONFIG.PAGES['master-placeholder-1']?.source || 'Data For SEC API + AVP Master Fund ID')}
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
              <p>อ้างอิง TER จาก Data For SEC API และ Ongoing Cost จาก AVP Master Fund ID</p>
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
        CONFIG.PAGES['master-placeholder-1']?.source || 'Data For SEC API + AVP Master Fund ID',
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
        buildSelectedFeeUniverse(),
      ]);
      const rawLookup = buildRawSecLookup(rawSecRows);
      const feeRows = buildFeeComparisonRows(universe, rawLookup, { includeThaiOnly: true })
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
        setError(area, 'ไม่พบข้อมูลค่าธรรมเนียมที่จับคู่ได้จาก Data For SEC API และ AVP Master Fund ID', 'master-placeholder-4');
        return;
      }

      const source = CONFIG.PAGES['master-placeholder-4']?.source || 'Data For SEC API + AVP Master Fund ID';
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
        .of-period-btn{padding:4px 13px;border:1px solid #d0d9e8;border-radius:20px;background:#fff;color:#334155;font-size:0.88rem;cursor:pointer;font-family:inherit;transition:all .15s}
        .of-period-btn.is-active{background:#1a2744;color:#fff;border-color:#1a2744;font-weight:700}
        .of-period-btn:hover:not(.is-active){background:#eef2f7}
        .of-axis-label{font-size:0.88rem;color:#64748b;font-weight:600;white-space:nowrap}
        .of-axis-select{padding:4px 8px;border:1px solid #d0d9e8;border-radius:6px;background:#fff;color:#1e293b;font-size:0.88rem;font-family:inherit;cursor:pointer;max-width:200px}
        .of-count{font-size:0.88rem;color:#475569;padding:4px 12px;background:#e8edf5;border-radius:20px;white-space:nowrap}
        .of-reset-label-btn{padding:6px 12px;font-size:0.88rem}
        .of-scatter-title{padding:12px 16px;font-size:1.375rem;font-weight:700;color:#1a2744;background:#eef2fa;border-bottom:1px solid #dbe4f0;letter-spacing:.01em;line-height:1.25}
        .of-scatter-section{display:flex;align-items:flex-start;gap:24px;padding:0 16px 0 0;border-bottom:2px solid #e2e8f0;background:#fff}
        .of-scatter-left{flex:4;min-width:0;padding:12px 0 12px 12px}
        .of-scatter-right{flex:3;min-width:260px;overflow-y:auto;max-height:530px;padding-left:4px}
        .of-svg{width:100%;display:block}
        .of-no-data{padding:60px 20px;text-align:center;color:#94a3b8;font-size:0.9rem}
        .of-side-table{width:100%;border-collapse:collapse;font-size:0.88rem;font-family:'Sarabun','THSarabunNew',sans-serif;table-layout:fixed}
        .of-side-col-fund{width:52%}
        .of-side-col-metric{width:24%}
        .of-side-th{padding:8px 10px;background:#1a2744;color:#fff;font-weight:600;font-size:0.88rem;white-space:nowrap;position:sticky;top:0;z-index:1}
        .of-side-th-sort{cursor:pointer;user-select:none}
        .of-side-th-sort:hover{background:#223257}
        .of-side-val{text-align:right}
        .of-side-val-mid{text-align:center!important}
        .of-side-table tbody tr{border-bottom:1px solid #eef2f7}
        .of-side-table tbody tr:nth-child(even){background:#f8fafc}
        .of-side-table tbody tr:hover td{background:#e8f0fb!important}
        .of-side-fund-td{padding:5px 10px;vertical-align:middle}
        .of-side-fund-inner{display:flex;align-items:center;gap:6px;font-size:0.88rem;font-weight:600;color:#1e293b}
        .of-side-table td.of-side-val{padding:5px 10px;text-align:right;color:#334155;font-weight:500;font-size:0.88rem}
        #of-table-container{margin-top:18px}
        .of-table-wrap{overflow-x:auto}
        .of-table{width:100%;border-collapse:collapse;font-size:0.88rem;font-family:'Sarabun','THSarabunNew',sans-serif}
        .of-th{padding:8px 10px;background:#1a2744;color:#fff;font-weight:600;font-size:0.88rem;white-space:nowrap;text-align:center;border-right:1px solid #2d3f6b}
        .of-th-cb{width:36px;text-align:center;background:#0f1f3d}
        .of-th-hl{width:72px;background:#0f1f3d}
        .of-th-name{text-align:left;min-width:180px;background:#0f1f3d}
        .of-th-thai{text-align:left;min-width:110px}
        .of-td{padding:6px 10px;border-bottom:1px solid #eef2f7;border-right:1px solid #f1f5fb;text-align:right;vertical-align:middle}
        .of-td-cb{text-align:center;background:#f8fafc;width:36px}
        .of-td-hl{text-align:center;background:#f8fafc;position:relative;overflow:visible}
        .of-td-name{text-align:left;background:#f8fafc;display:flex;align-items:center;gap:6px}
        .of-td-thai{text-align:left;color:#475569;font-size:0.88rem}
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
        .of-palette-swatch.is-clear{width:18px;height:18px;border-radius:999px;background:#f8fafc;color:#64748b;border:1px solid #cbd5e1;font-size:14px;line-height:1}
        .of-pos{color:#16a34a;font-weight:600}.of-neg{color:#dc2626;font-weight:600}.of-zero{color:#64748b}.of-na{color:#94a3b8;font-size:0.88rem}
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

  async otherFactorsFixedIncomeTable(area) {
    const pageKey = 'master-placeholder-9';
    setLoading(area, 'กำลังโหลดตารางปัจจัยประกอบกองทุนตราสารหนี้...');

    let rows;
    let editableColumns = [];
    try {
      await ensureSelectedFundsCatalog();
      const [thaiQualityRows, secRows] = await Promise.all([
        fetchCached('thai-annualized-v2'),
        fetchCached('master-placeholder-4'),
      ]);

      const buildThaiQualityLookup = (rawRows) => {
        const headers = rawRows[0] || [];
        const ci = {
          code: findColumnIndex(headers, ['Fund Code']),
          fundSize: findColumnIndex(headers, ['Fund Size']),
          fundSizeDate: findColumnIndex(headers, ['Fund Size Date', 'Fund SizeDate']),
        };
        const get = (row, index) => index >= 0 ? String(row[index] ?? '').trim() : '';
        const lookup = new Map();
        rawRows.slice(1).forEach(row => {
          const code = get(row, ci.code).toUpperCase();
          if (!code) return;
          lookup.set(code, {
            fundSize: get(row, ci.fundSize),
            fundSizeDate: get(row, ci.fundSizeDate),
          });
        });
        return lookup;
      };

      const formatFundSizeMillion = (value) => {
        const n = parseNum(value);
        if (Number.isNaN(n)) return '';
        return Math.round(n / 1000000).toLocaleString('en-US');
      };
      const parseFundSizeMillion = (value) => {
        const n = parseNum(value);
        return Number.isNaN(n) ? NaN : Math.round(n / 1000000);
      };
      const formatPercentText = (value) => {
        const text = String(value ?? '').trim();
        if (!text) return '';
        const n = parseNum(text);
        if (Number.isNaN(n)) return text;
        return `${n.toFixed(2)}%`;
      };

      const thaiQualityLookup = buildThaiQualityLookup(thaiQualityRows);
      const secLookup = buildRawSecLookup(secRows);
      const selectedFunds = Object.values(State.selectedFunds || {})
        .filter(fund => State.selectedKeys?.has(fund.key));

      rows = selectedFunds.length
        ? selectedFunds.map(fund => {
            const code = String(fund.code || '').trim().toUpperCase();
            const quality = thaiQualityLookup.get(code) || {};
            const sec = secLookup.get(code) || {};
            return {
              fundName: fund.code || fund.name || '-',
              highlightColor: HL_COLORS?.[State.highlights?.[fund.key]]?.bg || '',
              fundSize: formatFundSizeMillion(quality.fundSize),
              fundSizeValue: parseFundSizeMillion(quality.fundSize),
              fundSizeDate: quality.fundSizeDate || '',
              recoveringPeriod: sec.recoveringPeriod || '',
              averageDuration: sec.portfolioDurationPeriod || '',
              turnoverRatio: sec.portfolioTurnoverRatio || '',
              turnoverValue: parseNum(sec.portfolioTurnoverRatio),
              ytm: formatPercentText(sec.yieldToMaturity),
              ytmValue: parseNum(sec.yieldToMaturity),
            };
          })
        : Array.from({ length: 8 }, () => ({
            fundName: '-',
            highlightColor: '',
            fundSize: '',
            fundSizeDate: '',
            recoveringPeriod: '',
            averageDuration: '',
            turnoverRatio: '',
            turnoverValue: NaN,
            ytm: '',
            ytmValue: NaN,
          }));
    } catch (e) {
      setError(area, e.message, pageKey);
      return;
    }

    const source = 'AVP Thai Fund for Quality + Data For SEC API';

    const headers = [
      'ชื่อกอง',
      'Fund Size (ลบ.)',
      'Fund Size Date',
      'Recovering Period',
      'อายุของตราสารหนี้เฉลี่ย',
      'อัตราส่วนหมุนเวียนการลงทุน',
      'Yield to Maturity',
    ];

    const sortState = State.reportSorts[pageKey] || (State.reportSorts[pageKey] = { key: '', dir: 'asc' });
    const parseThaiPeriodMonths = (value) => {
      const text = String(value || '');
      if (!text.trim()) return '';
      const year = Number((/(\d+(?:\.\d+)?)\s*ปี/.exec(text) || [])[1] || 0);
      const month = Number((/(\d+(?:\.\d+)?)\s*เดือน/.exec(text) || [])[1] || 0);
      const day = Number((/(\d+(?:\.\d+)?)\s*วัน/.exec(text) || [])[1] || 0);
      const total = (year * 12) + month + (day / 30);
      return total || text;
    };
    const sortable = (row, key) => ({
      fundName: row.fundName,
      fundSize: row.fundSizeValue,
      fundSizeDate: row.fundSizeDate,
      recoveringPeriod: parseThaiPeriodMonths(row.recoveringPeriod),
      averageDuration: parseThaiPeriodMonths(row.averageDuration),
      turnoverRatio: row.turnoverValue,
      ytm: row.ytmValue,
    })[key];
    const sortedRows = sortState.key
      ? [...rows].sort((a, b) => compareValues(sortable(a, sortState.key), sortable(b, sortState.key), sortState.dir))
      : rows;
    const fundSizeValues = rows
      .map(row => row.fundSizeValue)
      .filter(value => !Number.isNaN(value));
    const minFundSize = fundSizeValues.length ? Math.min(...fundSizeValues) : NaN;
    const maxFundSize = fundSizeValues.length ? Math.max(...fundSizeValues) : NaN;
    const fundSizeHeatStyle = (value) => {
      if (Number.isNaN(value) || Number.isNaN(minFundSize) || Number.isNaN(maxFundSize)) return '';
      const ratio = maxFundSize === minFundSize ? 1 : Math.max(0, Math.min(1, (value - minFundSize) / (maxFundSize - minFundSize)));
      const start = { r: 255, g: 255, b: 255 };
      const end = { r: 122, g: 188, b: 129 };
      const mix = (a, b) => Math.round(a + (b - a) * ratio);
      return `background:rgb(${mix(start.r, end.r)}, ${mix(start.g, end.g)}, ${mix(start.b, end.b)});color:#24364f;font-weight:700;`;
    };
    const sortHeader = (key, labelHtml) =>
      `<th class="of2-sort ${sortState.key === key ? 'is-active' : ''}" data-of2-sort="${esc(key)}">${renderSortLabel(labelHtml, sortState.key === key, sortState.dir, false)}</th>`;

    area.innerHTML = `
      ${pageToolActions(pageKey, source)}
      <div class="card report-card" id="report-card">
        <div class="of2-table-wrap">
          <table class="of2-table">
            <colgroup>
              <col class="of2-col-name">
              <col class="of2-col-size">
              <col class="of2-col-date">
              <col class="of2-col-recovering">
              <col class="of2-col-duration">
              <col class="of2-col-turnover">
              <col class="of2-col-ytm">
            </colgroup>
            <thead>
              <tr>
                ${sortHeader('fundName', 'ชื่อกอง')}
                ${sortHeader('fundSize', '<span class="of2-th-lines"><span>Fund Size</span><span>(ลบ.)</span></span>')}
                ${sortHeader('fundSizeDate', 'Fund Size Date')}
                ${sortHeader('recoveringPeriod', '<span class="of2-th-lines"><span>Recovering</span><span>Period</span></span>')}
                ${sortHeader('averageDuration', '<span class="of2-th-lines"><span>อายุของ</span><span>ตราสารหนี้เฉลี่ย</span></span>')}
                ${sortHeader('turnoverRatio', '<span class="of2-th-lines"><span>อัตราส่วน</span><span>หมุนเวียนการลงทุน</span></span>')}
                ${sortHeader('ytm', '<span class="of2-th-lines"><span>Yield to</span><span>Maturity</span></span>')}
              </tr>
            </thead>
            <tbody>
              ${sortedRows.map((row, rowIndex) => {
                const baseRowBg = rowIndex % 2 === 0 ? '#ffffff' : '#f4f6fa';
                const nameStyle = row.highlightColor
                  ? `background:${esc(row.highlightColor)}`
                  : `background:${baseRowBg}`;
                return `
                <tr>
                  <td class="of2-name" style="${nameStyle}">${esc(row.fundName)}</td>
                  <td class="of2-size ${row.fundSize ? '' : 'of2-placeholder'}" style="${fundSizeHeatStyle(row.fundSizeValue)}">${esc(row.fundSize || '-')}</td>
                  <td class="${row.fundSizeDate ? '' : 'of2-placeholder'}">${esc(row.fundSizeDate || '-')}</td>
                  <td class="${row.recoveringPeriod ? '' : 'of2-placeholder'}">${esc(row.recoveringPeriod || '-')}</td>
                  <td class="${row.averageDuration ? '' : 'of2-placeholder'}">${esc(row.averageDuration || '-')}</td>
                  <td class="${row.turnoverRatio ? '' : 'of2-placeholder'}">${esc(row.turnoverRatio || '-')}</td>
                  <td class="${row.ytm ? '' : 'of2-placeholder'}">${esc(row.ytm || '-')}</td>
                </tr>
              `;}).join('')}
            </tbody>
          </table>
        </div>
      </div>
      <style>
        .of2-table-wrap{overflow-x:auto;background:#fff}
        .of2-table{width:100%;min-width:1120px;border-collapse:collapse;table-layout:fixed;font-family:'Sarabun','THSarabunNew',sans-serif;color:#24364f}
        .of2-col-name{width:15%}.of2-col-size{width:14%}.of2-col-date{width:14%}.of2-col-recovering{width:14%}.of2-col-duration{width:17%}.of2-col-turnover{width:16%}.of2-col-ytm{width:10%}
        .of2-table th{height:54px;padding:8px 10px;background:#244a80;color:#fff;font-size:.78rem;font-weight:800;text-align:center;vertical-align:middle;border:1px solid #dbe4f0;line-height:1.2}
        .of2-sort{cursor:pointer;user-select:none}
        .of2-sort:hover{background:#315d96}
        .of2-sort.is-active{background:#f7d774;color:#4e3500;border-color:#d79a12}
        .of2-table .sort-label{display:inline-flex;align-items:center;justify-content:center;gap:5px;width:100%;line-height:1.2}
        .of2-table .sort-text{display:inline-block}
        .of2-th-lines,.of2-th-lines span{display:block}
        .of2-table td{height:42px;padding:7px 10px;border:1px solid #dbe4f0;text-align:center;font-size:.82rem;font-weight:500;line-height:1.25;vertical-align:middle;background:#fff}
        .of2-table tbody tr:nth-child(even) td{background:#f4f6fa}
        .of2-table td.of2-name{font-weight:800;color:#173566}
        .of2-table td.of2-placeholder{color:#8b98aa;font-weight:700}
        @media(max-width:760px){.of2-table{min-width:980px}.of2-table th{font-size:.74rem;height:50px}.of2-table td{font-size:.78rem;height:40px}}
      </style>`;

    $$('.of2-sort', area).forEach(el => {
      el.addEventListener('click', () => {
        toggleNamedSort(sortState, el.dataset.of2Sort);
        Pages.otherFactorsFixedIncomeTable(area);
      });
    });

    App._currentExport = null;
    App._currentTableExport = () => buildSimpleTablePayload(
      CONFIG.PAGES[pageKey]?.title || 'ปัจจัยประกอบ กองทุนตราสารหนี้',
      source,
      headers,
      sortedRows.map(row => [
        row.fundName,
        row.fundSize,
        row.fundSizeDate,
        row.recoveringPeriod,
        row.averageDuration,
        row.turnoverRatio,
        row.ytm,
      ])
    );
    App._currentClipboardExport = null;
    bindPageImageActions(area, 'report-card', 'other-factors-fixed-income-table');
  },

  async otherFactorsFixedIncomeAllocationTable(area) {
    const pageKey = 'master-placeholder-10';
    const source = 'AVP Thai Fund for Quality + AVP Master Fund ID + Data For SEC API';
    setLoading(area, 'กำลังโหลดตารางปัจจัยประกอบ กองทุนตราสารหนี้...');

    let rows;
    let editableColumns = [];
    try {
      const [selectRows, thaiQualityRows, masterRows, secRows] = await Promise.all([
        fetchCached('select-fund'),
        fetchCached('thai-annualized-v2'),
        fetchCached('master-annualized-v2'),
        fetchCached('master-placeholder-4'),
      ]);
      await Promise.all([
        loadFundOverrides(),
        loadMasterAllocations(),
        loadFixedIncomeFactorsOverrides(),
      ]);
      const catalogByKey = Object.fromEntries(
        applyFundOverrides(buildSelectedFundsCatalog(selectRows)).map(fund => [fund.key, fund])
      );
      const previousSelectedFunds = State.selectedFunds || {};
      State.selectedFunds = Object.keys(previousSelectedFunds).length
        ? {
            ...previousSelectedFunds,
            ...Object.fromEntries(
              Object.keys(previousSelectedFunds).map(key => [key, catalogByKey[key] || previousSelectedFunds[key]])
            ),
          }
        : catalogByKey;
      const headers = thaiQualityRows[0] || [];
      const ci = {
        code: findColumnIndex(headers, ['Fund Code']),
        cash: findColumnIndex(headers, ['Asset Alloc Cash % (Net)']),
        bond: findColumnIndex(headers, ['Asset Alloc Bond % (Net)']),
        aaa: findColumnIndex(headers, ['Fixd-Inc Credit Rtg - Brkdwn AAA (Calc) (Net) (FI%)']),
        aa: findColumnIndex(headers, ['Fixd-Inc Credit Rtg - Brkdwn AA (Calc) (Net) (FI%)']),
        a: findColumnIndex(headers, ['Fixd-Inc Credit Rtg - Brkdwn A (Calc) (Net) (FI%)']),
        bbb: findColumnIndex(headers, ['Fixd-Inc Credit Rtg - Brkdwn BBB (Calc) (Net) (FI%)']),
        holdings: findColumnIndex(headers, ['# of Holdings (Long)']),
        top10: findColumnIndex(headers, ['Latest % Asset in Top 10 Holdings']),
        fundSize: findColumnIndex(headers, ['Fund Size']),
        maxDd3y: findColumnIndex(headers, ['Max Drawdown 3Y']),
        sd3y: findColumnIndex(headers, ['Std Dev(Annualized) 3Y']),
        maxDd5y: findColumnIndex(headers, ['Max Drawdown 5Y']),
        sd5y: findColumnIndex(headers, ['Std Dev(Annualized) 5Y']),
      };
      const masterHeaders = masterRows[0] || [];
      const mci = {
        isin: findColumnIndex(masterHeaders, ['ISIN']),
        fundId: findColumnIndex(masterHeaders, ['FundId', 'Fund ID']),
        holdings: findColumnIndex(masterHeaders, ['# of Holdings (Long)']),
        top10: findColumnIndex(masterHeaders, ['Latest % Asset in Top 10 Holdings']),
        maxDd3y: findColumnIndex(masterHeaders, ['Max Drawdown 3Y']),
        sd3y: findColumnIndex(masterHeaders, ['Std Dev(Annualized) 3Y']),
        maxDd5y: findColumnIndex(masterHeaders, ['Max Drawdown 5Y']),
        sd5y: findColumnIndex(masterHeaders, ['Std Dev(Annualized) 5Y']),
      };
      const get = (row, index) => index >= 0 ? String(row[index] ?? '').trim() : '';
      const qualityLookup = new Map();
      thaiQualityRows.slice(1).forEach(row => {
        const code = get(row, ci.code).toUpperCase();
        if (code) qualityLookup.set(code, row);
      });
      const secLookup = buildRawSecLookup(secRows);
      const masterLookup = new Map();
      masterRows.slice(1).forEach(row => {
        [get(row, mci.isin), get(row, mci.fundId)].forEach(key => {
          const normalized = String(key || '').trim().toUpperCase();
          if (normalized && !masterLookup.has(normalized)) masterLookup.set(normalized, row);
        });
      });
      const formatTwo = (value) => {
        const n = parseNum(value);
        return Number.isNaN(n) ? '' : n.toFixed(2);
      };
      const formatFundSizeMillion = (value) => {
        const n = parseNum(value);
        if (Number.isNaN(n)) return '';
        return Math.round(n / 1000000).toLocaleString('en-US');
      };
      const formatPercentText = (value) => {
        const text = String(value ?? '').trim();
        if (!text) return '';
        const n = parseNum(text);
        if (Number.isNaN(n)) return text;
        return `${n.toFixed(2)}%`;
      };
      const masterAllocationsForFund = (fund) => {
        const saved = State.masterAllocations?.items?.[fund.key];
        const savedAllocations = Array.isArray(saved?.allocations) ? saved.allocations : [];
        const validSaved = savedAllocations.filter(item => item?.masterId || item?.masterName);
        if (validSaved.length) return validSaved;
        return fund.masterId ? [{ masterId: fund.masterId, weight: 100 }] : [];
      };
      const weightedMasterMetric = (fund, index) => {
        if (index < 0) return '';
        const allocations = masterAllocationsForFund(fund);
        let totalWeight = 0;
        let weightedValue = 0;
        allocations.forEach(item => {
          const masterKey = String(item.masterId || '').trim().toUpperCase();
          const masterRow = masterLookup.get(masterKey);
          if (!masterRow) return;
          const value = parseNum(get(masterRow, index));
          if (Number.isNaN(value)) return;
          const weight = parseFloat(item.weight);
          const effectiveWeight = Number.isNaN(weight) ? 100 : weight;
          weightedValue += value * effectiveWeight;
          totalWeight += effectiveWeight;
        });
        return totalWeight ? (weightedValue / totalWeight).toFixed(2) : '';
      };
      const preferMasterMetric = (fund, raw, masterIndex, thaiIndex) =>
        weightedMasterMetric(fund, masterIndex) || formatTwo(get(raw, thaiIndex));
      editableColumns = [
        'cash',
        'bond',
        'sector2',
        'foreign',
        'aaa',
        'aa',
        'a',
        'bbb',
        'duration',
        'ytm',
        'holdings',
        'top10',
        'fundSize',
        'maxDd3y',
        'sd3y',
        'maxDd5y',
        'sd5y',
      ];
      const applyFixedIncomeFactorsOverride = (fund, row) => {
        const override = State.fixedIncomeFactorsOverrides?.items?.[fund.key] || State.fixedIncomeFactorsOverrides?.items?.[String(fund.code || '').trim().toUpperCase()];
        const baseValues = Object.fromEntries(editableColumns.map(key => [key, row[key] || '']));
        if (!override) return { ...row, baseValues };
        const next = { ...row, baseValues, isOverride: true };
        editableColumns.forEach(key => {
          if (Object.prototype.hasOwnProperty.call(override, key)) next[key] = String(override[key] ?? '').trim();
        });
        return next;
      };
      const selectedFunds = Object.values(State.selectedFunds || {})
        .filter(fund => State.selectedKeys?.has(fund.key));
      rows = selectedFunds.length
        ? selectedFunds.map(fund => {
            const code = String(fund.code || '').trim().toUpperCase();
            const raw = qualityLookup.get(code) || [];
            const sec = secLookup.get(code) || {};
            const row = {
              key: fund.key,
              code,
              fundName: fund.code || fund.name || '-',
              highlightColor: HL_COLORS?.[State.highlights?.[fund.key]]?.bg || '',
              cash: formatTwo(get(raw, ci.cash)),
              bond: formatTwo(get(raw, ci.bond)),
              sector2: '',
              foreign: '',
              aaa: formatTwo(get(raw, ci.aaa)),
              aa: formatTwo(get(raw, ci.aa)),
              a: formatTwo(get(raw, ci.a)),
              bbb: formatTwo(get(raw, ci.bbb)),
              duration: sec.portfolioDurationPeriod || '',
              ytm: formatPercentText(sec.yieldToMaturity),
              holdings: preferMasterMetric(fund, raw, mci.holdings, ci.holdings),
              top10: preferMasterMetric(fund, raw, mci.top10, ci.top10),
              fundSize: formatFundSizeMillion(get(raw, ci.fundSize)),
              maxDd3y: preferMasterMetric(fund, raw, mci.maxDd3y, ci.maxDd3y),
              sd3y: preferMasterMetric(fund, raw, mci.sd3y, ci.sd3y),
              maxDd5y: preferMasterMetric(fund, raw, mci.maxDd5y, ci.maxDd5y),
              sd5y: preferMasterMetric(fund, raw, mci.sd5y, ci.sd5y),
            };
            return applyFixedIncomeFactorsOverride(fund, row);
          })
        : Array.from({ length: 6 }, () => ({ fundName: '-', highlightColor: '' }));
    } catch (e) {
      setError(area, e.message, pageKey);
      return;
    }

    const columns = [
      'กองไทย',
      'Asset Alloc Cash',
      'Fixd-Inc Sector -',
      'Fixd-Inc Sector -',
      'ลงทุนต่างประเทศ',
      'AAA',
      'AA',
      'A',
      'BBB',
      'อายุของตราสารหนี้เฉลี่ย',
      'YTM',
      'Number of Holdings',
      '% Asset in Top 10',
      'Fund Size (ลบ.)',
      'Max DD 3Y',
      'SD 3Y',
      'Max DD 5Y',
      'SD 5Y',
    ];
    const isEditingFixedIncomeFactors = !!State.fixedIncomeFactorsEditMode;
    const editFundDataButton = isEditingFixedIncomeFactors
      ? `
        <button class="btn btn-primary" id="of3-save-overrides" type="button" title="บันทึกข้อมูลแก้ไขของหน้านี้">
          บันทึก Override
        </button>
        <button class="btn btn-ghost" id="of3-cancel-edit" type="button" title="ยกเลิกการแก้ไข">
          ยกเลิก
        </button>`
      : `
        <button class="btn btn-ghost" id="of3-edit-fund-data" type="button" title="แก้ไขข้อมูลในตารางนี้">
          แก้ไขข้อมูลกองทุน
        </button>`;
    const renderOf3Cell = (row, key) => {
      const value = row[key] || '';
      if (!isEditingFixedIncomeFactors) {
        return `<td class="${value ? '' : 'of3-placeholder'}">${esc(value || '-')}</td>`;
      }
      return `
        <td class="of3-edit-cell">
          <input class="fund-input fund-input-editable of3-override-input" data-key="${esc(key)}" value="${esc(value)}" placeholder="-">
        </td>`;
    };

    area.innerHTML = `
      ${pageToolActions(pageKey, source, '', editFundDataButton)}
      <div class="card report-card" id="report-card">
        <div class="of3-table-wrap">
          <table class="of3-table">
            <colgroup>
              <col class="of3-col-fund">
              <col class="of3-col-small">
              <col class="of3-col-small">
              <col class="of3-col-small">
              <col class="of3-col-small">
              <col class="of3-col-rating">
              <col class="of3-col-rating">
              <col class="of3-col-rating">
              <col class="of3-col-rating">
              <col class="of3-col-duration">
              <col class="of3-col-small">
              <col class="of3-col-holdings">
              <col class="of3-col-holdings">
              <col class="of3-col-size">
              <col class="of3-col-risk">
              <col class="of3-col-risk">
              <col class="of3-col-risk">
              <col class="of3-col-risk">
            </colgroup>
            <thead>
              <tr>
                <th rowspan="2">กองไทย</th>
                <th rowspan="2"><span>Asset Alloc</span><span>Cash</span></th>
                <th rowspan="2"><span>Fixd-Inc</span><span>Sector -</span></th>
                <th rowspan="2"><span>Fixd-Inc</span><span>Sector -</span></th>
                <th rowspan="2"><span>ลงทุน</span><span>ต่างประเทศ</span></th>
                <th colspan="4" class="of3-group">Rating</th>
                <th rowspan="2"><span>อายุของตรา</span><span>สารหนี้เฉลี่ย</span></th>
                <th rowspan="2">YTM</th>
                <th rowspan="2"><span>Number of</span><span>Holdings</span></th>
                <th rowspan="2"><span>% Asset in</span><span>Top 10</span></th>
                <th rowspan="2"><span>Fund Size</span><span>(ลบ.)</span></th>
                <th colspan="4" class="of3-group">ความเสี่ยง</th>
              </tr>
              <tr>
                <th>AAA</th>
                <th>AA</th>
                <th>A</th>
                <th>BBB</th>
                <th>Max DD 3Y</th>
                <th>SD 3Y</th>
                <th>Max DD 5Y</th>
                <th>SD 5Y</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map((row, rowIndex) => {
                const baseBg = rowIndex % 2 === 0 ? '#ffffff' : '#f4f6fa';
                const nameStyle = row.highlightColor
                  ? `background:${esc(row.highlightColor)}`
                  : `background:${baseBg}`;
                return `
                  <tr>
                    <td class="of3-name" style="${nameStyle}">${esc(row.fundName)}</td>
                    ${editableColumns.map(key => renderOf3Cell(row, key)).join('')}
                  </tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
      <style>
        .of3-table-wrap{overflow-x:auto;background:#fff}
        .of3-table{width:100%;min-width:1320px;border-collapse:collapse;table-layout:fixed;font-family:'Sarabun','THSarabunNew',sans-serif;color:#24364f}
        .of3-col-fund{width:11%}.of3-col-small{width:7%}.of3-col-rating{width:6%}.of3-col-duration{width:9%}.of3-col-holdings{width:8%}.of3-col-size{width:7%}.of3-col-risk{width:6%}
        .of3-table th{height:38px;padding:6px 8px;background:#123a73;color:#fff;font-size:.74rem;font-weight:800;text-align:center;vertical-align:middle;border:1px solid #dbe4f0;line-height:1.15}
        .of3-table th span{display:block}
        .of3-table th.of3-group{background:#143e79}
        .of3-table td{height:34px;padding:6px 8px;border:1px solid #dbe4f0;text-align:center;font-size:.8rem;font-weight:500;line-height:1.2;vertical-align:middle;background:#fff}
        .of3-table tbody tr:nth-child(even) td{background:#f4f6fa}
        .of3-table td.of3-name{font-weight:800;color:#173566;text-align:left}
        .of3-table td.of3-placeholder{color:#8b98aa;font-weight:700}
        .of3-table td.of3-edit-cell{padding:3px;background:#fff8df}
        .of3-table .of3-override-input{height:28px;min-width:0;width:100%;padding:3px 5px;text-align:center;font-size:.78rem}
        @media(max-width:760px){.of3-table{min-width:1120px}.of3-table th{font-size:.7rem}.of3-table td{font-size:.76rem}}
      </style>`;

    $('#of3-edit-fund-data', area)?.addEventListener('click', () => {
      State.fixedIncomeFactorsEditMode = true;
      Pages.otherFactorsFixedIncomeAllocationTable(area);
    });
    $('#of3-cancel-edit', area)?.addEventListener('click', () => {
      State.fixedIncomeFactorsEditMode = false;
      Pages.otherFactorsFixedIncomeAllocationTable(area);
    });
    $('#of3-save-overrides', area)?.addEventListener('click', async () => {
      const items = {};
      $$('.of3-table tbody tr', area).forEach((tr, idx) => {
        const row = rows[idx];
        const code = String(row?.code || row?.key || '').trim().toUpperCase();
        if (!code) return;
        const item = { key: code, code };
        $$('.of3-override-input', tr).forEach(input => {
          const key = input.dataset.key;
          if (!key) return;
          const nextValue = input.value.trim();
          const baseValue = String(row?.baseValues?.[key] ?? '').trim();
          if (nextValue !== baseValue) item[key] = nextValue;
        });
        if (row?.isOverride || Object.keys(item).length > 2) items[code] = item;
      });
      if (!Object.keys(items).length) {
        State.fixedIncomeFactorsEditMode = false;
        toast('ไม่มีข้อมูลที่ต้องบันทึก Override', 'info');
        Pages.otherFactorsFixedIncomeAllocationTable(area);
        return;
      }
      try {
        await saveFixedIncomeFactorsOverrides(items);
        State.fixedIncomeFactorsEditMode = false;
        toast('บันทึก Override ของปัจจัยประกอบ กองทุนตราสารหนี้แล้ว', 'success');
        Pages.otherFactorsFixedIncomeAllocationTable(area);
      } catch (err) {
        toast(err.message || 'บันทึก Override ไม่สำเร็จ', 'error');
      }
    });

    App._currentExport = null;
    App._currentTableExport = () => buildSimpleTablePayload(
      CONFIG.PAGES[pageKey]?.title || 'ปัจจัยประกอบ กองทุนตราสารหนี้',
      source,
      columns,
      rows.map(row => [
        row.fundName,
        row.cash || '-',
        row.bond || '-',
        row.sector2 || '-',
        row.foreign || '-',
        row.aaa || '-',
        row.aa || '-',
        row.a || '-',
        row.bbb || '-',
        row.duration || '-',
        row.ytm || '-',
        row.holdings || '-',
        row.top10 || '-',
        row.fundSize || '-',
        row.maxDd3y || '-',
        row.sd3y || '-',
        row.maxDd5y || '-',
        row.sd5y || '-',
      ])
    );
    App._currentClipboardExport = null;
    bindPageImageActions(area, 'report-card', 'other-factors-fixed-income-allocation-table');
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
	    let draftApiStatus = '';

    function loadLocalDrafts() {
      try { return JSON.parse(localStorage.getItem(DRAFTS_KEY) || '[]'); }
      catch { return []; }
    }

	    function saveLocalDrafts(arr) {
	      localStorage.setItem(DRAFTS_KEY, JSON.stringify(arr));
	    }

	    function hasDraftApi() {
	      return Boolean(String(DRAFT_API_WEB_APP_URL || '').trim());
	    }

	    function shouldUseLocalDraftProxy() {
	      return ['localhost', '127.0.0.1', '::1'].includes(window.location.hostname);
	    }

	    function draftApiJsonp(params) {
	      return new Promise((resolve, reject) => {
	        const callbackName = `__draftApiCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
	        const script = document.createElement('script');
	        const cleanup = () => {
	          delete window[callbackName];
	          script.remove();
	        };
	        const timer = setTimeout(() => {
	          cleanup();
	          reject(new Error('Draft API request timeout'));
	        }, 30000);
	        window[callbackName] = (data) => {
	          clearTimeout(timer);
	          cleanup();
	          resolve(data || {});
	        };
	        const url = new URL(DRAFT_API_WEB_APP_URL);
	        Object.entries({
	          key: DRAFT_API_SECRET_KEY,
	          callback: callbackName,
	          ...params,
	        }).forEach(([key, value]) => {
	          if (value !== undefined && value !== null && value !== '') {
	            url.searchParams.set(key, value);
	          }
	        });
	        script.onerror = () => {
	          clearTimeout(timer);
	          cleanup();
	          reject(new Error('Draft API script failed to load'));
	        };
	        script.src = url.toString();
	        document.head.appendChild(script);
	      });
	    }

	    async function draftApiFetch(params) {
	      const url = new URL(DRAFT_API_WEB_APP_URL);
	      Object.entries({
	        key: DRAFT_API_SECRET_KEY,
	        ...params,
	      }).forEach(([key, value]) => {
	        if (value !== undefined && value !== null && value !== '') {
	          url.searchParams.set(key, value);
	        }
	      });
	      const res = await fetch(url.toString(), { cache: 'no-store', redirect: 'follow' });
	      const text = await res.text();
	      if (!res.ok) throw new Error(`Draft API HTTP ${res.status}`);
	      return JSON.parse(text);
	    }

	    async function draftApiCompressedPayload(payload) {
	      const text = JSON.stringify(payload || {});
	      if (!('CompressionStream' in window)) return { payload: text };
	      const stream = new Blob([text], { type: 'application/json' }).stream().pipeThrough(new CompressionStream('gzip'));
	      const buffer = await new Response(stream).arrayBuffer();
	      let binary = '';
	      const bytes = new Uint8Array(buffer);
	      for (let i = 0; i < bytes.length; i += 0x8000) {
	        binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
	      }
	      return {
	        payloadGzipB64: btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, ''),
	      };
	    }

	    async function draftApiDirectRequest(action, payload = {}) {
	      const params = { action };
	      if (action === 'list') {
	        if (payload?.quarter) params.quarter = payload.quarter;
	      } else if (action === 'delete') {
	        params.id = payload?.id || '';
	        if (payload?.quarter) params.quarter = payload.quarter;
	      } else {
	        Object.assign(params, await draftApiCompressedPayload({
	          action,
	          key: DRAFT_API_SECRET_KEY,
	          draft: payload?.draft || payload || {},
	        }));
	      }
	      let data;
	      try {
	        data = await draftApiFetch(params);
	      } catch {
	        data = await draftApiJsonp(params);
	      }
	      if (data.ok === false) {
	        throw new Error(data.error || `Draft API ${action} failed`);
	      }
	      return data;
	    }

	    async function draftApiRequest(action, payload = {}) {
	      if (!shouldUseLocalDraftProxy()) {
	        return await draftApiDirectRequest(action, payload);
	      }
	      let res;
	      if (action === 'list') {
	        const quarter = payload?.quarter ? `?quarter=${encodeURIComponent(payload.quarter)}` : '';
	        res = await fetch(`/api/draft-drive${quarter}`, { cache: 'no-store' });
	      } else if (action === 'delete') {
	        const quarter = payload?.quarter ? `?quarter=${encodeURIComponent(payload.quarter)}` : '';
	        res = await fetch(`/api/draft-drive/${encodeURIComponent(payload.id || '')}${quarter}`, { method: 'DELETE' });
	      } else {
	        res = await fetch('/api/draft-drive', {
	          method: 'POST',
	          headers: { 'Content-Type': 'application/json' },
	          body: JSON.stringify(payload?.draft || payload || {}),
	        });
	      }
	      const data = await res.json().catch(() => ({}));
	      if (!res.ok || data.ok === false) {
	        throw new Error(data.error || `Draft API ${action} failed (${res.status})`);
	      }
	      return data;
	    }
	
	    async function loadDrafts() {
	      try {
	        if (hasDraftApi()) {
	          const data = await draftApiRequest('list');
	          storageMode = 'drive';
	          draftApiStatus = '';
	          const driveDrafts = Array.isArray(data.drafts) ? data.drafts : [];
	          const localDrafts = loadLocalDrafts();
	          const driveIds = new Set(driveDrafts.map(d => String(d.id || '')));
	          const draftsToMigrate = localDrafts.filter(d => d?.id && !driveIds.has(String(d.id)));
	          for (const draft of draftsToMigrate) {
	            try {
	              await draftApiRequest('save', {
	                draft: {
	                  ...draft,
	                  highlights: draft.highlights || {},
	                  filters: draft.filters || {},
	                  migratedFromLocalStorage: true,
	                },
	              });
	            } catch {
	              /* keep local copy if migration fails */
	            }
	          }
	          if (draftsToMigrate.length) {
	            localStorage.removeItem(DRAFTS_KEY);
	            const migratedData = await draftApiRequest('list');
	            return Array.isArray(migratedData.drafts) ? migratedData.drafts : driveDrafts;
	          }
	          return driveDrafts;
	        }
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
	      } catch (err) {
	        draftApiStatus = hasDraftApi() ? `Draft API ใช้งานไม่ได้: ${err.message || err}` : '';
	        storageMode = 'local';
	        return loadLocalDrafts();
	      }
	    }

	    async function saveDraft(draft) {
	      if (storageMode === 'drive') {
	        return await draftApiRequest('save', { draft });
	      }
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
	      if (storageMode === 'drive' && draft?.id) {
	        return await draftApiRequest('delete', {
	          id: draft.id,
	          quarter: draftQuarterOf(draft),
	        });
	      }
	      if (storageMode === 'file' && draft?.id) {
        const res = await fetch(`/api/drafts/${encodeURIComponent(draft.id)}`, { method: 'DELETE' });
        if (!res.ok) throw new Error(`ลบไฟล์ไม่สำเร็จ (${res.status})`);
        return res.json();
      }
      currentDrafts.splice(idx, 1);
      saveLocalDrafts(currentDrafts);
      return { ok: true };
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
	        ? 'บันทึกเป็นไฟล์กลางใน Drafts/ และ sync ไป Google Drive'
	        : storageMode === 'drive'
	          ? 'บันทึกเป็น JSON บน Google Drive ผ่าน Apps Script'
	        : 'บันทึกใน Browser นี้ชั่วคราว';
	      const storageWarning = draftApiStatus
	        ? `<div class="notes-storage-mode notes-storage-warning">${esc(draftApiStatus)}</div>`
	        : '';

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
	            ${storageWarning}
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
	          .notes-storage-warning{color:#b45309;background:#fffbeb;border-color:#fcd34d;margin-top:-8px}
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
          const result = await saveDraft(draft);
          if (result?.driveUploaded) {
            toast('บันทึกดราฟและ sync ไป Google Drive แล้ว', 'success', 5000);
          } else if (result?.warning) {
            toast(result.warning, 'warning', 7000);
          } else {
            toast('บันทึกดราฟแล้ว', 'success', 4000);
          }
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
            const result = await saveDraft(buildDraftPayload(existingDraft));
            currentDrafts = await loadDrafts();
            renderPage();
            if (result?.driveUploaded) {
              toast('บันทึกทับดราฟและ sync ไป Google Drive แล้ว', 'success', 5000);
            } else if (result?.warning) {
              toast(result.warning, 'warning', 7000);
            } else {
              toast('บันทึกทับดราฟเรียบร้อยแล้ว', 'success');
            }
          } catch (err) {
            toast(`บันทึกทับไม่สำเร็จ: ${err.message || err}`, 'error', 5000);
          }
        }
        if (delBtn) {
          if (!confirm('ลบบันทึกนี้?')) return;
          const idx = parseInt(delBtn.dataset.idx);
          const deletingDraft = currentDrafts[idx];
          try {
            const result = await deleteDraft(deletingDraft, idx);
            if (String(overwriteDraftId) === String(deletingDraft?.id || '')) overwriteDraftId = '';
            currentDrafts = await loadDrafts();
            renderPage();
            if (result?.warning) {
              toast(result.warning, 'warning', 7000);
            } else {
              toast('ลบดราฟแล้ว', 'success', 4000);
            }
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

  async incomeFund(area, pageKey = 'income-fund-1') {
    const source = 'Fund Key Performance AVP';
    setLoading(area, 'กำลังโหลดรายชื่อ Income Fund...');

    let rawRows;
    try {
      rawRows = await fetchIncomeFundUniverseRows(pageKey);
      await Promise.all([
        loadFundOverrides(),
        loadMasterAllocations(true),
        loadIncomeFundSecMetadata(),
        loadIncomeFundSelection(true),
      ]);
    } catch (e) {
      setError(area, e.message, pageKey);
      return;
    }

    const headers = rawRows[0] || [];
    const CI = {
      CATEGORY: findColumnIndex(headers, ['AVP® Category', 'AVP®  Category', 'AVP Category']),
    };
    const allFunds = applyFundOverrides(buildSelectedFundsCatalog(rawRows))
      .filter(fund => fund.dividend === 'Dividend' || fund.dividend === 'Redemption');

    State.selectedFunds = {
      ...State.selectedFunds,
      ...Object.fromEntries(allFunds.map(f => [f.key, f])),
    };

    const categories = [...new Set(allFunds.map(f => f.category).filter(Boolean))].sort();
    const filters = State.incomeFundFilters;
    const sortState = State.incomeFundSort;
    const optLocal = (val, label, cur) =>
      `<option value="${esc(val)}" ${cur === val ? 'selected' : ''}>${esc(label)}</option>`;
    const policyLabel = value => value === 'Redemption' ? 'Auto Redeem' : value;
    const getSavedMasterAllocations = (fund) => {
      const saved = State.masterAllocations?.items?.[fund.key];
      const allocations = Array.isArray(saved?.allocations) ? saved.allocations : [];
      return allocations.filter(item => item?.masterId || item?.masterName);
    };
    const masterMappingDisplay = (fund) => {
      const allocations = getSavedMasterAllocations(fund);
      if (!allocations.length) {
        return {
          isMapped: false,
          masterId: fund.masterId,
          masterName: fund.masterName,
          searchText: `${fund.masterId || ''} ${fund.masterName || ''}`,
        };
      }
      const masterId = allocations.map(item => item.masterId).filter(Boolean).join(', ') || '-';
      const masterName = allocations.map(item => {
        const name = item.masterName || item.masterId || '-';
        const weight = parseFloat(item.weight);
        return Number.isNaN(weight) ? name : `${name} (${weight.toFixed(2)}%)`;
      }).join(' + ');
      return {
        isMapped: true,
        masterId,
        masterName,
        searchText: allocations.map(item => `${item.masterId || ''} ${item.masterName || ''}`).join(' '),
      };
    };

    const render = (goPage = 1, opts = {}) => {
      const preserveScroll = !!opts.preserveScroll;
      const prevScrollTop = preserveScroll ? area.scrollTop : 0;
      State.tablePage = goPage;
      State.pageSize = filters.pageSize || SELECT_FUND_DEFAULT_PAGE_SIZE;

      const passesCurrentFilters = (f) => {
        const mappingText = masterMappingDisplay(f).searchText.toLowerCase();
        const policyText = policyLabel(f.dividend).toLowerCase();
        const secMeta = incomeFundSecMetadata(f.code);
        const secText = `${secMeta.projId} ${secMeta.fundClassName} ${secMeta.projAbbrName}`.toLowerCase();
        if (filters.query && !f.code.toLowerCase().includes(filters.query) && !f.name.toLowerCase().includes(filters.query) && !f.masterName.toLowerCase().includes(filters.query) && !mappingText.includes(filters.query) && !policyText.includes(filters.query) && !secText.includes(filters.query)) return false;
        if (filters.category && f.category !== filters.category) return false;
        if (filters.type && f.type !== filters.type) return false;
        if (filters.style && f.style !== filters.style) return false;
        if (filters.dividend && f.dividend !== filters.dividend) return false;
        return true;
      };
      let visible = allFunds.filter(passesCurrentFilters);
      if (sortState.key === 'focusGroup') {
        const visibleKeys = new Set(visible.map(f => f.key));
        const selectedOutsideFilters = allFunds.filter(f =>
          State.incomeFundSelectedKeys.has(f.key) && !visibleKeys.has(f.key)
        );
        visible = [...selectedOutsideFilters, ...visible];
      }

      if (sortState.key) {
        const sortableValue = (fund, key) => {
          if (key === 'highlight') {
            const idx = State.highlights[fund.key];
            return idx === undefined ? '' : HL_COLORS[idx]?.name || '';
          }
          if (key === 'masterId' || key === 'masterName') {
            return masterMappingDisplay(fund)[key];
          }
          if (key === 'dividend') return policyLabel(fund.dividend);
          return fund[key];
        };
        visible = [...visible].sort((a, b) => {
          if (sortState.key === 'focusGroup') {
            const av = State.incomeFundSelectedKeys.has(a.key) ? 1 : 0;
            const bv = State.incomeFundSelectedKeys.has(b.key) ? 1 : 0;
            if (av !== bv) return sortState.dir === 'desc' ? bv - av : av - bv;
            return compareValues(a.code, b.code, 'asc');
          }
          return compareValues(sortableValue(a, sortState.key), sortableValue(b, sortState.key), sortState.dir);
        });
      }

      const total = visible.length;
      const totalPages = Math.max(1, Math.ceil(total / State.pageSize));
      const pg = Math.min(Math.max(1, State.tablePage), totalPages);
      const si = (pg - 1) * State.pageSize;
      const pageData = visible.slice(si, si + State.pageSize);
      const allVisibleChecked = pageData.length > 0 && pageData.every(f => State.incomeFundSelectedKeys.has(f.key));
      const tRows = pageData.map(f => {
        const isSelected = State.incomeFundSelectedKeys.has(f.key);
        const mapping = masterMappingDisplay(f);
        return `
          <tr data-key="${esc(f.key)}" class="${isSelected ? 'row-selected' : ''}">
            <td class="td-check">
              <input type="checkbox" class="row-chk" data-key="${esc(f.key)}" ${isSelected ? 'checked' : ''}>
            </td>
            <td class="td-code">${esc(f.code)}</td>
            <td>${esc(f.name || '-')}</td>
            <td class="td-hl">${buildHighlightSelect(f.key, State.highlights[f.key])}</td>
            <td>${esc(policyLabel(f.dividend))}</td>
            <td class="td-isin">${esc(incomeFundSecMetadata(f.code).projId)}</td>
            <td class="td-isin">${esc(incomeFundSecMetadata(f.code).fundClassName)}</td>
            <td class="td-isin">${esc(incomeFundSecMetadata(f.code).projAbbrName)}</td>
            <td>${esc(f.style)}</td>
            <td class="td-isin">${esc(mapping.masterId)}</td>
            <td>${esc(mapping.masterName)}${mapping.isMapped ? ' <span class="badge badge-success fund-override-mini">Master Mapping</span>' : ''}${f.isOverride ? ' <span class="badge badge-accent fund-override-mini">แก้ไขแล้ว</span>' : ''}</td>
          </tr>`;
      }).join('');

      area.innerHTML = `
        <div class="card sf-card">
          <div class="sf-filterbar">
            <div class="sf-search">
              <span class="s-icon">${searchIcon()}</span>
              <input class="search-input" id="income-q" type="text"
                placeholder="ชื่อกองทุน / Fund Code..." value="${esc(filters.query)}" autocomplete="off">
            </div>
            <div class="sf-drop">
              <div class="sf-droplabel">หมวดหมู่ AVP</div>
              <select class="sf-select" id="income-category">
                ${optLocal('', 'ทั้งหมด', filters.category)}
                ${categories.map(g => optLocal(g, g, filters.category)).join('')}
              </select>
            </div>
            <div class="sf-drop">
              <div class="sf-droplabel">ประเภทกองทุน</div>
              <select class="sf-select" id="income-type">
                ${['','General','SSF','RMF','LTF','TESGX'].map((v,i) => optLocal(v, i===0?'ทั้งหมด':v, filters.type)).join('')}
              </select>
            </div>
            <div class="sf-drop">
              <div class="sf-droplabel">STYLE</div>
              <select class="sf-select" id="income-style">
                ${['','Active','Passive'].map((v,i) => optLocal(v, i===0?'ทั้งหมด':v, filters.style)).join('')}
              </select>
            </div>
            <div class="sf-drop">
              <div class="sf-droplabel">นโยบาย</div>
              <select class="sf-select" id="income-div">
                ${optLocal('', 'ทั้งหมด', filters.dividend)}
                ${optLocal('Dividend', 'Dividend', filters.dividend)}
                ${optLocal('Redemption', 'Auto Redeem', filters.dividend)}
              </select>
            </div>
            <button class="btn btn-ghost btn-sm" id="income-reset">↺ รีเซ็ต</button>
          </div>
          <div class="sf-meta">
            <span class="row-count-badge">${total.toLocaleString()} รายการ</span>
            <span class="row-count-badge is-info" id="income-selected-count">เลือกแล้ว ${State.incomeFundSelectedKeys.size.toLocaleString()} กองทุน</span>
            ${sourceBadgeHtml(pageKey, source)}
            ${getPageDataSourceBadge(pageKey) ? `<span class="badge badge-data-origin">${esc(getPageDataSourceBadge(pageKey))}</span>` : ''}
            <span class="badge badge-success">Dividend / Auto Redeem เท่านั้น</span>
            <span class="badge ${State.incomeFundSelectionSaving ? 'badge-warning' : 'badge-success'}" id="income-save-status">
              ${State.incomeFundSelectionSaving ? 'กำลังบันทึก selection...' : 'Selection synced'}
            </span>
            ${CI.CATEGORY === -1 ? `<span class="badge badge-warning">ยังไม่พบคอลัมน์ AVP® Category</span>` : ''}
          </div>
          <div class="table-wrapper">
            <table>
              <thead><tr>
                <th class="th-check sf-sort ${sortState.key === 'focusGroup' ? 'is-active' : ''}" data-sort-key="focusGroup">
                  <span style="display:inline-flex;align-items:center;gap:8px;white-space:nowrap">
                    <input type="checkbox" id="income-chk-all" title="เลือกทั้งหมดที่แสดง" ${allVisibleChecked ? 'checked' : ''} ${pageData.length === 0 ? 'disabled' : ''}>
                    ${renderSortLabel('Focus Group', sortState.key === 'focusGroup', sortState.dir)}
                  </span>
                </th>
                <th class="sf-sort ${sortState.key === 'code' ? 'is-active' : ''}" data-sort-key="code">${renderSortLabel('Fund Code', sortState.key === 'code', sortState.dir)}</th>
                <th class="sf-sort ${sortState.key === 'name' ? 'is-active' : ''}" data-sort-key="name">${renderSortLabel('ชื่อกองทุน', sortState.key === 'name', sortState.dir)}</th>
                <th class="sf-sort ${sortState.key === 'highlight' ? 'is-active' : ''}" data-sort-key="highlight">${renderSortLabel('Highlight', sortState.key === 'highlight', sortState.dir)}</th>
                <th class="sf-sort ${sortState.key === 'dividend' ? 'is-active' : ''}" data-sort-key="dividend">${renderSortLabel('นโยบาย', sortState.key === 'dividend', sortState.dir)}</th>
                <th>proj_id</th>
                <th>fund_class_name</th>
                <th>proj_abbr_name</th>
                <th class="sf-sort ${sortState.key === 'style' ? 'is-active' : ''}" data-sort-key="style">${renderSortLabel('Style', sortState.key === 'style', sortState.dir)}</th>
                <th class="sf-sort ${sortState.key === 'masterId' ? 'is-active' : ''}" data-sort-key="masterId">${renderSortLabel('ISIN', sortState.key === 'masterId', sortState.dir)}</th>
                <th class="sf-sort ${sortState.key === 'masterName' ? 'is-active' : ''}" data-sort-key="masterName">${renderSortLabel('Master Fund', sortState.key === 'masterName', sortState.dir)}</th>
              </tr></thead>
              <tbody>${tRows}</tbody>
            </table>
          </div>
          ${totalPages > 1 ? `
          <div class="pagination-bar">
            <label class="page-size-wrap">แถวต่อหน้า :
              <select class="page-size-select" id="income-page-size">
                ${SELECT_FUND_PAGE_SIZE_OPTIONS.map(size => `<option value="${size}" ${size === State.pageSize ? 'selected' : ''}>${size}</option>`).join('')}
              </select>
            </label>
            <button class="btn btn-ghost btn-sm" id="pg-prev" ${pg<=1?'disabled':''}>← ก่อนหน้า</button>
            <span class="pg-info">หน้า ${pg} / ${totalPages} &nbsp;(${si+1}–${Math.min(si+State.pageSize,total)} จาก ${total.toLocaleString()})</span>
            <button class="btn btn-ghost btn-sm" id="pg-next" ${pg>=totalPages?'disabled':''}>ถัดไป →</button>
          </div>` : ''}
        </div>`;

      const qEl = $('#income-q', area);
      let timer;
      qEl?.addEventListener('input', () => {
        clearTimeout(timer);
        qEl.value = qEl.value.toUpperCase();
        timer = setTimeout(() => {
          filters.query = qEl.value.trim().toLowerCase();
          render(1);
        }, 280);
      });
      $('#income-category', area)?.addEventListener('change', e => { filters.category = e.target.value; render(1); });
      $('#income-type', area)?.addEventListener('change', e => { filters.type = e.target.value; render(1); });
      $('#income-style', area)?.addEventListener('change', e => { filters.style = e.target.value; render(1); });
      $('#income-div', area)?.addEventListener('change', e => { filters.dividend = e.target.value; render(1); });
      $$('.sf-sort', area).forEach(el => {
        el.addEventListener('click', () => {
          if (el.dataset.sortKey === 'focusGroup' && sortState.key !== 'focusGroup') {
            sortState.key = 'focusGroup';
            sortState.dir = 'desc';
          } else {
            toggleNamedSort(sortState, el.dataset.sortKey);
          }
          render(1);
        });
      });

      $('#income-reset', area)?.addEventListener('click', () => {
        Object.assign(filters, {
          category: '',
          type: '',
          style: '',
          dividend: '',
          query: '',
          pageSize: SELECT_FUND_DEFAULT_PAGE_SIZE,
        });
        Object.assign(sortState, { key: '', dir: 'asc' });
        State.tablePage = 1;
        State.pageSize = SELECT_FUND_DEFAULT_PAGE_SIZE;
        render(1);
      });

      const syncSelectionUi = () => {
        const selectedCountEl = $('#income-selected-count', area);
        if (selectedCountEl) selectedCountEl.textContent = `เลือกแล้ว ${State.incomeFundSelectedKeys.size.toLocaleString()} กองทุน`;
        const saveStatusEl = $('#income-save-status', area);
        if (saveStatusEl) {
          saveStatusEl.textContent = State.incomeFundSelectionSaving ? 'กำลังบันทึก selection...' : 'Selection synced';
          saveStatusEl.className = `badge ${State.incomeFundSelectionSaving ? 'badge-warning' : 'badge-success'}`;
        }
        const rowCheckboxes = $$('.row-chk', area);
        const checkedCount = rowCheckboxes.filter(cb => cb.checked).length;
        const selectAllCheckbox = $('#income-chk-all', area);
        if (selectAllCheckbox) {
          selectAllCheckbox.checked = rowCheckboxes.length > 0 && checkedCount === rowCheckboxes.length;
          selectAllCheckbox.indeterminate = checkedCount > 0 && checkedCount < rowCheckboxes.length;
        }
      };

      $('#income-chk-all', area)?.addEventListener('change', e => {
        $$('.row-chk', area).forEach(cb => {
          const key = cb.dataset.key;
          cb.checked = e.target.checked;
          if (e.target.checked) State.incomeFundSelectedKeys.add(key);
          else State.incomeFundSelectedKeys.delete(key);
          cb.closest('tr')?.classList.toggle('row-selected', e.target.checked);
        });
        syncSelectionUi();
        debounceIncomeFundSelectionSave(syncSelectionUi);
      });
      $('#income-chk-all', area)?.addEventListener('click', e => e.stopPropagation());
      $$('.row-chk', area).forEach(el => {
        el.addEventListener('change', () => {
          const key = el.dataset.key;
          if (el.checked) State.incomeFundSelectedKeys.add(key);
          else State.incomeFundSelectedKeys.delete(key);
          el.closest('tr')?.classList.toggle('row-selected', el.checked);
          syncSelectionUi();
          debounceIncomeFundSelectionSave(syncSelectionUi);
        });
      });
      $$('.hl-select', area).forEach(el => {
        el.addEventListener('change', () => {
          const fund = el.dataset.fund;
          const rawValue = el.value;
          if (rawValue === '') delete State.highlights[fund];
          else State.highlights[fund] = parseInt(rawValue, 10);
          render(pg, { preserveScroll: true });
        });
      });
      $('#pg-prev', area)?.addEventListener('click', () => render(pg - 1));
      $('#pg-next', area)?.addEventListener('click', () => render(pg + 1));
      $('#income-page-size', area)?.addEventListener('change', e => {
        State.pageSize = parseInt(e.target.value, 10) || SELECT_FUND_DEFAULT_PAGE_SIZE;
        filters.pageSize = State.pageSize;
        render(1);
      });

      syncSelectionUi();
      if (preserveScroll) area.scrollTop = prevScrollTop;
    };

    render();
    App._currentExport = () => {
      const selectedRows = allFunds
        .filter(f => State.incomeFundSelectedKeys.has(f.key))
        .map(f => {
          const mapping = masterMappingDisplay(f);
          return [
            f.code,
            f.name,
            incomeFundSecMetadata(f.code).projId,
            incomeFundSecMetadata(f.code).fundClassName,
            incomeFundSecMetadata(f.code).projAbbrName,
            f.category,
            f.type,
            policyLabel(f.dividend),
            f.style,
            mapping.masterId,
            mapping.masterName,
            State.highlights[f.key] !== undefined ? HL_COLORS[State.highlights[f.key]].name : '',
          ];
        });
      if (!selectedRows.length) {
        toast('ยังไม่ได้เลือกกองทุนสำหรับ Export', 'warning');
        return;
      }
      exportExcel([
        ['Fund Code', 'Name', 'proj_id', 'fund_class_name', 'proj_abbr_name', 'AVP Category', 'Type', 'Policy', 'Style', 'Master Fund ID', 'Master Fund Name', 'Highlight Color'],
        ...selectedRows,
      ], 'income-funds');
    };
    App._currentTableExport = null;
    App._currentClipboardExport = null;
  },

  async incomeFundFocus(area, pageKey = 'income-fund-2') {
    setLoading(area, 'กำลังโหลด Focus Group จาก Income Fund...');

    let rawRows;
    let dividendDb = { dividend_history: [] };
    let dividendDbWarning = '';
    try {
      rawRows = await fetchIncomeFundUniverseRows(pageKey);
      await Promise.all([
        loadFundOverrides(),
        loadIncomeFundSecMetadata(),
        loadIncomeFundSelection(true),
      ]);
    } catch (e) {
      setError(area, e.message, pageKey);
      return;
    }

    try {
      dividendDb = await loadIncomeFundDividendDatabase(true);
    } catch (e) {
      dividendDbWarning = e.message || String(e);
    }

    const allFunds = applyFundOverrides(buildSelectedFundsCatalog(rawRows))
      .filter(fund => fund.dividend === 'Dividend' || fund.dividend === 'Redemption');
    const selectedFunds = allFunds.filter(fund => State.incomeFundSelectedKeys.has(fund.key));
    const dividendIndex = buildIncomeDividendIndex(dividendDb);
    const filters = State.incomeFund2Filters;
    const sortState = State.incomeFund2Sort;
    let dividendSyncRunning = false;

    const rows = selectedFunds.map(fund => {
      const historyRows = dividendIndex.get(fund.key) || dividendIndex.get(normalizeFundKey(fund.code)) || [];
      const summary = summarizeIncomeDividendRows(historyRows);
      const secMeta = incomeFundSecMetadata(fund.code);
      return {
        key: fund.key,
        code: fund.code,
        name: fund.name,
        policy: incomeFundPolicyLabel(fund.dividend),
        category: fund.category,
        type: fund.type,
        style: fund.style,
        projId: secMeta.projId,
        fundClassName: secMeta.fundClassName,
        projAbbrName: secMeta.projAbbrName,
        ...summary,
        latestDate: summary.latest?.dividend_date || '',
        latestValue: summary.latest?.dividend_value || 0,
      };
    });

    const render = () => {
      let visible = rows.filter(row => {
        const q = String(filters.query || '').trim().toLowerCase();
        const haystack = `${row.code} ${row.name} ${row.policy} ${row.projId} ${row.sources}`.toLowerCase();
        if (q && !haystack.includes(q)) return false;
        if (filters.source && !row.sources.toLowerCase().includes(filters.source.toLowerCase())) return false;
        return true;
      });

      if (sortState.key) {
        visible = [...visible].sort((a, b) => compareValues(a[sortState.key], b[sortState.key], sortState.dir));
      }

      const totalHistoryRows = rows.reduce((sum, row) => sum + row.count, 0);
      const fundsWithHistory = rows.filter(row => row.count > 0).length;
      const latestOverall = rows.map(row => row.latestDate).filter(Boolean).sort().pop() || '-';
      const sourceCounts = {
        SEC: rows.filter(row => row.sources.includes('SEC')).length,
        Finnomena: rows.filter(row => row.sources.includes('Finnomena')).length,
      };
      const opt = (value, label) => `<option value="${esc(value)}" ${filters.source === value ? 'selected' : ''}>${esc(label)}</option>`;

      const tableRows = visible.map(row => {
        const recent = row.history.slice(0, 3).map(item =>
          `<span class="badge badge-data-origin">${esc(item.dividend_date || '-')} · ${esc(formatDividendNumber(item.dividend_value))}</span>`
        ).join(' ');
        return `
          <tr>
            <td class="td-code">
              <button class="link-button income-focus-detail" type="button" data-key="${esc(row.key)}">${esc(row.code)}</button>
            </td>
            <td>
              <button class="link-button income-focus-detail" type="button" data-key="${esc(row.key)}">${esc(row.name || '-')}</button>
            </td>
            <td>${esc(row.policy || '-')}</td>
            <td class="td-isin">${esc(row.projId || '-')}</td>
            <td class="report-num">${esc(row.latestDate || '-')}</td>
            <td class="report-num">${esc(formatDividendNumber(row.latestValue))}</td>
            <td class="report-num">${row.count.toLocaleString()}</td>
            <td class="report-num">${esc(formatDividendNumber(row.total))}</td>
            <td class="report-num">${esc(formatDividendNumber(row.average))}</td>
            <td>${esc(row.sources)}</td>
            <td>${recent || '<span class="muted">ยังไม่มีประวัติ</span>'}</td>
          </tr>`;
      }).join('');

      area.innerHTML = `
        <div class="card sf-card">
          <div class="sf-filterbar">
            <div class="sf-search">
              <span class="s-icon">${searchIcon()}</span>
              <input class="search-input" id="income-focus-q" type="text"
                placeholder="ค้นหา Focus Group..." value="${esc(filters.query)}" autocomplete="off">
            </div>
            <div class="sf-drop">
              <div class="sf-droplabel">แหล่งข้อมูล</div>
              <select class="sf-select" id="income-focus-source">
                ${opt('', 'ทั้งหมด')}
                ${opt('SEC', 'SEC')}
                ${opt('Finnomena', 'Finnomena')}
              </select>
            </div>
            <div class="sf-search" style="max-width:360px">
              <input class="search-input" id="income-focus-github-token" type="password"
                placeholder="GitHub token สำหรับรัน Actions" value="${esc(State.incomeFundDividendGithubToken)}" autocomplete="off">
            </div>
            <button class="btn btn-ghost btn-sm" id="income-focus-sync-dividend"
              title="เรียก Apps Script เพื่อ trigger GitHub Actions/Python และอัปเดตไฟล์ Dividend History Database">
              ดึงข้อมูลปันผล
            </button>
            <button class="btn btn-ghost btn-sm" id="income-focus-refresh">รีเฟรช Selection</button>
          </div>
          <div class="sf-meta">
            <span class="row-count-badge is-info">Focus Group ${rows.length.toLocaleString()} กองทุน</span>
            <span class="row-count-badge">${fundsWithHistory.toLocaleString()} กองมีประวัติปันผล</span>
            <span class="row-count-badge">${totalHistoryRows.toLocaleString()} รายการหลัง dedupe</span>
            <span class="badge badge-data-origin">ล่าสุด ${esc(latestOverall)}</span>
            <span class="badge badge-success">SEC ${sourceCounts.SEC.toLocaleString()}</span>
            <span class="badge badge-success">Finnomena ${sourceCounts.Finnomena.toLocaleString()}</span>
            ${getPageDataSourceBadge(pageKey) ? `<span class="badge badge-data-origin">${esc(getPageDataSourceBadge(pageKey))}</span>` : ''}
            ${dividendDbWarning ? `<span class="badge badge-warning">${esc(dividendDbWarning)}</span>` : ''}
          </div>
          ${rows.length ? `
            <div class="table-wrapper">
              <table>
                <thead><tr>
                  <th class="sf-sort ${sortState.key === 'code' ? 'is-active' : ''}" data-sort-key="code">${renderSortLabel('Fund Code', sortState.key === 'code', sortState.dir)}</th>
                  <th class="sf-sort ${sortState.key === 'name' ? 'is-active' : ''}" data-sort-key="name">${renderSortLabel('ชื่อกองทุน', sortState.key === 'name', sortState.dir)}</th>
                  <th class="sf-sort ${sortState.key === 'policy' ? 'is-active' : ''}" data-sort-key="policy">${renderSortLabel('Policy', sortState.key === 'policy', sortState.dir)}</th>
                  <th>proj_id</th>
                  <th class="sf-sort ${sortState.key === 'latestDate' ? 'is-active' : ''}" data-sort-key="latestDate">${renderSortLabel('วันที่ล่าสุด', sortState.key === 'latestDate', sortState.dir)}</th>
                  <th class="sf-sort ${sortState.key === 'latestValue' ? 'is-active' : ''}" data-sort-key="latestValue">${renderSortLabel('ปันผลล่าสุด', sortState.key === 'latestValue', sortState.dir)}</th>
                  <th class="sf-sort ${sortState.key === 'count' ? 'is-active' : ''}" data-sort-key="count">${renderSortLabel('ครั้ง', sortState.key === 'count', sortState.dir)}</th>
                  <th class="sf-sort ${sortState.key === 'total' ? 'is-active' : ''}" data-sort-key="total">${renderSortLabel('รวม', sortState.key === 'total', sortState.dir)}</th>
                  <th class="sf-sort ${sortState.key === 'average' ? 'is-active' : ''}" data-sort-key="average">${renderSortLabel('เฉลี่ย', sortState.key === 'average', sortState.dir)}</th>
                  <th>Source</th>
                  <th>ล่าสุด 3 รายการ</th>
                </tr></thead>
                <tbody>${tableRows || '<tr><td colspan="11" class="empty-cell">ไม่พบข้อมูลตามเงื่อนไข</td></tr>'}</tbody>
              </table>
            </div>
          ` : `
            <div class="state-box">
              <strong>ยังไม่มี Focus Group</strong>
              <p>ไปที่หน้า Income Fund แล้วติ๊กกองทุนที่ต้องการก่อน จากนั้นกลับมาหน้านี้เพื่อดูค่าปันผลของกองที่เลือก</p>
            </div>
          `}
        </div>`;

      let timer;
      const qEl = $('#income-focus-q', area);
      qEl?.addEventListener('input', () => {
        clearTimeout(timer);
        qEl.value = qEl.value.toUpperCase();
        timer = setTimeout(() => {
          filters.query = qEl.value.trim().toLowerCase();
          render();
        }, 220);
      });
      $('#income-focus-source', area)?.addEventListener('change', e => {
        filters.source = e.target.value;
        render();
      });
      $('#income-focus-github-token', area)?.addEventListener('input', e => {
        State.incomeFundDividendGithubToken = e.target.value.trim();
      });
      $('#income-focus-refresh', area)?.addEventListener('click', () => Pages.incomeFundFocus(area, pageKey));
      const syncBtn = $('#income-focus-sync-dividend', area);
      if (syncBtn) {
        syncBtn.disabled = dividendSyncRunning;
        syncBtn.textContent = dividendSyncRunning ? 'กำลังสั่งดึงข้อมูล...' : 'ดึงข้อมูลปันผล';
        syncBtn.addEventListener('click', async () => {
          if (dividendSyncRunning) return;
          dividendSyncRunning = true;
          syncBtn.disabled = true;
          syncBtn.textContent = 'กำลังสั่งดึงข้อมูล...';
          try {
            const githubToken = $('#income-focus-github-token', area)?.value?.trim() || '';
            State.incomeFundDividendGithubToken = githubToken;
            const result = await triggerIncomeFundDividendSync(rows, { githubToken });
            toast(result.message || 'สั่งดึงข้อมูลปันผลแล้ว รอ GitHub Actions ทำงานเสร็จแล้วค่อยรีเฟรช', 'success', 6500);
          } catch (err) {
            toast(err.message || String(err), 'error', 8000);
          } finally {
            dividendSyncRunning = false;
            syncBtn.disabled = false;
            syncBtn.textContent = 'ดึงข้อมูลปันผล';
          }
        });
      }
      $$('.sf-sort', area).forEach(el => {
        el.addEventListener('click', () => {
          toggleNamedSort(sortState, el.dataset.sortKey);
          render();
        });
      });
      $$('.income-focus-detail', area).forEach(el => {
        el.addEventListener('click', () => {
          const row = rows.find(item => item.key === el.dataset.key);
          if (!row) return;
          Modal.openHtml(`ประวัติปันผล ${row.code}`, buildIncomeDividendDetailHtml(row));
        });
      });
    };

    render();
    App._currentExport = () => {
      if (!rows.length) {
        toast('ยังไม่มี Focus Group สำหรับ Export', 'warning');
        return;
      }
      exportExcel([
        ['Fund Code', 'Name', 'Policy', 'proj_id', 'Latest Dividend Date', 'Latest Dividend', 'History Count', 'Total Dividend', 'Average Dividend', 'Sources'],
        ...rows.map(row => [
          row.code,
          row.name,
          row.policy,
          row.projId,
          row.latestDate,
          row.latestValue,
          row.count,
          row.total,
          row.average,
          row.sources,
        ]),
      ], 'income-fund-focus');
    };
    App._currentTableExport = null;
    App._currentClipboardExport = null;
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

  robustnessFtImport(area) {
    const pageKey = 'robustness-ft-import';
    const today = new Date();
    const isoToday = today.toISOString().slice(0, 10);
    const threeYearsAgo = new Date(today);
    threeYearsAgo.setFullYear(today.getFullYear() - 3);
    const defaultStart = threeYearsAgo.toISOString().slice(0, 10);
    const state = State.ftHistoricalImport || {};
    if (!state.startDate) state.startDate = defaultStart;
    if (!state.endDate) state.endDate = isoToday;
    if (!state.url) state.url = 'https://markets.ft.com/data/etfs/tearsheet/summary?s=IXN:PCQ:USD';

    const parseFtSymbol = (value = '') => {
      const text = String(value || '').trim();
      if (!text) return '';
      try {
        const url = new URL(text);
        return (url.searchParams.get('s') || '').trim();
      } catch {
        const match = text.match(/[?&]s=([^&]+)/i);
        return match ? decodeURIComponent(match[1]).trim() : text;
      }
    };

    const symbolSlug = (symbol = '') => symbol.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    const parseIsoDateParts = (dateValue) => {
      const [year, month, day] = String(dateValue || '').split('-').map(Number);
      return { year, month, day };
    };
    const isoFromUtc = (utcMs) => new Date(utcMs).toISOString().slice(0, 10);
    const addDays = (dateValue, days) => {
      const { year, month, day } = parseIsoDateParts(dateValue);
      return isoFromUtc(Date.UTC(year, month - 1, day + days));
    };
    const addYears = (dateValue, years) => {
      const { year, month, day } = parseIsoDateParts(dateValue);
      return isoFromUtc(Date.UTC(year + years, month - 1, day));
    };
    const buildRanges = (startDate, endDate) => {
      if (!startDate || !endDate || startDate > endDate) return [];
      const ranges = [];
      let cursor = startDate;
      while (cursor <= endDate && ranges.length < 20) {
        const chunkEnd = [addDays(addYears(cursor, 1), -1), endDate].sort()[0];
        ranges.push({ start: cursor, end: chunkEnd });
        cursor = addDays(chunkEnd, 1);
      }
      return ranges;
    };

    const renderPreview = () => {
      const url = $('#ft-url', area)?.value || '';
      const startDate = $('#ft-start-date', area)?.value || '';
      const endDate = $('#ft-end-date', area)?.value || '';
      const symbol = parseFtSymbol(url);
      const slug = symbolSlug(symbol || 'SYMBOL');
      const ranges = buildRanges(startDate, endDate);
      State.ftHistoricalImport = { url, startDate, endDate, symbol };

      const symbolEl = $('#ft-symbol-preview', area);
      const commandEl = $('#ft-command-preview', area);
      const rangeEl = $('#ft-range-preview', area);
      const outputEl = $('#ft-output-preview', area);
      const statusEl = $('#ft-preview-status', area);
      const priceBtn = $('#ft-run-prices', area);
      const qualityBtn = $('#ft-run-qualitative', area);
      const copyBtn = $('#ft-copy-command', area);
      const priceValid = Boolean(symbol && startDate && endDate && startDate <= endDate);
      const qualityValid = Boolean(symbol);

      if (symbolEl) symbolEl.textContent = symbol || 'ยังอ่าน symbol ไม่ได้';
      if (rangeEl) {
        rangeEl.innerHTML = ranges.length
          ? ranges.map((range, idx) => `<span class="badge">${idx + 1}. ${esc(range.start)} ถึง ${esc(range.end)}</span>`).join(' ')
          : '<span class="badge badge-warning">เลือกวันที่เริ่มต้น/สิ้นสุดให้ถูกต้อง</span>';
      }
      if (outputEl) {
        outputEl.innerHTML = `
          <div><strong>Price CSV</strong> Data/ft_historical_prices/prices/${esc(slug)}.csv</div>
          <div><strong>Qualitative Snapshot</strong> Data/ft_historical_prices/qualitative/as_of_YYYY-MM-DD/${esc(slug)}/</div>
          <div><strong>SQLite</strong> Data/ft_historical_prices/ft_historical_prices.sqlite</div>
        `;
      }
      if (commandEl) {
        commandEl.textContent = symbol
          ? [
              `ราคา: python3 scripts/ft_historical_prices_store.py --symbol '${symbol}' --start-date ${startDate || 'YYYY-MM-DD'} --end-date ${endDate || 'YYYY-MM-DD'}`,
              `ข้อมูลเชิงคุณภาพ: python3 scripts/ft_historical_prices_store.py --symbol '${symbol}' --qualitative-only`,
            ].join('\n')
          : 'ใส่ FT link หรือ symbol ให้ครบก่อน';
      }
      if (statusEl) {
        statusEl.className = qualityValid ? 'badge badge-primary' : 'badge badge-warning';
        statusEl.textContent = priceValid ? `พร้อมดึงราคา ${ranges.length} ช่วง / ข้อมูลเชิงคุณภาพ` : (qualityValid ? 'พร้อมดึงข้อมูลเชิงคุณภาพ' : 'ยังไม่พร้อม');
      }
      if (priceBtn) priceBtn.disabled = !priceValid;
      if (qualityBtn) qualityBtn.disabled = !qualityValid;
      if (copyBtn) copyBtn.disabled = !qualityValid;
    };

    area.innerHTML = `
      <div class="grid grid-2">
        <section class="card">
          <div class="card-header">
            <div>
              <h3>FT Historical Prices</h3>
              <p>รับลิงก์ FT แล้วดึงตาราง Date, Open, High, Low, Close, Volume</p>
            </div>
            <span class="badge badge-primary" id="ft-preview-status">พร้อมตรวจ</span>
          </div>
          <div class="form-grid">
            <label class="form-field form-field-full">
              <span>FT.com link</span>
              <input id="ft-url" type="url" value="${esc(state.url)}" placeholder="https://markets.ft.com/data/etfs/tearsheet/summary?s=IXN:PCQ:USD">
            </label>
            <label class="form-field">
              <span>วันที่เริ่มต้น</span>
              <input id="ft-start-date" type="date" value="${esc(state.startDate)}">
            </label>
            <label class="form-field">
              <span>วันที่สิ้นสุด</span>
              <input id="ft-end-date" type="date" value="${esc(state.endDate)}">
            </label>
          </div>
          <div class="toolbar">
            <button class="btn btn-primary" id="ft-run-prices" type="button">ดึงราคา Historical Prices</button>
            <button class="btn btn-secondary" id="ft-run-qualitative" type="button">ดึงข้อมูลเชิงคุณภาพ</button>
            <button class="btn btn-secondary" id="ft-copy-command" type="button">คัดลอกคำสั่ง</button>
            <button class="btn btn-ghost" id="ft-reset-3y" type="button">ย้อนหลัง 3 ปี</button>
          </div>
        </section>

        <section class="card">
          <div class="card-header">
            <div>
              <h3>Preview Plan</h3>
              <p>FT จำกัดการดูครั้งละประมาณ 1 ปี ระบบจะแบ่งช่วงให้ก่อนดึงจริง</p>
            </div>
          </div>
          <div class="kv-list">
            <div><span>Symbol</span><strong id="ft-symbol-preview">-</strong></div>
            <div><span>แบ่งช่วง</span><div id="ft-range-preview" class="badge-row"></div></div>
            <div><span>ไฟล์ปลายทาง</span><div id="ft-output-preview" class="mono-small"></div></div>
          </div>
        </section>
      </div>

      <section class="card">
        <div class="card-header">
          <div>
            <h3>Local Command</h3>
            <p>ใช้คำสั่งนี้กับ script ที่เตรียมไว้เพื่อสร้าง CSV, SQLite และ raw JSON ใน Google Drive project</p>
          </div>
        </div>
        <pre class="code-preview" id="ft-command-preview"></pre>
      </section>

      <section class="card hidden" id="ft-result-card">
        <div class="card-header">
          <div>
            <h3>ผลการดึงข้อมูล</h3>
            <p id="ft-result-summary">-</p>
          </div>
        </div>
        <div class="kv-list" id="ft-result-detail"></div>
      </section>
    `;

    ['ft-url', 'ft-start-date', 'ft-end-date'].forEach(id => {
      $(`#${id}`, area)?.addEventListener('input', renderPreview);
      $(`#${id}`, area)?.addEventListener('change', renderPreview);
    });
    $('#ft-reset-3y', area)?.addEventListener('click', () => {
      $('#ft-start-date', area).value = defaultStart;
      $('#ft-end-date', area).value = isoToday;
      renderPreview();
    });
    $('#ft-copy-command', area)?.addEventListener('click', async () => {
      const command = $('#ft-command-preview', area)?.textContent || '';
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(command);
        toast('คัดลอกคำสั่งดึงข้อมูล FT แล้ว', 'success', 2200);
      } else {
        toast('เตรียมคำสั่งไว้ด้านล่างแล้ว', 'info', 2200);
      }
    });
    $('#ft-run-prices', area)?.addEventListener('click', async () => {
      renderPreview();
      const runBtn = $('#ft-run-prices', area);
      const resultCard = $('#ft-result-card', area);
      const resultSummary = $('#ft-result-summary', area);
      const resultDetail = $('#ft-result-detail', area);
      const payload = {
        url: $('#ft-url', area)?.value || '',
        symbol: State.ftHistoricalImport?.symbol || '',
        startDate: $('#ft-start-date', area)?.value || '',
        endDate: $('#ft-end-date', area)?.value || '',
      };
      if (!payload.symbol || !payload.startDate || !payload.endDate || payload.startDate > payload.endDate) {
        toast('กรุณาใส่ FT link และช่วงวันที่ให้ถูกต้อง', 'warning', 3000);
        return;
      }
      resultCard?.classList.remove('hidden');
      if (resultSummary) resultSummary.textContent = 'กำลังดึงข้อมูลจาก FT.com และเขียนไฟล์ลง Data/ft_historical_prices...';
      if (resultDetail) resultDetail.innerHTML = '';
      if (runBtn) {
        runBtn.disabled = true;
        runBtn.textContent = 'กำลังดึงข้อมูล...';
      }
      try {
        let result = {};
        if (ftHistoricalApiUrl()) {
          const data = await triggerFtHistoricalSync({
            ...payload,
            runPrices: true,
            runQualitative: true,
          });
          result = data.result || data;
          if (resultSummary) {
            resultSummary.textContent = data.message || `สั่ง GitHub Actions ให้ดึงราคา ${payload.symbol} แล้ว`;
          }
          if (resultDetail) {
            resultDetail.innerHTML = `
              <div><span>Workflow</span><strong>${esc(result.workflow || 'ft-historical-prices-database.yml')}</strong></div>
              <div><span>Drive Folder</span><strong>${esc(result.driveFolderId || FT_HISTORICAL_DRIVE_FOLDER_ID)}</strong></div>
              <div><span>ไฟล์ปลายทาง</span><strong>${esc(result.fileName || FT_HISTORICAL_DB_FILE_NAME)}</strong></div>
              <div><span>สถานะ</span><strong>รอ GitHub Actions ทำงานเสร็จ แล้วไฟล์ JSON จะถูกอัปโหลดกลับเข้า Google Drive</strong></div>
            `;
          }
          toast(data.message || 'สั่งดึงราคา FT ผ่าน GitHub Actions แล้ว', 'success', 6500);
          return;
        }
        const resp = await fetch('/api/ft-historical-prices', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        const data = await resp.json().catch(() => ({}));
        if (!resp.ok || data.ok === false) throw new Error(data.error || `FT sync failed (${resp.status})`);
        result = data.result || {};
        if (resultSummary) {
          resultSummary.textContent = `${result.symbol || payload.symbol}: ${Number(result.totalRows || 0).toLocaleString()} rows · ${result.startDate || payload.startDate} ถึง ${result.endDate || payload.endDate}`;
        }
        if (resultDetail) {
          resultDetail.innerHTML = `
            <div><span>CSV</span><strong>${esc(result.csvPath || '-')}</strong></div>
            <div><span>SQLite</span><strong>${esc(result.sqlitePath || '-')}</strong></div>
            <div><span>แบ่งช่วงจริง</span><div class="badge-row">${(result.ranges || []).map((range, idx) => `<span class="badge">${idx + 1}. ${esc(range.startDate)} ถึง ${esc(range.endDate)} · ${Number(range.rows || 0).toLocaleString()} rows</span>`).join('')}</div></div>
          `;
        }
        toast('ดึงข้อมูล FT สำเร็จ', 'success', 2600);
      } catch (err) {
        if (resultSummary) resultSummary.textContent = err.message || 'ดึงข้อมูล FT ไม่สำเร็จ';
        if (resultDetail) {
          resultDetail.innerHTML = '<div><span>สถานะ</span><strong>ตรวจว่าเปิดผ่าน local server และ network เข้าถึง markets.ft.com ได้</strong></div>';
        }
        toast(err.message || 'ดึงข้อมูล FT ไม่สำเร็จ', 'error', 6000);
      } finally {
        if (runBtn) {
          runBtn.textContent = 'ดึงราคา Historical Prices';
          renderPreview();
        }
      }
    });
    $('#ft-run-qualitative', area)?.addEventListener('click', async () => {
      renderPreview();
      const runBtn = $('#ft-run-qualitative', area);
      const resultCard = $('#ft-result-card', area);
      const resultSummary = $('#ft-result-summary', area);
      const resultDetail = $('#ft-result-detail', area);
      const payload = {
        url: $('#ft-url', area)?.value || '',
        symbol: State.ftHistoricalImport?.symbol || '',
      };
      if (!payload.symbol) {
        toast('กรุณาใส่ FT link หรือ symbol ให้ถูกต้อง', 'warning', 3000);
        return;
      }
      resultCard?.classList.remove('hidden');
      if (resultSummary) resultSummary.textContent = 'กำลังดึงข้อมูล profile/risk/holdings จาก FT.com และเขียนไฟล์ลง Data/ft_historical_prices...';
      if (resultDetail) resultDetail.innerHTML = '';
      if (runBtn) {
        runBtn.disabled = true;
        runBtn.textContent = 'กำลังดึงข้อมูล...';
      }
      try {
        let result = {};
        if (ftHistoricalApiUrl()) {
          const data = await triggerFtHistoricalSync({
            ...payload,
            startDate: $('#ft-start-date', area)?.value || '',
            endDate: $('#ft-end-date', area)?.value || '',
            runPrices: true,
            runQualitative: true,
          });
          result = data.result || data;
          if (resultSummary) {
            resultSummary.textContent = data.message || `สั่ง GitHub Actions ให้ดึงข้อมูลเชิงคุณภาพ ${payload.symbol} แล้ว`;
          }
          if (resultDetail) {
            resultDetail.innerHTML = `
              <div><span>Workflow</span><strong>${esc(result.workflow || 'ft-historical-prices-database.yml')}</strong></div>
              <div><span>Drive Folder</span><strong>${esc(result.driveFolderId || FT_HISTORICAL_DRIVE_FOLDER_ID)}</strong></div>
              <div><span>ไฟล์ปลายทาง</span><strong>${esc(result.fileName || FT_HISTORICAL_DB_FILE_NAME)}</strong></div>
              <div><span>สถานะ</span><strong>รอ GitHub Actions ทำงานเสร็จ แล้วไฟล์ JSON จะถูกอัปโหลดกลับเข้า Google Drive</strong></div>
            `;
          }
          toast(data.message || 'สั่งดึงข้อมูลเชิงคุณภาพ FT ผ่าน GitHub Actions แล้ว', 'success', 6500);
          return;
        }
        const resp = await fetch('/api/ft-qualitative-data', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        const data = await resp.json().catch(() => ({}));
        if (!resp.ok || data.ok === false) throw new Error(data.error || `FT qualitative sync failed (${resp.status})`);
        result = data.result || {};
        if (resultSummary) {
          const name = result.displayName ? ` · ${result.displayName}` : '';
          resultSummary.textContent = `${result.symbol || payload.symbol}${name}: profile ${Number(result.profileRows || 0).toLocaleString()} fields · risk ${Number(result.riskRows || 0).toLocaleString()} rows · holdings ${Number(result.holdingsRows || 0).toLocaleString()} rows`;
        }
        if (resultDetail) {
          resultDetail.innerHTML = `
            <div><span>ชื่อจาก FT</span><strong>${esc(result.displayName || '-')}</strong></div>
            <div><span>Snapshot As of</span><strong>${esc(result.asOfDate || '-')}</strong></div>
            <div><span>Profile CSV</span><strong>${esc(result.profileCsvPath || '-')} · ${Number(result.profileRows || 0).toLocaleString()} fields</strong></div>
            <div><span>Risk CSV</span><strong>${esc(result.riskCsvPath || '-')} · ${Number(result.riskRows || 0).toLocaleString()} rows</strong></div>
            <div><span>Holdings CSV</span><strong>${esc(result.holdingsCsvPath || '-')} · ${Number(result.holdingsRows || 0).toLocaleString()} rows</strong></div>
            <div><span>Risk As of</span><strong>${esc(result.riskAsOfDate || '-')}</strong></div>
            <div><span>Holdings As of</span><strong>${esc(result.holdingsAsOfDate || '-')}</strong></div>
            <div><span>SQLite</span><strong>${esc(result.sqlitePath || '-')}</strong></div>
            <div><span>Snapshot Folder</span><strong>${esc(result.snapshotDir || '-')}</strong></div>
          `;
        }
        toast('ดึงข้อมูลเชิงคุณภาพ FT สำเร็จ', 'success', 2600);
      } catch (err) {
        if (resultSummary) resultSummary.textContent = err.message || 'ดึงข้อมูลเชิงคุณภาพ FT ไม่สำเร็จ';
        if (resultDetail) {
          resultDetail.innerHTML = '<div><span>สถานะ</span><strong>ตรวจว่าเปิดผ่าน local server และ network เข้าถึง markets.ft.com ได้</strong></div>';
        }
        toast(err.message || 'ดึงข้อมูลเชิงคุณภาพ FT ไม่สำเร็จ', 'error', 6000);
      } finally {
        if (runBtn) {
          runBtn.textContent = 'ดึงข้อมูลเชิงคุณภาพ';
          renderPreview();
        }
      }
    });
    renderPreview();
    State._pageDataSource[pageKey] = 'FT Markets historical prices';
    App._currentExport = null;
    App._currentTableExport = null;
  },

  async upsideDownsideCapture(area) {
    const pageKey = 'upside-downside-capture';
    setLoading(area, 'กำลังโหลดข้อมูล Historical Prices จาก Google Drive JSON...');

    const pct = (value) => Number.isFinite(value) ? `${(value * 100).toFixed(2)}%` : '-';
    const num = (value, digits = 2) => Number.isFinite(value) ? Number(value).toLocaleString(undefined, { maximumFractionDigits: digits, minimumFractionDigits: digits }) : '-';
    const int = (value) => Number.isFinite(value) ? Number(value).toLocaleString() : '-';

    let payload;
    try {
      payload = await loadFtHistoricalPayloadForPage(5000);
    } catch (err) {
      area.innerHTML = `
        <div class="card">
          <div class="state-box">
            <div class="state-icon">!</div>
            <strong>ยังโหลดข้อมูล FT จาก Google Drive JSON ไม่ได้</strong>
            <p>${esc(err.message || 'กรุณาตรวจว่า ft_historical_prices_database.json ถูกสร้างใน Drive แล้ว')}</p>
          </div>
        </div>`;
      return;
    }

    const symbols = payload.symbols || [];
    const rows = payload.prices || payload.rows || [];
    const profileRows = payload.profile || [];
    const symbolItems = symbols.filter(item => item.symbol);
    const symbolOptions = symbolItems.map(item => item.symbol);
    const symbolMetaBySymbol = new Map(symbolItems.map(item => [item.symbol, item]));
    const defaultSymbol = payload.selectedSymbol || symbolOptions[0] || '';
    const selectedRowsForStats = rows
      .filter(row => row.symbol === defaultSymbol && row.date && Number.isFinite(Number(row.close)))
      .sort((a, b) => String(a.date).localeCompare(String(b.date)));
    const computedReturns = [];
    let prevClose = null;
    selectedRowsForStats.forEach(row => {
      const close = Number(row.close);
      if (prevClose && close) computedReturns.push((close / prevClose) - 1);
      if (close) prevClose = close;
    });
    const positiveReturns = computedReturns.filter(value => value > 0);
    const negativeReturns = computedReturns.filter(value => value < 0);
    const firstClose = selectedRowsForStats[0]?.close;
    const lastClose = selectedRowsForStats[selectedRowsForStats.length - 1]?.close;
    const stat = (payload.stats || []).find(item => item.symbol === defaultSymbol) || {
      rowCount: selectedRowsForStats.length,
      startDate: selectedRowsForStats[0]?.date || '',
      endDate: selectedRowsForStats[selectedRowsForStats.length - 1]?.date || '',
      cumulativeReturn: firstClose && lastClose ? (Number(lastClose) / Number(firstClose)) - 1 : null,
      dailyReturnCount: computedReturns.length,
      upDays: positiveReturns.length,
      downDays: negativeReturns.length,
      avgUpDayReturn: positiveReturns.length ? positiveReturns.reduce((sum, value) => sum + value, 0) / positiveReturns.length : null,
      avgDownDayReturn: negativeReturns.length ? negativeReturns.reduce((sum, value) => sum + value, 0) / negativeReturns.length : null,
    };
    const symbolLabel = (symbol) => symbolMetaBySymbol.get(symbol)?.displayLabel || symbol || '-';
    const symbolSelectOptions = (selectedSymbol = defaultSymbol) => symbolOptions
      .map(symbol => `<option value="${esc(symbol)}"${symbol === selectedSymbol ? ' selected' : ''}>${esc(symbolLabel(symbol))}</option>`)
      .join('');
    const captureRowHtml = (rowIdx) => `
      <tr data-uc-row="${rowIdx}">
        <td class="uc-benchmark-cell">${esc(symbolLabel(defaultSymbol))}</td>
        ${[1, 2, 3].map(point => `
          <td>
            <div class="uc-nav-cell">
              <input class="uc-nav" data-point="${point}" type="text" readonly placeholder="-">
              <small class="uc-actual-date" data-point="${point}"></small>
            </div>
          </td>
          <td>
            <input class="uc-date" data-point="${point}" type="date">
          </td>
        `).join('')}
        <td class="td-num uc-calc uc-downside-gap">-</td>
        <td class="td-num uc-calc uc-downside-capture">-</td>
        <td class="td-num uc-calc uc-upside-gap">-</td>
        <td class="td-num uc-calc uc-upside-capture">-</td>
        <td class="td-num uc-calc uc-robustness-ratio">-</td>
        <td class="td-num uc-calc uc-total-cycle-fund">-</td>
        <td class="td-num uc-calc uc-total-cycle-benchmark">-</td>
        <td class="td-num uc-calc uc-robustness-score">-</td>
        <td><button class="btn btn-ghost btn-xs uc-delete-row" type="button">ลบ</button></td>
      </tr>
    `;
    const captureCompareRowHtml = (rowIdx, dates = ['', '', '']) => `
      <tr data-uc-compare-row="${rowIdx}">
        <td class="uc-compare-cell">${esc(symbolLabel(defaultSymbol))}</td>
        ${[1, 2, 3].map(point => `
          <td>
            <div class="uc-nav-cell">
              <input class="uc-nav" data-point="${point}" type="text" readonly placeholder="-">
              <small class="uc-actual-date" data-point="${point}"></small>
            </div>
          </td>
          <td>
            <input class="uc-date uc-date-synced" data-point="${point}" type="date" value="${esc(dates[point - 1] || '')}" disabled>
          </td>
        `).join('')}
        <td class="td-num uc-calc uc-downside-gap">-</td>
        <td class="td-num uc-calc uc-downside-capture">-</td>
        <td class="td-num uc-calc uc-upside-gap">-</td>
        <td class="td-num uc-calc uc-upside-capture">-</td>
        <td class="td-num uc-calc uc-robustness-ratio">-</td>
        <td class="td-num uc-calc uc-total-cycle-fund">-</td>
        <td class="td-num uc-calc uc-total-cycle-benchmark">-</td>
        <td class="td-num uc-calc uc-robustness-score">-</td>
        <td></td>
      </tr>
    `;
    const localPriceRows = rows
      .filter(row => row.symbol && row.date && Number.isFinite(Number(row.close)))
      .sort((a, b) => String(a.date).localeCompare(String(b.date)));

    if (!symbols.length) {
      area.innerHTML = `
        <div class="card">
          <div class="state-box">
            <div class="state-icon">🗂</div>
            <strong>ยังไม่มีข้อมูล FT Historical Prices</strong>
            <p>ไปที่เมนูเตรียมข้อมูลจาก FT.com แล้วดึงข้อมูลอย่างน้อย 1 symbol ก่อน</p>
          </div>
        </div>`;
      return;
    }

    area.innerHTML = `
      <div class="stats-grid">
        <div class="stat-card">
          <div class="stat-label">Symbol ที่เลือก</div>
          <div class="stat-value stat-value-small">${esc(payload.selectedDisplayLabel || symbolLabel(payload.selectedSymbol || defaultSymbol))}</div>
        </div>
        <div class="stat-card">
          <div class="stat-label">จำนวนราคา</div>
          <div class="stat-value">${int(stat.rowCount || 0)}</div>
        </div>
        <div class="stat-card">
          <div class="stat-label">ช่วงข้อมูล</div>
          <div class="stat-value stat-value-small">${esc(stat.startDate || '-')} ถึง ${esc(stat.endDate || '-')}</div>
        </div>
        <div class="stat-card">
          <div class="stat-label">ผลตอบแทนสะสมจาก Close</div>
          <div class="stat-value">${pct(stat.cumulativeReturn)}</div>
        </div>
      </div>

      <div class="grid grid-2">
        <section class="card">
          <div class="card-header">
            <div>
              <h3>Profile and Investment</h3>
              <p>Metadata จากหน้า FT summary สำหรับช่วยเลือก benchmark/fund</p>
            </div>
          </div>
          <div class="kv-list">
            ${profileRows.length ? profileRows.map(item => `
              <div>
                <span>${esc(item.field || '')}</span>
                <strong>${esc(item.value || '--').replace(/\n/g, '<br>')}</strong>
              </div>
            `).join('') : `
              <div><span>สถานะ</span><strong>ยังไม่มี Profile and investment สำหรับ symbol นี้ ให้ดึงข้อมูล FT ใหม่อีกครั้ง</strong></div>
            `}
          </div>
        </section>

        <section class="card">
          <div class="card-header">
            <div>
              <h3>ค่าที่มีสำหรับคำนวณ Capture</h3>
              <p>ชุดนี้คือข้อมูลจาก SQLite หลังแปลงจาก FT Historical Prices แล้ว</p>
            </div>
            <span class="badge badge-data-origin">${esc(payload.source || 'SQLite')}</span>
          </div>
          <div class="kv-list">
            <div><span>Open / High / Low / Close</span><strong>มีครบในระดับรายวัน</strong></div>
            <div><span>Volume</span><strong>มีสำหรับตรวจ liquidity/ความผิดปกติ</strong></div>
            <div><span>Daily Return</span><strong>คำนวณจาก Close เทียบวันก่อนหน้าได้ ${int(stat.dailyReturnCount || 0)} ค่า</strong></div>
            <div><span>Up Days</span><strong>${int(stat.upDays || 0)} วัน · เฉลี่ย ${pct(stat.avgUpDayReturn)}</strong></div>
            <div><span>Down Days</span><strong>${int(stat.downDays || 0)} วัน · เฉลี่ย ${pct(stat.avgDownDayReturn)}</strong></div>
          </div>
        </section>

        <section class="card">
          <div class="card-header">
            <div>
              <h3>Available Symbols</h3>
              <p>ตอนนี้เลือกให้ดูตัวแรกก่อน รอบถัดไปค่อยเพิ่มตัวเลือก fund/benchmark</p>
            </div>
          </div>
          <div class="table-wrapper" style="max-height:320px">
            <table>
              <thead>
                <tr>
                  <th>Symbol</th>
                  <th>Name</th>
                  <th>Rows</th>
                  <th>Start</th>
                  <th>End</th>
                </tr>
              </thead>
              <tbody>
                ${symbols.map(item => `
                  <tr>
                    <td>${esc(item.symbol || '')}</td>
                    <td>${esc(item.displayName || '-')}</td>
                    <td class="td-num">${int(item.rowCount || 0)}</td>
                    <td>${esc(item.startDate || '')}</td>
                    <td>${esc(item.endDate || '')}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </section>
      </div>

      <section class="card" id="uc-nav-chart-card">
        <div class="card-header">
          <div>
            <h3>NAV ช่วงที่เลือก</h3>
            <p>กราฟเส้นจากวันที่ NAV (1) ถึง NAV (3) ของ cycle ที่เลือก</p>
          </div>
          <label class="form-field uc-chart-cycle-field">
            <span>Cycle</span>
            <select id="uc-chart-row-select"></select>
          </label>
        </div>
        <div class="uc-chart-summary" id="uc-chart-summary">เลือกวันที่ NAV (1) และ NAV (3) เพื่อแสดงกราฟ</div>
        <div class="uc-nav-chart" id="uc-nav-chart">
          <div class="state-box compact">ยังไม่มีช่วงวันที่สำหรับวาดกราฟ</div>
        </div>
      </section>

      <section class="card">
        <div class="card-header">
          <div>
            <h3>Capture Date Builder</h3>
            <p>เลือก Benchmark ครั้งเดียว แล้วเพิ่มวันที่แต่ละ cycle เพื่อดึง Close มาเป็น NAV</p>
          </div>
          <span class="badge badge-data-origin">NAV = Close</span>
        </div>
        <div class="uc-toolbar">
          <label class="form-field">
            <span>Benchmark</span>
            <select id="uc-benchmark-select">
              ${symbolSelectOptions(defaultSymbol)}
            </select>
          </label>
          <button class="btn btn-secondary" id="uc-add-compare-asset" type="button">เพิ่มสินทรัพย์เปรียบเทียบ</button>
        </div>
        <div class="table-wrapper uc-table-wrapper">
          <table class="uc-table">
            <thead>
              <tr>
                <th rowspan="2">Benchmark</th>
                <th colspan="2">Peak</th>
                <th colspan="2">Bottom</th>
                <th colspan="2">Peak</th>
                <th rowspan="2">Downside Gap<br>[(2)-(1)]/(1)</th>
                <th rowspan="2">Downside Capture</th>
                <th rowspan="2">Upside Gap<br>[(3)-(2)]/(2)</th>
                <th rowspan="2">Upside Capture</th>
                <th rowspan="2">Robustness Ratio</th>
                <th rowspan="2">Total Cycle Return<br>(Fund)</th>
                <th rowspan="2">Total Cycle Return<br>(Benchmark)</th>
                <th rowspan="2">Robustness Score<br>(Fund/Benchmark)</th>
                <th rowspan="2">จัดการ</th>
              </tr>
              <tr>
                <th>NAV (1)</th>
                <th>Date</th>
                <th>NAV (2)</th>
                <th>Date</th>
                <th>NAV (3)</th>
                <th>Date</th>
              </tr>
            </thead>
            <tbody id="uc-cycle-body">
              ${Array.from({ length: 5 }, (_, rowIdx) => captureRowHtml(rowIdx)).join('')}
              <tr class="uc-add-row">
                <td colspan="16">
                  <button class="btn btn-secondary" id="uc-add-row" type="button">เพิ่มวันที่</button>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        <div class="uc-compare-assets" id="uc-compare-assets"></div>
      </section>
    `;

    const symbolRowsCache = new Map();
    symbolOptions.forEach(symbol => {
      const priceRows = localPriceRows.filter(row => row.symbol === symbol);
      if (priceRows.length) symbolRowsCache.set(symbol, priceRows);
    });
    if (defaultSymbol && !symbolRowsCache.has(defaultSymbol)) symbolRowsCache.set(defaultSymbol, localPriceRows);

    const lookupClose = async (symbol, dateValue) => {
      if (!symbol || !dateValue) return null;
      const localMatch = [...localPriceRows]
        .filter(row => row.symbol === symbol && row.date <= dateValue)
        .pop();
      if (localMatch) {
        return {
          ...localMatch,
          requestedDate: dateValue,
          isExactDate: localMatch.date === dateValue,
          source: 'loaded-page',
        };
      }
      const url = `/api/ft-price-on-date?symbol=${encodeURIComponent(symbol)}&date=${encodeURIComponent(dateValue)}`;
      const resp = await fetch(url, { cache: 'no-store' });
      const data = await resp.json().catch(() => ({}));
      if (resp.status === 404) {
        throw new Error('ไม่พบ API /api/ft-price-on-date กรุณา restart fund_server.py');
      }
      if (!resp.ok || data.ok === false) throw new Error(data.error || `โหลด NAV ไม่สำเร็จ (${resp.status})`);
      return data.price || null;
    };

    const getSymbolPriceRows = async (symbol) => {
      if (!symbol) return [];
      if (symbolRowsCache.has(symbol)) return symbolRowsCache.get(symbol);
      const url = `/api/ft-historical-prices?symbol=${encodeURIComponent(symbol)}&limit=5000`;
      const resp = await fetch(url, { cache: 'no-store' });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok || data.ok === false) throw new Error(data.error || `โหลดราคา ${symbol} ไม่สำเร็จ`);
      const priceRows = (data.rows || [])
        .filter(row => row.symbol && row.date && Number.isFinite(Number(row.close)))
        .sort((a, b) => String(a.date).localeCompare(String(b.date)));
      symbolRowsCache.set(symbol, priceRows);
      return priceRows;
    };

    const getSelectedChartRow = () => {
      const select = $('#uc-chart-row-select', area);
      const rowsInTable = $$('#uc-cycle-body tr[data-uc-row]', area);
      const selectedIndex = Number(select?.value || 0);
      return rowsInTable[selectedIndex] || rowsInTable[0] || null;
    };

    const getRowDateValue = (rowEl, point) => $(`.uc-date[data-point="${point}"]`, rowEl)?.value || '';

    const getChartSymbols = () => {
      const list = [$('#uc-benchmark-select', area)?.value || defaultSymbol]
        .concat($$('.uc-compare-symbol', area).map(select => select.value).filter(Boolean));
      return [...new Set(list.filter(Boolean))];
    };

    const formatShortDate = (value) => {
      if (!value) return '';
      const parts = String(value).split('-');
      return parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0].slice(2)}` : value;
    };

    const buildNavChartSvg = (series, markers, startDate, endDate, highlightRanges = []) => {
      const width = 1080;
      const height = 360;
      const pad = { left: 68, right: 22, top: 28, bottom: 48 };
      const chartW = width - pad.left - pad.right;
      const chartH = height - pad.top - pad.bottom;
      const allPoints = series.flatMap(item => item.points);
      if (!allPoints.length) {
        return '<div class="state-box compact">ไม่มีราคาในช่วงวันที่นี้สำหรับ symbol ที่เลือก</div>';
      }
      const times = allPoints.map(point => new Date(`${point.date}T00:00:00`).getTime()).filter(Number.isFinite);
      const values = allPoints.map(point => Number(point.close)).filter(Number.isFinite);
      const minT = Math.min(...times);
      const maxT = Math.max(...times);
      const minVRaw = Math.min(...values);
      const maxVRaw = Math.max(...values);
      const valuePad = (maxVRaw - minVRaw) * 0.08 || Math.max(maxVRaw * 0.04, 1);
      const minV = minVRaw - valuePad;
      const maxV = maxVRaw + valuePad;
      const x = (date) => {
        const t = new Date(`${date}T00:00:00`).getTime();
        return pad.left + ((t - minT) / ((maxT - minT) || 1)) * chartW;
      };
      const y = (value) => pad.top + (1 - ((value - minV) / ((maxV - minV) || 1))) * chartH;
      const colors = ['#1a3c6e', '#2d9e6b', '#d63b3b', '#8b5cf6', '#e8a317', '#4a90d9'];
      const yTicks = Array.from({ length: 5 }, (_, idx) => minV + ((maxV - minV) * idx / 4));
      const markerItems = markers.filter(marker => marker.date);
      const rangeHighlight = (fromDate, toDate, className) => {
        if (!fromDate || !toDate) return '';
        const x1 = Math.max(pad.left, Math.min(pad.left + chartW, x(fromDate)));
        const x2 = Math.max(pad.left, Math.min(pad.left + chartW, x(toDate)));
        const left = Math.min(x1, x2);
        const widthValue = Math.abs(x2 - x1);
        if (!Number.isFinite(widthValue) || widthValue <= 0) return '';
        return `<rect x="${left.toFixed(2)}" y="${pad.top}" width="${widthValue.toFixed(2)}" height="${chartH}" class="${className}"></rect>`;
      };
      const ranges = highlightRanges.length ? highlightRanges : [
        { from: markers[0]?.date, to: markers[1]?.date, className: 'uc-chart-highlight-downside' },
        { from: markers[1]?.date, to: markers[2]?.date, className: 'uc-chart-highlight-upside' },
      ];
      const highlights = ranges
        .map(range => rangeHighlight(range.from, range.to, range.className))
        .join('');
      const markerLines = markerItems.map((marker) => {
        const date = marker.date;
        const markerX = x(date);
        return `
          <line x1="${markerX.toFixed(2)}" y1="${pad.top}" x2="${markerX.toFixed(2)}" y2="${pad.top + chartH}" class="uc-chart-marker-line"></line>
          <text x="${markerX.toFixed(2)}" y="${height - 14}" text-anchor="middle" class="uc-chart-marker-label">${esc(marker.label)}</text>
        `;
      }).join('');
      const paths = series.map((item, idx) => {
        const path = item.points.map((point, pointIdx) => {
          const cmd = pointIdx === 0 ? 'M' : 'L';
          return `${cmd}${x(point.date).toFixed(2)},${y(Number(point.close)).toFixed(2)}`;
        }).join(' ');
        const color = colors[idx % colors.length];
        const dots = markers
          .map(marker => item.points.find(point => point.date === marker.date))
          .filter(Boolean)
          .map(point => `<circle cx="${x(point.date).toFixed(2)}" cy="${y(Number(point.close)).toFixed(2)}" r="4" fill="${color}"></circle>`)
          .join('');
        return `
          <path d="${path}" fill="none" stroke="${color}" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"></path>
          ${dots}
        `;
      }).join('');
      const legend = series.map((item, idx) => {
        const color = colors[idx % colors.length];
        return `
          <span class="uc-chart-legend-item">
            <i style="background:${color}"></i>${esc(item.label || item.symbol)}
          </span>
        `;
      }).join('');
      const yGrid = yTicks.map(value => {
        const yy = y(value);
        return `
          <line x1="${pad.left}" y1="${yy.toFixed(2)}" x2="${pad.left + chartW}" y2="${yy.toFixed(2)}" class="uc-chart-grid"></line>
          <text x="${pad.left - 10}" y="${(yy + 4).toFixed(2)}" text-anchor="end" class="uc-chart-axis-label">${num(value, 2)}</text>
        `;
      }).join('');
      return `
        <div class="uc-chart-legend">${legend}</div>
        <svg viewBox="0 0 ${width} ${height}" class="uc-chart-svg" role="img" aria-label="NAV from ${esc(startDate)} to ${esc(endDate)}">
          <rect x="0" y="0" width="${width}" height="${height}" rx="8" class="uc-chart-bg"></rect>
          ${highlights}
          ${yGrid}
          ${markerLines}
          ${paths}
          <line x1="${pad.left}" y1="${pad.top + chartH}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" class="uc-chart-axis"></line>
          <text x="${pad.left}" y="${height - 14}" text-anchor="start" class="uc-chart-axis-label">${esc(formatShortDate(startDate))}</text>
          <text x="${pad.left + chartW}" y="${height - 14}" text-anchor="end" class="uc-chart-axis-label">${esc(formatShortDate(endDate))}</text>
        </svg>
      `;
    };

    const refreshChartRowOptions = () => {
      const select = $('#uc-chart-row-select', area);
      if (!select) return;
      const current = select.value;
      const count = $$('#uc-cycle-body tr[data-uc-row]', area).length;
      select.innerHTML = [
        '<option value="all">All</option>',
        ...Array.from({ length: count }, (_, idx) => `<option value="${idx}">Cycle ${idx + 1}</option>`),
      ].join('');
      if (current === 'all') {
        select.value = 'all';
      } else if (current && Number(current) < count) {
        select.value = current;
      } else if (!current) {
        select.value = 'all';
      }
    };

    const renderNavChart = async () => {
      const chartEl = $('#uc-nav-chart', area);
      const summaryEl = $('#uc-chart-summary', area);
      const selectedChartMode = $('#uc-chart-row-select', area)?.value || 'all';
      const cycleRows = selectedChartMode === 'all'
        ? $$('#uc-cycle-body tr[data-uc-row]', area)
        : [getSelectedChartRow()].filter(Boolean);
      const chartCycles = cycleRows
        .map((rowEl, idx) => {
          const rowIndex = selectedChartMode === 'all'
            ? $$('#uc-cycle-body tr[data-uc-row]', area).indexOf(rowEl)
            : Number(selectedChartMode || 0);
          return {
            label: `Cycle ${Number.isFinite(rowIndex) ? rowIndex + 1 : idx + 1}`,
            nav1Date: getRowDateValue(rowEl, 1),
            nav2Date: getRowDateValue(rowEl, 2),
            nav3Date: getRowDateValue(rowEl, 3),
          };
        })
        .filter(cycle => cycle.nav1Date && cycle.nav3Date);
      if (!chartEl || !summaryEl) return;
      if (!chartCycles.length) {
        summaryEl.textContent = 'เลือกวันที่ NAV (1) และ NAV (3) เพื่อแสดงกราฟ';
        chartEl.innerHTML = '<div class="state-box compact">ยังไม่มีช่วงวันที่สำหรับวาดกราฟ</div>';
        return;
      }
      const cycleBoundaryDates = chartCycles.flatMap(cycle => [cycle.nav1Date, cycle.nav3Date]).filter(Boolean).sort();
      const startDate = cycleBoundaryDates[0];
      const endDate = cycleBoundaryDates[cycleBoundaryDates.length - 1];
      const symbolsForChart = getChartSymbols();
      summaryEl.textContent = selectedChartMode === 'all'
        ? `ทุก Cycle · ช่วง ${formatShortDate(startDate)} ถึง ${formatShortDate(endDate)} · ${symbolsForChart.length} สินทรัพย์`
        : `${chartCycles[0].label} · ช่วง ${formatShortDate(startDate)} ถึง ${formatShortDate(endDate)} · ${symbolsForChart.length} สินทรัพย์`;
      chartEl.innerHTML = '<div class="state-box compact">กำลังวาดกราฟ...</div>';
      try {
        const series = [];
        for (const symbol of symbolsForChart) {
          const priceRows = await getSymbolPriceRows(symbol);
          const points = priceRows
            .filter(row => row.date >= startDate && row.date <= endDate && Number.isFinite(Number(row.close)))
            .map(row => ({ date: row.date, close: Number(row.close) }));
          if (points.length) series.push({ symbol, label: symbolLabel(symbol), points });
        }
        const markers = chartCycles.flatMap(cycle => [
          { label: `${cycle.label} NAV (1)`, date: cycle.nav1Date },
          { label: `${cycle.label} NAV (2)`, date: cycle.nav2Date },
          { label: `${cycle.label} NAV (3)`, date: cycle.nav3Date },
        ]);
        const highlightRanges = chartCycles.flatMap(cycle => [
          { from: cycle.nav1Date, to: cycle.nav2Date, className: 'uc-chart-highlight-downside' },
          { from: cycle.nav2Date, to: cycle.nav3Date, className: 'uc-chart-highlight-upside' },
        ]);
        chartEl.innerHTML = buildNavChartSvg(
          series,
          selectedChartMode === 'all'
            ? markers
            : [
              { label: 'NAV (1)', date: chartCycles[0].nav1Date },
              { label: 'NAV (2)', date: chartCycles[0].nav2Date },
              { label: 'NAV (3)', date: chartCycles[0].nav3Date },
            ],
          startDate,
          endDate,
          highlightRanges,
        );
      } catch (err) {
        chartEl.innerHTML = `<div class="state-box compact">${esc(err.message || 'วาดกราฟไม่สำเร็จ')}</div>`;
      }
    };

    const parseNavValue = (rowEl, point) => {
      const raw = $(`.uc-nav[data-point="${point}"]`, rowEl)?.value || '';
      const value = Number(raw.replace(/,/g, ''));
      return Number.isFinite(value) && value > 0 ? value : null;
    };

    const setCalcText = (rowEl, selector, value, formatter = 'pct') => {
      const el = $(selector, rowEl);
      if (!el) return;
      if (!Number.isFinite(value)) {
        el.textContent = '-';
        return;
      }
      el.textContent = formatter === 'ratio' ? value.toFixed(2) : pct(value);
    };

    const updateCaptureCalculations = (rowEl) => {
      const nav1 = parseNavValue(rowEl, 1);
      const nav2 = parseNavValue(rowEl, 2);
      const nav3 = parseNavValue(rowEl, 3);
      const downsideGap = nav1 && nav2 ? (nav2 / nav1) - 1 : NaN;
      const upsideGap = nav2 && nav3 ? (nav3 / nav2) - 1 : NaN;
      const totalCycleFund = nav1 && nav3 ? (nav3 / nav1) - 1 : NaN;
      const downsideCapture = Number.isFinite(downsideGap) ? 1 : NaN;
      const upsideCapture = Number.isFinite(upsideGap) ? 1 : NaN;
      const robustnessRatio = Number.isFinite(upsideCapture) && Number.isFinite(downsideCapture) && downsideCapture !== 0
        ? upsideCapture / downsideCapture
        : NaN;

      setCalcText(rowEl, '.uc-downside-gap', downsideGap);
      setCalcText(rowEl, '.uc-downside-capture', downsideCapture, 'ratio');
      setCalcText(rowEl, '.uc-upside-gap', upsideGap);
      setCalcText(rowEl, '.uc-upside-capture', upsideCapture, 'ratio');
      setCalcText(rowEl, '.uc-robustness-ratio', robustnessRatio, 'ratio');
      setCalcText(rowEl, '.uc-total-cycle-fund', totalCycleFund);
      setCalcText(rowEl, '.uc-total-cycle-benchmark', NaN);
      setCalcText(rowEl, '.uc-robustness-score', NaN);
    };

    const updateCaptureCell = async (rowEl, point, symbolOverride = '') => {
      const symbol = symbolOverride || $('#uc-benchmark-select', area)?.value || defaultSymbol;
      const dateValue = $(`.uc-date[data-point="${point}"]`, rowEl)?.value || '';
      const navInput = $(`.uc-nav[data-point="${point}"]`, rowEl);
      const actualEl = $(`.uc-actual-date[data-point="${point}"]`, rowEl);
      if (!navInput || !actualEl) return;
      navInput.value = '';
      actualEl.textContent = '';
      updateCaptureCalculations(rowEl);
      if (!symbol || !dateValue) return;
      navInput.placeholder = '...';
      try {
        const price = await lookupClose(symbol, dateValue);
        if (!price) {
          navInput.placeholder = '-';
          actualEl.textContent = 'ไม่มีข้อมูลก่อนวันที่นี้';
          updateCaptureCalculations(rowEl);
          return;
        }
        navInput.value = num(price.close, 3);
        navInput.placeholder = '-';
        actualEl.textContent = price.isExactDate ? '' : `ใช้ ${price.date}`;
        updateCaptureCalculations(rowEl);
      } catch (err) {
        navInput.placeholder = '-';
        actualEl.textContent = err.message || 'โหลดไม่ได้';
        updateCaptureCalculations(rowEl);
      }
    };

    const getBenchmarkDateRows = () => $$('#uc-cycle-body tr[data-uc-row]', area).map(rowEl => (
      [1, 2, 3].map(point => $(`.uc-date[data-point="${point}"]`, rowEl)?.value || '')
    ));

    const refreshComparisonLabels = (blockEl) => {
      const symbol = $('.uc-compare-symbol', blockEl)?.value || '-';
      $$('.uc-compare-cell', blockEl).forEach(cell => {
        cell.textContent = symbolLabel(symbol);
      });
    };

    const updateComparisonBlock = (blockEl) => {
      const symbol = $('.uc-compare-symbol', blockEl)?.value || defaultSymbol;
      refreshComparisonLabels(blockEl);
      $$('.uc-compare-body tr[data-uc-compare-row]', blockEl).forEach(rowEl => {
        [1, 2, 3].forEach(point => updateCaptureCell(rowEl, String(point), symbol));
        updateCaptureCalculations(rowEl);
      });
    };

    const syncComparisonBlocks = () => {
      const dateRows = getBenchmarkDateRows();
      $$('.uc-compare-block', area).forEach(blockEl => {
        const body = $('.uc-compare-body', blockEl);
        if (!body) return;
        body.innerHTML = dateRows.map((dates, rowIdx) => captureCompareRowHtml(rowIdx, dates)).join('');
        updateComparisonBlock(blockEl);
      });
    };

    const compareBlockHtml = (compareIdx) => {
      const dateRows = getBenchmarkDateRows();
      return `
        <section class="uc-compare-block" data-uc-compare-id="${compareIdx}">
          <div class="uc-compare-header">
            <label class="form-field">
              <span>สินทรัพย์เปรียบเทียบ ${compareIdx}</span>
              <select class="uc-compare-symbol">
                ${symbolSelectOptions(defaultSymbol)}
              </select>
            </label>
            <button class="btn btn-ghost uc-remove-compare" type="button">ลบสินทรัพย์</button>
          </div>
          <div class="table-wrapper uc-table-wrapper">
            <table class="uc-table">
              <thead>
                <tr>
                  <th rowspan="2">สินทรัพย์เปรียบเทียบ</th>
                  <th colspan="2">Peak</th>
                  <th colspan="2">Bottom</th>
                  <th colspan="2">Peak</th>
                  <th rowspan="2">Downside Gap<br>[(2)-(1)]/(1)</th>
                  <th rowspan="2">Downside Capture</th>
                  <th rowspan="2">Upside Gap<br>[(3)-(2)]/(2)</th>
                  <th rowspan="2">Upside Capture</th>
                  <th rowspan="2">Robustness Ratio</th>
                  <th rowspan="2">Total Cycle Return<br>(Fund)</th>
                  <th rowspan="2">Total Cycle Return<br>(Benchmark)</th>
                  <th rowspan="2">Robustness Score<br>(Fund/Benchmark)</th>
                  <th rowspan="2">จัดการ</th>
                </tr>
                <tr>
                  <th>NAV (1)</th>
                  <th>Date</th>
                  <th>NAV (2)</th>
                  <th>Date</th>
                  <th>NAV (3)</th>
                  <th>Date</th>
                </tr>
              </thead>
              <tbody class="uc-compare-body">
                ${dateRows.map((dates, rowIdx) => captureCompareRowHtml(rowIdx, dates)).join('')}
              </tbody>
            </table>
          </div>
        </section>
      `;
    };

    const refreshBenchmarkLabels = () => {
      const symbol = $('#uc-benchmark-select', area)?.value || defaultSymbol || '-';
      $$('.uc-benchmark-cell', area).forEach(cell => {
        cell.textContent = symbolLabel(symbol);
      });
    };

    $('#uc-cycle-body', area)?.addEventListener('change', async event => {
      const input = event.target.closest('.uc-date');
      if (!input) return;
      const rowEl = input.closest('tr');
      if (rowEl) {
        await updateCaptureCell(rowEl, input.dataset.point);
        syncComparisonBlocks();
        renderNavChart();
      }
    });

    $('#uc-benchmark-select', area)?.addEventListener('change', () => {
      refreshBenchmarkLabels();
      $$('#uc-cycle-body tr[data-uc-row]', area).forEach(rowEl => {
        [1, 2, 3].forEach(point => updateCaptureCell(rowEl, String(point)));
        updateCaptureCalculations(rowEl);
      });
      syncComparisonBlocks();
      renderNavChart();
    });

    $('#uc-cycle-body', area)?.addEventListener('click', event => {
      const deleteBtn = event.target.closest('.uc-delete-row');
      if (!deleteBtn) return;
      const body = $('#uc-cycle-body', area);
      const rowEl = deleteBtn.closest('tr[data-uc-row]');
      if (!body || !rowEl) return;
      if (body.querySelectorAll('tr[data-uc-row]').length <= 1) {
        toast('ต้องเหลืออย่างน้อย 1 แถว', 'warning', 2200);
        return;
      }
      rowEl.remove();
      refreshChartRowOptions();
      syncComparisonBlocks();
      renderNavChart();
    });

    $('#uc-add-row', area)?.addEventListener('click', () => {
      const body = $('#uc-cycle-body', area);
      if (!body) return;
      const rowIdx = body.querySelectorAll('tr[data-uc-row]').length;
      const addRow = $('.uc-add-row', body);
      if (addRow) {
        addRow.insertAdjacentHTML('beforebegin', captureRowHtml(rowIdx));
      } else {
        body.insertAdjacentHTML('beforeend', captureRowHtml(rowIdx));
      }
      refreshBenchmarkLabels();
      refreshChartRowOptions();
      syncComparisonBlocks();
      renderNavChart();
    });

    $('#uc-add-compare-asset', area)?.addEventListener('click', () => {
      const container = $('#uc-compare-assets', area);
      if (!container) return;
      const compareIdx = container.querySelectorAll('.uc-compare-block').length + 1;
      container.insertAdjacentHTML('beforeend', compareBlockHtml(compareIdx));
      const blockEl = container.lastElementChild;
      if (blockEl) updateComparisonBlock(blockEl);
      renderNavChart();
    });

    $('#uc-compare-assets', area)?.addEventListener('change', event => {
      const select = event.target.closest('.uc-compare-symbol');
      if (!select) return;
      const blockEl = select.closest('.uc-compare-block');
      if (blockEl) updateComparisonBlock(blockEl);
      renderNavChart();
    });

    $('#uc-compare-assets', area)?.addEventListener('click', event => {
      const removeBtn = event.target.closest('.uc-remove-compare');
      if (!removeBtn) return;
      const blockEl = removeBtn.closest('.uc-compare-block');
      if (blockEl) blockEl.remove();
      renderNavChart();
    });

    $('#uc-chart-row-select', area)?.addEventListener('change', () => {
      renderNavChart();
    });

    refreshChartRowOptions();
    renderNavChart();
    State._pageDataSource[pageKey] = `SQLite: ${payload.source || 'ft_historical_prices.sqlite'}`;
    App._currentExport = null;
    App._currentTableExport = null;
  },

  /* ─────────────────────────────────────────────────────────
     TOP 10 HOLDING V3  –  Multi-Fund Comparison Dashboard
     ───────────────────────────────────────────────────────── */
  async masterMenu02V3(area, pageKey = 'master-placeholder-8') {
    const isLocalFtPage = pageKey === 'ft-top10-holding';
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
      setError(area, e.message, pageKey);
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

    let ftSymbolItems = [];
    try {
      const payload = await loadFtHistoricalPayloadForPage(1);
      ftSymbolItems = (payload.symbols || [])
        .filter(item => {
          if (!item.symbol) return false;
          if (Object.prototype.hasOwnProperty.call(item, 'profileRowCount')) {
            return Number(item.profileRowCount || 0) > 0 || Number(item.holdingsRowCount || 0) > 0;
          }
          return Boolean(item.displayName || item.displayLabel || item.symbol);
        })
        .map(item => ({
          symbol: String(item.symbol || '').trim(),
          label: String(item.displayLabel || item.symbol || '').trim(),
          displayName: String(item.displayName || '').trim(),
          profileRowCount: Object.prototype.hasOwnProperty.call(item, 'profileRowCount')
            ? Number(item.profileRowCount || 0)
            : null,
          riskRowCount: Number(item.riskRowCount || 0),
          holdingsRowCount: Number(item.holdingsRowCount || 0),
        }));
    } catch (_) {}
    const ftSymbolMeta = new Map(ftSymbolItems.map(item => [item.symbol, item]));

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

    // ── Default selections from existing universe ──
    const defaultThaiCodes = Array.isArray(savedV3State?.selections?.thai) && savedV3State.selections.thai.length
      ? savedV3State.selections.thai.slice(0, 7)
      : thaiCodes.slice(0, 7).map(item => item.code);
    while (defaultThaiCodes.length < 7) defaultThaiCodes.push('');
    const savedFtSymbol = String(savedV3State?.selections?.iShare || '').trim();
    const selectedFtSymbol = ftSymbolMeta.has(savedFtSymbol) ? savedFtSymbol : '';
    const ftSymbolOptions = ftSymbolItems.length
      ? ftSymbolItems.map(item => {
        const riskLabel = item.riskRowCount ? ` · risk ${item.riskRowCount}` : '';
          const profileLabel = item.profileRowCount == null ? '' : ` · profile ${item.profileRowCount}`;
          return `<option value="${esc(item.symbol)}"${item.symbol === selectedFtSymbol ? ' selected' : ''}>${esc(item.label)}${profileLabel}${riskLabel}</option>`;
        }).join('')
      : '<option value="">ยังไม่มีข้อมูล FT qualitative ในเครื่อง</option>';

    // ── Selector Card HTML builder ──
    const selectorCard = (idx, color, label, inputHtml) => `
      <div style="background:#fff;border-radius:12px;border:1px solid var(--border);border-top:3px solid ${color};padding:12px 14px;display:flex;flex-direction:column;gap:6px;">
        <div style="font-size:0.88rem;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.04em;">${label}</div>
        ${inputHtml}
        <div id="v3-card-info-${idx}" style="font-size:0.84rem;color:var(--text-muted);min-height:18px;"></div>
      </div>`;

    const iShareCard = selectorCard(0, fundColors[0], 'iShare Index',
      `<select id="v3-sel-0" style="width:100%;padding:8px 10px;border-radius:6px;border:1px solid var(--border);font-size:0.92rem;font-weight:600;color:var(--primary-dark);background:#fff;cursor:pointer;">
        <option value="">-- เลือกจาก FT qualitative --</option>
        ${ftSymbolOptions}
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

      ${pageToolActions(pageKey, CONFIG.PAGES[pageKey]?.source || 'Multi-Fund Compare API')}

      <div class="card thv2-card" id="report-card">
        <div class="thv2-wrap">

          <!-- Fund Selectors (2×4 grid) -->
          <section class="thv2-panel">
            <div class="thv2-panel-head">
              <div>
                <h3>เลือกกองทุนที่ต้องการเปรียบเทียบ</h3>
                <p>${isLocalFtPage ? 'อ่านข้อมูลจาก FT qualitative local snapshot โดยอัตโนมัติเมื่อเลือกกองทุน' : 'ช่องที่ 1: iShare Index (Dropdown) &nbsp;|&nbsp; ช่องที่ 2-8: กองทุนไทย (พิมพ์หรือเลือกจาก Dropdown)'}</p>
              </div>
              <div style="display:flex;flex-direction:column;align-items:flex-end;gap:8px;">
                ${isLocalFtPage ? '' : `<button id="v3-load-btn" class="btn btn-primary" type="button" style="white-space:nowrap;">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
                  โหลดข้อมูลเปรียบเทียบ
                </button>`}
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

            <section class="thv2-panel" id="v3-ft-risk-section" style="display:none;">
              <div class="thv2-panel-head">
                <div>
                  <h3>FT Risk Measures</h3>
                  <p>ข้อมูล risk/profile จาก FT.com qualitative snapshot ใน SQLite</p>
                </div>
              </div>
              <div style="overflow-x:auto;">
                <table style="width:100%;border-collapse:collapse;font-size:0.92rem;">
                  <thead id="v3-ft-risk-thead"></thead>
                  <tbody id="v3-ft-risk-tbody"></tbody>
                </table>
              </div>
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
      const allSectionIds = ['v3-master-info-section','v3-holdings-section','v3-holdings-structure-section','v3-ft-risk-section'];
      const visibleSectionIds = isLocalFtPage
        ? ['v3-master-info-section','v3-holdings-section','v3-holdings-structure-section','v3-ft-risk-section']
        : allSectionIds;
      allSectionIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = visibleSectionIds.includes(id) ? '' : 'none';
      });

      // Master Fund Info Table — transposed (funds = columns, fields = rows)
      const masterThead = document.getElementById('v3-master-thead');
      const masterTbody = document.getElementById('v3-master-tbody');
      const dash = '<span style="color:#cbd5e1;">—</span>';
      const fmtPct = v => { if (v == null) return dash; const n = parseFloat(String(v).replace(/%/g,'')); return isNaN(n) ? dash : n.toFixed(2) + '%'; };
      const fmtStr = v => v ? esc(v) : dash;
      const fmtIsin = v => v ? `<span style="font-family:monospace;font-size:0.86rem;color:#1a3c6e;">${esc(v)}</span>` : dash;
      const ftProfile = (f, field) => f?._ftProfileMap?.[field] || '';
      const firstFtProfile = (f, fields) => {
        for (const field of fields) {
          const value = ftProfile(f, field);
          if (value) return value;
        }
        return '';
      };

      if (masterThead) masterThead.innerHTML = `
        <tr style="background:#1a3c6e;color:#fff;font-size:0.86rem;">
          <th style="padding:10px 14px;min-width:130px;border-right:1px solid rgba(255,255,255,0.15);background:transparent;"></th>
          ${data.map(f => `<th style="padding:10px 10px;text-align:center;font-weight:700;min-width:90px;">
            <span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${f.color};margin-right:5px;vertical-align:middle;opacity:0.9;"></span>${esc((f._selectorName || f.name).split('–')[0].trim())}
          </th>`).join('')}
        </tr>`;

      const masterRows = [
        { label: 'ISIN Code',        fmt: f => fmtIsin(f._isin) },
        { label: 'FT Symbol',        fmt: f => fmtIsin(f._ftSymbol) },
        { label: 'FT Display Name',  fmt: f => fmtStr(f._ftDisplayName) },
        { label: 'Morningstar Category', fmt: f => fmtStr(ftProfile(f, 'Morningstar category')) },
        { label: 'Investment Style', fmt: f => fmtStr(ftProfile(f, 'Investment style (stocks)')) },
        { label: 'Fund Size', fmt: f => fmtStr(firstFtProfile(f, ['Fund size', 'Total net assets'])) },
        { label: 'Share Class Size', fmt: f => fmtStr(firstFtProfile(f, ['Share class size'])) },
        { label: 'Ongoing Charge',   fmt: f => fmtPct(firstFtProfile(f, ['Ongoing charge', 'Net expense ratio']) || f._feesRaw) },
        { label: 'Initial Charge',   fmt: f => fmtPct(firstFtProfile(f, ['Initial charge', 'Front end load']) || f._feesInitial) },
        { label: 'Max Annual Charge',fmt: f => fmtPct(firstFtProfile(f, ['Max annual charge']) || f._feesMaxAnnual) },
        { label: 'Exit Charge',      fmt: f => fmtPct(firstFtProfile(f, ['Exit charge']) || f._feesExit) },
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
                ${esc(name)}${wt ? ` <span style="font-size:0.88rem;color:var(--text-muted);font-weight:500;">${esc(wt)}</span>` : ''}
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

      if (holdingsDatasets.length) {
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
      }

      const top10Legend = document.getElementById('v3-top10-asset-legend');
      if (top10Legend) top10Legend.innerHTML = holdingsDatasets.map(ds => `
        <div style="display:flex;align-items:center;gap:6px;font-size:0.88rem;color:var(--text-muted);max-width:320px;">
          <span style="width:12px;height:12px;border-radius:3px;background:${ds.backgroundColor};display:inline-block;flex-shrink:0;"></span>
          <span style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${esc(ds.label)}</span>
        </div>`).join('');

      const riskThead = document.getElementById('v3-ft-risk-thead');
      const riskTbody = document.getElementById('v3-ft-risk-tbody');
      if (riskThead) riskThead.innerHTML = `
        <tr style="background:#1a3c6e;color:#fff;font-size:0.86rem;">
          <th style="padding:10px 14px;min-width:180px;border-right:1px solid rgba(255,255,255,0.15);background:transparent;text-align:left;">Metric</th>
          ${data.map(f => `<th style="padding:10px 10px;text-align:center;font-weight:700;min-width:90px;">
            <span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${f.color};margin-right:5px;vertical-align:middle;opacity:0.9;"></span>${esc((f._selectorName || f.name).split('–')[0].trim())}
          </th>`).join('')}
        </tr>`;
      const riskMetrics = [
        ['Sharpe ratio', '1 year'],
        ['Sharpe ratio', '3 year'],
        ['Sharpe ratio', '5 years'],
        ['Standard deviation', '1 year'],
        ['Standard deviation', '3 year'],
        ['Standard deviation', '5 years'],
        ['Alpha', '3 year'],
        ['Beta', '3 year'],
      ];
      const riskValue = (f, metric, period) => {
        const found = (f._ftRisk || []).find(row => row.metric === metric && row.period === period);
        return found?.fundValue || '';
      };
      const anyRisk = data.some(f => Array.isArray(f._ftRisk) && f._ftRisk.length);
      if (riskTbody) riskTbody.innerHTML = anyRisk
        ? riskMetrics.map(([metric, period], ri) => `
          <tr style="border-bottom:1px solid var(--border-light);${ri % 2 === 0 ? '' : 'background:#fafbfd;'}">
            <td style="padding:9px 14px;font-weight:700;font-size:0.9rem;color:var(--text-muted);background:#f8fafc;white-space:nowrap;border-right:1px solid var(--border-light);">${esc(metric)} · ${esc(period)}</td>
            ${data.map(f => `<td style="padding:9px 10px;text-align:center;font-size:0.9rem;color:var(--text);">${fmtStr(riskValue(f, metric, period))}</td>`).join('')}
          </tr>`).join('')
        : `<tr><td colspan="${data.length + 1}" style="padding:14px;text-align:center;color:var(--text-muted);">ยังไม่มี FT risk measures ใน SQLite สำหรับกองที่เลือก</td></tr>`;

      latestExportPayload = buildTop10V3ExportPayload(
        CONFIG.PAGES[pageKey]?.title || 'Top 10 Holding',
        CONFIG.PAGES[pageKey]?.source || 'Multi-Fund Compare API',
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

    const loadFtQualitative = async (symbol) => {
      const cleanSymbol = String(symbol || '').trim();
      if (!cleanSymbol) return null;
      try {
        if (ftHistoricalApiUrl()) {
          const payload = await loadFtHistoricalDatabase();
          return getFtQualitativeFromPayload(payload, cleanSymbol);
        }
        const resp = await fetch(`/api/ft-historical-prices?symbol=${encodeURIComponent(cleanSymbol)}&limit=1`, { cache: 'no-store' });
        const payload = await resp.json().catch(() => ({}));
        if (!resp.ok || payload.ok === false) return null;
        const profileMap = Object.fromEntries((payload.profile || []).map(row => [row.field, row.value]));
        return {
          symbol: payload.selectedSymbol || cleanSymbol,
          displayName: payload.selectedDisplayName || profileMap['FT display name'] || '',
          profile: payload.profile || [],
          profileMap,
          risk: payload.risk || [],
          holdings: payload.holdings || [],
        };
      } catch {
        return null;
      }
    };

    // ── Load button click ──
    const loadBtn  = document.getElementById('v3-load-btn');
    const statusWrap = document.getElementById('v3-api-status-wrap');
    const statusEl   = document.getElementById('v3-api-status');
    const explorerSec = document.getElementById('v3-explorer-section');
    const rawEl     = document.getElementById('v3-raw');
    const toggleBtn = document.getElementById('v3-toggle-raw');
    const buildFundListFromSelectors = () => {
      const fundList = [];
      const iShareTicker = area.querySelector('#v3-sel-0')?.value?.trim() || '';
      if (iShareTicker) {
        const ftMeta = ftSymbolMeta.get(iShareTicker);
        fundList.push({
          idx: 0,
          name: ftMeta?.displayName || ftMeta?.label || iShareTicker,
          isin: iShareTicker,
          apiId: iShareTicker.split(':')[0],
          color: fundColors[0],
        });
      }
      for (let fi = 1; fi <= 7; fi++) {
        const inp = area.querySelector(`#v3-sel-${fi}`);
        const code = inp?.value?.trim() || '';
        if (code && !code.startsWith('กองทุนไทย #')) {
          const lu = thaiLookup[code];
          const isin = lu?.isin || code;
          const shortName = lu?.name ? code + ' – ' + lu.name.substring(0, 18) : code;
          fundList.push({ idx: fi, name: shortName, isin, apiId: isin, color: fundColors[fi] });
        }
      }
      return fundList;
    };
    const updateCardInfo = (fundList, liveData) => {
      (fundList || []).forEach((f, i) => {
        const infoEl = area.querySelector(`#v3-card-info-${f.idx}`);
        if (!infoEl) return;
        const d = liveData?.[i];
        if (d?._ftSymbol || d?._apiOk) {
          infoEl.textContent = d._ftSymbol || d._isin || f.isin || '—';
          infoEl.style.color = '#1a3c6e';
          infoEl.style.fontWeight = '600';
        } else {
          infoEl.textContent = '(ข้อมูลจำลอง)';
          infoEl.style.color = '#94a3b8';
          infoEl.style.fontWeight = 'normal';
        }
      });
    };
    const buildLocalFtRow = (fund, ft) => ({
      name: ft?.displayName || fund.name,
      color: fund.color,
      _selectorName: fund.name,
      _isin: fund.isin,
      _feesRaw: null,
      _feesInitial: null,
      _feesMaxAnnual: null,
      _feesExit: null,
      _manager: null,
      _topHoldings: (ft?.holdings || []).map(item => ({
        name: item.holdingName || '',
        companyName: item.holdingName || '',
        symbol: item.holdingSymbol || '',
        weight: parseFloat(String(item.portfolioWeight || '').replace(/%/g, '').trim()),
        weightText: item.portfolioWeight || '',
        oneYearChange: item.oneYearChange || '',
      })).filter(item => item.name),
      _ftSymbol: ft?.symbol || '',
      _ftDisplayName: ft?.displayName || '',
      _ftProfile: ft?.profile || [],
      _ftProfileMap: ft?.profileMap || {},
      _ftRisk: ft?.risk || [],
    });
    const renderLocalFtData = async () => {
      const fundList = buildFundListFromSelectors();
      saveV3State({ selections: getSelections() });
      if (!fundList.length) {
        if (statusWrap) statusWrap.style.display = '';
        if (statusEl) {
          statusEl.style.background = '#fef3c7';
          statusEl.style.color = '#92400e';
          statusEl.textContent = 'กรุณาเลือกกองทุนอย่างน้อย 1 กอง';
        }
        return;
      }
      if (statusWrap) statusWrap.style.display = '';
      if (statusEl) {
        statusEl.style.background = 'var(--primary-faint)';
        statusEl.style.color = 'var(--primary)';
        statusEl.textContent = `กำลังอ่าน FT qualitative จาก ${ftHistoricalApiUrl() ? 'Google Drive JSON' : 'Local'} (${fundList.length} กองทุน)...`;
      }
      const ftCalls = await Promise.allSettled(fundList.map(f => loadFtQualitative(f.isin)));
      const liveData = fundList.map((fund, index) => {
        const ft = ftCalls[index]?.status === 'fulfilled' ? ftCalls[index].value : null;
        return buildLocalFtRow(fund, ft);
      });
      render(liveData);
      updateCardInfo(fundList, liveData);
      const ftOkCount = liveData.filter(item => item._ftProfile?.length).length;
      if (statusEl) {
        const allOk = ftOkCount === fundList.length;
        statusEl.style.background = allOk ? '#f0fdf4' : '#fef3c7';
        statusEl.style.color = allOk ? '#15803d' : '#92400e';
        statusEl.textContent = allOk
          ? `อ่านข้อมูล FT qualitative สำเร็จครบทุกกองทุน (${ftOkCount}/${fundList.length})`
          : `อ่านข้อมูล FT qualitative ได้ ${ftOkCount}/${fundList.length} กองทุน`;
      }
      saveV3State({
        selections: getSelections(),
        fundList,
        liveData,
        showExplorer: false,
        rawVisible: false,
        statusText: statusEl?.textContent || '',
        statusTone: ftOkCount === fundList.length ? 'success' : 'warning',
      });
    };

    if (toggleBtn && rawEl) {
      toggleBtn.addEventListener('click', () => {
        rawEl.style.display = rawEl.style.display === 'none' ? 'block' : 'none';
        saveV3State({ rawVisible: rawEl.style.display !== 'none' });
      });
    }

    if (isLocalFtPage) {
      area.querySelector('#v3-sel-0')?.addEventListener('change', () => {
        renderLocalFtData();
      });
      for (let si = 1; si <= 7; si++) {
        area.querySelector(`#v3-sel-${si}`)?.addEventListener('change', () => {
          renderLocalFtData();
        });
      }
      renderLocalFtData();
    }

    if (loadBtn) {
      loadBtn.addEventListener('click', async () => {

        // ── Build fund list from selectors ──
        saveV3State({ selections: getSelections() });
        const fundList = buildFundListFromSelectors();

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
          statusEl.textContent = `⏳ กำลังดึงข้อมูลจาก Code.gs API + FT qualitative (${fundList.length} กองทุน)...`;
        }
        saveV3State({
          statusText: `⏳ กำลังดึงข้อมูลจาก Code.gs API + FT qualitative (${fundList.length} กองทุน)...`,
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
              fetch(`${TOP_10_HOLDING_API_URL}?isin=${encodeURIComponent(f.apiId || f.isin)}&fields=${FIELDS}`)
                .then(r => r.json())
                .catch(() => ({ ok: false, error: 'network error' }))
            )
          );
          const ftCalls = await Promise.allSettled(fundList.map(f => loadFtQualitative(f.isin)));

          // ── Map API responses → dashboard format (fallback to demoData) ──
          const liveData = fundList.map((f, i) => {
            const base = demoData[i];
            const raw  = apiCalls[i];
            const ft = ftCalls[i]?.status === 'fulfilled' ? ftCalls[i].value : null;
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
            const ftFields = {
              _ftSymbol: ft?.symbol || '',
              _ftDisplayName: ft?.displayName || '',
              _ftProfile: ft?.profile || [],
              _ftProfileMap: ft?.profileMap || {},
              _ftRisk: ft?.risk || [],
            };
            if (!api) return { ...base, _selectorName: f.name, _isin: f.isin, _feesRaw: null, _feesInitial: null, _feesMaxAnnual: null, _feesExit: null, _manager: null, ...ftFields };

            // Name
            const fundName = ft?.displayName || api.summary?.fundName || f.name;

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
                     _manager:       api.manager?.name || null,
                     ...ftFields };
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
          const ftOkCount = ftCalls.filter(r => r.status === 'fulfilled' && r.value?.profile?.length).length;
          if (statusEl) {
            const allOk = okCount === fundList.length;
            statusEl.style.background = allOk ? '#f0fdf4' : '#fef3c7';
            statusEl.style.color      = allOk ? '#15803d' : '#92400e';
            statusEl.textContent      = allOk
              ? `✓ ดึงข้อมูลสำเร็จครบทุกกองทุน (${okCount}/${fundList.length}) · FT qualitative ${ftOkCount}/${fundList.length}`
              : `⚠️ ดึงข้อมูลสำเร็จ ${okCount}/${fundList.length} กองทุน · FT qualitative ${ftOkCount}/${fundList.length} — กองที่เหลือแสดงข้อมูลจำลองแทน`;
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
      if (isLocalFtPage) return;
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
    $('#btn-export-modal')?.classList.remove('hidden');
    $('#modal-overlay').classList.remove('hidden');
  },

  openHtml(title, html) {
    this._rows = null;
    $('#modal-title').textContent = title;
    $('#modal-body').innerHTML = html;
    $('#btn-export-modal')?.classList.add('hidden');
    $('#modal-overlay').classList.remove('hidden');
  },

  close() {
    $('#modal-overlay').classList.add('hidden');
    $('#btn-export-modal')?.classList.remove('hidden');
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

    /* Update nav active state */
    $$('.nav-item').forEach(el =>
      el.classList.toggle('active', el.dataset.page === page)
    );

    /* Update page title */
    const titles = {
      'dashboard':         { title: 'แดชบอร์ด', subtitle: '' },
      'select-fund':       { title: 'เลือกกองทุน', subtitle: '' },
      'data-import':       { title: 'เตรียมข้อมูล', subtitle: 'นำข้อมูล Raw เข้า Google Sheets ปลายทางตาม Quarter' },
      'sec-data-import':   { title: 'เตรียมข้อมูลจาก SEC', subtitle: 'เลือกชุดข้อมูล SEC API และออกแบบ data_preparation ก่อนเขียนลง Google Sheet' },
      'fund-data-manager': { title: 'แก้ไขข้อมูลกองทุน', subtitle: 'กำหนด Master Fund หลายกองพร้อม weight สำหรับกองทุนไทย' },
      'master-fund-data-manager': { title: 'แก้ไขข้อมูลกอง Master Fund', subtitle: 'เตรียมหน้าไว้สำหรับแก้ไขข้อมูล Master Fund' },
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
      'income-fund-1': { title: 'Income Fund', subtitle: 'แสดงเฉพาะกองทุน Dividend และ Auto Redeem จาก Fund Key Performance AVP' },
      'income-fund-2': { title: 'Income Fund 2', subtitle: 'Focus Group จากกองที่ติ๊กไว้ใน Income Fund พร้อมประวัติปันผล SEC/Finnomena' },
      'robustness-ft-import': { title: 'เตรียมข้อมูลจาก FT.com', subtitle: 'เตรียม Historical Prices สำหรับ Robustness Fund จาก FT Markets' },
      'ft-top10-holding': { title: 'Top 10 Holding', subtitle: 'อ่านข้อมูลจาก FT qualitative local snapshot เป็นหลัก' },
      'upside-downside-capture': { title: 'Upside Downside Capture', subtitle: 'เตรียมหน้าวิเคราะห์การจับ upside/downside ของกองทุนเทียบ benchmark' },
      'master-placeholder-9': { title: 'ปัจจัยประกอบ กองทุนตราสารหนี้', subtitle: 'Fund Size, Duration, Turnover และ Yield to Maturity' },
      'master-placeholder-10': { title: 'ปัจจัยประกอบ กองทุนตราสารหนี้', subtitle: 'Asset Allocation, Rating, Holdings และความเสี่ยง' },
      'master-placeholder-11': { title: 'ปัจจัยประกอบอื่นๆ 4', subtitle: 'Max DD vs Sortino' },
      'notes': { title: 'บันทึกข้อมูล', subtitle: 'ดราฟงานค้างและโหลดกลับมาทำต่อ' },
      'guide':             { title: 'คู่มือการใช้งาน', subtitle: '' },
      'fund-list-update':  { title: 'Fund List Update', subtitle: 'ติดตามการเพิ่ม นำออก และสลับตำแหน่ง Fund List ราย Quarter' },
      'fund-selection-logs': { title: 'Fund Selection Logs', subtitle: 'บันทึกเหตุผลการคัดเลือกกองทุนราย Quarter จาก Investment Committee' },
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
      case 'data-import':       Pages.dataImport(area);                     break;
      case 'sec-data-import':   Pages.secDataImport(area);                  break;
      case 'fund-data-manager': Pages.fundDataManager(area);                break;
      case 'master-fund-data-manager': Pages.placeholder(area, 'แก้ไขข้อมูลกอง Master Fund'); break;
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
      case 'income-fund-1': Pages.incomeFund(area, 'income-fund-1');        break;
      case 'income-fund-2': Pages.incomeFundFocus(area, 'income-fund-2');  break;
      case 'robustness-ft-import': Pages.robustnessFtImport(area);          break;
      case 'ft-top10-holding': Pages.masterMenu02V3(area, 'ft-top10-holding'); break;
      case 'upside-downside-capture': Pages.upsideDownsideCapture(area);    break;
      case 'master-placeholder-9': Pages.otherFactorsFixedIncomeTable(area); break;
      case 'master-placeholder-10': Pages.otherFactorsFixedIncomeAllocationTable(area); break;
      case 'master-placeholder-11': Pages.comingSoon(area);                 break;
      case 'notes':             Pages.notesPage(area);                      break;
      case 'guide':             Pages.guide(area);                          break;
      case 'fund-list-update':  Pages.fundListUpdate(area);                 break;
      case 'fund-selection-logs': Pages.fundSelectionLogs(area);             break;
      default:
        area.innerHTML = '<div class="card"><div class="state-box">ไม่พบหน้าที่ต้องการ</div></div>';
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

  /* ดึงรายชื่อ Quarter ที่พร้อมครบทั้ง Google Sheets และ Drive JSON */
  async detect() {
    const sel    = $('#quarter-select');
    const status = $('#quarter-status');
    if (!sel) return;

    status.className = 'quarter-status loading';
    sel.disabled = true;

    try {
      const quarters = await detectReadyQuarters();
      if (!quarters.length) {
        throw new Error('ไม่พบ Quarter ที่พร้อมครบทั้ง 4 Google Sheets และ 4 JSON files ใน Drive');
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
        ไม่มี Quarter พร้อมครบ</option>`;
      sel.disabled = true;
      State.currentQuarter = CONFIG.PAGES?.['select-fund']?.tabName || null;
      console.warn('Quarter auto-detect failed:', err.message);
      toast(err.message || 'ตรวจ Quarter ไม่สำเร็จ', 'error', 7000);
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
