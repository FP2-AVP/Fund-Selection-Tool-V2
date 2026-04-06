/* ============================================================
   Fund Selection Tool – FP2
   Main Application Logic
   ============================================================ */
'use strict';

/* ── DOM helpers ── */
const $  = (sel, ctx = document) => ctx.querySelector(sel);
const $$ = (sel, ctx = document) => [...ctx.querySelectorAll(sel)];

/* ── Application State ── */
const State = {
  page:         'dashboard',
  sortCol:      null,
  sortDir:      'asc',
  selectedKeys: new Set(),   // keys of selected rows in select-fund page
  _cache:       {},
  _compareRows: null,        // rows currently in comparison modal
};

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

/* ============================================================
   DATA CACHE
   ============================================================ */
async function fetchCached(sheetId, tabName) {
  const key = `${sheetId}::${tabName}`;
  const now = Date.now();
  if (State._cache[key] && now - State._cache[key].ts < CONFIG.CACHE_TTL) {
    return State._cache[key].data;
  }
  const data = await SheetsAPI.fetchSheetData(sheetId, tabName);
  State._cache[key] = { data, ts: now };
  return data;
}

function clearCache() {
  State._cache = {};
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
  const { selectable = false, selectedKeys = new Set() } = opts;

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
    html += `<th class="th-sortable" data-col="${i}">${esc(h)}</th>`;
  });
  html += '</tr></thead>';

  /* ── Body ── */
  html += '<tbody>';
  dataRows.forEach((row, ri) => {
    const key    = esc(String(row[0] ?? ri));
    const selCls = (selectable && selectedKeys.has(String(row[0] ?? ri))) ? 'row-selected' : '';
    html += `<tr data-ri="${ri}" data-key="${key}" class="${selCls}">`;
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
  const { selectable = false, onSelChange } = opts;

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
        tArea.innerHTML = buildTable(sorted, { selectable, selectedKeys: State.selectedKeys });
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

/* ============================================================
   PAGES
   ============================================================ */
const Pages = {

  /* ── DASHBOARD ── */
  async dashboard(area) {
    area.innerHTML = `
      <div class="stats-grid" id="stats-grid">
        ${Array(4).fill('<div class="stat-card"><div class="spinner" style="width:22px;height:22px;border-width:2px;margin:auto"></div></div>').join('')}
      </div>
      <div class="section-title">เข้าถึงข้อมูลได้เลย</div>
      <div class="quick-grid" id="quick-grid"></div>`;

    /* Quick links */
    const links = [
      { page: 'select-fund',       icon: quickIcon('check'), title: 'เลือกกองทุน',                sub: 'AVP Master Fund ID' },
      { page: 'thai-annualized',   icon: quickIcon('trend'),  title: 'กองทุนไทย Annualized',        sub: 'AVP Thai Fund for Quality' },
      { page: 'thai-calendar',     icon: quickIcon('cal'),    title: 'กองทุนไทย Calendar Year',      sub: 'AVP Thai Fund for Quality' },
      { page: 'master-annualized', icon: quickIcon('globe'),  title: 'Master Fund Annualized',       sub: 'AVP Master Fund ID' },
      { page: 'master-calendar',   icon: quickIcon('list'),   title: 'Master Fund Calendar Year',    sub: 'AVP Master Fund ID' },
      { page: 'guide',             icon: quickIcon('book'),   title: 'คู่มือการใช้งาน',               sub: 'ขั้นตอนการตั้งค่าและใช้งาน' },
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

    /* Fetch row counts async */
    const pageKeys = ['select-fund','thai-annualized','thai-calendar','master-annualized'];
    const labels   = ['กองทุนที่เลือกได้','กองทุนไทย Annualized','กองทุนไทย Calendar','Master Annualized'];
    const classes  = ['','c-accent','c-success','c-gold'];

    const results = await Promise.allSettled(
      pageKeys.map(k => {
        const p = CONFIG.PAGES[k];
        return fetchCached(p.sheetId, p.tabName);
      })
    );

    $('#stats-grid', area).innerHTML = results.map((r, i) => {
      const count = r.status === 'fulfilled' ? Math.max(0, r.value.length - 1) : '–';
      return `
        <div class="stat-card ${classes[i]}">
          <div class="stat-label">${labels[i]}</div>
          <div class="stat-value">${count}</div>
          <div class="stat-desc">รายการ</div>
        </div>`;
    }).join('');

    App._currentExport = null;
  },

  /* ── GENERIC TABLE ── */
  async genericTable(area, pageKey) {
    const cfg = CONFIG.PAGES[pageKey];
    setLoading(area, `กำลังโหลด ${cfg.title}...`);

    let rawRows;
    try {
      rawRows = await fetchCached(cfg.sheetId, cfg.tabName);
    } catch (e) {
      setError(area, e.message, pageKey);
      return;
    }

    State.sortCol = null;
    State.sortDir = 'asc';

    const render = (query = '') => {
      const filtered = filterRows(rawRows, query);
      const sorted   = State.sortCol !== null
        ? sortRows(filtered, State.sortCol, State.sortDir)
        : filtered;
      const count    = Math.max(0, sorted.length - 1);

      area.innerHTML = `
        <div class="card">
          <div class="card-header">
            <span class="card-title">${esc(cfg.title)}</span>
            <div class="filter-bar">
              <div class="search-wrap">
                <span class="s-icon">${searchIcon()}</span>
                <input class="search-input" id="tbl-search" type="text"
                  placeholder="ค้นหา..." value="${esc(query)}" autocomplete="off">
              </div>
              <span class="row-count-badge">${count} รายการ</span>
              <span class="badge badge-primary">${esc(cfg.source)}</span>
            </div>
          </div>
          <div id="tbl-area">${buildTable(sorted)}</div>
        </div>`;

      bindTable(area, () => filterRows(rawRows, $('#tbl-search', area)?.value.trim() ?? ''));

      const inp = $('#tbl-search', area);
      let timer;
      inp.addEventListener('input', () => {
        clearTimeout(timer);
        timer = setTimeout(() => render(inp.value.trim()), 280);
      });
    };

    render();
    App._currentExport = () => exportExcel(rawRows, cfg.title);
  },

  /* ── SELECT FUND ── */
  async selectFund(area) {
    const cfg = CONFIG.PAGES['select-fund'];
    setLoading(area, 'กำลังโหลดรายการกองทุน...');

    let rawRows;
    try {
      rawRows = await fetchCached(cfg.sheetId, cfg.tabName);
    } catch (e) {
      setError(area, e.message, 'select-fund');
      return;
    }

    State.selectedKeys = new Set();
    State.sortCol      = null;
    State.sortDir      = 'asc';

    const updateBar = () => {
      const n    = State.selectedKeys.size;
      const countEl = $('#sel-count', area);
      const cmpBtn  = $('#btn-compare', area);
      const expBtn  = $('#btn-export-sel', area);
      if (countEl) countEl.textContent = `เลือกแล้ว ${n} รายการ`;
      if (cmpBtn)  cmpBtn.disabled     = n < 2;
      if (expBtn)  expBtn.disabled     = n === 0;
    };

    const render = (query = '') => {
      const filtered = filterRows(rawRows, query);
      const sorted   = State.sortCol !== null
        ? sortRows(filtered, State.sortCol, State.sortDir)
        : filtered;
      const count    = Math.max(0, sorted.length - 1);

      area.innerHTML = `
        <div class="card">
          <div class="card-header">
            <span class="card-title">เลือกกองทุน</span>
            <div class="filter-bar">
              <div class="search-wrap">
                <span class="s-icon">${searchIcon()}</span>
                <input class="search-input" id="tbl-search" type="text"
                  placeholder="ค้นหากองทุน..." value="${esc(query)}" autocomplete="off">
              </div>
              <span class="row-count-badge">${count} รายการ</span>
              <span class="badge badge-primary">${esc(cfg.source)}</span>
            </div>
          </div>
          <div class="card-body" style="padding:10px 16px 0">
            <div class="sel-bar">
              <span class="sel-bar-count" id="sel-count">เลือกแล้ว ${State.selectedKeys.size} รายการ</span>
              <button class="btn btn-accent btn-sm" id="btn-compare"
                ${State.selectedKeys.size < 2 ? 'disabled' : ''}>
                ⚖ เปรียบเทียบ
              </button>
              <button class="btn btn-primary btn-sm" id="btn-export-sel"
                ${State.selectedKeys.size === 0 ? 'disabled' : ''}>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                Export ที่เลือก
              </button>
              <button class="btn btn-ghost btn-sm" id="btn-clear-sel">✕ ล้าง</button>
            </div>
          </div>
          <div id="tbl-area">${buildTable(sorted, { selectable: true, selectedKeys: State.selectedKeys })}</div>
        </div>`;

      bindTable(
        area,
        () => filterRows(rawRows, $('#tbl-search', area)?.value.trim() ?? ''),
        { selectable: true, onSelChange: updateBar }
      );

      const inp = $('#tbl-search', area);
      let timer;
      inp.addEventListener('input', () => {
        clearTimeout(timer);
        timer = setTimeout(() => render(inp.value.trim()), 280);
      });

      /* Compare button */
      $('#btn-compare', area)?.addEventListener('click', () => {
        const selData = rawRows.slice(1).filter(r =>
          State.selectedKeys.has(String(r[0] ?? ''))
        );
        if (selData.length < 2) { toast('กรุณาเลือกอย่างน้อย 2 รายการ', 'warning'); return; }
        Modal.open('เปรียบเทียบกองทุน', [rawRows[0], ...selData]);
      });

      /* Export selected */
      $('#btn-export-sel', area)?.addEventListener('click', () => {
        const selData = rawRows.slice(1).filter(r =>
          State.selectedKeys.has(String(r[0] ?? ''))
        );
        if (selData.length === 0) { toast('ยังไม่ได้เลือกกองทุน', 'warning'); return; }
        exportExcel([rawRows[0], ...selData], 'selected-funds');
      });

      /* Clear selection */
      $('#btn-clear-sel', area)?.addEventListener('click', () => {
        State.selectedKeys.clear();
        render(inp?.value.trim() ?? '');
      });
    };

    render();
    App._currentExport = () => exportExcel(rawRows, 'fund-list');
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
  _toastTimer:    null,

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

    /* Refresh */
    $('#btn-refresh').addEventListener('click', () => {
      clearCache();
      App.navigate(State.page);
      toast('รีเฟรชข้อมูลแล้ว', 'success');
    });

    /* Export */
    $('#btn-export').addEventListener('click', () => {
      if (App._currentExport) App._currentExport();
      else toast('ไม่มีข้อมูลสำหรับ Export ในหน้านี้', 'warning');
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
    App._currentExport = null;

    /* Update nav active state */
    $$('.nav-item').forEach(el =>
      el.classList.toggle('active', el.dataset.page === page)
    );

    /* Update page title */
    const titles = {
      'dashboard':         'แดชบอร์ด',
      'select-fund':       'เลือกกองทุน',
      'thai-annualized':   'กองทุนไทย Annualized Return',
      'thai-calendar':     'กองทุนไทย Calendar Year',
      'master-annualized': 'Master Fund Annualized Return',
      'master-calendar':   'Master Fund Calendar Year',
      'guide':             'คู่มือการใช้งาน',
    };
    $('#page-title').textContent = titles[page] || page;

    const area = $('#content-area');

    switch (page) {
      case 'dashboard':         Pages.dashboard(area);                      break;
      case 'select-fund':       Pages.selectFund(area);                     break;
      case 'thai-annualized':   Pages.genericTable(area, 'thai-annualized');   break;
      case 'thai-calendar':     Pages.genericTable(area, 'thai-calendar');     break;
      case 'master-annualized': Pages.genericTable(area, 'master-annualized'); break;
      case 'master-calendar':   Pages.genericTable(area, 'master-calendar');   break;
      case 'guide':             Pages.guide(area);                          break;
      default:
        area.innerHTML = '<div class="card"><div class="state-box">ไม่พบหน้าที่ต้องการ</div></div>';
    }
  },
};

/* ============================================================
   BOOT – Login & Auth
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => {

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
