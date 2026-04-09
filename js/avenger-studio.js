'use strict';

// Fix: Fabric.js 5.x sets ctx.textBaseline = 'alphabetical' (typo).
// Chrome 112+ treats this as invalid → breaks text metrics → cascade crash.
// Patch the setter to silently normalise it to the correct 'alphabetic' value.
(function patchTextBaseline() {
  try {
    const desc = Object.getOwnPropertyDescriptor(CanvasRenderingContext2D.prototype, 'textBaseline');
    if (desc && desc.set) {
      Object.defineProperty(CanvasRenderingContext2D.prototype, 'textBaseline', {
        get: desc.get,
        set(v) { desc.set.call(this, v === 'alphabetical' ? 'alphabetic' : v); },
        configurable: true,
        enumerable: desc.enumerable,
      });
    }
  } catch {}
})();

(function initAvengerStudio(global) {
  const CW = 960;
  const CH = 540;
  const HDR_Y = 88;
  const FTR_Y = 497;
  const FONT = 'THSarabunNew';
  const STORAGE_KEY = 'avenger_slide_db';
  const LOGO_KEY = 'avenger_logo';
  const DB_NAME = 'avenger-studio-db';
  const DB_VERSION = 2;   // bumped to 2 so onupgradeneeded fires and creates all stores
  const SLIDES_STORE = 'slides';
  const QUEUE_STORE  = 'presentationQueue';

  let stylesInjected = false;
  let mounted = false;
  let root = null;
  let canvas = null;
  let canvasC = null;
  let canvasCReady = false;
  let currentDrawColor = '#1a3c6e';
  let currentEditIdx = null;
  let previewIdx = 0;
  let clipboard = null;
  let slides = [];
  let stateListener = null;
  let currentSlideKind = null;
  let dragIndex = null;

  // Safety patch: make fabric's stylesToArray null-safe.
  // In Fabric.js 5.3.1, stylesToArray crashes with TypeError when a style map has
  // undefined entries (common in old serialised textbox/i-text objects).
  // The function lives at different paths across Fabric versions, so we patch all of them.
  (function patchStylesToArray() {
    try {
      const safe = (orig) => function safeStylesToArray(styles, lineCount) {
        if (!styles) return [];
        try { return orig.call(this, styles, lineCount); } catch (e) {
          console.warn('[Studio] stylesToArray error suppressed:', e);
          return [];
        }
      };
      // Fabric.js 5.x: fabric.util.stylesToArray
      if (fabric.util && typeof fabric.util.stylesToArray === 'function') {
        fabric.util.stylesToArray = safe(fabric.util.stylesToArray);
      }
      // Fabric.js 4.x: fabric.util.object.stylesToArray
      if (fabric.util && fabric.util.object && typeof fabric.util.object.stylesToArray === 'function') {
        fabric.util.object.stylesToArray = safe(fabric.util.object.stylesToArray);
      }
      // Fabric.js: fabric.Object.stylesToArray (static)
      if (typeof fabric.Object.stylesToArray === 'function') {
        fabric.Object.stylesToArray = safe(fabric.Object.stylesToArray);
      }
    } catch (e) { /* non-fatal */ }
  })();

  function injectStyles() {
    if (stylesInjected) return;
    const style = document.createElement('style');
    style.textContent = `
      .studio-root { display:flex; flex-direction:column; height:100%; padding:12px; gap:12px; background:#f1f5f9; }
      .studio-editor-layout { display:flex; flex:1; gap:12px; min-height:0; }
      .studio-tab-content { display:none; flex:1; flex-direction:column; gap:12px; min-height:0; overflow:hidden; }
      .studio-tab-content.active { display:flex; }
      .studio-tab-nav { display:flex; gap:4px; background:#f1f5f9; padding:4px; border-radius:10px; border:1px solid #e2e8f0; }
      .studio-tab-btn { padding:8px 18px; border-radius:8px; border:none; font-family:'THSarabunNew',sans-serif; font-size:14px; font-weight:bold; cursor:pointer; color:#64748b; background:transparent; transition:all .2s; }
      .studio-tab-btn.active { background:white; color:#1a3c6e; box-shadow:0 2px 8px rgba(0,0,0,.08); }
      .studio-header { display:flex; justify-content:space-between; align-items:center; background:white; padding:10px 20px; border-radius:12px; box-shadow:0 2px 10px rgba(0,0,0,.05); border:1px solid #e2e8f0; flex-shrink:0; }
      .studio-canvas-wrapper { flex:1; background:#94a3b8; border-radius:12px; overflow:auto; display:flex; align-items:center; justify-content:center; padding:30px; }
      .studio-canvas-shell { box-shadow:0 10px 30px rgba(0,0,0,.25); flex-shrink:0; background:white; }
      .studio-panel { width:280px; background:white; border-radius:12px; box-shadow:0 4px 15px rgba(0,0,0,.05); padding:14px; display:flex; flex-direction:column; gap:10px; overflow-y:auto; }
      .studio-prop-section { padding:12px; background:#f8fafc; border-radius:10px; border:1px solid #e2e8f0; }
      .studio-section-title { font-size:11px; font-weight:bold; color:#94a3b8; text-transform:uppercase; letter-spacing:.5px; margin-bottom:10px; display:block; }
      .studio-tool-btn { width:100%; padding:10px; display:flex; align-items:center; gap:8px; border-radius:8px; transition:all .2s; color:#475569; background:#fff; border:1px solid #e2e8f0; cursor:pointer; font-size:14px; font-weight:bold; font-family:'THSarabunNew',sans-serif; }
      .studio-tool-btn:hover { background:#f1f5f9; border-color:#cbd5e1; transform:translateY(-1px); }
      .studio-tool-btn.active-mode { background:#eff6ff !important; color:#1a3c6e !important; border-color:#bfdbfe !important; }
      .studio-sidebar-label { font-size:12px; color:#64748b; font-weight:bold; display:block; margin-top:8px; margin-bottom:3px; }
      .studio-sidebar-input { width:100%; padding:7px 10px; border:1px solid #e2e8f0; border-radius:6px; font-family:'THSarabunNew',sans-serif; font-size:14px; color:#1e293b; background:white; }
      .studio-sidebar-divider { height:1px; background:#e2e8f0; flex-shrink:0; }
      .studio-palette-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:7px; margin-top:8px; }
      .studio-color-box { aspect-ratio:1; border-radius:5px; cursor:pointer; border:1px solid rgba(0,0,0,.1); transition:transform .1s; }
      .studio-color-box:hover { transform:scale(1.12); }
      .studio-color-box.selected { outline:3px solid #1a3c6e; outline-offset:2px; }
      .studio-control-disabled { opacity:.4; pointer-events:none; filter:grayscale(1); }
      .studio-library-section { background:white; border-radius:12px; padding:12px 16px; border:1px solid #e2e8f0; flex-shrink:0; }
      .studio-slides-scroll { display:flex; gap:14px; overflow-x:auto; padding:10px 5px 16px; min-height:115px; scroll-behavior:smooth; }
      .studio-slide-wrap { position:relative; flex-shrink:0; width:155px; }
      .studio-slide-wrap.dragging { opacity:.45; transform:scale(.97); }
      .studio-slide-item { width:100%; height:88px; background:#f8fafc; border-radius:8px; border:2px solid #e2e8f0; cursor:pointer; overflow:hidden; transition:all .2s; }
      .studio-slide-item.active { border-color:#1a3c6e; box-shadow:0 0 0 3px rgba(26,60,110,.12); transform:translateY(-3px); }
      .studio-slide-item img { width:100%; height:100%; object-fit:contain; pointer-events:none; }
      .studio-slide-del { position:absolute; top:-8px; right:-8px; width:24px; height:24px; border-radius:50%; background:#ef4444; border:2px solid white; color:white; font-size:15px; font-weight:bold; cursor:pointer; display:flex; align-items:center; justify-content:center; z-index:10; box-shadow:0 2px 6px rgba(0,0,0,.2); }
      .studio-slide-badge { position:absolute; bottom:-10px; left:50%; transform:translateX(-50%); background:#1a3c6e; color:white; font-size:10px; font-weight:bold; padding:1px 8px; border-radius:10px; font-family:'THSarabunNew',sans-serif; white-space:nowrap; }
      .studio-slide-kind { position:absolute; top:8px; left:8px; padding:2px 8px; border-radius:999px; font-size:10px; font-weight:bold; color:white; background:rgba(15,23,42,.78); letter-spacing:.04em; }
      .studio-slide-kind.kind-table { background:rgba(21,101,192,.86); }
      .studio-slide-kind.kind-image { background:rgba(22,101,52,.86); }
      .studio-preview-overlay { position:fixed; inset:0; background:rgba(0,0,0,.8); z-index:1000; display:none; align-items:center; justify-content:center; }
      .studio-preview-overlay.open { display:flex; }
      .studio-preview-modal { background:#1e293b; border-radius:16px; width:90%; max-width:1000px; max-height:90vh; display:flex; flex-direction:column; overflow:hidden; }
      .studio-preview-content { flex:1; background:white; overflow:auto; display:flex; justify-content:center; align-items:flex-start; padding:20px; }
      .studio-preview-footer { background:#0f172a; padding:14px; display:flex; align-items:center; gap:20px; color:white; }
      .studio-toast { position:fixed; bottom:24px; right:24px; background:#1e293b; color:white; padding:11px 18px; border-radius:10px; font-size:14px; font-weight:bold; font-family:'THSarabunNew',sans-serif; z-index:9999; opacity:0; transition:opacity .3s, transform .3s; pointer-events:none; transform:translateY(8px); box-shadow:0 4px 20px rgba(0,0,0,.25); }
      .studio-toast.show { opacity:1; transform:translateY(0); }
      .studio-logo-preview { margin-top:6px; background:#f1f5f9; border-radius:6px; padding:6px; display:flex; align-items:center; justify-content:center; min-height:52px; border:1px dashed #cbd5e1; }
      .studio-logo-preview img { max-height:44px; max-width:100%; object-fit:contain; }
      .studio-table-tools { display:none; }
      .studio-table-tools.active { display:block; }
    `;
    document.head.appendChild(style);
    stylesInjected = true;
  }

  function renderLayout() {
    return `
      <div class="studio-root">
        <div class="studio-header">
          <div style="display:flex; align-items:center; gap:12px">
            <div style="width:38px;height:38px;background:#1a3c6e;border-radius:9px;display:flex;align-items:center;justify-content:center;color:white;font-weight:bold;font-size:18px;font-family:'THSarabunNew',sans-serif">A</div>
            <div>
              <h1 style="font-size:16px;color:#1e293b;font-family:'THSarabunNew',sans-serif;font-weight:bold">Avenger Studio</h1>
              <p style="font-size:11px;color:#94a3b8;font-weight:bold;text-transform:uppercase;letter-spacing:.5px">Presentation Slide Designer</p>
            </div>
          </div>
          <div class="studio-tab-nav">
            <button class="studio-tab-btn active" data-studio-tab="studio" type="button">Studio</button>
            <button class="studio-tab-btn" data-studio-tab="canvas" type="button">Canvas</button>
          </div>
          <div style="display:flex; gap:10px; align-items:center">
            <div id="studio-edit-indicator" style="display:none;font-size:13px;color:#1a3c6e;background:#eff6ff;padding:5px 14px;border-radius:20px;border:1px solid #bfdbfe;font-weight:bold;font-family:'THSarabunNew',sans-serif">
              แก้ไขหน้า: <span id="studio-page-display">1</span>
            </div>
            <button id="studio-preview-btn" type="button" class="studio-tool-btn" style="width:auto;padding:8px 14px">Preview</button>
            <button id="studio-export-btn" type="button" class="studio-tool-btn" style="width:auto;padding:8px 18px;background:#1a3c6e;color:white;border:none">Export PDF</button>
          </div>
        </div>
        <div id="studio-tab-studio" class="studio-tab-content active">
          <div class="studio-editor-layout">
            <div class="studio-canvas-wrapper">
              <div id="studio-canvas-container" class="studio-canvas-shell"><canvas id="studio-canvas"></canvas></div>
            </div>
            <aside class="studio-panel">
              <button id="studio-new-page" type="button" class="studio-tool-btn" style="background:#eff6ff;color:#1a3c6e;border-color:#bfdbfe;justify-content:center">สร้างหน้าใหม่</button>
              <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px">
                <button id="studio-add-rect" type="button" class="studio-tool-btn" style="flex-direction:column;padding:12px;justify-content:center;align-items:center"><span style="font-size:18px">⬜</span><span style="font-size:12px">สี่เหลี่ยม</span></button>
                <button id="studio-add-circle" type="button" class="studio-tool-btn" style="flex-direction:column;padding:12px;justify-content:center;align-items:center"><span style="font-size:18px">◯</span><span style="font-size:12px">วงกลม</span></button>
                <button id="studio-add-line" type="button" class="studio-tool-btn" style="flex-direction:column;padding:12px;justify-content:center;align-items:center"><span style="font-size:18px">／</span><span style="font-size:12px">เส้นตรง</span></button>
              </div>
              <button id="studio-add-text" type="button" class="studio-tool-btn">เพิ่มข้อความ</button>
              <div class="studio-sidebar-divider"></div>
              <div class="studio-prop-section">
                <span class="studio-section-title">Template</span>
                <label class="studio-sidebar-label">ข้อมูล ณ :</label>
                <input class="studio-sidebar-input" type="text" id="studio-footer-date" placeholder="28/02/2026">
                <label class="studio-sidebar-label">ที่มา :</label>
                <input class="studio-sidebar-input" type="text" id="studio-footer-source" placeholder="Percent Rank">
                <button id="studio-apply-footer" type="button" class="studio-tool-btn" style="margin-top:10px;font-size:13px;background:#f0fdf4;color:#166534;border-color:#bbf7d0">อัพเดต Footer</button>
              </div>
              <div class="studio-sidebar-divider"></div>
              <div id="studio-table-tools" class="studio-prop-section studio-table-tools">
                <span class="studio-section-title">Table Adjust</span>
                <label class="studio-sidebar-label">ปรับทั้งตารางแบบเร็ว</label>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
                  <button id="studio-table-scale-down" type="button" class="studio-tool-btn" style="justify-content:center">ตารางเล็กลง</button>
                  <button id="studio-table-scale-up" type="button" class="studio-tool-btn" style="justify-content:center">ตารางใหญ่ขึ้น</button>
                </div>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px">
                  <button id="studio-table-font-down" type="button" class="studio-tool-btn" style="justify-content:center">ฟอนต์เล็กลง</button>
                  <button id="studio-table-font-up" type="button" class="studio-tool-btn" style="justify-content:center">ฟอนต์ใหญ่ขึ้น</button>
                </div>
                <div style="display:flex;justify-content:space-between;align-items:center;font-size:13px;margin-top:10px;color:#64748b">
                  <label>ขนาดตัวอักษรโดยประมาณ</label><span id="studio-table-font-estimate">-</span>
                </div>
                <label class="studio-sidebar-label" style="margin-top:10px">ขยับตาราง</label>
                <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;text-align:center">
                  <div></div>
                  <button id="studio-table-move-up"    type="button" class="studio-tool-btn" style="justify-content:center">▲</button>
                  <div></div>
                  <button id="studio-table-move-left"  type="button" class="studio-tool-btn" style="justify-content:center">◀</button>
                  <button id="studio-table-move-center" type="button" class="studio-tool-btn" style="justify-content:center;font-size:11px">กึ่งกลาง</button>
                  <button id="studio-table-move-right" type="button" class="studio-tool-btn" style="justify-content:center">▶</button>
                  <div></div>
                  <button id="studio-table-move-down"  type="button" class="studio-tool-btn" style="justify-content:center">▼</button>
                  <div></div>
                </div>
                <p style="font-size:12px;color:#64748b;line-height:1.4;margin-top:10px">ใช้สำหรับจูน Presentation ของตารางข้อมูลที่ส่งเข้ามา</p>
              </div>
              <div class="studio-sidebar-divider"></div>
              <div class="studio-prop-section">
                <span class="studio-section-title">Properties</span>
                <div id="studio-text-tools" class="studio-control-disabled">
                  <div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:4px">
                    <label>ขนาดตัวอักษร</label><span id="studio-font-val">28</span>
                  </div>
                  <input type="range" id="studio-font-slider" min="10" max="200" value="28">
                </div>
                <div id="studio-general-tools" class="studio-control-disabled" style="margin-top:12px">
                  <label style="font-size:13px;display:block;margin-bottom:6px">สีพื้นหลัง / ข้อความ</label>
                  <div class="studio-palette-grid" id="studio-palette"></div>
                  <div style="margin-top:12px">
                    <div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:4px">
                      <label>ความโปร่งใส</label><span id="studio-alpha-val">100%</span>
                    </div>
                    <input type="range" id="studio-alpha-slider" min="0" max="100" value="100">
                  </div>
                  <div style="margin-top:12px">
                    <label style="font-size:13px;display:block;margin-bottom:6px">เส้นขอบ</label>
                    <button id="studio-stroke-toggle" type="button" class="studio-tool-btn" style="justify-content:center">มีขอบ</button>
                    <div style="margin-top:8px">
                      <input type="color" id="studio-stroke-color" value="#94a3b8" style="width:100%;height:38px;border:1px solid #dbe4f0;border-radius:8px;background:#fff;padding:4px">
                    </div>
                  </div>
                  <div style="margin-top:12px">
                    <div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:6px">
                      <label>ความหนาเส้นขอบ</label><span id="studio-stroke-width-val">1 px</span>
                    </div>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
                      <button id="studio-stroke-width-down" type="button" class="studio-tool-btn" style="justify-content:center">ลดลง</button>
                      <button id="studio-stroke-width-up" type="button" class="studio-tool-btn" style="justify-content:center">เพิ่มขึ้น</button>
                    </div>
                  </div>
                </div>
              </div>
              <button id="studio-save-library" type="button" class="studio-tool-btn" style="margin-top:auto;background:#1a3c6e;color:white;border:none;justify-content:center;box-shadow:0 4px 12px rgba(26,60,110,.3);font-size:15px">บันทึกลงคลัง</button>
            </aside>
          </div>
          <div class="studio-library-section">
            <p style="font-size:11px;font-weight:bold;color:#94a3b8;text-transform:uppercase;margin-bottom:6px;letter-spacing:.5px">Library (รายการหน้าทั้งหมด)</p>
            <div id="studio-slides-grid" class="studio-slides-scroll"></div>
          </div>
        </div>
        <div id="studio-tab-canvas" class="studio-tab-content">
          <div class="studio-editor-layout">
            <div class="studio-canvas-wrapper">
              <div id="studio-canvas-container-c" class="studio-canvas-shell"><canvas id="studio-canvas-c"></canvas></div>
            </div>
            <aside class="studio-panel">
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
                <button id="studio-draw-mode" type="button" class="studio-tool-btn active-mode" style="flex-direction:column;padding:12px;justify-content:center;align-items:center">วาดอิสระ</button>
                <button id="studio-select-mode" type="button" class="studio-tool-btn" style="flex-direction:column;padding:12px;justify-content:center;align-items:center">เลือก/ย้าย</button>
              </div>
              <div class="studio-sidebar-divider"></div>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
                <button id="studio-add-rect-c" type="button" class="studio-tool-btn" style="flex-direction:column;padding:12px;justify-content:center;align-items:center">สี่เหลี่ยม</button>
                <button id="studio-add-circle-c" type="button" class="studio-tool-btn" style="flex-direction:column;padding:12px;justify-content:center;align-items:center">วงกลม</button>
              </div>
              <button id="studio-add-text-c" type="button" class="studio-tool-btn">เพิ่มข้อความ</button>
              <div class="studio-sidebar-divider"></div>
              <div class="studio-prop-section">
                <span class="studio-section-title">Canvas Properties</span>
                <div id="studio-c-brush-tools">
                  <div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:4px"><label>ขนาดพู่กัน</label><span id="studio-brush-val">3</span></div>
                  <input type="range" id="studio-brush-slider" min="1" max="50" value="3">
                </div>
                <div id="studio-c-text-tools" class="studio-control-disabled" style="margin-top:10px">
                  <div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:4px"><label>ขนาดตัวอักษร</label><span id="studio-c-font-val">28</span></div>
                  <input type="range" id="studio-c-font-slider" min="10" max="200" value="28">
                </div>
                <div style="margin-top:10px">
                  <label style="font-size:13px;display:block;margin-bottom:6px">สี</label>
                  <div class="studio-palette-grid" id="studio-c-palette"></div>
                </div>
                <div id="studio-c-obj-tools" class="studio-control-disabled" style="margin-top:10px">
                  <div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:4px"><label>ความโปร่งใส</label><span id="studio-c-alpha-val">100%</span></div>
                  <input type="range" id="studio-c-alpha-slider" min="0" max="100" value="100">
                </div>
              </div>
              <div class="studio-sidebar-divider"></div>
              <button id="studio-clear-canvas-c" type="button" class="studio-tool-btn" style="color:#ef4444;border-color:#fecaca">ล้าง Canvas</button>
              <button id="studio-copy-canvas-c" type="button" class="studio-tool-btn" style="background:#f0fdf4;color:#166534;border-color:#bbf7d0">คัดลอกรูป</button>
              <button id="studio-send-to-slide" type="button" class="studio-tool-btn" style="margin-top:auto;background:#1a3c6e;color:white;border:none;justify-content:center;box-shadow:0 4px 12px rgba(26,60,110,.3);font-size:15px">ใส่ใน Studio Slide</button>
            </aside>
          </div>
        </div>
        <div id="studio-preview-overlay" class="studio-preview-overlay">
          <div class="studio-preview-modal">
            <div style="padding:14px 18px;background:#0f172a;display:flex;justify-content:space-between;align-items:center;color:white">
              <span style="font-family:'THSarabunNew',sans-serif;font-size:16px">Preview Mode</span>
              <button id="studio-close-preview" type="button" style="background:none;border:none;color:white;cursor:pointer;font-size:22px">&times;</button>
            </div>
            <div class="studio-preview-content"><img id="studio-preview-img" alt="preview" style="max-width:100%;box-shadow:0 5px 15px rgba(0,0,0,.2)"></div>
            <div class="studio-preview-footer">
              <button id="studio-prev-preview" type="button" class="studio-tool-btn" style="width:auto">ก่อนหน้า</button>
              <span id="studio-preview-info" style="flex:1;text-align:center;font-family:'THSarabunNew',sans-serif">หน้า 1 / 1</span>
              <button id="studio-next-preview" type="button" class="studio-tool-btn" style="width:auto">ถัดไป</button>
            </div>
          </div>
        </div>
        <div id="studio-toast" class="studio-toast"></div>
      </div>`;
  }

  function $(sel) {
    return root?.querySelector(sel);
  }

  function showToast(msg) {
    const toast = $('#studio-toast');
    if (!toast) return;
    toast.textContent = msg;
    toast.classList.add('show');
    clearTimeout(showToast._timer);
    showToast._timer = setTimeout(() => toast.classList.remove('show'), 3000);
  }

  function openSlidesDb() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        // Create both stores so whichever file opens the DB first, all stores exist.
        if (!db.objectStoreNames.contains(SLIDES_STORE)) {
          db.createObjectStore(SLIDES_STORE, { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains(QUEUE_STORE)) {
          db.createObjectStore(QUEUE_STORE, { keyPath: 'id' });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function getStoredSlides() {
    const db = await openSlidesDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(SLIDES_STORE, 'readonly');
      const req = tx.objectStore(SLIDES_STORE).getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  }

  async function replaceStoredSlides(nextSlides) {
    const db = await openSlidesDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(SLIDES_STORE, 'readwrite');
      tx.objectStore(SLIDES_STORE).clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
    await new Promise((resolve, reject) => {
      const tx = db.transaction(SLIDES_STORE, 'readwrite');
      const store = tx.objectStore(SLIDES_STORE);
      nextSlides.forEach(slide => store.put(slide));
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function loadSlides() {
    try {
      const dbSlides = await getStoredSlides();
      if (dbSlides.length) {
        slides = dbSlides;
        return;
      }
      const legacySlides = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
      slides = Array.isArray(legacySlides) ? legacySlides : [];
      if (slides.length) {
        await replaceStoredSlides(slides);
      }
    } catch {
      slides = [];
    }
  }

  async function persistSlides() {
    await replaceStoredSlides(slides);
  }

  function emitState() {
    if (!stateListener) return;
    stateListener({
      slideCount: slides.length,
      currentEditIdx,
      previewIdx,
      activeTab: $('#studio-tab-canvas')?.classList.contains('active') ? 'canvas' : 'studio',
      processedQueueIds: [],
    });
  }

  async function loadFonts() {
    const defs = [
      { src: 'fonts/THSarabunNew.woff', weight: 'normal', style: 'normal' },
      { src: 'fonts/THSarabunNew-Bold.woff', weight: 'bold', style: 'normal' },
      { src: 'fonts/THSarabunNew-Italic.woff', weight: 'normal', style: 'italic' },
      { src: 'fonts/THSarabunNew-BoldItalic.woff', weight: 'bold', style: 'italic' },
    ];
    await Promise.all(defs.map(async def => {
      try {
        const face = new FontFace(FONT, `url(${def.src})`, { weight: def.weight, style: def.style });
        document.fonts.add(await face.load());
      } catch {
        /* noop */
      }
    }));
  }

  function buildImportedSlide(dataURL) {
    return {
      id: `studio-slide-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      kind: 'image',
      json: {
        version: fabric.version,
        objects: [{
          type: 'image',
          version: fabric.version,
          originX: 'left',
          originY: 'top',
          left: 0,
          top: 0,
          width: CW,
          height: CH,
          scaleX: 1,
          scaleY: 1,
          src: dataURL,
          selectable: true,
          evented: true,
        }],
      },
      preview: dataURL,
    };
  }

  function createRectObject(left, top, width, height, fill, extra = {}) {
    return { type: 'rect', version: fabric.version, left, top, width, height, fill, stroke: extra.stroke || '#d7dee9', strokeWidth: extra.strokeWidth ?? 1, selectable: extra.selectable ?? true, evented: extra.evented ?? true, id: extra.id };
  }

  // Non-blocking canvas.loadFromJSON replacement.
  // Processes objects in batches of CHUNK_SIZE with setTimeout(0) between batches so the
  // main thread yields periodically — prevents Chrome "Page Unresponsive" dialog on large
  // table slides (500-800 objects).
  function loadFromJSONChunked(targetCanvas, json, callback) {
    const serialized = (typeof json === 'string') ? JSON.parse(json) : json;
    targetCanvas.clear();
    if (serialized && serialized.background) {
      targetCanvas.setBackgroundColor(serialized.background, () => {});
    }
    const allObjects = (serialized && serialized.objects) ? serialized.objects : [];
    if (!allObjects.length) {
      targetCanvas.renderAll();
      if (callback) callback();
      return;
    }
    const CHUNK_SIZE = 25;
    let batchStart = 0;
    function processBatch() {
      const batch = allObjects.slice(batchStart, batchStart + CHUNK_SIZE);
      fabric.util.enlivenObjects(batch, (enlivened) => {
        enlivened.forEach(obj => { if (obj) targetCanvas.add(obj); });
        batchStart += CHUNK_SIZE;
        if (batchStart < allObjects.length) {
          setTimeout(processBatch, 0); // yield main thread between batches
        } else {
          targetCanvas.renderAll();
          if (callback) callback();
        }
      }, 'fabric');
    }
    processBatch();
  }

  function createTextObject(text, left, top, width, fontSize, color, extra = {}) {
    // Use 'text' (fabric.Text) instead of 'textbox' (fabric.Textbox) — Textbox has a
    // callSuper→toObject infinite-recursion bug in Fabric.js 5.3.1 that causes
    // "Maximum call stack size exceeded" when serialising or rendering many text objects.
    // Table cells are non-editable so the simpler Text class is sufficient.
    //
    // fabric.Text ignores `width` for layout, so centering must be done via originX:'center'
    // with `left` pointing to the horizontal midpoint of the cell.
    const align = extra.textAlign || 'center';
    const finalLeft = align === 'center' ? left + width / 2 : left + 4;
    const originX   = align === 'center' ? 'center' : 'left';
    const originY   = extra.originY || 'top';
    return {
      type: 'text',
      version: fabric.version,
      left: finalLeft,
      top,
      originX,
      originY,
      fontSize,
      fontFamily: FONT,
      fill: color,
      text: String(text ?? ''),
      fontWeight: extra.fontWeight || 'normal',
      textAlign: align,
      selectable: extra.selectable ?? true,
      evented: extra.evented ?? true,
      lineHeight: extra.lineHeight || 1.1,
      textBackgroundColor: extra.textBackgroundColor || '',
      cellBgId: extra.cellBgId || '',
      id: extra.id,
    };
  }

  // Migrate slide JSON created before the textbox→text fix so that old stored slides
  // can be serialised without triggering the Fabric.js callSuper stack-overflow bug.
  function migrateSlideJson(json) {
    if (!json || !Array.isArray(json.objects)) return json;
    let changed = false;
    const objects = json.objects.map(obj => {
      if (obj.type !== 'textbox' && obj.type !== 'i-text') return obj;
      changed = true;
      const align = obj.textAlign || 'center';
      // Strip `styles` entirely — old textbox objects carry per-character style maps
      // that are often malformed (undefined entries) causing stylesToArray → TypeError.
      const { styles: _discard, ...rest } = obj;
      const migrated = { ...rest, type: 'text' };
      // fabric.Text uses originX:'center' + left=midpoint to centre within a cell.
      // Old textbox objects stored left=cellLeft+padding, width=cellWidth-padding*2.
      if (align === 'center' && obj.width && obj.originX !== 'center') {
        migrated.left = obj.left + obj.width / 2;
        migrated.originX = 'center';
      }
      return migrated;
    });
    return changed ? { ...json, objects } : json;
  }

  function buildTableSlideBase(title, subtitle, source, pageNum, totalPages) {
    return [
      { ...createRectObject(0, 0, CW, CH, '#ffffff', { stroke: '#ffffff', selectable: false, evented: false }), id: 'tpl-bg' },
      { type: 'line', version: fabric.version, x1: 0, y1: HDR_Y, x2: CW, y2: HDR_Y, stroke: '#1a3c6e', strokeWidth: 2, selectable: false, evented: false, id: 'tpl-hdrline' },
      { ...createTextObject(title, 24, 16, 690, 34, '#1a3c6e', { fontWeight: 'bold', textAlign: 'left', selectable: false, evented: false }), id: 'tpl-title' },
      { ...createTextObject(subtitle || '', 24, 54, 420, 16, '#64748b', { textAlign: 'left', selectable: false, evented: false }), id: 'tpl-subtitle' },
      { type: 'line', version: fabric.version, x1: 0, y1: FTR_Y, x2: CW, y2: FTR_Y, stroke: '#94a3b8', strokeWidth: 1, selectable: false, evented: false, id: 'tpl-ftrline' },
      { ...createTextObject(`ที่มา : ${source || '-'}`, 24, FTR_Y + 10, 420, 15, '#64748b', { textAlign: 'left', selectable: false, evented: false }), id: 'tpl-ftrleft' },
      { ...createTextObject(`${pageNum}/${totalPages}`, CW - 70, FTR_Y + 8, 40, 18, '#1a3c6e', { fontWeight: 'bold', selectable: false, evented: false }), id: 'tpl-ftrright' },
    ];
  }

  // Max rows per slide — cap large tables to 3 slides when paired with the
  // sender-side 90-row limit, reducing the chance of browser hangs.
  const TABLE_ROWS_PER_SLIDE = 30;

  function estimateTextWidth(text, fontSize) {
    return Math.max(fontSize * 0.72, String(text || '').length * fontSize * 0.58);
  }

  function getMaxFontSizeForBounds(text, bg, widthPadding = 8, heightPadding = 6) {
    if (!bg) return 44;
    const safeWidth = Math.max(14, (bg.width || 0) - widthPadding);
    const safeHeight = Math.max(10, (bg.height || 0) - heightPadding);
    const textLength = Math.max(1, String(text || '').length);
    const widthCap = Math.floor(safeWidth / Math.max(0.58, textLength * 0.58));
    const heightCap = Math.floor(safeHeight / 1.2);
    return Math.max(7, Math.min(44, widthCap, heightCap));
  }

  function buildInlineCellTextObjects(cell, pageIndex, rowIndex, cellIndex, left, top, width, height, fontSize) {
    const fragments = Array.isArray(cell.fragments) ? cell.fragments : [];
    if (!fragments.length) {
      return [createTextObject(cell.text || '', left + 4, top + height / 2, width - 8, fontSize, cell.color || '#334155', {
        fontWeight: cell.strong ? 'bold' : 'normal',
        textAlign: cell.align === 'left' ? 'left' : 'center',
        originY: 'center',
        selectable: false,
        evented: false,
        id: `tbl-cell-text-${pageIndex}-${rowIndex}-${cellIndex}`,
        cellBgId: `tbl-cell-bg-${pageIndex}-${rowIndex}-${cellIndex}`,
      })];
    }

    const objects = [];
    const innerLeft = left + 4;
    const innerWidth = Math.max(24, width - 8);
    const lineHeight = Math.max(12, Math.round(fontSize * 1.6));
    const maxLines = Math.max(1, Math.floor((height - 4) / lineHeight));
    const totalLines = Math.min(
      maxLines,
      Math.max(1, fragments.reduce((count, fragment) => {
        const fragWidth = estimateTextWidth(fragment.text, fontSize) + 14;
        return count + (fragWidth > innerWidth ? Math.ceil(fragWidth / innerWidth) : 0);
      }, 0))
    );
    const blockHeight = totalLines * lineHeight;
    let cursorX = innerLeft;
    let cursorY = top + Math.max(2, (height - blockHeight) / 2);
    let currentLine = 1;

    fragments.forEach((fragment, fragIndex) => {
      const text = String(fragment.text || '');
      const chipWidth = Math.min(innerWidth, estimateTextWidth(text, fontSize) + 14);
      if (cursorX > innerLeft && cursorX + chipWidth > innerLeft + innerWidth) {
        cursorX = innerLeft;
        if (currentLine < totalLines) {
          cursorY += lineHeight;
          currentLine += 1;
        }
      }
      objects.push(createTextObject(text, cursorX, cursorY + (lineHeight / 2), chipWidth, fontSize, fragment.color || cell.color || '#334155', {
        fontWeight: fragment.strong || cell.strong ? 'bold' : 'normal',
        textAlign: 'left',
        originY: 'center',
        selectable: false,
        evented: false,
        id: `tbl-cell-text-${pageIndex}-${rowIndex}-${cellIndex}-frag-${fragIndex}`,
        cellBgId: `tbl-cell-bg-${pageIndex}-${rowIndex}-${cellIndex}`,
        textBackgroundColor: fragment.bg || '',
      }));
      cursorX += chipWidth + 6;
    });

    return objects;
  }

  function buildTableSlidesFromPayload(payload) {
    const rows = payload.rows || [];
    const chunks = [];
    const rowsPerSlide = Math.max(1, payload.rowsPerSlide || TABLE_ROWS_PER_SLIDE);
    for (let i = 0; i < rows.length; i += rowsPerSlide) {
      chunks.push(rows.slice(i, i + rowsPerSlide));
    }
    if (!chunks.length) chunks.push([]);
    const totalPages = chunks.length;
    const tableWidth = CW - 48;
    const explicitWidthTotal = payload.columns.reduce((sum, column) => sum + (column.widthPx || 0), 0);
    const weightedColumns = payload.columns.filter(column => !column.widthPx);
    const totalWeight = weightedColumns.reduce((sum, column) => sum + (column.weight || 1), 0) || 1;
    const remainingWidth = Math.max(0, tableWidth - explicitWidthTotal);
    const widths = payload.columns.map(column => {
      if (column.widthPx) return column.widthPx;
      return remainingWidth * ((column.weight || 1) / totalWeight);
    });
    return chunks.map((rowsChunk, pageIndex) => {
      const groupTop = 96;
      const groupHeight = payload.groupHeightPx || 30;
      const headerTop = groupTop + groupHeight + 2;
      const headerHeight = payload.headerHeightPx || 34;
      const bodyTop = headerTop + headerHeight;
      const bodyBottom = FTR_Y - 10;
      const availableBodyHeight = Math.max(120, bodyBottom - bodyTop);
      const rowCount = Math.max(1, rowsChunk.length);
      const rowHeight = Math.max(14, Math.floor(availableBodyHeight / rowCount));
      const bodyFontSize = payload.bodyFontSizePx || Math.max(8, Math.min(13, rowHeight - 8));
      const headerFontSize = payload.headerFontSizePx || Math.max(10, Math.min(15, rowHeight - 2));
      const titleFontSize = payload.titleFontSizePx || 34;
      const objects = buildTableSlideBase(payload.title, payload.subtitle, payload.source, pageIndex + 1, totalPages);
      const titleObj = objects.find(obj => obj.id === 'tpl-title');
      if (titleObj) titleObj.fontSize = titleFontSize;
      let groupX = 24;
      let offset = 0;
      payload.headerGroups.forEach(group => {
        const width = widths.slice(offset, offset + group.span).reduce((sum, cur) => sum + cur, 0);
        objects.push({ ...createRectObject(groupX, groupTop, width, groupHeight, group.bg || '#dbe4f0', { selectable: false, evented: false }), id: `tbl-group-bg-${pageIndex}-${offset}` });
        objects.push({ ...createTextObject(group.label || '', groupX + 4, groupTop + groupHeight / 2, width - 8, Math.max(10, headerFontSize - 1), group.color || '#334155', { fontWeight: 'bold', originY: 'center', selectable: false, evented: false }), id: `tbl-group-text-${pageIndex}-${offset}` });
        groupX += width;
        offset += group.span;
      });
      let headerX = 24;
      payload.columns.forEach((column, index) => {
        const width = widths[index];
        objects.push({ ...createRectObject(headerX, headerTop, width, headerHeight, column.bg || '#3f5d8c', { selectable: false, evented: false }), id: `tbl-head-bg-${pageIndex}-${index}` });
        objects.push({ ...createTextObject(column.label || '', headerX + 4, headerTop + headerHeight / 2, width - 8, headerFontSize, column.color || '#ffffff', { fontWeight: 'bold', textAlign: column.align === 'left' ? 'left' : 'center', originY: 'center', selectable: false, evented: false }), id: `tbl-head-text-${pageIndex}-${index}` });
        headerX += width;
      });
      rowsChunk.forEach((row, rowIndex) => {
        let cellX = 24;
        const y = bodyTop + (rowIndex * rowHeight);
        row.cells.forEach((cell, cellIndex) => {
          const width = widths[cellIndex];
          const bgId = `tbl-cell-bg-${pageIndex}-${rowIndex}-${cellIndex}`;
          objects.push({ ...createRectObject(cellX, y, width, rowHeight, cell.bg || '#ffffff', { selectable: false, evented: false }), id: bgId });
          buildInlineCellTextObjects(cell, pageIndex, rowIndex, cellIndex, cellX, y, width, rowHeight, bodyFontSize)
            .forEach(textObj => objects.push(textObj));
          cellX += width;
        });
      });
      return { id: `studio-table-${Date.now()}-${pageIndex}-${Math.random().toString(36).slice(2, 7)}`, json: { version: fabric.version, objects } };
    });
  }

  async function createPreviewFromJson(json, multiplier = 1) {
    const previewCanvas = new fabric.StaticCanvas(null, { width: CW, height: CH, backgroundColor: '#ffffff' });
    try {
      return await new Promise((resolve, reject) => {
        previewCanvas.loadFromJSON(json, () => {
          try {
            previewCanvas.renderAll();
            resolve(previewCanvas.toDataURL({ format: 'png', multiplier }));
          } catch (err) {
            reject(err);
          }
        });
      });
    } catch (err) {
      // Preview generation failed — import continues without a thumbnail
      console.warn('[AvengerStudio] createPreviewFromJson failed:', err);
      return '';
    } finally {
      // Always dispose to free canvas memory
      try { previewCanvas.dispose(); } catch {}
    }
  }

  async function importImageToLibrary(dataURL) {
    slides.push(buildImportedSlide(dataURL));
    await persistSlides();
    renderLibrary();
    if (slides.length) loadFromLibrary(slides.length - 1);
    emitState();
  }

  async function importTableToLibrary(payload, options = {}) {
    if (options.replaceImageSlides) {
      // Only remove slides explicitly marked as 'image' — slides with no kind
      // (manually-created blank pages) must be preserved.
      slides = slides.filter(slide => {
        const kind = String(slide?.kind || '').toLowerCase();
        return kind !== 'image';
      });
    }
    const defs = buildTableSlidesFromPayload(payload);
    for (const slide of defs) {
      slide.kind = 'table';
      slide.preview = '';  // skip heavy preview render — generated lazily after loadFromLibrary
      slides.push(slide);
    }
    await persistSlides();
    renderLibrary();
    if (slides.length) {
      switchTab('studio');
      loadFromLibrary(slides.length - 1);
    }
    emitState();
  }

  function initSlideTemplate(pageNum) {
    canvas.getObjects().filter(obj => obj.id && obj.id.startsWith('tpl-')).forEach(obj => canvas.remove(obj));
    const pNum = pageNum !== undefined ? pageNum : (slides.length + 1);
    const dateVal = $('#studio-footer-date')?.value || '';
    const sourceVal = $('#studio-footer-source')?.value || '';
    canvas.add(new fabric.Line([0, HDR_Y, CW, HDR_Y], { stroke: '#1a3c6e', strokeWidth: 2, selectable: false, evented: false, id: 'tpl-hdrline' }));
    canvas.add(new fabric.IText('ชื่อหัวข้อสไลด์', { left: 20, top: 14, width: 680, fontSize: 44, fontWeight: 'bold', fontFamily: FONT, fill: '#1a3c6e', id: 'tpl-title' }));
    canvas.add(new fabric.Line([0, FTR_Y, CW, FTR_Y], { stroke: '#94a3b8', strokeWidth: 1, selectable: false, evented: false, id: 'tpl-ftrline' }));
    canvas.add(new fabric.Text('www.avenger-planner.com', { left: 20, top: FTR_Y + 12, fontSize: 16, fontFamily: FONT, fill: '#64748b', selectable: false, evented: false, id: 'tpl-ftrleft' }));
    canvas.add(new fabric.IText(`ข้อมูล ณ :  ${dateVal}          ที่มา :  ${sourceVal}`, { left: CW / 2 - 170, top: FTR_Y + 12, fontSize: 16, fontFamily: FONT, fill: '#64748b', id: 'tpl-ftrcenter' }));
    canvas.add(new fabric.Text(String(pNum), { left: CW - 44, top: FTR_Y + 7, fontSize: 22, fontWeight: 'bold', fontFamily: FONT, fill: '#1a3c6e', selectable: false, evented: false, id: 'tpl-ftrright' }));
    const savedLogo = localStorage.getItem(LOGO_KEY);
    if (savedLogo) addLogoToCanvas(savedLogo, false);
    else addLogoPlaceholder();
    canvas.renderAll();
  }

  function addLogoToCanvas(dataURL, render = true) {
    ['tpl-logo', 'tpl-logo-ph', 'tpl-logo-ph-txt'].forEach(id => {
      const obj = canvas.getObjects().find(item => item.id === id);
      if (obj) canvas.remove(obj);
    });
    fabric.Image.fromURL(dataURL, img => {
      const maxW = 148;
      const maxH = 68;
      const scale = Math.min(maxW / img.width, maxH / img.height);
      img.set({ left: CW - 20 - img.width * scale, top: (HDR_Y - img.height * scale) / 2, scaleX: scale, scaleY: scale, id: 'tpl-logo', selectable: true, evented: true });
      canvas.add(img);
      if (render) canvas.renderAll();
    });
  }

  function addLogoPlaceholder() {
    canvas.add(new fabric.Rect({ left: CW - 168, top: 10, width: 148, height: 68, fill: '#f8fafc', stroke: '#bfdbfe', strokeWidth: 1.5, rx: 6, selectable: false, evented: false, id: 'tpl-logo-ph' }));
    canvas.add(new fabric.Text('[ Logo ]', { left: CW - 128, top: 36, fontSize: 18, fontFamily: FONT, fill: '#94a3b8', selectable: false, evented: false, id: 'tpl-logo-ph-txt' }));
  }

  function updateLogoPreview(dataURL) {
    $('#studio-logo-preview').innerHTML = dataURL ? `<img src="${dataURL}" alt="logo">` : '<span style="font-size:12px;color:#94a3b8">ยังไม่มีโลโก้</span>';
  }

  function createNewPage() {
    currentSlideKind = null;
    currentEditIdx = null;
    canvas.clear();
    canvas.setBackgroundColor('#ffffff', () => {
      initSlideTemplate();
      $('#studio-edit-indicator').style.display = 'none';
      updateSlideModeUI();
      updateTableFontEstimate();
      renderLibrary();
      emitState();
    });
  }

  function getTableObjects() {
    return canvas.getObjects().filter(obj => {
      const id = String(obj.id || '');
      if (id.startsWith('tbl-')) return true;
      if (currentSlideKind !== 'table') return false;
      if (id.startsWith('tpl-')) return false;
      if (id === 'tpl-logo' || id === 'tpl-logo-ph' || id === 'tpl-logo-ph-txt') return false;
      return true;
    });
  }

  function updateSlideModeUI() {
    const section = $('#studio-table-tools');
    if (!section) return;
    const isTable = currentSlideKind === 'table';
    section.classList.toggle('active', isTable);
    if (isTable) {
      requestAnimationFrame(() => {
        section.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      });
    }
  }

  function updateTableFontEstimate() {
    const node = $('#studio-table-font-estimate');
    if (!node) return;
    const tableTexts = getTableObjects().filter(obj => obj.type && obj.type.includes('text'));
    if (!tableTexts.length) {
      node.textContent = '-';
      return;
    }
    const avg = tableTexts.reduce((sum, obj) => sum + (obj.fontSize || 0), 0) / tableTexts.length;
    node.textContent = `ประมาณ ${Math.round(avg)} px`;
  }

  function getObjectsBounds(objects) {
    if (!objects.length) return null;
    const xs = [];
    const ys = [];
    objects.forEach(obj => {
      const bounds = obj.getBoundingRect(true, true);
      xs.push(bounds.left, bounds.left + bounds.width);
      ys.push(bounds.top, bounds.top + bounds.height);
    });
    const minX = Math.min(...xs);
    const maxX = Math.max(...xs);
    const minY = Math.min(...ys);
    const maxY = Math.max(...ys);
    return {
      minX,
      maxX,
      minY,
      maxY,
      width: maxX - minX,
      height: maxY - minY,
      centerX: (minX + maxX) / 2,
      centerY: (minY + maxY) / 2,
    };
  }

  function translateTableObjects(objects, dx, dy) {
    objects.forEach(obj => {
      obj.set({
        left: obj.left + dx,
        top: obj.top + dy,
      });
      obj.setCoords();
    });
  }

  function keepTableObjectsInBounds(objects, anchorBounds = null) {
    const current = getObjectsBounds(objects);
    if (!current) return;
    const target = anchorBounds || current;
    const workMinX = 20;
    const workMaxX = CW - 20;
    const workMinY = HDR_Y + 6;
    const workMaxY = FTR_Y - 6;
    let dx = target.centerX - current.centerX;
    let dy = target.centerY - current.centerY;
    const nextMinX = current.minX + dx;
    const nextMaxX = current.maxX + dx;
    const nextMinY = current.minY + dy;
    const nextMaxY = current.maxY + dy;
    if (nextMinX < workMinX) dx += workMinX - nextMinX;
    if (nextMaxX > workMaxX) dx -= nextMaxX - workMaxX;
    if (nextMinY < workMinY) dy += workMinY - nextMinY;
    if (nextMaxY > workMaxY) dy -= nextMaxY - workMaxY;
    translateTableObjects(objects, dx, dy);
  }

  function transformTableObjects(scaleFactor) {
    const objects = getTableObjects();
    if (!objects.length) {
      showToast('ไม่พบวัตถุตารางสำหรับปรับ');
      return;
    }
    const anchorBounds = getObjectsBounds(objects);
    const centerX = anchorBounds.centerX;
    const centerY = anchorBounds.centerY;
    objects.forEach(obj => {
      const nextLeft = centerX + ((obj.left - centerX) * scaleFactor);
      const nextTop = centerY + ((obj.top - centerY) * scaleFactor);
      obj.set({
        left: nextLeft,
        top: nextTop,
      });
      if (obj.type === 'rect') {
        obj.set({
          width: Math.max(8, (obj.width || 0) * scaleFactor),
          height: Math.max(8, (obj.height || 0) * scaleFactor),
        });
      } else if (obj.type && obj.type.includes('text')) {
        obj.set({
          width: Math.max(12, (obj.width || 0) * scaleFactor),
          fontSize: Math.max(7, Math.min(44, Math.round((obj.fontSize || 12) * scaleFactor))),
        });
      }
      obj.setCoords();
    });
    keepTableObjectsInBounds(objects, anchorBounds);
    syncCurrentSlideSnapshot(false);
    canvas.renderAll();
    updateTableFontEstimate();
  }

  function centerTableObjects() {
    const objects = getTableObjects();
    if (!objects.length) return;
    const xs = [];
    const ys = [];
    objects.forEach(obj => {
      const bounds = obj.getBoundingRect(true, true);
      xs.push(bounds.left, bounds.left + bounds.width);
      ys.push(bounds.top, bounds.top + bounds.height);
    });
    const minX = Math.min(...xs);
    const maxX = Math.max(...xs);
    const minY = Math.min(...ys);
    const maxY = Math.max(...ys);
    const centerX = (minX + maxX) / 2;
    const centerY = (minY + maxY) / 2;
    const targetCenterX = CW / 2;
    const targetCenterY = HDR_Y + ((FTR_Y - HDR_Y) / 2);
    const dx = targetCenterX - centerX;
    const dy = targetCenterY - centerY;
    objects.forEach(obj => {
      obj.set({
        left: obj.left + dx,
        top: obj.top + dy,
      });
      obj.setCoords();
    });
    canvas.renderAll();
    updateTableFontEstimate();
  }

  // Move every table object by (dx, dy) pixels and save snapshot.
  // STEP is the nudge amount per button press (px on the 960×540 canvas).
  const TABLE_MOVE_STEP = 10;
  function moveTableObjects(dx, dy) {
    const objects = getTableObjects();
    if (!objects.length) { showToast('ไม่พบวัตถุตารางสำหรับขยับ'); return; }
    objects.forEach(obj => {
      obj.set({ left: obj.left + dx, top: obj.top + dy });
      obj.setCoords();
    });
    canvas.renderAll();
    syncCurrentSlideSnapshot(false);
    updateTableFontEstimate();
  }

  function adjustTableFont(fontFactor) {
    const allTable = getTableObjects();
    const objects = allTable.filter(obj => obj.type && obj.type.includes('text'));
    console.log('[Studio] adjustTableFont: allTable=' + allTable.length + ' text=' + objects.length);
    objects.forEach(obj => console.log('  id=' + obj.id + ' type=' + obj.type + ' fs=' + obj.fontSize));
    if (!objects.length) {
      showToast('ไม่พบข้อความตารางสำหรับปรับ');
      return;
    }
    const anchorBounds = getObjectsBounds(allTable);
    objects.forEach(obj => {
      const current = obj.fontSize || 12;
      let next = Math.round(current * fontFactor);
      // Ensure at least 1px change so small fonts (e.g. 8px × 1.05 = 8.4 → 8) still move
      if (fontFactor > 1 && next <= current) next = current + 1;
      if (fontFactor < 1 && next >= current) next = current - 1;
      next = Math.max(7, Math.min(44, next));
      obj.set('fontSize', next);
      const id = String(obj.id || '');
      let bgId = obj.cellBgId || '';
      if (!bgId && id.startsWith('tbl-cell-text-')) bgId = id.replace(/tbl-cell-text-/, 'tbl-cell-bg-').replace(/-frag-\d+$/, '');
      else if (id.startsWith('tbl-head-text-')) bgId = id.replace('tbl-head-text-', 'tbl-head-bg-');
      else if (id.startsWith('tbl-group-text-')) bgId = id.replace('tbl-group-text-', 'tbl-group-bg-');
      const bg = bgId ? canvas.getObjects().find(item => item.id === bgId) : null;
      next = Math.min(next, getMaxFontSizeForBounds(obj.text, bg, id.includes('-frag-') ? 14 : 8, 6));
      if (bg) {
        const isCenteredX = obj.originX === 'center';
        const isCenteredY = obj.originY === 'center';
        // Promote to originY:'center' if not already, so future adjustments stay centered
        if (!isCenteredY) {
          obj.set('originY', 'center');
        }
        obj.set({
          left: isCenteredX ? bg.left + bg.width / 2 : bg.left + 4,
          top: bg.top + bg.height / 2,
          width: Math.max(12, bg.width - 8),
        });
      }
      obj.setCoords();
    });
    keepTableObjectsInBounds(allTable, anchorBounds);
    syncCurrentSlideSnapshot(false);
    canvas.renderAll();
    updateTableFontEstimate();
    showToast('ปรับขนาดฟอนต์ตารางแล้ว');
  }

  // For TABLE slides: reconstruct JSON by patching the stored JSON with current canvas
  // positions/sizes rather than calling canvas.toJSON(). This avoids the Fabric.js 5.3.1
  // bug where IText.toObject() → stylesToArray() crashes with TypeError even after clearing
  // obj.styles — because stylesToArray is a non-patchable local inside the minified bundle.
  function snapshotTableSlide(index) {
    const stored = slides[index];
    if (!stored || !stored.json || !Array.isArray(stored.json.objects)) return null;
    const canvasObjects = canvas.getObjects();
    const updatedObjects = stored.json.objects.map(jsonObj => {
      const co = canvasObjects.find(o => o.id === jsonObj.id);
      if (!co) return jsonObj;
      const patch = { left: co.left, top: co.top };
      if (co.width  !== undefined) patch.width  = co.width;
      if (co.height !== undefined) patch.height = co.height;
      if (co.scaleX !== undefined && co.scaleX !== 1) patch.scaleX = co.scaleX;
      if (co.scaleY !== undefined && co.scaleY !== 1) patch.scaleY = co.scaleY;
      if (co.fontSize !== undefined) patch.fontSize = co.fontSize;
      if (co.originX !== undefined) patch.originX = co.originX;
      if (co.originY !== undefined) patch.originY = co.originY;
      return { ...jsonObj, ...patch };
    });
    return { ...stored.json, objects: updatedObjects };
  }

  function syncCurrentSlideSnapshot(showMessage = false) {
    if (currentEditIdx === null || !slides[currentEditIdx]) return;
    if (!canvas) return;

    const slide = slides[currentEditIdx];

    // TABLE slides: use custom JSON reconstruction to avoid canvas.toJSON() which
    // triggers Fabric.js 5.3.1's unfixable stylesToArray TypeError.
    if (slide.kind === 'table') {
      const json = snapshotTableSlide(currentEditIdx);
      if (!json) return;
      let preview = slide.preview || '';
      try { preview = canvas.toDataURL({ format: 'png', multiplier: 1 }); } catch {}
      slides[currentEditIdx] = { ...slide, json, preview };
      persistSlides()
        .then(() => { renderLibrary(); if (showMessage) showToast('บันทึกการปรับตารางแล้ว'); })
        .catch(err => showToast(`บันทึกไม่สำเร็จ: ${err.message || err}`));
      return;
    }

    // Non-table slides: use normal canvas serialisation
    let dataURL, json;
    try {
      dataURL = canvas.toDataURL({ format: 'png', multiplier: 1 });
      json = canvas.toJSON(['id', 'selectable', 'evented']);
    } catch (err) {
      console.warn('[Studio] syncCurrentSlideSnapshot: canvas serialisation failed, skipping:', err);
      return;
    }
    slides[currentEditIdx] = { ...(slide || {}), json, preview: dataURL };
    persistSlides()
      .then(() => {
        renderLibrary();
        if (showMessage) showToast('บันทึกการปรับตารางแล้ว');
      })
      .catch(err => showToast(`บันทึกไม่สำเร็จ: ${err.message || err}`));
  }

  function reorderSlides(fromIndex, toIndex) {
    if (fromIndex === toIndex || fromIndex === null || toIndex === null || fromIndex < 0 || toIndex < 0) return;
    const [moved] = slides.splice(fromIndex, 1);
    slides.splice(toIndex, 0, moved);
    if (currentEditIdx === fromIndex) currentEditIdx = toIndex;
    else if (fromIndex < currentEditIdx && toIndex >= currentEditIdx) currentEditIdx -= 1;
    else if (fromIndex > currentEditIdx && toIndex <= currentEditIdx) currentEditIdx += 1;
    persistSlides()
      .then(() => {
        renderLibrary();
        if (currentEditIdx !== null && slides[currentEditIdx]) {
          currentSlideKind = slides[currentEditIdx]?.kind || null;
          const pageNo = canvas.getObjects().find(obj => obj.id === 'tpl-ftrright');
          if (pageNo) pageNo.set('text', String(currentEditIdx + 1));
          $('#studio-page-display').textContent = String(currentEditIdx + 1);
          updateSlideModeUI();
          canvas.renderAll();
        }
        showToast('ย้ายลำดับสไลด์แล้ว');
      })
      .catch(err => showToast(`ย้ายลำดับไม่สำเร็จ: ${err.message || err}`));
  }

  function renderLibrary() {
    const grid = $('#studio-slides-grid');
    grid.innerHTML = '';
    if (!slides.length) {
      grid.innerHTML = '<p style="font-size:13px;color:#cbd5e1;padding:20px">ยังไม่มีหน้าในคลัง...</p>';
      return;
    }
    slides.forEach((slide, index) => {
      const wrap = document.createElement('div');
      wrap.className = 'studio-slide-wrap';
      wrap.draggable = true;
      wrap.addEventListener('dragstart', () => {
        dragIndex = index;
        wrap.classList.add('dragging');
      });
      wrap.addEventListener('dragend', () => {
        dragIndex = null;
        wrap.classList.remove('dragging');
      });
      wrap.addEventListener('dragover', e => {
        e.preventDefault();
      });
      wrap.addEventListener('drop', e => {
        e.preventDefault();
        reorderSlides(dragIndex, index);
      });
      const item = document.createElement('div');
      item.className = `studio-slide-item${currentEditIdx === index ? ' active' : ''}`;
      item.innerHTML = `<img src="${slide.preview}">`;
      item.addEventListener('click', () => loadFromLibrary(index));
      const kind = document.createElement('div');
      const kindLabel = slide.kind === 'table' ? 'TABLE' : 'IMAGE';
      kind.className = `studio-slide-kind ${slide.kind === 'table' ? 'kind-table' : 'kind-image'}`;
      kind.textContent = kindLabel;
      const del = document.createElement('button');
      del.className = 'studio-slide-del';
      del.type = 'button';
      del.innerHTML = '&times;';
      del.addEventListener('click', e => {
        e.stopPropagation();
        slides.splice(index, 1);
        persistSlides().then(() => {
          if (currentEditIdx === index) createNewPage();
          else renderLibrary();
          emitState();
        }).catch(err => showToast(`ลบสไลด์ไม่สำเร็จ: ${err.message || err}`));
      });
      const badge = document.createElement('div');
      badge.className = 'studio-slide-badge';
      badge.textContent = index + 1;
      wrap.append(item, kind, del, badge);
      grid.appendChild(wrap);
    });
  }

  function loadFromLibrary(index) {
    if (currentEditIdx !== null && currentEditIdx !== index) {
      try { syncCurrentSlideSnapshot(false); } catch (err) {
        console.warn('[Studio] snapshot skipped during load:', err);
      }
    }
    currentEditIdx = index;
    currentSlideKind = slides[index]?.kind || null;
    // Migrate old textbox/i-text objects to text so canvas.toJSON() won't stack-overflow
    const slideJson = migrateSlideJson(slides[index].json);
    if (slideJson !== slides[index].json) slides[index] = { ...slides[index], json: slideJson };
    loadFromJSONChunked(canvas, slideJson, () => {
      // Strip per-character styles immediately after load so that canvas.toJSON()
      // never triggers the Fabric.js 5.3.1 stylesToArray TypeError.
      canvas.getObjects().forEach(obj => {
        if (obj && obj.styles !== undefined) obj.styles = {};
      });
      // Upgrade old text objects to use originY:'center' for proper vertical centering
      if (slides[index]?.kind === 'table') {
        canvas.getObjects().forEach(obj => {
          if (!obj.type || !obj.type.includes('text')) return;
          const id = String(obj.id || '');
          let bgId = '';
          if (id.startsWith('tbl-cell-text-')) bgId = id.replace('tbl-cell-text-', 'tbl-cell-bg-');
          else if (id.startsWith('tbl-head-text-')) bgId = id.replace('tbl-head-text-', 'tbl-head-bg-');
          else if (id.startsWith('tbl-group-text-')) bgId = id.replace('tbl-group-text-', 'tbl-group-bg-');
          if (!bgId) return;
          const bg = canvas.getObjects().find(o => o.id === bgId);
          if (!bg) return;
          if (obj.originY !== 'center') {
            const isCenteredX = obj.originX === 'center';
            obj.set({
              originY: 'center',
              top: bg.top + bg.height / 2,
              left: isCenteredX ? bg.left + bg.width / 2 : bg.left + 4,
            });
            obj.setCoords();
          }
        });
      }
      const pageNo = canvas.getObjects().find(obj => obj.id === 'tpl-ftrright');
      if (pageNo) pageNo.set('text', String(index + 1));
      canvas.renderAll();
      $('#studio-edit-indicator').style.display = 'block';
      $('#studio-page-display').textContent = String(index + 1);
      updateSlideModeUI();
      updateTableFontEstimate();
      renderLibrary();
      if (slides[index]?.kind === 'table') {
        canvas.discardActiveObject();
        showToast('สไลด์ตารางถูกล็อกเป็นก้อนเดียว แก้ไขผ่านต้นทางแล้วส่งใหม่จะเสถียรกว่า');
      }
      // Lazily generate preview thumbnail if missing (e.g. just imported without preview)
      if (!slides[index]?.preview) {
        requestAnimationFrame(() => {
          try {
            const preview = canvas.toDataURL({ format: 'png', multiplier: 1 });
            if (slides[index]) {
              slides[index] = { ...slides[index], preview };
              persistSlides().then(() => renderLibrary()).catch(() => {});
            }
          } catch {}
        });
      }
      emitState();
    });
  }

  function saveToLibrary() {
    const dataURL = canvas.toDataURL({ format: 'png', multiplier: 1 });
    const json = canvas.toJSON(['id', 'selectable', 'evented']);
    if (currentEditIdx !== null) slides[currentEditIdx] = { ...(slides[currentEditIdx] || {}), json, preview: dataURL };
    else slides.push({ id: `studio-slide-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`, json, preview: dataURL });
    persistSlides().then(() => {
      createNewPage();
      renderLibrary();
      showToast('บันทึกแล้ว');
    }).catch(err => showToast(`บันทึกไม่สำเร็จ: ${err.message || err}`));
  }

  function switchTab(tabName) {
    if (tabName !== 'studio' && currentEditIdx !== null) syncCurrentSlideSnapshot(false);
    root.querySelectorAll('.studio-tab-content').forEach(node => node.classList.remove('active'));
    root.querySelectorAll('.studio-tab-btn').forEach(node => node.classList.remove('active'));
    $(`#studio-tab-${tabName}`).classList.add('active');
    root.querySelector(`[data-studio-tab="${tabName}"]`).classList.add('active');
    if (tabName === 'canvas') setTimeout(initCanvasTab, 50);
    emitState();
  }

  function openPreview() {
    if (!slides.length) {
      showToast('คลังว่างเปล่า');
      return;
    }
    previewIdx = 0;
    $('#studio-preview-overlay').classList.add('open');
    updatePreview();
  }

  function updatePreview() {
    $('#studio-preview-img').src = slides[previewIdx].preview;
    $('#studio-preview-info').textContent = `หน้า ${previewIdx + 1} / ${slides.length}`;
  }

  async function buildPdfSlideImage(slide) {
    if (!slide) return '';
    if (slide.json) {
      try {
        return await createPreviewFromJson(slide.json, 3);
      } catch {}
    }
    return slide.preview || '';
  }

  async function exportToPDF() {
    if (!slides.length) {
      showToast('กรุณาบันทึกงานลงคลังอย่างน้อย 1 หน้า');
      return;
    }
    if (currentEditIdx !== null && slides[currentEditIdx]) {
      try { syncCurrentSlideSnapshot(false); } catch {}
    }
    const { jsPDF } = global.jspdf;
    // PDF page size matches the rendered image (multiplier × canvas) so the PNG is embedded
    // at 1:1 — no downscaling, no quality loss.  multiplier=3 gives 2880×1620 px per page.
    const PDF_MULT = 3;
    const pdfW = CW * PDF_MULT;
    const pdfH = CH * PDF_MULT;
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'px', format: [pdfW, pdfH] });
    let addedPages = 0;
    for (let index = 0; index < slides.length; index += 1) {
      const slide = slides[index];
      const image = await buildPdfSlideImage(slide);
      if (!image) continue;
      if (addedPages > 0) pdf.addPage([pdfW, pdfH], 'landscape');
      pdf.addImage(image, 'PNG', 0, 0, pdfW, pdfH, undefined, 'NONE');
      addedPages += 1;
    }
    if (!addedPages) {
      showToast('Export PDF ไม่สำเร็จ: ไม่พบภาพสไลด์ที่พร้อมใช้งาน');
      return;
    }
    pdf.save('avenger-presentation.pdf');
  }

  function initCanvasTab() {
    if (canvasCReady) return;
    canvasC = new fabric.Canvas('studio-canvas-c', { width: CW, height: CH, backgroundColor: '#ffffff', isDrawingMode: true, preserveObjectStacking: true });
    canvasC.freeDrawingBrush.color = currentDrawColor;
    canvasC.freeDrawingBrush.width = 3;
    canvasCReady = true;
    canvasC.on('selection:created', updateCanvasToolbar);
    canvasC.on('selection:updated', updateCanvasToolbar);
    canvasC.on('selection:cleared', () => {
      $('#studio-c-text-tools').classList.add('studio-control-disabled');
      $('#studio-c-obj-tools').classList.add('studio-control-disabled');
    });
  }

  function setDrawMode(drawMode) {
    if (!canvasCReady) return;
    canvasC.isDrawingMode = drawMode;
    $('#studio-draw-mode').classList.toggle('active-mode', drawMode);
    $('#studio-select-mode').classList.toggle('active-mode', !drawMode);
    const brushTools = $('#studio-c-brush-tools');
    brushTools.style.opacity = drawMode ? '1' : '.45';
    brushTools.style.pointerEvents = drawMode ? 'auto' : 'none';
  }

  function updateToolbar() {
    const obj = canvas.getActiveObject();
    if (!obj) return;
    $('#studio-general-tools').classList.remove('studio-control-disabled');
    $('#studio-alpha-slider').value = (obj.opacity || 1) * 100;
    $('#studio-alpha-val').textContent = `${Math.round((obj.opacity || 1) * 100)}%`;
    $('#studio-stroke-width-val').textContent = `${obj.strokeWidth ?? 1} px`;
    $('#studio-stroke-toggle').classList.toggle('active-mode', (obj.strokeWidth ?? 0) > 0);
    if (obj.type && obj.type.includes('text')) {
      $('#studio-text-tools').classList.remove('studio-control-disabled');
      $('#studio-font-slider').value = obj.fontSize;
      $('#studio-font-val').textContent = String(obj.fontSize);
    } else {
      $('#studio-text-tools').classList.add('studio-control-disabled');
    }
  }

  function updateCanvasToolbar() {
    const obj = canvasC?.getActiveObject();
    if (!obj) return;
    $('#studio-c-obj-tools').classList.remove('studio-control-disabled');
    $('#studio-c-alpha-slider').value = (obj.opacity || 1) * 100;
    $('#studio-c-alpha-val').textContent = `${Math.round((obj.opacity || 1) * 100)}%`;
    if (obj.type && obj.type.includes('text')) {
      $('#studio-c-text-tools').classList.remove('studio-control-disabled');
      $('#studio-c-font-slider').value = obj.fontSize;
      $('#studio-c-font-val').textContent = String(obj.fontSize);
    } else {
      $('#studio-c-text-tools').classList.add('studio-control-disabled');
    }
  }

  function changeColor(hex) {
    const obj = canvas.getActiveObject();
    if (!obj) return;
    obj.set('fill', hex);
    canvas.renderAll();
  }

  function toggleStroke() {
    const obj = canvas.getActiveObject();
    if (!obj) return;
    const hasStroke = (obj.strokeWidth ?? 0) > 0;
    if (hasStroke) {
      obj.set({ stroke: 'rgba(0,0,0,0)', strokeWidth: 0 });
    } else {
      const picked = $('#studio-stroke-color')?.value || '#94a3b8';
      obj.set({
        stroke: picked,
        strokeWidth: Math.max(1, obj.strokeWidth ?? 1),
      });
    }
    $('#studio-stroke-width-val').textContent = `${obj.strokeWidth ?? 0} px`;
    $('#studio-stroke-toggle').classList.toggle('active-mode', (obj.strokeWidth ?? 0) > 0);
    canvas.renderAll();
    syncCurrentSlideSnapshot(false);
  }

  function nudgeStrokeWidth(delta) {
    const obj = canvas.getActiveObject();
    if (!obj) return;
    const next = Math.max(0, Math.min(20, (obj.strokeWidth ?? 1) + delta));
    obj.set('strokeWidth', next);
    if (next > 0 && (!obj.stroke || obj.stroke === 'rgba(0,0,0,0)')) {
      obj.set('stroke', $('#studio-stroke-color')?.value || '#94a3b8');
    }
    $('#studio-stroke-width-val').textContent = `${next} px`;
    $('#studio-stroke-toggle').classList.toggle('active-mode', next > 0);
    canvas.renderAll();
    syncCurrentSlideSnapshot(false);
  }

  function normalizeHexColor(value) {
    const raw = String(value || '').trim();
    if (/^#[0-9a-f]{6}$/i.test(raw)) return raw;
    return '#d7dee9';
  }

  function setCanvasColor(hex, element) {
    currentDrawColor = hex;
    if (canvasC) {
      canvasC.freeDrawingBrush.color = hex;
      const obj = canvasC.getActiveObject();
      if (obj) {
        obj.set('fill', hex);
        canvasC.renderAll();
      }
    }
    root.querySelectorAll('#studio-c-palette .studio-color-box').forEach(node => node.classList.remove('selected'));
    element.classList.add('selected');
  }

  function bindPalette(containerId, onPick) {
    const colors = ['transparent', '#1a3c6e', '#ef4444', '#10b981', '#e8a317', '#0f172a', '#4a90d9', '#ec4899', '#ffffff'];
    const container = $(containerId);
    colors.forEach((hex, index) => {
      const node = document.createElement('button');
      node.type = 'button';
      node.className = `studio-color-box${index === 0 && containerId === '#studio-c-palette' ? ' selected' : ''}`;
      node.style.background = hex === 'transparent'
        ? 'linear-gradient(135deg, #ffffff 0 45%, #fecaca 45% 55%, #ffffff 55% 100%)'
        : hex;
      if (hex === '#ffffff' || hex === 'transparent') node.style.border = '1px solid #ddd';
      node.addEventListener('click', () => onPick(hex, node));
      container.appendChild(node);
    });
  }

  function bindEvents() {
    root.querySelectorAll('[data-studio-tab]').forEach(node => node.addEventListener('click', () => switchTab(node.dataset.studioTab)));
    $('#studio-preview-btn').addEventListener('click', openPreview);
    $('#studio-export-btn').addEventListener('click', exportToPDF);
    $('#studio-new-page').addEventListener('click', createNewPage);
    $('#studio-add-rect').addEventListener('click', () => { const obj = new fabric.Rect({ left: 220, top: 200, width: 130, height: 100, fill: '#1a3c6e' }); canvas.add(obj).setActiveObject(obj); });
    $('#studio-add-circle').addEventListener('click', () => { const obj = new fabric.Circle({ left: 260, top: 220, radius: 55, fill: '#e8a317' }); canvas.add(obj).setActiveObject(obj); });
    $('#studio-add-line').addEventListener('click', () => {
      const obj = new fabric.Line([220, 240, 380, 240], { stroke: '#1a3c6e', strokeWidth: 3 });
      canvas.add(obj).setActiveObject(obj);
      canvas.renderAll();
    });
    $('#studio-add-text').addEventListener('click', () => { const obj = new fabric.IText('พิมพ์ข้อความที่นี่...', { left: 220, top: 250, fontSize: 28, fontFamily: FONT, fill: '#1e293b' }); canvas.add(obj).setActiveObject(obj); });
    $('#studio-apply-footer').addEventListener('click', () => {
      const obj = canvas.getObjects().find(item => item.id === 'tpl-ftrcenter');
      if (!obj) return;
      obj.set('text', `ข้อมูล ณ :  ${$('#studio-footer-date').value}          ที่มา :  ${$('#studio-footer-source').value}`);
      canvas.renderAll();
      showToast('อัพเดต Footer แล้ว');
    });
    $('#studio-save-library').addEventListener('click', saveToLibrary);
    $('#studio-table-scale-down').addEventListener('click', () => { transformTableObjects(0.95); showToast('ย่อตารางแล้ว'); });
    $('#studio-table-scale-up').addEventListener('click', () => { transformTableObjects(1.05); showToast('ขยายตารางแล้ว'); });
    $('#studio-table-font-down').addEventListener('click', () => adjustTableFont(0.95));
    $('#studio-table-font-up').addEventListener('click', () => adjustTableFont(1.05));
    $('#studio-table-move-up').addEventListener('click',     () => { moveTableObjects(0, -TABLE_MOVE_STEP); showToast('ขยับขึ้นแล้ว'); });
    $('#studio-table-move-down').addEventListener('click',   () => { moveTableObjects(0,  TABLE_MOVE_STEP); showToast('ขยับลงแล้ว'); });
    $('#studio-table-move-left').addEventListener('click',   () => { moveTableObjects(-TABLE_MOVE_STEP, 0); showToast('ขยับซ้ายแล้ว'); });
    $('#studio-table-move-right').addEventListener('click',  () => { moveTableObjects( TABLE_MOVE_STEP, 0); showToast('ขยับขวาแล้ว'); });
    $('#studio-table-move-center').addEventListener('click', () => { centerTableObjects(); syncCurrentSlideSnapshot(false); showToast('จัดกึ่งกลางแล้ว'); });
    $('#studio-font-slider').addEventListener('input', e => {
      const obj = canvas.getActiveObject();
      if (!obj || !obj.type?.includes('text')) return;
      obj.set('fontSize', parseInt(e.target.value, 10));
      $('#studio-font-val').textContent = e.target.value;
      canvas.renderAll();
    });
    $('#studio-alpha-slider').addEventListener('input', e => {
      const obj = canvas.getActiveObject();
      if (!obj) return;
      obj.set('opacity', e.target.value / 100);
      $('#studio-alpha-val').textContent = `${e.target.value}%`;
      canvas.renderAll();
    });
    $('#studio-stroke-color').addEventListener('input', e => {
      const obj = canvas.getActiveObject();
      if (!obj || (obj.strokeWidth ?? 0) <= 0) return;
      obj.set('stroke', e.target.value);
      canvas.renderAll();
      syncCurrentSlideSnapshot(false);
    });
    $('#studio-stroke-toggle').addEventListener('click', toggleStroke);
    $('#studio-stroke-width-down').addEventListener('click', () => nudgeStrokeWidth(-1));
    $('#studio-stroke-width-up').addEventListener('click', () => nudgeStrokeWidth(1));
    $('#studio-close-preview').addEventListener('click', () => $('#studio-preview-overlay').classList.remove('open'));
    $('#studio-preview-overlay').addEventListener('click', e => { if (e.target.id === 'studio-preview-overlay') $('#studio-preview-overlay').classList.remove('open'); });
    $('#studio-prev-preview').addEventListener('click', () => { previewIdx = (previewIdx - 1 + slides.length) % slides.length; updatePreview(); });
    $('#studio-next-preview').addEventListener('click', () => { previewIdx = (previewIdx + 1) % slides.length; updatePreview(); });
    $('#studio-draw-mode').addEventListener('click', () => setDrawMode(true));
    $('#studio-select-mode').addEventListener('click', () => setDrawMode(false));
    $('#studio-add-rect-c').addEventListener('click', () => { if (!canvasCReady) initCanvasTab(); setDrawMode(false); canvasC.add(new fabric.Rect({ left: 130, top: 160, width: 160, height: 110, fill: currentDrawColor })); });
    $('#studio-add-circle-c').addEventListener('click', () => { if (!canvasCReady) initCanvasTab(); setDrawMode(false); canvasC.add(new fabric.Circle({ left: 170, top: 170, radius: 65, fill: currentDrawColor })); });
    $('#studio-add-text-c').addEventListener('click', () => { if (!canvasCReady) initCanvasTab(); setDrawMode(false); const obj = new fabric.IText('พิมพ์ข้อความ...', { left: 220, top: 230, fontSize: 28, fontFamily: FONT, fill: currentDrawColor }); canvasC.add(obj).setActiveObject(obj); });
    $('#studio-brush-slider').addEventListener('input', e => { if (!canvasCReady) initCanvasTab(); canvasC.freeDrawingBrush.width = parseInt(e.target.value, 10); $('#studio-brush-val').textContent = e.target.value; });
    $('#studio-c-font-slider').addEventListener('input', e => { const obj = canvasC?.getActiveObject(); if (!obj?.type?.includes('text')) return; obj.set('fontSize', parseInt(e.target.value, 10)); $('#studio-c-font-val').textContent = e.target.value; canvasC.renderAll(); });
    $('#studio-c-alpha-slider').addEventListener('input', e => { const obj = canvasC?.getActiveObject(); if (!obj) return; obj.set('opacity', e.target.value / 100); $('#studio-c-alpha-val').textContent = `${e.target.value}%`; canvasC.renderAll(); });
    $('#studio-clear-canvas-c').addEventListener('click', () => { if (!canvasCReady) return; if (!confirm('ล้าง Canvas ทั้งหมด?')) return; canvasC.clear(); canvasC.setBackgroundColor('#ffffff', () => canvasC.renderAll()); });
    $('#studio-copy-canvas-c').addEventListener('click', async () => {
      if (!canvasCReady) return;
      try {
        const blob = await fetch(canvasC.toDataURL({ format: 'png' })).then(r => r.blob());
        await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
        showToast('คัดลอกรูปแล้ว');
      } catch {
        showToast('คัดลอกไม่สำเร็จ');
      }
    });
    $('#studio-send-to-slide').addEventListener('click', () => {
      if (!canvasCReady) return;
      const dataURL = canvasC.toDataURL({ format: 'png' });
      fabric.Image.fromURL(dataURL, img => {
        const scale = Math.min(CW / CW, CH / CH) * 0.72;
        img.set({ left: (CW - CW * scale) / 2, top: HDR_Y + 10 + (FTR_Y - HDR_Y - CH * scale - 20) / 2, scaleX: scale, scaleY: scale, selectable: true, evented: true });
        canvas.add(img).setActiveObject(img);
        canvas.renderAll();
      });
      switchTab('studio');
      showToast('ส่งรูปไปยัง Studio แล้ว');
    });
    bindPalette('#studio-palette', changeColor);
    bindPalette('#studio-c-palette', setCanvasColor);
    canvas.on('selection:created', updateToolbar);
    canvas.on('selection:updated', updateToolbar);
    canvas.on('selection:cleared', () => {
      $('#studio-text-tools').classList.add('studio-control-disabled');
      $('#studio-general-tools').classList.add('studio-control-disabled');
    });
    global.addEventListener('keydown', e => {
      if (!mounted) return;
      const active = canvas.getActiveObject();
      if (active && !active.isEditing && (e.key === 'Delete' || e.key === 'Backspace')) {
        if (!(active.id && active.id.startsWith('tpl-'))) {
          if (active.type === 'activeSelection') {
            active.forEachObject(obj => canvas.remove(obj));
            canvas.discardActiveObject();
          } else canvas.remove(active);
          canvas.renderAll();
        }
      }
      if (active && !active.isEditing && e.ctrlKey && e.key === 'c') active.clone(clone => { clipboard = clone; });
      if (e.ctrlKey && e.key === 'v' && clipboard) {
        clipboard.clone(clone => {
          canvas.discardActiveObject();
          clone.set({ left: clone.left + 20, top: clone.top + 20, evented: true });
          if (clone.type === 'activeSelection') {
            clone.canvas = canvas;
            clone.forEachObject(obj => canvas.add(obj));
            clone.setCoords();
          } else canvas.add(clone);
          clipboard.top += 20;
          clipboard.left += 20;
          canvas.setActiveObject(clone);
          canvas.requestRenderAll();
        });
      }
    });
  }

  async function mount(area, options = {}) {
    injectStyles();
    // Dispose old Fabric instances BEFORE destroying DOM nodes — Fabric attaches
    // document-level mousemove/mouseup listeners that must be removed or they will
    // intercept all subsequent clicks (including nav items) on every re-mount.
    if (canvas) { try { canvas.dispose(); } catch {} canvas = null; }
    if (canvasC) { try { canvasC.dispose(); } catch {} canvasC = null; }
    canvasCReady = false;
    root = area;
    stateListener = typeof options.onStateChange === 'function' ? options.onStateChange : null;
    root.innerHTML = renderLayout();
    // Only load from IndexedDB on first mount — subsequent mounts reuse in-memory slides
    // to avoid race condition with in-flight persistSlides() and to restore edits correctly
    if (!slides.length) {
      await loadSlides();
    }
    await loadFonts();
    canvas = new fabric.Canvas('studio-canvas', { width: CW, height: CH, backgroundColor: '#ffffff', preserveObjectStacking: true });
    bindEvents();
    const savedLogo = localStorage.getItem(LOGO_KEY);
    if (savedLogo) updateLogoPreview(savedLogo);
    renderLibrary();
    // Restore the previously active slide, or create a blank page if none exist
    if (slides.length > 0) {
      const restoreIdx = (currentEditIdx !== null && slides[currentEditIdx]) ? currentEditIdx : slides.length - 1;
      loadFromLibrary(restoreIdx);
    } else {
      currentEditIdx = null;
      previewIdx = 0;
      createNewPage();
    }
    mounted = true;
    emitState();
  }

  function isMounted() {
    return mounted;
  }

  async function importQueueItem(item) {
    if (!mounted) return false;
    if (item.kind === 'table' && item.payload) await importTableToLibrary(item.payload, { replaceImageSlides: item.replaceImageSlides });
    else if (item.image) await importImageToLibrary(item.image);
    return true;
  }

  global.AvengerStudio = {
    mount,
    isMounted,
    importQueueItem,
    persistCurrent: () => syncCurrentSlideSnapshot(false),
    getState: () => ({ slideCount: slides.length, currentEditIdx, previewIdx }),
  };
})(window);
