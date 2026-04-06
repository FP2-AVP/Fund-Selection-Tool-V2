/* ============================================================
   Fund Selection Tool – FP2
   Google Sheets API Handler
   ============================================================ */
'use strict';

const SheetsAPI = {
  accessToken:  null,
  tokenClient:  null,
  _userInfo:    null,

  _notSignedInMessage() {
    if (CONFIG.BYPASS_LOGIN) {
      return 'ยังไม่ได้เข้าสู่ระบบ: ตอนนี้เปิด BYPASS_LOGIN=true อยู่ ซึ่งข้ามเฉพาะหน้า login แต่ไม่ได้สร้าง Google access token จริง ให้ตั้ง BYPASS_LOGIN=false แล้วกด Sign in ผ่านปุ่มของแอป';
    }
    return 'ยังไม่ได้เข้าสู่ระบบ: โปรเจ็กต์นี้ใช้ Google Identity Services แบบ popup token flow ผ่านปุ่ม Sign in ของแอป ไม่ได้อ่านสถานะจาก redirect URL ที่ส่งเอง';
  },

  /* ── Token Client (lazy init) ── */
  _ensureClient() {
    if (this.tokenClient) return this.tokenClient;
    if (typeof google === 'undefined') {
      throw new Error('Google Identity Services ยังไม่โหลด กรุณารีเฟรชหน้าเว็บ');
    }
    this.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.CLIENT_ID,
      scope:     CONFIG.SCOPES,
      callback:  () => {},   // overridden per call
    });
    return this.tokenClient;
  },

  /* ── Request access token (shows Google popup) ── */
  requestToken(silent = false) {
    return new Promise((resolve, reject) => {
      const client = this._ensureClient();
      client.callback = (resp) => {
        if (resp.error) {
          reject(new Error(resp.error_description || resp.error));
          return;
        }
        this.accessToken = resp.access_token;
        this._userInfo = null; // reset cached user info
        resolve(resp);
      };
      client.requestAccessToken({ prompt: silent ? '' : 'select_account' });
    });
  },

  /* ── Get user profile from Google ── */
  async getUserInfo() {
    if (this._userInfo) return this._userInfo;
    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');

    const resp = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { 'Authorization': `Bearer ${this.accessToken}` },
    });
    if (!resp.ok) throw new Error('ไม่สามารถดึงข้อมูลผู้ใช้ได้');
    this._userInfo = await resp.json();
    return this._userInfo;
  },

  /* ── Fetch sheet values ── */
  async fetchSheetData(sheetId, tabName = 'Sheet1') {
    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');

    const range = encodeURIComponent(tabName);
    const url   = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}`;

    let resp = await fetch(url, {
      headers: { 'Authorization': `Bearer ${this.accessToken}` },
    });

    /* ── Token expired → silent refresh ── */
    if (resp.status === 401) {
      try {
        await this.requestToken(true);
      } catch {
        throw new Error('Session หมดอายุ กรุณาเข้าสู่ระบบอีกครั้ง');
      }
      resp = await fetch(url, {
        headers: { 'Authorization': `Bearer ${this.accessToken}` },
      });
    }

    if (!resp.ok) {
      let errMsg = `HTTP ${resp.status}`;
      try {
        const body = await resp.json();
        errMsg = body?.error?.message || errMsg;
      } catch { /* ignore */ }

      if (resp.status === 403) {
        throw new Error(`ไม่มีสิทธิ์เข้าถึง Sheet นี้ (${errMsg})\nตรวจสอบว่าบัญชีที่ใช้มีสิทธิ์อ่าน Google Sheet`);
      }
      if (resp.status === 404) {
        throw new Error(`ไม่พบ Sheet หรือ Tab ที่ระบุ (${errMsg})\nตรวจสอบ Sheet ID และชื่อ Tab ใน config.js`);
      }
      throw new Error(errMsg);
    }

    const data = await resp.json();
    return data.values || [];
  },

  /* ── Get spreadsheet metadata (list of tab names) ── */
  async getSheetTabs(sheetId) {
    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');

    const url  = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties.title,properties.title`;
    const resp = await fetch(url, {
      headers: { 'Authorization': `Bearer ${this.accessToken}` },
    });
    if (!resp.ok) throw new Error(`ดึง metadata ไม่สำเร็จ: ${resp.status}`);

    const data = await resp.json();
    return {
      title: data.properties?.title || '',
      tabs:  (data.sheets || []).map(s => s.properties.title),
    };
  },

  /* ── Sign out ── */
  signOut() {
    if (this.accessToken) {
      google.accounts.oauth2.revoke(this.accessToken, () => {});
    }
    this.accessToken = null;
    this._userInfo   = null;
    this.tokenClient = null;
  },
};
