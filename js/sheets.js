/* ============================================================
   Fund Selection Tool – FP2
   Google Sheets API Handler
   ============================================================ */
'use strict';

const SheetsAPI = {
  accessToken:  null,
  tokenClient:  null,
  tokenProvider: null,
  _userInfo:    null,

  _notSignedInMessage() {
    if (CONFIG.BYPASS_LOGIN) {
      return 'ยังไม่ได้เข้าสู่ระบบ: ตอนนี้เปิด BYPASS_LOGIN=true อยู่ ซึ่งข้ามเฉพาะหน้า login แต่ไม่ได้สร้าง Google access token จริง ให้ตั้ง BYPASS_LOGIN=false แล้วกด Sign in ผ่านปุ่มของแอป';
    }
    return 'ยังไม่ได้เข้าสู่ระบบ: โปรเจ็กต์นี้ใช้ Google Identity Services แบบ popup token flow ผ่านปุ่ม Sign in ของแอป ไม่ได้อ่านสถานะจาก redirect URL ที่ส่งเอง';
  },

  _quoteSheetName(tabName) {
    return `'${String(tabName || '').replace(/'/g, "''")}'`;
  },

  async _fetchGoogleJson(url, options = {}) {
    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');

    const buildOptions = () => ({
      ...options,
      headers: {
        ...(options.headers || {}),
        'Authorization': `Bearer ${this.accessToken}`,
      },
    });

    let resp = await fetch(url, buildOptions());

    if (resp.status === 401) {
      try {
        await this.requestToken(true);
      } catch {
        throw new Error('Session หมดอายุ กรุณาเข้าสู่ระบบอีกครั้ง');
      }
      resp = await fetch(url, buildOptions());
    }

    if (!resp.ok) {
      let errMsg = `HTTP ${resp.status}`;
      try {
        const body = await resp.json();
        errMsg = body?.error?.message || errMsg;
      } catch { /* ignore */ }
      if (/drive\.googleapis\.com\/overview|Google Drive API has not been used|is disabled/i.test(errMsg)) {
        throw new Error('ยังไม่ได้เปิดใช้ Google Drive API ใน Google Cloud Project นี้ กรุณาไปที่ Google Cloud Console > APIs & Services > Library > Google Drive API แล้วกด Enable จากนั้นรอสักครู่และ Login เว็บใหม่');
      }
      if (resp.status === 403) {
        throw new Error(`ไม่มีสิทธิ์เข้าถึงหรือแก้ไข Sheet นี้ (${errMsg})`);
      }
      if (resp.status === 404) {
        throw new Error(`ไม่พบ Sheet หรือ Tab ที่ระบุ (${errMsg})`);
      }
      throw new Error(errMsg);
    }

    return await resp.json().catch(() => ({}));
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
    if (this.tokenProvider) return this.tokenProvider();
    return new Promise((resolve, reject) => {
      const client = this._ensureClient();
      let settled = false;
      const timeoutMs = silent ? 12000 : 120000;
      const timer = setTimeout(() => {
        if (settled) return;
        settled = true;
        reject(new Error(silent
          ? 'Session หมดอายุ กรุณาเข้าสู่ระบบอีกครั้ง'
          : 'หมดเวลารอ Google Sign in กรุณาลองเข้าสู่ระบบใหม่'));
      }, timeoutMs);
      const finish = (fn, value) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        fn(value);
      };
      client.callback = (resp) => {
        if (resp.error) {
          finish(reject, new Error(resp.error_description || resp.error));
          return;
        }
        this.accessToken = resp.access_token;
        this._userInfo = null; // reset cached user info
        finish(resolve, resp);
      };
      client.error_callback = (err) => {
        finish(reject, new Error(err?.message || err?.type || 'Google Sign in ไม่สำเร็จ'));
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
  async fetchSheetData(sheetId, tabName = 'Sheet1', options = {}) {
    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');

    const range = encodeURIComponent(tabName);
    const valueRenderOption = options.valueRenderOption === 'UNFORMATTED_VALUE'
      ? 'UNFORMATTED_VALUE'
      : 'FORMATTED_VALUE';
    const params = `valueRenderOption=${valueRenderOption}&dateTimeRenderOption=FORMATTED_STRING`;
    const url   = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}?${params}`;

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

  async fetchSheetRange(sheetId, rangeA1) {
    const range = encodeURIComponent(rangeA1);
    const params = 'valueRenderOption=FORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING';
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}?${params}`;
    const data = await this._fetchGoogleJson(url);
    return data.values || [];
  },

  /* ── Get spreadsheet metadata (list of tab names) ── */
  async getSheetTabs(sheetId) {
    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');

    const url  = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties(title,sheetId),properties.title`;
    let data;
    for (let attempt = 1; attempt <= 3; attempt += 1) {
      try {
        data = await this._fetchGoogleJson(url);
        break;
      } catch (err) {
        if (attempt === 3 || !/429|500|502|503|504/.test(String(err?.message || ''))) throw err;
        await new Promise(resolve => setTimeout(resolve, attempt * 700));
      }
    }
    return {
      title: data.properties?.title || '',
      tabs:  (data.sheets || []).map(s => s.properties.title),
      sheetTabs: (data.sheets || []).map(s => ({
        title: s.properties.title,
        sheetId: s.properties.sheetId,
      })),
    };
  },

  async addSheetTab(sheetId, tabName) {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}:batchUpdate`;
    return await this._fetchGoogleJson(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        requests: [
          { addSheet: { properties: { title: tabName } } },
        ],
      }),
    });
  },

  async clearSheetValues(sheetId, tabName) {
    const range = encodeURIComponent(`${this._quoteSheetName(tabName)}!A:ZZZ`);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}:clear`;
    return await this._fetchGoogleJson(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({}),
    });
  },

  async updateSheetValues(sheetId, tabName, rows) {
    const range = encodeURIComponent(`${this._quoteSheetName(tabName)}!A1`);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}?valueInputOption=RAW`;
    return await this._fetchGoogleJson(url, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        range: `${this._quoteSheetName(tabName)}!A1`,
        majorDimension: 'ROWS',
        values: rows,
      }),
    });
  },

  async updateSheetRange(sheetId, rangeA1, rows, valueInputOption = 'RAW') {
    const range = encodeURIComponent(rangeA1);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}?valueInputOption=${encodeURIComponent(valueInputOption)}`;
    return await this._fetchGoogleJson(url, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        range: rangeA1,
        majorDimension: 'ROWS',
        values: rows,
      }),
    });
  },

  async batchUpdateSheetRanges(sheetId, updates) {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchUpdate`;
    return await this._fetchGoogleJson(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        valueInputOption: 'RAW',
        data: updates.map(update => ({
          range: update.range,
          majorDimension: 'ROWS',
          values: update.values,
        })),
      }),
    });
  },

  async appendSheetValues(sheetId, tabName, rows) {
    const range = encodeURIComponent(`${this._quoteSheetName(tabName)}!A1`);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
    return await this._fetchGoogleJson(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        majorDimension: 'ROWS',
        values: rows,
      }),
    });
  },

	  async listDriveFolderFiles(folderId) {
	    const query = encodeURIComponent(`'${folderId}' in parents and trashed = false`);
	    const fields = encodeURIComponent('files(id,name,mimeType,modifiedTime,size,webViewLink),nextPageToken');
	    const orderBy = encodeURIComponent('modifiedTime desc,name');
	    const url = `https://www.googleapis.com/drive/v3/files?q=${query}&fields=${fields}&orderBy=${orderBy}&pageSize=100&supportsAllDrives=true&includeItemsFromAllDrives=true`;
	    const data = await this._fetchGoogleJson(url);
	    return data.files || [];
	  },

	  async findDriveFolder(parentFolderId, folderName) {
	    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');
	    const safeName = String(folderName || '').trim();
	    if (!safeName) throw new Error('ไม่ได้ระบุชื่อโฟลเดอร์');
	    const escapedName = safeName.replace(/'/g, "\\'");
	    const query = encodeURIComponent(`'${parentFolderId}' in parents and name='${escapedName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`);
	    const fields = encodeURIComponent('files(id,name,webViewLink),nextPageToken');
	    const url = `https://www.googleapis.com/drive/v3/files?q=${query}&fields=${fields}&pageSize=1&supportsAllDrives=true&includeItemsFromAllDrives=true`;
	    const data = await this._fetchGoogleJson(url);
	    return data.files?.[0] || null;
	  },

	  async resolveExistingDriveFolderPath(rootFolderId, pathSegments = []) {
	    let current = { id: rootFolderId };
	    for (const segment of pathSegments) {
	      const next = await this.findDriveFolder(current.id, segment);
	      if (!next) {
	        throw new Error(`ไม่พบโฟลเดอร์ Drive: ${pathSegments.join('/')}`);
	      }
	      current = next;
	    }
	    return current;
	  },

	  async findDriveFileInFolder(folderId, fileName) {
	    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');
	    const safeName = String(fileName || '').trim();
	    if (!safeName) throw new Error('ไม่ได้ระบุชื่อไฟล์');
	    const escapedName = safeName.replace(/'/g, "\\'");
	    const query = encodeURIComponent(`'${folderId}' in parents and name='${escapedName}' and trashed=false`);
	    const fields = encodeURIComponent('files(id,name,mimeType,modifiedTime,webViewLink),nextPageToken');
	    const url = `https://www.googleapis.com/drive/v3/files?q=${query}&fields=${fields}&pageSize=1&supportsAllDrives=true&includeItemsFromAllDrives=true`;
	    const data = await this._fetchGoogleJson(url);
	    return data.files?.[0] || null;
	  },

	  async fetchDriveJsonRows(rootFolderId, pathSegments, fileName) {
	    const folder = await this.resolveExistingDriveFolderPath(rootFolderId, pathSegments);
	    const file = await this.findDriveFileInFolder(folder.id, fileName);
	    if (!file) {
	      throw new Error(`ไม่พบไฟล์ Drive: ${[...pathSegments, fileName].join('/')}`);
	    }
	    const buffer = await this.downloadDriveFile(file.id);
	    const text = new TextDecoder('utf-8').decode(buffer);
	    const payload = JSON.parse(text);
	    const rows = Array.isArray(payload) ? payload : payload?.values;
	    if (!Array.isArray(rows)) {
	      throw new Error(`รูปแบบ JSON ไม่ถูกต้อง: ${fileName}`);
	    }
	    return rows;
	  },

	  async findOrCreateDriveFolder(parentFolderId, folderName) {
	    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');
	    const safeName = String(folderName || '').trim();
	    if (!safeName) throw new Error('ไม่ได้ระบุชื่อโฟลเดอร์');
	    const escapedName = safeName.replace(/'/g, "\\'");
	    const query = encodeURIComponent(`'${parentFolderId}' in parents and name='${escapedName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`);
	    const fields = encodeURIComponent('files(id,name,webViewLink),nextPageToken');
	    const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${query}&fields=${fields}&pageSize=1&supportsAllDrives=true&includeItemsFromAllDrives=true`;
	    const existing = await this._fetchGoogleJson(searchUrl);
	    if (existing.files?.[0]) return existing.files[0];

	    const createUrl = 'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&fields=id,name,webViewLink';
	    return await this._fetchGoogleJson(createUrl, {
	      method: 'POST',
	      headers: { 'Content-Type': 'application/json; charset=utf-8' },
	      body: JSON.stringify({
	        name: safeName,
	        mimeType: 'application/vnd.google-apps.folder',
	        parents: [parentFolderId],
	      }),
	    });
	  },

	  async resolveDriveFolderPath(rootFolderId, pathSegments = []) {
	    let current = { id: rootFolderId };
	    for (const segment of pathSegments) {
	      current = await this.findOrCreateDriveFolder(current.id, segment);
	    }
	    return current;
	  },

	  async uploadJsonToDriveFolder(folderId, fileName, payload) {
	    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');
	    const safeName = String(fileName || 'data_preparation.json').trim() || 'data_preparation.json';
	    const query = encodeURIComponent(`'${folderId}' in parents and name='${safeName.replace(/'/g, "\\'")}' and trashed=false`);
	    const fields = encodeURIComponent('files(id,name,webViewLink),nextPageToken');
	    const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${query}&fields=${fields}&pageSize=1&supportsAllDrives=true&includeItemsFromAllDrives=true`;
	    const existing = await this._fetchGoogleJson(searchUrl);
	    const fileId = existing.files?.[0]?.id || '';
	    const jsonText = JSON.stringify(payload, null, 2);

	    if (fileId) {
	      const updateUrl = `https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(fileId)}?uploadType=media&supportsAllDrives=true&fields=id,name,webViewLink`;
	      return await this._fetchGoogleJson(updateUrl, {
	        method: 'PATCH',
	        headers: { 'Content-Type': 'application/json; charset=utf-8' },
	        body: jsonText,
	      });
	    }

	    const boundary = `fund_json_${Date.now()}`;
	    const metadata = {
	      name: safeName,
	      mimeType: 'application/json',
	      parents: [folderId],
	    };
	    const body = [
	      `--${boundary}`,
	      'Content-Type: application/json; charset=UTF-8',
	      '',
	      JSON.stringify(metadata),
	      `--${boundary}`,
	      'Content-Type: application/json; charset=UTF-8',
	      '',
	      jsonText,
	      `--${boundary}--`,
	      '',
	    ].join('\r\n');
	    const createUrl = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name,webViewLink';
	    return await this._fetchGoogleJson(createUrl, {
	      method: 'POST',
	      headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
	      body,
	    });
	  },
	
	  async downloadDriveFile(fileId) {
    if (!this.accessToken) throw new Error('ยังไม่ได้เข้าสู่ระบบ');

    const url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media&supportsAllDrives=true`;
    const buildOptions = () => ({
      headers: { 'Authorization': `Bearer ${this.accessToken}` },
    });

    let resp = await fetch(url, buildOptions());
    if (resp.status === 401) {
      try {
        await this.requestToken(true);
      } catch {
        throw new Error('Session หมดอายุ กรุณาเข้าสู่ระบบอีกครั้ง');
      }
      resp = await fetch(url, buildOptions());
    }

    if (!resp.ok) {
      let errMsg = `HTTP ${resp.status}`;
      try {
        const body = await resp.json();
        errMsg = body?.error?.message || errMsg;
      } catch { /* ignore */ }
      if (resp.status === 403) {
        throw new Error(`ไม่มีสิทธิ์ดาวน์โหลดไฟล์จาก Drive (${errMsg})`);
      }
      if (resp.status === 404) {
        throw new Error(`ไม่พบไฟล์ใน Drive (${errMsg})`);
      }
      throw new Error(errMsg);
    }

    return await resp.arrayBuffer();
  },

  /* ── Sign out ── */
  signOut() {
    if (this.accessToken) {
      google.accounts.oauth2.revoke(this.accessToken, () => {});
    }
    this.accessToken = null;
    this._userInfo   = null;
    this.tokenClient = null;
    this.tokenProvider = null;
  },
};
