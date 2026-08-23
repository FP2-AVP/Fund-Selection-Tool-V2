/* ============================================================
   FT Historical Prices Web App
   ------------------------------------------------------------
   Deploy as a Google Apps Script web app.

   Script Properties to set:
   - API_SECRET_KEY: optional, must match CONFIG.FT_HISTORICAL_API_SECRET_KEY
   - GITHUB_TOKEN: optional fallback GitHub PAT with Actions workflow dispatch permission
   - GITHUB_REPO: optional, default FP2-AVP/Fund-Selection-Tool-V2
   - GITHUB_WORKFLOW: optional, default ft-historical-prices-database.yml
   ============================================================ */

const FT_HISTORICAL_VERSION = '2026-08-23-ytd-batch-status';
const DEFAULT_DRIVE_FOLDER_ID = '1Locig3aW7hVs0SxCoeFpa76OJ0XcF1rg';
const DEFAULT_FILE_NAME = 'ft_historical_prices_database.json';
const DEFAULT_GITHUB_REPO = 'FP2-AVP/Fund-Selection-Tool-V2';
const DEFAULT_GITHUB_WORKFLOW = 'ft-historical-prices-database.yml';

function doGet(e) {
  return handleFtHistoricalRequest_(e);
}

function doPost(e) {
  return handleFtHistoricalRequest_(e);
}

function handleFtHistoricalRequest_(e) {
  try {
    const params = (e && e.parameter) || {};
    const body = parseBody_(e);
    const action = String(params.action || body.action || 'ping').trim();
    if (action !== 'ping' && !isAuthorized_(params.key || body.key || '')) {
      return json_({ ok: false, error: 'Unauthorized' });
    }

    if (action === 'ping') {
      return json_({
        ok: true,
        message: 'FT Historical Prices API ready',
        version: FT_HISTORICAL_VERSION,
        driveFolderId: DEFAULT_DRIVE_FOLDER_ID,
        fileName: DEFAULT_FILE_NAME,
      });
    }

    if (action === 'database') {
      return readFtDatabase_(params, body);
    }

    if (action === 'uploadDatabase') {
      return uploadFtDatabase_(params, body);
    }

    if (action === 'uploadFile') {
      return uploadFtFile_(params, body);
    }

    if (action === 'downloadFile') {
      return downloadFtFile_(params, body);
    }

    if (action === 'sync') {
      return triggerFtWorkflow_(params, body);
    }

    if (action === 'workflowStatus') {
      return readFtWorkflowStatus_(params, body);
    }

    return json_({ ok: false, error: `Unknown action: ${action}` });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function parseBody_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch (_) {
    return {};
  }
}

function scriptProp_(name, fallback) {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  return value === null || value === undefined || value === '' ? fallback : value;
}

function isAuthorized_(key) {
  const expected = scriptProp_('API_SECRET_KEY', '');
  return !expected || String(key || '') === expected;
}

function newestFileByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  let newest = null;
  while (files.hasNext()) {
    const file = files.next();
    if (!newest || file.getLastUpdated().getTime() > newest.getLastUpdated().getTime()) {
      newest = file;
    }
  }
  return newest;
}

function upsertNewestFile_(folder, fileName, content, mimeType) {
  const files = folder.getFilesByName(fileName);
  let newest = null;
  const duplicates = [];
  while (files.hasNext()) {
    const file = files.next();
    if (!newest || file.getLastUpdated().getTime() > newest.getLastUpdated().getTime()) {
      if (newest) duplicates.push(newest);
      newest = file;
    } else {
      duplicates.push(file);
    }
  }
  duplicates.forEach(file => file.setTrashed(true));
  if (newest) {
    newest.setContent(content);
    return newest;
  }
  return folder.createFile(Utilities.newBlob(content, mimeType || 'application/octet-stream', fileName));
}

function parseCsvLine_(line) {
  const values = [];
  let value = '';
  let quoted = false;
  for (let idx = 0; idx < line.length; idx += 1) {
    const ch = line[idx];
    const next = line[idx + 1];
    if (ch === '"' && quoted && next === '"') {
      value += '"';
      idx += 1;
    } else if (ch === '"') {
      quoted = !quoted;
    } else if (ch === ',' && !quoted) {
      values.push(value);
      value = '';
    } else {
      value += ch;
    }
  }
  values.push(value);
  return values;
}

function priceCsvSummary_(file) {
  const text = file.getBlob().getDataAsString('UTF-8').trim();
  if (!text) return null;
  const lines = text.split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return null;
  const headers = parseCsvLine_(lines[0]).map(header => String(header || '').trim().toLowerCase());
  const symbolIdx = headers.indexOf('symbol');
  const dateIdx = headers.indexOf('date');
  if (symbolIdx < 0 || dateIdx < 0) return null;

  let symbol = '';
  let startDate = '';
  let endDate = '';
  let rowCount = 0;
  for (let idx = 1; idx < lines.length; idx += 1) {
    const cells = parseCsvLine_(lines[idx]);
    const rowSymbol = String(cells[symbolIdx] || '').trim();
    const date = String(cells[dateIdx] || '').trim();
    if (!rowSymbol || !date) continue;
    if (!symbol) symbol = rowSymbol;
    rowCount += 1;
    if (!startDate || date < startDate) startDate = date;
    if (!endDate || date > endDate) endDate = date;
  }
  if (!symbol || !rowCount) return null;
  return {
    symbol,
    rowCount,
    rows: rowCount,
    startDate,
    endDate,
    start: startDate,
    end: endDate,
    source: `Google Drive prices/${file.getName()}`,
  };
}

function augmentPayloadFromPricesFolder_(folder, payload) {
  const priceFolders = folder.getFoldersByName('prices');
  if (!priceFolders.hasNext()) return payload;
  const pricesFolder = priceFolders.next();
  const existing = {};
  payload.symbols = Array.isArray(payload.symbols) ? payload.symbols : [];
  payload.symbols.forEach(item => {
    if (item && item.symbol) existing[String(item.symbol).trim().toUpperCase()] = item;
  });

  let added = 0;
  let updated = 0;
  const files = pricesFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (!/\.csv$/i.test(file.getName())) continue;
    const summary = priceCsvSummary_(file);
    if (!summary) continue;
    const key = String(summary.symbol).trim().toUpperCase();
    if (existing[key]) {
      const currentRows = Number(existing[key].rowCount || existing[key].rows || 0);
      if (!currentRows || summary.rowCount > currentRows) {
        existing[key].rowCount = summary.rowCount;
        existing[key].rows = summary.rowCount;
        existing[key].startDate = summary.startDate || existing[key].startDate || '';
        existing[key].endDate = summary.endDate || existing[key].endDate || '';
        updated += 1;
      }
    } else {
      payload.symbols.push(summary);
      existing[key] = summary;
      added += 1;
    }
  }
  payload.counts = payload.counts || {};
  payload.counts.symbols = payload.symbols.length;
  payload.drivePriceFallback = { added, updated };
  return payload;
}

function readFtDatabase_(params, body) {
  const folderId = String(params.folderId || body.folderId || DEFAULT_DRIVE_FOLDER_ID).trim();
  const fileName = String(params.fileName || body.fileName || DEFAULT_FILE_NAME).trim();
  const folder = DriveApp.getFolderById(folderId);
  const file = newestFileByName_(folder, fileName);
  if (!file) {
    return json_({
      ok: false,
      notFound: true,
      error: `File not found: ${fileName}`,
      driveFolderId: folderId,
      fileName,
    });
  }

  const payload = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
  augmentPayloadFromPricesFolder_(folder, payload);
  payload.ok = true;
  payload.source = `Google Drive: ${fileName}`;
  payload.drive = {
    folderId,
    fileId: file.getId(),
    fileName,
    updatedAt: file.getLastUpdated().toISOString(),
    url: file.getUrl(),
  };
  return json_(payload);
}

function uploadFtDatabase_(params, body) {
  const folderId = String(params.folderId || body.folderId || DEFAULT_DRIVE_FOLDER_ID).trim();
  const fileName = String(params.fileName || body.fileName || DEFAULT_FILE_NAME).trim();
  const database = body.database || body.payload;
  if (!database || !Array.isArray(database.symbols)) {
    throw new Error('Invalid FT historical database payload');
  }

  const folder = DriveApp.getFolderById(folderId);
  const blob = Utilities.newBlob(
    JSON.stringify(database, null, 2) + '\n',
    'application/json',
    fileName,
  );
  const file = upsertNewestFile_(folder, fileName, blob.getDataAsString('UTF-8'), 'application/json');

  return json_({
    ok: true,
    message: 'อัปโหลด FT Historical Prices Database ไป Google Drive แล้ว',
    version: FT_HISTORICAL_VERSION,
    driveUploaded: true,
    drive: {
      folderId,
      fileId: file.getId(),
      fileName,
      updatedAt: file.getLastUpdated().toISOString(),
      url: file.getUrl(),
    },
    counts: database.counts || {},
    generatedAt: database.generated_at || database.generatedAt || '',
  });
}

function uploadFtFile_(params, body) {
  const folderId = String(params.folderId || body.folderId || DEFAULT_DRIVE_FOLDER_ID).trim();
  const relativePath = normalizeRelativePath_(params.relativePath || body.relativePath || body.fileName || '');
  const fileName = String(params.fileName || body.fileName || lastPathPart_(relativePath)).trim();
  const mimeType = String(params.mimeType || body.mimeType || 'application/octet-stream').trim();
  const contentBase64 = String(body.contentBase64 || '').trim();
  if (!relativePath) throw new Error('Missing relativePath');
  if (!fileName) throw new Error('Missing fileName');
  if (!contentBase64) throw new Error('Missing contentBase64');

  const root = DriveApp.getFolderById(folderId);
  const parent = ensureParentFolder_(root, relativePath);
  const bytes = Utilities.base64Decode(contentBase64);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const files = parent.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
  const file = parent.createFile(blob);

  return json_({
    ok: true,
    message: 'อัปโหลดไฟล์ FT ไป Google Drive แล้ว',
    version: FT_HISTORICAL_VERSION,
    drive: {
      folderId,
      fileId: file.getId(),
      fileName,
      relativePath,
      updatedAt: file.getLastUpdated().toISOString(),
      url: file.getUrl(),
    },
  });
}

function downloadFtFile_(params, body) {
  const folderId = String(params.folderId || body.folderId || DEFAULT_DRIVE_FOLDER_ID).trim();
  const relativePath = normalizeRelativePath_(params.relativePath || body.relativePath || body.fileName || '');
  const fileName = String(params.fileName || body.fileName || lastPathPart_(relativePath)).trim();
  if (!relativePath && !fileName) throw new Error('Missing relativePath or fileName');

  const root = DriveApp.getFolderById(folderId);
  const file = findFileByRelativePath_(root, relativePath || fileName);
  if (!file) {
    return json_({
      ok: false,
      notFound: true,
      error: `File not found: ${relativePath || fileName}`,
      driveFolderId: folderId,
      relativePath: relativePath || fileName,
    });
  }

  const blob = file.getBlob();
  return json_({
    ok: true,
    fileName: file.getName(),
    relativePath: relativePath || file.getName(),
    mimeType: blob.getContentType() || 'application/octet-stream',
    contentBase64: Utilities.base64Encode(blob.getBytes()),
    drive: {
      folderId,
      fileId: file.getId(),
      updatedAt: file.getLastUpdated().toISOString(),
      url: file.getUrl(),
    },
  });
}

function triggerFtWorkflow_(params, body) {
  const token = String(body.githubToken || params.githubToken || scriptProp_('GITHUB_TOKEN', '')).trim();
  if (!token) {
    throw new Error('Missing GitHub token. ใส่ token ในหน้าเว็บ หรือเพิ่ม Apps Script property GITHUB_TOKEN');
  }

  const repo = scriptProp_('GITHUB_REPO', DEFAULT_GITHUB_REPO);
  const workflow = scriptProp_('GITHUB_WORKFLOW', DEFAULT_GITHUB_WORKFLOW);
  const quarter = String(params.quarter || body.quarter || '2026-Q1').trim();
  const symbol = String(params.symbol || body.symbol || 'IXN:PCQ:USD').trim();
  const symbols = String(params.symbols || body.symbols || '').trim();
  const allSymbolsFromDb = String(params.allSymbolsFromDb || params.all_symbols_from_db || body.allSymbolsFromDb || body.all_symbols_from_db || 'false').trim();
  const ftUrl = String(params.url || body.url || '').trim();
  const startDate = String(params.startDate || body.startDate || '').trim();
  const endDate = String(params.endDate || body.endDate || '').trim();
  const runPrices = String(params.runPrices || body.runPrices || 'true').trim();
  const runQualitative = String(params.runQualitative || body.runQualitative || 'true').trim();
  const continueOnError = String(params.continueOnError || params.continue_on_error || body.continueOnError || body.continue_on_error || 'true').trim();
  const sleepSeconds = String(params.sleepSeconds || params.sleep_seconds || body.sleepSeconds || body.sleep_seconds || '1').trim();
  const folderId = String(params.folderId || body.folderId || DEFAULT_DRIVE_FOLDER_ID).trim();
  const fileName = String(params.fileName || body.fileName || DEFAULT_FILE_NAME).trim();
  const callbackUrl = String(params.callbackUrl || body.callbackUrl || ScriptApp.getService().getUrl()).trim();
  const apiKey = String(params.key || body.key || '').trim();
  const url = `https://api.github.com/repos/${repo}/actions/workflows/${encodeURIComponent(workflow)}/dispatches`;
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
    payload: JSON.stringify({
      ref: 'main',
      inputs: {
        quarter,
        symbol,
        symbols,
        all_symbols_from_db: allSymbolsFromDb,
        ft_url: ftUrl,
        start_date: startDate,
        end_date: endDate,
        run_prices: runPrices,
        run_qualitative: runQualitative,
        drive_folder_id: folderId,
        file_name: fileName,
        app_script_url: callbackUrl,
        app_script_key: apiKey,
        continue_on_error: continueOnError,
        sleep_seconds: sleepSeconds,
      },
    }),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error(`GitHub workflow dispatch failed (${status}): ${text}`);
  }

  return json_({
    ok: true,
    message: 'สั่ง GitHub Actions ให้ดึง FT Historical Prices แล้ว',
    version: FT_HISTORICAL_VERSION,
    workflow,
    repo,
    quarter,
    symbol,
    symbols,
    allSymbolsFromDb,
    driveFolderId: folderId,
    fileName,
  });
}

function readFtWorkflowStatus_(params, body) {
  const token = String(body.githubToken || params.githubToken || scriptProp_('GITHUB_TOKEN', '')).trim();
  if (!token) {
    throw new Error('Missing GitHub token. เพิ่ม Apps Script property GITHUB_TOKEN เพื่ออ่านสถานะ GitHub Actions');
  }

  const repo = scriptProp_('GITHUB_REPO', DEFAULT_GITHUB_REPO);
  const workflow = scriptProp_('GITHUB_WORKFLOW', DEFAULT_GITHUB_WORKFLOW);
  const perPage = Math.max(1, Math.min(Number(params.perPage || body.perPage || 5), 10));
  const url = `https://api.github.com/repos/${repo}/actions/workflows/${encodeURIComponent(workflow)}/runs?branch=main&per_page=${perPage}`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error(`GitHub workflow status failed (${status}): ${text}`);
  }

  const payload = JSON.parse(text || '{}');
  const runs = (payload.workflow_runs || []).map(run => ({
    id: run.id,
    name: run.name,
    status: run.status,
    conclusion: run.conclusion || '',
    event: run.event,
    branch: run.head_branch,
    createdAt: run.created_at,
    updatedAt: run.updated_at,
    runStartedAt: run.run_started_at,
    htmlUrl: run.html_url,
  }));
  return json_({
    ok: true,
    version: FT_HISTORICAL_VERSION,
    repo,
    workflow,
    latestRun: runs[0] || null,
    runs,
  });
}

function normalizeRelativePath_(value) {
  return String(value || '')
    .replace(/\\/g, '/')
    .split('/')
    .map(part => part.trim())
    .filter(part => part && part !== '.' && part !== '..')
    .join('/');
}

function lastPathPart_(relativePath) {
  const parts = normalizeRelativePath_(relativePath).split('/').filter(Boolean);
  return parts.length ? parts[parts.length - 1] : '';
}

function ensureParentFolder_(root, relativePath) {
  const parts = normalizeRelativePath_(relativePath).split('/').filter(Boolean);
  parts.pop();
  return parts.reduce((folder, name) => {
    const children = folder.getFoldersByName(name);
    return children.hasNext() ? children.next() : folder.createFolder(name);
  }, root);
}

function findFileByRelativePath_(root, relativePath) {
  const parts = normalizeRelativePath_(relativePath).split('/').filter(Boolean);
  if (!parts.length) return null;
  const fileName = parts.pop();
  const folder = parts.reduce((current, name) => {
    if (!current) return null;
    const children = current.getFoldersByName(name);
    return children.hasNext() ? children.next() : null;
  }, root);
  if (!folder) return null;
  const files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}
