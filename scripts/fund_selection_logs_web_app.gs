/*
 * Fund Selection Logs JSON API
 *
 * Deploy as a Google Apps Script Web App:
 * - Execute as: Me
 * - Who has access: Anyone with Google account / organization users
 */

const FUND_SELECTION_LOGS_ROOT_FOLDER_ID = '12ciJQq-dpBr-DpdnzXCOXqtW_ijctJN6';
const FUND_SELECTION_LOGS_API_KEY = 'change-this-fund-selection-logs-api-key';

function doGet(e) {
  return handleLogsRequest_(e);
}

function doPost(e) {
  return handleLogsRequest_(e);
}

function handleLogsRequest_(e) {
  let request = {};
  try {
    request = parseRequest_(e);
    if (FUND_SELECTION_LOGS_API_KEY && request.key !== FUND_SELECTION_LOGS_API_KEY) {
      return jsonResponse_({ ok: false, error: 'Unauthorized' }, request.callback);
    }

    const action = String(request.action || 'get').toLowerCase();
    if (action === 'get') return jsonResponse_(getLog_(request.quarter), request.callback);
    if (action === 'save') {
      return jsonResponse_(saveLog_(request.log || request, request.baseRevision), request.callback);
    }
    if (action === 'delete') return jsonResponse_(deleteLog_(request.quarter), request.callback);
    if (action === 'ping') return jsonResponse_({ ok: true, message: 'Fund Selection Logs API ready' }, request.callback);

    return jsonResponse_({ ok: false, error: `Unknown action: ${action}` }, request.callback);
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message || String(err) }, request.callback);
  }
}

function parseRequest_(e) {
  const params = Object.assign({}, e && e.parameter ? e.parameter : {});
  if (params.payloadGzipB64) {
    try {
      const bytes = Utilities.base64DecodeWebSafe(params.payloadGzipB64);
      const blob = Utilities.newBlob(bytes, 'application/gzip', 'payload.json.gz');
      const text = Utilities.ungzip(blob).getDataAsString('UTF-8');
      return Object.assign(params, JSON.parse(text));
    } catch (err) {
      return params;
    }
  }
  if (params.payload) {
    try {
      return Object.assign(params, JSON.parse(params.payload));
    } catch (err) {
      return params;
    }
  }
  const body = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (!body) return params;

  try {
    return Object.assign(params, JSON.parse(body));
  } catch (err) {
    return params;
  }
}

function jsonResponse_(payload, callback) {
  callback = typeof callback === 'string' ? callback : '';
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback)) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeQuarter_(value) {
  const quarter = String(value || '').trim().toUpperCase();
  if (!/^\d{4}-Q[1-4]$/.test(quarter)) {
    throw new Error('quarter must be in YYYY-QN format, e.g. 2026-Q1');
  }
  return quarter;
}

function fileNameForQuarter_(quarter) {
  return `Fund Selection Logs - ${normalizeQuarter_(quarter)}.json`;
}

function rootFolder_() {
  return DriveApp.getFolderById(FUND_SELECTION_LOGS_ROOT_FOLDER_ID);
}

function findLogFile_(quarter) {
  const files = rootFolder_().getFilesByName(fileNameForQuarter_(quarter));
  return files.hasNext() ? files.next() : null;
}

function readJsonFile_(file) {
  return JSON.parse(file.getBlob().getDataAsString('UTF-8'));
}

function getLog_(quarter) {
  quarter = normalizeQuarter_(quarter);
  const file = findLogFile_(quarter);
  if (!file) {
    return {
      ok: true,
      log: null,
      notFound: true,
      source: 'New file',
      fileName: fileNameForQuarter_(quarter),
      quarter,
    };
  }

  const log = readJsonFile_(file);
  log.driveFileId = file.getId();
  log.driveFileName = file.getName();
  return {
    ok: true,
    log,
    source: `Google Drive: ${file.getName()}`,
    drive: {
      fileId: file.getId(),
      fileName: file.getName(),
      folderId: FUND_SELECTION_LOGS_ROOT_FOLDER_ID,
    },
    quarter,
  };
}

function saveLog_(log, baseRevision) {
  if (!log || typeof log !== 'object') throw new Error('log is required');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const quarter = normalizeQuarter_(log.quarter);
    const fileName = fileNameForQuarter_(quarter);
    const existing = findLogFile_(quarter);
    const current = existing ? readJsonFile_(existing) : {};
    const currentRevision = Number(current.revision || 0);

    if (
      baseRevision !== undefined
      && baseRevision !== null
      && currentRevision
      && Number(baseRevision) !== currentRevision
    ) {
      return {
        ok: false,
        conflict: true,
        error: 'revision conflict',
        currentRevision,
        currentLog: current,
      };
    }

    const now = new Date().toISOString();
    const payload = Object.assign({}, log, {
      schemaVersion: Number(log.schemaVersion || 1),
      quarter,
      revision: currentRevision + 1,
      createdAt: log.createdAt || current.createdAt || now,
      updatedAt: now,
    });

    const content = JSON.stringify(payload, null, 2);
    let file;
    if (existing) {
      existing.setName(fileName);
      existing.setContent(content);
      file = existing;
    } else {
      file = rootFolder_().createFile(
        Utilities.newBlob(content, 'application/json', fileName)
      );
    }

    return {
      ok: true,
      log: payload,
      driveUploaded: true,
      drive: {
        fileId: file.getId(),
        fileName: file.getName(),
        folderId: FUND_SELECTION_LOGS_ROOT_FOLDER_ID,
      },
      quarter,
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteLog_(quarter) {
  quarter = normalizeQuarter_(quarter);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const file = findLogFile_(quarter);
    if (file) file.setTrashed(true);
    return {
      ok: true,
      removed: Boolean(file),
      quarter,
    };
  } finally {
    lock.releaseLock();
  }
}
