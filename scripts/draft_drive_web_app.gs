/*
 * Fund List Tool Draft JSON API
 *
 * Deploy as a Google Apps Script Web App:
 * - Execute as: Me
 * - Who has access: Anyone with Google account / organization users
 */

const DRAFT_ROOT_FOLDER_ID = '1TwC7V8gpcDswftoweT-VcG89OnIVW5Ol';
const DRAFT_API_KEY = 'change-this-draft-api-key';

function doGet(e) {
  return handleDraftRequest_(e);
}

function doPost(e) {
  return handleDraftRequest_(e);
}

function handleDraftRequest_(e) {
  let request = {};
  try {
    request = parseRequest_(e);
    if (DRAFT_API_KEY && request.key !== DRAFT_API_KEY) {
      return jsonResponse_({ ok: false, error: 'Unauthorized' }, request.callback);
    }

    const action = String(request.action || 'list').toLowerCase();
    if (action === 'list') return jsonResponse_({ ok: true, drafts: listDrafts_(request.quarter) }, request.callback);
    if (action === 'get') return jsonResponse_({ ok: true, draft: getDraft_(request.id, request.quarter) }, request.callback);
    if (action === 'save') return jsonResponse_(saveDraft_(request.draft || request), request.callback);
    if (action === 'delete') return jsonResponse_(deleteDraft_(request.id, request.quarter), request.callback);
    if (action === 'ping') return jsonResponse_({ ok: true, message: 'Draft API ready' }, request.callback);

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
    const body = `${callback}(${JSON.stringify(payload)});`;
    return ContentService
      .createTextOutput(body)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeSlug_(value, fallback) {
  const text = String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '-')
    .replace(/[^a-z0-9ก-๙._-]+/g, '')
    .replace(/^[._-]+|[._-]+$/g, '');
  return (text || fallback || 'draft').slice(0, 80);
}

function titleSlug_(value, fallback) {
  const text = String(value || '')
    .trim()
    .replace(/&/g, ' and ')
    .replace(/[@.].*$/g, '')
    .replace(/\s+/g, '-')
    .replace(/[^A-Za-z0-9ก-๙._-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^[._-]+|[._-]+$/g, '');
  return (text || fallback || 'draft').slice(0, 80);
}

function normalizeQuarter_(value) {
  const quarter = String(value || '').trim().toUpperCase();
  return /^\d{4}-Q[1-4]$/.test(quarter) ? quarter : 'unspecified';
}

function quarterYear_(quarter) {
  return /^\d{4}-Q[1-4]$/.test(quarter) ? quarter.slice(0, 4) : 'unspecified';
}

function getOrCreateChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function rootFolder_() {
  return DriveApp.getFolderById(DRAFT_ROOT_FOLDER_ID);
}

function quarterFolder_(quarter, create) {
  const normalizedQuarter = normalizeQuarter_(quarter);
  const root = rootFolder_();
  const yearName = quarterYear_(normalizedQuarter);

  if (create) {
    const yearFolder = getOrCreateChildFolder_(root, yearName);
    return getOrCreateChildFolder_(yearFolder, normalizedQuarter);
  }

  const years = root.getFoldersByName(yearName);
  if (!years.hasNext()) return null;
  const quarters = years.next().getFoldersByName(normalizedQuarter);
  return quarters.hasNext() ? quarters.next() : null;
}

function fileNameForDraftId_(id) {
  return `${safeSlug_(id, 'draft')}.json`;
}

function fileNameForDraft_(draftId, draft) {
  const quarter = normalizeQuarter_(draft.currentQuarter || draft.dataQuarter || draft.quarter);
  const category = titleSlug_(draft.avpCategory || draft.filters?.selectFundFilters?.category || draft.asset, 'All-AVP');
  const type = titleSlug_(draft.fundType || draft.filters?.selectFundFilters?.type, 'All-Type');
  const dateText = titleSlug_(draft.userDate || String(draft.createdAt || '').slice(0, 10), 'No-Date');
  const fundCount = Array.isArray(draft.selectedKeys)
    ? draft.selectedKeys.length
    : Object.keys(draft.selectedFunds || {}).length;
  const highlightCount = Object.keys(draft.highlights || {}).length;
  const author = titleSlug_(draft.authorEmail || draft.author || draft.authorName, 'unknown');
  const id = safeSlug_(draftId, 'draft');
  return `${quarter}__${category}__${type}__${dateText}__${fundCount}F-${highlightCount}H__${author}__${id}.json`;
}

function fileNameSuffixForDraftId_(id) {
  return `__${safeSlug_(id, 'draft')}.json`;
}

function readJsonFile_(file) {
  return JSON.parse(file.getBlob().getDataAsString('UTF-8'));
}

function listDrafts_(quarter) {
  const drafts = [];
  if (quarter) {
    collectDraftsFromFolder_(quarterFolder_(quarter, false), drafts);
  } else {
    const root = rootFolder_();
    const years = root.getFolders();
    while (years.hasNext()) {
      const yearFolder = years.next();
      const quarters = yearFolder.getFolders();
      while (quarters.hasNext()) collectDraftsFromFolder_(quarters.next(), drafts);
    }
  }

  drafts.sort((a, b) => String(b.createdAt || b.updatedAt || '').localeCompare(String(a.createdAt || a.updatedAt || '')));
  return drafts;
}

function collectDraftsFromFolder_(folder, drafts) {
  if (!folder) return;
  const files = folder.getFilesByType('application/json');
  while (files.hasNext()) {
    const file = files.next();
    try {
      const draft = readJsonFile_(file);
      if (draft && typeof draft === 'object') {
        draft.driveFileId = file.getId();
        draft.driveFileName = file.getName();
        drafts.push(draft);
      }
    } catch (err) {
      // Skip malformed JSON files so one bad draft does not break the list.
    }
  }
}

function getDraft_(id, quarter) {
  if (!id) throw new Error('id is required');
  const file = findDraftFile_(id, quarter);
  if (!file) throw new Error(`Draft not found: ${id}`);
  const draft = readJsonFile_(file);
  draft.driveFileId = file.getId();
  draft.driveFileName = file.getName();
  return draft;
}

function findDraftFile_(id, quarter) {
  const legacyName = fileNameForDraftId_(id);
  const suffix = fileNameSuffixForDraftId_(id);
  const folders = [];

  if (quarter) {
    const folder = quarterFolder_(quarter, false);
    if (folder) folders.push(folder);
  } else {
    const root = rootFolder_();
    const years = root.getFolders();
    while (years.hasNext()) {
      const quarters = years.next().getFolders();
      while (quarters.hasNext()) folders.push(quarters.next());
    }
  }

  for (let i = 0; i < folders.length; i += 1) {
    const files = folders[i].getFilesByName(legacyName);
    if (files.hasNext()) return files.next();
  }
  for (let i = 0; i < folders.length; i += 1) {
    const files = folders[i].getFilesByType('application/json');
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().endsWith(suffix)) return file;
    }
  }
  return null;
}

function saveDraft_(draft) {
  if (!draft || typeof draft !== 'object') throw new Error('draft is required');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const now = new Date().toISOString();
    const draftId = safeSlug_(draft.id || Date.now(), 'draft');
    const quarter = normalizeQuarter_(draft.currentQuarter || draft.dataQuarter || draft.quarter);
    const folder = quarterFolder_(quarter, true);
    const payload = Object.assign({}, draft, {
      id: draftId,
      currentQuarter: quarter,
      createdAt: draft.createdAt || now,
      updatedAt: now,
    });
    const fileName = fileNameForDraft_(draftId, payload);

    const existing = findDraftFile_(draftId, quarter);
    const blob = Utilities.newBlob(JSON.stringify(payload, null, 2), 'application/json', fileName);
    let file;

    if (existing) {
      existing.setName(fileName);
      existing.setContent(blob.getDataAsString('UTF-8'));
      file = existing;
    } else {
      file = folder.createFile(blob);
    }

    return {
      ok: true,
      draft: payload,
      driveUploaded: true,
      driveFileId: file.getId(),
      driveFileName: file.getName(),
      quarter,
      path: `${quarterYear_(quarter)}/${quarter}/${fileName}`,
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteDraft_(id, quarter) {
  if (!id) throw new Error('id is required');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const file = findDraftFile_(id, quarter);
    if (file) file.setTrashed(true);
    return {
      ok: true,
      removed: Boolean(file),
      id,
      quarter: quarter || '',
    };
  } finally {
    lock.releaseLock();
  }
}
