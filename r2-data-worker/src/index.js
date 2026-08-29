export default {
  async fetch(request, env) {
    try {
      if (request.method === 'OPTIONS') return cors(request, env, null, 204);
      const session = await readSession(request, env);
      if (!session) return cors(request, env, {ok:false,error:'Session expired'}, 401);
      const url = new URL(request.url);
      if (url.pathname === '/drafts' && request.method === 'GET') {
        return cors(request, env, await listDrafts(env, url.searchParams.get('quarter') || ''), 200);
      }
      if (url.pathname === '/drafts/rebuild-index' && request.method === 'POST') {
        return cors(request, env, await rebuildDraftIndex(env), 200);
      }
      if (url.pathname === '/drafts' && request.method === 'POST') {
        return cors(request, env, await saveDraft(request, env, session), 200);
      }
      if (url.pathname.startsWith('/drafts/') && request.method === 'GET') {
        return cors(request, env, await getDraft(env, decodeURIComponent(url.pathname.slice('/drafts/'.length)), url.searchParams.get('quarter') || ''), 200);
      }
      if (url.pathname.startsWith('/drafts/') && request.method === 'DELETE') {
        return cors(request, env, await deleteDraft(env, decodeURIComponent(url.pathname.slice('/drafts/'.length)), url.searchParams.get('quarter') || ''), 200);
      }
      if (url.pathname === '/master-fund-overrides' && request.method === 'GET') {
        return cors(request, env, await getMasterFundOverrides(env, url.searchParams.get('quarter') || ''), 200);
      }
      if (url.pathname === '/master-fund-overrides' && request.method === 'POST') {
        return cors(request, env, await saveMasterFundOverride(request, env, session), 200);
      }
      if (url.pathname.startsWith('/master-fund-overrides/') && request.method === 'DELETE') {
        return cors(request, env, await deleteMasterFundOverride(env, decodeURIComponent(url.pathname.slice('/master-fund-overrides/'.length)), url.searchParams.get('quarter') || ''), 200);
      }
      if (url.pathname === '/master-allocations' && request.method === 'GET') {
        return cors(request, env, await getMasterAllocations(env, url.searchParams.get('quarter') || ''), 200);
      }
      if (url.pathname === '/master-allocations' && request.method === 'POST') {
        return cors(request, env, await saveMasterAllocation(request, env, session), 200);
      }
      if (url.pathname.startsWith('/master-allocations/') && request.method === 'DELETE') {
        return cors(request, env, await deleteMasterAllocation(env, decodeURIComponent(url.pathname.slice('/master-allocations/'.length)), url.searchParams.get('quarter') || ''), 200);
      }
      if (url.pathname === '/report-data-overrides' && request.method === 'GET') {
        return cors(request, env, await getReportDataOverrides(env, url.searchParams.get('quarter') || ''), 200);
      }
      if (url.pathname === '/report-data-overrides' && request.method === 'POST') {
        return cors(request, env, await saveReportDataOverrides(request, env, session), 200);
      }
      if (url.pathname.startsWith('/report-data-overrides/') && request.method === 'DELETE') {
        const parts = url.pathname.slice('/report-data-overrides/'.length).split('/').map(decodeURIComponent);
        return cors(request, env, await deleteReportDataOverride(env, parts[0] || '', parts[1] || '', url.searchParams.get('quarter') || '', url.searchParams.get('section') || ''), 200);
      }
      if (url.pathname.startsWith('/ft/jobs/') && request.method === 'POST') {
        return cors(request, env, await createFtJob(request, env, session, decodeURIComponent(url.pathname.slice('/ft/jobs/'.length))), 202);
      }
      if (url.pathname.startsWith('/ft/jobs/') && request.method === 'GET') {
        return cors(request, env, await getFtJob(env, decodeURIComponent(url.pathname.slice('/ft/jobs/'.length))), 200);
      }
      const ftKey = ftObjectKey(url.pathname);
      if (ftKey && request.method === 'GET') {
        return r2Object(request, env, ftKey, ftKey.endsWith('/index.json') ? 'no-store' : 'private, max-age=300');
      }
      if (request.method !== 'GET') return cors(request, env, {ok:false,error:'Method not allowed'}, 405);
      const key = objectKey(url.pathname);
      if (!key) return cors(request, env, {ok:true,service:'fund-selection-r2-data'});
      const object = await env.FUND_DATA.get(key);
      if (!object) return cors(request, env, {ok:false,error:`Object not found: ${key}`}, 404);
      const headers = new Headers();
      object.writeHttpMetadata(headers);
      headers.set('Content-Type', 'application/json; charset=utf-8');
      headers.set('Cache-Control', key === 'manifest.json' ? 'no-store' : 'private, max-age=300');
      addCors(headers, request, env);
      return new Response(object.body, {headers});
    } catch (error) {
      return cors(request, env, {ok:false,error:String(error?.message || error)}, 500);
    }
  },
};

function normalizeQuarter(value) {
  const quarter = String(value || '').trim().toUpperCase();
  if (!/^\d{4}-Q[1-4]$/.test(quarter)) throw new Error('quarter must look like 2026-Q3');
  return quarter;
}

function safeSlug(value, fallback='draft') {
  const slug = String(value || '').trim().replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '');
  return slug || fallback;
}

function masterFundOverrideObjectKey(quarter) {
  return `Data/${quarter.slice(0, 4)}/${quarter}/overrides/master_fund_overrides.json`;
}

function masterFundOverrideKey(profile) {
  const fundId = String(profile?.masterFundId || '').trim().toUpperCase();
  const isin = String(profile?.isin || '').trim().toUpperCase();
  if (fundId) return fundId;
  if (isin) return `isin:${isin}`;
  throw new Error('Master FundId or ISIN is required');
}

function normalizeMasterFundOverrideDocument(document, quarter) {
  const rawItems = document?.funds || document?.items || document;
  const items = rawItems && typeof rawItems === 'object' && !Array.isArray(rawItems)
    ? Object.fromEntries(Object.entries(rawItems).filter(([, value]) => value && typeof value === 'object' && !Array.isArray(value)))
    : {};
  return {version:Number(document?.version || 1), quarter:String(document?.quarter || quarter), updatedAt:String(document?.updatedAt || ''), items};
}

async function readMasterFundOverrideDocument(env, quarter) {
  const path = masterFundOverrideObjectKey(quarter);
  const object = await env.FUND_DATA.get(path);
  if (!object) return {path, exists:false, document:{version:1, quarter, updatedAt:'', items:{}}};
  try {
    return {path, exists:true, document:normalizeMasterFundOverrideDocument(await object.json(), quarter)};
  } catch {
    throw new Error(`Invalid Master Fund Override JSON: ${path}`);
  }
}

async function writeMasterFundOverrideDocument(env, quarter, items) {
  const path = masterFundOverrideObjectKey(quarter);
  const updatedAt = new Date().toISOString();
  await env.FUND_DATA.put(path, JSON.stringify({version:1, quarter, updatedAt, funds:items}, null, 2), {
    httpMetadata:{contentType:'application/json; charset=utf-8', cacheControl:'no-store'},
    customMetadata:{quarter, dataType:'master-fund-overrides'},
  });
  return {path, updatedAt};
}

async function getMasterFundOverrides(env, quarterInput) {
  const quarter = normalizeQuarter(quarterInput);
  const {path, exists, document} = await readMasterFundOverrideDocument(env, quarter);
  return {ok:true, quarter, items:document.items, source:'Cloudflare R2', r2Exists:exists, r2Path:path, updatedAt:document.updatedAt};
}

async function saveMasterFundOverride(request, env, session) {
  const input = await request.json();
  const quarter = normalizeQuarter(input?.quarter);
  const profile = input?.profile && typeof input.profile === 'object' ? input.profile : null;
  if (!profile) throw new Error('profile is required');
  const key = masterFundOverrideKey(profile);
  const {document} = await readMasterFundOverrideDocument(env, quarter);
  const savedProfile = {...profile, masterFundId:String(profile.masterFundId || '').trim(), isin:String(profile.isin || '').trim().toUpperCase(), updatedAt:new Date().toISOString(), updatedBy:session?.user?.email || ''};
  const items = {...document.items};
  Object.keys(items).forEach(existingKey => {
    const existing = items[existingKey] || {};
    const sameFundId = savedProfile.masterFundId && String(existing.masterFundId || '').trim().toUpperCase() === savedProfile.masterFundId.toUpperCase();
    const sameIsin = savedProfile.isin && String(existing.isin || '').trim().toUpperCase() === savedProfile.isin;
    if (existingKey !== key && (sameFundId || sameIsin)) delete items[existingKey];
  });
  items[key] = savedProfile;
  const stored = await writeMasterFundOverrideDocument(env, quarter, items);
  return {ok:true, quarter, key, profile:savedProfile, items, source:'Cloudflare R2', r2Uploaded:true, r2Exists:true, r2Path:stored.path, updatedAt:stored.updatedAt};
}

async function deleteMasterFundOverride(env, keyInput, quarterInput) {
  const quarter = normalizeQuarter(quarterInput);
  const requestedKey = String(keyInput || '').trim();
  if (!requestedKey) throw new Error('key is required');
  const {document} = await readMasterFundOverrideDocument(env, quarter);
  const actualKey = Object.keys(document.items).find(key => key.toUpperCase() === requestedKey.toUpperCase());
  if (!actualKey) return {ok:true, notFound:true, quarter, key:requestedKey, items:document.items, source:'Cloudflare R2'};
  const items = {...document.items};
  delete items[actualKey];
  const stored = await writeMasterFundOverrideDocument(env, quarter, items);
  return {ok:true, deleted:true, quarter, key:actualKey, items, source:'Cloudflare R2', r2Uploaded:true, r2Exists:true, r2Path:stored.path, updatedAt:stored.updatedAt};
}

function masterAllocationsObjectKey(quarter) {
  return `Data/${quarter.slice(0, 4)}/${quarter}/overrides/fund_master_allocations.json`;
}

function normalizeAllocationKey(value) {
  return String(value || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
}

async function readMasterAllocationsDocument(env, quarter) {
  const path = masterAllocationsObjectKey(quarter);
  const object = await env.FUND_DATA.get(path);
  if (!object) return {path, exists:false, document:{version:1, quarter, updatedAt:'', items:{}}};
  try {
    const payload = await object.json();
    const rawItems = payload?.items || payload;
    const items = rawItems && typeof rawItems === 'object' && !Array.isArray(rawItems)
      ? Object.fromEntries(Object.entries(rawItems).filter(([, value]) => value && typeof value === 'object' && !Array.isArray(value)))
      : {};
    return {path, exists:true, document:{version:Number(payload?.version || 1), quarter:String(payload?.quarter || quarter), updatedAt:String(payload?.updatedAt || ''), items}};
  } catch {
    throw new Error(`Invalid Master Allocations JSON: ${path}`);
  }
}

async function writeMasterAllocationsDocument(env, quarter, items) {
  const path = masterAllocationsObjectKey(quarter);
  const updatedAt = new Date().toISOString();
  await env.FUND_DATA.put(path, JSON.stringify({version:1, quarter, updatedAt, items}, null, 2), {
    httpMetadata:{contentType:'application/json; charset=utf-8', cacheControl:'no-store'},
    customMetadata:{quarter, dataType:'fund-master-allocations'},
  });
  return {path, updatedAt};
}

async function getMasterAllocations(env, quarterInput) {
  const quarter = normalizeQuarter(quarterInput);
  const {path, exists, document} = await readMasterAllocationsDocument(env, quarter);
  return {ok:true, quarter, items:document.items, source:'Cloudflare R2', r2Exists:exists, r2Path:path, updatedAt:document.updatedAt};
}

async function saveMasterAllocation(request, env, session) {
  const input = await request.json();
  const item = input?.item && typeof input.item === 'object' ? input.item : input;
  const quarter = normalizeQuarter(input?.quarter || item?.quarter);
  const key = normalizeAllocationKey(item?.key || item?.thaiFundCode);
  if (!key) throw new Error('Thai Fund Code is required');
  const allocations = Array.isArray(item?.allocations) ? item.allocations : [];
  if (!allocations.length) throw new Error('At least one Master Fund allocation is required');
  if (allocations.some(allocation => !String(allocation?.masterId || '').trim())) throw new Error('Every allocation must have Master ID or ISIN');
  const totalWeight = allocations.reduce((sum, allocation) => sum + (Number(allocation?.weight) || 0), 0);
  if (Math.abs(totalWeight - 100) > 0.01) throw new Error(`Allocation weight must total 100% (current ${totalWeight.toFixed(2)}%)`);
  const {document} = await readMasterAllocationsDocument(env, quarter);
  const now = new Date().toISOString();
  const savedItem = {...item, key, quarter, thaiFundCode:String(item.thaiFundCode || key).trim(), allocations, createdAt:document.items[key]?.createdAt || item.createdAt || now, updatedAt:now, updatedBy:session?.user?.email || ''};
  const items = {...document.items, [key]:savedItem};
  const stored = await writeMasterAllocationsDocument(env, quarter, items);
  return {ok:true, quarter, key, item:savedItem, items, source:'Cloudflare R2', r2Uploaded:true, r2Exists:true, r2Path:stored.path, updatedAt:stored.updatedAt};
}

async function deleteMasterAllocation(env, keyInput, quarterInput) {
  const quarter = normalizeQuarter(quarterInput);
  const requestedKey = normalizeAllocationKey(keyInput);
  if (!requestedKey) throw new Error('key is required');
  const {document} = await readMasterAllocationsDocument(env, quarter);
  const actualKey = Object.keys(document.items).find(key => normalizeAllocationKey(key) === requestedKey);
  if (!actualKey) return {ok:true, notFound:true, quarter, key:requestedKey, items:document.items, source:'Cloudflare R2'};
  const items = {...document.items};
  delete items[actualKey];
  const stored = await writeMasterAllocationsDocument(env, quarter, items);
  return {ok:true, deleted:true, quarter, key:actualKey, items, source:'Cloudflare R2', r2Uploaded:true, r2Exists:true, r2Path:stored.path, updatedAt:stored.updatedAt};
}

function reportDataOverrideObjectKey(quarter) {
  return `Data/${quarter.slice(0, 4)}/${quarter}/overrides/report_data_overrides.json`;
}

function normalizeReportEntityType(value) {
  const entityType = String(value || '').trim();
  if (!['thaiFunds', 'masterFunds'].includes(entityType)) throw new Error('entityType must be thaiFunds or masterFunds');
  return entityType;
}

function normalizeReportDataOverrideDocument(payload, quarter) {
  const cleanEntities = value => value && typeof value === 'object' && !Array.isArray(value) ? value : {};
  return {
    schemaVersion: Number(payload?.schemaVersion || 1),
    quarter: String(payload?.quarter || quarter),
    updatedAt: String(payload?.updatedAt || ''),
    thaiFunds: cleanEntities(payload?.thaiFunds),
    masterFunds: cleanEntities(payload?.masterFunds),
  };
}

async function readReportDataOverrideDocument(env, quarter) {
  const path = reportDataOverrideObjectKey(quarter);
  const object = await env.FUND_DATA.get(path);
  if (!object) return {path, exists:false, document:normalizeReportDataOverrideDocument({}, quarter)};
  try {
    return {path, exists:true, document:normalizeReportDataOverrideDocument(await object.json(), quarter)};
  } catch {
    throw new Error(`Invalid Report Data Override JSON: ${path}`);
  }
}

async function writeReportDataOverrideDocument(env, quarter, document) {
  const path = reportDataOverrideObjectKey(quarter);
  const updatedAt = new Date().toISOString();
  const payload = {...normalizeReportDataOverrideDocument(document, quarter), quarter, updatedAt};
  await env.FUND_DATA.put(path, JSON.stringify(payload, null, 2), {
    httpMetadata:{contentType:'application/json; charset=utf-8', cacheControl:'no-store'},
    customMetadata:{quarter, dataType:'report-data-overrides'},
  });
  return {path, updatedAt, document:payload};
}

async function getReportDataOverrides(env, quarterInput) {
  const quarter = normalizeQuarter(quarterInput);
  const {path, exists, document} = await readReportDataOverrideDocument(env, quarter);
  return {ok:true, ...document, source:'Cloudflare R2', r2Exists:exists, r2Path:path};
}

async function saveReportDataOverrides(request, env, session) {
  const input = await request.json();
  const quarter = normalizeQuarter(input?.quarter);
  const rawChanges = Array.isArray(input?.changes) ? input.changes : [input];
  if (!rawChanges.length) throw new Error('At least one change is required');
  const {document} = await readReportDataOverrideDocument(env, quarter);
  const next = {...document, thaiFunds:{...document.thaiFunds}, masterFunds:{...document.masterFunds}};
  const now = new Date().toISOString();
  const updatedBy = session?.user?.email || '';
  rawChanges.forEach(change => {
    const entityType = normalizeReportEntityType(change?.entityType);
    const key = String(change?.key || '').trim().toUpperCase();
    const section = String(change?.section || '').trim();
    const values = change?.values && typeof change.values === 'object' && !Array.isArray(change.values) ? change.values : null;
    if (!key) throw new Error('key is required');
    if (!section) throw new Error('section is required');
    if (!values) throw new Error('values is required');
    const existing = next[entityType][key] || {key, sections:{}};
    const sections = {...(existing.sections || {})};
    const previous = sections[section] || {};
    const cleanValues = Object.fromEntries(Object.entries(values).filter(([field]) => field && !field.startsWith('_')));
    if (!Object.keys(cleanValues).length) {
      delete sections[section];
      if (Object.keys(sections).length) next[entityType][key] = {...existing, key, sections, updatedAt:now, updatedBy};
      else delete next[entityType][key];
    } else {
      sections[section] = {
        ...previous,
        values: cleanValues,
        updatedAt: now,
        updatedBy,
        note: String(change?.note ?? previous.note ?? '').trim(),
      };
      next[entityType][key] = {...existing, key, sections, updatedAt:now, updatedBy};
    }
  });
  const stored = await writeReportDataOverrideDocument(env, quarter, next);
  return {ok:true, ...stored.document, source:'Cloudflare R2', r2Uploaded:true, r2Exists:true, r2Path:stored.path};
}

async function deleteReportDataOverride(env, entityTypeInput, keyInput, quarterInput, sectionInput) {
  const quarter = normalizeQuarter(quarterInput);
  const entityType = normalizeReportEntityType(entityTypeInput);
  const key = String(keyInput || '').trim().toUpperCase();
  const section = String(sectionInput || '').trim();
  if (!key) throw new Error('key is required');
  const {document} = await readReportDataOverrideDocument(env, quarter);
  const next = {...document, thaiFunds:{...document.thaiFunds}, masterFunds:{...document.masterFunds}};
  const actualKey = Object.keys(next[entityType]).find(itemKey => itemKey.toUpperCase() === key);
  if (!actualKey) return {ok:true, notFound:true, quarter, entityType, key, source:'Cloudflare R2'};
  if (section) {
    const existing = next[entityType][actualKey] || {};
    const sections = {...(existing.sections || {})};
    delete sections[section];
    if (Object.keys(sections).length) next[entityType][actualKey] = {...existing, sections};
    else delete next[entityType][actualKey];
  } else {
    delete next[entityType][actualKey];
  }
  const stored = await writeReportDataOverrideDocument(env, quarter, next);
  return {ok:true, deleted:true, quarter, entityType, key:actualKey, section, ...stored.document, source:'Cloudflare R2', r2Uploaded:true, r2Exists:true, r2Path:stored.path};
}

function titleSlug(value, fallback) {
  const slug = String(value || '').trim().replace(/[^A-Za-z0-9ก-๙._-]+/g, '-').replace(/^-+|-+$/g, '');
  return slug || fallback;
}

function draftFileName(id, draft) {
  const quarter = normalizeQuarter(draft.currentQuarter || draft.dataQuarter || draft.quarter);
  const category = titleSlug(draft.avpCategory || draft.filters?.selectFundFilters?.category || draft.asset, 'All-AVP');
  const type = titleSlug(draft.fundType || draft.filters?.selectFundFilters?.type, 'All-Type');
  const dateText = titleSlug(draft.userDate || String(draft.createdAt || '').slice(0, 10), 'No-Date');
  const fundCount = Array.isArray(draft.selectedKeys) ? draft.selectedKeys.length : Object.keys(draft.selectedFunds || {}).length;
  const highlightCount = Object.keys(draft.highlights || {}).length;
  const author = titleSlug(draft.authorEmail || draft.author || draft.authorName, 'unknown');
  return `${quarter}__${category}__${type}__${dateText}__${fundCount}F-${highlightCount}H__${author}__${safeSlug(id)}.json`;
}

async function listAllObjects(env, prefix) {
  const objects = [];
  let cursor;
  do {
    const page = await env.FUND_DATA.list({prefix, cursor, limit:1000});
    objects.push(...page.objects);
    cursor = page.truncated ? page.cursor : undefined;
  } while (cursor);
  return objects;
}

async function listDrafts(env, quarterInput) {
  const quarter = quarterInput ? normalizeQuarter(quarterInput) : '';
  let index = await readDraftIndex(env);
  if (!index) index = await rebuildDraftIndex(env);
  const drafts = (Array.isArray(index.drafts) ? index.drafts : [])
    .filter(draft => !quarter || draft.currentQuarter === quarter)
    .map(draft => ({...draft, storageSource:'r2', indexOnly:true}));
  return {ok:true, drafts, source:'Cloudflare R2 index', quarter, indexUpdatedAt:index.updatedAt || ''};
}

function draftSummary(draft, key) {
  return {
    id:String(draft.id || ''),
    draftKind:draft.draftKind || 'fund-selection',
    name:draft.name || '',
    avpCategory:draft.avpCategory || draft.filters?.selectFundFilters?.category || draft.asset || '',
    fundType:draft.fundType || draft.filters?.selectFundFilters?.type || '',
    userDate:draft.userDate || '',
    author:draft.author || draft.authorEmail || '',
    authorEmail:draft.authorEmail || draft.author || '',
    authorName:draft.authorName || '',
    notes:draft.notes || '',
    createdAt:draft.createdAt || '',
    updatedAt:draft.updatedAt || draft.createdAt || '',
    currentQuarter:draft.currentQuarter || draft.dataQuarter || draft.quarter || '',
    selectedFundCount:Array.isArray(draft.selectedKeys) ? draft.selectedKeys.length : Object.keys(draft.selectedFunds || {}).length,
    highlightCount:Object.keys(draft.highlights || {}).length,
    r2ObjectKey:key,
    r2FileName:key.split('/').pop(),
  };
}

async function readDraftIndex(env) {
  const object = await env.FUND_DATA.get('Draft/index.json');
  if (!object) return null;
  try {
    const index = await object.json();
    return index && Array.isArray(index.drafts) ? index : null;
  } catch {
    return null;
  }
}

async function writeDraftIndex(env, drafts) {
  const sorted = [...drafts].sort((a, b) => String(b.updatedAt || b.createdAt || '').localeCompare(String(a.updatedAt || a.createdAt || '')));
  const index = {version:1, updatedAt:new Date().toISOString(), count:sorted.length, drafts:sorted};
  await env.FUND_DATA.put('Draft/index.json', JSON.stringify(index, null, 2), {
    httpMetadata:{contentType:'application/json; charset=utf-8', cacheControl:'no-store'},
  });
  return index;
}

async function rebuildDraftIndex(env) {
  const prefix = 'Draft/';
  const objects = await listAllObjects(env, prefix);
  const drafts = [];
  const jsonObjects = objects.filter(object => object.key !== 'Draft/index.json' && object.key.endsWith('.json'));
  for (let i = 0; i < jsonObjects.length; i += 20) {
    const pageObjects = jsonObjects.slice(i, i + 20);
    const batch = await Promise.all(pageObjects.map(object => env.FUND_DATA.get(object.key)));
    for (let j = 0; j < batch.length; j += 1) {
      const object = batch[j];
      if (!object) continue;
      try {
        const draft = await object.json();
        if (!draft || typeof draft !== 'object' || draft.deletedAt) continue;
        drafts.push(draftSummary(draft, pageObjects[j].key));
      } catch {
        // One malformed draft must not break the complete list.
      }
    }
  }
  const index = await writeDraftIndex(env, drafts);
  return {...index, ok:true, rebuilt:true};
}

async function findDraftObject(env, id, quarter) {
  const prefix = `Draft/${quarter.slice(0, 4)}/${quarter}/`;
  const suffix = `__${safeSlug(id)}.json`;
  const legacyName = `${safeSlug(id)}.json`;
  const objects = await listAllObjects(env, prefix);
  return objects.find(object => object.key.endsWith(suffix) || object.key.endsWith(`/${legacyName}`)) || null;
}

async function getDraft(env, idInput, quarterInput) {
  const id = safeSlug(idInput, '');
  if (!id) throw new Error('id is required');
  const quarter = normalizeQuarter(quarterInput);
  const existing = await findDraftObject(env, id, quarter);
  if (!existing) throw new Error(`Draft not found: ${id}`);
  const object = await env.FUND_DATA.get(existing.key);
  if (!object) throw new Error(`Draft not found: ${id}`);
  const draft = await object.json();
  return {ok:true, draft:{...draft, r2ObjectKey:existing.key, r2FileName:existing.key.split('/').pop(), storageSource:'r2'}};
}

async function saveDraft(request, env, session) {
  const length = Number(request.headers.get('Content-Length') || 0);
  if (length > 5 * 1024 * 1024) throw new Error('Draft payload is too large');
  const input = await request.json();
  const draft = input?.draft && typeof input.draft === 'object' ? input.draft : input;
  if (!draft || typeof draft !== 'object') throw new Error('draft is required');
  const now = new Date().toISOString();
  const id = safeSlug(draft.id || Date.now());
  const quarter = normalizeQuarter(draft.currentQuarter || draft.dataQuarter || draft.quarter);
  const existing = await findDraftObject(env, id, quarter);
  const payload = {
    ...draft,
    id,
    currentQuarter:quarter,
    createdAt:draft.createdAt || now,
    updatedAt:now,
    updatedBy:session?.user?.email || '',
  };
  const fileName = draftFileName(id, payload);
  const key = `Draft/${quarter.slice(0, 4)}/${quarter}/${fileName}`;
  await env.FUND_DATA.put(key, JSON.stringify(payload, null, 2), {
    httpMetadata:{contentType:'application/json; charset=utf-8'},
    customMetadata:{draftId:id, quarter},
  });
  if (existing && existing.key !== key) await env.FUND_DATA.delete(existing.key);
  const index = (await readDraftIndex(env)) || {drafts:[]};
  const summaries = (Array.isArray(index.drafts) ? index.drafts : []).filter(item => String(item.id) !== id);
  summaries.push(draftSummary(payload, key));
  await writeDraftIndex(env, summaries);
  return {ok:true, draft:{...payload, r2ObjectKey:key, r2FileName:fileName, storageSource:'r2'}, r2Uploaded:true, fileName, path:key, quarter};
}

async function deleteDraft(env, idInput, quarterInput) {
  const id = safeSlug(idInput, '');
  if (!id) throw new Error('id is required');
  const quarter = normalizeQuarter(quarterInput);
  const existing = await findDraftObject(env, id, quarter);
  if (!existing) return {ok:true, notFound:true, id, quarter};
  await env.FUND_DATA.delete(existing.key);
  const index = await readDraftIndex(env);
  if (index) await writeDraftIndex(env, index.drafts.filter(item => String(item.id) !== id));
  return {ok:true, deleted:true, id, quarter, path:existing.key};
}

function objectKey(pathname) {
  if (pathname === '/manifest.json') return 'manifest.json';
  if (pathname.startsWith('/data/')) return decodeURIComponent(pathname.slice('/data/'.length)).replace(/^\/+/, '');
  return '';
}

function ftObjectKey(pathname) {
  if (pathname === '/ft' || pathname === '/ft/' || pathname === '/ft/index') {
    return 'Data For FT.com/index.json';
  }
  const parts = pathname.split('/').filter(Boolean).map(decodeURIComponent);
  if (parts[0] !== 'ft' || parts[1] !== 'symbols' || !parts[2]) return '';
  const slug = safeFtPathPart(parts[2], 'symbol');
  if (parts.length === 3 || (parts.length === 4 && parts[3] === 'metadata')) {
    return `Data For FT.com/symbols/${slug}/metadata.json`;
  }
  if (parts.length === 5 && parts[3] === 'prices' && /^\d{4}$/.test(parts[4])) {
    return `Data For FT.com/symbols/${slug}/prices/${parts[4]}.json.gz`;
  }
  if (parts.length === 5 && parts[3] === 'qualitative' && parts[4] === 'latest') {
    return `Data For FT.com/symbols/${slug}/qualitative/latest.json`;
  }
  if (parts.length === 6 && parts[3] === 'qualitative' && parts[4] === 'snapshots' && /^\d{4}-\d{2}-\d{2}$/.test(parts[5])) {
    return `Data For FT.com/symbols/${slug}/qualitative/snapshots/${parts[5]}.json`;
  }
  return '';
}

function safeFtPathPart(value, label) {
  const clean = String(value || '').trim().toUpperCase();
  if (!/^[A-Z0-9_]+$/.test(clean)) throw new Error(`Invalid FT ${label}`);
  return clean;
}

async function r2Object(request, env, key, cacheControl) {
  const object = await env.FUND_DATA.get(key);
  if (!object) return cors(request, env, {ok:false,error:`Object not found: ${key}`}, 404);
  const headers = new Headers();
  object.writeHttpMetadata(headers);
  if (!headers.has('Content-Type')) headers.set('Content-Type', 'application/json; charset=utf-8');
  headers.set('Cache-Control', cacheControl);
  headers.set('ETag', object.httpEtag);
  addCors(headers, request, env);
  return new Response(object.body, {headers});
}

function ftJobId(kind) {
  const stamp = new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14);
  const random = crypto.randomUUID().slice(0, 8);
  return `ft-${kind}-${stamp}-${random}`;
}

function ftJobObjectKey(jobId) {
  return `Data For FT.com/jobs/${jobId}.json`;
}

function normalizeFtJobKind(value) {
  const kind = String(value || '').trim().toLowerCase();
  if (!['historical', 'qualitative-all', 'ytd-all'].includes(kind)) throw new Error('Unknown FT job type');
  return kind;
}

function isoDate(value, label, required=false) {
  const clean = String(value || '').trim();
  if (!clean && !required) return '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(clean)) throw new Error(`${label} must be YYYY-MM-DD`);
  return clean;
}

async function createFtJob(request, env, session, kindInput) {
  const kind = normalizeFtJobKind(kindInput);
  if (!env.GITHUB_TOKEN) throw new Error('Worker secret GITHUB_TOKEN is not configured');
  const input = await request.json().catch(() => ({}));
  const now = new Date();
  const today = now.toISOString().slice(0, 10);
  const symbol = String(input.symbol || '').trim();
  const startDate = kind === 'ytd-all'
    ? `${today.slice(0, 4)}-01-01`
    : isoDate(input.startDate, 'startDate', kind === 'historical');
  const endDate = kind === 'ytd-all'
    ? today
    : isoDate(input.endDate, 'endDate', kind === 'historical');
  if (kind === 'historical' && !symbol) throw new Error('FT symbol is required');
  if (startDate && endDate && startDate > endDate) throw new Error('startDate must be before endDate');

  const jobId = ftJobId(kind);
  const runPrices = kind === 'historical' || kind === 'ytd-all';
  const runQualitative = kind === 'qualitative-all';
  const allSymbolsFromDb = kind !== 'historical';
  const repo = String(env.GITHUB_REPO || 'FP2-AVP/Fund-Selection-Tool-V2').trim();
  const workflow = String(env.GITHUB_WORKFLOW || 'ft-historical-prices-database.yml').trim();
  const ref = String(env.GITHUB_REF || 'main').trim();
  const quarter = String(input.quarter || `${today.slice(0, 4)}-Q1`).trim();
  const workflowInputs = {
    quarter,
    symbol: symbol || 'IXN:PCQ:USD',
    symbols: '',
    all_symbols_from_db: allSymbolsFromDb ? 'true' : 'false',
    ft_url: String(input.url || ''),
    start_date: startDate,
    end_date: endDate,
    run_prices: runPrices ? 'true' : 'false',
    run_qualitative: runQualitative ? 'true' : 'false',
    drive_folder_id: String(env.FT_DRIVE_FOLDER_ID || '1Locig3aW7hVs0SxCoeFpa76OJ0XcF1rg'),
    file_name: String(env.FT_DATABASE_FILE_NAME || 'ft_historical_prices_database.json'),
    app_script_url: '',
    app_script_key: '',
    continue_on_error: 'true',
    sleep_seconds: String(input.sleepSeconds || '1'),
    job_id: jobId,
    job_kind: kind,
  };
  const response = await fetch(`https://api.github.com/repos/${repo}/actions/workflows/${encodeURIComponent(workflow)}/dispatches`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${env.GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'fund-selection-data-worker',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ref, inputs:workflowInputs}),
  });
  if (!response.ok) throw new Error(`GitHub workflow dispatch failed (${response.status}): ${await response.text()}`);
  const job = {
    ok:true, jobId, kind, status:'queued', conclusion:'', createdAt:now.toISOString(),
    createdBy:session?.user?.email || '', repo, workflow, ref, symbol, startDate, endDate,
    allSymbolsFromDb, runPrices, runQualitative,
  };
  await env.FUND_DATA.put(ftJobObjectKey(jobId), JSON.stringify(job, null, 2), {
    httpMetadata:{contentType:'application/json; charset=utf-8', cacheControl:'no-store'},
    customMetadata:{jobId, kind},
  });
  return job;
}

async function getFtJob(env, jobIdInput) {
  const jobId = String(jobIdInput || '').trim();
  if (!/^ft-[a-z0-9-]+$/.test(jobId)) throw new Error('Invalid FT job id');
  const object = await env.FUND_DATA.get(ftJobObjectKey(jobId));
  if (!object) throw new Error(`FT job not found: ${jobId}`);
  const job = await object.json();
  if (!env.GITHUB_TOKEN) return job;
  const response = await fetch(`https://api.github.com/repos/${job.repo}/actions/workflows/${encodeURIComponent(job.workflow)}/runs?event=workflow_dispatch&per_page=30`, {
    headers: {
      Authorization: `Bearer ${env.GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'fund-selection-data-worker',
    },
  });
  if (!response.ok) return {...job, statusError:`GitHub runs HTTP ${response.status}`};
  const data = await response.json();
  const run = (data.workflow_runs || []).find(item => String(item.display_title || item.name || '').includes(jobId));
  if (!run) return job;
  return {
    ...job,
    status:run.status || job.status,
    conclusion:run.conclusion || '',
    runId:run.id,
    runNumber:run.run_number,
    runUrl:run.html_url || '',
    runStartedAt:run.run_started_at || '',
    updatedAt:run.updated_at || '',
  };
}

async function readSession(request, env) {
  const match = (request.headers.get('Authorization') || '').match(/^Bearer\s+([A-Za-z0-9_-]+)$/);
  if (!match) return null;
  return env.AUTH_SESSIONS.get(`session:${match[1]}`, 'json');
}

function addCors(headers, request, env) {
  const origin = request.headers.get('Origin') || '';
  const primary = new URL(env.FRONTEND_URL).origin;
  const configured = String(env.ADDITIONAL_FRONTEND_ORIGINS || 'http://localhost:8080,http://127.0.0.1:8080')
    .split(',')
    .map(value => value.trim())
    .filter(Boolean)
    .map(value => new URL(value).origin);
  const allowedOrigins = [...new Set([primary, ...configured])];
  headers.set('Access-Control-Allow-Origin', allowedOrigins.includes(origin) ? origin : primary);
  headers.set('Access-Control-Allow-Headers', 'Authorization, Content-Type');
  headers.set('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  headers.set('Vary', 'Origin');
}

function cors(request, env, body, status=200) {
  const headers = new Headers({'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});
  addCors(headers, request, env);
  return body === null ? new Response(null,{status,headers}) : new Response(JSON.stringify(body),{status,headers});
}
