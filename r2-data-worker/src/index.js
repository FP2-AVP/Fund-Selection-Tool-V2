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

async function readSession(request, env) {
  const match = (request.headers.get('Authorization') || '').match(/^Bearer\s+([A-Za-z0-9_-]+)$/);
  if (!match) return null;
  return env.AUTH_SESSIONS.get(`session:${match[1]}`, 'json');
}

function addCors(headers, request, env) {
  const origin = request.headers.get('Origin') || '';
  const allowed = new URL(env.FRONTEND_URL).origin;
  headers.set('Access-Control-Allow-Origin', origin === allowed ? origin : allowed);
  headers.set('Access-Control-Allow-Headers', 'Authorization, Content-Type');
  headers.set('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  headers.set('Vary', 'Origin');
}

function cors(request, env, body, status=200) {
  const headers = new Headers({'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});
  addCors(headers, request, env);
  return body === null ? new Response(null,{status,headers}) : new Response(JSON.stringify(body),{status,headers});
}
