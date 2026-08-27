export default {
  async fetch(request, env) {
    try {
      if (request.method === 'OPTIONS') return cors(request, env, null, 204);
      if (request.method !== 'GET') return cors(request, env, {ok:false,error:'Method not allowed'}, 405);
      const session = await readSession(request, env);
      if (!session) return cors(request, env, {ok:false,error:'Session expired'}, 401);
      const url = new URL(request.url);
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
  headers.set('Access-Control-Allow-Methods', 'GET, OPTIONS');
  headers.set('Vary', 'Origin');
}

function cors(request, env, body, status=200) {
  const headers = new Headers({'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});
  addCors(headers, request, env);
  return body === null ? new Response(null,{status,headers}) : new Response(JSON.stringify(body),{status,headers});
}
