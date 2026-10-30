import {handle} from '../functions/api/[[path]].js';

const MAX_BODY_BYTES = 150000;
async function readBody(req) {
  let body;
  if (req.body !== undefined) {
    body = Buffer.isBuffer(req.body) ? req.body.toString('utf8')
      : typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
  } else {
    const chunks = [];
    let size = 0;
    for await (const chunk of req) {
      const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
      size += buffer.length;
      if (size > MAX_BODY_BYTES) throw new RangeError('Body too large');
      chunks.push(buffer);
    }
    body = Buffer.concat(chunks).toString('utf8');
  }
  if (Buffer.byteLength(body || '', 'utf8') > MAX_BODY_BYTES) throw new RangeError('Body too large');
  return body || '';
}

export function createVercelHandler(route, options = {}) {
  return async function vercelHandler(req, res) {
    let response;
    try {
      const headers = new Headers();
      for (const [name, value] of Object.entries(req.headers || {})) {
        if (value !== undefined) headers.set(name, Array.isArray(value) ? value.join(', ') : String(value));
      }
      const host = headers.get('host');
      if (!host || !/^[a-z0-9.-]+(?::\d+)?$/i.test(host)) {
        response = Response.json({error:'Invalid request host.'}, {status:400});
      } else {
        const protocol = headers.get('x-forwarded-proto') === 'http' ? 'http' : 'https';
        const method = String(req.method || 'GET').toUpperCase();
        const init = {method, headers};
        if (method !== 'GET' && method !== 'HEAD') init.body = await readBody(req);
        const request = new Request(protocol + '://' + host + route, init);
        const original = options.env || process.env;
        const env = {...original,
          SUPABASE_URL: String(original.SUPABASE_URL || '').trim().replace(/\/rest\/v1\/?$/, '').replace(/\/$/, ''),
          SUPABASE_ANON_KEY: String(original.SUPABASE_ANON_KEY || '').trim()
        };
        response = await handle(request, env, options.fetcher || fetch);
      }
    } catch (error) {
      response = Response.json({error:error instanceof RangeError ? 'Request is too large.' : 'The service is temporarily unavailable. Please try again.'}, {status:error instanceof RangeError ? 413 : 503});
    }
    res.statusCode = response.status;
    response.headers.forEach((value, name) => res.setHeader(name, value));
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.end(await response.text());
  };
}
