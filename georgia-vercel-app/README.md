Georgia’s Support App — Vercel setup

Status: prepared and locally tested, not deployed. No keys or secrets are included.

USE YOUR EXISTING GITHUB REPOSITORY (no ZIP needed)
1. Keep your existing public/auth.js, public/bridge.js, public/calendar-input.js and functions/api/[[path]].js.
2. Replace ALL of public/app.html with the latest app-ready-for-vercel.txt file. Do not append it. The first character must be < in <!doctype html>.
3. Add the four files printed below, using GitHub → Add file → Create new file. Folder paths are part of each filename. Commit each to main.
4. In Vercel, choose Add New → Project → Import your GitHub repository.
5. Framework Preset: Other. Root Directory: repository root. Build command: blank. Output Directory: public. vercel.json contains these settings and API function time limits.
6. Add these four Environment Variables for Production, with the same existing values:
   SUPABASE_URL — project URL only, e.g. https://your-project.supabase.co (not /rest/v1/).
   SUPABASE_ANON_KEY — Supabase publishable key, not secret/service_role.
   APPS_SCRIPT_URL — existing deployed Web app URL ending /exec.
   BRIDGE_SECRET — must exactly equal Apps Script property SUPPORT_WEB_BRIDGE_SECRET.
   Keep all values out of GitHub and chat. Add Preview scope only if you intend preview deployments to use those services.
7. Deploy. If variables are added afterward, redeploy. Open the main production URL.
8. Verify /api/config displays JSON rather than HTML. Then sign in with your existing app email and password, load the calendar and notes, and check finance access. Test saves deliberately; this package’s tests do not write real records.
9. In Apps Script → Project Settings → Script properties, change SUPPORT_WEB_SITE_URL to the final Vercel HTTPS address so generated links point there. Keep the existing /exec deployment and bridge secret; changing the host does not require replacing Apps Script code.

COMPLETE ZIP OPTION
If you prefer a new private repository, unzip the complete Vercel package and upload its contents. Keep api, lib, functions, public and vercel.json at the repository root. It contains the latest welcome page, sign-in page and interface.

The ChatGPT working draft remains separate and uses example records. A new host does not resolve the original phishing classification. Live service connectivity and visual browser checks still need verification after deployment.

Official configuration reference: https://vercel.com/docs/project-configuration/vercel-json
Function duration reference: https://vercel.com/docs/functions/configuring-functions/duration

FOUR FILES TO ADD TO YOUR EXISTING REPOSITORY


FILE: api/config.js

import {createVercelHandler} from '../lib/vercel-api.js';
export default createVercelHandler('/api/config');


FILE: api/rpc.js

import {createVercelHandler} from '../lib/vercel-api.js';
export default createVercelHandler('/api/rpc');


FILE: lib/vercel-api.js

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


FILE: vercel.json

{
  "$schema": "https://openapi.vercel.sh/vercel.json",
  "version": 2,
  "framework": null,
  "buildCommand": "",
  "outputDirectory": "public",
  "functions": {
    "api/*.js": {
      "maxDuration": 120
    }
  },
  "headers": [
    {
      "source": "/(.*)",
      "headers": [
        {
          "key": "Cache-Control",
          "value": "no-store"
        },
        {
          "key": "X-Content-Type-Options",
          "value": "nosniff"
        },
        {
          "key": "Referrer-Policy",
          "value": "no-referrer"
        },
        {
          "key": "X-Frame-Options",
          "value": "DENY"
        },
        {
          "key": "Permissions-Policy",
          "value": "camera=(), microphone=(), geolocation=()"
        }
      ]
    }
  ]
}
