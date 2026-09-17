/**
 * Handler serverless Vercel (Node) + Hono.
 * Export em formato Node para evitar hang do adapter edge.
 */
import app from '../server/app.js';

export const config = {
  api: {
    bodyParser: false,
  },
  maxDuration: 30,
};

function buildRequest(req) {
  const host = req.headers['x-forwarded-host'] || req.headers.host || 'localhost';
  const proto = req.headers['x-forwarded-proto'] || 'https';
  const url = `${proto}://${host}${req.url}`;

  const headers = new Headers();
  Object.entries(req.headers || {}).forEach(([key, value]) => {
    if (value === undefined || value === null) return;
    if (Array.isArray(value)) value.forEach((v) => headers.append(key, v));
    else headers.set(key, String(value));
  });

  const method = req.method || 'GET';
  const init = { method, headers };

  if (method !== 'GET' && method !== 'HEAD') {
    init.body = req;
    init.duplex = 'half';
  }

  return new Request(url, init);
}

export default async function handler(req, res) {
  try {
    const request = buildRequest(req);
    const response = await app.fetch(request);
    res.statusCode = response.status;
    response.headers.forEach((value, key) => {
      if (key.toLowerCase() === 'set-cookie') {
        const current = res.getHeader('Set-Cookie');
        if (!current) res.setHeader('Set-Cookie', value);
        else if (Array.isArray(current)) res.setHeader('Set-Cookie', [...current, value]);
        else res.setHeader('Set-Cookie', [current, value]);
      } else {
        res.setHeader(key, value);
      }
    });
    const buf = Buffer.from(await response.arrayBuffer());
    res.end(buf);
  } catch (err) {
    console.error('API error:', err);
    res.statusCode = 500;
    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify({ error: err.message || 'Internal error' }));
  }
}
