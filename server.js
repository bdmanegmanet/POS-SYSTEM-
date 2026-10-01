const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = Number(process.env.PORT || 10000);
const GAS_WEB_APP_URL = String(process.env.GAS_WEB_APP_URL || '').trim();
const indexPath = path.join(__dirname, 'Index.html');

function send(res, status, body, type = 'text/plain; charset=utf-8') {
  res.writeHead(status, {
    'Content-Type': type,
    'Cache-Control': 'no-store'
  });
  if (res.req && res.req.method === 'HEAD') return res.end();
  res.end(body);
}

function getIndex() {
  if (!fs.existsSync(indexPath)) {
    throw new Error('Index.html is missing from the deployment.');
  }
  return fs.readFileSync(indexPath, 'utf8');
}

async function proxyToGas(body) {
  if (!GAS_WEB_APP_URL) {
    throw new Error('GAS_WEB_APP_URL is not configured on Render.');
  }

  const response = await fetch(GAS_WEB_APP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
    redirect: 'follow'
  });

  const text = await response.text();
  return { status: response.status, text };
}

const server = http.createServer(async (req, res) => {
  try {
    const requestUrl = new URL(req.url || '/', 'http://localhost');
    const pathname = requestUrl.pathname;

    // Render health check / uptime endpoint.
    if (pathname === '/healthz' && (req.method === 'GET' || req.method === 'HEAD')) {
      return send(
        res,
        200,
        JSON.stringify({
          ok: true,
          service: 'POS System',
          gasConfigured: Boolean(GAS_WEB_APP_URL)
        }),
        'application/json; charset=utf-8'
      );
    }

    // Apps Script API proxy.
    if (pathname === '/api' && req.method === 'POST') {
      let raw = '';
      req.setEncoding('utf8');

      req.on('data', chunk => {
        raw += chunk;
        if (raw.length > 2_000_000) req.destroy(new Error('Request body too large.'));
      });

      req.on('end', async () => {
        try {
          let body = {};
          try {
            body = JSON.parse(raw || '{}');
          } catch (_) {
            return send(
              res,
              400,
              JSON.stringify({ success: false, error: 'Invalid JSON request.' }),
              'application/json; charset=utf-8'
            );
          }

          const out = await proxyToGas(body);
          return send(res, out.status, out.text, 'application/json; charset=utf-8');
        } catch (e) {
          return send(
            res,
            503,
            JSON.stringify({ success: false, error: e.message || 'Backend unavailable.' }),
            'application/json; charset=utf-8'
          );
        }
      });
      return;
    }

    // GET /api gives a clear diagnostic instead of a generic 404.
    if (pathname === '/api' && (req.method === 'GET' || req.method === 'HEAD')) {
      return send(
        res,
        200,
        JSON.stringify({
          ok: true,
          endpoint: '/api',
          method: 'POST',
          gasConfigured: Boolean(GAS_WEB_APP_URL)
        }),
        'application/json; charset=utf-8'
      );
    }

    // Main frontend. URL query strings such as /?page=dashboard are supported.
    if ((req.method === 'GET' || req.method === 'HEAD') && (pathname === '/' || pathname === '/index.html')) {
      return send(res, 200, getIndex(), 'text/html; charset=utf-8');
    }

    // SPA fallback: serve the POS frontend for client-side routes.
    if (req.method === 'GET' || req.method === 'HEAD') {
      return send(res, 200, getIndex(), 'text/html; charset=utf-8');
    }

    return send(res, 404, 'Not found');
  } catch (e) {
    console.error('Server error:', e);
    return send(res, 500, e.message || 'Internal server error');
  }
});

server.on('clientError', (err, socket) => {
  console.error('Client error:', err.message);
  if (socket.writable) socket.end('HTTP/1.1 400 Bad Request\r\n\r\n');
});

server.listen(PORT, '0.0.0.0', () => {
  console.log('POS System listening on ' + PORT);
  console.log('Index file: ' + indexPath + ' | exists=' + fs.existsSync(indexPath));
  console.log('GAS backend configured: ' + Boolean(GAS_WEB_APP_URL));
});
