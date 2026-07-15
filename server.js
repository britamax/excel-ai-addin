const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = 3000;
const DIR = __dirname;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'application/javascript; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png':  'image/png',
  '.jpg':  'image/jpeg',
  '.svg':  'image/svg+xml',
  '.xml':  'application/xml',
  '.ico':  'image/x-icon',
};

const opts = {
  key:  fs.readFileSync(path.join(DIR, 'cert', 'key.pem')),
  cert: fs.readFileSync(path.join(DIR, 'cert', 'cert.pem')),
};

https.createServer(opts, (req, res) => {
  const uri = url.parse(req.url).pathname;
  let fp = path.join(DIR, uri === '/' ? 'src/taskpane/taskpane.html' : uri);
  // Security: prevent path traversal
  if (!fp.startsWith(DIR)) { res.writeHead(403); res.end(); return; }

  const ext = path.extname(fp);
  const ct = MIME[ext] || 'application/octet-stream';

  // CORS headers for Excel Web
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  fs.readFile(fp, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not Found: ' + uri);
      return;
    }
    res.writeHead(200, { 'Content-Type': ct });
    res.end(data);
  });
}).listen(PORT, () => {
  console.log(`HTTPS server running on https://localhost:${PORT}`);
});
