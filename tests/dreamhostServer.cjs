// Test-only HTTP approximation of the Apache routing contract, not a production server.
// Does not execute .htaccess or implement TLS. Live Apache/header verification is required.
const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const root = path.resolve(__dirname, "../dist/dreamhost");
const mime = { ".html": "text/html", ".js": "application/javascript", ".css": "text/css", ".svg": "image/svg+xml", ".png": "image/png", ".jpg": "image/jpeg", ".woff2": "font/woff2" };
exports.startStaticServer = async () => {
  const config = fs.readFileSync(path.join(root, ".htaccess"), "utf8");
  const headers = Object.fromEntries([...config.matchAll(/^Header always set ([\w-]+) "([^"]+)"$/gm)].map(match => [match[1], match[2]]));
  const server = http.createServer((req, res) => {
    for (const [name, value] of Object.entries(headers)) res.setHeader(name, value);
    let pathname;
    try { pathname = decodeURIComponent(new URL(req.url, "http://localhost").pathname); }
    catch { res.writeHead(400); res.end(); return; }
    const relative = pathname.slice(1);
    const blocked = /(^|\/)\.(?!well-known(?:\/|$))/.test(relative)
      || /^(?:src|node_modules|config|env|scripts|tests|infra|appPackage)(?:\/|$)/i.test(relative)
      || /\.(?:pem|key|pfx|p12|crt|map|ts|tsx|cs|bicep)$/i.test(relative);
    let file = path.resolve(root, relative || "index.html");
    const exists = file.startsWith(root + path.sep) && fs.existsSync(file) && fs.statSync(file).isFile();
    if (blocked || (!exists && (/^(?:assets|api|\.well-known)(?:\/|$)/i.test(relative) || /(^|\/)[^/]*\.[^/]+$/.test(relative)))) {
      res.writeHead(404, { "Content-Type": "text/plain" }); res.end("Not found"); return;
    }
    if (!exists) file = path.join(root, "index.html");
    if (/-[\w-]{8,}\.(?:js|css|png|jpe?g|svg|gif|webp|ico|woff2?|ttf|avif)$/.test(file))
      res.setHeader("Cache-Control", "public, max-age=31536000, immutable");
    res.setHeader("Content-Type", mime[path.extname(file)] || "application/octet-stream");
    if (req.method === "HEAD") res.end(); else fs.createReadStream(file).pipe(res);
  });
  await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
  return { origin: `http://127.0.0.1:${server.address().port}`, close: () => new Promise(resolve => { server.close(resolve); server.closeAllConnections(); }) };
};
