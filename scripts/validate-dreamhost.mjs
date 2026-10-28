import fs from "node:fs";
import path from "node:path";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

export function validateDreamhost(directory) {
  const root = path.resolve(directory);
  const files = [];
  function walk(folder) {
    for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
      assert.ok(!entry.isSymbolicLink(), "Output must not contain links to other directories/files");
      const full = path.join(folder, entry.name);
      if (entry.isDirectory()) walk(full);
      else files.push(path.relative(root, full).replaceAll(path.sep, "/"));
    }
  }
  walk(root);
  assert.ok(files.includes("index.html") && files.includes(".htaccess"), "HTML and Apache configuration are required");
  for (const name of files) {
    assert.match(name, /^(?:index\.html|\.htaccess|assets\/[^/]+-[\w-]{8,}\.(?:js|css|png|jpe?g|svg|gif|webp|ico|woff2?|ttf|avif))$/, `Unexpected upload file: ${name}`);
    if (!/\.(?:html|js|css|svg)$/.test(name)) continue;
    const content = fs.readFileSync(path.join(root, name), "utf8");
    assert.doesNotMatch(content, /https?:\/\/(?:localhost|127\.0\.0\.1|0\.0\.0\.0)|\.devtunnels\.ms|\.ngrok\.|\$\{\{|<!--app-(?:head|html)-->|BEGIN (?:RSA |EC )?PRIVATE KEY|-----BEGIN CERTIFICATE|[A-Za-z]:[\\/](?:Users|source)[\\/]|sourceMappingURL=/i, `Development/secret marker in ${name}`);
  }
  const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
  const assetUrls = [...html.matchAll(/(?:src|href)="([^"]+)"/g)].map(match => match[1]);
  assert.ok(assetUrls.some(url => url.endsWith(".js")) && assetUrls.some(url => url.endsWith(".css")), "Compiled JS and CSS references are required");
  for (const url of assetUrls) {
    assert.match(url, /^\/assets\/[\w.-]+$/, "Production assets must use the root /assets/ path");
    assert.ok(files.includes(url.slice(1)), "Referenced asset must exist");
  }
  const javascript = files.filter(name => name.endsWith(".js")).map(name => fs.readFileSync(path.join(root, name), "utf8")).join("\n");
  assert.ok(javascript.includes("https://signaltuner-back.azurewebsites.net"), "Approved backend URL is missing");
  assert.doesNotMatch(javascript, /VITE_SIGNALTUNER_API_URL/, "API URL must be compiled, not read at runtime");
  assert.equal(fs.readFileSync(path.join(root, ".htaccess"), "utf8"), fs.readFileSync(new URL("../hosting/dreamhost/.htaccess", import.meta.url), "utf8"), "Upload headers/routing must match the source configuration");
  console.log(`DreamHost static validation passed (${files.length} files; root assets, approved API, no source maps/development markers).`);
  return files;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  validateDreamhost(fileURLToPath(new URL("../dist/dreamhost/", import.meta.url)));
}
