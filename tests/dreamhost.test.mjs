import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { fileURLToPath } from "node:url";
import { resolveConfig } from "vite";
import { validateDreamhost } from "../scripts/validate-dreamhost.mjs";
import { startStaticServer } from "./dreamhostServer.cjs";

test("production artifact has only compiled assets, HTML and Apache config", () => {
  validateDreamhost(fileURLToPath(new URL("../dist/dreamhost/", import.meta.url)));
});

test("DreamHost mode isolates environment and local Teams base remains unchanged", async () => {
  const production = await resolveConfig({ mode: "dreamhost" }, "build");
  assert.equal(production.base, "/");
  assert.deepEqual(production.envPrefix, []);
  assert.equal(production.publicDir, "");
  assert.equal(production.build.sourcemap, false);
  assert.equal(production.define["import.meta.env.VITE_SIGNALTUNER_API_URL"], '"https://signaltuner-back.azurewebsites.net"');
  const local = await resolveConfig({ mode: "production" }, "build");
  assert.equal(local.base, "/tabs/home/");
});

test("a stale shell API override cannot create a localhost production build", async () => {
  const previous = process.env.VITE_SIGNALTUNER_API_URL;
  process.env.VITE_SIGNALTUNER_API_URL = "http://localhost:4317/backend";
  try { await assert.rejects(resolveConfig({ mode: "dreamhost" }, "build"), /approved HTTPS Azure backend/); }
  finally { if (previous === undefined) delete process.env.VITE_SIGNALTUNER_API_URL; else process.env.VITE_SIGNALTUNER_API_URL = previous; }
});

test("Apache source permits only documented framing hosts, protects assets and preserves query strings", () => {
  const config = fs.readFileSync(new URL("../hosting/dreamhost/.htaccess", import.meta.url), "utf8");
  const policy = config.match(/Header always set Content-Security-Policy "([^"]+)"/)[1];
  for (const host of ["https://teams.microsoft.com", "https://*.teams.microsoft.com", "https://*.cloud.microsoft", "https://*.office.com", "https://*.microsoft365.com", "https://outlook.office365.com", "https://outlook-sdf.office365.com"]) assert.ok(policy.includes(host));
  assert.doesNotMatch(policy, /frame-ancestors\s+\*|https:\/\/\*\s|https:\/\/\*\.microsoft\.com/);
  assert.match(config, /^Header unset X-Frame-Options$/m);
  assert.match(config, /^Header always unset X-Frame-Options$/m);
  assert.match(config, /Options -Indexes -MultiViews/);
  assert.match(config, /RewriteRule \^ index\.html \[END\]/);
  assert.doesNotMatch(config, /\[.*QSD.*\]|index\.html\?/);
  assert.ok(config.indexOf("REQUEST_FILENAME") < config.indexOf("RewriteRule ^ index.html"));
  assert.match(config, /REQUEST_STATUS.*200/);
});

test("static HTTP contract preserves app routes, query strings, refreshes and real assets", async t => {
  const server = await startStaticServer(); t.after(server.close);
  const routes = ["/", "/tabs/home", "/tabs/config", "/tabs/login", "/tabs/create-account", "/tabs/support", "/tabs/privacy", "/tabs/terms", "/login", "/create-account", "/forgot-password", "/support", "/privacy", "/terms", "/tabs/forgot-password"];
  for (const route of routes) {
    const target = `${server.origin}${route}?returnUrl=${encodeURIComponent(server.origin + "/tabs/home?meeting=synthetic")}`;
    for (let refresh = 0; refresh < 2; refresh++) {
      const response = await fetch(target);
      assert.equal(response.status, 200, route); assert.equal(response.url, target);
      assert.match(response.headers.get("content-type"), /text\/html/);
      assert.match(response.headers.get("cache-control"), /no-cache/);
      assert.equal(response.headers.get("x-frame-options"), null);
      const html = await response.text(); assert.ok(html.includes('<div id="root">'));
      for (const [, url] of html.matchAll(/(?:src|href)="([^"]+)"/g)) {
        const asset = await fetch(server.origin + url);
        assert.equal(asset.status, 200); assert.doesNotMatch(asset.headers.get("content-type"), /text\/html/);
        assert.match(asset.headers.get("cache-control"), /immutable/);
        await asset.arrayBuffer();
      }
    }
  }
  for (const route of ["/assets/missing.js", "/assets/missing", "/missing.css", "/tabs/missing.png", "/missing.woff2", "/api/anything", "/.env", "/.htaccess", "/src/Tab/App.tsx", "/local.key"]) {
    const response = await fetch(server.origin + route); assert.equal(response.status, 404, route);
    assert.doesNotMatch(await response.text(), /<html|id="root"/); assert.doesNotMatch(response.headers.get("cache-control"), /immutable/);
  }
});

test("production Toolkit has no Azure or package operations and local debugger is retained", () => {
  const remote = fs.readFileSync(new URL("../m365agents.yml", import.meta.url), "utf8");
  assert.doesNotMatch(remote, /uses:\s*(?:arm\/|azureAppService\/|teamsApp\/)/);
  assert.match(remote, /run build:dreamhost/); assert.match(remote, /toolkit:production-disabled/);
  assert.match(remote, /projectId: d394d4cb-2010-4a84-b662-6d5806379557/);
  const local = fs.readFileSync(new URL("../m365agents.local.yml", import.meta.url), "utf8");
  assert.match(local, /localhost:3978/); assert.match(local, /devTool\/install/);
  assert.match(fs.readFileSync(new URL("../src/index.ts", import.meta.url), "utf8"), /https\.createServer/);
});
