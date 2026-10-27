const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");
const Module = require("node:module");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");

const filename = path.resolve(__dirname, "../src/Tab/MeetingServiceStatus.tsx");
const loaded = new Module(filename, module);
loaded.filename = filename;
loaded.paths = module.paths;
const originalRequire = loaded.require.bind(loaded);
loaded.require = name => name.endsWith(".css") ? {} : name.endsWith(".png") ? { default: "teams-logo.png" } : originalRequire(name);
loaded._compile(ts.transpileModule(fs.readFileSync(filename, "utf8"), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.ReactJSX },
}).outputText, filename);
const { MeetingServiceStatusView: View, incidentPreview, incidentReportLink, loadTeamsServiceStatus } = loaded.exports;
const incident = (id, status) => ({ IncidentId: id, Title: `Synthetic incident ${id}`, Status: status,
  PublishDate: `2026-01-${String(id).padStart(2, "0")}T12:00:00Z`, Description: "Synthetic description." });
const details = (status, incidents = []) => ({ DisplayName: "Microsoft Teams", CurrentStatus: status, incidents });
const render = props => renderToStaticMarkup(React.createElement(View, props));

test("operational empty row shows zero real incidents and a collapsed website-style disclosure", () => {
  const markup = render({ details: details("Operational") });
  for (const text of ["Microsoft Teams", "Operational", "0 active incidents", "View incidents", "No incidents to display.", "See full report history"])
    assert.ok(markup.includes(text), text);
  assert.ok(markup.includes('data-service-status="healthy"'));
  assert.ok(markup.includes('aria-expanded="false"'));
  assert.ok(markup.includes('hidden=""'));
  assert.ok(markup.includes('href="https://www.signaltuner.com/teamsstatus"'));
  assert.ok(!markup.includes("Synthetic incident"));
});

test("preview keeps every unresolved incident and the five newest resolved reports", () => {
  const values = Array.from({ length: 9 }, (_, i) => incident(i + 1, "Resolved"));
  values.push(incident(10, "Outage"), incident(11, "Investigating"), incident(12, "Active"));
  const preview = incidentPreview(values);
  assert.deepEqual(preview.map(item => item.IncidentId), [12, 11, 10, 9, 8, 7, 6, 5]);
  assert.equal(values[0].IncidentId, 1);
  const markup = render({ details: details("Active incident", preview) });
  assert.ok(markup.includes("3 active incidents"));
  assert.ok(markup.includes('teamsServiceCount teamsServiceStatus teamsServiceStatus--outage'));
  assert.ok(markup.includes('data-incident-state="unknown"'));
  assert.ok(markup.includes("Updated"));
  assert.ok(markup.includes("Synthetic description."));
});

test("loading and failure stay unavailable instead of claiming operational status or zero incidents", () => {
  const loading = render({ details: null });
  assert.ok(loading.includes("Loading service status..."));
  const failed = render({ details: null, error: "Service could not be loaded." });
  for (const markup of [loading, failed]) {
    assert.ok(markup.includes("Incident count unavailable"));
    assert.ok(markup.includes('disabled=""'));
    assert.ok(!markup.includes("0 active incidents"));
  }
  assert.ok(failed.includes("Status unavailable"));
  assert.ok(failed.includes('role="alert"'));
});

test("provider links allow only HTTP(S), and incident content is rendered as text", () => {
  for (const link of ["javascript:alert(1)", "data:text/html,hello", "/relative", "bad", undefined]) assert.equal(incidentReportLink(link), null);
  assert.equal(incidentReportLink("https://example.invalid/report"), "https://example.invalid/report");
  const value = { ...incident(1, "Active"), Title: "<img onerror=alert(1)>", Link: "javascript:alert(1)" };
  const markup = render({ details: details("Active incident", [value]) });
  assert.ok(markup.includes("&lt;img onerror=alert(1)&gt;"));
  assert.ok(!markup.includes('href="javascript:'));
});

test("loads the same Microsoft Teams service endpoints as the website with auth and cancellation", async t => {
  const calls = [], controller = new AbortController();
  t.mock.method(global, "fetch", async (url, options) => {
    calls.push([url, options]);
    return { ok: true, json: async () => url.includes("GetServiceStatus") ? details("Operational") : [incident(1, "Resolved")] };
  });
  const result = await loadTeamsServiceStatus("https://api.example.invalid/", "synthetic-token", controller.signal);
  assert.deepEqual(calls.map(([url]) => url), ["https://api.example.invalid/api/User/GetServiceStatus/3", "https://api.example.invalid/api/User/GetServiceIncidents/3"]);
  for (const [, options] of calls) {
    assert.equal(options.headers.Authorization, "Bearer synthetic-token");
    assert.equal(options.signal, controller.signal);
  }
  assert.equal(result.CurrentStatus, "Operational");
  assert.equal(result.incidents.length, 1);
});

test("HTTP and malformed responses fail rather than inventing empty incident data", async t => {
  t.mock.method(global, "fetch", async () => ({ ok: false }));
  await assert.rejects(loadTeamsServiceStatus("https://api.example.invalid", "test", new AbortController().signal));
  global.fetch = async () => ({ ok: true, json: async () => ({}) });
  await assert.rejects(loadTeamsServiceStatus("https://api.example.invalid", "test", new AbortController().signal));
});
