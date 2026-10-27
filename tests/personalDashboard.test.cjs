const { test, afterEach } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");
const Module = require("node:module");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
function load(name) {
  const filename = path.resolve(__dirname, "../src/Tab", name);
  const loaded = new Module(filename, module); loaded.filename = filename; loaded.paths = module.paths;
  const original = loaded.require.bind(loaded);
  loaded.require = name => {
    if (/\.(css|svg|png)$/.test(name)) return name;
    if (name.startsWith(".")) {
      const source = [".tsx", ".ts"].map(ext => path.resolve(path.dirname(filename), name + ext)).find(fs.existsSync);
      if (source) return load(source);
    }
    return original(name);
  };
  loaded._compile(ts.transpileModule(fs.readFileSync(filename, "utf8"), { compilerOptions: {
    module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.ReactJSX, esModuleInterop: true,
  } }).outputText, filename);
  return loaded.exports;
}
const { createPersonalApi, signalPresentation, teamsHistory, readPendingAnalysis, pendingAnalysisKey, sendSelfAnalysis } = load("personalDashboardData.ts");
const Summary = load("ConnectivitySummary.tsx").default;
const originalFetch = global.fetch;
afterEach(() => global.fetch = originalFetch);
const storage = () => { const values = new Map(); return { getItem: key => values.get(key) || null, setItem: (key, value) => values.set(key, value), removeItem: key => values.delete(key) }; };
const operation = { requestId: "11111111111141118111111111111111", targetUserId: 17, meetingSessionId: null, mode: "Report" };

test("authenticated reads use the supplied account token, cancellation, and no-store", async () => {
  const signal = new AbortController().signal;
  global.fetch = async (url, options) => {
    assert.equal(url, "https://example.invalid/api/User/signal/current");
    assert.equal(options.headers.Authorization, "Bearer synthetic"); assert.equal(options.signal, signal); assert.equal(options.cache, "no-store");
    return Response.json({ SignalScore: 92 });
  };
  assert.deepEqual(await createPersonalApi("https://example.invalid/", "synthetic", () => assert.fail())("User/signal/current", { signal }), { SignalScore: 92 });
});
for (const status of [401, 403, 404, 503]) test(`HTTP ${status} returns safe feedback and only expiry clears the session`, async () => {
  let expired = 0;
  global.fetch = async () => new Response("private diagnostic details", { status });
  await assert.rejects(createPersonalApi("https://example.invalid", "synthetic", () => expired++)("meetings/teams-1"), error => error.status === status && !error.message.includes("private"));
  assert.equal(expired, status === 401 ? 1 : 0);
});
test("lost response and reload retain the exact paid request, isolated by account", async () => {
  const saved = storage(), calls = [];
  const api = async (path, options) => { calls.push(JSON.parse(options.body)); if (calls.length === 1) throw new Error("lost"); return { state: "Approved" }; };
  await assert.rejects(sendSelfAnalysis(api, saved, operation, new AbortController().signal));
  assert.equal(readPendingAnalysis(saved, 18), null);
  await sendSelfAnalysis(api, saved, readPendingAnalysis(saved, 17), new AbortController().signal);
  assert.deepEqual(calls, [operation, operation]); assert.equal(readPendingAnalysis(saved, 17), null);
});
test("free Test and paid Report keep the server mode and self target, with no meeting", async () => {
  for (const mode of ["Test", "Report"]) {
    const saved = storage(); let submitted;
    await sendSelfAnalysis(async (_, options) => { submitted = JSON.parse(options.body); assert.deepEqual(readPendingAnalysis(saved, 17), submitted); return {}; }, saved, { ...operation, mode }, new AbortController().signal);
    assert.equal(submitted.mode, mode); assert.equal(submitted.meetingSessionId, null); assert.equal(submitted.targetUserId, 17);
  }
});
test("storage failure prevents sending a request that could not be recovered", async () => {
  await assert.rejects(sendSelfAnalysis(() => assert.fail(), { setItem() { throw new Error("unavailable"); } }, operation, new AbortController().signal));
});
test("malformed and cross-context pending operations cannot target another account or meeting", () => {
  const saved = storage();
  for (const value of ["{", JSON.stringify({ ...operation, targetUserId: 18 }), JSON.stringify({ ...operation, meetingSessionId: 2 }), JSON.stringify({ ...operation, mode: "Unknown" })]) {
    saved.setItem(pendingAnalysisKey(17), value); assert.equal(readPendingAnalysis(saved, 17), null);
  }
});
test("history contains only Teams sessions, newest first", () => {
  const meeting = { source: "Teams", meetingId: "teams-1", createdAt: "2026-01-01T00:00:00Z" };
  assert.deepEqual(teamsHistory([meeting, { ...meeting, source: "Website" }, { ...meeting, meetingId: "teams-02" }, { ...meeting, meetingId: "teams-2", createdAt: "2026-02-01T00:00:00Z" }]).map(m => m.meetingId), ["teams-2", "teams-1"]);
});
test("missing, offline and stale signals never render a zero or healthy fallback", () => {
  for (const signal of [null, { SignalScore: 95, SignalOverallStatus: "Offline" }, { SignalScore: 95, SignalOverallStatus: "Healthy", LastUpdated: "2026-01-01T00:00:00Z" }]) {
    const value = signalPresentation(signal, Date.parse("2026-01-02T00:00:00Z")); assert.equal(value.score, null); assert.equal(value.hasData, false);
    const html = renderToStaticMarkup(React.createElement(Summary, { name: "Synthetic", score: value.score, hasData: value.hasData, overallStatus: value.overall, workspaceStatus: null, networkStatus: null, deviceStatus: null }));
    assert.ok(!html.includes("0%") && !html.includes(">Healthy<"));
  }
});
test("shared telemetry summary retains real scores and independent category statuses", () => {
  const html = renderToStaticMarkup(React.createElement(Summary, { name: "Synthetic", score: 92, hasData: true, overallStatus: "Healthy", workspaceStatus: null, networkStatus: "Poor", deviceStatus: "Fair" }));
  assert.ok(html.includes("92%") && html.includes(">Unknown<") && html.includes(">Poor<") && html.includes(">Fair<"));
});
