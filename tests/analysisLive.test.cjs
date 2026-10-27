const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");
const Module = require("node:module");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");

function load(name) {
  const filename = path.resolve(__dirname, "../src/Tab", name);
  const compiled = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.ReactJSX },
  }).outputText;
  const loaded = new Module(filename, module);
  loaded.filename = filename;
  loaded.paths = module.paths;
  const originalRequire = loaded.require.bind(loaded);
  loaded.require = name => {
    if (name.endsWith(".css")) return {};
    if (name.startsWith(".")) {
      const source = [".tsx", ".ts"].map(extension => path.resolve(path.dirname(filename), name + extension)).find(fs.existsSync);
      if (source) return load(source);
    }
    return originalRequire(name);
  };
  loaded._compile(compiled, filename);
  return loaded.exports;
}
const { pollAnalysisRequest, analysisStatusPath } = load("analysisPolling.ts");
const { default: Charts, liveMetrics, jitterColor } = load("LiveAnalysisCharts.tsx");
const { default: RemoteAnalysisButton } = load("RemoteAnalysisButton.tsx");
const settle = async () => { for (let i = 0; i < 8; i++) await Promise.resolve(); };
const active = { state: "Running", deliveries: [] };
test("idle participant analysis keeps all charts and displays the guidance only once", () => {
  for (const expanded of [false, true]) {
    const markup = renderToStaticMarkup(React.createElement(RemoteAnalysisButton, {
      apiBaseUrl: "https://api.example.invalid", token: "synthetic-session", targetUserId: 2, targetName: "Test member",
      expanded, renderInline: ({ action, panel }) => React.createElement("div", null, action, panel),
    }));
    assert.equal((markup.match(/role="img"/g) || []).length, 8);
    assert.equal((markup.match(/Click Analyze to begin/g) || []).length, 1);
    assert.ok(!markup.includes("<circle") && !markup.includes("<rect"));
    assert.equal(markup.includes('hidden=""'), !expanded);
  }
});

test("empty charts use matching idle, waiting and finished guidance", () => {
  for (const requestState of [undefined, "Pending", "Approved", "Running", "Completed", "Cancelled", "Failed", "Denied", "Expired", "NoCredits"]) {
    const markup = renderToStaticMarkup(React.createElement(Charts, { requestState, running: requestState === "Running" }));
    const expected = !requestState ? "Click Analyze to begin a 5-minute rapid sampling session. See live connection health data and receive a detailed PDF report."
      : ["Pending", "Approved", "Running"].includes(requestState) ? "Waiting for live samples. Live charts require the updated SignalTuner desktop app."
      : "No samples were received for this analysis.";
    assert.ok(markup.includes(expected), requestState);
    assert.equal((markup.match(/role="img"/g) || []).length, 8);
    assert.ok(!markup.includes("<circle") && !markup.includes("<rect"));
  }
});

test("sample freshness and completion use the website messages", () => {
  for (const [running, ageSeconds, expected] of [
    [true, 0, "Live updates every two seconds"],
    [true, 9, "Live data is delayed. Last update 9 seconds ago."],
    [false, 9, "Last reported samples. Analysis is no longer running."],
  ]) {
    const markup = renderToStaticMarkup(React.createElement(Charts, {
      running, telemetry: { ageSeconds, snapshot: { elapsedSeconds: 20, cpuPercent: [42] } },
    }));
    assert.ok(markup.includes(expected));
    assert.ok(markup.includes("20s / 300s sampled"));
    assert.ok(!markup.includes("Click Analyze to begin"));
  }
});
function setup(t, read) {
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const visibility = new EventTarget();
  visibility.hidden = false;
  const data = [], errors = [], lost = [];
  let version = 0, mutating = false;
  const stop = pollAnalysisRequest({ read, visibility, onData: value => data.push(value),
    onError: () => errors.push(true), onAccessLost: () => lost.push(true),
    version: () => version, isMutating: () => mutating });
  t.after(stop);
  return { visibility, data, errors, lost, stop,
    mutate: () => { version++; mutating = true; }, resume: () => { mutating = false; } };
}

test("requests live telemetry immediately, every two seconds while active, and stops at completion", async t => {
  assert.equal(analysisStatusPath("synthetic-id"), "/synthetic-id?includeLive=true");
  let calls = 0;
  const h = setup(t, async () => ++calls < 3 ? active : { state: "Completed", deliveries: [] });
  await settle(); assert.equal(calls, 1);
  t.mock.timers.tick(1999); await settle(); assert.equal(calls, 1);
  t.mock.timers.tick(1); await settle(); assert.equal(calls, 2);
  t.mock.timers.tick(2000); await settle(); assert.equal(calls, 3);
  t.mock.timers.tick(30000); await settle(); assert.equal(calls, 3);
  assert.equal(h.data.at(-1).state, "Completed");
});

test("does not overlap requests and discards responses older than a cancellation", async t => {
  let resolve, calls = 0;
  const h = setup(t, () => { calls++; return new Promise(done => { resolve = done; }); });
  t.mock.timers.tick(6000); await settle(); assert.equal(calls, 1);
  h.mutate(); resolve(active); await settle(); assert.equal(h.data.length, 0);
  t.mock.timers.tick(2000); await settle(); assert.equal(calls, 1);
  h.resume(); t.mock.timers.tick(2000); await settle(); assert.equal(calls, 2);
  h.stop(); resolve(active); await settle(); assert.equal(h.data.length, 0);
});

test("pauses hidden tabs, resumes immediately and aborts on cleanup", async t => {
  let calls = 0, signal;
  const h = setup(t, async value => { calls++; signal = value; return active; });
  await settle();
  h.visibility.hidden = true; h.visibility.dispatchEvent(new Event("visibilitychange"));
  assert.equal(signal.aborted, true);
  t.mock.timers.tick(20000); await settle(); assert.equal(calls, 1);
  h.visibility.hidden = false; h.visibility.dispatchEvent(new Event("visibilitychange"));
  await settle(); assert.equal(calls, 2);
  h.stop(); assert.equal(signal.aborted, true);
  t.mock.timers.tick(20000); await settle(); assert.equal(calls, 2);
});

test("backs off failures, times out stalled reads, and clears access on authorization failure", async t => {
  let calls = 0, signal;
  const h = setup(t, value => {
    calls++; signal = value;
    if (calls === 1) return Promise.reject(new Error("synthetic failure"));
    if (calls === 2) return new Promise((_, reject) => value.addEventListener("abort", () => reject(new Error("timeout"))));
    return Promise.reject(Object.assign(new Error("unavailable"), { status: 403 }));
  });
  await settle(); assert.equal(h.errors.length, 1);
  t.mock.timers.tick(3999); await settle(); assert.equal(calls, 1);
  t.mock.timers.tick(1); await settle(); assert.equal(calls, 2);
  t.mock.timers.tick(8000); await settle(); assert.equal(signal.aborted, true);
  t.mock.timers.tick(8000); await settle(); assert.equal(h.lost.length, 1);
  t.mock.timers.tick(60000); await settle(); assert.equal(calls, 3);
});

test("delivery-only status uses ten seconds and does not restart an analysis", async t => {
  let calls = 0;
  setup(t, async () => { calls++; return { state: "Completed", deliveries: [{ state: calls === 1 ? "Queued" : "Accepted" }] }; });
  await settle();
  t.mock.timers.tick(9999); await settle(); assert.equal(calls, 1);
  t.mock.timers.tick(1); await settle(); assert.equal(calls, 2);
  t.mock.timers.tick(30000); await settle(); assert.equal(calls, 2);
});

test("renders website chart set with real zeros, gaps, jitter, VPN and completed sample labels", () => {
  const markup = renderToStaticMarkup(React.createElement(Charts, { running: false, telemetry: {
    ageSeconds: 0, snapshot: { elapsedSeconds: 300, wifiStrength: [61, 70], downloadMbps: [0, 4], uploadMbps: [3.8],
      packetLossPercent: [0, 10], latencyMs: [10, null, 1000], cpuPercent: [0, 95], memoryPercent: [98], jitterMs: 0, vpnStatus: "Detected" },
  } }));
  assert.equal((markup.match(/role="img"/g) || []).length, 8);
  for (const value of ["70%", "0 ms", "VPN Detected", "Last reported samples", "300s / 300s sampled"]) assert.ok(markup.includes(value));
  assert.equal((markup.match(/<rect /g) || []).length, 3);
  assert.ok(!markup.includes("NaN") && !markup.includes("Infinity"));
  const latency = markup.match(/<svg[^>]*aria-label="Latency chart[^]*?<\/svg>/)[0];
  assert.equal((latency.match(/<circle /g) || []).length, 2);
  assert.ok(!latency.includes("<g><line"));
});

test("uses the website severity boundaries and keeps unavailable data distinct from zero", () => {
  const boundaries = [[61, 46, 31, 30], [3, 1.8, .8, 0], [3.8, 2.6, 1.2, 0], [1.9, 3.9, 5.9, 6], [20, 50, 99, 100], [15, 30, 50, 51], [79, 80, 90, 95], [92, 93, 96, 98]];
  liveMetrics.forEach((metric, i) => assert.deepEqual(boundaries[i].map(metric.color), ["#3BB537", "#F9AE00", "#FC6F20", "#EC2F3E"]));
  assert.deepEqual([15, 30, 50, 51].map(jitterColor), ["#3BB537", "#F9AE00", "#FC6F20", "#EC2F3E"]);
  const markup = renderToStaticMarkup(React.createElement(Charts, { running: true }));
  assert.ok(markup.includes("Detection Unavailable"));
  assert.ok(markup.includes("require the updated SignalTuner desktop app"));
  assert.ok(!markup.includes("<circle") && !markup.includes("<rect"));
});

test("jitter plots real bounded histories with the latency scale, severity colors, zeros and gaps", () => {
  const render = snapshot => renderToStaticMarkup(React.createElement(Charts, { running: true, telemetry: { snapshot } }));
  const chart = snapshot => render(snapshot).match(/<svg[^>]*aria-label="Jitter chart[^]*?<\/svg>/)[0];
  const history = chart({ jitterMs: 9, jitterHistoryMs: [0, 15, 30, null, 50, 500] });
  assert.ok(history.includes("Jitter chart: 500 ms"));
  assert.ok(history.includes('text-anchor="end">500</text>'));
  assert.equal((history.match(/<circle /g) || []).length, 5);
  assert.equal((history.match(/<g><line /g) || []).length, 3);
  assert.ok(history.includes('cy="38"'));
  for (const color of ["#3BB537", "#F9AE00", "#FC6F20", "#EC2F3E"]) assert.ok(history.includes(`fill="${color}"`));
  assert.equal((chart({ jitterHistoryMs: Array(15).fill(0) }).match(/<circle /g) || []).length, 10);
  assert.equal((chart({ jitterMs: 0 }).match(/<circle /g) || []).length, 1);
  assert.ok(chart({ jitterMs: 9, jitterHistoryMs: [0, null] }).includes("Jitter chart: Unavailable"));
});
