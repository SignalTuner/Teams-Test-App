const { test, afterEach } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");
const Module = require("node:module");
const filename = path.resolve(__dirname, "../src/Tab/meetingAnalysis.ts");
const compiled = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
}).outputText;
const loaded = new Module(filename, module);
loaded.filename = filename;
loaded.paths = module.paths;
loaded._compile(compiled, filename);
const { activeMeetingParticipantIds, readMeetingAnalysisOperation, validateMeetingAnalysisResults, requestMeetingAnalysis } = loaded.exports;
const originalFetch = global.fetch;
afterEach(() => { global.fetch = originalFetch; });
const operation = { requestId: "11111111111141118111111111111111", participantUserIds: [1, 2] };
const accepted = { targetUserId: 1, request: { requestId: "22222222222242228222222222222222", targetUserId: 1, state: "Approved" } };
const rejected = { targetUserId: 2, error: "The desktop client is not connected.", statusCode: 409 };
const response = () => ({ batchId: operation.requestId, meetingSessionId: 5, results: [accepted, rejected] });

test("selects all active accounts including self, excluding offline, missing and duplicate identities", () => {
  assert.deepEqual(activeMeetingParticipantIds([
    { userId: 1, clientDataStatus: "active" }, { userId: 2, clientDataStatus: "active" },
    { userId: 3, clientDataStatus: "no_data" }, { userId: 4, clientDataStatus: "offline" },
    { userId: 0, clientDataStatus: "active" }, { userId: NaN, clientDataStatus: "active" },
    { userId: 1, clientDataStatus: "active" },
  ]), [1, 2]);
  assert.deepEqual(activeMeetingParticipantIds([]), []);
});

test("reload retains the confirmed roster and batch; malformed saved operations cannot create new charges", () => {
  assert.deepEqual(readMeetingAnalysisOperation(JSON.stringify(operation)), operation);
  assert.equal(readMeetingAnalysisOperation(null), null);
  for (const value of ["{", "null", "{}", JSON.stringify({ ...operation, participantUserIds: [1, 1] }),
    JSON.stringify({ ...operation, participantUserIds: [] }), JSON.stringify({ ...operation, participantUserIds: [0] }),
    JSON.stringify({ ...operation, requestId: "invalid" })]) assert.throws(() => readMeetingAnalysisOperation(value));
});

test("partial acceptance remains actionable; incomplete, duplicate and cross-meeting results stay uncertain", () => {
  assert.deepEqual(validateMeetingAnalysisResults(response(), operation, 5), [accepted, rejected]);
  for (const value of [null, {}, { ...response(), batchId: "other" }, { ...response(), meetingSessionId: 6 },
    { ...response(), results: [accepted] }, { ...response(), results: [accepted, accepted] },
    { ...response(), results: [accepted, { targetUserId: 3, error: "Offline" }] },
    { ...response(), results: [accepted, { targetUserId: 2 }] },
    { ...response(), results: [{ ...accepted, request: { ...accepted.request, targetUserId: 2 } }, rejected] }])
    assert.throws(() => validateMeetingAnalysisResults(value, operation, 5));
});

test("retry after a lost response sends the exact original batch and meeting with authenticated requests", async () => {
  const calls = [];
  global.fetch = async (url, options) => {
    calls.push({ url, options });
    if (calls.length === 1) throw new TypeError("Synthetic lost response");
    return { ok: true, json: async () => response() };
  };
  const signal = new AbortController().signal;
  await assert.rejects(requestMeetingAnalysis("https://api.example.invalid/", "synthetic-session", 5, operation, signal));
  const restored = readMeetingAnalysisOperation(JSON.stringify(operation));
  const results = await requestMeetingAnalysis("https://api.example.invalid/", "synthetic-session", 5, restored, signal);
  assert.deepEqual(results, [accepted, rejected]);
  assert.equal(calls[0].options.body, calls[1].options.body);
  assert.equal(calls[1].url, "https://api.example.invalid/api/remote-analysis/teams-meeting");
  assert.equal(calls[1].options.headers.Authorization, "Bearer synthetic-session");
  assert.deepEqual(JSON.parse(calls[1].options.body), { ...operation, meetingSessionId: 5 });
  assert.equal(calls[1].options.signal, signal);
});

test("server failures do not expose internal response bodies or discard the retry identifier", async () => {
  for (const status of [401, 403, 404, 429, 503]) {
    global.fetch = async () => ({ ok: false, status, json: async () => ({ message: "private-server-detail" }) });
    await assert.rejects(requestMeetingAnalysis("https://api.example.invalid", "synthetic-session", 5, operation, new AbortController().signal),
      error => !error.message.includes("private-server-detail"));
    assert.equal(operation.requestId, "11111111111141118111111111111111");
  }
});

test("malformed response bodies show a safe retry message", async () => {
  global.fetch = async () => ({ ok: true, json: async () => { throw new SyntaxError("private-server-detail"); } });
  await assert.rejects(requestMeetingAnalysis("https://api.example.invalid", "synthetic-session", 5, operation, new AbortController().signal),
    error => error.message.includes("Retry this same meeting analysis") && !error.message.includes("private-server-detail"));
});
