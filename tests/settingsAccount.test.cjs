const { test, afterEach } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");
const Module = require("node:module");
const filename = path.resolve(__dirname, "../src/Tab/settingsAccount.ts");
const loaded = new Module(filename, module);
loaded.filename = filename;
loaded.paths = module.paths;
loaded._compile(ts.transpileModule(fs.readFileSync(filename, "utf8"), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
}).outputText, filename);
const { loadReportAccount, saveAccountProfile, subscriptionStatus, periodDate } = loaded.exports;
const originalFetch = global.fetch;
afterEach(() => { global.fetch = originalFetch; });

test("account overview uses the current Teams bearer and preserves server credit totals", async () => {
  const account = { purchasedCredits: 3, teamAllocatedCredits: 2, teamFundAvailable: 4,
    availableReports: 16, subscription: { remaining: 7, allowance: 10, status: "active" } };
  const controller = new AbortController();
  global.fetch = async (url, options) => {
    assert.equal(url, "https://backend.example/api/credits/account");
    assert.equal(options.headers.Authorization, "Bearer synthetic-session");
    assert.equal(options.signal, controller.signal);
    return Response.json(account);
  };
  assert.deepEqual(await loadReportAccount("https://backend.example", "synthetic-session", controller.signal), account);
});

for (const status of [401, 403, 503]) test(`account HTTP ${status} uses safe actionable feedback`, async () => {
  global.fetch = async () => new Response("private backend detail", { status });
  await assert.rejects(loadReportAccount("https://backend.example", "synthetic-session", new AbortController().signal),
    status === 401 ? /session has expired/ : /could not be loaded/);
});

test("invalid credit data cannot masquerade as a zero balance", async () => {
  global.fetch = async () => Response.json({ purchasedCredits: "3", availableReports: null });
  await assert.rejects(loadReportAccount("https://backend.example", "synthetic-session", new AbortController().signal), /could not be loaded/);
});

test("Save profile submits all fields together using the website contract and Teams session", async () => {
  let calls = 0;
  global.fetch = async (url, options) => {
    calls++;
    assert.equal(url, "https://backend.example/api/User/UpdateUser");
    assert.equal(options.method, "PUT");
    assert.deepEqual(options.headers, { Authorization: "Bearer synthetic-session" });
    assert.deepEqual(Object.fromEntries(options.body), {
      userId: "17", userFirstName: "Test", userLastName: "Person", userEmail: "test@example.invalid",
    });
    return Response.json({ status: "success" });
  };
  await saveAccountProfile("https://backend.example", "synthetic-session", 17, " Test ", " Person ", " test@example.invalid ");
  assert.equal(calls, 1);
});

for (const message of ["Create a password before updating your email.", "That email is already in use."])
  test(`profile preserves server validation: ${message}`, async () => {
    global.fetch = async () => Response.json({ message }, { status: 409 });
    await assert.rejects(saveAccountProfile("https://backend.example", "synthetic-session", 17, "Test", "Person", "test@example.invalid"),
      error => error.message === message);
  });

test("profile hides arbitrary server errors and does not retry mutations", async () => {
  let calls = 0;
  global.fetch = async () => { calls++; return Response.json({ message: "private backend detail" }, { status: 500 }); };
  await assert.rejects(saveAccountProfile("https://backend.example", "synthetic-session", 17, "Test", "Person", "test@example.invalid"), /could not be saved/);
  assert.equal(calls, 1);
});

test("subscription display distinguishes cancellation, delinquency and unknown state", () => {
  assert.equal(subscriptionStatus(null), "Pay as you go");
  assert.equal(subscriptionStatus({ status: "active", cancelAtPeriodEnd: true }), "Ends after this period");
  assert.equal(subscriptionStatus({ status: "past_due" }), "Payment overdue");
  assert.equal(subscriptionStatus({ status: "unknown" }), "Needs attention");
  assert.equal(periodDate("invalid"), null);
  assert.equal(periodDate(null), null);
});
