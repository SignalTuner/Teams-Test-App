const { test } = require("node:test");
const assert = require("node:assert/strict");
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "playwright");
const { startStaticServer } = require("./dreamhostServer.cjs");

async function scenario(t, state = { request: null }) {
  const server = await startStaticServer(); t.after(server.close);
  const browser = await chromium.launch({ headless: true, ...(process.env.BROWSER_EXECUTABLE ? { executablePath: process.env.BROWSER_EXECUTABLE } : {}) });
  t.after(() => browser.close());
  const page = await browser.newPage(); page.setDefaultTimeout(15000);
  const errors = [], calls = [], assets = [];
  page.on("pageerror", error => errors.push(error.message));
  page.on("response", response => { if (response.url().includes("/assets/")) assets.push({ url: response.url(), status: response.status() }); });
  // All backend traffic is synthetic. The production bundle and real TeamsJS stay intact.
  await page.route("https://signaltuner-back.azurewebsites.net/**", async route => {
    const request = route.request(), url = new URL(request.url());
    calls.push({ path: url.pathname, method: request.method(), authorization: request.headers().authorization });
    const user = { userId: 17, displayName: "Synthetic Person", firstName: "Synthetic", lastName: "Person", email: "synthetic@example.invalid", authProvider: "email", clientIsActive: true, credits: 2, hasPassword: true };
    const data = url.pathname.endsWith("/User/Login") ? { token: "synthetic-session", user }
      : url.pathname.endsWith("/auth/me") ? user
      : url.pathname.endsWith("/credits/account") ? { purchasedCredits: 2, availableReports: 2, subscription: null }
      : url.pathname.endsWith("/current-device") ? { connected: true, supportsTestAnalysis: true }
      : url.pathname.endsWith("/remote-analysis/latest") ? { request: state.request }
      : url.pathname.includes("/remote-analysis/") ? state.request
      : url.pathname.endsWith("/meetings") ? { meetings: [] }
      : url.pathname.includes("GetServiceIncidents") ? []
      : url.pathname.includes("GetServiceStatus") ? { CurrentStatus: "Operational", DisplayName: "Microsoft Teams" }
      : { SignalScore: 94, SignalOverallStatus: "Healthy", LastUpdated: new Date().toISOString() };
    await route.fulfill({ contentType: "application/json", body: JSON.stringify(data) });
  });
  return { page, origin: server.origin, errors, calls, assets };
}

test("compiled static app renders direct entry/legal/config/auth routes and survives refresh", async t => {
  const { page, origin, errors, assets } = await scenario(t);
  for (const [route, selector] of [
    ["/tabs/home", "#login-email"], ["/tabs/config", ".configPage"],
    ["/tabs/privacy", ".legalPanel"], ["/tabs/terms", ".legalPanel"],
    ["/tabs/login", "#login-email"], ["/tabs/create-account", "#register-email"],
    ["/tabs/support", "#support-email"], ["/forgot-password", "#forgot-password-email"],
    ["/create-account", "#register-email"], ["/support", "#support-email"],
  ]) {
    await page.goto(origin + route); await page.locator(selector).waitFor();
    await page.reload(); await page.locator(selector).waitFor();
  }
  assert.ok(assets.length > 0); assert.ok(assets.every(asset => asset.status === 200));
  assert.deepEqual(errors, []);
});

test("static dashboard renders complete jitter histories across live updates and reloads", async t => {
  const state = { request: { requestId: "synthetic-jitter-analysis", mode: "Test", state: "Running", creditState: "None", deliveries: [],
    liveTelemetry: { ageSeconds: 0, snapshot: { elapsedSeconds: 20, jitterMs: 9, jitterHistoryMs: [0, 15, 30, null, 50, 9] } } } };
  const { page, origin, errors } = await scenario(t, state);
  await page.addInitScript(() => localStorage.setItem("signaltunerSessionToken", "synthetic-session"));
  await page.goto(origin + "/tabs/home");
  const chart = page.locator(".live-metric--jitterHistoryMs svg");
  await chart.waitFor();
  await page.waitForFunction(() => document.querySelectorAll(".live-metric--jitterHistoryMs circle").length === 5);
  assert.equal(await chart.locator("g > line").count(), 3);

  // A fresh snapshot replaces the window, including real zeros and missing probes.
  state.request.liveTelemetry.snapshot = { elapsedSeconds: 22, jitterMs: 5, jitterHistoryMs: [0, 15, 30, null, 50, 9, 0, 5] };
  await page.waitForFunction(() => document.querySelectorAll(".live-metric--jitterHistoryMs circle").length === 7);
  assert.match(await chart.getAttribute("aria-label"), /5 ms/);
  await page.reload();
  await page.waitForFunction(() => document.querySelectorAll(".live-metric--jitterHistoryMs circle").length === 7);

  // Legacy telemetry supplies only one real measurement; polling cannot invent history.
  state.request.liveTelemetry.snapshot = { elapsedSeconds: 24, jitterMs: 4 };
  await page.waitForFunction(() => document.querySelectorAll(".live-metric--jitterHistoryMs circle").length === 1);
  assert.match(await chart.getAttribute("aria-label"), /4 ms/);
  assert.deepEqual(errors, []);
});

test("static email sign-in preserves same-origin returnUrl and uses Azure bearer API", async t => {
  const { page, origin, errors, calls } = await scenario(t);
  const destination = origin + "/tabs/home?view=synthetic%20meeting&source=return";
  await page.goto(origin + "/login?returnUrl=" + encodeURIComponent(destination));
  await page.locator("#login-email").fill("synthetic@example.invalid");
  await page.locator("#login-password").fill("Synthetic-test-only-123!");
  await page.locator(".authForm button[type=submit]").click();
  await page.getByRole("heading", { name: "Your Connectivity Health" }).waitFor();
  assert.equal(page.url(), destination);
  await page.reload(); await page.getByRole("heading", { name: "Your Connectivity Health" }).waitFor();
  assert.equal(page.url(), destination);
  assert.ok(calls.some(call => call.path.endsWith("/User/Login") && call.method === "POST"));
  assert.ok(calls.some(call => call.path.endsWith("/credits/account") && call.authorization === "Bearer synthetic-session"));
  assert.equal(await page.getByRole("button", { name: "Continue with Microsoft Teams" }).count(), 0);
  assert.deepEqual(errors, []);
});
