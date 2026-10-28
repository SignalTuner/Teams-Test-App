// Run against localhost Vite with VITE_SIGNALTUNER_API_URL=http://127.0.0.1:4317/backend.
// API and Teams SDK responses are synthetic and intercepted only in this test browser.
const { test } = require("node:test");
const assert = require("node:assert/strict");
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "playwright");
const origin = process.env.PERSONAL_TEST_ORIGIN || "http://127.0.0.1:4317";
async function scenario(t, options = {}) {
  const browser = await chromium.launch({ headless: true, timeout: 15000, ...(process.env.BROWSER_EXECUTABLE ? { executablePath: process.env.BROWSER_EXECUTABLE } : {}) });
  t.after(() => browser.close());
  const context = await browser.newContext({ viewport: { width: options.width || 1100, height: 900 } });
  t.after(() => context.close());
  const page = await context.newPage();
  page.setDefaultTimeout(10000);
  const state = { connected: options.connected ?? true, credits: options.credits ?? 2, request: null, calls: [], errors: [], ...options };
  page.on("pageerror", error => state.errors.push(error.message));
  await page.addInitScript(({ saved }) => { if (saved && !localStorage.getItem("initialized")) { localStorage.setItem("signaltunerSessionToken", "synthetic-session"); localStorage.setItem("initialized", "yes"); } }, { saved: options.saved !== false });
  await page.route("**/@microsoft_teams-js.js*", route => route.fulfill({ contentType: "application/javascript", body: `
    export const app = { initialize: async () => {}, getContext: async () => {
      ${options.failContext ? 'throw new Error("Synthetic unavailable context");' : ''}
      ${options.delayContext ? 'await new Promise(resolve => { window.resolveSyntheticHost = resolve; });' : ''}
      return {app:{theme:"default",host:{name:"Teams",clientType:${JSON.stringify(options.clientType ?? "desktop")}}},user:{id:"synthetic-oid",tenant:{id:"synthetic-tenant"}}, ${options.meeting ? 'meeting:{id:"synthetic-meeting"},' : ''}};
    }, registerOnThemeChangeHandler: () => {} };
    window.syntheticSsoTokenRequests = 0;
    export const authentication = { getAuthToken: async () => { window.syntheticSsoTokenRequests++; return "synthetic-teams-token"; } };
    export const meeting = { getMeetingDetails: callback => callback(null, null) };
    export const pages = {};` }));
  await page.route("**/api/**", async route => {
    const request = route.request(), url = new URL(request.url());
    const path = url.pathname.split("/api/")[1];
    state.calls.push({ path, method: request.method(), body: request.postData() ? JSON.parse(request.postData()) : null });
    const user = { userId: state.userId || 17, displayName: state.userId === 18 ? "Second Synthetic" : "Synthetic Person", firstName: "Synthetic", lastName: "Person", email: "synthetic@example.invalid", authProvider: "teams_sso", clientIsActive: state.connected, credits: state.credits, hasPassword: true };
    const reply = (value, status = 200) => route.fulfill({ status, contentType: "application/json", body: JSON.stringify(value) });
    if (path === "auth/me") return reply(user, state.expired ? 401 : state.restorationFailure ? 503 : 200);
    if (path === "User/Login") return reply({ token: "synthetic-session", user: { ...user, authProvider: "email" }, profileRequired: state.profileRequired || false });
    if (path === "User/teams-sso") return state.ssoFailure ? reply({ message: "Link your existing account before using Teams sign-in." }, 409)
      : reply({ token: "synthetic-session", user, profileRequired: state.profileRequired || false });
    if (path === "User/signal/current") return state.noSignal ? reply({}, 404) : reply({ SignalScore: 94, SignalOverallStatus: state.connected ? "Healthy" : "Offline", SignalWifiStatus: "Healthy", SignalBandwidthStatus: "Fair", SignalSystemStatus: "Healthy", LastUpdated: new Date().toISOString() });
    if (path === "credits/account") return reply({ purchasedCredits: state.purchasedCredits ?? state.credits, availableReports: state.credits, subscription: state.subscription || null, isCustomer: state.isCustomer || false });
    if (path === "User/GetServiceStatus/3") return reply({ CurrentStatus: "Operational", DisplayName: "Microsoft Teams" });
    if (path === "User/GetServiceIncidents/3") return reply([]);
    if (path === "User/activation-code") return reply({ activationCode: request.method() === "POST" ? "SYNTHETIC-CODE" : null, expiresAtUtc: new Date(Date.now() + 300000).toISOString(), serverNowUtc: new Date().toISOString() });
    if (path === "remote-analysis/current-device") return reply({ connected: state.connected, supportsTestAnalysis: state.supportsTest !== false });
    if (path === "remote-analysis/latest") return reply({ request: state.request });
    if (path === "remote-analysis" && request.method() === "POST") {
      const body = JSON.parse(request.postData());
      state.request = { ...body, state: "Running", creditState: body.mode === "Test" ? "None" : "Reserved", deliveries: [], liveTelemetry: { ageSeconds: 0, snapshot: { elapsedSeconds: 15, cpuPercent: [10, 12] } } };
      if (state.loseResponse) { state.loseResponse = false; return route.abort(); }
      return reply(state.request);
    }
    if (path.startsWith("remote-analysis/")) {
      if (path.endsWith("/cancel")) state.request = { ...state.request, state: "Cancelled", creditState: "Released" };
      return reply(state.request || {});
    }
    const meeting = { meetingId: "teams-4", meetingSessionId: 4, source: "Teams", name: "Synthetic retrospective", createdAt: "2026-01-01T12:00:00Z", participantCount: 1 };
    if (path === "meetings") return reply({ meetings: state.emptyMeetings ? [] : [meeting] });
    if (path === "meetings/teams-4") return state.denyMeeting ? reply({}, 404) : reply({ meeting, canAnalyze: true, participants: [{ participantId: "teams-user-17", userId: 17, name: "Synthetic Person", email: user.email, meetingRole: "Presenter", signal: { signalScore: 94, signalOverallStatus: "Healthy", signalWifiStatus: "Healthy" } }], requests: [] });
    if (path === "TeamsMeetings/join" || path === "TeamsMeetings/4/dashboard") return reply({ meetingSessionId: 4, currentUser: user, teamsServiceHealth: { activeIncidents: [] }, participants: [] });
    return reply({}, 404);
  });
  await page.goto(`${origin}/tabs/home`);
  return { page, state };
}
const home = page => page.getByRole("heading", { name: "Your Connectivity Health" }).waitFor();
async function accountMenu(page, name) { await page.locator(".accountButton").click(); await page.getByRole("menuitem", { name, exact: true }).click(); }

for (const clientType of ["ios", "android", "ipados"]) test(`${clientType} hides direct and indirect commerce while retaining allowance and analysis`, async t => {
  const { page, state } = await scenario(t, { clientType, width: 1440, credits: 8, purchasedCredits: 1,
    isCustomer: true, subscription: { name: "Synthetic Plan", status: "trialing", remaining: 7, allowance: 10 } });
  await home(page);
  assert.equal(await page.locator('a[href^="https:"]').count(), 0);
  const report = page.getByRole("button", { name: "Analysis + Report", exact: true });
  await report.waitFor();
  await page.waitForFunction(() => ![...document.querySelectorAll("button")].find(b => b.textContent === "Analysis + Report")?.disabled);
  assert.equal(await report.isEnabled(), true);
  await accountMenu(page, "Settings");
  await page.getByText("Available report allowance", { exact: true }).waitFor();
  assert.match(await page.locator(".settings-credit-value").innerText(), /8 reports/);
  assert.equal(await page.locator('a[href^="https:"], .settings-plan-summary, [data-payment-disclosure]').count(), 0);
  assert.doesNotMatch(await page.locator(".settings-account").innerText(), /buy|purchase|upgrade|pricing|trial|payment|billing|pay as you go/i);
  await page.getByRole("button", { name: "Contact support about account deletion" }).click();
  await page.getByRole("heading", { name: "Contact support" }).waitFor();
  await page.goto(`${origin}/tabs/home`); await home(page);
  await page.getByRole("button", { name: "Analysis + Report", exact: true }).click();
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  assert.equal(state.request.mode, "Report");
  assert.deepEqual(state.errors, []);
});

test("mobile zero allowance remains neutral and Test Analysis still runs", async t => {
  const { page, state } = await scenario(t, { clientType: "android", credits: 0 });
  await home(page);
  await page.getByText(/No report allowance is available/).waitFor();
  assert.equal(await page.getByRole("button", { name: "Analysis + Report", exact: true }).isDisabled(), true);
  assert.equal(await page.locator('a[href^="https:"]').count(), 0);
  await page.getByRole("button", { name: "Test Analysis", exact: true }).click();
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  assert.equal(state.request.mode, "Test");
});

for (const clientType of ["desktop", "web", "macos"]) test(`${clientType} retains disclosed external purchase links even at phone width`, async t => {
  const { page } = await scenario(t, { clientType, width: 390, isCustomer: true });
  await home(page);
  await page.getByRole("button", { name: "View incidents", exact: false }).click();
  await page.getByRole("link", { name: /See full report history on website.*outside Microsoft Teams/ }).waitFor();
  await accountMenu(page, "Settings");
  const purchase = page.getByRole("link", { name: /Buy credits on website.*outside Microsoft Teams/ });
  await purchase.waitFor();
  assert.equal(await purchase.getAttribute("target"), "_blank");
  assert.match(await purchase.innerText(), /↗/);
  await page.getByRole("link", { name: /Manage billing on Stripe.*outside Microsoft Teams/ }).waitFor();
  assert.match(await page.locator("[data-payment-disclosure]").innerText(), /outside Microsoft Teams.*Stripe/s);
});

for (const options of [{ clientType: "unknown" }, { failContext: true }]) test(`unrecognized or failed host context hides commerce: ${JSON.stringify(options)}`, async t => {
  const { page } = await scenario(t, options);
  await home(page); await accountMenu(page, "Settings");
  await page.getByText("Available report allowance", { exact: true }).waitFor();
  assert.equal(await page.locator('a[href^="https:"], [data-payment-disclosure]').count(), 0);
});

test("unresolved host never flashes commerce, then desktop receives disclosure before agreement", async t => {
  const { page } = await scenario(t, { saved: false, delayContext: true });
  await page.getByRole("heading", { name: "Welcome back" }).waitFor();
  assert.equal(await page.locator('[data-payment-disclosure], a[href^="https:"]').count(), 0);
  await page.waitForFunction(() => Boolean(window.resolveSyntheticHost));
  await page.evaluate(() => window.resolveSyntheticHost());
  await page.getByRole("button", { name: /Create.*account/i }).last().click();
  await page.locator("#email-terms").waitFor();
  await page.locator("[data-payment-disclosure]").waitFor();
  assert.equal(await page.locator("#email-terms").isChecked(), false);
  assert.equal(await page.locator("[data-payment-disclosure]").evaluate(p => Boolean(p.compareDocumentPosition(document.querySelector("#email-terms")) & Node.DOCUMENT_POSITION_FOLLOWING)), true);
});

test("mobile setup and legal pages contain no routes to external purchasing", async t => {
  const { page } = await scenario(t, { clientType: "ios", connected: false });
  await page.getByRole("heading", { name: "Connect the desktop client", exact: true }).waitFor();
  assert.equal(await page.locator('a[href^="https:"]').count(), 0);
  assert.ok(await page.getByRole("button", { name: "Refresh status", exact: true }).last().isVisible());
  for (const route of ["privacy", "terms"]) {
    // The production server serves these SPA routes; Vite only falls back under its base.
    await page.route(`${origin}/tabs/${route}`, async request => request.fulfill({ response: await request.fetch({ url: `${origin}/tabs/home` }) }));
    await page.goto(`${origin}/tabs/${route}`);
    await page.locator(".legalPanel").waitFor();
    assert.equal(await page.locator('a[href^="https:"], [data-payment-disclosure]').count(), 0);
  }
});

test("restored personal session validates, supports Settings and history, and never joins", async t => {
  const { page, state } = await scenario(t);
  await home(page);
  assert.equal(await page.getByRole("heading", { name: "Connect the desktop client", exact: true }).count(), 0);
  assert.ok(!state.calls.some(call => call.path === "User/activation-code"));
  await page.getByRole("button", { name: "Your Meetings", exact: true }).click();
  await page.getByRole("button", { name: "View Dashboard" }).click();
  const meetingDialog = page.getByRole("dialog", { name: "Synthetic retrospective", exact: true });
  await meetingDialog.waitFor();
  assert.equal(await meetingDialog.locator(".teamsMeetingService").count(), 0);
  assert.ok((await page.locator("main").innerText()).includes("Historical meeting"));
  assert.equal(await page.locator("#personal-health-title").count(), 1);
  for (const width of [390, 1100]) {
    await page.setViewportSize({ width, height: width === 390 ? 600 : 900 });
    assert.ok(await meetingDialog.evaluate(dialog => dialog.scrollWidth <= dialog.clientWidth), `${width}px modal overflow`);
    if (width === 390) {
      assert.ok(await meetingDialog.locator(".personalMeetingModalBody").evaluate(body => {
        body.scrollTop = body.scrollHeight;
        return body.scrollHeight > body.clientHeight && body.scrollTop > 0;
      }), "overflowing meeting content scrolls inside the modal");
      assert.ok(await meetingDialog.getByRole("button", { name: "Close meeting dashboard" }).isVisible());
      await meetingDialog.locator(".personalMeetingModalBody").evaluate(body => { body.scrollTop = 0; });
    }
    if (process.env.PERSONAL_TEST_SCREENSHOTS) {
      require("node:fs").mkdirSync(process.env.PERSONAL_TEST_SCREENSHOTS, { recursive: true });
      await page.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, `meeting-modal-${width}.png`) });
    }
  }
  await meetingDialog.getByRole("button", { name: "Analyze", exact: true }).click();
  await page.getByRole("dialog", { name: "Analyze Synthetic Person", exact: true }).waitFor();
  await page.keyboard.press("Escape");
  await page.getByRole("dialog", { name: "Analyze Synthetic Person", exact: true }).waitFor({ state: "detached" });
  assert.equal(await meetingDialog.isVisible(), true);
  await meetingDialog.getByRole("button", { name: "Close meeting dashboard" }).click();
  await meetingDialog.waitFor({ state: "detached" });
  assert.equal(await page.getByRole("button", { name: "View Dashboard" }).evaluate(button => button === document.activeElement), true);
  await page.getByRole("button", { name: "View Dashboard" }).click();
  await meetingDialog.waitFor();
  await page.keyboard.press("Escape");
  await meetingDialog.waitFor({ state: "detached" });
  await accountMenu(page, "Settings");
  await page.getByRole("heading", { name: "Settings", exact: true }).waitFor();
  await accountMenu(page, "Support");
  await page.getByRole("heading", { name: "Contact support" }).waitFor();
  assert.ok(state.calls.some(call => call.path === "auth/me"));
  assert.ok(!state.calls.some(call => /teams-sso|TeamsMeetings\/join/.test(call.path)));
  assert.deepEqual(state.errors, []);
});
test("email sign-in works while SSO is hidden and remains disabled after sign-out and reload", async t => {
  const { page, state } = await scenario(t, { saved: false });
  await page.getByRole("heading", { name: "Welcome back" }).waitFor();
  assert.equal(await page.getByRole("button", { name: "Continue with Microsoft Teams" }).count(), 0);
  assert.equal(await page.locator(".authSeparator").count(), 0);
  await page.locator("#login-email").fill("synthetic@example.invalid");
  await page.locator("#login-password").fill("Synthetic-test-only-123!");
  await page.locator(".authForm button[type=submit]").click();
  await home(page);
  assert.ok(state.calls.some(call => call.path === "User/Login"));
  await accountMenu(page, "Sign out");
  await page.reload();
  await page.getByRole("heading", { name: "Welcome back" }).waitFor();
  await page.waitForFunction(() => window.syntheticSsoTokenRequests !== undefined);
  assert.equal(await page.evaluate(() => window.syntheticSsoTokenRequests), 0);
  assert.equal(state.calls.filter(call => call.path === "User/teams-sso").length, 0);
});
test("meeting sign-in and account creation expose email only without requesting Teams tokens", async t => {
  const { page, state } = await scenario(t, { saved: false, meeting: true });
  await page.getByRole("heading", { name: "Welcome back" }).waitFor();
  assert.equal(await page.getByRole("button", { name: "Continue with Microsoft Teams" }).count(), 0);
  await page.getByRole("button", { name: /Create.*account/i }).last().click();
  await page.getByRole("heading", { name: /Create/i }).waitFor();
  assert.equal(await page.getByRole("button", { name: "Create with Microsoft Teams" }).count(), 0);
  assert.equal(await page.locator(".teamsCreateBlock, .authSeparator").count(), 0);
  assert.ok(await page.locator("#register-email").isVisible());
  await page.waitForFunction(() => window.syntheticSsoTokenRequests !== undefined);
  assert.equal(await page.evaluate(() => window.syntheticSsoTokenRequests), 0);
  assert.ok(!state.calls.some(call => call.path === "User/teams-sso"));
});
test("activation is explicit, dismissible, reopenable, and disappears on reconnection", async t => {
  const { page, state } = await scenario(t, { connected: false, noSignal: true, emptyMeetings: true });
  await page.getByRole("heading", { name: "Connect the desktop client", exact: true }).waitFor();
  assert.equal(await page.getByRole("heading", { name: "Your Connectivity Health" }).count(), 0);
  assert.equal(await page.getByRole("heading", { name: "Account Information" }).count(), 0);
  await page.getByRole("button", { name: "Generate activation code" }).waitFor();
  assert.ok(!state.calls.some(call => call.path === "User/activation-code" && call.method === "POST"));
  await page.getByRole("button", { name: "Generate activation code" }).click();
  await page.getByText("SYNTHETIC-CODE", { exact: true }).waitFor();
  await page.getByRole("button", { name: "Continue without syncing" }).click();
  await home(page);
  await page.getByRole("heading", { name: "Account Information" }).waitFor();
  await page.getByRole("button", { name: "Your Meetings", exact: true }).click();
  await page.getByText(/No Teams meetings yet/).waitFor();
  await page.getByRole("button", { name: "Connect desktop client" }).click();
  await page.getByRole("button", { name: "Continue without syncing" }).waitFor();
  assert.equal(await page.getByRole("heading", { name: "Your Connectivity Health" }).count(), 0);
  state.connected = true; state.noSignal = false;
  await page.getByRole("button", { name: "Refresh status", exact: true }).first().click();
  await page.getByText("Desktop connected", { exact: true }).waitFor();
  assert.equal(await page.getByRole("button", { name: "Continue without syncing" }).count(), 0);
});
test("a reporting client skips setup even when the separate telemetry read is unavailable", async t => {
  const { page, state } = await scenario(t, { connected: true, noSignal: true });
  await home(page);
  await page.getByText("No data is available.", { exact: true }).waitFor();
  assert.equal(await page.getByRole("heading", { name: "Connect the desktop client", exact: true }).count(), 0);
  assert.equal(await page.getByRole("button", { name: "Connect desktop client", exact: true }).count(), 0);
  assert.ok(!state.calls.some(call => call.path === "User/activation-code"));
});
test("zero-credit Test runs, restores on reload and cancels without enabling paid analysis", async t => {
  const { page, state } = await scenario(t, { credits: 0 });
  await home(page);
  const testButton = page.getByRole("button", { name: "Test Analysis", exact: true });
  await testButton.click();
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  assert.equal(state.request.mode, "Test"); assert.equal(state.request.meetingSessionId, null);
  assert.equal(await page.getByRole("button", { name: "Analysis + Report", exact: true }).isDisabled(), true);
  await page.reload(); await home(page);
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  await page.getByRole("button", { name: "Cancel analysis" }).click();
  await page.getByText("Analysis cancelled", { exact: true }).waitFor();
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 1);
});
test("lost report response survives reload and retries the original ID", async t => {
  const { page, state } = await scenario(t, { loseResponse: true });
  await home(page); await page.getByRole("button", { name: "Analysis + Report", exact: true }).click();
  await page.getByRole("button", { name: "Retry same analysis" }).waitFor();
  await page.reload(); await home(page);
  await page.getByRole("button", { name: "Retry same analysis" }).click();
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  const submitted = state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST");
  assert.equal(submitted.length, 2); assert.deepEqual(submitted[0].body, submitted[1].body);
});
test("session expiry clears private data and history access failure stays in its own section", async t => {
  const { page, state } = await scenario(t, { denyMeeting: true });
  await home(page);
  await page.getByRole("button", { name: "Your Meetings", exact: true }).click();
  await page.getByRole("button", { name: "View Dashboard" }).click();
  await page.getByText("No data is available.", { exact: true }).waitFor();
  await page.getByRole("button", { name: "Close meeting dashboard" }).click();
  state.expired = true;
  await page.getByRole("button", { name: "Refresh status", exact: true }).first().click();
  await page.getByRole("heading", { name: "Welcome back" }).waitFor();
  assert.equal(await page.getByRole("heading", { name: "Account Information" }).count(), 0);
});
test("meeting context continues to join and show the existing meeting dashboard", async t => {
  const { page, state } = await scenario(t, { meeting: true });
  await page.getByRole("heading", { name: "Meeting Overview" }).waitFor();
  await page.getByRole("button", { name: "Invite", exact: true }).waitFor();
  assert.ok(state.calls.some(call => call.path === "TeamsMeetings/join"));
  assert.equal(await page.getByRole("heading", { name: "Your Connectivity Health" }).count(), 0);
  assert.deepEqual(state.errors, []);
});
test("personal dashboard fits narrow light and dark tabs", async t => {
  const { page, state } = await scenario(t);
  await home(page);
  for (const width of [390, 760, 1440]) {
    await page.setViewportSize({ width, height: 1000 });
    for (const theme of ["light", "dark"]) {
      await page.evaluate(theme => { document.documentElement.dataset.teamsTheme = theme === "dark" ? "dark" : "default"; }, theme);
      assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), `${width} ${theme} overflow`);
      if (process.env.PERSONAL_TEST_SCREENSHOTS) {
        require("node:fs").mkdirSync(process.env.PERSONAL_TEST_SCREENSHOTS, { recursive: true });
        await page.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, `personal-${width}-${theme}.png`), fullPage: true });
      }
    }
  }
  assert.deepEqual(state.errors, []);
});

test("temporary restoration failure keeps the session for retry", async t => {
  const { page, state } = await scenario(t, { restorationFailure: true });
  await page.getByRole("heading", { name: "Unable to restore your session" }).waitFor();
  assert.equal(await page.evaluate(() => localStorage.getItem("signaltunerSessionToken")), "synthetic-session");
  state.restorationFailure = false;
  await page.getByRole("button", { name: "Retry", exact: true }).click(); await home(page);
  assert.ok(!state.calls.some(call => call.path === "User/teams-sso"));
});
test("profile-required email sign-in retains profile completion", async t => {
  const { page } = await scenario(t, { saved: false, profileRequired: true });
  await page.getByRole("heading", { name: "Welcome back" }).waitFor();
  await page.locator("#login-email").fill("synthetic@example.invalid");
  await page.locator("#login-password").fill("Synthetic-test-only-123!");
  await page.locator(".authForm button[type=submit]").click();
  await page.locator("#profile-first-name").waitFor();
  assert.equal(await page.getByRole("heading", { name: "Your Connectivity Health" }).count(), 0);
  assert.equal(await page.evaluate(() => localStorage.getItem("signaltunerSessionToken")), null);
});
test("switching accounts removes previous analysis and never replays its pending request", async t => {
  const { page, state } = await scenario(t, { loseResponse: true });
  await home(page); await page.getByRole("button", { name: "Analysis + Report", exact: true }).click();
  await page.getByRole("button", { name: "Retry same analysis" }).waitFor();
  state.userId = 18; state.request = null;
  await page.evaluate(() => {
    localStorage.setItem("signaltunerSessionToken", "synthetic-other");
    window.dispatchEvent(new StorageEvent("storage", { key: "signaltunerSessionToken", newValue: "synthetic-other" }));
  });
  await page.locator(".accountButton").filter({ hasText: "Second Synthetic" }).waitFor();
  assert.equal(await page.getByRole("button", { name: "Retry same analysis" }).count(), 0);
  assert.equal(await page.getByText("Five-minute analysis in progress", { exact: true }).count(), 0);
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 1);
});
test("old desktop capability disables only Test; report completion distinguishes provider acceptance", async t => {
  const { page, state } = await scenario(t, { supportsTest: false });
  await home(page);
  await page.getByText("Update your desktop client to enable Test Analysis.").waitFor();
  assert.equal(await page.getByRole("button", { name: "Test Analysis", exact: true }).isDisabled(), true);
  await page.getByRole("button", { name: "Analysis + Report", exact: true }).click();
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  state.request = { ...state.request, state: "Completed", creditState: "Consumed", deliveries: [{ recipient: "synthetic@example.invalid", state: "Accepted" }] };
  await page.getByRole("button", { name: "Refresh analysis status", exact: true }).click();
  await page.getByText("Report email: Accepted by email provider", { exact: true }).waitFor();
  await page.getByText("Credit status: Consumed", { exact: true }).waitFor();
});
