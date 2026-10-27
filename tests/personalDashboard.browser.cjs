// Run against localhost Vite with VITE_SIGNALTUNER_API_URL=http://127.0.0.1:4317/backend.
// API and Teams SDK responses are synthetic and intercepted only in this test browser.
const { test } = require("node:test");
const assert = require("node:assert/strict");
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "playwright");
const origin = process.env.PERSONAL_TEST_ORIGIN || "http://127.0.0.1:4317";
async function scenario(t, options = {}) {
  const browser = await chromium.launch({ headless: true, timeout: 15000, ...(process.env.BROWSER_EXECUTABLE ? { executablePath: process.env.BROWSER_EXECUTABLE } : {}) });
  t.after(() => browser.close());
  const context = await browser.newContext({ viewport: { width: 1100, height: 900 } });
  t.after(() => context.close());
  const page = await context.newPage();
  page.setDefaultTimeout(10000);
  const state = { connected: options.connected ?? true, credits: options.credits ?? 2, request: null, calls: [], errors: [], ...options };
  page.on("pageerror", error => state.errors.push(error.message));
  await page.addInitScript(({ saved }) => { if (saved && !localStorage.getItem("initialized")) { localStorage.setItem("signaltunerSessionToken", "synthetic-session"); localStorage.setItem("initialized", "yes"); } }, { saved: options.saved !== false });
  await page.route("**/@microsoft_teams-js.js*", route => route.fulfill({ contentType: "application/javascript", body: `
    export const app = { initialize: async () => {}, getContext: async () => ({app:{theme:"default"},user:{id:"synthetic-oid",tenant:{id:"synthetic-tenant"}}, ${options.meeting ? 'meeting:{id:"synthetic-meeting"},' : ''}}), registerOnThemeChangeHandler: () => {} };
    export const authentication = { getAuthToken: async () => "synthetic-teams-token" };
    export const meeting = { getMeetingDetails: callback => callback(null, null) };
    export const pages = {};` }));
  await page.route("**/api/**", async route => {
    const request = route.request(), url = new URL(request.url());
    const path = url.pathname.split("/api/")[1];
    state.calls.push({ path, method: request.method(), body: request.postData() ? JSON.parse(request.postData()) : null });
    const user = { userId: state.userId || 17, displayName: state.userId === 18 ? "Second Synthetic" : "Synthetic Person", firstName: "Synthetic", lastName: "Person", email: "synthetic@example.invalid", authProvider: "teams_sso", clientIsActive: state.connected, credits: state.credits, hasPassword: true };
    const reply = (value, status = 200) => route.fulfill({ status, contentType: "application/json", body: JSON.stringify(value) });
    if (path === "auth/me") return reply(user, state.expired ? 401 : state.restorationFailure ? 503 : 200);
    if (path === "User/teams-sso") return state.ssoFailure ? reply({ message: "Link your existing account before using Teams sign-in." }, 409)
      : reply({ token: "synthetic-session", user, profileRequired: state.profileRequired || false });
    if (path === "User/signal/current") return state.noSignal ? reply({}, 404) : reply({ SignalScore: 94, SignalOverallStatus: state.connected ? "Healthy" : "Offline", SignalWifiStatus: "Healthy", SignalBandwidthStatus: "Fair", SignalSystemStatus: "Healthy", LastUpdated: new Date().toISOString() });
    if (path === "credits/account") return reply({ purchasedCredits: state.credits, availableReports: state.credits, subscription: null });
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
test("automatic SSO works without meeting context and explicit sign-out suppresses it after reload", async t => {
  const { page, state } = await scenario(t, { saved: false });
  await home(page);
  const sso = state.calls.find(call => call.path === "User/teams-sso");
  assert.equal(sso.body.teamsMeetingId, null); assert.equal(sso.body.teamsTenantId, null); assert.equal(sso.body.allowAccountCreation, false);
  await accountMenu(page, "Sign out");
  await page.reload();
  await page.getByRole("heading", { name: "Welcome back" }).waitFor();
  assert.equal(state.calls.filter(call => call.path === "User/teams-sso").length, 1);
});
test("unlinked account falls back to existing login and account creation", async t => {
  const { page, state } = await scenario(t, { saved: false, ssoFailure: true });
  await page.getByText("Link your existing account before using Teams sign-in.", { exact: false }).waitFor();
  await page.getByRole("button", { name: /Create.*account/i }).last().click();
  await page.getByRole("heading", { name: /Create/i }).waitFor();
  assert.ok(!state.calls.some(call => call.path === "User/teams-sso" && call.body.allowAccountCreation));
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
test("profile-required SSO stays in the existing profile completion flow", async t => {
  const { page } = await scenario(t, { saved: false, profileRequired: true });
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
