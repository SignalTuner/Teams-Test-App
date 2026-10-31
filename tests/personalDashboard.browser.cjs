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
      window.syntheticHostContext ||= {app:{appId:"11111111-1111-4111-8111-111111111111",theme:"default",host:{name:"Teams",clientType:${JSON.stringify(options.clientType ?? "desktop")}}},
        page:{id:"home",frameContext:${JSON.stringify(options.frameContext ?? "content")},subPageId:${JSON.stringify(options.subPageId ?? "")}},
        user:{id:"synthetic-oid",tenant:{id:"synthetic-tenant"}}, ${options.meeting ? 'meeting:{id:"synthetic-meeting"},chat:{id:"synthetic-chat"},' : ''}};
    return structuredClone(window.syntheticHostContext); }, openLink: async link => { if (window.syntheticNavigationFailure) throw new Error("Synthetic navigation failure"); (window.syntheticLinks ||= []).push(link); }, registerOnThemeChangeHandler: () => {} };
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
    if (path === "remote-analysis/latest") return reply({ request: state.meetingRequests?.[url.searchParams.get("targetUserId")] || state.request });
    if (path === "remote-analysis/teams-meeting" && request.method() === "POST") {
      const body = JSON.parse(request.postData());
      if (state.loseMeetingResponse) { state.loseMeetingResponse = false; return route.abort(); }
      if (state.acceptMeeting) {
        state.meetingRequests ||= {};
        state.meetingBatches ||= {};
        if (state.meetingBatches[body.requestId]) return reply(state.meetingBatches[body.requestId]);
        const results = body.participantUserIds.map(targetUserId => {
          if (state.rateLimitedParticipants?.includes(targetUserId)) return { targetUserId, statusCode: 429, error: "This device has received too many requests. Try again later." };
          if (state.meetingFailures?.includes(targetUserId)) return { targetUserId, statusCode: 409, error: "The desktop client is not connected." };
          const request = state.meetingRequests[targetUserId] = { requestId: require("node:crypto").createHash("md5").update(`${body.requestId}:${targetUserId}`).digest("hex"),
            targetUserId, state: "Pending", creditState: "None", deliveries: [] };
          return { targetUserId, request };
        });
        return reply(state.meetingBatches[body.requestId] = { batchId: body.requestId, meetingSessionId: body.meetingSessionId, results });
      }
      return reply({ batchId: body.requestId, meetingSessionId: body.meetingSessionId,
        results: body.participantUserIds.map(targetUserId => ({ targetUserId, statusCode: 409, error: "The desktop client is not connected." })) });
    }
    if (path === "remote-analysis" && request.method() === "POST") {
      const body = JSON.parse(request.postData());
      state.request = { ...body, state: state.singleRequestState || "Running", creditState: body.mode === "Test" ? "None" : "Reserved", deliveries: [], liveTelemetry: state.singleRequestState ? null : { ageSeconds: 0, snapshot: { elapsedSeconds: 15, cpuPercent: [10, 12] } } };
      if (state.loseResponse) { state.loseResponse = false; return route.abort(); }
      return reply(state.request);
    }
    if (path.startsWith("remote-analysis/")) {
      const meetingRequest = Object.values(state.meetingRequests || {}).find(value => path === `remote-analysis/${value.requestId}`);
      if (meetingRequest) return reply(meetingRequest);
      if (path.endsWith("/cancel")) state.request = { ...state.request, state: "Cancelled", creditState: "Released" };
      return reply(state.request || {});
    }
    const meeting = { meetingId: "teams-4", meetingSessionId: 4, source: "Teams", name: "Synthetic retrospective", createdAt: "2026-01-01T12:00:00Z", participantCount: 1 };
    if (path === "meetings") return reply({ meetings: state.emptyMeetings ? [] : [meeting] });
    if (path === "meetings/teams-4") return state.denyMeeting ? reply({}, 404) : reply({ meeting, canAnalyze: true, participants: [{ participantId: "teams-user-17", userId: 17, name: "Synthetic Person", email: user.email, meetingRole: "Presenter", signal: { signalScore: 94, signalOverallStatus: "Healthy", signalWifiStatus: "Healthy" } }], requests: [] });
    if (path === "TeamsMeetings/join" || path === "TeamsMeetings/4/dashboard") return reply({ meetingSessionId: 4, currentUser: user, teamsServiceHealth: { activeIncidents: [] }, participants: state.participants || [] });
    return reply({}, 404);
  });
  await page.goto(`${origin}/tabs/home`);
  return { page, state };
}
const home = page => page.getByRole("heading", { name: "Your Connectivity Health" }).waitFor();
async function startPersonalReport(page) {
  await page.getByRole("button", { name: "Analysis + Report", exact: true }).click();
  await page.getByRole("dialog", { name: "Analyze Synthetic Person", exact: true }).getByRole("button", { name: /Request analysis/ }).click();
}
const meetingParticipants = [
  { userId: 17, displayName: "Synthetic Person", email: "synthetic@example.invalid", clientDataStatus: "active", signalScore: 94, overallStatus: "Healthy", meetingRole: "Organizer" },
  { userId: 18, displayName: "Offline Person", email: "offline@example.invalid", clientDataStatus: "offline", signalScore: 70, meetingRole: "Attendee" },
];

for (const credits of [0, null]) test(`personal report requires confirmation and keeps idle actions available with credits ${credits}`, async t => {
  const { page, state } = await scenario(t, { credits });
  await home(page);
  assert.equal(await page.getByRole("button", { name: /^Refresh (analysis )?status$/ }).count(), 0);
  const report = page.getByRole("button", { name: "Analysis + Report", exact: true });
  const testButton = page.getByRole("button", { name: "Test Analysis", exact: true });
  assert.equal(await report.isEnabled(), true); assert.equal(await testButton.isEnabled(), true);
  await report.click();
  const dialog = page.getByRole("dialog", { name: "Analyze Synthetic Person", exact: true });
  await dialog.getByText(`Available credits: ${credits === null ? "Unavailable" : credits}`, { exact: true }).waitFor();
  for (const text of ["Connectivity Health Analysis runs for 5 minutes, starting when the user grants permission on the SignalTuner desktop client.",
    "Upon completion, a Connectivity Health Report will be emailed to you and the analyzed user.", "Each Connectivity Health Report requires 1 credit.",
    "Analysis disrupted by permission denial, cancellation, or disconnection will not deduct credits."]) await dialog.getByText(text, { exact: true }).waitFor();
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 0);
  await dialog.getByRole("button", { name: "Close analysis request" }).click();
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 0);
  await page.route("**/api/remote-analysis", route => route.fulfill({ status: 402, contentType: "application/json", body: "{}" }));
  await startPersonalReport(page);
  await dialog.getByRole("alert").getByText(/No report allowance is available/).waitFor();
  if (process.env.PERSONAL_TEST_SCREENSHOTS) {
    require("node:fs").mkdirSync(process.env.PERSONAL_TEST_SCREENSHOTS, { recursive: true });
    await dialog.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, `personal-report-confirmation-${credits}.png`) });
  }
  await dialog.getByRole("button", { name: "Close analysis request" }).click();
  assert.equal(await report.isEnabled(), true); assert.equal(await testButton.isEnabled(), true);
  await page.unroute("**/api/remote-analysis");
  await testButton.click();
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  assert.equal(await report.isEnabled(), false); assert.equal(await testButton.isEnabled(), false);
  await page.getByRole("button", { name: "Cancel analysis", exact: true }).click();
  await page.getByText("Canceled", { exact: true }).waitFor();
  assert.equal(await report.isEnabled(), true); assert.equal(await testButton.isEnabled(), true);
  assert.equal(await page.getByText(/Credit status:/).count(), 0);
  assert.deepEqual(state.errors, []);
});

test("personal idle actions remain enabled when offline or status cannot be checked", async t => {
  const { page, state } = await scenario(t, { connected: false });
  await page.getByRole("button", { name: "Continue without syncing" }).click();
  await home(page);
  const testButton = page.getByRole("button", { name: "Test Analysis", exact: true });
  const report = page.getByRole("button", { name: "Analysis + Report", exact: true });
  assert.equal(await report.isEnabled(), true); assert.equal(await testButton.isEnabled(), true);
  await testButton.click();
  await page.getByRole("alert").getByText("Connect your desktop client to start an analysis.", { exact: true }).waitFor();
  assert.equal(await testButton.isEnabled(), true);
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 0);
  state.connected = true;
  await page.route("**/api/remote-analysis/latest?**", route => route.fulfill({ status: 503, contentType: "application/json", body: "{}" }));
  await testButton.click();
  await page.getByRole("alert").getByText(/request could not be confirmed/i).waitFor();
  assert.equal(await report.isEnabled(), true); assert.equal(await testButton.isEnabled(), true);
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 0);
  await page.unroute("**/api/remote-analysis/latest?**");
  await testButton.click();
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 1);
  assert.deepEqual(state.errors, []);
});

for (const partial of [false, true]) test(`meeting request limits list only affected participants${partial ? " alongside accepted analyses" : ""}`, async t => {
  const participants = [...meetingParticipants,
    { userId: 19, displayName: "Another Person", email: "another@example.invalid", clientDataStatus: "active", signalScore: 88 },
    { userId: 20, displayName: "Third Person", email: "third@example.invalid", clientDataStatus: "active", signalScore: 90 }];
  const { page, state } = await scenario(t, { meeting: true, participants, acceptMeeting: true, rateLimitedParticipants: partial ? [19, 20] : [17, 19, 20] });
  await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Analyze Meeting", exact: true });
  await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).click();
  const alert = dialog.getByRole("alert");
  await alert.getByText("The following users cannot be analyzed due to too many requests. Deselect them from the list above or try again later.", { exact: false }).waitFor();
  assert.deepEqual(await alert.getByRole("listitem").allTextContents(), partial ? ["another@example.invalid", "third@example.invalid"] : ["synthetic@example.invalid", "another@example.invalid", "third@example.invalid"]);
  assert.equal(await dialog.getByText(/Check that the selected desktops are connected/).count(), 0);
  const self = dialog.getByRole("checkbox", { name: /synthetic@example.invalid/ });
  if (partial) {
    assert.equal(await self.isChecked(), false);
    assert.equal(await self.isEnabled(), false);
    await dialog.getByText("Analysis requested", { exact: true }).waitFor();
  } else await self.uncheck();
  await dialog.getByRole("checkbox", { name: "third@example.invalid", exact: true }).uncheck();
  await dialog.getByText("Credits required: 1", { exact: true }).waitFor();
  if (process.env.PERSONAL_TEST_SCREENSHOTS) {
    require("node:fs").mkdirSync(process.env.PERSONAL_TEST_SCREENSHOTS, { recursive: true });
    await page.evaluate(() => { document.documentElement.dataset.teamsTheme = "dark"; });
    await dialog.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, `meeting-request-limit-${partial}.png`) });
  }
  state.rateLimitedParticipants = [];
  await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).click();
  await dialog.waitFor({ state: "detached" });
  const calls = state.calls.filter(call => call.path === "remote-analysis/teams-meeting");
  assert.equal(calls.length, 2);
  assert.deepEqual(calls[1].body.participantUserIds, [19]);
  assert.notEqual(calls[1].body.requestId, calls[0].body.requestId);
  assert.deepEqual(state.errors, []);
});

test("a whole meeting request limit does not blame individual participants", async t => {
  const { page } = await scenario(t, { meeting: true, participants: meetingParticipants });
  await page.route("**/api/remote-analysis/teams-meeting", route => route.fulfill({ status: 429, contentType: "application/json", body: "{}" }));
  await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Analyze Meeting", exact: true });
  await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).click();
  await dialog.getByRole("alert").getByText("Please wait before retrying this meeting analysis.", { exact: true }).waitFor();
  assert.equal(await dialog.getByRole("alert").getByRole("listitem").count(), 0);
});

for (const action of ["Test Analysis", "Analysis + Report", "meeting-individual"]) test(`${action} shows the same device request-limit message`, async t => {
  const { page } = await scenario(t, { meeting: action === "meeting-individual", participants: meetingParticipants });
  const message = "This device has received too many requests. Try again later.";
  await page.route("**/api/remote-analysis", route => route.fulfill({ status: 429, contentType: "application/json", body: JSON.stringify({ message }) }));
  if (action === "meeting-individual") {
    await page.getByRole("button", { name: "Analyze", exact: true }).first().click();
    await page.getByRole("button", { name: /Request analysis/ }).click();
  } else if (action === "Analysis + Report") await startPersonalReport(page);
  else await page.getByRole("button", { name: action, exact: true }).click();
  await page.getByRole("alert").getByText(message, { exact: true }).waitFor();
  assert.equal(await page.getByText("The request was declined. Refresh status and check desktop availability and report allowance.", { exact: true }).count(), 0);
});

for (const frameContext of ["content", "sidePanel"]) test(`individual analysis hands off approval and running status in ${frameContext}`, async t => {
  const { page, state } = await scenario(t, { meeting: true, frameContext, participants: meetingParticipants, singleRequestState: "Pending" });
  const participant = page.locator(".participantListItem").first();
  await participant.getByRole("button", { name: "Analyze", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Analyze Synthetic Person", exact: true });
  await dialog.getByRole("button", { name: /Request analysis/ }).click();
  await dialog.waitFor({ state: "detached" });
  const status = participant.locator(".remoteAnalysisProgress");
  await status.getByText("Waiting for approval...", { exact: true }).waitFor();
  assert.equal(await status.evaluate(el => getComputedStyle(el).color), "rgb(250, 188, 44)");
  assert.equal(await participant.getByRole("timer").count(), 0);
  if (frameContext === "content") {
    assert.equal(await participant.locator(".remoteAnalysisInlineSummary").evaluate(el => el.previousElementSibling.textContent.trim()), "Analysis for Synthetic Person");
    await participant.getByRole("button", { name: "Analysis for Synthetic Person", exact: true }).click();
    assert.equal(await status.isVisible(), true, "status stays visible when charts are collapsed");
  } else await participant.getByText("Analysis for Synthetic Person", { exact: true }).waitFor();
  state.request = { ...state.request, state: "Approved" };
  await page.waitForResponse(response => response.url().includes(state.request.requestId));
  await status.getByText("Waiting for approval...", { exact: true }).waitFor();
  // Running must display immediately, even before metric arrays arrive.
  state.request = { ...state.request, state: "Running", liveTelemetry: { ageSeconds: 0, snapshot: { elapsedSeconds: 30 } } };
  await status.getByText("Five-minute analysis in progress", { exact: false }).waitFor();
  await status.getByRole("timer").waitFor();
  assert.match(await status.getByRole("timer").textContent(), /^4:[23]\d remaining$/);
  assert.equal(await status.evaluate(el => getComputedStyle(el).color), "rgb(59, 181, 55)");
  assert.equal(await participant.getByText("Waiting for approval...", { exact: true }).count(), 0);
  assert.equal(await page.getByRole("dialog").count(), 0);
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 1);
  assert.deepEqual(state.errors, []);
});

for (const partial of [false, true]) test(`meeting analysis closes and tracks each accepted participant${partial ? " with partial failure" : ""}`, async t => {
  const participants = [...meetingParticipants,
    { userId: 19, displayName: "Another Person", email: "another@example.invalid", clientDataStatus: "active", signalScore: 88 }];
  const { page, state } = await scenario(t, { meeting: true, participants, acceptMeeting: true, meetingFailures: partial ? [19] : [] });
  await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Analyze Meeting", exact: true });
  await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).click();
  await dialog.waitFor({ state: "detached" });
  const self = page.getByRole("region", { name: "Analysis for Synthetic Person", exact: true });
  const other = page.getByRole("region", { name: "Analysis for Another Person", exact: true });
  await self.getByText("Waiting for approval...", { exact: true }).waitFor();
  if (partial) {
    await page.getByText("1 of 2 requests accepted.", { exact: false }).waitFor();
    assert.equal(await other.getByText("Waiting for approval...", { exact: true }).count(), 0);
    await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
    await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).waitFor();
    assert.equal(await dialog.getByRole("checkbox", { name: "another@example.invalid", exact: true }).isEnabled(), true);
    await dialog.getByRole("button", { name: "Close meeting analysis", exact: true }).click();
  } else await other.getByText("Waiting for approval...", { exact: true }).waitFor();
  state.meetingRequests[17] = { ...state.meetingRequests[17], state: "Running", creditState: "Reserved",
    liveTelemetry: { ageSeconds: 0, snapshot: { elapsedSeconds: 45 } } };
  await self.getByText("Five-minute analysis in progress", { exact: false }).waitFor();
  await self.getByRole("timer").waitFor();
  if (!partial) await other.getByText("Waiting for approval...", { exact: true }).waitFor();
  assert.equal(await self.getByText("Waiting for approval...", { exact: true }).count(), 0);
  assert.equal(state.calls.filter(call => call.path === "remote-analysis/teams-meeting").length, 1);
  if (process.env.PERSONAL_TEST_SCREENSHOTS) {
    require("node:fs").mkdirSync(process.env.PERSONAL_TEST_SCREENSHOTS, { recursive: true });
    await page.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, `meeting-analysis-status-${partial}.png`), fullPage: true });
  }
  assert.deepEqual(state.errors, []);
});

for (const terminal of ["Cancelled", "Completed"]) test(`individual confirmation keeps the start form after ${terminal}`, async t => {
  const { page, state } = await scenario(t, { meeting: true, participants: meetingParticipants, singleRequestState: "Pending" });
  const participant = page.locator(".participantListItem").first();
  const open = () => participant.getByRole("button", { name: "Analyze", exact: true }).click();
  await open();
  const dialog = page.getByRole("dialog", { name: "Analyze Synthetic Person", exact: true });
  await dialog.getByRole("button", { name: /Request analysis/ }).click();
  await dialog.waitFor({ state: "detached" });
  state.request = { ...state.request, state: terminal, creditState: terminal === "Completed" ? "Consumed" : "Released" };
  await open();
  const start = dialog.getByRole("button", { name: /Request analysis/ });
  await start.waitFor();
  assert.equal(await start.isEnabled(), true);
  assert.equal(await dialog.getByRole("button", { name: /New.*analysis|Retry.*analysis/i }).count(), 0);
  assert.equal(await dialog.getByText(/PDF generated|reserved credit was returned|One report credit used/).count(), 0);
  await start.click();
  await dialog.waitFor({ state: "detached" });
  const calls = state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST");
  assert.equal(calls.length, 2);
  assert.notEqual(calls[0].body.requestId, calls[1].body.requestId);
  // Reopening during a live run keeps the same form but prevents a second start.
  await open();
  await start.waitFor();
  assert.equal(await start.isEnabled(), false);
  assert.equal(await dialog.getByRole("button", { name: /New.*analysis|Retry.*analysis/i }).count(), 0);
  assert.deepEqual(state.errors, []);
});

for (const terminal of ["Cancelled", "Completed"]) test(`meeting confirmation allows a fresh selection after ${terminal} and reload`, async t => {
  const participants = [...meetingParticipants,
    { userId: 19, displayName: "Another Person", email: "another@example.invalid", clientDataStatus: "active", signalScore: 88 }];
  const { page, state } = await scenario(t, { meeting: true, participants, acceptMeeting: true });
  await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Analyze Meeting", exact: true });
  await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).click();
  await dialog.waitFor({ state: "detached" });
  for (const request of Object.values(state.meetingRequests)) request.state = terminal;
  assert.equal(await page.evaluate(() => sessionStorage.getItem("teamsMeetingAnalysis:17:4")), null);
  await page.reload();
  await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
  await dialog.getByRole("checkbox", { name: "synthetic@example.invalid", exact: true }).uncheck();
  await dialog.getByText("Credits required: 1", { exact: true }).waitFor();
  assert.equal(await dialog.getByRole("button", { name: /New.*analysis|Retry.*analysis/i }).count(), 0);
  assert.equal(await dialog.getByRole("list", { name: "Meeting analysis results" }).count(), 0);
  await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).click();
  await dialog.waitFor({ state: "detached" });
  const calls = state.calls.filter(call => call.path === "remote-analysis/teams-meeting");
  assert.equal(calls.length, 2);
  assert.notEqual(calls[0].body.requestId, calls[1].body.requestId);
  assert.deepEqual(calls[1].body.participantUserIds, [19]);
  assert.deepEqual(state.errors, []);
});

test("unconfirmed individual analysis stays actionable without claiming approval is pending", async t => {
  const { page, state } = await scenario(t, { meeting: true, participants: meetingParticipants, singleRequestState: "Pending", loseResponse: true });
  await page.locator(".participantListItem").first().getByRole("button", { name: "Analyze", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Analyze Synthetic Person", exact: true });
  await dialog.getByRole("button", { name: /Request analysis/ }).click();
  await dialog.getByRole("alert").waitFor();
  assert.equal(await page.getByText("Waiting for approval...", { exact: true }).count(), 0);
  assert.equal(await dialog.isVisible(), true);
  await dialog.getByRole("button", { name: /Request analysis/ }).click();
  await dialog.waitFor({ state: "detached" });
  const calls = state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST");
  assert.equal(calls.length, 2);
  assert.equal(calls[0].body.requestId, calls[1].body.requestId);
});

for (const width of [320, 1100]) test(`meeting confirmation selects participants and preserves the submitted roster on retry at ${width}px`, async t => {
  const participants = [...meetingParticipants,
    { userId: 19, displayName: "Another Person", email: "another@example.invalid", clientDataStatus: "active", signalScore: 88 }];
  const { page, state } = await scenario(t, { meeting: true, width, participants, loseMeetingResponse: true });
  await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Analyze Meeting", exact: true });
  await dialog.getByText("Available credits: 2", { exact: true }).waitFor();
  for (const copy of [
    "Connectivity Health Analysis runs for 5 minutes, starting when the user grants permission on the SignalTuner desktop client.",
    "Upon completion, a Connectivity Health Report will be emailed to you and the analyzed user.",
    "Each Connectivity Health Report requires 1 credit.",
    "Analysis disrupted by permission denial, cancellation, or disconnection will not deduct credits.",
  ]) await dialog.getByText(copy, { exact: true }).waitFor();
  const self = dialog.getByRole("checkbox", { name: "synthetic@example.invalid", exact: true });
  const other = dialog.getByRole("checkbox", { name: "another@example.invalid", exact: true });
  assert.equal(await self.isChecked(), true);
  assert.equal(await other.isChecked(), true);
  const offline = dialog.getByRole("checkbox", { name: /offline@example.invalid/ });
  assert.equal(await offline.isEnabled(), false);
  assert.equal(await offline.isChecked(), false);
  await dialog.getByText("Credits required: 2", { exact: true }).waitFor();
  await self.uncheck();
  await dialog.getByText("Credits required: 1", { exact: true }).waitFor();
  await other.uncheck();
  await dialog.getByText("Credits required: 0", { exact: true }).waitFor();
  assert.equal(await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).isEnabled(), false);
  await other.check();
  assert.equal(await dialog.evaluate(el => el.scrollWidth <= el.clientWidth), true);
  assert.ok(!state.calls.some(call => call.path === "remote-analysis/teams-meeting"));
  if (process.env.PERSONAL_TEST_SCREENSHOTS) {
    require("node:fs").mkdirSync(process.env.PERSONAL_TEST_SCREENSHOTS, { recursive: true });
    await page.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, `meeting-confirmation-${width}.png`), fullPage: true });
  }
  await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).click();
  await dialog.getByRole("alert").waitFor();
  const submitted = () => state.calls.filter(call => call.path === "remote-analysis/teams-meeting");
  assert.equal(submitted().length, 1);
  assert.deepEqual(submitted()[0].body.participantUserIds, [19]);
  assert.equal(submitted()[0].body.meetingSessionId, 4);
  assert.equal(await other.isEnabled(), false);
  // Reload must recover the same selected roster and ID, without auto-submitting.
  await page.reload();
  await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
  await dialog.getByText("Credits required: 1", { exact: true }).waitFor();
  assert.equal(await self.isChecked(), false);
  assert.equal(await other.isChecked(), true);
  assert.equal(await other.isEnabled(), false);
  assert.equal(submitted().length, 1);
  await dialog.getByRole("button", { name: "Start meeting analysis", exact: true }).click();
  await dialog.getByRole("alert").filter({ hasText: "No participant analyses could be started" }).waitFor();
  assert.equal(submitted().length, 2);
  assert.deepEqual(submitted()[1].body, submitted()[0].body);
  await dialog.getByRole("button", { name: "Close meeting analysis", exact: true }).click();
  await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
  await dialog.getByText("Credits required: 2", { exact: true }).waitFor();
  assert.equal(await self.isEnabled(), true);
  assert.equal(await other.isEnabled(), true);
  assert.equal(await dialog.getByRole("button", { name: /New.*analysis|Retry.*analysis/i }).count(), 0);
  assert.deepEqual(state.errors, []);
});

for (const width of [320, 1100]) test(`meeting side panel stays compact at ${width}px and opens confirmation in the full tab`, async t => {
  const { page, state } = await scenario(t, { meeting: true, frameContext: "sidePanel", width, participants: meetingParticipants });
  await page.getByRole("heading", { name: "Meeting Overview" }).waitFor();
  assert.equal(await page.locator(".meetingSidePanel").count(), 1);
  assert.equal(await page.locator(".meetingHealthMetrics, .participantRole, .participantAvatar, .participantMetrics, .remoteAnalysisInline").count(), 0);
  assert.equal(await page.getByRole("button", { name: "Invite", exact: true }).count(), 0);
  assert.equal(await page.locator(".sidePanelParticipant").count(), 2);
  assert.equal(await page.locator(".sidePanelParticipant").first().getByRole("button", { name: "Analyze", exact: true }).count(), 1);
  await page.getByText("Available credits:", { exact: false }).waitFor();
  await page.getByRole("button", { name: /View incidents/ }).waitFor();
  const openTab = page.getByRole("button", { name: "Open in Teams tab", exact: true });
  assert.equal(await page.evaluate(() => Boolean(document.querySelector(".teamsMeetingService").compareDocumentPosition(document.querySelector(".sidePanelOpenTab")) & Node.DOCUMENT_POSITION_FOLLOWING)), true);
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  if (process.env.PERSONAL_TEST_SCREENSHOTS) {
    require("node:fs").mkdirSync(process.env.PERSONAL_TEST_SCREENSHOTS, { recursive: true });
    await page.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, `meeting-side-panel-${width}.png`), fullPage: true });
  }
  await page.getByRole("button", { name: "Analyze Meeting", exact: true }).click();
  const link = new URL(await page.evaluate(() => window.syntheticLinks.at(-1)));
  assert.equal(link.searchParams.get("openInMeeting"), "false");
  const target = JSON.parse(link.searchParams.get("context"));
  assert.equal(target.chatId, "synthetic-chat");
  assert.match(target.subEntityId, /^confirm-meeting-analysis:/);
  assert.equal(await page.getByRole("dialog").count(), 0);
  assert.ok(!state.calls.some(call => call.path.startsWith("remote-analysis") && call.method === "POST"));
  await openTab.click();
  assert.equal(JSON.parse(new URL(await page.evaluate(() => window.syntheticLinks.at(-1))).searchParams.get("context")).subEntityId, undefined);
  await page.evaluate(() => { window.syntheticNavigationFailure = true; });
  await openTab.click();
  await page.getByRole("alert").filter({ hasText: "The meeting tab could not be opened" }).waitFor();
  // Participant analysis retains the existing confirmation workflow in the panel.
  await page.locator(".sidePanelParticipant").first().getByRole("button", { name: "Analyze", exact: true }).click();
  const individualDialog = page.getByRole("dialog", { name: "Analyze Synthetic Person", exact: true });
  await individualDialog.waitFor();
  await individualDialog.getByText("Available credits: 2", { exact: true }).waitFor();
  for (const copy of [
    "Connectivity Health Analysis runs for 5 minutes, starting when the user grants permission on the SignalTuner desktop client.",
    "Upon completion, a Connectivity Health Report will be emailed to you and the analyzed user.",
    "Each Connectivity Health Report requires 1 credit.",
    "Analysis disrupted by permission denial, cancellation, or disconnection will not deduct credits.",
  ]) await individualDialog.getByText(copy, { exact: true }).waitFor();
  assert.equal(await individualDialog.evaluate(el => el.scrollWidth <= el.clientWidth), true);
  if (process.env.PERSONAL_TEST_SCREENSHOTS) {
    await page.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, `individual-confirmation-${width}.png`) });
  }
  state.credits = 0;
  await page.evaluate(() => window.dispatchEvent(new Event("focus")));
  await individualDialog.getByText("Available credits: 0", { exact: true }).waitFor();
  await page.route("**/api/credits/account", route => route.fulfill({ status: 503, contentType: "application/json", body: "{}" }));
  await page.evaluate(() => window.dispatchEvent(new Event("focus")));
  await individualDialog.getByText("Available credits: Unavailable", { exact: true }).waitFor();
  assert.ok(!state.calls.some(call => call.path.startsWith("remote-analysis") && call.method === "POST"));
  assert.deepEqual(state.errors, []);

  const destination = await scenario(t, { meeting: true, subPageId: target.subEntityId, width, participants: meetingParticipants });
  await destination.page.getByRole("dialog", { name: "Analyze Meeting", exact: true }).waitFor();
  assert.equal(await destination.page.locator(".meetingHealthMetrics").count(), 1);
  assert.equal(await destination.page.locator(".meetingSidePanel").count(), 0);
  await destination.page.getByRole("button", { name: "Start meeting analysis", exact: true }).waitFor();
  assert.ok(!destination.state.calls.some(call => call.path.startsWith("remote-analysis") && call.method === "POST"));
  await destination.page.getByRole("button", { name: "Close meeting analysis", exact: true }).click();
  await destination.page.evaluate(() => window.dispatchEvent(new Event("focus")));
  assert.equal(await destination.page.getByRole("dialog").count(), 0);
});

test("a cached full meeting tab opens a new confirmation hint on focus without starting analysis", async t => {
  const { page, state } = await scenario(t, { meeting: true, participants: meetingParticipants });
  await page.getByRole("heading", { name: "Meeting Overview" }).waitFor();
  assert.equal(await page.getByRole("dialog").count(), 0);
  await accountMenu(page, "Settings");
  await page.evaluate(() => {
    window.syntheticHostContext.page.subPageId = "confirm-meeting-analysis:11111111-1111-4111-8111-111111111111";
    window.dispatchEvent(new Event("focus"));
  });
  await page.getByRole("dialog", { name: "Analyze Meeting", exact: true }).waitFor();
  assert.ok(!state.calls.some(call => call.path.startsWith("remote-analysis") && call.method === "POST"));
  await page.getByRole("button", { name: "Close meeting analysis", exact: true }).click();
  await accountMenu(page, "Settings");
  await accountMenu(page, "Dashboard");
  await page.getByRole("heading", { name: "Meeting Overview" }).waitFor();
  assert.equal(await page.getByRole("dialog").count(), 0);
  await page.evaluate(() => {
    window.syntheticHostContext = { ...window.syntheticHostContext, chat: { id: "another-chat" },
      page: { ...window.syntheticHostContext.page, subPageId: "confirm-meeting-analysis:22222222-2222-4222-8222-222222222222" } };
    window.dispatchEvent(new Event("focus"));
  });
  assert.equal(await page.getByRole("dialog").count(), 0);
  assert.deepEqual(state.errors, []);
});

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
  await page.getByText("Available credits", { exact: true }).waitFor();
  assert.match(await page.locator(".settings-credit-value").innerText(), /8 reports/);
  assert.equal(await page.locator('a[href^="https:"], .settings-plan-summary, [data-payment-disclosure]').count(), 0);
  assert.doesNotMatch(await page.locator(".settings-account").innerText(), /buy|purchase|upgrade|pricing|trial|payment|billing|pay as you go/i);
  await page.getByRole("button", { name: "Contact support about account deletion" }).click();
  await page.getByRole("heading", { name: "Contact support" }).waitFor();
  await page.goto(`${origin}/tabs/home`); await home(page);
  await startPersonalReport(page);
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  assert.equal(state.request.mode, "Report");
  assert.deepEqual(state.errors, []);
});

test("mobile zero allowance remains neutral and Test Analysis still runs", async t => {
  const { page, state } = await scenario(t, { clientType: "android", credits: 0 });
  await home(page);
  await page.getByText(/No report allowance is available/).waitFor();
  assert.equal(await page.getByRole("button", { name: "Analysis + Report", exact: true }).isDisabled(), false);
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
  await page.getByText("Available credits", { exact: true }).waitFor();
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
  const analysis = page.getByRole("region", { name: "Analysis for Synthetic Person", exact: true });
  assert.equal(await analysis.getByRole("heading", { name: "Test Analysis", exact: true }).count(), 0);
  await analysis.locator(".remoteAnalysisProgress").getByRole("timer").waitFor();
  assert.equal(await analysis.evaluate(element => {
    const body = element.querySelector(".remoteAnalysisInlineBody");
    return body.firstElementChild.className === "remoteAnalysisProgress"
      && body.querySelector(".live-analysis").nextElementSibling.className === "buttonRow";
  }), true);
  if (process.env.PERSONAL_TEST_SCREENSHOTS) {
    require("node:fs").mkdirSync(process.env.PERSONAL_TEST_SCREENSHOTS, { recursive: true });
    await analysis.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, "personal-analysis-running.png") });
  }
  assert.equal(state.request.mode, "Test"); assert.equal(state.request.meetingSessionId, null);
  assert.equal(await page.getByRole("button", { name: "Analysis + Report", exact: true }).isDisabled(), true);
  await page.reload(); await home(page);
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  await page.getByRole("button", { name: "Cancel analysis" }).click();
  await page.getByText("Canceled", { exact: true }).waitFor();
  assert.equal(await analysis.locator(".remoteAnalysisProgress .analysis-canceled").count(), 1);
  assert.equal(await analysis.getByRole("timer").count(), 0);
  if (process.env.PERSONAL_TEST_SCREENSHOTS) {
    await analysis.evaluate(element => window.scrollTo(0, window.scrollY + element.getBoundingClientRect().top - 100));
    await analysis.screenshot({ path: require("node:path").join(process.env.PERSONAL_TEST_SCREENSHOTS, "personal-analysis-canceled.png") });
  }
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 1);
});
test("lost report response survives reload and the normal report action recovers the original ID", async t => {
  const { page, state } = await scenario(t, { loseResponse: true });
  await home(page); await startPersonalReport(page);
  await page.getByRole("alert").waitFor();
  state.request = null; // Simulate a latest read that cannot yet see the submission.
  await page.reload(); await home(page);
  await startPersonalReport(page);
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
  await page.evaluate(() => window.dispatchEvent(new Event("focus")));
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
  await home(page); await startPersonalReport(page);
  await page.getByRole("alert").waitFor();
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
test("old desktop capability gives feedback on click; report completion distinguishes provider acceptance", async t => {
  const { page, state } = await scenario(t, { supportsTest: false });
  await home(page);
  await page.getByText("Update your desktop client to run Test Analysis.").waitFor();
  assert.equal(await page.getByRole("button", { name: "Test Analysis", exact: true }).isDisabled(), false);
  await page.getByRole("button", { name: "Test Analysis", exact: true }).click();
  await page.getByRole("alert").getByText("Update your desktop client to run Test Analysis.", { exact: true }).waitFor();
  assert.equal(state.calls.filter(call => call.path === "remote-analysis" && call.method === "POST").length, 0);
  await startPersonalReport(page);
  await page.getByText("Five-minute analysis in progress", { exact: true }).waitFor();
  state.request = { ...state.request, state: "Completed", creditState: "Consumed", deliveries: [{ recipient: "synthetic@example.invalid", state: "Accepted" }] };
  await page.getByText("Report email: Accepted by email provider", { exact: true }).waitFor();
  assert.equal(await page.getByText(/Credit status:/).count(), 0);
  assert.equal(await page.getByRole("button", { name: "Analysis + Report", exact: true }).isDisabled(), false);
});
