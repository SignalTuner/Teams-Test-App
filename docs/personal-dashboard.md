# Teams personal dashboard

The existing personal tab validates a saved SignalTuner session with `/api/auth/me`
independently of meeting context. As of 2026-10-27, Teams SSO is temporarily paused
by `TEAMS_SSO_ENABLED = false` in `src/Tab/App.tsx`: sign-in/account-creation buttons,
SSO notices and separators are hidden, and automatic and manual token acquisition
are blocked. With no valid session, use email/password sign-in or email account
creation. Password reset, legal consent and profile completion remain available.
Separate storage partitions sign in independently; no tokens pass between frames.
Existing sessions, including those originally issued through SSO, still restore.
The backend and manifest SSO contracts remain intact for a later coordinated rollout.

Explicit sign-out suppresses automatic SSO. Transient restoration failures retain
the token for retry; HTTP 401 clears it. Shared-storage account changes clear
private state before validating the replacement token. Personal requests abort on
unmount and account-keyed components isolate data and pending analyses.

## Shared components

- `DashboardHeader`: existing logo, account dropdown and navigation.
- `ConnectivityPresentation`: existing gauge, trend chart, telemetry groups,
  colors, thresholds, icons and recommendations, extracted from `App.tsx`.
- `ConnectivitySummary`: category/score display; the meeting view retains its
  existing category fallback behavior.
- `ClientPrompt`: existing download, temporary activation code and polling flow.
  Codes are never generated automatically. As in meeting sign-in, setup appears
  as a separate screen only when the authenticated account's `clientIsActive`
  reporting-data flag is false. Reporting clients open the dashboard directly.
  Connecting or choosing Continue without syncing opens the dashboard; disconnected
  users can reopen setup later. A telemetry-read failure does not trigger setup.
- `AnalysisSection`: shared disclosure layout for personal and meeting charts.
- `AnalysisConfirmationDialog`: shared individual-analysis confirmation shell,
  wording, credit balance, error styling and close controls. Personal Analysis +
  Report opens this dialog without submitting; Request analysis confirms it.
- Existing Settings, Support, AppearanceSelector, MeetingServiceStatus,
  LiveAnalysisCharts, AnalysisCountdown and RemoteAnalysisButton are reused.

`PersonalDashboard`, `PersonalAnalysis` and `personalDashboardData` contain personal
view orchestration. Each section loads and fails independently. Visible tabs poll
signals every ten seconds, account/credits/device capability every thirty seconds,
and history every minute. Returning to the tab refreshes immediately. Signals older
than two minutes are stale; missing values never fall back to zero or Healthy.

Personal connectivity uses the same participant identity, score/category row and
disclosure styling as the meeting dashboard. Test Analysis and Analysis + Report
sit at the right, with Available credits underneath. The analysis disclosure sits
directly below that row; there is no separate Analysis card, last-update label or
telemetry-details toggle. Active requests and recovery errors open the disclosure
automatically. This layout change requires only Teams frontend deployment.

The personal analysis disclosure uses `Analysis for <display name>` and the same
compact running status/countdown and Canceled indicator as the meeting view.
Telemetry follows immediately; cancellation and report-delivery controls remain
below the charts. Refresh status, Refresh analysis status and the Credit status
label/value are removed. The redundant Test/Report heading is removed;
Test completion still explicitly says no report was generated or credit used.
Both analysis buttons remain enabled when idle, including with zero/unknown
credits, a disconnected/older desktop, or a previous error. Pending/Approved/
Running analysis and an in-flight action disable them to prevent overlap. Clicking
an action checks the latest analysis and desktop capability and shows actionable
feedback; enabling a button does not bypass backend entitlement enforcement.
Idle analysis status refreshes every thirty seconds and on focus/visibility;
active analysis retains its two-second polling. There is no automatic submission.
Test starts directly. Report requires the same explicit confirmation as in-meeting
individual analysis and closes on acceptance. Opening/closing the dialog alone
never creates a request or spends credits.

The [Teams mobile commerce policy](teams-mobile-commerce.md) hides direct and
indirect purchase routes on iOS/Android/iPadOS using TeamsJS host context. Existing
allowance and entitled reports remain usable; zero allowance does not disable
supported Test Analysis. Desktop/web external links and consent disclose leaving
Teams for website/Stripe payment flows. This follow-up changes only the Teams UI.

## Contracts and recovery

The only new route is **`GET /api/User/signal/current`**. It resolves the authenticated
active account and reuses the website's signal reader. This is necessary because
`/api/auth/me` does not return the legacy code required by `/api/User/GetSignal/{code}`.
It neither generates nor exposes activation codes. Both signal routes select the
latest row and populate the existing `LastUpdated`/`Timestamp` properties as UTC.
Hardware identifiers remain excluded by SignalModel serialization. No schema change.

Other contracts are unchanged: `/api/credits/account`; remote-analysis current-device,
latest with `mode=Any`, creation, live status, cancel and retry-delivery; and
`/api/meetings?includeTeams=true` with `/api/meetings/teams-{id}` details.

Self-analysis submits the authenticated target with `meetingSessionId: null`, mode
Test or Report and a request ID. The exact operation is saved in account-scoped
sessionStorage **before sending**. Storage failure prevents sending; uncertain
responses/reloads recover through the normal action buttons using the same ID/body
if read-only checks cannot find the request. A confirmed existing active request
is restored without another POST; a pending operation cannot switch modes.
Sign-out retains uncertain operation IDs
for safe recovery by the same account, never another account. New analyses require
a successful prior-status check. Backend ownership, conflict, approval, capability
and billing checks remain authoritative. Test is free and capability-checked on
click; Report credit eligibility remains enforced by the backend. Active status polls every two seconds and
pending delivery every ten seconds while visible. Provider acceptance is not
described as confirmed inbox delivery.

Historical meetings show previously signed-in membership and latest available
participant signals, **not** meeting-time telemetry or live attendance. Opening
history opens a closable modal above the personal dashboard and never calls join.
Close or Escape dismisses it and restores focus to the triggering button; personal
analysis stays mounted behind it. Your Meetings starts collapsed on page load.
Participant analysis uses the numeric meeting session ID
and existing approval/credit rules. No calendar, Graph, bot or invitation changes.

## Validation commands

```powershell
./node_modules/.bin/tsc.cmd -p tsconfig.app.json
npm run build
npm run test:personal
node --test tests/meetingAnalysis.test.cjs tests/analysisLive.test.cjs tests/serviceStatus.test.cjs tests/meetingInvitations.test.cjs tests/settingsAccount.test.cjs
```

Browser tests use an existing Playwright installation, without a project dependency.
Set `PLAYWRIGHT_MODULE` to its module directory if needed and `BROWSER_EXECUTABLE`
to an installed Chromium browser executable. In a separate shell:

```powershell
$env:VITE_SIGNALTUNER_API_URL='http://127.0.0.1:4317/backend'
./node_modules/.bin/vite.cmd --host 127.0.0.1 --port 4317 --strictPort
```

Run `npm run test:personal-browser`. Only the test browser intercepts SDK/API
responses with synthetic accounts. Optional `PERSONAL_TEST_SCREENSHOTS` sets an
output directory. Coverage includes restoration, email sign-in/profile completion,
hidden SSO controls and absence of automatic token/API calls,
activation/reconnection, free tests, report retry/recovery, account isolation,
Settings/Support, historical authorization, the original meeting view and
390/760/1440px light/dark layouts. This does not verify a live Teams tenant.

Backend: `dotnet test signaltuner-back.Tests/signaltuner-back.Tests.csproj`.
`PersonalSignalTests` additionally uses the existing guarded disposable MySQL
fixture in backend `docs/web-meetings.md`; with that fixture running, set
`SIGNALTUNER_LOCAL_MEETING_DB=1` and filter `FullyQualifiedName~PersonalSignalTests`.
Never run against production. No migration is required.

Local validation on 2026-10-27: TypeScript and the frontend build passed; all 48
frontend tests and 13 browser scenarios passed. Backend build passed; the default
suite passed 325 tests and skipped 13 opt-in cases. Both PersonalSignalTests passed
with the isolated database enabled. Synthetic light/dark screenshots were reviewed
at desktop and narrow widths. Existing Vite deprecation and backend compiler/package
advisory warnings remain; this work does not upgrade dependencies.

The sign-in setup follow-up passed TypeScript, the frontend build, 12 personal-data
tests and 14 browser scenarios. Setup now follows the same separate-screen gate as
the meeting app. This follow-up changes only Teams frontend behavior; it needs no
additional backend, schema, Azure configuration or manifest change.

## Deployment impact

Personal analysis controls follow-up (2026-10-28): removed the personal-view
refresh controls and Credit status display, and reused the in-meeting individual
confirmation dialog for Analysis + Report. Idle actions remain enabled; active
analysis or an in-flight action prevents duplicate submissions. Validation passed
TypeScript, 24 personal-data/live-analysis unit tests, all 48 browser scenarios,
and the DreamHost production build. The confirmation screenshot was reviewed.
This follow-up requires only a Teams static frontend upload; it adds no backend,
database, Azure configuration, manifest, or package change. No upload was performed;
hosted Teams verification remains separate. The original feature's broader
deployment prerequisites are recorded below.

- Changed projects: Teams app and backend.
- Dependent projects: website retains its signal route; WPF supplies existing Test
  capability/telemetry. Website, WPF and SEO need no code change or release.
- Database/configuration prerequisites: existing remote-analysis/self-Test features
  must be deployed and enabled. No new SQL, setting, permission, manifest or package.
- Backward compatibility: additive route and populated optional timestamps. Older
  desktops can report but cannot run Test. Existing clients remain compatible.
- Required deployment order: every backend instance, then Teams hosted code.
  Personal connectivity will be unavailable against an older backend.
- Azure: publish through the existing pipeline to **App Service `signaltuner-back`,
  resource group `SignalTuner`**. Verify the intended production slot via **Azure
  portal > App Services > signaltuner-back > Deployment Center > Logs**. No
  Environment variables or Entra settings change. The Teams frontend now targets
  DreamHost Apache at `teams.signaltuner.com`; follow [the static hosting guide](dreamhost-hosting.md).
  No Azure frontend App Service is needed. Do not invoke package/publish lifecycle
  operations as part of a frontend upload. This hosting conversion requires no
  additional backend deployment beyond the feature prerequisites above.
- Post-deployment verification: email/password sign-in with SSO hidden in separate personal/meeting storage
  partitions, activation, full five-minute zero-credit Test, Report reservation/
  consumption/refund and email, history authorization, account changes and Teams
  high-contrast theme. Local checks do not establish hosted readiness.
- Rollback: revert Teams hosted code first; the additive backend route may remain.
  Keep existing Test-capable backend/desktop support while requests are running.
- Projects still requiring deployment: backend and Teams hosted app. No production
  deployment, Azure change, package or live analysis was performed.
