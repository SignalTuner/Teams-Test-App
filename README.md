# Overview of the Basic Tab template

SignalTuner documentation: [Temporary activation codes](docs/temporary-activation-codes.md).

The Teams Settings and Support pages match the website's content and styling,
with responsive meeting-sidebar layouts and Teams high-contrast support.
Settings replaces the former Account menu entry: it reads `/api/credits/account`,
saves profile details with `/api/User/UpdateUser`, and retains the existing password
endpoint and password-before-email-change restriction. It uses the signed-in Teams
SignalTuner session, not website local storage. Plans, team management, billing and
account deletion open their existing external flows; the website may require sign-in.
Support includes the full name in the existing support-message body and enforces
the combined 1,000-character limit. No support message is sent until form submission.
The placeholder Teams automation options and testing-credit control have been removed.
Only the Teams tab frontend needs rebuilding/releasing; no backend, Azure setting,
database, permission, or manifest change is required.

Validate these pages with `npm run test:settings`,
`npx tsc -p tsconfig.app.json --noEmit`, and `npm run build:frontend`.
Local browser verification uses synthetic accounts and mocked requests; it does
not confirm live profile updates, support delivery, billing, or account deletion.

Meeting Overview displays available credits above Analyze Meeting using the
backend's `availableReports` total (eligible team credits, monthly allowance, and
personal credits), matching the analysis dialog. It refreshes on load, every minute
while visible, when the tab regains focus, when the dialog opens/closes, and after
a meeting-analysis request. Requests time out after eight seconds; loading and
unavailable states never appear as a zero balance. The backend still enforces all
analysis charges. This display needs only a Teams frontend rebuild/release.

This template showcases how Microsoft Teams supports the ability to run web-based UI inside "custom tabs" that users can install either for just themselves (personal tabs) or within a team or group chat context.

## Get started with the Basic Tab template

> **Prerequisites**
>
> To run the basic tab template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: >=20.
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
>   Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take couple days for the enrollment to take effect.
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 6.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.

**Congratulations**! You are running an application that can now show a basic web page in Teams, Outlook and the Microsoft 365 app.

![Basic Tab](https://github.com/user-attachments/assets/e8121c82-ddbc-493f-9afb-34db2a4b8e73)

## What's included in the template

| Folder       | Contents                                     |
| ------------ | -------------------------------------------- |
| `.vscode`    | VSCode files for debugging                   |
| `appPackage` | Templates for the application manifest |
| `env`        | Environment files                            |
| `infra`      | Templates for provisioning Azure resources   |
| `src`        | The source code for the application    |

The following files can be customized and demonstrate an example implementation to get you started.

| File                             | Contents                                                                                                        |
| -------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| `index.html`                     | HTML file.                                                                                                      |
| `src/Tab/App.css`                | CSS file for the app.                                                                                           |
| `src/Tab/App.tsx`                | Tab source file. It calls `teamsjs` SDK to get the context of on which Microsoft 365 application your app is running.      |
| `src/index.ts`                   | Starting the app using [Microsoft Teams SDK](https://aka.ms/teams-ai-library-v2).                                                                     |
| `vite.config.js`                 | Configuration for Vite build tool.                                                                              |  
| `nodemon.json`                   | Configuration for Nodemon to watch and restart the server.                                                      |

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                 | Contents                                                                                                                                  |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `m365agents.yml`       | This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions. |
| `m365agents.local.yml` | This overrides `m365agents.yml` with actions that enable local execution and debugging.                                                     |

## Extend the Basic Tab template

Following documentation will help you to extend the Basic Tab template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Access data in Microsoft Graph](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk#microsoft-graph-scenarios)
- [Use an existing Microsoft Entra application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Enable the app for multi-tenant](https://aka.ms/m365-agent-multi-tenancy-entra-app)
- [Preview the app on mobile clients](https://aka.ms/teamsfx-mobile)

## Remote analysis with device consent

The meeting participant list includes **Analyze** beside each SignalTuner user.
`src/Tab/RemoteAnalysisButton.tsx` uses the existing SignalTuner Bearer token and
meeting session ID to request a five-minute WPF analysis. The backend verifies
that both users belong to the recorded meeting; Teams context alone is not
authorization. Unregistered guests must sign in first. Every participant row has
an **Analysis for ...** section, initially collapsed even when no analysis is active.
Expand its header for request status and available live telemetry; the separate
**Live details** buttons have been removed. Collapsed headers match the website's
20% shorter height (44px desktop, 40.8px mobile for a single line).

The requester pays one report credit. The target device prompts **Deny**, **Just
once**, or **Always**; Always is revocable in desktop Settings for that requester
and device. The PDF is sent to both server-selected account addresses. The request
dialog handles device selection, cost disclosure, progress, cancellation, and
retries of definitely rejected email copies. An accepted or uncertain copy is
never automatically resent. Reopening the dialog restores the latest request.

The authoritative contract is in `../signaltuner-back/README.md`, **Remote analysis
with device consent**. No Entra scopes, Teams manifests, authentication provider,
origins or API base URL have changed. No new package or external provider is used.

The remote-analysis migration was confirmed applied on 2026-10-08. Rollout: deploy the updated
backend and enable `RemoteAnalysis__Enabled`, release WPF, then deploy the website
and Teams app. The feature is incomplete until all four are available. Roll back
browser controls first; retain the backend worker until reservations settle.
No manifest packaging, publication, deployment or production migration was run by the agent.

Run `npm run build`. The frontend type configuration is `tsconfig.app.json`;
the root project-reference type check also includes unrelated server declarations.
Verify consent, Always/revocation, account switching, requester billing and both
PDF deliveries using two test accounts and an updated desktop client.

## SignalTuner team roles (2026-10-16, pending rollout)

Analyze now uses the per-participant five-minute PDF flow. The old paid live-details
and legacy full-meeting analysis buttons are removed. Server-resolved `approvalBasis`
explains immediate Owner/Admin authority or per-request Member approval; Microsoft
Teams meeting roles do not grant SignalTuner team authority. Outside-team device
consent remains unchanged. Credit spending uses the requester's team allocation,
Owner/Admin fund, monthly allowance, then personal credits. Account links to the
website for team roles, invites, allocations, and credit requests.

Requires the updated backend and additive team migration. The user confirmed
backend deployment and all pending migrations applied on 2026-10-16; this does
not confirm a Teams app release. See sibling
backend `docs/team-roles-and-credits.md` for deployment/rollback and real-tenant
verification requirements. No manifest, identity scope, or package was changed.

## Single desktop sessions and device privacy

Desktop sign-in now replaces the previous session through a backend sign-out flag.
Website and Teams use the current desktop automatically and do not display or
receive physical device IDs. The desktop-session migration was confirmed applied
on 2026-10-21. Deploy the
backend before the updated WPF, website and Teams builds. See the
[contract, compatibility limits, validation and rollout](../signaltuner-back/docs/desktop-sessions.md).

## Analyze Meeting

The meeting health panel's **Analyze Meeting** button opens a confirmation for all
participants currently marked `clientDataStatus = "active"`, including the requester
when active. Confirmation shows the participant count, available account credits,
maximum cost (one credit per report), and existing consent rules. Starting submits
one batch of five-minute PDF analyses; each accepted request appears beside its
participant's Analyze button. That button opens the existing progress/cancel dialog.
The meeting action does not purchase legacy live-detail leases or a combined PDF.

The tab saves only the batch ID and original participant user IDs in session storage,
scoped by signed-in user and meeting. Retrying after a timeout, partial failure, or
reload reuses that operation, even when an accepted report already completed.
**New meeting analysis** explicitly selects the current active roster for a new run.
Unavailable devices do not prevent other participants from being requested.

Deploy all backend instances with `POST /api/remote-analysis/teams-meeting` before
deploying this Teams build. Existing remote-analysis configuration and updated WPF
clients are prerequisites; no new migration, manifest, Entra scope, or WPF/website
release is needed. Before backend rollout the new control cannot work. Roll back
the Teams control first and let accepted requests finish on the existing workers.
See [the API and verification checklist](../signaltuner-back/docs/teams-meeting-analysis.md).

Validation commands: `npm run test:meeting-analysis`, `npm run build`, and
`npx tsc -p tsconfig.app.json --noEmit`. The focused tests use Node's built-in test
runner and the existing TypeScript dependency; they make no real API calls.
Real-tenant consent, credits, desktop execution and PDF delivery require separate
end-to-end verification. No deployment or live analysis is performed by these checks.

## Meeting participant invitations

Invite now opens a native Teams bot dialog: Load participants, select up to 20
people from the current meeting chat, and Send invitation. The bot posts one
message mentioning those people and asking them to open SignalTuner and sign in.
Selections persist across pages. The roster includes chat members who might not
currently be in the call; channel meetings are not supported by this flow.

Deploy the backend bot handler before the Teams tab. Configure the optional
backend `TeamsInvitations__TeamsAppId` with the installed manifest app ID to add an
Open SignalTuner link; without it, the invitation contains opening instructions.
Run `npm run test:meeting-invitations`, `npm run build`, and
`npx tsc -p tsconfig.app.json --noEmit`. Real Teams bot installation, mentions,
and meeting-context link navigation still need tenant verification. See
[the invitation workflow](docs/meeting-participant-roster.md#invite-participants)
and the backend's `docs/teams-meeting-invitations.md` for rollout and limits.

Following the fresh investigation on 2026-10-24, the user configured the hosted
backend for Azure Bot `signaltuner-teams-bot` (app ID
`e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`, SingleTenant). Read-only inspection
confirmed the identity settings and a nonempty credential setting without
revealing its value. Source manifest 1.0.14 and the invitation dialog match.
The last generated local package still references the original bot. The user subsequently confirmed successful backend publication. Local package
regeneration/installation is the next release step; live credential validity and
roster loading are not yet verified. VS Code
serves the local tab, which calls the hosted backend. See the backend invitation
document for the precise Azure locations, evidence, rollout, and rollback.

The subsequent approved 2026-10-25 correction moves the source manifest to
**1.0.15**, aligns SSO/RSC with the bot, and validates Teams tokens on the backend.
The agent applied the local Entra SSO configuration; the user confirmed publication
of the updated backend on 2026-10-25. Teams package 1.0.15 installation and live
verification remain outstanding. See `signaltuner-back/docs/teams-sso.md` for exact
Azure settings, compatibility, and rollout. Manifest preflight without packaging:
`powershell -NoProfile -ExecutionPolicy Bypass -File .\.vscode\scripts\generate-local-manifest.ps1 -ValidateOnly`.

## Microsoft Teams service-monitor row

The former Status Incidents panel now matches the website meeting service row:
Microsoft Teams logo/name, current status, unresolved-incident count, and a
collapsed View incidents disclosure. Expanded reports show title, status, updated
time, plain-text description, a safe HTTP(S) provider-report link, and a link to
the website's `/teamsstatus` history. All unresolved reports and the five newest
resolved reports are shown, using the website's ordering and severity rules.
There are no sample incidents or forced active-incident states.

`MeetingServiceStatus.tsx` consumes the existing website endpoints
`GET /api/User/GetServiceStatus/3` and `GET /api/User/GetServiceIncidents/3`
from the configured backend. It refreshes every 30 seconds, skips hidden tabs,
refreshes on visibility return, avoids overlapping reads, and cancels reads after
15 seconds or unmount. Failures show unavailable status/count instead of a false
healthy/zero result. This row uses Teams theme colors and the website's responsive
layout. The Meeting Health panel continues to use the meeting dashboard data.

Run `npm run test:service-status`, `npm run build:frontend`, and
`npx tsc --project tsconfig.app.json --noEmit`. Release only the Teams frontend;
existing backend routes, website, WPF, manifests, dependencies and database are
unchanged. No deployment was performed. After release, verify status, disclosure,
provider/history links, empty/error states, and light/dark/narrow layouts in Teams.
Rollback is a frontend-only rollback with no data/configuration changes.

## Live participant analysis charts

The Meeting Health panel's connectivity-issues count includes only participants
with active client data and a finite signal score in the Poor (70 to below 80)
or Critical (below 70) band. Disconnected participants and missing/invalid scores
do not contribute. When no participants are connected, this metric's icon and
number are light blue. This uses the existing score bands and requires only a Teams
frontend release; no backend, WPF, website or database change is required.

The Analysis section uses the website's `LiveAnalysisCharts` and `AnalysisCountdown`
presentation, ported to TypeScript with Teams theme colors. It includes Wi-Fi,
download/upload, packet loss, latency, CPU and memory histories, Jitter and VPN
cards, sampled duration, countdown, credit state and email delivery controls.
Window sizes, gaps, units, scale caps and severity colors match the website.
Keep these components aligned with `signaltuner-front/src/LiveAnalysisCharts.js`
and `AnalysisCountdown.js` when changing the product's telemetry presentation.

Participant sections always contain all eight charts and the VPN card, even when
idle or empty; collapsing the section hides its body. Empty values remain
unavailable, without fabricated zero samples. Above-chart text follows the website:
idle Analyze/access guidance appears once; pending/approved/running requests wait
for live samples; finished requests with no data say "No samples were received for
this analysis." Running requests show progress/countdown, and canceled requests
show "Canceled" without a countdown. Samples older than eight seconds are delayed;
finished data is labeled as last reported samples. Row error/updating text is shown
when the request dialog is closed. No credit message appears for an unset/None
credit state, matching the website.

Deployment impact for persistent charts and matching text:
- Changed projects: website and Microsoft Teams app.
- Dependent projects: existing backend and WPF telemetry producers, unchanged.
- Database/configuration prerequisites: none added.
- Backward compatibility: existing APIs and older/missing telemetry supported.
- Required deployment order: either frontend first; deploy both for consistent UI.
- Post-deployment verification: expand idle, waiting, running, canceled and
  completed participant rows; check empty plots, messages and two-second updates.
- Rollback considerations: either frontend can roll back independently.
- Projects still requiring deployment: website and Teams app; not deployed here.

Each authorized request reads `GET /api/remote-analysis/{id}?includeLive=true`
immediately, then every two seconds after each active reply, including when the
section is collapsed. First samples open the row; manual collapse is preserved.
Opening a row after reload retrieves the requester's meeting-scoped `/latest`
analysis and its samples. Finished analyses retain last reported samples.
Legacy telemetry remains available only when no remote request is loaded.
This does not add the website's third-party sharing permissions to Teams.

Requests never overlap, time out after eight seconds, back off up to thirty
seconds after failures, pause while the tab is hidden and resume on return.
Delivery-only updates use ten seconds; finished deliveries stop polling.
Late replies cannot overwrite cancellations/new requests; revoked access clears
cached samples. Data older than eight seconds is marked delayed. None of these
reads starts an analysis, spends credits or performs additional speed tests.

The existing WPF client publishes snapshots every two seconds during an approved
remote analysis. Individual sensor cadences remain unchanged (bandwidth tests
and VPN detection retain their existing cadence). Updated WPF and backend live
telemetry support must already be deployed, as for the website. No backend/WPF
code, schema, authentication, manifest, configuration or dependencies change here.
Deploy the Teams frontend; it can be rolled back independently. Older clients
show unavailable/waiting data instead of fabricated zeros. Existing Analyze
Meeting backend deployment requirements above still apply.

Run `npm run test:analysis-live` for synthetic chart and polling checks, alongside
the existing meeting-analysis tests, build and TypeScript check. After release,
verify real approved analysis samples, collapse/reopen, countdown, disconnect,
cancellation, revoked access and report completion in Teams. Local checks do not
verify a live tenant, execute billable analysis or publish a Teams package.

### Jitter chart follow-up

Jitter now uses the same responsive line-and-dot chart as Latency, with ten probe
slots, a 100–1000 ms adaptive ceiling and the existing <=15/30/50 ms green/yellow/
orange bands (red above 50 ms). Values, lines and points share these jitter colors.
The optional `jitterHistoryMs` live field supplies real WPF samples and null gaps;
older payloads fall back to the unchanged scalar `jitterMs` as one point. Repeated
polls do not fabricate history. The two-second polling behavior is unchanged.

This follow-up requires the WPF producer and backend field support as well as
website/Teams chart releases, unlike the initial Teams live-chart integration.
Deploy every backend instance first, then update WPF and website/Teams. No new
schema, permissions or configuration are required; until rollout completes only
the scalar point may appear. Each component remains independently reversible.
See `../signaltuner-back/docs/remote-analysis-live.md` for the complete deployment
impact and verification checklist. No deployment was performed.
