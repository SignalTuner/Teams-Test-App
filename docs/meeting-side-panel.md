# Compact meeting side panel

The Teams frontend selects the compact layout only when the TeamsJS host context
reports `page.frameContext === "sidePanel"`. Screen width does not select it.
Normal meeting tabs and the personal dashboard retain their existing layouts.

- Meeting Overview contains Meeting Score, Analyze Meeting, and Available credits.
- Each participant shows name, email, Signal Score and the existing Analyze action.
  Name and email share a wrapping row. The centered score matches the Meeting
  Score gauge size, with Analyze beneath it. Both analysis buttons use the same
  compact 160px width as the tab's Analyze Meeting button, capped by available space.
  Participant confirmation/approval and entitlement checks remain unchanged. Live
  inline telemetry, role/avatar, category metrics, Invite and Prompt are omitted
  from the compact list; the full meeting tab retains them.
- The Microsoft Teams service-status component is unchanged.
- Open in Teams tab follows the service-status section.
- The label Available report allowance is now Available credits in the meeting
  controls, personal analysis and the restricted/mobile Settings view. The backend
  still owns the balance (`availableReports`); credit semantics are unchanged.

## Tab navigation and confirmation

`meetingTabNavigation.ts` constructs a Teams `/l/entity/` link using the deployed
`context.app.appId`, configured `context.page.id`, current `chat.id` and tenant.
It sets `openInMeeting=false`, with URL-encoded `contextType: chat` and chat ID.
Missing navigation identity produces safe feedback; it never falls back to a
different app, meeting, personal dashboard or external purchasing page.

Analyze Meeting uses the same link with a unique
`subEntityId: confirm-meeting-analysis:<UUID>`. The destination reads the supported
TeamsJS `page.subPageId` field and opens its existing confirmation dialog only in
`content` context with a meeting/chat present. This hint does not send analysis
requests or consume credits. The user must still select Start meeting analysis.
Existing retries retain their saved roster and request ID. The plain Open button
has no confirmation hint.

The confirmation explains the five-minute Connectivity Health Analysis, report
email recipients, one credit per report, and no charge for disrupted analyses.
Participants appear by email (name fallback when email is unavailable), with
eligible connected accounts initially selected. Users can deselect participants;
Credits required tracks the selection and an empty selection disables Start.
Disconnected accounts are visible but cannot be selected for a new run. Only the
selected IDs are sent through the existing `participantUserIds` API field. While
a submission is unconfirmed, its roster and request ID remain locked, including
after reload; Start meeting analysis safely resubmits that same operation. A
validated response clears the saved operation, so the next explicit start uses
a fresh request ID and the participant selection is editable again. The dialog
always keeps its start form, including after cancellation/completion; it never
switches to a results view or shows New/Retry analysis buttons. Per-participant
progress stays in the analysis sections. No email addresses are added to persisted
operations. Backend consent/authority, self-approval, credit reservation/release,
and report delivery rules remain unchanged; this copy does not require a new
desktop permission prompt when existing permission rules already approve a run.

The individual Analyze dialog uses the same explanation and shows the requesting
account's Available credits, loaded through the existing account endpoint. It
refreshes on opening, focus/visibility, every minute while open, and request credit
state changes. Loading and Unavailable are distinct from a confirmed zero balance.
This display does not change backend authorization, consent, credit enforcement,
request creation, cancellation or email retry behavior. No purchasing links are
added, including on Teams mobile.

Individual-dialog validation on 2026-10-28: TypeScript, ten account helper tests,
three targeted browser scenarios and the DreamHost build passed. Narrow/wide
screenshots were reviewed. Browser checks cover the requested copy, updated and
zero balances, unavailable balance handling, no implicit analysis submission,
and opening/closing the dialog from personal meeting history. Tests use synthetic
API responses; live desktop consent and report delivery remain separate checks.

After an individual request is accepted in Pending, Approved or Running state,
its confirmation closes and the participant's analysis section opens. Meeting
confirmation closes when at least one participant request is accepted. Each
device-rate-limited response is an exception: the confirmation remains open with
the requested explanation and a list of affected participant emails (name fallback).
Only participant results with status 429 enter that list; a whole-request HTTP
429 does not identify or blame participants. Accepted participants are marked
Analysis requested, unchecked and disabled for the remaining dialog session, so
adjusting the selection does not submit them again. A new explicit start sends
only the remaining selected participants. Each
accepted participant uses its own server state: Pending/Approved displays yellow
"Waiting for approval..." below Analysis for {name}; Running replaces it with
green "Five-minute analysis in progress" and the existing bold countdown. The
status stays visible with collapsed charts, and Running no longer waits for metric
samples before the UI updates. The countdown uses server telemetry elapsed time
when available; existing consent/self-approval and billing rules are unchanged.

The compact side panel shows the same status beneath an Analysis for {name} label
after an individual request, and continues polling with the confirmation closed.
Full-tab progress polling also continues. Unconfirmed/failed submissions keep the
dialog actionable without claiming approval is pending. A partially accepted
meeting run leaves a visible summary with instructions to reopen Analyze Meeting
and select participants. Reopening a participant's Analyze control shows its
normal start form; Start is disabled while that participant has an active analysis.
Cancelled/completed requests do not replace the start form or require a separate
New analysis action. Cancellation and report-delivery actions remain in the
participant analysis section (also beneath the compact side-panel status).
Unconfirmed individual submissions preserve their request ID across closing and
reopening; pressing the normal Request analysis button safely confirms the same
request. A confirmed response permits a fresh ID for the next explicit start.

Stable-confirmation validation on 2026-10-28: TypeScript, 18 analysis unit tests,
all 39 dashboard browser scenarios and the DreamHost build passed. Coverage
includes cancelled/completed individual requests, active-run start disablement,
cancelled/completed meeting runs across reload, fresh request IDs and participant
selection, and unchanged uncertain-response IDs. Only the Teams frontend needs
deployment; no backend, WPF, database, Azure or manifest changes were made.

Status handoff validation on 2026-10-28: TypeScript, 11 live-analysis tests, seven
meeting-analysis tests, all 35 dashboard browser scenarios and the DreamHost
build passed. Browser coverage verifies yellow Pending/Approved status, green
Running status and countdown before metric samples, collapsed-section visibility,
closed-dialog polling in the side panel, independent batch participant states,
partial failures and uncertain-response retries. The rendered meeting status
layout was reviewed. These are frontend-only changes with the existing APIs;
no backend/WPF release, database migration, Azure setting or Teams package change
is needed. Nothing was deployed; live Teams/device approval remains to be verified.

Dialog errors and inline analysis errors share the alert styling in `App.css`:
transparent background, red border, and white text in dark/high-contrast themes.
Light theme keeps dark text for readability. The rule outranks dialog paragraph
styles and also covers personal analysis and historical-meeting dialogs outside
a live meeting. On 2026-10-28, TypeScript and the DreamHost build passed; synthetic
browser checks verified computed colors/borders in five meeting/personal error
views across all three themes, and the dark dialog rendering was reviewed.

Personal Test Analysis and Analysis + Report use the same public device-limit
message as individual meeting analysis: "This device has received too many
requests. Try again later." The shared rate-limit helper allows only known public
messages; unknown/malformed 429 responses use requester-level wait guidance and
never expose server diagnostics. Device limits remain backend-controlled; no API,
authorization, rate-limit threshold, credit or schema change is involved.

Participant-limit validation on 2026-10-28: TypeScript, 20 personal/meeting unit
tests, all 45 dashboard browser scenarios and the DreamHost build passed. Tests
cover all-limited and partially accepted batches, affected-email lists, editable
selection without resubmitting accepted targets, whole-request 429 handling,
matching personal Test/Report messages, and safe malformed/unknown response text.
Dark dialog screenshots were reviewed. No deployment or backend change was made.

The dashboard consumes each hint once, returning from Settings/Support if needed
and opening the confirmation. Returning to those screens does not replay it.
On window focus/visibility
return, the dashboard refreshes Teams context for a cached tab, accepting changes
only for the same meeting/chat. No internal TeamsJS context-change API is used.
Navigation failures leave the panel available for retry. These hints are UI input,
not authentication or authorization; server enforcement is unchanged.

References: Microsoft's [Teams deep links](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-link-application)
document `openInMeeting=false` and group/meeting chat targeting;
[tab context](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/access-teams-context)
documents `page.subPageId`.

## Validation and rollout

Run `npm run test:meeting-analysis`, `npm run test:personal-browser`, the app
TypeScript check and `npm run build:dreamhost`. Browser fixtures intercept all API
and SDK operations with synthetic accounts; they verify narrow/wide side panels,
normal tabs, encoded links, navigation failure, cold/cached confirmation opening,
and the absence of automatic analysis POSTs. Screenshots cover the compact layout.

Only the Teams frontend needs deployment: upload the rebuilt `dist/dreamhost`
contents to the existing DreamHost site. No backend, WPF, website or SEO source
change, database migration, Azure setting, new permission, dependency, or Teams
manifest/package update is required. Existing meeting-chat tab installation is
required. Rollback restores the previous static frontend files without changing
the existing API, saved analysis requests or credit state.

Live Teams verification remains required: check installed-app ID availability,
side-panel to meeting-chat-tab navigation on desktop/web, a reused tab receiving
the confirmation hint, and the normal explicit confirmation/approval flow.
Local tests do not validate Teams host routing. Nothing was deployed or installed.

Local validation on 2026-10-28: TypeScript, seven meeting-analysis/navigation tests,
all 28 dashboard browser scenarios and the DreamHost build passed. The three
navigation scenarios were rerun after adding the Settings-to-confirmation handoff
and preventing hint replay. Narrow/wide screenshots were reviewed. The existing
Vite CJS deprecation warning remains.

Confirmation redesign validation on 2026-10-28: TypeScript, seven meeting-analysis
tests, five targeted browser scenarios (selection/retry at 320px and 1100px plus
the three navigation scenarios), and the DreamHost build passed. The initial
320px page load timed out; both selection scenarios passed on rerun. Screenshots
were reviewed for narrow layout and paragraph spacing. Fixtures verify selected
IDs, live credit totals, empty-selection disablement, offline exclusion, and
unchanged request IDs/rosters after reload and retry. Live desktop permission,
report email and credit settlement were not exercised by these mocked UI tests.
