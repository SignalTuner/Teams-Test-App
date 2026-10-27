# SignalTuner Teams meeting session workflow

Website integration (2026-10-26): once a user signs in and joins a meeting session,
that same session can appear in their website Meetings list. The website reads
existing signed-in account memberships; bot-only roster entries are not enrolled.
Microsoft Teams is selected as the service. The backend shares its best-effort
Organizer/Co-organizer/Presenter/Attendee labels across both views and no longer
guesses that the first SignalTuner participant is the organizer. Unknown roles remain
Participant. These labels never grant SignalTuner team administration or analysis
approval privileges. Deploy the updated backend, then the website; this feature
requires no new Teams package, Azure setting, permissions or database migration.
See `signaltuner-back/docs/teams-website-meetings.md` for the contract and validation.

Title correction (2026-10-26): the installed TeamsJS meeting-details methods are
Microsoft-internal and cannot be relied on to supply a title. The backend now reads
meeting names through the existing bot using authenticated conversation routing
from TeamsBotActivityLog. It enriches names during join and website reads, without
sending messages or changing SSO. Deploy only the backend for this correction; no
new Teams package or Azure setting is needed. For an older meeting with no bot
routing history, Invite > Load Participants records authenticated context; refresh
the website after one minute. Existing bot access/RSC requirements still apply.

The user also confirmed that Invite > Load Participants succeeded in a new meeting
on 2026-10-26. This does not independently verify sending/delivery of invitations.

SSO update (2026-10-26): the user confirmed package 1.0.15 installation. Fresh
token diagnostics now identify the correct application but Teams tenant `a827...`,
different from the bot/backend tenant `71cd...`; Azure logs show issuer rejection.
A backend-only `TeamsSso:TenantId` correction is prepared, tested, and approved.
The agent saved and verified the approved production Azure setting. The user
subsequently confirmed SSO works again on 2026-10-26; all-instance deployment was
not independently inspected. No new Teams package is required, and this does not fix
bot meeting access. See `signaltuner-back/docs/teams-sso.md` for exact values,
Azure location, compatibility, and verification. Existing SSO validation remains.

The health dashboard lists authenticated SignalTuner users. The Invite dialog separately retrieves the current meeting chat's roster through the existing Teams bot so participants who have not opened SignalTuner can be invited. This invitation-only lookup does not enroll people, create SignalTuner accounts, or synchronize them into the health dashboard.

The source of truth for the dashboard is:

```text
Authenticated SignalTuner users who opened the SignalTuner Teams app and joined the SignalTuner meeting session for the current Teams meeting ID.
```

## Teams tab flow

1. The tab initializes TeamsJS.
2. The tab reads the Teams meeting ID, conversation ID if available, tenant ID if available, and meeting title if available.
3. The user signs in or creates a SignalTuner account.
4. The backend creates or finds a `TeamsMeetingSessions` row keyed by `teams_meeting_id`.
5. The backend creates or updates a `TeamsMeetingParticipants` row for the authenticated SignalTuner user.
6. The dashboard returns only authenticated SignalTuner participants in that meeting session.

Users who do not open SignalTuner and authenticate do not appear in the dashboard.

## Authentication endpoints expected by the tab

```http
POST /api/User/teams-sso
PUT  /api/User/profile
POST /api/auth/google/start
POST /api/auth/github/start
POST /api/auth/email-magic-code/request
POST /api/auth/email-magic-code/verify
GET  /api/auth/me
POST /api/auth/logout
```

For Teams SSO, the tab initializes TeamsJS, reads context, calls `teamsJs.authentication.getAuthToken()`, and sends that Teams access token once to `/api/User/teams-sso`. The backend should validate provider tokens server-side and return a SignalTuner session token or secure session cookie. Provider tokens must not be exposed back to the frontend after login, and the Teams SSO token should not be used as the long-term SignalTuner app session.

Teams SSO sign-in and automatic SSO send `allowAccountCreation: false`. If no SignalTuner account exists for the Microsoft Teams identity, the backend returns `409` with `code: "teams_sso_account_required"` so the tab can send the user through account creation. For Teams SSO account creation, the tab must first show the local "Finish your profile" step and must not call `/api/User/teams-sso` until the user enters first and last name, checks the consent box, and clicks Continue. The create request must send `allowAccountCreation: true`, `firstName`, `lastName`, `termsOfServiceAccepted: true`, `privacyPolicyAccepted: true`, and the legal document versions. The backend records the profile names, acceptance UTC timestamps, and versions on the newly created `Users` row.

When `/api/User/teams-sso` creates a SignalTuner user, or resolves an SSO user with no recorded first or last name, the response includes `profileRequired: true`. The Teams tab must collect first and last name before joining a meeting session, then call `PUT /api/User/profile` with the SignalTuner bearer token. The backend writes `Users.user_first_name`, `Users.user_last_name`, and `Users.user_initials`, and participant/dashboard display should use the recorded name instead of falling back to email.

First-time emails receive 3 free personal credits; previously deleted/reused emails receive no new signup bonus. On the first UTC day of each month, the backend grants 1 credit only to eligible accounts with a zero personal balance. This requires the free-credit migration and all-instance backend deployment; the Teams response shape is unchanged. See `signaltuner-back/docs/free-account-credits.md` for eligibility and rollout. Provider identities should be stored in `UserAuthIdentities` using stable provider subject IDs.

## Meeting session endpoints expected by the tab

```http
POST /api/TeamsMeetings/join
Authorization: Bearer {signaltuner_session_token}
Content-Type: application/json

{
  "teamsMeetingId": "...",
  "teamsConversationId": "...",
  "teamsTenantId": "...",
  "meetingTitle": "..."
}
```

Returns the joined session, current user, credit balance, Microsoft Teams health, unresolved Teams incidents, and authenticated SignalTuner participants.

```http
GET /api/TeamsMeetings/{meetingSessionId}/dashboard
Authorization: Bearer {signaltuner_session_token}
```

Requires the authenticated user to be associated with the session.

## Invite participants

The previous `/api/TeamsMeetings/{meetingSessionId}/invite` call had no backend implementation and has been removed from the tab. There is no new REST invitation endpoint.

1. **Invite** (or **Prompt** on a participant row) opens a native Teams Adaptive Card dialog.
2. **Load participants** sends `signaltuner.invite.load` to the existing bot at `/api/messages` using Teams' authenticated `task/submit` flow. This explicit first step establishes the trusted bot conversation; it does not send a chat message.
3. The bot verifies that the requested meeting resolves to the invoking conversation and that the sender belongs to it, then returns the meeting chat roster. The sender and this bot are excluded. Names are displayed without email addresses. This roster can include people not currently in the call; it is not an online-presence list.
4. Select up to 20 people. Previous/Next shows 40 people per page and retains selections and deselections across pages. Nothing is sent during navigation.
5. **Send invitation** rechecks membership and sends one message mentioning the selected people in the verified meeting chat. The dialog explains that everyone in the chat can see it. Recipients are asked to open SignalTuner in this meeting and sign in. A configured app deep link adds an **Open SignalTuner** button.

This flow supports meeting chats with the bot installed. Channel/team contexts are explicitly rejected to avoid treating an entire team's roster as meeting participants. The client reports unsupported hosts, bot errors, and cancellation appropriately. Rosters over 2,000 entries fail explicitly rather than silently returning a partial list. Anonymous participants unavailable to the Teams chat roster cannot be selected.

The bot uses existing manifest permissions and identity. No new Graph permission, SSO flow, database migration, or credit charge is introduced. Provider-authenticated bot context is the authorization boundary; client-supplied meeting-session metadata is not used as proof of membership.

Backend configuration: optional non-secret `TeamsInvitations:TeamsAppId` (App Service `TeamsInvitations__TeamsAppId`) must equal the installed manifest's `id` / environment `TEAMS_APP_ID`, **not** its `bots[0].botId` or `webApplicationInfo.id`. Set it separately for local/test/production. If absent or invalid, invitations still contain explicit instructions to open the app but omit the link. The existing configurable tab uses entity ID `home`. Verify link navigation in the target Teams client before release.

**Fresh configuration confirmed, 2026-10-24:** the user applied the backend
settings for `signaltuner-teams-bot`, app ID
`e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`, SingleTenant, tenant
`71cdea9d-1b1d-4c9d-ba89-8860aa809e67`. Read-only checks confirmed the identity
and presence of a nonempty secret setting; credential validity remains a live
Teams check. The source manifest (1.0.14) and dialog match. The last generated
local manifest (1.0.13) still targets the original bot. Deploy the current backend,
then regenerate/install the local package when authorized. The user runs the tab
through VS Code against the hosted backend. No application deployment or package
installation is inferred from the configuration confirmation.

The user subsequently confirmed successful backend publication on 2026-10-24.
The next release step is regenerating and installing the local Teams package;
live bot authentication and roster loading remain unverified.

See `signaltuner-back/docs/teams-meeting-invitations.md` for evidence, Azure
configuration locations, retry limits, and tenant validation. Local checks:
`npm run test:meeting-invitations`, TypeScript checks using `tsconfig.app.json`,
and `npm run build`. No package, deployment, or live invitation is authorized by
this documentation. A tab reload alone does not update an installed manifest.

## Approved SSO and meeting-permission correction (2026-10-25)

The user approved the correction. Source manifest **1.0.15** aligns bot, dialog,
and `webApplicationInfo.id` to `signaltuner-teams-bot`. `TEAMS_SSO_RESOURCE` is
configured per environment; the local generator checks it against the HTTPS tab
origin and bot ID. `-ValidateOnly` performs preflight without generating a package.

The agent configured and verified the approved Entra API URI, `access_as_user`
scope, Teams desktop/web preauthorization, v2 tokens, and optional email/UPN claims.
These changes are in **Azure portal > Microsoft Entra ID > App registrations >
signaltuner-teams-bot > Expose an API / Token configuration / Manifest**.
Only the local `api://localhost:3978/botid-e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`
URI is configured; hosted environments require their own verified URI.

The backend now validates the Teams access token before either Teams sign-in
entry point accesses the database. Old SSO identities and UPN-only legacy requests
are rejected. Existing account matching uses verified tenant/object identity;
unlinked email collisions ask for existing SignalTuner email/password sign-in.

On 2026-10-25 the user confirmed publication of the updated backend. Deployed
binaries and all-instance rollout were not independently verified. Generate/install
the 1.0.15 package and re-consent in the actual meeting next. No backend deployment
or Teams package generation was performed by the agent. Live sign-in, meeting membership/consent, and roster loading
remain to be verified. Full contract, exact Azure values, validation results,
compatibility, and rollback: `signaltuner-back/docs/teams-sso.md`.

## Dashboard data

Use the Microsoft Teams service record:

```text
service_id = 3
service_name = Microsoftteams
service_display_name = Microsoft Teams
```

Return unresolved `ServiceIncidents` rows for `service_id = 3` using the project convention for unresolved incidents. If no convention exists, treat `service_incident_status != 'Resolved'` as unresolved.

Participant rows should include only users from the same SignalTuner meeting session:

```json
{
  "userId": 123,
  "displayName": "Alex Wilber",
  "email": "alex@example.com",
  "authProvider": "teams_sso",
  "joinedAt": "2026-07-10T15:00:00Z",
  "lastSeenAt": "2026-07-10T15:10:00Z",
  "signalScore": 92,
  "overallStatus": "Excellent",
  "deviceStatus": "Excellent",
  "workspaceStatus": "Excellent",
  "networkStatus": "Excellent",
  "clientDataStatus": "active",
  "clientIsActive": true
}
```

If the participant has no active desktop telemetry, return `clientDataStatus = "no_data"` and no signal score.

## Analysis endpoints expected by the tab

```http
POST /api/analysis/user
Authorization: Bearer {signaltuner_session_token}

{
  "meetingSessionId": 1,
  "targetUserId": 123
}
```

```http
POST /api/analysis/full
Authorization: Bearer {signaltuner_session_token}

{
  "meetingSessionId": 1
}
```

Access rules:

- The requesting user must be associated with the meeting session.
- The target user must be associated with the same meeting session.
- The requesting user's own telemetry is free.
- Analyzing another active user costs 1 credit.
- Inactive/no-data users are not charged.
- Deduct credits only after validating analyzable telemetry exists.

For now, analysis should return full telemetry and issue objects for non-Excellent values in:

- `signal_bandwidth_status`
- `signal_system_status`
- `signal_network_status`

Do not add AI-generated analysis yet.

## Development-only account endpoints

```http
POST /api/account/test-add-credits
POST /api/account/test-subscription-plan
```

These are test/development-only controls used by the tab account menu. Record test credit increases in `CreditTransactions` with `reason = test_credit_grant`.

TODO: restrict or remove these endpoints before production.

## Database migration

Apply:

```text
sql/2026-07-10_signal_tuner_meeting_sessions_auth_credits.sql
```

That migration replaces the obsolete bot roster cache table with authenticated meeting-session participants.
