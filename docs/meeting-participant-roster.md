# SignalTuner Teams meeting session workflow

SignalTuner no longer attempts to identify or enumerate Microsoft Teams meeting attendees from Teams, Graph, meeting roster APIs, bot roster APIs, chat roster APIs, or any meeting-chat workaround.

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
POST /api/auth/google/start
POST /api/auth/github/start
POST /api/auth/email-magic-code/request
POST /api/auth/email-magic-code/verify
GET  /api/auth/me
POST /api/auth/logout
```

For Teams SSO, the tab initializes TeamsJS, reads context, calls `teamsJs.authentication.getAuthToken()`, and sends that Teams access token once to `/api/User/teams-sso`. The backend should validate provider tokens server-side and return a SignalTuner session token or secure session cookie. Provider tokens must not be exposed back to the frontend after login, and the Teams SSO token should not be used as the long-term SignalTuner app session.

New users receive 5 free credits. Provider identities should be stored in `UserAuthIdentities` using stable provider subject IDs.

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

```http
POST /api/TeamsMeetings/{meetingSessionId}/invite
Authorization: Bearer {signaltuner_session_token}
```

Preferred behavior is to post a general invite message or Adaptive Card into the meeting chat:

```text
SignalTuner is available for this meeting. Open SignalTuner to join the meeting health view and share your connectivity status.
```

This endpoint must not enumerate, target, list, infer, or synchronize Teams attendees. If chat posting is unavailable, return a copyable message and app deep link.

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
