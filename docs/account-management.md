# Account Management Contracts

The Teams app account page uses the SignalTuner backend as the source of truth
for account identity and profile updates.

## Profile name

- `PUT /api/User/profile`
- Request: `{ "firstName": string, "lastName": string }`
- Updates `Users.user_first_name`, `Users.user_last_name`,
  `Users.user_initials`, and `Users.user_display_name`.
- Existing Teams meeting participant display-name snapshots are refreshed for
  the user.

## Email

- `PUT /api/User/email`
- Request: `{ "email": string }`
- Requires an authenticated SignalTuner user with `Users.user_password` already
  set. This prevents a Teams-SSO-only user from changing their email before they
  have an email/password sign-in path.
- Updates `Users.user_email` only. It must not change `Users.user_m365_upn`,
  `Users.user_m365_object_id`, or `Users.user_m365_tenant_id`.
- Rejects email values that already match another user's `Users.user_email` or
  `Users.user_m365_upn`.
- Email/password registration performs the same collision check, so an address
  retained as another account's Microsoft UPN cannot create a second user row.
- Returns a replacement SignalTuner session token for the same `Users.user_id`
  with email claims reflecting the updated `Users.user_email`.

Teams SSO resolution continues to match by Microsoft object/tenant identifiers
and `Users.user_m365_upn`, preferring those stable Microsoft identifiers over an
email-only match. The SSO response returns `Users.user_email` as the account
email and exposes `Users.user_m365_upn` separately. A user who updates their
SignalTuner email therefore remains on the same account when signing in through
Teams.
