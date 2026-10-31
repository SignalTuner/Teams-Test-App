# Public Microsoft Teams Store preparation

Production package updated 2026-10-28. **ZIP prepared and schema-validated;
Store certification readiness remains unverified.** Customer-tenant functionality,
live mobile verification and certification evidence remain open. A new production
ZIP was generated at the user's request; no upload, publication, installation,
Azure change or migration was performed. Existing generated local files are stale
and must not be submitted.

## Production package 1.0.18 (2026-10-28)

- User-confirmed production `TEAMS_APP_ID` in `env/.env.prod`:
  `12a01999-ead3-4e8a-9fca-0263cc470316`. Local environment IDs are unchanged.
- Source manifest version advanced from 1.0.17 to 1.0.18. Bot and
  `webApplicationInfo.id` remain `e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`;
  groupChat-only notification bot, meeting contexts, personal dashboard, RSC
  permissions, hosted SSO resource and domains are preserved. SSO stays disabled.
- Fresh ZIP: `appPackage/build/SignalTuner-production-1.0.18.zip`.
  Resolved manifest and icons: `appPackage/build/production-1.0.18/`.
  Exactly `manifest.json`, `color.png`, and `outline.png` are at the ZIP root.
  No development URLs, unresolved variables, source files or credentials are included.
  The prior local and production-test archives were not reused or overwritten.
- ZIP SHA-256: `2B8B782BD85E3DA56EEF494478E4B8252498E146EC6975894F84355CB38EF11B`.
- Production preflight and icon pixel/dimension checks passed. The resolved
  manifest passed Microsoft's published 1.27 schema using Ajv 6 / Draft-04;
  schema SHA-256: `ADC66F66F44B0FCD42EF36DB4319EDBB5225BFFEA95FB64A83EA2F2706B7F9A5`.
  The packed manifest matches the validated file. Ajv reported informational
  warnings for sibling keywords beside `$ref` in Microsoft's schema.
- Live read-only HTTPS checks returned 200 for `/tabs/home`, `/tabs/config`,
  `/tabs/privacy`, and `/tabs/terms` on `teams.signaltuner.com`, with a CSP allowing
  Teams framing and no X-Frame-Options blocker. These checks do not establish
  in-client rendering, installed app identity, or backend workflow success.

This is a distinct Teams app identity from the earlier local/test app, not an
in-place upgrade of that app. Validate/install the new identity in the intended
test tenant and meeting; existing tabs under the old identity do not automatically
become this app. Do not change the bot/Entra registration ID to the Teams app ID.

### Required backend configuration alignment

After the new app is available in the intended tenant, update the existing
non-secret **`TeamsInvitations__TeamsAppId`** setting to
**`12a01999-ead3-4e8a-9fca-0263cc470316`** at **Azure portal > App Services >
signaltuner-back (resource group SignalTuner, production, not signaltuner-back-stage)
> Settings > Environment variables > App settings**. Apply the change and allow
the App Service restart. Verify Load Participants, selection, Send Invitation,
meeting-chat delivery, and the message's Open SignalTuner button in a real meeting.
This is an existing configuration change, not a backend code release or SQL migration.
The old identity's invitation links and the new package will be inconsistent until
the setting is aligned. Reverting that setting restores links to the old identity;
retain the existing bot identity and authorization checks during rollback.

No Azure setting was changed in this task. The latest recorded backend inspection
still named the local app ID; current live settings were not independently inspected.
The backend invitation document records this required production value as well.

Before submitting, run Developer Portal validation against this exact ZIP and
complete the open cross-tenant, mobile, bot-dialog/delivery and reviewer-evidence
checks below. Local schema validation is not Microsoft Store approval. See
[Microsoft's package requirements](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/apps-package)
and [Store validation guidelines](https://learn.microsoft.com/en-us/microsoftteams/platform/publishing/office-store-approval).
Only Teams package/configuration and backend documentation changed for this
release; no frontend hosting, backend, WPF, website, or SEO deployment was performed.

## Prepared changes

- Source manifest 1.0.18: SignalTuner name without a local/dev suffix, SignalTuner
  Inc. publisher (from the existing legal pages), actual website, and descriptions
  covering account, desktop, credit, consent and meeting-invitation requirements.
- Color icon: 192 x 192 PNG, opaque white background, symbol within the centered
  120 x 120 safe area. Outline: 32 x 32 PNG, white symbol on transparency. Both use
  the symbol from `src/Tab/assets/signaltuner-logo-horizontal.png`, already used
  by this product. Rebuild with
  `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/build-brand-icons.ps1`.
  These are production assets, not a new generated brand or third-party artwork.
- Development project is SignalTuner Teams App; npm name is `signaltuner-teams-app`.
  Physical directory rename is pending an open-folder Windows lock; the prepared
  helper updates all five AGENTS maps and workspace together after the move.
  See [rename instructions](dreamhost-hosting.md#rename-and-workspace).
  Local configuration files, ports and app IDs remain distinct. A Teams-generated
  "custom app" label belongs to the distribution channel and cannot be removed
  through the app's display name.
- Read-only preflight and `env/.env.prod.example` with the confirmed DreamHost
  origin/resource URI and a deliberately unfilled production app ID.
- [DreamHost static build](dreamhost-hosting.md) at `dist/dreamhost`, with Apache
  routes, framing/security/cache headers and isolated public Azure API config.
  No Node runtime or Azure frontend is needed. Toolkit remote provision/publish
  is disabled; deploy only builds files. Local debugging remains intact.
  Static build/TypeScript, six static checks, two compiled-browser checks, 50
  frontend tests, 25 existing browser scenarios and the rename-helper test passed.
  Live Apache/TLS/iframe/CORS verification remains a release gate.

## Submission blockers and verification gaps

Update 2026-10-27: the user deferred multi-tenant SSO. `TEAMS_SSO_ENABLED = false`
now hides SSO sign-in/account-creation controls and blocks automatic/manual Teams
token requests. Email/password sign-in, registration and existing sessions remain.
Only the Teams frontend needs deployment for this pause; no Azure, backend,
permission, manifest or database change is required. Hosted SSO configuration below
is deferred until re-enabling SSO. Cross-tenant bot access, invitations and the other
Store gates still require resolution; hiding SSO does not establish Store readiness.
Validation for the pause initially passed frontend build, TypeScript and 12
personal-dashboard unit tests. The mobile-commerce follow-up now runs the browser
coverage using isolated temporary Playwright tooling and Edge, with synthetic SDK
and API responses; application dependencies are unchanged. No live sign-in,
package or deployment was performed.

The [mobile commerce policy](teams-mobile-commerce.md) now gates purchase, plan,
trial, billing and indirect website routes through TeamsJS host context. Mobile
retains server allowance and entitled functionality with neutral zero-allowance
feedback. Desktop/web links identify leaving Teams; registration, profile consent,
Settings and the Teams terms disclose external payment handling. Actual Teams
mobile and external checkout verification remain necessary before certification.
The completed follow-up passed 50 frontend tests, 25 synthetic browser scenarios,
full build, TypeScript, manifest readiness and Microsoft's 1.27 schema for local
and synthetic hosted input. Manifest description wording now directs signup
inside the app instead of linking to the website; source version remains the
undistributed 1.0.17. No generated package was changed.

| Priority | Finding | Required action / owner |
| --- | --- | --- |
| Verify | Production ID is confirmed and the new 1.0.18 ZIP is schema-valid. Basic live route/header checks passed; in-client hosting verification remains outstanding. The old local ZIP points to localhost. | Use only `SignalTuner-production-1.0.18.zip`; validate the installed production identity and hosted frontend in Teams. Teams app owner. |
| Deferred SSO / bot blocker | SSO entry points are disabled pending multi-tenant work. Backend `TeamsSsoTokenValidator` still trusts one tenant; Azure Bot is documented as SingleTenant. | Keep SSO disabled. Verify email/password and bot/meeting workflows with a second independent tenant before submission. Design and approve multi-tenant SSO before re-enabling it. Never disable issuer/audience checks or substitute a wildcard tenant. Backend and Entra/Bot owners. |
| Blocker | Bot invitation access remains unresolved in the latest SSO record. | Demonstrate Load participants, selected mentions, meeting deep link and authorization failures in a real supported meeting before advertising the flow as working. Backend plus Teams app. |
| Verify | TeamsJS host gating now hides commerce and indirect purchase routes on iOS/Android/iPadOS, with desktop/web external-navigation and payment disclosures. | Verify actual mobile Settings, signup, legal/support, zero-allowance Test Analysis, entitled reports, setup and incident routes. Unknown/failed context must remain restricted. See the [host policy and test coverage](teams-mobile-commerce.md). Teams frontend owner. |
| Verify | Manifest 1.0.18 retains the background notification-only bot in `groupChat`; no personal bot or channel scope. | No conversational commands/welcome/Help are planned. Validate the existing tab-launched Adaptive Card `task/submit` flow, participant selection and chat delivery in Teams. See the [compatibility assessment](meeting-participant-roster.md#background-bot-and-meeting-scopes-2026-10-27); schema validation is not runtime certification. |
| Certification follow-up | Backend invitation text mentions recipients but does not identify the initiating user. | Review user attribution required for user-triggered bot messages; do not conflate notification-only classification with exemption from notification-content rules. Backend implementation remains unchanged by request. |
| Verify | Configurable tabs now use `groupChat` only; `supportsChannelFeatures` removed; personal static tab and meeting contexts retained. | Test standard meeting installation and upgrades from older manifests, personal dashboard access and absence of channel installation. Explain the narrowed supported surfaces to reviewers and existing users. |
| Verify | User confirms DreamHost subdomain exists; static build/routes are locally validated. No upload or live hosting verification performed. Backend source CORS includes the origin. | Verify DreamHost web directory, DNS, certificate, Apache `.htaccess`, real asset 404s, route refreshes, framing headers and live backend CORS. Check config/privacy/terms/support and root auth routes. |
| Verify | Desktop setup offers a macOS DMG although the supplied project map documents the Windows client. | Verify the downloadable Mac client and supported features; otherwise correct the download flow/support wording before release. Test mobile/no-client users can continue without syncing. If a maintained Mac project exists, add it to the shared project map separately. |
| Verify | External navigation/download actions exist throughout Settings and client setup. | Clearly mark leaving Teams; warn before executable downloads; test keyboard, focus, high contrast, dark mode and narrow meeting/mobile layouts. |
| Verify | Publisher identity, support SLA/contact, policy publication and current deployment evidence are not independently verified. | Confirm the exact Partner Center verified publisher, reachable public support/privacy/terms pages, current data handling disclosures, reviewer accounts and feature deployments. |

This review preserves backend authentication, permissions, API routes and
billing rules. Changes to those need a coordinated implementation/release after the specific
public-tenant design is approved. Cosmetic readiness cannot resolve these blockers.

## Production configuration, DreamHost and Azure locations

All values in `.env.prod` are non-secret, Teams-project configuration. This is a
separate environment, not a replacement for `.env.local`.

| Key | Source and purpose | Default |
| --- | --- | --- |
| `TEAMSFX_ENV` | `prod`; toolkit environment and preflight | No production fallback |
| `TEAMS_APP_ID` | Confirmed production app ID from Teams Developer Portal; consumed by the manifest/toolkit | Blank; do not substitute bot ID or assume local ID |
| `TAB_DOMAIN` | Production tab DNS hostname; manifest valid domains | `teams.signaltuner.com` |
| `TAB_ENDPOINT` | HTTPS origin, no trailing slash; manifest routes | `https://teams.signaltuner.com` |
| `TEAMS_SSO_RESOURCE` | Entra exposed API URI matching that origin and bot identity | `api://teams.signaltuner.com/botid-e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`; not yet live-verified |

The static build uses the separate public
`config/dreamhost/.env.dreamhost` API URL. No Azure frontend resource ID is needed.

1. **Teams tab hosting:** DreamHost > Websites > Manage Websites >
   `teams.signaltuner.com` > Manage. Confirm the assigned website user/web directory
   and HTTPS certificate, then upload only `dist/dreamhost` contents by SFTP when
   authorized. See [the full hosting and rollback guide](dreamhost-hosting.md).
   Test real Apache headers, route/asset behavior and Teams iframe loading before
   packaging. Legacy Bicep is not part of the production workflow.
2. **Entra SSO:** Azure portal > Microsoft Entra ID > App registrations > All
   applications > `signaltuner-teams-bot`
   (`e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`) > Expose an API. Once the host and
   identity design are approved, add the hosted URI
   `api://teams.signaltuner.com/botid-e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`
   and set the identical `TEAMS_SSO_RESOURCE`. Do not overwrite local configuration
   or assume multi-domain Teams SSO support; resolve the separate environment and
   tenant rollout before re-enabling SSO. Verify `access_as_user` and Teams client
   preauthorizations. Save and allow propagation, obtain new tokens through normal
   sign-in and verify audience/issuer; never collect or commit tokens. Supported
   account types and bot tenant architecture require a separately approved design;
   changing Entra alone will not change the backend's single-tenant validator.
3. **Backend:** Azure portal > App Services > `signaltuner-back` (resource group
   `SignalTuner`, production, not `signaltuner-back-stage`) > Settings > Environment
   variables > App settings. Preserve `TeamsSso__TenantId` and `MicrosoftApp*`
   settings until the coordinated identity design is approved. For invitation
   links, set/update `TeamsInvitations__TeamsAppId` to the final production manifest
   **app ID**, not the bot ID, after confirming that app's installation. Apply and
   restart as prompted, after compatible backend code is on every instance.
   Verify the Open SignalTuner link in the intended meeting. Backend CORS already
   includes `https://teams.signaltuner.com`; a different host requires a reviewed
   backend CORS change and deployment.
4. **Bot:** Azure portal > Azure Bot > `signaltuner-teams-bot` (documented resource
   group `DefaultResourceGroup-EUS`) > Configuration and Channels > Microsoft Teams.
   Verify identity/tenant, channel and the existing backend `/api/messages`
   endpoint. Review the bot profile/display name/icon in the available bot profile
   settings so it matches the manifest branding. Do not rotate credentials or
   change bot type/endpoints as a cosmetic step. Save any approved profile changes
   and verify actual Teams rendering. Public-tenant compatibility remains open.

## Validation before creating a ZIP

Run from this repository:

```powershell
npm run validate:store
powershell -NoProfile -ExecutionPolicy Bypass -File .vscode/scripts/generate-local-manifest.ps1 -ValidateOnly
npm run build
npx --no-install tsc -p tsconfig.app.json --noEmit
npm run build:dreamhost
npm run test:dreamhost
# Once .env.prod contains confirmed values:
npm run validate:store -- -RequireProduction -EnvironmentFile env/.env.prod
```

The preflight validates branding, icon dimensions/pixels and, when supplied,
production inputs, origin/resource alignment, unresolved placeholders and
development hosts. It renders only in memory. It does not validate Microsoft's
full JSON schema, ownership, live endpoints, existing app/version uniqueness,
tenant consent or certification policy. Source-only success is not release approval.
The production example deliberately fails validation until the existing approved
production Teams app ID is supplied.

Validation completed during this review: source preflight, local manifest
`-ValidateOnly`, full `npm run build`, frontend TypeScript check and Git whitespace
check passed. Preflight exercised six cases: source-only success; missing, blank
and local production-input rejection; synthetic valid inputs; and mismatched SSO
resource rejection. Both icons were visually inspected, including the white icon
against a dark background. Build emitted the existing Vite CommonJS API deprecation
warning. The later 1.0.17 scope change additionally passed Microsoft's published
1.27 JSON schema using Ajv/Draft-04 for local and synthetic hosted inputs in memory,
three invitation tests, six meeting-analysis tests, TypeScript and frontend build.
See the roster document for the schema hash and compatibility limits. Developer
Portal package/Store validation and live Teams tests were not run; production
configuration remains incomplete. Existing local ZIP/manifest files were not regenerated.

After explicit authorization to package, generate the production manifest and
ZIP containing `manifest.json`, `color.png` and `outline.png` at its root. Do not
zip the repository or the entire `appPackage` directory. Retain the correct app
identity and increment the manifest version if 1.0.17 has already been distributed.
Run Microsoft schema/package validation and Developer Portal Store validation
against that exact ZIP. Resolve all errors and assess every warning.

`m365agents.yml` now disables remote provision/publish and only prepares static
files in its deploy lifecycle. Packaging, Developer Portal updates and public
Store submission must be separately authorized. The local Toolkit workflow is
retained for debugging and still has package/install actions; it is not a
production packaging or read-only validation command.

## Store listing and reviewer material

- Use the prepared manifest name/descriptions consistently; verify all claims
  against the final released behavior, including account/desktop/credit limitations.
- Establish Partner Center access and publisher verification. Prepare publisher
  attestation; Microsoft documents different timing for new and existing listings.
- Capture actual Teams screens with synthetic accounts: personal dashboard,
  meeting health, analysis consent/progress and service incidents. Provide up to
  five screenshots at 1366 x 768, each no larger than 1,024 KB. Do not reuse generic
  template screenshots, live meeting data or unverified promotional mockups.
- Supply a walkthrough recording and detailed test notes privately. Include new
  account and returning-account flows, two-person meeting setup, desktop activation,
  credits/consent, invite/roster, cancellation, PDF completion, unsupported contexts,
  sign-out and support. Arrange reviewer access and enough credits/devices to test
  every advertised role. Never commit passwords or reviewer credentials.
- Test desktop/web and iOS/Android, multiple themes, keyboard/screen reader,
  slow/offline responses and a second tenant. Confirm privacy/terms are publicly
  readable and the support path works. Send test mail/invitations only with explicit
  authorization. Public Store submission happens through the documented Developer
  Portal/Partner Center process after these gates, with separate user authorization.

## Cross-project follow-up required

- Project: `signaltuner-back`, with the Teams app and external Entra/Bot resources.
- Contract or behavior affected: deferred public-tenant SSO, bot installation/invitations and
  production app deep links.
- Required code/configuration changes: approve and implement customer-tenant
  bot identity/consent handling and invitation access; multi-tenant SSO is deferred.
  No conversational bot functionality is required by the chosen design.
  Review initiating-user attribution in `TeamsMeetingInvitations.Invitation` before
  certification; the user requested preserving this implementation in this task.
- Suggested search terms or likely files: `Services/TeamsSsoTokenValidator.cs`,
  `Controllers/BotMessagesController.cs`, `Services/TeamsMeetingInvitations.cs`,
  `docs/teams-sso.md`, `TeamsSso__TenantId`, `MicrosoftAppType`,
  `TeamsInvitations__TeamsAppId` (code beneath `signaltuner-back/`).
- Compatibility risk: existing local and installed clients use the current bot
  and SSO identity. Preserve signed-token validation and explicit tenant isolation.
- Required deployment action/order: approved identity configuration and all backend
  instances first, hosted Teams frontend next, production package/install tests,
  then public submission. Do not publish an incompatible manifest ahead of backend.

## Deployment impact

- Changed projects: Teams app. Five AGENTS path maps/workspace remain on the
  current directory until the prepared rename helper can run after the folder
  lock is cleared. Existing application work preserved; sibling code unchanged.
- Dependent projects: backend and installed desktop client for advertised workflows;
  website for external account flows. No sibling code changed in this preparation.
- Database/configuration prerequisites: no migration or schema change; confirmed
  production environment and resolved public-tenant design still required.
- Backward compatibility: app/bot IDs, routes and backend authentication unchanged.
  The approved 1.0.17 manifest narrows bot/configurable-tab scopes to `groupChat`
  and sets notification-only mode. Personal dashboard remains; old channel/personal
  bot installations need upgrade review. The frontend SSO pause is separate.
- Required deployment order: resolve blockers, backend/configuration if changed,
  hosted frontend, production package and installation checks, Store submission.
- Post-deployment verification: cross-tenant sign-in, all install scopes, bot
  invitations, consent/credits, telemetry, reports, mobile commerce and legal/support.
- Rollback considerations: retain existing identities and security checks. Restore
  earlier branding through a new higher manifest version if an update is required;
  frontend and backend compatibility must be assessed separately.
- Projects still requiring deployment: DreamHost static frontend and later Teams package; backend
  will require deployment for the unresolved bot/public-tenant work. This task did
  not deploy any component. Functionality may be inconsistent until required
  components and configuration are released together in the stated order.

## Official references

- [Teams icon requirements](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/design/design-teams-app-icon-store-appbar)
- [Submission checklist](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/prepare/submission-checklist)
- [Store validation guidelines](https://learn.microsoft.com/en-us/microsoftteams/platform/publishing/office-store-approval)
- [Package contents](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/apps-package)
- [Publishing overview](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-publish-overview)
