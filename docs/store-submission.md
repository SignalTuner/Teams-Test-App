# Public Microsoft Teams Store preparation

Reviewed 2026-10-27. **Not ready to submit.** Source branding is prepared; production
configuration, customer-tenant functionality, mobile policy and live certification
evidence remain open. No package, upload, publication, Azure change or migration
was performed during this preparation. Existing generated local files are stale
and must not be submitted.

## Prepared changes

- Source manifest 1.0.16: SignalTuner name without a local/dev suffix, SignalTuner
  Inc. publisher (from the existing legal pages), actual website, and descriptions
  covering account, desktop, credit, consent and meeting-invitation requirements.
- Color icon: 192 x 192 PNG, opaque white background, symbol within the centered
  120 x 120 safe area. Outline: 32 x 32 PNG, white symbol on transparency. Both use
  the symbol from `src/Tab/assets/signaltuner-logo-horizontal.png`, already used
  by this product. Rebuild with
  `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/build-brand-icons.ps1`.
  These are production assets, not a new generated brand or third-party artwork.
- Toolkit creation names and internal npm name no longer say Teams Test App.
  Local configuration files, ports and app IDs remain distinct. A Teams-generated
  "custom app" label belongs to the distribution channel and cannot be removed
  through the app's display name.
- Read-only preflight and a blank `env/.env.prod.example`; no guessed app ID or
  production resource. Deployment exclusions now include `.localConfigs`, tests,
  docs and agent tooling so local certificate paths/configuration are not shipped.

## Submission blockers and verification gaps

| Priority | Finding | Required action / owner |
| --- | --- | --- |
| Blocker | Only local environment is populated; `.env.dev` has no hosting or app identity. The last local ZIP points to localhost. | Confirm a production Teams app ID and tab host; populate a separate `.env.prod`. Do not submit the local ZIP. Teams app owner. |
| Blocker | Backend `TeamsSsoTokenValidator` accepts exactly one trusted tenant. Existing Azure Bot is documented as SingleTenant; only a localhost SSO API URI is confirmed. | Design and approve customer-tenant sign-in, bot installation and consent. Test with a second independent tenant. Email/password fallback alone does not prove meeting/bot functionality works across tenants. Backend and Entra/Bot owners. Never disable issuer/audience checks or substitute a wildcard tenant. |
| Blocker | Bot invitation access remains unresolved in the latest SSO record. | Demonstrate Load participants, selected mentions, meeting deep link and authorization failures in a real supported meeting before advertising the flow as working. Backend plus Teams app. |
| Blocker | `SettingsPage.tsx` always exposes Buy credits, View plans and a billing link; it has no Teams host-client gating. | Review and correct iOS/Android commerce/upsell flows using Teams client context, including linked account/legal flows. Keep report balances and consumption separate from purchases. Desktop purchase links must disclose leaving Teams. Teams frontend owner. |
| Blocker | Manifest declares a conversational bot in personal/groupChat/team scopes, but `SignalTunerTeamsBot` has no message/help/welcome handler. | Implement suitable onboarding/help/fallback for the advertised scopes, or approve a narrower accurate capability declaration after consumer review. Do not silently remove released scopes. Backend and manifest owners. |
| Verify | `supportsChannelFeatures: tier1`, team scope and meeting contexts are declared, while invitations explicitly reject channel meetings. | Test every declared install surface. Explain supported limits in first-run UI and listing; resolve mismatches before submission. Current 1.27 schema meets the documented 1.25 minimum for channel-enabled submissions; still validate against Microsoft's current schema. |
| Verify | Production host `teams.signaltuner.com` appears in CORS/manifest, but a public fetch of `/tabs/home` was inaccessible during review. | Verify DNS, certificate, hosting resource, route refreshes and iframe headers; inability to fetch is not proof the service is down. Check `/tabs/config`, `/tabs/privacy`, `/tabs/terms`, `/tabs/support` without relying on localhost. |
| Verify | Desktop setup offers a macOS DMG although the supplied project map documents the Windows client. | Verify the downloadable Mac client and supported features; otherwise correct the download flow/support wording before release. Test mobile/no-client users can continue without syncing. If a maintained Mac project exists, add it to the shared project map separately. |
| Verify | External navigation/download actions exist throughout Settings and client setup. | Clearly mark leaving Teams; warn before executable downloads; test keyboard, focus, high contrast, dark mode and narrow meeting/mobile layouts. |
| Verify | Publisher identity, support SLA/contact, policy publication and current deployment evidence are not independently verified. | Confirm the exact Partner Center verified publisher, reachable public support/privacy/terms pages, current data handling disclosures, reviewer accounts and feature deployments. |

This review does not change authentication, permissions, scopes, API routes or
billing rules. Those need a coordinated implementation/release after the specific
public-tenant design is approved. Cosmetic readiness cannot resolve these blockers.

## Production configuration and Azure locations

All values in `.env.prod` are non-secret, Teams-project configuration. This is a
separate environment, not a replacement for `.env.local`.

| Key | Source and purpose | Default |
| --- | --- | --- |
| `TEAMSFX_ENV` | `prod`; toolkit environment and preflight | No production fallback |
| `TEAMS_APP_ID` | Confirmed production app ID from Teams Developer Portal; consumed by the manifest/toolkit | Blank; do not substitute bot ID or assume local ID |
| `TAB_DOMAIN` | Confirmed production tab DNS hostname; manifest valid domains | Blank |
| `TAB_ENDPOINT` | Matching HTTPS origin, no trailing slash; manifest routes | Blank |
| `TEAMS_SSO_RESOURCE` | Entra exposed API URI matching that origin and bot identity | Blank |
| `AZURE_APP_SERVICE_RESOURCE_ID` | Full Azure resource ID of the production Teams tab host; deployment workflow | Blank; required before deployment, not manifest validation |

1. **Teams tab hosting:** Azure portal > App Services > the app serving the
   confirmed tab hostname > Settings > Custom domains and Certificates. The
   resource name/slot is not established by repository evidence. Identify it by
   its custom-domain binding; do not select `signaltuner-back` or its staging slot
   as the tab frontend. Verify hostname and TLS; obtain the resource ID from
   Overview > JSON View. The existing Bicep can create a paid B1 service and outputs
   the default Azure hostname; do not run Provision as a harmless packaging step.
   Production resource creation/domain changes need separate approval. Build and
   deploy to the confirmed slot, then test route and iframe loading before packaging.
2. **Entra SSO:** Azure portal > Microsoft Entra ID > App registrations > All
   applications > `signaltuner-teams-bot`
   (`e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`) > Expose an API. Once the host and
   identity design are approved, add the hosted URI
   `api://<confirmed-tab-host>/botid-e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`, if that
   identity is retained, and set the identical `TEAMS_SSO_RESOURCE`. Retain local
   development compatibility. Verify `access_as_user` and existing Teams client
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
# Once .env.prod contains confirmed values:
npm run validate:store -- -RequireProduction -EnvironmentFile env/.env.prod
```

The preflight validates branding, icon dimensions/pixels and, when supplied,
production inputs, origin/resource alignment, unresolved placeholders and
development hosts. It renders only in memory. It does not validate Microsoft's
full JSON schema, ownership, live endpoints, existing app/version uniqueness,
tenant consent or certification policy. Source-only success is not release approval.
The empty production example deliberately fails production validation.

Validation completed during this review: source preflight, local manifest
`-ValidateOnly`, full `npm run build`, frontend TypeScript check and Git whitespace
check passed. Preflight exercised six cases: source-only success; missing, blank
and local production-input rejection; synthetic valid inputs; and mismatched SSO
resource rejection. Both icons were visually inspected, including the white icon
against a dark background. Build emitted the existing Vite CommonJS API deprecation
warning. No Microsoft schema/package validator or live Teams/browser test was run;
the production environment remains incomplete. Existing local ZIP/manifest files
were not regenerated.

After explicit authorization to package, generate the production manifest and
ZIP containing `manifest.json`, `color.png` and `outline.png` at its root. Do not
zip the repository or the entire `appPackage` directory. Retain the correct app
identity and increment the manifest version if 1.0.16 has already been distributed.
Run Microsoft schema/package validation and Developer Portal Store validation
against that exact ZIP. Resolve all errors and assess every warning.

`m365agents.yml` **publish** calls `teamsApp/publishAppPackage`, which submits to
the organization's Teams Admin Center. It is not a public Store submission.
Provision also deploys infrastructure and updates Developer Portal. Neither is
a read-only validation command.

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
- Contract or behavior affected: public-tenant SSO, bot onboarding/invitations and
  production app deep links.
- Required code/configuration changes: approve and implement customer-tenant
  identity/consent handling; resolve invitation access and conversational bot gaps.
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

- Changed projects: Teams app only; existing uncommitted dashboard work preserved.
- Dependent projects: backend and installed desktop client for advertised workflows;
  website for external account flows. No sibling code changed in this preparation.
- Database/configuration prerequisites: no migration or schema change; confirmed
  production environment and resolved public-tenant design still required.
- Backward compatibility: branding changes are additive; app/bot IDs, scopes, routes
  and authentication behavior unchanged. Only source manifest version advanced.
- Required deployment order: resolve blockers, backend/configuration if changed,
  hosted frontend, production package and installation checks, Store submission.
- Post-deployment verification: cross-tenant sign-in, all install scopes, bot
  invitations, consent/credits, telemetry, reports, mobile commerce and legal/support.
- Rollback considerations: retain existing identities and security checks. Restore
  earlier branding through a new higher manifest version if an update is required;
  frontend and backend compatibility must be assessed separately.
- Projects still requiring deployment: hosted Teams frontend and package; backend
  will require deployment for the unresolved bot/public-tenant work. This task did
  not deploy any component. Functionality may be inconsistent until required
  components and configuration are released together in the stated order.

## Official references

- [Teams icon requirements](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/design/design-teams-app-icon-store-appbar)
- [Submission checklist](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/prepare/submission-checklist)
- [Store validation guidelines](https://learn.microsoft.com/en-us/microsoftteams/platform/publishing/office-store-approval)
- [Package contents](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/apps-package)
- [Publishing overview](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-publish-overview)
