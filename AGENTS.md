# AGENTS.md — SignalTuner Microsoft Teams app

These instructions apply to this repository and all files beneath it unless a more specific nested `AGENTS.md` overrides them.

<!-- BEGIN SIGNALTUNER SHARED INSTRUCTIONS -->
## SignalTuner ecosystem mandate

This repository is one component of the broader **SignalTuner product ecosystem**. Do not treat it as an isolated application. Before changing behavior, contracts, configuration, dependencies, or deployment assumptions, consider every SignalTuner producer and consumer that may be affected.

The current project map is:

| Component | Current local path | Product responsibility | Primary technology / source of truth |
|---|---|---|---|
| SignalTuner WPF local client | `C:\Users\georg\source\repos\SignalTuner` | Windows endpoint telemetry collection, local diagnostics, activation, and desktop behavior | WPF on .NET Framework 4.7.2; this repository owns Windows-client behavior |
| SignalTuner backend | `C:\Users\georg\source\repos\signaltuner-back` | APIs, authentication enforcement, database access, server-side business rules, background services, and SignalR | .NET 6 Web API with MySQL; this repository owns API and database implementation |
| SignalTuner website frontend | `C:\Users\georg\source\repos\signaltuner-front` | Public website and browser dashboard | React SPA; this repository owns browser-specific presentation and interaction |
| SignalTuner service-status SEO frontend | `C:\Users\georg\source\repos\is-service-down` | SEO-friendly public service-status/front-door experience that should share style, messaging patterns, and browser functionality with the website where applicable | React SPA / Create React App; this repository owns SEO-oriented public status pages and domain-specific frontend behavior |
| SignalTuner Microsoft Teams app | `C:\Users\georg\source\repos\Teams Test App` | Teams tab/meeting experience, Teams context, Teams SSO initiation, manifests, and Teams-specific UI | Microsoft 365/Teams Agents Toolkit project; inspect `package.json` and manifests for the exact runtime |

The project map may expand. When another SignalTuner client, service, website, integration, deployment project, or shared package is discovered, evaluate whether it belongs in this map and tell the user when the shared documentation should be updated.

### Local path maintenance

The paths above are operational development metadata and must remain synchronized across all five `AGENTS.md` files and `SignalTuner.code-workspace`.

When any SignalTuner project is moved, renamed, cloned to a new canonical location, or otherwise changes its local directory:

1. Update the project map in **every** SignalTuner `AGENTS.md`.
2. Update all absolute path references elsewhere in those files.
3. Update `C:\Users\georg\source\repos\SignalTuner.code-workspace`.
4. Search all SignalTuner repositories for the old path and update relevant scripts, documentation, launch profiles, and tooling.
5. Report which files were updated and whether any inaccessible project still contains the old path.
6. Do not consider the relocation complete until all shared documentation and workspace references are aligned.

Do not update only the currently open repository when shared path information changes.

## Work across repositories

Use the multi-root workspace whenever practical:

`C:\Users\georg\source\repos\SignalTuner.code-workspace`

Before implementing a change, identify:

- The component that owns the behavior.
- Every known producer and consumer of the affected contract.
- Tests, configuration, documentation, and deployment artifacts tied to that contract.
- Whether independently deployed or older clients require backward compatibility.
- Database, authentication, telemetry, privacy, billing, or operational consequences.

When sibling repositories are available in the active workspace, inspect and update them in the same task where the requested change requires alignment. Use workspace-wide search rather than assuming a contract has only one consumer.

When an affected sibling project is not accessible:

- Complete all safe work available in the current repository.
- Do not pretend the ecosystem is fully aligned.
- Report the exact repository, likely files or search terms, contract change, and required follow-up.
- Tell the user which additional project must be opened or updated.
- Include deployment implications.

Use this reporting format when cross-project work remains:

```text
Cross-project follow-up required
- Project:
- Contract or behavior affected:
- Required code/configuration changes:
- Suggested search terms or likely files:
- Compatibility risk:
- Required deployment action/order:
```

## Shared-contract policy

Treat the following as ecosystem contracts rather than local implementation details:

- API routes, versions, HTTP methods, status codes, headers, and query parameters.
- Request/response DTOs, JSON names, data types, nullability, enums, and default values.
- Authentication, authorization, OAuth, Microsoft Entra, Teams SSO, session, cookie, token, and logout behavior.
- SignalR hub routes, events, payloads, reconnect behavior, and connection identity.
- Database tables, columns, keys, constraints, indexes, relationships, defaults, and status values.
- Environment variables, configuration keys, feature flags, public URLs, callback URLs, origins, and local ports.
- Teams manifests, app IDs, resource/application IDs, valid domains, content URLs, and permission declarations.
- Telemetry payloads, units, sampling intervals, thresholds, timestamps, and interpretation.
- Error identifiers, retry behavior, idempotency rules, and user-visible failure semantics.
- Billing, subscription, entitlement, plan, quantity, and credit semantics.
- Shared domain terminology.

Before changing a shared contract:

1. Search all accessible SignalTuner repositories for exact names and semantic equivalents.
2. Identify the authoritative owner.
3. Prefer additive, backward-compatible changes.
4. Update producers, consumers, tests, examples, configuration, and documentation together.
5. Preserve transitional compatibility when releases cannot be deployed atomically.
6. State whether the change is breaking.
7. Provide the required deployment order and rollback considerations.

For example, moving or renaming an endpoint is not complete when only the backend compiles. Every calling client, environment configuration, integration test, mock, manifest, and operational document must either be updated or explicitly reported as outstanding.

## Source-of-truth ownership

Use these ownership rules:

| Concern | Authoritative source |
|---|---|
| API implementation, server-side authorization, shared server business rules | SignalTuner backend |
| Canonical MySQL schema and database migrations | `C:\Users\georg\source\repos\signaltuner-back\database` |
| Personal service monitoring selections | Backend `UserServiceSelections`, shared by the website and WPF for the same account; legacy team `ServiceSelections` remain separate |
| Windows telemetry collection and desktop behavior | SignalTuner WPF local client |
| Browser-specific UI and routing | SignalTuner website frontend |
| SEO-friendly public status/front-door pages and domain-specific public website variants | SignalTuner service-status SEO frontend, aligned with SignalTuner website frontend style and shared browser functionality |
| Teams context, Teams manifests, Teams-specific UI, and Teams SSO initiation | SignalTuner Microsoft Teams app |
| Deployed secrets and credentials | Approved secret/configuration stores; never source files or `AGENTS.md` |
| Exact build/test commands | The relevant repository's solution, project, and package manifests |
| Production configuration | The relevant deployment platform and checked-in non-secret templates |

A consumer must not silently redefine an owner-controlled contract. When local behavior must diverge, document why and keep the divergence explicit.

## Database documentation and migration policy

The backend repository owns the canonical schema:

`C:\Users\georg\source\repos\signaltuner-back\database\schema.sql`

Related documentation and migrations belong under:

```text
signaltuner-back\database\
  schema.sql
  README.md
  SCHEMA_CHANGELOG.md
  migrations\
    pending\
    applied\
```

All projects must treat `schema.sql` as the documented state of the **confirmed live database**, not as a scratch design file.

### Current confirmed database notes

As of 2026-07-31, the user confirmed execution of the July 25 Teams meeting
migrations. The confirmed live schema includes:

- `TeamsMeetingRosterParticipants` for Teams meeting roster cache entries,
  keyed by `meeting_id` and `participant_cache_key`.
- `TeamsMeetingSessions.organizer_m365_object_id` and
  `TeamsMeetingSessions.organizer_tenant_id`, indexed by
  `ix_teams_meeting_sessions_organizer`.
- `TeamsMeetingParticipants.meeting_role` for the signed-in participant's
  resolved Teams meeting role.

These fields are additive contracts shared by the SignalTuner backend and
SignalTuner Microsoft Teams app. Keep future meeting roster, organizer, and role
changes aligned across both projects.

### Required database-change workflow

When a database change is requested:

1. Inspect `database/schema.sql` and affected backend/data-access code.
2. Create a narrowly scoped SQL migration in `database/migrations/pending`.
3. Use a sortable filename such as `YYYYMMDD_HHMM_short_description.sql`.
4. Include comments describing purpose, prerequisites, compatibility, rollback strategy, and affected applications.
5. Make the migration safe and idempotent where practical. Where idempotency is not practical, include explicit precondition/verification queries.
6. Update required backend code, tests, DTOs, queries, and consumers.
7. Do **not** run the migration against the live database unless the user explicitly instructs that production action.
8. Do **not** update `database/schema.sql` to the proposed state merely because the migration file was written.
9. Report the pending migration and the exact production action the user must perform.
10. After the user explicitly confirms that the migration succeeded against the live database:
    - Move the migration from `pending` to `applied`.
    - Update `database/schema.sql` to the confirmed live structure.
    - Add an entry to `database/SCHEMA_CHANGELOG.md`.
    - Re-run relevant tests and schema checks.
    - Update all five `AGENTS.md` files only when shared database instructions, paths, ownership, or conventions changed.

Never include production rows, personal data, credentials, hostnames, GTID state, or full data dumps in the repository. Do not commit the supplied source dump. Keep migration scripts and `schema.sql` structure-only.

If the backend/database repository is unavailable while working in another project, document the exact required migration and backend updates as cross-project follow-up rather than creating an unofficial schema copy.

### MySQL Workbench command compatibility

When generating one-off MySQL cleanup, verification, or test-data commands for the user to run in MySQL Workbench, account for common Workbench and managed-MySQL restrictions:

- If temporary tables are needed, define explicit primary keys instead of using `CREATE TEMPORARY TABLE ... AS SELECT`, because `sql_require_primary_key` may be enabled.
- Do not force `ENGINE=MEMORY`; some environments disable the MEMORY engine. Let MySQL choose the default temporary table engine unless a specific engine is required and confirmed available.
- Start scripts with `DROP TEMPORARY TABLE IF EXISTS ...` for any temp tables they create, because a failed prior Workbench run can leave temp tables in the current session.
- Prefer transactions for multi-table cleanup scripts and order deletes around foreign keys using the canonical `database/schema.sql`.
- MySQL Workbench safe update mode may reject valid joined deletes. For intentional one-off cleanup scripts, include `SET SQL_SAFE_UPDATES = 0;` at the start and `SET SQL_SAFE_UPDATES = 1;` after `COMMIT;`, and explain that this changes only the current Workbench session.
- Do not ask the user to globally disable `sql_require_primary_key`, safe updates, or other safety settings unless there is no session-scoped alternative.

## Reuse of tools, vendors, and patterns

Prefer technologies, vendors, libraries, infrastructure, conventions, and implementation patterns already established in the SignalTuner ecosystem.

Before introducing a new tool or provider:

1. Search sibling projects for an existing solution.
2. Determine whether that solution meets the new functional, security, privacy, compatibility, cost, licensing, and operational requirements.
3. Reuse or extend it when suitable.
4. Introduce a parallel solution only when the established option cannot meet a material requirement.
5. Document the reason for the exception, its operational burden, and whether it should become the new ecosystem standard.

Apply this policy to analytics, authentication, email delivery, logging, error reporting, payments, storage/CDN, feature flags, date/time handling, HTTP clients, serialization, UI libraries, monitoring, alerting, and testing.

Google Analytics is the preferred analytics platform for browser-capable SignalTuner projects when analytics are required and it can satisfy the use case. Do not force browser-only analytics into the WPF client or another environment where it cannot provide essential functionality. In that case, choose the most compatible established alternative or explain why a new approach is necessary.

Avoid duplicating business rules in multiple clients when they can be centralized safely in the backend or a shared package. Do not centralize platform-specific telemetry or UI behavior that properly belongs to a client.

## Backward compatibility and coordinated releases

SignalTuner components may be deployed or installed independently. Design shared-contract changes accordingly.

- Prefer adding before removing.
- Prefer optional fields with safe defaults over immediate renames/removals.
- Keep old API routes or fields during a documented transition when released clients still depend on them.
- Do not assume the WPF client, website, service-status SEO frontend, Teams app, and backend can be deployed simultaneously.
- Explicitly identify minimum compatible versions when compatibility changes.
- For database migrations, distinguish expand, migrate/backfill, and contract/removal phases when appropriate.
- Avoid destructive changes until all known consumers have migrated.
- Include rollback implications and whether the old version remains compatible.

Whenever more than one project is affected, provide a deployment impact report:

```text
Deployment impact
- Changed projects:
- Dependent projects:
- Database/configuration prerequisites:
- Backward compatibility:
- Required deployment order:
- Post-deployment verification:
- Rollback considerations:
- Projects still requiring deployment:
```

Alert the user whenever a change requires deployment or release of a project other than the one currently open. State that functionality may be inconsistent until all required components are deployed. Never deploy, publish, package, submit, or migrate production systems unless the user explicitly asks.

## Environment and configuration discipline

Treat development, test, staging, and production as distinct environments.

- Inspect authoritative configuration files before changing URLs, ports, origins, application IDs, scopes, redirect URIs, manifests, or feature flags.
- Keep local and production Teams manifests/configuration separate.
- Keep environment-specific values out of shared source code where practical.
- Never copy production secrets into development configuration.
- Keep non-secret example configuration current when keys are added or renamed.
- Update CORS, valid domains, redirect URLs, Entra registrations, and client configuration together when an origin or route changes.
- Validate that Teams manifest resource/application IDs and iframe origins match the intended environment.
- Do not assume a value documented here overrides the actual project configuration; investigate discrepancies and update stale documentation.

When adding a configuration key, document its owner, purpose, allowed environments, default behavior, whether it is secret, and every consuming project.

### Azure deployment instructions

Whenever deployment requires a change in Azure, explicitly tell the user where to make it. Include this in the deployment guidance and final response, not only in linked documentation.

- Identify the affected SignalTuner component, Azure service/resource type, and target resource name when known, including the environment and deployment slot when applicable.
- Give the Azure portal navigation path to the relevant setting or operation. Distinguish backend App Service settings from frontend hosting, Microsoft Entra app registrations, Azure Bot resources, and other Azure services.
- State the exact setting/key or operation, the required non-secret value or authoritative source of that value, and whether to add, update, or remove it. Never expose secret values.
- Explain any required save/apply, restart, consent, or redeployment step, its order relative to application deployment, and how to verify the change.
- If the resource name or location is not known, say so and explain how to identify the correct resource; do not guess or give only a vague instruction such as "configure this in Azure."

This is a reporting requirement, not permission to change Azure resources or deploy without the user's authorization.

## Security and privacy

SignalTuner processes user, device, network, account, meeting, and telemetry information. Treat this data as sensitive.

- Never commit API keys, passwords, connection strings, signing keys, tokens, session secrets, certificates, private keys, OAuth client secrets, or magic codes.
- Never place live user records, email addresses, meeting identifiers, device identifiers, or production telemetry in fixtures, examples, logs, screenshots, or committed SQL.
- Never log passwords, authorization headers, cookies, session tokens, one-time codes, complete access tokens, or secret configuration.
- Enforce authorization on the backend. Frontend visibility is not an authorization boundary.
- Validate all externally supplied identifiers and ownership relationships.
- Use parameterized database operations.
- Preserve least privilege for OAuth/Graph/Teams scopes and database accounts.
- Explain security consequences before weakening authentication, authorization, validation, encryption, or privacy controls.
- Redact sensitive values when reporting errors.
- Do not expose internal exception details to end users; preserve actionable server-side diagnostics.

## Observability standards

Prefer consistent, structured, actionable telemetry.

- Use stable event/error identifiers where practical.
- Include correlation/request IDs across client and backend calls when available.
- Record component version, environment, and platform without recording secrets or unnecessary personal data.
- Log enough context to diagnose authentication, Teams integration, API, SignalR, and telemetry failures.
- Avoid duplicate high-frequency logs and unbounded payload logging.
- Use UTC for persisted server timestamps unless a contract explicitly requires otherwise.
- Distinguish transient, validation, authorization, dependency, and internal failures.
- Preserve causal exceptions in server logs while returning safe user-facing responses.
- When adding a background process, define health signals, expected cadence, failure behavior, and alerting needs.

## Dependency policy

Before adding or replacing a dependency:

- Check whether the ecosystem already uses a suitable package or service.
- Prefer maintained, commercially common, well-documented dependencies.
- Minimize production dependencies and avoid overlapping packages.
- Verify framework/runtime compatibility.
- Review licensing, security posture, maintenance activity, bundle/runtime impact, and deployment requirements.
- Explain major new dependencies and provider lock-in.
- Do not upgrade frameworks or replace stable foundational libraries as incidental cleanup.
- Keep lockfiles and package manifests synchronized.
- Remove unused dependencies only after confirming no project, build script, or runtime path relies on them.

## Canonical terminology

Use these terms consistently unless an authoritative existing contract requires a different name:

- **STMSTA**: shorthand for the SignalTuner Microsoft Teams app. Use it only as an abbreviation in discussion or notes; keep formal product names, manifests, identifiers, and user-facing labels unchanged unless explicitly requested.
- **User**: a SignalTuner account holder.
- **Team**: a SignalTuner grouping of users/devices.
- **Device**: an endpoint running a SignalTuner local client.
- **Signal**: the current telemetry/status record associated with a user/device.
- **Diagnostic session**: a bounded period in which enhanced diagnostics are requested or collected.
- **Meeting session**: the SignalTuner server-side record representing a joined meeting context.
- **Meeting participant**: a person represented within a meeting session.
- **Service incident**: a tracked availability/degradation event for an external service.
- **Activation code**: a code linking a local client/device to a SignalTuner user or account context.
- **Entitlement**: permission to use a feature based on account/subscription state.
- **Credit**: a metered unit used by applicable meeting/analysis functionality.

When an existing database/API name differs, preserve compatibility and map it clearly rather than silently creating a competing term.

### Team roles and team-owned report credits (confirmed schema, 2026-10-16)

SignalTuner team roles are Owner, Admin, and Member; they are separate from
Microsoft Teams meeting roles. Owner is the canonical `Teams.team_admin` user.
The backend owns role authorization, invitation approval, credit allocation and
request decisions. Members require per-analysis target approval; same-team Owners
and Admins do not. Credits allocated to a member or the Owner/Admin general fund
belong to the team; unused allocations return to its fund when a member leaves.
See `signaltuner-back/docs/team-roles-and-credits.md` for the API, compatibility,
financial invariants, and coordinated release order. On 2026-10-16 the user
confirmed all pending migrations applied, including
`20261016_1600_team_roles_and_credit_wallets.sql` and
`20261016_1200_remove_low_signal_email_alerts.sql`. Both are recorded under the
backend's `database/migrations/applied`; `schema.sql` reflects the team-role,
wallet/approval/audit additions and removal of retired low-signal alert columns.
No SQL migrations remain pending as of that confirmation. The user separately
confirmed backend deployment and working member credit totals; other client
release and end-to-end verification status remains separate.

### Single desktop sessions and team analysis history (confirmed schema, 2026-10-21)

The backend owns the DS1 per-desktop credential and `sign_out_requested` flag.
Replacement revokes the former desktop before authorizing the new one. Updated
WPF polls the flag and clears local sign-in; an offline desktop does so on reconnect.
Website/Teams must not receive physical device IDs and use server-selected desktop
analysis targets. See `signaltuner-back/docs/desktop-sessions.md` for API contracts,
legacy-client restrictions, validation and rollback.

On 2026-10-21, the user confirmed that both remaining migrations ran successfully
against the live database: `20261021_1200_single_desktop_session.sql` and
`20261021_1800_team_analysis_batches.sql`. Both are recorded unchanged under
the backend's `database/migrations/applied`; `schema.sql` includes
`DesktopClientSessions` and `TeamAnalysisBatches`. No SQL migrations remain
pending at this confirmation. Team analysis history retains the original roster
for full-team completion summaries; see `signaltuner-back/docs/team-analysis.md`.
This records user confirmation, not an independent live database inspection.

Application deployment and end-to-end verification remain separate: deploy all
backend instances before releasing the updated WPF and browser/Teams clients.
For team analysis history, deploy the backend before the website; no additional
WPF/Teams/SEO release is required for that feature. Retain the additive tables
and desktop-session enforcement during rollback as documented.
### Online administrative account deletion (confirmed, 2026-10-22)

The backend owns the online deletion barrier and the administrative purge script.
The user confirmed applying
`signaltuner-back/database/migrations/applied/20261022_1500_online_account_deletion.sql`
and deploying the backend. This records user confirmation, not independent live
inspection. `schema.sql` includes the guard tables, routines and 82 triggers;
no SQL migrations remain pending at this confirmation. Confirm every backend
instance is updated, then explicitly set `@purge_backend_ready=TRUE` in the
private `database/operations/remove_user.sql` copy. Services can remain running
during a purge after that rollout. No client release is needed.
Shared hardware defaults to refusal; the administrative script supports the
explicit `@purge_preserve_shared_devices=TRUE` retention option after diagnostic
review. It preserves peer history and ambiguous device-only logs, detaches only
the target's shared-device mapping transactionally, and keeps account/email
suppression. This is not complete erasure of shared-device history; account
identity collisions still stop the purge. No new migration or deployment is
required for this option. See the backend operations README for limits.
See `signaltuner-back/docs/online-account-deletion.md` for concurrency, privacy,
retained suppression markers, blocked re-registration, external billing, tests
and rollback. Never manually remove guard objects/markers during rollback after a
committed purge: late events could recreate deleted records. Future account-bearing
schema changes must update the guard generator, purge coverage and tests together.
No production migration, deployment or deletion has been performed or confirmed.

### Account lifecycle and free credits (confirmed schema, 2026-10-23)

On 2026-10-23 the user confirmed all pending migrations applied, including
`20261022_1900_account_lifecycle.sql` and `20261023_1200_free_account_credits.sql`.
Both are retained unchanged under the backend's `database/migrations/applied`;
no SQL migrations remain pending at this confirmation. This records user
confirmation, not independent live inspection or application deployment.
`schema.sql` now includes Users creation time and monthly evaluation columns,
AccountRegistrationEmails, AccountCancellationTokens, FreeCreditEligibility,
and the guarded self-deletion routine. Legacy creation timestamps remain unknown.

Signup remains immediate without email verification. Welcome cancellation and
Settings deletion require explicit confirmation. Billing-bearing accounts require
assisted deletion. Email reuse releases only the exact deleted email mapping;
old account/device/billing guards remain. See `signaltuner-back/docs/account-lifecycle.md`.
Verify website routes and all backend instances before enabling
`AccountLifecycle:AllowEmailReuse` (default false). Keep the new legacy Stripe
webhook protection on rollback after any reuse.

The backend owns credit eligibility: first-time normalized emails receive three
personal credits; deleted/reused emails receive zero signup credits. On the first
UTC day of the month, active accounts with exactly zero personal credits receive
one, with durable per-email and per-account evaluation markers. New/recreated
accounts enter monthly eligibility in the following month. Team wallets and paid
monthly allowances are unchanged. Retain hash-only FreeCreditEligibility history
across purges. See `signaltuner-back/docs/free-account-credits.md`.
Deploy every backend instance for this policy and keep a worker running on the
first UTC day. Existing API shapes are unchanged; no additional client release is
required for free credits. Retain eligibility data and signup checks on rollback.
Application deployment and hosted verification remain separate from the database
confirmation; no production SQL or deployment was performed by the agent.

### Teams SSO and meeting permissions (approved configuration, 2026-10-25)

The backend owns validation for both Teams sign-in routes. Tokens must be signed,
valid for the configured application and one explicitly trusted Teams tenant,
and carry the Teams delegated scope/client claims. Client context and email/UPN alone are not identity proof.
The Teams manifest's bot and `webApplicationInfo.id` must align for meeting RSC;
SSO resource URIs are environment-specific. The user approved this correction and
the agent configured the existing Entra bot registration for local tab SSO.
See `signaltuner-back/docs/teams-sso.md` for exact settings, intentional legacy
restrictions, account matching, tests, rollout, and rollback. On 2026-10-25 the
user confirmed publication of the updated backend; deployed binaries were not
independently verified. On 2026-10-26 the user confirmed package 1.0.15 installation;
fresh token diagnostics and Azure logs identify a Teams/bot tenant mismatch and
issuer rejection. The user approved the backend-only `TeamsSso:TenantId` override;
the agent saved and verified the production Azure setting on 2026-10-26. The user
subsequently confirmed SSO works again; all-instance deployment was not independently
inspected. When absent the setting defaults to
`MicrosoftAppTenantId`; an invalid explicit value fails closed. See the SSO document
for exact tenant values and rollout. Bot invitation access remains unresolved.
No database migration is involved. Do not roll back to unsigned-token or UPN-only sign-in handling.

### Teams meeting website projection (implemented, 2026-10-26)

The backend projects existing TeamsMeetingSessions/TeamsMeetingParticipants into
`GET /api/meetings?includeTeams=true` for the same authenticated SignalTuner account.
Teams website IDs use `teams-<session ID>` and retain the numeric meetingSessionId
for analysis. Microsoft Teams is service ID 3. Membership remains historical;
bot-only roster entries do not enroll accounts. Best-effort Teams role labels are
shared display metadata, separate from SignalTuner team roles and authorization.
The first signed-in participant is no longer guessed to be Organizer.
See `signaltuner-back/docs/teams-website-meetings.md` for API compatibility, privacy,
tests and rollback. Deploy backend before website; no SQL migration, Azure setting,
new permission, Teams package, WPF release or SEO release is required for this
feature. Implementation/testing does not confirm deployment or live verification.

Title capture correction: installed TeamsJS meeting-details methods are internal.
The backend now uses the existing bot to recover titles on join and website reads,
using only adapter-authenticated TeamsBotActivityLog routing metadata and matching
meeting/conversation IDs. Do not replace that source with pre-authentication request
logs. This correction requires backend publishing only, with no Azure settings,
SSO changes, migration or Teams package. Missing/purged routing history needs a new
authenticated bot activity (for example Invite > Load Participants). See the same
backend document for bounded retries, privacy, validation and rollback.

## Confirmation gates

Obtain explicit user approval before performing or finalizing any of the following:

- Running a migration against a live database.
- Dropping, truncating, renaming, or irreversibly transforming production data.
- Removing or breaking a public API contract.
- Removing backward compatibility for a released client.
- Changing authentication architecture, identity provider, scopes, or authorization model.
- Adding a paid service or a new vendor with material recurring cost.
- Replacing an ecosystem-standard provider.
- Changing production domains, callback URLs, app IDs, or public endpoints.
- Rotating or invalidating credentials.
- Publishing or deploying any project.
- Submitting or distributing a Teams package or desktop installer.
- Performing bulk changes to live accounts, subscriptions, credits, or telemetry.

Code, migration drafts, deployment plans, and documentation may be prepared without production execution unless the user says otherwise.

## Build, test, and validation rules

Use commands defined by the repository itself. Inspect solution files, project files, `package.json`, lockfiles, CI workflows, and existing documentation before selecting commands.

- Do not invent a test command that the repository does not support.
- Prefer reproducible installs (`npm ci`) when a compatible lockfile exists.
- Build the narrowest affected project first, then run the broader relevant suite.
- Do not report tests as passing unless they were actually run successfully.
- Distinguish not run, unavailable, failed, and passed.
- Do not repair unrelated failures unless necessary; report them separately.
- Avoid broad formatting churn and unrelated refactoring.
- Inspect the final diff for secrets, generated output, accidental data, and unrelated changes.

Definition of done:

1. The requested behavior is implemented in the authoritative owner.
2. Every accessible affected producer and consumer is aligned.
3. Relevant builds, tests, linting, type checks, and packaging checks pass, or limitations are explicitly reported.
4. Shared contracts and configuration are documented.
5. Database changes follow the pending/confirmed workflow.
6. Backward compatibility and deployment order are evaluated.
7. Security/privacy and observability implications are addressed.
8. No production data or secrets are introduced.
9. All affected documentation, including shared `AGENTS.md` path information when applicable, is current.
10. The final response lists changed projects, validation performed, remaining work, and deployment actions.

## Documentation maintenance

Update documentation when changing:

- Project paths or repository names.
- Architecture or source-of-truth ownership.
- Build, test, run, package, or deployment commands.
- API contracts and compatibility windows.
- Database schema or migration procedure.
- Authentication/authorization flows.
- URLs, ports, origins, redirects, app registrations, or manifest requirements.
- Configuration keys and secrets-handling expectations.
- Supported runtimes or framework versions.
- Shared vendors, libraries, and operational providers.
- Domain terminology.

Shared changes must be reflected consistently in all five `AGENTS.md` files. Project-specific changes belong only in the affected repository's project section unless they alter an ecosystem rule.

## Agent completion report

For meaningful changes, finish with:

```text
Work completed
- Current project changes:
- Other SignalTuner projects changed:
- Shared contracts affected:
- Database migration/schema status:
- Validation performed:
- Security/privacy considerations:
- Deployment/release actions:
- Outstanding cross-project work:
```

Do not claim the ecosystem is complete or deployable if required projects remain inaccessible, unmodified, untested, or undeployed.
<!-- END SIGNALTUNER SHARED INSTRUCTIONS -->

## Current repository: SignalTuner Microsoft Teams app

**Path:** `C:\Users\georg\source\repos\Teams Test App`

This repository owns the Microsoft Teams-specific experience: Teams tab/meeting UI, Teams context, local and production manifests, Teams SSO initiation, valid domains/content URLs, and integration with the SignalTuner backend.

### Repository-specific constraints

- Inspect `package.json`, Microsoft 365/Teams Agents Toolkit configuration, environment files, and every manifest variant before changing commands or identifiers.
- Keep local-development and production manifests/configuration separate.
- Ensure iframe/content origins, valid domains, resource/application IDs, redirect URLs, and backend CORS/session settings describe the same environment.
- Do not place Entra client secrets, session-signing keys, certificates, access tokens, or production credentials in source control.
- Do not trust Teams client context as authorization by itself; the backend must validate identity, tenant/account relationships, meeting access, entitlements, and requested actions.
- Request the minimum Graph/Teams permissions required.
- Keep meeting session, participant identity, invitation, analysis, and credit contracts aligned with the backend.
- Handle users who cannot install, access, authenticate to, or view the app without exposing internal errors.
- Do not package, publish, upload, or submit the Teams app unless the user explicitly asks.

### Integration checks

For Teams/authentication changes, review together:

- `manifest*.json` content URLs, valid domains, app IDs, `webApplicationInfo`, permissions, and scopes.
- Local tunnel/localhost origin and production origin.
- Entra app registrations, exposed API scopes, authorized client applications, redirect URIs, and tenant settings.
- Backend CORS, forwarded headers, cookie SameSite/Secure configuration, session storage, signing keys, OAuth callbacks, and SSO token validation.
- Frontend API base URLs and credential/cookie behavior.
- Meeting context parsing, participant identity, meeting-session creation, invitations, analysis calls, and entitlements/credits.

A Teams-origin or SSO change commonly requires coordinated updates to this repository, the backend, and external Entra/Teams configuration. Report all three categories and the deployment/configuration order.

### Build and validation

Use only scripts actually defined in `package.json` and the Agents Toolkit configuration. A typical project may use commands such as:

```powershell
npm ci
npm run build
```

Local Teams debugging commands vary by generated template; do not invent `dev`, `start`, or `dev:teamsfx` scripts without confirming they exist.

Validate manifest generation/validation, local and production environment separation, build/type checks, Teams initialization, authentication/session behavior, API failure states, and meeting-context behavior. Clearly separate code validation from manual Teams/Entra tenant validation.
