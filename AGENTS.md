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
| SignalTuner Microsoft Teams app | `C:\Users\georg\source\repos\Teams Test App` | Teams tab/meeting experience, Teams context, Teams SSO initiation, manifests, and Teams-specific UI | Microsoft 365/Teams Agents Toolkit project; inspect `package.json` and manifests for the exact runtime |

The project map may expand. When another SignalTuner client, service, website, integration, deployment project, or shared package is discovered, evaluate whether it belongs in this map and tell the user when the shared documentation should be updated.

### Local path maintenance

The paths above are operational development metadata and must remain synchronized across all four `AGENTS.md` files and `SignalTuner.code-workspace`.

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
| Windows telemetry collection and desktop behavior | SignalTuner WPF local client |
| Browser-specific UI and routing | SignalTuner website frontend |
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
    - Update all four `AGENTS.md` files only when shared database instructions, paths, ownership, or conventions changed.

Never include production rows, personal data, credentials, hostnames, GTID state, or full data dumps in the repository. Do not commit the supplied source dump. Keep migration scripts and `schema.sql` structure-only.

If the backend/database repository is unavailable while working in another project, document the exact required migration and backend updates as cross-project follow-up rather than creating an unofficial schema copy.

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
- Do not assume the WPF client, website, Teams app, and backend can be deployed simultaneously.
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

Shared changes must be reflected consistently in all four `AGENTS.md` files. Project-specific changes belong only in the affected repository's project section unless they alter an ecosystem rule.

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
