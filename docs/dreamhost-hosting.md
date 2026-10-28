# SignalTuner Teams App: DreamHost static hosting

Prepared 2026-10-27 for `https://teams.signaltuner.com` on standard DreamHost
Apache hosting. The subdomain already exists per the user. No live upload,
DNS/hosting/Azure/Entra change, database operation or Teams package generation
is part of this preparation.

## Architecture and preserved behavior

DreamHost serves only the React app, assets and Apache configuration. Node.js is
a development/build tool and is not needed on DreamHost. The existing Azure
backend remains `https://signaltuner-back.azurewebsites.net`; bot messaging remains
`https://signaltuner-back.azurewebsites.net/api/messages`. Bot and Entra application
ID stays `e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`. All manifest IDs, permissions,
notification-only/groupChat declarations and template substitutions are retained.

TeamsJS initialization, meeting context, themes, personal/meeting dashboards,
activation, telemetry, analysis, service health/history, invitation dialog,
settings and support are unchanged. SSO implementation/configuration is retained
but remains **paused** by the previously requested `TEAMS_SSO_ENABLED = false`.
This hosting work does not re-enable it or implement public-tenant SSO. Existing
email/password and bearer-session behavior remains in use. Mobile commerce
restrictions remain active.

## Build and upload contents

Use Node.js 20+ and the existing lockfile. From the renamed checkout:

```powershell
Set-Location 'C:\Users\georg\source\repos\SignalTuner Teams App'
npm ci
npm run build:dreamhost
npm run validate:dreamhost
npm run test:dreamhost
npx --no-install tsc -p tsconfig.app.json --noEmit
```

`npm ci` is needed for a fresh checkout or changed lockfile, not every update.
The output is `dist/dreamhost/`: `index.html`, `.htaccess`, and `assets/` with
content-hashed JS, CSS and images. Upload **the contents**, not a parent `dist`
folder. No source, dependencies, Node entry point, configuration credentials,
certificates, source maps or Teams ZIP belongs on DreamHost.

`build:dreamhost` is also the repeatable SFTP preparation script: it builds and
validates the directory without needing credentials or opening any connection.
It cleans only `dist/dreamhost`, refuses redirected output paths, uses `/` as
Vite's asset base and copies `hosting/dreamhost/.htaccess` automatically.
`npm run build` retains the local Node build and cleans `dist`, so run
`build:dreamhost` **after** a local full build when preparing upload files.

## Configuration ownership

| Configuration | Owner / consumer | Value and handling |
| --- | --- | --- |
| `config/dreamhost/.env.dreamhost` / `VITE_SIGNALTUNER_API_URL` | Teams Vite production build | `https://signaltuner-back.azurewebsites.net`; public, not a secret |
| Root `.env.local`, `.localConfigs`, `env/.env.local` | Existing local Vite/Node/Toolkit workflow | Remain separate; not copied or loaded by DreamHost mode |
| `env/.env.prod` | Teams manifest preparation only | Copy the non-secret example and confirm the existing production app ID; not needed to build static files |
| `TAB_DOMAIN` | Production manifest valid domains | `teams.signaltuner.com` |
| `TAB_ENDPOINT` | Production manifest URLs | `https://teams.signaltuner.com` without trailing slash |
| `TEAMS_SSO_RESOURCE` | Production manifest/Entra exposed API alignment | `api://teams.signaltuner.com/botid-e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`; verify in Entra before use |

DreamHost mode loads a separate env directory, exposes no automatic `VITE_*`
variables and explicitly substitutes only the approved API URL. A conflicting
shell API override fails the build. No secret belongs in any frontend variable.
No runtime `.env` file is uploaded or required. Endpoint changes require a
coordinated code/configuration review, not a DreamHost environment setting.

`m365agents.yml` cannot provision Azure or create/update/publish a Teams app:
its provision/publish actions fail with guidance; deploy only prepares static
files. `m365agents.local.yml`, `src/index.ts`, local certificates and VS Code
debug tasks are retained. Legacy `infra/`, `web.config` and `.webappignore` remain
inert historical tooling; do not run them or upload them to DreamHost.

## SFTP upload and HTTPS

1. In DreamHost **Websites > Manage Websites > teams.signaltuner.com > Manage**,
   inspect hosting settings and record the assigned website user and **web
   directory**. It is commonly `/home/<website-user>/teams.signaltuner.com`, but
   use the actual configured directory, not the account home directory or an
   assumed `public_html`. Confirm this is Apache hosting, not DreamPress/Nginx.
2. Verify DNS points to the assigned DreamHost host. In the site's SSL/Security
   settings, enable/verify the certificate for `teams.signaltuner.com` (a valid
   DreamHost/Let's Encrypt certificate is sufficient). Wait for HTTPS to work
   before enabling the app for users. The upload includes an HTTP-to-HTTPS 302
   redirect to the fixed hostname, with an ACME challenge exception.
3. Use FileZilla, WinSCP or another SFTP client with the hostname and website-user
   credentials from DreamHost's SFTP access panel, **port 22**. Keep credentials
   in the client/approved credential store, never the repo or build directory.
4. Back up existing site files, including hidden `.htaccess`, outside the public
   web directory. Upload all new `assets/` first, then `.htaccess`, then replace
   `index.html` last (upload as a temporary file and rename if supported). Do not
   replace the entire site directory or remove `.well-known`/certificate files.
   Ensure hidden-file transfer is enabled. Typical permissions are 755 for
   directories and 644 for files, subject to the hosting account's policy.
5. Keep previous hashed assets while tabs may still run older HTML. Archive each
   tested upload directory outside the served root. To roll back, restore that
   release's assets and matching `.htaccess`, then its `index.html` last. Do not
   roll back backend authentication or change identities as part of a static rollback.

DreamHost applies `.htaccess` immediately; no Node process or Apache restart is
needed. Parent-directory or hosting-layer headers can still affect the final
response. Resolve a conflicting upstream `X-Frame-Options` or second restrictive
CSP at that layer; do not broaden the framing policy to `*`.

## Apache routes and security

The file requires Apache 2.4, `mod_rewrite` and `mod_headers`. It serves existing
files normally, disables directory listing/MultiViews, and internally rewrites
extensionless app paths to `index.html` without changing query strings. This
covers `/tabs/home`, `/tabs/config`, `/tabs/login`, `/tabs/create-account`,
`/tabs/support`, `/tabs/privacy`, `/tabs/terms`, plus root `/login`,
`/create-account`, `/forgot-password`, `/support`, `/privacy`, `/terms` and
`/tabs/forgot-password`. Password reset currently requests an email through the
existing backend; no new callback/reset endpoint is invented.

`/assets/...`, `/api/...`, paths with file extensions and private/source files
return real errors when missing; they never receive SPA HTML. API calls go to
Azure, not an Apache proxy. Existing `.well-known` challenge files pass through.
HTML revalidates on each visit; successful content-hashed assets cache immutably
for a year. Missing assets are not given immutable cache headers.

CSP `frame-ancestors` permits self plus Microsoft's documented Teams, cloud,
Microsoft 365 and Outlook hosts, using HTTPS host-specific sources. It does not
allow arbitrary framing. `X-Frame-Options` is unset in both Apache header tables;
`nosniff`, strict-origin referrers, `object-src 'none'` and `base-uri 'self'` remain.
The CSP does not add script/network restrictions that would break TeamsJS or the
existing Azure bearer API. No wildcard CORS, cross-origin isolation or new browser
permission policy is introduced. Installed TeamsJS is 2.52.0, above Microsoft's
2.19 minimum for the current cloud host transition.

## Backend and future hosted SSO

Inspected `signaltuner-back/signaltuner-back/Program.cs`: the named CORS policy
already includes exactly `https://teams.signaltuner.com`, permits the existing
methods/headers and applies before the endpoints. No backend change or deployment
is required for this origin. Source inspection does not prove every running
instance has this version. Verify a real browser OPTIONS/authenticated request
after upload; never relax CORS or token validation to make a failing test pass.

The production SSO resource is preparation guidance, not evidence of live Entra
configuration. Before future re-enablement, use **Azure portal > Microsoft Entra
ID > App registrations > All applications > signaltuner-teams-bot** (app ID above)
**> Expose an API > Application ID URI** and configure/verify:

`api://teams.signaltuner.com/botid-e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999`

Confirm enabled `access_as_user`, token version 2 and existing Teams client
preauthorizations: desktop/mobile `1fec8e78-bce4-4aaf-ab1b-5451cc387264` and web
`5e3ce6c0-2b1f-4285-8d4b-75ee78787346`. Save approved changes and allow propagation,
then obtain fresh tokens through normal sign-in and verify backend issuer/audience
validation. Do not record tokens. No new secret, redirect URI, provider or scope is
required by the static hosting conversion itself.

The documented existing URI is local. Microsoft documents that a Teams SSO app
does not support multiple hosting domains; do not assume adding another URI alone
provides concurrent local/production SSO. Preserve local configuration and resolve
the environment/tenant rollout separately before changing the registration. This
task changes no registration or tenant trust. See `../signaltuner-back/docs/teams-sso.md`.
Public-tenant SSO remains deferred, and public-tenant bot/invitation access remains
a separate Store blocker. Hosting files on DreamHost does not resolve either.

The Azure Bot remains under **Azure portal > Azure Bot > signaltuner-teams-bot >
Configuration**, messaging endpoint unchanged. Backend App Service
`signaltuner-back` (resource group `SignalTuner`) and its settings need no hosting
change. No Azure frontend App Service/resource ID is needed.

## Production manifest and live verification

After the site is uploaded and verified, copy `env/.env.prod.example` to
`env/.env.prod`, supply the **existing approved production Teams app ID** from
Developer Portal, and run the read-only check:

```powershell
npm run validate:store -- -RequireProduction -EnvironmentFile env/.env.prod
```

The production app ID is deliberately unfilled; no local ID or bot ID is guessed.
Keep source `appPackage/manifest.json` substitutions and its pending 1.0.17
version unchanged unless release history requires a later version. Do not submit
the old local ZIP. Package generation/Developer Portal update/installation remain
separate authorized release steps, after the existing Store blockers are resolved.

Live checks before packaging:

- HTTPS certificate/chain and HTTP redirect work without loops. All app routes
  and refreshes return HTML with the expected framing CSP and no conflicting XFO.
- Inspect `curl.exe -I https://teams.signaltuner.com/tabs/home` and privacy/terms/
  config; check `Cache-Control`. Open an actual hashed JS URL and a missing JS URL:
  the former must be JS/200/immutable, the latter 404 and not HTML/200. Inspect
  response headers after every proxy/CDN layer, not just the uploaded file.
- In Teams desktop/web/iOS/Android, verify iframe load, initialization, themes,
  personal and meeting context, email sign-in/signup/reset, same-origin returnUrl,
  activation, telemetry, zero-allowance Test Analysis, entitled reports, history,
  mobile commerce restrictions and support.
- In a supported non-channel meeting verify the existing bot dialog's Load
  Participants, selection, Send Invitation and meeting-chat delivery. Check
  installation scopes and customer-tenant access. No bot flow was changed here.
- Verify backend OPTIONS and authorized API traffic from the hosted origin,
  consent/privacy routes and absence of development URLs in browser requests.

Local validation uses a test-only static server and headless Edge, not the Node
application server. Its routing/headers approximate this `.htaccess`; Apache is
not installed locally, so this does **not** certify Apache directive execution,
DreamHost configuration, TLS, live CORS or a live Teams host. SDK/API mocks in the
broader regression suite are also distinct from live verification. No real account,
purchase, report, support message or invitation is generated by the tests.

Completed local checks (2026-10-27):

- Full local Node/frontend build, TypeScript, DreamHost build and output validator
  passed. The upload directory contains 11 files with root-relative compiled assets,
  `.htaccess`, the approved backend URL and no detected source/development/secret
  markers. No source maps, env files, certificates or dependencies are included.
- All 50 existing frontend unit/render/contract tests and 25 existing browser
  scenarios passed. Six static/environment/routing tests and two browser tests of
  the actual compiled DreamHost bundle passed, including direct routes, refreshes,
  query-preserving returnUrl and bearer API use with synthetic responses.
- Existing local Node HTTPS server started successfully; all seven registered tab
  routes returned React HTML after normal trailing-slash redirects. The localhost
  smoke check alone bypassed local certificate trust; production checks must not.
  No local Teams package was created or installed. Test servers were stopped.
- Source Store readiness and local manifest `-ValidateOnly` passed. Actual
  production identity/Entra/live-host readiness remains unverified. Existing local
  ZIP and generated manifest remained unchanged (last modified October 26).
- Rename helper passed a synthetic five-project workspace test, duplicate-target
  refusal, identity preservation and rerun recovery. Actual `-ValidateOnly` passed;
  actual rename was blocked by a Windows open-folder lock. All current path maps
  remain accurate until that manual completion step.
- Existing Vite CommonJS API and Node experimental warnings are non-failing. The
  negative stale-API test intentionally logs a Vite configuration-load rejection.

Browser tests require Playwright test tooling and a browser, not application
runtime dependencies. This session used an isolated temporary Playwright 1.56.1
installation with headless Edge. Set `PLAYWRIGHT_MODULE` to that installation and
`BROWSER_EXECUTABLE` to your Edge executable, then run
`npm run test:dreamhost-browser`. The existing `test:personal-browser` additionally
requires the localhost Vite/synthetic-API setup documented in `teams-mobile-commerce.md`.

## Rename and workspace

Canonical target: `C:\Users\georg\source\repos\SignalTuner Teams App`.
**Directory rename pending:** Windows refused the rename because another process
holds the checkout open. Current directory remains
`C:\Users\georg\source\repos\Teams Test App`; all five project maps and the shared
workspace intentionally still point there. Internal project/npm naming and static
hosting preparation are complete. The upload directory currently resides under
that existing checkout and will move intact with it.
Internal npm name: `signaltuner-teams-app`; public Teams name: **SignalTuner**.
The five shared `AGENTS.md` maps and `SignalTuner.code-workspace` must use the
canonical path together. GitHub repository and Git remote stay unchanged.

Close VS Code and terminals/debuggers using the old folder. In a separate
PowerShell window, run the prepared helper from its parent:

```powershell
Set-Location 'C:\Users\georg\source\repos'
powershell -NoProfile -ExecutionPolicy Bypass -File '.\Teams Test App\scripts\complete-project-rename.ps1'
code 'C:\Users\georg\source\repos\SignalTuner.code-workspace'
```

The helper checks the exact source/target, renames the folder, updates all five
AGENTS maps and the workspace together, and checks Git configuration, manifest,
local env and debugger files remain byte-for-byte intact. `-ValidateOnly` performs
the preflight without changes. If a metadata write fails after the move, rerun
the helper from `SignalTuner Teams App\scripts` to finish synchronization; it
refuses to merge two checkouts. Do not discard uncommitted changes or manually
rename the remote repository. Reopen the shared workspace after success.

## Deployment impact

- Changed projects: Teams app only so far. The prepared rename helper updates
  all five AGENTS maps and multi-root workspace when the folder lock is cleared.
  No sibling application code change is required.
- Dependent projects: Azure backend/bot and connected desktop provide existing
  APIs/telemetry; website remains the external account/payment destination.
- Database/configuration prerequisites: no SQL or Azure setting change for this
  hosting conversion. DreamHost HTTPS, web-directory and header checks remain.
- Backward compatibility: unchanged identities, APIs, auth and local debugging.
  A new origin has separate browser storage; users may need to sign in again.
- Required order: upload/verify static host, confirm production manifest identity
  and deferred SSO/Store prerequisites, then separately authorize packaging and
  installation. No backend release is needed for the current CORS configuration.
- Rollback: restore matching static release files; retain old hashed assets and
  authorization safeguards. A manifest already distributed requires an appropriate
  higher-version rollback package, never silently switching IDs or permissions.
- Still requiring deployment: user upload to DreamHost and later Teams release;
  unrelated customer-tenant/bot blockers remain in the Store checklist.

## Sources

- [Microsoft tab hosting and frame-ancestor requirements](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/tab-requirements)
- [Microsoft tab SSO registration and resource URI](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/tab-sso-register-aad)
- [DreamHost .htaccess support](https://help.dreamhost.com/hc/en-us/articles/216456227--htaccess-overview)
- [DreamHost response headers](https://help.dreamhost.com/hc/en-us/articles/215747598-Setting-headers-with-an-htaccess-file)
- [DreamHost SFTP and port 22](https://help.dreamhost.com/hc/en-us/articles/115001051531-FTP-and-SFTP-at-DreamHost)
- [DreamHost website upload directory](https://help.dreamhost.com/hc/en-us/articles/214920778-Where-do-I-upload-my-files-once-connected-via-FTP)
- [Apache header tables and conditional headers](https://httpd.apache.org/docs/2.4/mod/mod_headers.html)
