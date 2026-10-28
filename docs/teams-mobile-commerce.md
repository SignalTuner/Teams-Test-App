# Teams host-specific commerce presentation

Implemented 2026-10-27 in the Teams frontend. The backend remains authoritative
for account access, report allowance and analysis authorization. This is a UI
policy, not an authorization boundary or a change to billing/credit contracts.

## Host detection

`teamsHost.ts` shares TeamsJS `app.initialize()` / `app.getContext()` with the app.
`TeamsHostProvider` allows external commerce only when `context.app.host.name`
is `Teams` or `TeamsModern` and `clientType` is `desktop`, `web` or `macos`.
Android, iOS and iPadOS are restricted, as are unsupported, missing, failed or
pending contexts. Links start hidden, so slow initialization cannot flash them.
Neither viewport width nor user agent enables commerce. Desktop installer
selection still uses user agent, only after the host policy permits that link.

## User experience

| Surface | Mobile / unknown host | Confirmed Teams desktop / web |
| --- | --- | --- |
| Settings | Server `availableReports` total and existing account/profile/password functions; no plan, trial, payment, billing, buy or upgrade UI | Existing plan/balance and external purchase/billing links |
| Analysis | Existing entitled reports remain available; zero allowance has neutral wording and Test Analysis remains available on supported connected desktops | Same entitlement and consumption rules |
| Account deletion | Opens existing in-app Support, without navigating to website commerce | External account-deletion page |
| Team administration, service history and provider reports | External links hidden; in-app status/incident content retained | External links retained |
| Desktop setup | Activation, refresh and continue controls remain; no website download/launch links | Download/launch links identify leaving Teams and installer download |
| Registration / profile consent | No purchase prompt; internal legal pages remain accessible | Payment disclosure appears before the agreement checkbox |

Website pages such as `/team`, `/teamsstatus` and `/account/delete` expose shared
website navigation to plans, so they are treated as indirect commerce routes.
Provider report destinations are also hidden conservatively on mobile rather
than assuming every external destination has commerce-free navigation.
The pending 1.0.17 manifest description also directs account creation into the
app instead of promoting a visit to the website. Required publisher website
metadata and existing identifiers/scopes are unchanged by this policy.

`ExternalWebsiteLink` includes a visible external-link arrow, “outside Teams”
text, accessible label, new-window target and `noopener noreferrer`. Settings
and consent explain external Stripe checkout and payment-details handling.
The Teams terms include a product-specific payment-flow clarification in section
7 (last updated October 27); existing effective date and general legal text remain.
Legal pages contain no commerce links or actionable purchase UI. This clarification
does not change website billing implementation or imply Marketplace integration.

## Validation and live checks

Run `npm run test:commerce` for host policy and initial-render checks and
`npm run test:personal-browser` for synthetic Teams/API browser scenarios. The
browser harness uses Playwright supplied through `PLAYWRIGHT_MODULE` and optionally
`BROWSER_EXECUTABLE`. Start Vite on localhost port 4317 with
`VITE_SIGNALTUNER_API_URL=http://127.0.0.1:4317/backend`; no real backend is called.
No Playwright runtime dependency is added to the application.

Browser scenarios cover iOS/Android/iPadOS at wide widths, desktop/web/macOS at
phone width, missing/failing/delayed host context, trial/payment-bearing account
fixtures, available allowance, entitled reports, zero-allowance Test Analysis,
in-app deletion support, setup, legal pages and consent disclosure. Existing
personal-dashboard, SSO-pause, meeting, settings, service-status and live-analysis
regressions remain applicable. These tests mock the host and are not live Teams
certification evidence.

Final local validation passed: 50 frontend unit/render/contract tests, all 25
browser scenarios in headless Edge, full `npm run build`, TypeScript, source
Store readiness and local manifest `-ValidateOnly`. Microsoft schema 1.27 also
passed for the final 1.0.17 description using local and synthetic hosted inputs
rendered in memory. Existing Vite CommonJS deprecation and Microsoft Draft-04
`$ref` sibling warnings are non-failing. Production readiness was not asserted.

Before submission, verify on actual Teams iOS, Android and iPadOS: sign-in/signup,
legal/support navigation, positive and zero allowance, personal and meeting
reports, Test Analysis on a supported connected desktop, account deletion support
and all incident/setup routes. On desktop/web verify external-navigation labels,
pre-consent disclosure and the actual website/Stripe flow. Re-audit new external
links and website navigation for indirect commerce whenever those flows change.
No purchase, support message, live analysis or invitation was sent by these tests.

## Release impact

- Changed project: Teams frontend only. Backend, website, WPF and SEO contracts
  and code are unchanged by this host policy.
- Prerequisites: existing backend entitlement and Test Analysis support; no new
  SQL, permission, environment key or Azure setting. The description-only wording
  cleanup is included in the already pending 1.0.17 manifest, not an installed package.
- Release: build and upload the static Teams frontend to DreamHost when authorized,
  following [the hosting guide](dreamhost-hosting.md). No Azure frontend is needed.
- Compatibility: independently releasable UI changes. Keep backend authorization
  intact. Rolling back this frontend can restore prohibited mobile links, so it
  requires a commerce audit before any Store release.
- No deployment, Azure modification, package generation/installation or Store
  submission was performed. The separate manifest 1.0.17 and bot/public-tenant
  verification work in `store-submission.md` still applies.

## Microsoft references

- [Teams Store financial information and mobile experience requirements](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/prepare/teams-store-validation-guidelines)
- [TeamsJS HostClientType values](https://learn.microsoft.com/en-us/javascript/api/@microsoft/teams-js/hostclienttype?view=msteams-client-js-latest)
- [TeamsJS host context](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/access-teams-context)
