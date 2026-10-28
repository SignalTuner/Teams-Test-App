import * as teamsJs from "@microsoft/teams-js";

let contextPromise: Promise<teamsJs.app.Context> | null = null;

export function initializeTeams(): Promise<teamsJs.app.Context> {
  if (!contextPromise) contextPromise = teamsJs.app.initialize().then(() => teamsJs.app.getContext());
  return contextPromise;
}

// Presentation only. The backend remains authoritative for access and allowances.
// Unknown/failed context stays restricted, including while Teams initializes.
export function allowsExternalCommerce(context: { app?: { host?: { clientType?: string; name?: string } } } | null): boolean {
  const host = context?.app?.host;
  return Boolean(host && ["Teams", "TeamsModern"].includes(host.name || "")
    && ["desktop", "web", "macos"].includes(host.clientType || ""));
}
