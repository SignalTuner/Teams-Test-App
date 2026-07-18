import React from "react";
import * as teamsJs from "@microsoft/teams-js";

import "./App.css";

type AuthProvider = "teams_sso" | "google" | "github" | "email_magic_code";
type ClientDataStatus = "active" | "inactive" | "no_data";
type AnalysisStatus = "Excellent" | "Fair" | "Poor" | "Critical" | "Offline" | "No data" | string;

type CurrentUser = {
  userId: number;
  email: string | null;
  displayName: string | null;
  firstName?: string | null;
  lastName?: string | null;
  credits: number;
  subscriptionPlan?: string | null;
  activationCode?: string | null;
  clientIsActive: boolean;
};

type AuthResponse = {
  token?: string;
  sessionToken?: string;
  jwt?: string;
  jwtToken?: string;
};

type ActivationCodeResponse = {
  activationCode?: string;
  ActivationCode?: string;
};

type TeamsMeetingContext = {
  teamsMeetingId: string;
  teamsConversationId: string | null;
  teamsTenantId: string | null;
  meetingTitle: string | null;
};

type ServiceIncident = {
  incidentId: number;
  title: string;
  status: string;
  impact?: string | null;
  startedAt?: string | null;
};

type TeamsServiceHealth = {
  serviceId: 3;
  serviceName: "Microsoftteams";
  serviceDisplayName: "Microsoft Teams";
  currentStatus: AnalysisStatus;
  unresolvedIncidents: ServiceIncident[];
};

type MeetingParticipant = {
  userId: number;
  displayName: string | null;
  email: string | null;
  authProvider: AuthProvider | string | null;
  joinedAt: string;
  lastSeenAt: string;
  signalScore: number | null;
  clientDataStatus: ClientDataStatus;
  clientIsActive: boolean;
};

type DashboardData = {
  meetingSessionId: number;
  currentUser: CurrentUser;
  teamsServiceHealth: TeamsServiceHealth;
  participants: MeetingParticipant[];
};

type JoinResponse = DashboardData & {
  message?: string;
};

type Issue = {
  field: "signal_bandwidth_status" | "signal_system_status" | "signal_network_status" | string;
  currentValue: string;
  affectedParticipant: string;
  severity: AnalysisStatus;
  recommendation: string;
};

type TelemetryRecord = Record<string, string | number | boolean | null>;

type IndividualAnalysisResult = {
  targetUser: MeetingParticipant;
  telemetry: TelemetryRecord;
  issues: Issue[];
  creditsUsed: number;
  remainingCredits: number;
};

type FullAnalysisResult = {
  analyzedUsers: Array<{
    participant: MeetingParticipant;
    telemetry: TelemetryRecord;
    issues: Issue[];
  }>;
  groupSummary: {
    activeUsersAnalyzed: number;
    usersWithBandwidthIssues: number;
    usersWithSystemIssues: number;
    usersWithNetworkIssues: number;
    teamsHasActiveServiceIncident: boolean;
  };
  creditsUsed: number;
  remainingCredits: number;
};

type AnalysisResult =
  | { mode: "user"; data: IndividualAnalysisResult }
  | { mode: "full"; data: FullAnalysisResult };

type SubscriptionPrompt = {
  requiredCredits: number;
  availableCredits: number;
};

type EmailCodeStep = "request" | "verify";
type JsonRecord = Record<string, unknown>;

const SIGNALTUNER_SESSION_TOKEN_KEY = "signaltunerSessionToken";

function normalizeBaseUrl(baseUrl: string | undefined): string {
  return baseUrl ? baseUrl.replace(/\/+$/, "") : "";
}

function formatValue(value: unknown): string {
  if (value === undefined || value === null || value === "") {
    return "No data";
  }

  if (typeof value === "boolean") {
    return value ? "Yes" : "No";
  }

  return String(value);
}

function getDownloadUrl(): string | null {
  const userAgent = window.navigator.userAgent;

  if (/Windows NT/i.test(userAgent)) {
    return "https://signaltuner.com/update/SignalTuner.exe";
  }

  if (/Mac OS|Macintosh|MacIntel/i.test(userAgent)) {
    return "https://signaltuner.com/update/macos/SignalTuner.dmg";
  }

  return null;
}

function getParticipantName(participant: MeetingParticipant): string {
  return participant.displayName ?? participant.email ?? `User ${participant.userId}`;
}

function getSeverityClass(value: AnalysisStatus | null | undefined): string {
  const normalized = String(value ?? "").toLowerCase();

  if (normalized === "excellent") {
    return "severityGreen";
  }

  if (normalized === "fair") {
    return "severityYellow";
  }

  if (normalized === "poor") {
    return "severityOrange";
  }

  if (normalized === "critical") {
    return "severityRed";
  }

  return "severityBlue";
}

async function fetchJson<T>(url: string, options?: RequestInit): Promise<T> {
  const response = await fetch(url, options);
  const text = await response.text();

  if (!response.ok) {
    const method = options?.method ?? "GET";
    let message = text || `${method} ${url} failed with status ${response.status}`;

    try {
      const parsed = JSON.parse(text) as { message?: string; requiredCredits?: number; availableCredits?: number };
      message = parsed.message ? `${method} ${url} failed with status ${response.status}: ${parsed.message}` : message;
    } catch {
      // Keep the plain response text.
    }

    const error = new Error(message) as Error & { status?: number; body?: string };
    error.status = response.status;
    error.body = text;
    throw error;
  }

  return text ? (JSON.parse(text) as T) : ({} as T);
}

function buildAuthHeaders(token: string | null): HeadersInit {
  return token ? { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } : { "Content-Type": "application/json" };
}

function asRecord(value: unknown): JsonRecord {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as JsonRecord) : {};
}

function readString(record: JsonRecord, ...keys: string[]): string | null {
  for (const key of keys) {
    const value = record[key];

    if (typeof value === "string" && value.trim()) {
      return value;
    }
  }

  return null;
}

function readNumber(record: JsonRecord, fallback: number, ...keys: string[]): number {
  for (const key of keys) {
    const value = record[key];

    if (typeof value === "number" && Number.isFinite(value)) {
      return value;
    }

    if (typeof value === "string" && value.trim() && Number.isFinite(Number(value))) {
      return Number(value);
    }
  }

  return fallback;
}

function readBoolean(record: JsonRecord, fallback: boolean, ...keys: string[]): boolean {
  for (const key of keys) {
    const value = record[key];

    if (typeof value === "boolean") {
      return value;
    }
  }

  return fallback;
}

function readArray(record: JsonRecord, ...keys: string[]): unknown[] {
  for (const key of keys) {
    const value = record[key];

    if (Array.isArray(value)) {
      return value;
    }
  }

  return [];
}

function normalizeCurrentUser(value: unknown): CurrentUser {
  const record = asRecord(value);

  return {
    userId: readNumber(record, 0, "userId", "UserId", "UserID", "user_id"),
    email: readString(record, "email", "Email", "userEmail", "UserEmail", "user_email"),
    displayName: readString(record, "displayName", "DisplayName", "userDisplayName", "UserDisplayName", "user_display_name"),
    firstName: readString(record, "firstName", "FirstName", "userFirstName", "UserFirstName", "user_first_name"),
    lastName: readString(record, "lastName", "LastName", "userLastName", "UserLastName", "user_last_name"),
    credits: readNumber(record, 0, "credits", "Credits", "userCredits", "UserCredits", "user_credits"),
    subscriptionPlan: readString(record, "subscriptionPlan", "SubscriptionPlan", "userStripePlan", "UserStripePlan"),
    activationCode: readString(record, "activationCode", "ActivationCode", "userActivationCode", "UserActivationCode"),
    clientIsActive: readBoolean(record, false, "clientIsActive", "ClientIsActive", "clientIsActivated", "ClientIsActivated"),
  };
}

function normalizeActivationCodeResponse(value: unknown): string | null {
  const record = asRecord(value);
  return readString(record, "activationCode", "ActivationCode", "userActivationCode", "UserActivationCode");
}

function mergeCurrentUser(primary: CurrentUser, fallback: CurrentUser | null): CurrentUser {
  if (!fallback) {
    return primary;
  }

  return {
    userId: primary.userId || fallback.userId,
    email: primary.email ?? fallback.email,
    displayName: primary.displayName ?? fallback.displayName,
    firstName: primary.firstName ?? fallback.firstName,
    lastName: primary.lastName ?? fallback.lastName,
    credits: primary.credits || fallback.credits,
    subscriptionPlan: primary.subscriptionPlan ?? fallback.subscriptionPlan,
    activationCode: primary.activationCode ?? fallback.activationCode,
    clientIsActive: primary.clientIsActive || fallback.clientIsActive,
  };
}

function normalizeServiceIncident(value: unknown): ServiceIncident {
  const record = asRecord(value);

  return {
    incidentId: readNumber(record, 0, "incidentId", "IncidentId", "serviceIncidentId", "ServiceIncidentId"),
    title: readString(record, "title", "Title", "serviceIncidentTitle", "ServiceIncidentTitle") ?? "Microsoft Teams incident",
    status: readString(record, "status", "Status", "serviceIncidentStatus", "ServiceIncidentStatus") ?? "Unresolved",
    impact: readString(record, "impact", "Impact", "serviceIncidentImpact", "ServiceIncidentImpact"),
    startedAt: readString(record, "startedAt", "StartedAt", "createdAt", "CreatedAt"),
  };
}

function normalizeTeamsServiceHealth(value: unknown): TeamsServiceHealth {
  const record = asRecord(value);
  const unresolvedIncidents = readArray(record, "unresolvedIncidents", "UnresolvedIncidents", "incidents", "Incidents").map(
    normalizeServiceIncident
  );

  return {
    serviceId: 3,
    serviceName: "Microsoftteams",
    serviceDisplayName: readString(record, "serviceDisplayName", "ServiceDisplayName") ?? "Microsoft Teams",
    currentStatus: readString(record, "currentStatus", "CurrentStatus", "serviceCurrentStatus", "ServiceCurrentStatus") ?? "No data",
    unresolvedIncidents,
  };
}

function normalizeMeetingParticipant(value: unknown): MeetingParticipant {
  const record = asRecord(value);
  const clientDataStatus =
    readString(record, "clientDataStatus", "ClientDataStatus") ??
    (readBoolean(record, false, "clientIsActive", "ClientIsActive") ? "active" : "no_data");

  return {
    userId: readNumber(record, 0, "userId", "UserId", "user_id"),
    displayName: readString(record, "displayName", "DisplayName", "displayNameSnapshot", "DisplayNameSnapshot"),
    email: readString(record, "email", "Email", "userEmail", "UserEmail"),
    authProvider: readString(record, "authProvider", "AuthProvider"),
    joinedAt: readString(record, "joinedAt", "JoinedAt") ?? "",
    lastSeenAt: readString(record, "lastSeenAt", "LastSeenAt") ?? "",
    signalScore: readNumber(record, Number.NaN, "signalScore", "SignalScore"),
    clientDataStatus: clientDataStatus as ClientDataStatus,
    clientIsActive: readBoolean(record, clientDataStatus === "active", "clientIsActive", "ClientIsActive"),
  };
}

function normalizeDashboardData(value: unknown): DashboardData {
  const record = asRecord(value);
  const currentUser = normalizeCurrentUser(record.currentUser ?? record.CurrentUser ?? record.user ?? record.User);
  const participants = readArray(record, "participants", "Participants").map(normalizeMeetingParticipant);
  const teamsServiceHealth = normalizeTeamsServiceHealth(
    record.teamsServiceHealth ?? record.TeamsServiceHealth ?? record.microsoftTeamsHealth ?? record.MicrosoftTeamsHealth
  );

  return {
    meetingSessionId: readNumber(record, 0, "meetingSessionId", "MeetingSessionId"),
    currentUser,
    teamsServiceHealth,
    participants,
  };
}

function getSignalTunerSessionToken(response: AuthResponse): string {
  const token = response.token ?? response.sessionToken ?? response.jwt ?? response.jwtToken;

  if (!token) {
    throw new Error(
      "Teams SSO succeeded, but the backend response did not include a SignalTuner session token. Expected token, sessionToken, jwt, or jwtToken."
    );
  }

  return token;
}

function DataList({ items }: { items: Array<{ label: string; value: unknown }> }) {
  return (
    <dl className="dataList">
      {items.map((item) => (
        <React.Fragment key={item.label}>
          <dt>{item.label}</dt>
          <dd>{formatValue(item.value)}</dd>
        </React.Fragment>
      ))}
    </dl>
  );
}

function ConfigPage() {
  React.useEffect(() => {
    teamsJs.app.initialize().then(() => {
      teamsJs.pages.config.setValidityState(true);
      teamsJs.pages.config.registerOnSaveHandler((saveEvent) => {
        const contentUrl = `${window.location.origin}/tabs/home`;

        teamsJs.pages.config
          .setConfig({
            entityId: "home",
            suggestedDisplayName: "SignalTuner",
            contentUrl,
            websiteUrl: contentUrl,
          })
          .then(() => saveEvent.notifySuccess())
          .catch((error) => saveEvent.notifyFailure(error instanceof Error ? error.message : String(error)));
      });
    });
  }, []);

  return (
    <main className="pageShell compactShell">
      <section className="centerPanel">
        <h1>Configure SignalTuner</h1>
        <p>This tab is ready to add to your Teams meeting.</p>
      </section>
    </main>
  );
}

function SignInPanel({
  error,
  isLoading,
  meetingContext,
  onEmailRequest,
  onEmailVerify,
  onOAuthStart,
  onTeamsSignIn,
}: {
  error: string | null;
  isLoading: boolean;
  meetingContext: TeamsMeetingContext | null;
  onEmailRequest: (email: string) => Promise<void>;
  onEmailVerify: (email: string, code: string) => Promise<void>;
  onOAuthStart: (provider: "google" | "github") => void;
  onTeamsSignIn: () => Promise<void>;
}) {
  const [email, setEmail] = React.useState("");
  const [code, setCode] = React.useState("");
  const [step, setStep] = React.useState<EmailCodeStep>("request");

  return (
    <main className="pageShell authShell">
      <section className="authPanel">
        <div>
          <p className="eyebrow">SignalTuner Teams session</p>
          <h1>Sign in to join this meeting health view.</h1>
          <p className="mutedText">Sign in to join the SignalTuner session for this Teams meeting.</p>
        </div>

        <div className="consentNotice">
          By joining this SignalTuner meeting session, your connectivity status and diagnostic information may be visible
          to other authenticated SignalTuner users in this meeting.
        </div>

        <div className="authActions">
          <button className="primaryButton" disabled={isLoading || !meetingContext} onClick={onTeamsSignIn} type="button">
            Continue with Teams
          </button>
          <div className="secondaryGrid">
            <button className="secondaryButton" disabled={isLoading} onClick={() => onOAuthStart("google")} type="button">
              Continue with Google
            </button>
            <button className="secondaryButton" disabled={isLoading} onClick={() => onOAuthStart("github")} type="button">
              Continue with GitHub
            </button>
          </div>
        </div>

        <form
          className="emailCodeForm"
          onSubmit={(event) => {
            event.preventDefault();
            if (step === "request") {
              void onEmailRequest(email).then(() => setStep("verify"));
              return;
            }

            void onEmailVerify(email, code);
          }}
        >
          <label>
            Email magic code
            <input
              autoComplete="email"
              onChange={(event) => setEmail(event.target.value)}
              placeholder="you@example.com"
              type="email"
              value={email}
            />
          </label>
          {step === "verify" && (
            <label>
              Code
              <input
                autoComplete="one-time-code"
                onChange={(event) => setCode(event.target.value)}
                placeholder="123456"
                type="text"
                value={code}
              />
            </label>
          )}
          <button className="secondaryButton" disabled={isLoading || !email} type="submit">
            {step === "request" ? "Email me a code" : "Verify code"}
          </button>
        </form>

        {error && <p className="inlineError">{error}</p>}
        {!meetingContext && <p className="inlineNote">Open this tab inside a Teams meeting to join a meeting session.</p>}
      </section>
    </main>
  );
}

function ClientPrompt({
  activationCodeError,
  isLoading,
  user,
  onContinue,
  onRefresh,
  onSignOut,
}: {
  activationCodeError: string | null;
  isLoading: boolean;
  user: CurrentUser;
  onContinue: () => void;
  onRefresh: () => Promise<void>;
  onSignOut: () => void;
}) {
  const downloadUrl = getDownloadUrl();
  const activationCode = user.activationCode?.trim() || "Loading...";

  return (
    <section className="panel clientPrompt">
      <div>
        <h2>Activate the desktop client</h2>
        <p>
          SignalTuner can show your live connection data after the desktop client is installed and paired with your
          account.
        </p>
      </div>
      <div className="activationCodeBlock">
        <span>Activation code</span>
        <strong>{activationCode}</strong>
        {activationCodeError && <p>{activationCodeError}</p>}
      </div>
      <div className="buttonRow">
        {downloadUrl ? (
          <a className="primaryButton buttonLink" href={downloadUrl}>
            Download SignalTuner
          </a>
        ) : (
          <span className="inlineNote">Desktop downloads are available for Windows and macOS.</span>
        )}
        <button className="secondaryButton" disabled={isLoading} onClick={() => void onRefresh()} type="button">
          Refresh status
        </button>
        <button className="secondaryButton" disabled={isLoading} onClick={onContinue} type="button">
          Continue without data
        </button>
        <button className="secondaryButton" disabled={isLoading} onClick={onSignOut} type="button">
          Sign out
        </button>
      </div>
    </section>
  );
}

function Dashboard({
  analysis,
  dashboard,
  error,
  isLoading,
  onAddCredits,
  onAnalyzeAll,
  onAnalyzeUser,
  onInvite,
  onRefresh,
  onSetPlan,
  onSignOut,
  subscriptionPrompt,
}: {
  analysis: AnalysisResult | null;
  dashboard: DashboardData;
  error: string | null;
  isLoading: boolean;
  onAddCredits: (amount: number) => Promise<void>;
  onAnalyzeAll: () => Promise<void>;
  onAnalyzeUser: (targetUserId: number) => Promise<void>;
  onInvite: () => Promise<void>;
  onRefresh: () => Promise<void>;
  onSetPlan: (plan: string) => Promise<void>;
  onSignOut: () => void;
  subscriptionPrompt: SubscriptionPrompt | null;
}) {
  const [accountOpen, setAccountOpen] = React.useState(false);
  const user = dashboard.currentUser;
  const activeParticipants = dashboard.participants.filter((participant) => participant.clientDataStatus === "active");

  return (
    <main className="pageShell">
      <header className="topBar">
        <div>
          <p className="eyebrow">SignalTuner meeting session</p>
          <h1>Microsoft Teams health dashboard</h1>
        </div>
        <div className="topActions">
          <span className="creditBadge">{user.credits} credits</span>
          <button className="secondaryButton" onClick={onRefresh} type="button">
            Refresh
          </button>
          <button className="secondaryButton" onClick={() => setAccountOpen((current) => !current)} type="button">
            Account
          </button>
        </div>
      </header>

      {accountOpen && (
        <section className="panel accountPanel">
          <div>
            <h2>{user.displayName ?? user.email ?? "SignalTuner account"}</h2>
            <DataList
              items={[
                { label: "Email", value: user.email },
                { label: "Credits remaining", value: user.credits },
                { label: "Subscription", value: user.subscriptionPlan ?? "Free" },
                { label: "Client active", value: user.clientIsActive },
              ]}
            />
          </div>
          <div className="accountControls">
            <p className="inlineNote">
              Test/development controls. TODO: restrict or remove before production subscription checkout is enabled.
            </p>
            <div className="buttonRow">
              {[1, 5, 10].map((amount) => (
                <button className="secondaryButton" key={amount} onClick={() => void onAddCredits(amount)} type="button">
                  Add Test +{amount}
                </button>
              ))}
            </div>
            <div className="buttonRow">
              {["Free", "Pro Test", "Team Test"].map((plan) => (
                <button className="secondaryButton" key={plan} onClick={() => void onSetPlan(plan)} type="button">
                  {plan}
                </button>
              ))}
            </div>
            <div className="buttonRow">
              {getDownloadUrl() && (
                <a className="secondaryButton buttonLink" href={getDownloadUrl() ?? undefined}>
                  Download client
                </a>
              )}
              <button className="secondaryButton" onClick={onSignOut} type="button">
                Sign out
              </button>
            </div>
          </div>
        </section>
      )}

      {error && <section className="panel panelAlert">{error}</section>}

      {subscriptionPrompt && (
        <section className="panel subscriptionPrompt">
          <h2>More credits required</h2>
          <p>
            This analysis requires {subscriptionPrompt.requiredCredits} credits. You have{" "}
            {subscriptionPrompt.availableCredits} available.
          </p>
          <button className="primaryButton" onClick={() => setAccountOpen(true)} type="button">
            Manage subscription
          </button>
          <p className="inlineNote">TODO: connect this action to production Stripe checkout.</p>
        </section>
      )}

      <section className="dashboardGrid">
        <article className="panel">
          <h2>Microsoft Teams health</h2>
          <span className={`statusBadge ${getSeverityClass(dashboard.teamsServiceHealth.currentStatus)}`}>
            {dashboard.teamsServiceHealth.currentStatus}
          </span>
          <DataList
            items={[
              { label: "Service", value: dashboard.teamsServiceHealth.serviceDisplayName },
              { label: "Service ID", value: dashboard.teamsServiceHealth.serviceId },
              { label: "Unresolved incidents", value: dashboard.teamsServiceHealth.unresolvedIncidents.length },
            ]}
          />
        </article>

        <article className="panel">
          <h2>Session</h2>
          <DataList
            items={[
              { label: "Session ID", value: dashboard.meetingSessionId },
              { label: "Participants", value: dashboard.participants.length },
              { label: "With active data", value: activeParticipants.length },
              { label: "Credits", value: user.credits },
            ]}
          />
        </article>
      </section>

      <section className="panel">
        <div className="sectionTitleRow">
          <div>
            <h2>Unresolved Microsoft Teams incidents</h2>
            <p>Shown from the SignalTuner service health tables for service ID 3.</p>
          </div>
        </div>
        {dashboard.teamsServiceHealth.unresolvedIncidents.length > 0 ? (
          <div className="incidentList">
            {dashboard.teamsServiceHealth.unresolvedIncidents.map((incident) => (
              <article className="incidentItem" key={incident.incidentId}>
                <strong>{incident.title}</strong>
                <span>{incident.status}</span>
                <p>{formatValue(incident.impact)}</p>
              </article>
            ))}
          </div>
        ) : (
          <p className="inlineNote">No unresolved Microsoft Teams incidents are currently reported.</p>
        )}
      </section>

      <section className="panel">
        <div className="sectionTitleRow">
          <div>
            <h2>Authenticated participants</h2>
            <p>Only SignalTuner users who opened this app and joined the current meeting session appear here.</p>
          </div>
          <div className="buttonRow inlineButtons">
            <button className="secondaryButton" disabled={isLoading} onClick={onInvite} type="button">
              Invite Participants
            </button>
            <button className="primaryButton" disabled={isLoading || activeParticipants.length === 0} onClick={onAnalyzeAll} type="button">
              Run Full Analysis
            </button>
          </div>
        </div>

        <div className="tableWrap">
          <table className="participantTable">
            <thead>
              <tr>
                <th>Participant</th>
                <th>Status / Signal Score</th>
                <th>Client Data</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {dashboard.participants.map((participant) => {
                const hasData = participant.clientDataStatus === "active";

                return (
                  <tr key={participant.userId}>
                    <td>
                      <strong>{getParticipantName(participant)}</strong>
                      <span>{participant.authProvider ?? "Signed in"}</span>
                    </td>
                    <td>{hasData ? participant.signalScore ?? "Active" : "No data"}</td>
                    <td>
                      <span className={`statusBadge ${hasData ? "severityGreen" : "severityBlue"}`}>
                        {hasData ? "Active telemetry" : "No data"}
                      </span>
                    </td>
                    <td>
                      {hasData ? (
                        <button className="secondaryButton" disabled={isLoading} onClick={() => onAnalyzeUser(participant.userId)} type="button">
                          Analyze Connection
                        </button>
                      ) : (
                        <button className="secondaryButton" disabled={isLoading} onClick={onInvite} type="button">
                          Prompt to Activate
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </section>

      {analysis && <AnalysisPanel analysis={analysis} />}
    </main>
  );
}

function AnalysisPanel({ analysis }: { analysis: AnalysisResult }) {
  if (analysis.mode === "user") {
    return (
      <section className="panel">
        <h2>Connection analysis: {getParticipantName(analysis.data.targetUser)}</h2>
        <p className="inlineNote">
          Credits used: {analysis.data.creditsUsed}. Remaining: {analysis.data.remainingCredits}.
        </p>
        <IssueList issues={analysis.data.issues} />
        <TelemetryTable telemetry={analysis.data.telemetry} />
      </section>
    );
  }

  return (
    <section className="panel">
      <h2>Full session analysis</h2>
      <DataList
        items={[
          { label: "Credits used", value: analysis.data.creditsUsed },
          { label: "Remaining credits", value: analysis.data.remainingCredits },
          { label: "Active users analyzed", value: analysis.data.groupSummary.activeUsersAnalyzed },
          { label: "Bandwidth issue users", value: analysis.data.groupSummary.usersWithBandwidthIssues },
          { label: "System issue users", value: analysis.data.groupSummary.usersWithSystemIssues },
          { label: "Network issue users", value: analysis.data.groupSummary.usersWithNetworkIssues },
          { label: "Teams active incident", value: analysis.data.groupSummary.teamsHasActiveServiceIncident },
        ]}
      />
      {analysis.data.analyzedUsers.map((result) => (
        <article className="analysisUser" key={result.participant.userId}>
          <h3>{getParticipantName(result.participant)}</h3>
          <IssueList issues={result.issues} />
          <TelemetryTable telemetry={result.telemetry} />
        </article>
      ))}
    </section>
  );
}

function IssueList({ issues }: { issues: Issue[] }) {
  if (issues.length === 0) {
    return <p className="inlineNote">No non-Excellent status fields were found.</p>;
  }

  return (
    <div className="issueGrid">
      {issues.map((issue) => (
        <article className="issueCard" key={`${issue.affectedParticipant}-${issue.field}`}>
          <span className={`statusBadge ${getSeverityClass(issue.severity)}`}>{issue.currentValue}</span>
          <strong>{issue.field}</strong>
          <p>{issue.recommendation}</p>
        </article>
      ))}
    </div>
  );
}

function TelemetryTable({ telemetry }: { telemetry: TelemetryRecord }) {
  const entries = Object.entries(telemetry);

  return (
    <div className="telemetryGrid">
      {entries.map(([key, value]) => (
        <React.Fragment key={key}>
          <span>{key}</span>
          <strong>{formatValue(value)}</strong>
        </React.Fragment>
      ))}
    </div>
  );
}

export default function App() {
  const isConfigPage = window.location.pathname.toLowerCase().startsWith("/tabs/config");
  const apiBaseUrl = normalizeBaseUrl(import.meta.env.VITE_SIGNALTUNER_API_URL);
  const [sessionToken, setSessionToken] = React.useState<string | null>(() =>
    window.localStorage.getItem(SIGNALTUNER_SESSION_TOKEN_KEY)
  );
  const [dashboard, setDashboard] = React.useState<DashboardData | null>(null);
  const [meetingContext, setMeetingContext] = React.useState<TeamsMeetingContext | null>(null);
  const [analysis, setAnalysis] = React.useState<AnalysisResult | null>(null);
  const [subscriptionPrompt, setSubscriptionPrompt] = React.useState<SubscriptionPrompt | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [isLoading, setIsLoading] = React.useState(false);
  const [isClientPromptDismissed, setIsClientPromptDismissed] = React.useState(false);
  const [activationCodeError, setActivationCodeError] = React.useState<string | null>(null);
  const [accountUser, setAccountUser] = React.useState<CurrentUser | null>(null);

  const parseCreditError = React.useCallback((caught: unknown): boolean => {
    const errorWithBody = caught as Error & { body?: string };

    if (!errorWithBody.body) {
      return false;
    }

    try {
      const parsed = JSON.parse(errorWithBody.body) as { requiredCredits?: number; availableCredits?: number };

      if (typeof parsed.requiredCredits === "number" && typeof parsed.availableCredits === "number") {
        setSubscriptionPrompt({
          requiredCredits: parsed.requiredCredits,
          availableCredits: parsed.availableCredits,
        });
        return true;
      }
    } catch {
      return false;
    }

    return false;
  }, []);

  const mergeActivationCode = React.useCallback((activationCode: string | null) => {
    if (!activationCode) {
      return;
    }

    setDashboard((current) =>
      current
        ? {
            ...current,
            currentUser: {
              ...current.currentUser,
              activationCode,
            },
          }
        : current
    );
    setAccountUser((current) =>
      current
        ? {
            ...current,
            activationCode,
          }
        : current
    );
  }, []);

  const refreshAccountInfo = React.useCallback(
    async (token: string) => {
      const account = normalizeCurrentUser(
        await fetchJson<unknown>(`${apiBaseUrl}/api/auth/me`, {
          headers: buildAuthHeaders(token),
        })
      );

      setAccountUser(account);
      setDashboard((current) =>
        current
          ? {
              ...current,
              currentUser: mergeCurrentUser(current.currentUser, account),
            }
          : current
      );

      return account;
    },
    [apiBaseUrl]
  );

  const refreshActivationCode = React.useCallback(
    async (token: string) => {
      const activationCode = normalizeActivationCodeResponse(
        await fetchJson<ActivationCodeResponse>(`${apiBaseUrl}/api/User/activation-code`, {
          headers: buildAuthHeaders(token),
        })
      );

      mergeActivationCode(activationCode);
      return activationCode;
    },
    [apiBaseUrl, mergeActivationCode]
  );

  const refreshDashboard = React.useCallback(
    async (meetingSessionId: number, token: string) => {
      const data = normalizeDashboardData(await fetchJson<unknown>(`${apiBaseUrl}/api/TeamsMeetings/${meetingSessionId}/dashboard`, {
        headers: buildAuthHeaders(token),
      }));
      setDashboard((current) => ({
        ...data,
        currentUser: mergeCurrentUser(data.currentUser, current?.currentUser ?? accountUser),
      }));
      return data;
    },
    [accountUser, apiBaseUrl]
  );

  const joinMeetingSession = React.useCallback(
    async (token: string, context: TeamsMeetingContext, fallbackUser: CurrentUser | null = null) => {
      const data = normalizeDashboardData(await fetchJson<unknown>(`${apiBaseUrl}/api/TeamsMeetings/join`, {
        method: "POST",
        headers: buildAuthHeaders(token),
        body: JSON.stringify(context),
      }));
      const mergedData = {
        ...data,
        currentUser: mergeCurrentUser(data.currentUser, fallbackUser ?? accountUser),
      };
      setDashboard(mergedData);
      return mergedData;
    },
    [accountUser, apiBaseUrl]
  );

  const completeAuth = React.useCallback(
    async (response: AuthResponse) => {
      if (!meetingContext) {
        throw new Error("Teams meeting context is unavailable.");
      }

      const signalTunerSessionToken = getSignalTunerSessionToken(response);

      window.localStorage.setItem(SIGNALTUNER_SESSION_TOKEN_KEY, signalTunerSessionToken);
      setSessionToken(signalTunerSessionToken);
      setIsClientPromptDismissed(false);
      setActivationCodeError(null);
      const account = await refreshAccountInfo(signalTunerSessionToken);
      await joinMeetingSession(signalTunerSessionToken, meetingContext, account);
    },
    [joinMeetingSession, meetingContext, refreshAccountInfo]
  );

  React.useEffect(() => {
    if (isConfigPage) {
      return;
    }

    const initializeTeams = async () => {
      setIsLoading(true);

      try {
        await teamsJs.app.initialize();
        const context = await teamsJs.app.getContext();
        const teamsMeetingId = context.meeting?.id;

        if (!teamsMeetingId) {
          setMeetingContext(null);
          setError("Open SignalTuner from inside a Microsoft Teams meeting to join a meeting session.");
          return;
        }

        setMeetingContext({
          teamsMeetingId,
          teamsConversationId: context.chat?.id ?? context.channel?.id ?? null,
          teamsTenantId: context.user?.tenant?.id ?? null,
          meetingTitle: context.meeting?.details?.title ?? null,
        });
        setError(null);
      } catch (caught) {
        setError(caught instanceof Error ? caught.message : String(caught));
      } finally {
        setIsLoading(false);
      }
    };

    void initializeTeams();
  }, [isConfigPage]);

  React.useEffect(() => {
    if (!sessionToken || !meetingContext || dashboard) {
      return;
    }

    setIsLoading(true);
    refreshAccountInfo(sessionToken)
      .then((account) => joinMeetingSession(sessionToken, meetingContext, account))
      .catch((caught) => {
        window.localStorage.removeItem(SIGNALTUNER_SESSION_TOKEN_KEY);
        setSessionToken(null);
        setError(caught instanceof Error ? caught.message : String(caught));
      })
      .finally(() => setIsLoading(false));
  }, [dashboard, joinMeetingSession, meetingContext, refreshAccountInfo, sessionToken]);

  React.useEffect(() => {
    if (!dashboard || !sessionToken) {
      return;
    }

    const currentUser = mergeCurrentUser(dashboard.currentUser, accountUser);

    if (!currentUser.clientIsActive && !currentUser.activationCode) {
      setIsLoading(true);
      refreshActivationCode(sessionToken)
        .then((activationCode) => {
          setActivationCodeError(activationCode ? null : "Unable to load your activation code.");
        })
        .catch((caught) => {
          setActivationCodeError(caught instanceof Error ? caught.message : String(caught));
        })
        .finally(() => setIsLoading(false));
    }
  }, [accountUser, dashboard, refreshActivationCode, sessionToken]);

  React.useEffect(() => {
    if (!dashboard || !sessionToken) {
      return;
    }

    const intervalId = window.setInterval(() => {
      void refreshDashboard(dashboard.meetingSessionId, sessionToken);
    }, 10000);

    return () => window.clearInterval(intervalId);
  }, [dashboard, refreshDashboard, sessionToken]);

  const signInWithTeams = React.useCallback(async () => {
    if (!apiBaseUrl || !meetingContext) {
      setError("SignalTuner API URL or Teams meeting context is unavailable.");
      return;
    }

    setIsLoading(true);
    setError(null);

    try {
      const teamsSsoToken = await teamsJs.authentication.getAuthToken();
      const response = await fetchJson<AuthResponse>(`${apiBaseUrl}/api/User/teams-sso`, {
        method: "POST",
        headers: buildAuthHeaders(null),
        body: JSON.stringify({
          teamsSsoToken,
          teamsTenantId: meetingContext.teamsTenantId,
          teamsMeetingId: meetingContext.teamsMeetingId,
        }),
      });

      await completeAuth(response);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
    } finally {
      setIsLoading(false);
    }
  }, [apiBaseUrl, completeAuth, meetingContext]);

  const requestEmailCode = React.useCallback(
    async (email: string) => {
      setIsLoading(true);
      setError(null);

      try {
        await fetchJson(`${apiBaseUrl}/api/auth/email-magic-code/request`, {
          method: "POST",
          headers: buildAuthHeaders(null),
          body: JSON.stringify({ email }),
        });
      } catch (caught) {
        setError(caught instanceof Error ? caught.message : String(caught));
        throw caught;
      } finally {
        setIsLoading(false);
      }
    },
    [apiBaseUrl]
  );

  const verifyEmailCode = React.useCallback(
    async (email: string, code: string) => {
      setIsLoading(true);
      setError(null);

      try {
        const response = await fetchJson<AuthResponse>(`${apiBaseUrl}/api/auth/email-magic-code/verify`, {
          method: "POST",
          headers: buildAuthHeaders(null),
          body: JSON.stringify({ email, code }),
        });
        await completeAuth(response);
      } catch (caught) {
        setError(caught instanceof Error ? caught.message : String(caught));
      } finally {
        setIsLoading(false);
      }
    },
    [apiBaseUrl, completeAuth]
  );

  const startOAuth = React.useCallback(
    (provider: "google" | "github") => {
      const returnUrl = encodeURIComponent(window.location.href);
      window.location.assign(`${apiBaseUrl}/api/auth/${provider}/start?returnUrl=${returnUrl}`);
    },
    [apiBaseUrl]
  );

  const analyzeUser = React.useCallback(
    async (targetUserId: number) => {
      if (!dashboard || !sessionToken) {
        return;
      }

      setIsLoading(true);
      setError(null);
      setSubscriptionPrompt(null);

      try {
        const data = await fetchJson<IndividualAnalysisResult>(`${apiBaseUrl}/api/analysis/user`, {
          method: "POST",
          headers: buildAuthHeaders(sessionToken),
          body: JSON.stringify({ meetingSessionId: dashboard.meetingSessionId, targetUserId }),
        });
        setAnalysis({ mode: "user", data });
        await refreshDashboard(dashboard.meetingSessionId, sessionToken);
      } catch (caught) {
        if (!parseCreditError(caught)) {
          setError(caught instanceof Error ? caught.message : String(caught));
        }
      } finally {
        setIsLoading(false);
      }
    },
    [apiBaseUrl, dashboard, parseCreditError, refreshDashboard, sessionToken]
  );

  const analyzeAll = React.useCallback(async () => {
    if (!dashboard || !sessionToken) {
      return;
    }

    setIsLoading(true);
    setError(null);
    setSubscriptionPrompt(null);

    try {
      const data = await fetchJson<FullAnalysisResult>(`${apiBaseUrl}/api/analysis/full`, {
        method: "POST",
        headers: buildAuthHeaders(sessionToken),
        body: JSON.stringify({ meetingSessionId: dashboard.meetingSessionId }),
      });
      setAnalysis({ mode: "full", data });
      await refreshDashboard(dashboard.meetingSessionId, sessionToken);
    } catch (caught) {
      if (!parseCreditError(caught)) {
        setError(caught instanceof Error ? caught.message : String(caught));
      }
    } finally {
      setIsLoading(false);
    }
  }, [apiBaseUrl, dashboard, parseCreditError, refreshDashboard, sessionToken]);

  const inviteParticipants = React.useCallback(async () => {
    if (!dashboard || !sessionToken) {
      return;
    }

    setIsLoading(true);
    setError(null);

    try {
      await fetchJson(`${apiBaseUrl}/api/TeamsMeetings/${dashboard.meetingSessionId}/invite`, {
        method: "POST",
        headers: buildAuthHeaders(sessionToken),
      });
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
    } finally {
      setIsLoading(false);
    }
  }, [apiBaseUrl, dashboard, sessionToken]);

  const addCredits = React.useCallback(
    async (amount: number) => {
      if (!sessionToken || !dashboard) {
        return;
      }

      await fetchJson(`${apiBaseUrl}/api/account/test-add-credits`, {
        method: "POST",
        headers: buildAuthHeaders(sessionToken),
        body: JSON.stringify({ amount }),
      });
      await refreshDashboard(dashboard.meetingSessionId, sessionToken);
    },
    [apiBaseUrl, dashboard, refreshDashboard, sessionToken]
  );

  const setPlan = React.useCallback(
    async (plan: string) => {
      if (!sessionToken || !dashboard) {
        return;
      }

      await fetchJson(`${apiBaseUrl}/api/account/test-subscription-plan`, {
        method: "POST",
        headers: buildAuthHeaders(sessionToken),
        body: JSON.stringify({ plan }),
      });
      await refreshDashboard(dashboard.meetingSessionId, sessionToken);
    },
    [apiBaseUrl, dashboard, refreshDashboard, sessionToken]
  );

  const signOut = React.useCallback(() => {
    window.localStorage.removeItem(SIGNALTUNER_SESSION_TOKEN_KEY);
    setSessionToken(null);
    setDashboard(null);
    setAnalysis(null);
    setSubscriptionPrompt(null);
    setIsClientPromptDismissed(false);
    setActivationCodeError(null);
    setAccountUser(null);
  }, []);

  if (isConfigPage) {
    return <ConfigPage />;
  }

  if (!sessionToken || !dashboard) {
    return (
      <SignInPanel
        error={error}
        isLoading={isLoading}
        meetingContext={meetingContext}
        onEmailRequest={requestEmailCode}
        onEmailVerify={verifyEmailCode}
        onOAuthStart={startOAuth}
        onTeamsSignIn={signInWithTeams}
      />
    );
  }

  const currentUser = mergeCurrentUser(dashboard.currentUser, accountUser);
  const displayDashboard = { ...dashboard, currentUser };
  const shouldShowClientPrompt = !currentUser.clientIsActive && !isClientPromptDismissed;

  return (
    <>
      {shouldShowClientPrompt && (
        <main className="pageShell">
          <ClientPrompt
            activationCodeError={activationCodeError}
            isLoading={isLoading}
            user={currentUser}
            onContinue={() => setIsClientPromptDismissed(true)}
            onRefresh={async () => {
              setActivationCodeError(null);
              const refreshedDashboard = await refreshDashboard(dashboard.meetingSessionId, sessionToken);

              if (!mergeCurrentUser(refreshedDashboard.currentUser, accountUser).activationCode) {
                const activationCode = await refreshActivationCode(sessionToken);

                if (!activationCode) {
                  setActivationCodeError("Unable to load your activation code.");
                }
              }
            }}
            onSignOut={signOut}
          />
        </main>
      )}
      {!shouldShowClientPrompt && (
        <Dashboard
          analysis={analysis}
          dashboard={displayDashboard}
          error={error}
          isLoading={isLoading}
          onAddCredits={addCredits}
          onAnalyzeAll={analyzeAll}
          onAnalyzeUser={analyzeUser}
          onInvite={inviteParticipants}
          onRefresh={async () => {
            await refreshDashboard(dashboard.meetingSessionId, sessionToken);
          }}
          onSetPlan={setPlan}
          onSignOut={signOut}
          subscriptionPrompt={subscriptionPrompt}
        />
      )}
    </>
  );
}
