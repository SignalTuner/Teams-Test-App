import TemporaryActivationCode, { ActivationContext } from "./TemporaryActivationCode";
import React from "react";
import * as teamsJs from "@microsoft/teams-js";

import "./App.css";
import deviceCoresIcon from "./assets/device-cores-icon.svg";
import deviceCpuIcon from "./assets/device-cpu-icon.svg";
import deviceDivIcon from "./assets/device-div-icon.svg";
import deviceMemoryIcon from "./assets/device-memory-icon.svg";
import deviceProcessorIcon from "./assets/device-processor-icon.svg";
import downloadStepIcon from "./assets/download.svg";
import networkDownloadIcon from "./assets/network-download-icon.svg";
import networkDivIcon from "./assets/network-div-icon.svg";
import networkLatencyIcon from "./assets/network-latency-icon.svg";
import networkPacketLossIcon from "./assets/network-packet-loss-icon.svg";
import networkUploadIcon from "./assets/network-upload-icon.svg";
import personCheckIcon from "./assets/person-check.svg";
import refreshStepIcon from "./assets/refresh.svg";
import signalRecommendationIcon from "./assets/signal-reccommendation-icon.svg";
import signalTunerDarkLogo from "./assets/signaltuner-logo-horizontal-darkmode.png";
import signalTunerLogo from "./assets/signaltuner-logo-horizontal.png";
import clientPreview from "./assets/client.png";
import analysisPreview from "./assets/analysis.png";
import loginDashboardPreview from "./assets/login-dashboard-preview.png";
import microsoftTeamsLogo from "./assets/microsoft-teams.png";
import privacyPolicyMarkdown from "./content/privacy-policy.md?raw";
import termsOfServiceMarkdown from "./content/terms-of-service.md?raw";
import workspaceCurrentNetworkIcon from "./assets/workspace-current-network-icon.svg";
import workspaceDivIcon from "./assets/workspace-div-icon.svg";
import workspaceNetworkFrequencyIcon from "./assets/workspace-network-frequency-icon.svg";
import workspaceVpnIcon from "./assets/workspace-vpn-icon.svg";
import workspaceWifiStrengthIcon from "./assets/workspace-wifi-strength-icon.svg";

type AuthProvider = "teams_sso" | "google" | "github" | "email_magic_code";
type ClientDataStatus = "active" | "inactive" | "no_data";
type AnalysisStatus = "Excellent" | "Fair" | "Poor" | "Critical" | "Offline" | "No data" | string;
type SignalDisplayStatus = "Excellent" | "Fair" | "Poor" | "Critical" | "Offline" | "Unknown";
type TeamsServiceStatus = "operational" | "activeIncident" | "outage";

type CurrentUser = {
  userId: number;
  email: string | null;
  displayName: string | null;
  firstName?: string | null;
  lastName?: string | null;
  hasPassword?: boolean;
  m365Upn?: string | null;
  authProvider?: string | null;
  credits: number;
  subscriptionPlan?: string | null;
  activationCode?: string | null;
  clientIsActive: boolean;
};

type AuthResponse = {
  token?: string;
  Token?: string;
  sessionToken?: string;
  jwt?: string;
  jwtToken?: string;
  email?: string;
  Email?: string;
  UserID?: number;
  UserId?: number;
  userId?: number;
  ActivationCode?: string;
  activationCode?: string;
  UserFirstName?: string;
  UserLastName?: string;
  firstName?: string;
  lastName?: string;
  displayName?: string;
  IsTeamAdmin?: boolean;
  UserTeamID?: number;
  UserHasSeenOnboarding?: boolean;
  profileRequired?: boolean;
  ProfileRequired?: boolean;
  authProvider?: string;
  userCreated?: boolean;
  user?: unknown;
  User?: unknown;
};

type PendingProfileAuth = {
  token: string;
  user: CurrentUser;
};

type PendingTeamsSsoAccountCreation = {
  user: CurrentUser;
};



type TeamsMeetingContext = {
  teamsMeetingId: string;
  teamsConversationId: string | null;
  teamsTenantId: string | null;
  meetingTitle: string | null;
  organizerM365ObjectId: string | null;
  organizerTenantId: string | null;
  currentUserM365ObjectId: string | null;
  currentUserTenantId: string | null;
  currentUserMeetingRole: string | null;
};

type TeamsSsoOptions = {
  allowAccountCreation?: boolean;
  firstName?: string;
  lastName?: string;
  termsOfServiceAccepted?: boolean;
  privacyPolicyAccepted?: boolean;
  termsOfServiceVersion?: string;
  privacyPolicyVersion?: string;
};

type ServiceIncident = {
  incidentId: number;
  title: string;
  status: string;
  impact?: string | null;
  startedAt?: string | null;
  link?: string | null;
  uniqueIdentifier?: string | null;
};

type TeamsServiceHealth = {
  serviceId: 3;
  serviceName: "Microsoftteams";
  serviceDisplayName: string;
  currentStatus: AnalysisStatus;
  unresolvedIncidents: ServiceIncident[];
  activeIncidents: ServiceIncident[];
  recentIncidents: ServiceIncident[];
};

type MeetingParticipant = {
  userId: number;
  displayName: string | null;
  email: string | null;
  authProvider: AuthProvider | string | null;
  meetingRole?: string | null;
  joinedAt: string;
  lastSeenAt: string;
  signalScore: number | null;
  overallStatus: AnalysisStatus | null;
  deviceStatus: AnalysisStatus | null;
  workspaceStatus: AnalysisStatus | null;
  networkStatus: AnalysisStatus | null;
  liveTelemetry: TelemetryRecord | null;
  clientDataStatus: ClientDataStatus;
  clientIsActive: boolean;
  analysisSessionExpiresAt: string | null;
};

type DashboardData = {
  meetingSessionId: number;
  currentUser: CurrentUser;
  teamsServiceHealth: TeamsServiceHealth;
  participants: MeetingParticipant[];
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

type SignalScoreTrendPoint = {
  timestampUtc: string;
  averageScore: number;
};

type SignalScoreTrendResponse = {
  userId: number;
  deviceId: string | null;
  points: SignalScoreTrendPoint[];
};

type SubscriptionPrompt = {
  requiredCredits: number;
  availableCredits: number;
};

const EXAMPLE_ACTIVE_INCIDENTS: ServiceIncident[] = [
  {
    incidentId: -1,
    title: "Users may experience choppy audio in meetings",
    status: "Active",
    impact: "Some calls may have degraded audio quality.",
    startedAt: "2026-07-24T10:32:00Z",
    uniqueIdentifier: "ST-DEMO-TEAMS-AUDIO-001",
  },
  {
    incidentId: -2,
    title: "Delays in message delivery in Teams",
    status: "Active",
    impact: "Chat messages may be delayed for some users.",
    startedAt: "2026-07-24T09:15:00Z",
    uniqueIdentifier: "ST-DEMO-TEAMS-CHAT-002",
  },
  {
    incidentId: -3,
    title: "Slow loading of Teams channels and tabs",
    status: "Active",
    impact: "Teams channels and embedded app tabs may load slowly.",
    startedAt: "2026-07-23T21:47:00Z",
    uniqueIdentifier: "ST-DEMO-TEAMS-TABS-003",
  },
  {
    incidentId: -4,
    title: "Meeting join failures for a subset of users",
    status: "Active",
    impact: "Some users may need to retry joining meetings.",
    startedAt: "2026-07-23T18:25:00Z",
    uniqueIdentifier: "ST-DEMO-TEAMS-JOIN-004",
  },
  {
    incidentId: -5,
    title: "Intermittent screen sharing degradation",
    status: "Active",
    impact: "Screen sharing may briefly freeze or reduce quality.",
    startedAt: "2026-07-23T15:10:00Z",
    uniqueIdentifier: "ST-DEMO-TEAMS-SHARE-005",
  },
];

type AuthPageMode = "login" | "create-account" | "forgot-password" | "support";
type InAppPage = "dashboard" | "account" | "settings" | "support";
type AuthBusyState =
  | "idle"
  | "initializing-teams"
  | "checking-session"
  | "auto-sso"
  | "teams-sso"
  | "email-login"
  | "email-register"
  | "password-reset"
  | "success";
type TeamsTheme = "default" | "dark" | "contrast";
type SignalTunerThemePreference = "light" | "dark";
type JsonRecord = Record<string, unknown>;

const SIGNALTUNER_SESSION_TOKEN_KEY = "signaltunerSessionToken";
const SIGNALTUNER_AUTO_SSO_FAILED_KEY = "signaltunerAutoSsoFailed";
const SIGNALTUNER_EXPLICIT_SIGN_OUT_KEY = "signaltunerExplicitSignOut";
const SIGNALTUNER_THEME_PREFERENCE_KEY = "signaltunerThemePreference";
const PRIVACY_POLICY_PATH = "/tabs/privacy";
const TERMS_OF_SERVICE_PATH = "/tabs/terms";
const LEGAL_DOCUMENT_EFFECTIVE_DATE = "2026-08-12";
const PASSWORD_REQUIREMENT_TEXT = "Use at least 8 characters, including uppercase, lowercase, number, and symbol.";
const CLIENT_PROMPT_REFRESH_INTERVAL_MS = 20000;
const CLIENT_PROMPT_COPIED_REFRESH_INTERVAL_MS = 2000;
const CLIENT_PROMPT_COPIED_REFRESH_DURATION_MS = 60000;

let teamsInitializationPromise: Promise<teamsJs.app.Context> | null = null;

function getCurrentYear(): number {
  return new Date().getFullYear();
}

function getReturnUrl(): string {
  const params = new URLSearchParams(window.location.search);
  return params.get("returnUrl") || `${window.location.origin}/tabs/home`;
}

function buildAuthPath(path: AuthPageMode, returnUrl: string): string {
  const params = new URLSearchParams({ returnUrl });
  return `/${path}?${params.toString()}`;
}

function navigateAuth(path: AuthPageMode, returnUrl: string): void {
  window.history.pushState({}, "", buildAuthPath(path, returnUrl));
  window.dispatchEvent(new PopStateEvent("popstate"));
}

function navigateToSupport(returnUrl: string): void {
  navigateAuth("support", returnUrl);
}

function restoreReturnUrl(returnUrl: string): void {
  try {
    const target = new URL(returnUrl, window.location.origin);

    if (target.origin === window.location.origin) {
      window.history.replaceState({}, "", `${target.pathname}${target.search}${target.hash}`);
    }
  } catch {
    window.history.replaceState({}, "", "/tabs/home");
  }
}

function getAuthPageMode(): AuthPageMode {
  const path = window.location.pathname.toLowerCase();
  if (path.includes("create-account")) {
    return "create-account";
  }

  if (path.includes("support")) {
    return "support";
  }

  return path.includes("forgot-password") ? "forgot-password" : "login";
}

function sanitizeAuthError(caught: unknown, fallback: string): string {
  if (!(caught instanceof Error)) {
    return fallback;
  }

  const status = (caught as Error & { status?: number }).status;

  if (status === 401 || status === 403) {
    return "Your email or password was not recognized.";
  }

  if (status && status >= 500) {
    return "SignalTuner is temporarily unavailable. Please try again.";
  }

  return fallback;
}

function getFriendlyErrorMessage(caught: unknown, fallback: string): string {
  if (!(caught instanceof Error)) {
    return fallback;
  }

  const body = (caught as Error & { body?: string }).body;

  if (body) {
    try {
      const parsed = JSON.parse(body) as { message?: string };
      if (parsed.message) {
        return parsed.message;
      }
    } catch {
      if (!body.trim().startsWith("{")) {
        return body;
      }
    }
  }

  return caught.message || fallback;
}

function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}

function validatePassword(password: string): string | undefined {
  if (password.length < 8) {
    return "Password must be at least 8 characters long.";
  }

  if (!/[A-Z]/.test(password)) {
    return "Password must contain at least one uppercase letter.";
  }

  if (!/[a-z]/.test(password)) {
    return "Password must contain at least one lowercase letter.";
  }

  if (!/[0-9]/.test(password)) {
    return "Password must contain at least one number.";
  }

  if (!/[~`!@#$%^&*()\-_+={[}\]|:;"'<,>.?/]/.test(password)) {
    return "Password must contain at least one special symbol.";
  }

  return undefined;
}

function mapTeamsTheme(theme: string | undefined): TeamsTheme {
  const normalized = (theme ?? "default").toLowerCase();

  if (normalized === "dark") {
    return "dark";
  }

  if (normalized === "contrast" || normalized === "highcontrast" || normalized === "high-contrast") {
    return "contrast";
  }

  return "default";
}

function applyTeamsTheme(theme: TeamsTheme): void {
  document.documentElement.dataset.teamsTheme = theme;
}

function getStoredThemePreference(): SignalTunerThemePreference {
  return window.localStorage.getItem(SIGNALTUNER_THEME_PREFERENCE_KEY) === "dark" ? "dark" : "light";
}

function getEffectiveTheme(teamsTheme: TeamsTheme, preference: SignalTunerThemePreference): TeamsTheme {
  if (teamsTheme === "contrast") {
    return "contrast";
  }

  return preference === "dark" ? "dark" : "default";
}

function normalizeTeamsObjectId(value: string | null | undefined): string | null {
  const trimmedValue = value?.trim();

  if (!trimmedValue) {
    return null;
  }

  const guidMatch = trimmedValue.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
  return (guidMatch?.[0] ?? trimmedValue).toLowerCase();
}

function sameTeamsId(left: string | null | undefined, right: string | null | undefined): boolean {
  const normalizedLeft = normalizeTeamsObjectId(left);
  const normalizedRight = normalizeTeamsObjectId(right);
  return Boolean(normalizedLeft && normalizedRight && normalizedLeft === normalizedRight);
}

async function initializeTeams(): Promise<teamsJs.app.Context> {
  if (!teamsInitializationPromise) {
    teamsInitializationPromise = teamsJs.app.initialize().then(() => teamsJs.app.getContext());
  }

  return teamsInitializationPromise;
}

async function getTeamsMeetingDetails(): Promise<JsonRecord | null> {
  await initializeTeams();

  const meetingApi = teamsJs.meeting as typeof teamsJs.meeting & {
    getMeetingDetailsVerbose?: () => Promise<unknown>;
  };

  if (typeof meetingApi.getMeetingDetailsVerbose === "function") {
    try {
      const meetingDetails = await meetingApi.getMeetingDetailsVerbose();
      if (meetingDetails) {
        return asRecord(meetingDetails);
      }
    } catch {
      // Fall back to the callback API below for Teams clients without verbose details.
    }
  }

  return new Promise((resolve) => {
    teamsJs.meeting.getMeetingDetails((error, meetingDetails) => {
      if (error || !meetingDetails) {
        resolve(null);
        return;
      }

      resolve(asRecord(meetingDetails));
    });
  });
}

async function getTeamsSsoToken(): Promise<string> {
  await initializeTeams();
  return teamsJs.authentication.getAuthToken();
}

function subscribeToTeamsThemeChanges(onThemeChange: (theme: TeamsTheme) => void): () => void {
  teamsJs.app.registerOnThemeChangeHandler((theme) => {
    onThemeChange(mapTeamsTheme(theme));
  });

  return () => teamsJs.app.registerOnThemeChangeHandler(() => undefined);
}

async function authenticateWithTeamsSso(
  apiBaseUrl: string,
  meetingContext: TeamsMeetingContext | null,
  options: TeamsSsoOptions = {}
): Promise<AuthResponse> {
  const teamsSsoToken = await getTeamsSsoToken();
  return fetchJson<AuthResponse>(`${apiBaseUrl}/api/User/teams-sso`, {
    method: "POST",
    headers: buildAuthHeaders(null),
    body: JSON.stringify({
      teamsSsoToken,
      teamsTenantId: meetingContext?.teamsTenantId ?? null,
      teamsMeetingId: meetingContext?.teamsMeetingId ?? null,
      allowAccountCreation: options.allowAccountCreation ?? false,
      firstName: options.firstName?.trim() || null,
      lastName: options.lastName?.trim() || null,
      termsOfServiceAccepted: options.termsOfServiceAccepted ?? false,
      privacyPolicyAccepted: options.privacyPolicyAccepted ?? false,
      termsOfServiceVersion: options.termsOfServiceVersion ?? null,
      privacyPolicyVersion: options.privacyPolicyVersion ?? null,
    }),
  });
}

async function signInWithEmail(apiBaseUrl: string, email: string, password: string): Promise<AuthResponse> {
  return fetchJson<AuthResponse>(`${apiBaseUrl}/api/User/Login`, {
    method: "POST",
    headers: buildAuthHeaders(null),
    body: JSON.stringify({ email, password }),
  });
}

async function requestPasswordResetEmail(apiBaseUrl: string, email: string): Promise<string> {
  const response = await fetch(`${apiBaseUrl}/api/User/RequestPasswordReset`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams({ email: email.trim() }),
  });

  const text = await response.text();

  if (!response.ok) {
    throw new Error(text || "Failed to send password reset email.");
  }

  return text || "Password reset email sent successfully.";
}

async function sendSupportEmail(
  apiBaseUrl: string,
  request: { sender: string; subject: string; body: string }
): Promise<string> {
  const response = await fetch(`${apiBaseUrl}/api/User/SendSupportEmail`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
  });

  const text = await response.text();

  if (!response.ok) {
    const error = new Error(text || `Support request failed with status ${response.status}`) as Error & {
      status?: number;
      body?: string;
    };
    error.status = response.status;
    error.body = text;
    throw error;
  }

  try {
    const parsed = text ? (JSON.parse(text) as { message?: string }) : {};
    return parsed.message || "Support request sent successfully.";
  } catch {
    return text || "Support request sent successfully.";
  }
}

async function updateAccountPassword(apiBaseUrl: string, token: string, userId: number, newPassword: string): Promise<string> {
  const formData = new FormData();
  formData.append("newPassword", newPassword);

  const response = await fetch(`${apiBaseUrl}/api/User/UpdatePassword/${userId}`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    body: formData,
  });

  const text = await response.text();

  if (!response.ok) {
    const error = new Error(text || "Failed to update password.") as Error & { body?: string };
    error.body = text;
    throw error;
  }

  return text || "Password updated successfully.";
}

async function createAccountWithEmail(
  apiBaseUrl: string,
  request: { email: string; password: string; firstName: string; lastName: string }
): Promise<AuthResponse> {
  const formData = new FormData();
  formData.append("Email", request.email.trim());
  formData.append("Password", request.password);
  formData.append("FirstName", request.firstName.trim());
  formData.append("LastName", request.lastName.trim());
  formData.append("SmsAlerts", "false");

  const result = await fetchJson<{ status?: string; message?: string }>(`${apiBaseUrl}/api/User/AddUser`, {
    method: "POST",
    body: formData,
  });

  if (result.status?.toLowerCase() === "error") {
    throw new Error(result.message || "Unable to create account.");
  }

  const response = await signInWithEmail(apiBaseUrl, request.email, request.password);
  response.email = response.email ?? response.Email ?? request.email.trim();
  response.firstName = response.firstName ?? response.UserFirstName ?? request.firstName.trim();
  response.lastName = response.lastName ?? response.UserLastName ?? request.lastName.trim();
  response.displayName = response.displayName ?? `${request.firstName.trim()} ${request.lastName.trim()}`;
  return response;
}

async function completeUserProfile(
  apiBaseUrl: string,
  token: string,
  request: { firstName: string; lastName: string }
): Promise<AuthResponse> {
  return fetchJson<AuthResponse>(`${apiBaseUrl}/api/User/profile`, {
    method: "PUT",
    headers: buildAuthHeaders(token),
    body: JSON.stringify({
      firstName: request.firstName.trim(),
      lastName: request.lastName.trim(),
    }),
  });
}

async function updateAccountEmail(apiBaseUrl: string, token: string, email: string): Promise<AuthResponse> {
  return fetchJson<AuthResponse>(`${apiBaseUrl}/api/User/email`, {
    method: "PUT",
    headers: buildAuthHeaders(token),
    body: JSON.stringify({ email: email.trim() }),
  });
}

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

function formatTrendTooltip(timestampUtc: string, score: number): string {
  const timestamp = new Date(timestampUtc);
  const timestampLabel = Number.isNaN(timestamp.getTime())
    ? timestampUtc
    : timestamp.toLocaleString([], { hour: "numeric", minute: "2-digit", second: "2-digit" });

  return `${timestampLabel}: ${Math.round(score)} Signal Score`;
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

function formatParticipantMeetingRole(role: string | null | undefined): string {
  const normalizedRole = role?.trim();

  if (!normalizedRole) {
    return "Participant";
  }

  const lowerRole = normalizedRole.toLowerCase();
  const knownRoles: Record<string, string> = {
    attendee: "Attendee",
    organizer: "Organizer",
    presenter: "Presenter",
  };

  if (knownRoles[lowerRole]) {
    return knownRoles[lowerRole];
  }

  return normalizedRole
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function getParticipantMeetingRoleRank(role: string | null | undefined): number {
  const normalizedRole = role
    ?.trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, "");

  const roleRanks: Record<string, number> = {
    organizer: 400,
    coorganizer: 350,
    presenter: 300,
    attendee: 200,
    participant: 100,
  };

  return normalizedRole ? roleRanks[normalizedRole] ?? 0 : 100;
}

function sortParticipantsByMeetingRole(participants: MeetingParticipant[]): MeetingParticipant[] {
  return participants
    .map((participant, index) => ({ participant, index }))
    .sort((left, right) => {
      const roleDifference =
        getParticipantMeetingRoleRank(right.participant.meetingRole) -
        getParticipantMeetingRoleRank(left.participant.meetingRole);

      if (roleDifference !== 0) {
        return roleDifference;
      }

      return left.index - right.index;
    })
    .map(({ participant }) => participant);
}

function getInitials(name: string | null | undefined, email: string | null | undefined): string {
  const source = name?.trim() || email?.trim() || "SignalTuner User";
  const parts = source
    .replace(/@.*/, "")
    .split(/\s|[._-]/)
    .filter(Boolean);

  return parts
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase())
    .join("") || "ST";
}

function getSignalTone(score: number | null | undefined): "good" | "fair" | "poor" | "none" {
  if (score === null || score === undefined || Number.isNaN(score)) {
    return "none";
  }

  if (score >= 75) {
    return "good";
  }

  if (score >= 55) {
    return "fair";
  }

  return "poor";
}

function getSignalDisplayStatusFromScore(score: number | null | undefined): "Excellent" | "Fair" | "Poor" | "Critical" | "Unknown" {
  if (score === null || score === undefined || Number.isNaN(score)) {
    return "Unknown";
  }

  if (score >= 90) {
    return "Excellent";
  }

  if (score >= 80) {
    return "Fair";
  }

  if (score >= 70) {
    return "Poor";
  }

  return "Critical";
}

function getSignalStatusLabel(status: AnalysisStatus | null | undefined, score: number | null | undefined): "Excellent" | "Fair" | "Poor" | "Critical" | "Unknown" {
  const normalizedStatus = status?.trim().toLowerCase();

  if (normalizedStatus === "excellent") {
    return "Excellent";
  }

  if (normalizedStatus === "fair") {
    return "Fair";
  }

  if (normalizedStatus === "poor") {
    return "Poor";
  }

  if (normalizedStatus === "critical") {
    return "Critical";
  }

  return getSignalDisplayStatusFromScore(score);
}

function getSignalOverallStatusLabel(status: AnalysisStatus | null | undefined, score: number | null | undefined): SignalDisplayStatus {
  const normalizedStatus = status?.trim().toLowerCase();

  if (normalizedStatus === "excellent") {
    return "Excellent";
  }

  if (normalizedStatus === "fair") {
    return "Fair";
  }

  if (normalizedStatus === "poor") {
    return "Poor";
  }

  if (normalizedStatus === "critical") {
    return "Critical";
  }

  if (normalizedStatus === "offline") {
    return "Offline";
  }

  return getSignalStatusLabel(null, score);
}

function getSignalStatusColor(status: SignalDisplayStatus): string {
  switch (status) {
    case "Excellent":
      return "#3BB537";
    case "Fair":
      return "#F9AE00";
    case "Poor":
      return "#FC6F20";
    case "Critical":
      return "#EC2F3E";
    case "Offline":
      return "#61708A";
    default:
      return "#6B96C1";
  }
}

function normalizeTeamsServiceStatus(health: TeamsServiceHealth): TeamsServiceStatus {
  const status = String(health.currentStatus ?? "").toLowerCase();
  const hasActiveIncidents = health.activeIncidents.length > 0 || health.unresolvedIncidents.length > 0;

  if (status.includes("outage") || status.includes("critical") || status.includes("down")) {
    return "outage";
  }

  if (status.includes("incident") || status.includes("degrad") || status.includes("issue") || hasActiveIncidents) {
    return "activeIncident";
  }

  return "operational";
}

function getTeamsStatusMeta(status: TeamsServiceStatus): { label: string; className: string; description: string } {
  if (status === "outage") {
    return {
      label: "Outage",
      className: "statusOutage",
      description: "Microsoft Teams is reporting a service outage that may interrupt meeting quality.",
    };
  }

  if (status === "activeIncident") {
    return {
      label: "Active Incident",
      className: "statusIncident",
      description: "Some users may experience degraded performance.",
    };
  }

  return {
    label: "Operational",
    className: "statusOperational",
    description: "Microsoft Teams service health is currently operational.",
  };
}

function getTelemetryValue(telemetry: TelemetryRecord | null, keys: string[], fallback = "No data"): string {
  if (!telemetry) {
    return fallback;
  }

  for (const key of keys) {
    const value = telemetry[key];

    if (value !== undefined && value !== null && value !== "") {
      return formatValue(value);
    }
  }

  return fallback;
}

function getTelemetryRawValue(telemetry: TelemetryRecord | null, keys: string[]): string | number | boolean | null {
  if (!telemetry) {
    return null;
  }

  for (const key of keys) {
    const value = telemetry[key];

    if (value !== undefined && value !== null && value !== "") {
      return value;
    }
  }

  return null;
}

function getTelemetryNumber(telemetry: TelemetryRecord | null, keys: string[]): number | null {
  const value = getTelemetryRawValue(telemetry, keys);

  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string") {
    const parsed = Number.parseFloat(value.replace(/,/g, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }

  return null;
}

function formatTelemetryMeasurement(telemetry: TelemetryRecord | null, keys: string[], unit = ""): string {
  const value = getTelemetryRawValue(telemetry, keys);

  if (value === null) {
    return "No data";
  }

  if (typeof value === "number") {
    return `${Number.isInteger(value) ? value : value.toFixed(1)}${unit}`;
  }

  if (typeof value === "boolean") {
    return value ? "Yes" : "No";
  }

  const text = value.trim();
  const isNumericText = text !== "" && Number.isFinite(Number.parseFloat(text.replace(/,/g, "")));
  return isNumericText && unit && !text.toLowerCase().includes(unit.trim().toLowerCase()) ? `${text}${unit}` : text;
}

function truncateTelemetryValue(value: string, maxLength = 12): string {
  return value.length > maxLength ? `${value.slice(0, maxLength - 1)}...` : value;
}

function getTelemetryValueTitle(value: string): string | undefined {
  return value.length > 12 ? value : undefined;
}

const signalMetricColors = {
  unknown: "#6b96c1",
  excellent: "#3BB537",
  fair: "#F9AE00",
  poor: "#FC6F20",
  critical: "#EC2F3E",
} as const;

type TelemetryInsight = {
  title: "Device" | "Workspace" | "Network";
  text: string;
  severity: "Informational" | "Fair" | "Poor" | "Critical";
};

type TelemetryRecommendation = {
  text: string;
  severity: TelemetryInsight["severity"];
  priority: number;
};

function getTelemetryInsightColor(severity: TelemetryInsight["severity"]): string {
  switch (severity) {
    case "Critical":
      return signalMetricColors.critical;
    case "Poor":
      return signalMetricColors.poor;
    case "Fair":
      return signalMetricColors.fair;
    default:
      return signalMetricColors.excellent;
  }
}

function buildTelemetryInsights(telemetry: TelemetryRecord | null): {
  overall: string;
  categories: TelemetryInsight[];
  recommendations: TelemetryRecommendation[];
} {
  const cpu = getTelemetryNumber(telemetry, ["cpu", "cpuUsage", "signal_cpu", "cpu_percent", "SignalCPU"]);
  const memory = getTelemetryNumber(telemetry, ["memory", "memoryUsage", "signal_memory", "memory_percent", "SignalMemory"]);
  const wifiStrength = getTelemetryNumber(telemetry, ["wifiStrength", "wiFiStrength", "signal_wifi_strength", "wifi_strength", "SignalWifiStrength"]);
  const download = getTelemetryNumber(telemetry, ["downloadSpeed", "signal_download_speed", "download_speed", "SignalDownloadSpeed"]);
  const upload = getTelemetryNumber(telemetry, ["uploadSpeed", "signal_upload_speed", "upload_speed", "SignalUploadSpeed"]);
  const latency = getTelemetryNumber(telemetry, ["latency", "ping", "signal_latency", "latencyMs", "SignalPing"]);
  const packetLoss = getTelemetryNumber(telemetry, ["packetLoss", "signal_packet_loss", "packet_loss", "SignalPacketLoss"]);
  const vpn = getTelemetryValue(telemetry, ["vpn", "vpnStatus", "vpnDetected", "signal_vpn", "vpn_status", "signal_vpn_detected", "SignalVpnDetected"]);
  const categories: TelemetryInsight[] = [];
  const recommendations: TelemetryRecommendation[] = [];

  if (cpu !== null && memory !== null && cpu >= 85 && memory >= 85) {
    categories.push({ title: "Device", text: "Your device is under heavy CPU and memory load, which may cause degraded audio/video performance or application responsiveness.", severity: cpu >= 95 || memory >= 95 ? "Critical" : "Poor" });
    recommendations.push({ text: "Close unnecessary or resource-intensive applications to free system resources.", severity: "Poor", priority: 96 });
  } else if (cpu !== null && cpu >= 95) {
    categories.push({ title: "Device", text: "Your CPU is nearly fully utilized and may not have enough processing capacity available for reliable audio or video processing.", severity: "Critical" });
    recommendations.push({ text: "Close CPU-intensive applications before joining or continuing your meeting.", severity: "Critical", priority: 88 });
  } else if (memory !== null && memory >= 95) {
    categories.push({ title: "Device", text: "Available memory is critically low. Windows may begin relying heavily on virtual memory, which can reduce meeting performance.", severity: "Critical" });
    recommendations.push({ text: "Close memory-intensive applications to free RAM.", severity: "Critical", priority: 86 });
  } else if ((cpu !== null && cpu >= 85) || (memory !== null && memory >= 85)) {
    categories.push({ title: "Device", text: cpu !== null && cpu >= 85 ? "CPU utilization is high and may affect meeting performance if demand increases further." : "Memory usage is high and could reduce application responsiveness during a meeting.", severity: "Poor" });
    recommendations.push({ text: cpu !== null && cpu >= 85 ? "Close unnecessary applications to free processing capacity." : "Close unused applications or browser tabs.", severity: "Poor", priority: 60 });
  } else {
    categories.push({ title: "Device", text: "Your device currently has ample processing and memory capacity for a video meeting.", severity: "Informational" });
  }

  if (wifiStrength !== null && wifiStrength < 40) {
    categories.push({ title: "Workspace", text: "Your Wi-Fi connection is consistently weak rather than experiencing a temporary signal drop.", severity: "Poor" });
    recommendations.push({ text: "Move closer to your Wi-Fi access point or improve Wi-Fi coverage in this workspace.", severity: "Poor", priority: 88 });
  } else if (wifiStrength !== null && wifiStrength < 55) {
    categories.push({ title: "Workspace", text: "Your Wi-Fi signal is weak. Distance or obstructions between your device and access point may be reducing connection stability.", severity: "Poor" });
    recommendations.push({ text: "Move closer to the access point or switch to a more reliable Wi-Fi band or access point.", severity: "Poor", priority: 84 });
  } else {
    categories.push({ title: "Workspace", text: "Your workspace connection currently provides suitable Wi-Fi coverage for a video meeting.", severity: "Informational" });
  }

  const networkValues = [download, upload, latency, packetLoss];
  if (networkValues.some((value) => value !== null)) {
    const criticalPacketLoss = packetLoss !== null && packetLoss > 5.9;
    const criticalLatency = latency !== null && latency > 99;
    const poorBandwidth = (download !== null && download < 0.8) || (upload !== null && upload < 1.2);
    if (criticalPacketLoss && wifiStrength !== null && wifiStrength >= 70) {
      categories.push({ title: "Network", text: "Your Wi-Fi signal is strong, so wireless coverage is unlikely to be the primary source of your connection problem.", severity: "Critical" });
      recommendations.push({ text: "Check router or ISP performance and temporarily test without the VPN if one is connected.", severity: "Critical", priority: 90 });
    } else if (criticalLatency && vpn === "Connected") {
      categories.push({ title: "Network", text: "Your connection is experiencing critically high latency while routed through a VPN.", severity: "Critical" });
      recommendations.push({ text: "If permitted by your organization, temporarily disconnect the VPN and retest.", severity: "Critical", priority: 94 });
    } else if (criticalPacketLoss || criticalLatency) {
      categories.push({ title: "Network", text: "Your connection is experiencing broader instability that may affect real-time audio and video communication.", severity: "Critical" });
      recommendations.push({ text: "Reduce network congestion or switch to a more stable connection.", severity: "Critical", priority: 96 });
    } else if (poorBandwidth) {
      categories.push({ title: "Network", text: "Your connection is experiencing a bandwidth bottleneck that may affect meeting quality.", severity: "Poor" });
      recommendations.push({ text: "Reduce other network activity or switch to a faster internet connection.", severity: "Poor", priority: 90 });
    } else {
      categories.push({ title: "Network", text: vpn === "Connected" ? "Your network is performing well even with the VPN connected." : "Your internet connection is currently well suited for real-time audio and video communication.", severity: "Informational" });
    }
  } else {
    categories.push({ title: "Network", text: "Network telemetry is not available for this analysis window.", severity: "Informational" });
  }

  const rankedRecommendations = recommendations
    .sort((first, second) => second.priority - first.priority)
    .filter((recommendation, index, all) => all.findIndex((candidate) => candidate.text === recommendation.text) === index)
    .slice(0, 3);
  const overall = rankedRecommendations.length > 0
    ? rankedRecommendations.map((recommendation) => recommendation.text).join(" ")
    : "No immediate recommendations. Your current workspace, network, and device appear healthy.";
  return { overall, categories, recommendations: rankedRecommendations };
}

function getCpuMetricColor(value: number | null): string {
  if (value === null || value === 0) {
    return signalMetricColors.unknown;
  }

  if (value >= 95) {
    return signalMetricColors.critical;
  }

  if (value >= 85) {
    return signalMetricColors.poor;
  }

  if (value >= 70) {
    return signalMetricColors.fair;
  }

  return signalMetricColors.excellent;
}

function getMemoryMetricColor(value: number | null): string {
  if (value === null || value === 0) {
    return signalMetricColors.unknown;
  }

  if (value >= 98) {
    return signalMetricColors.critical;
  }

  if (value >= 96) {
    return signalMetricColors.poor;
  }

  if (value >= 93) {
    return signalMetricColors.fair;
  }

  return signalMetricColors.excellent;
}

function getWifiStrengthMetricColor(value: number | null): string {
  if (value === null || value === 0) {
    return signalMetricColors.unknown;
  }

  if (value >= 90) {
    return signalMetricColors.excellent;
  }

  if (value >= 80) {
    return signalMetricColors.fair;
  }

  if (value >= 70) {
    return signalMetricColors.poor;
  }

  return signalMetricColors.critical;
}

function getWifiBandMetricColor(value: string, wifiStrength: number | null): string {
  if (value === "5 GHz" || value === "6 GHz") {
    return wifiStrength !== null && wifiStrength >= 90 ? signalMetricColors.excellent : signalMetricColors.fair;
  }

  if (value === "2.4 GHz") {
    return signalMetricColors.excellent;
  }

  return signalMetricColors.unknown;
}

function getVpnMetricColor(value: string): string {
  if (value === "Not Connected") {
    return signalMetricColors.excellent;
  }

  if (value === "Connected") {
    return signalMetricColors.fair;
  }

  return signalMetricColors.unknown;
}

function getDownloadMetricColor(value: number | null): string {
  if (value === null || value === 0) {
    return signalMetricColors.unknown;
  }

  if (value >= 3) {
    return signalMetricColors.excellent;
  }

  if (value >= 1.8) {
    return signalMetricColors.fair;
  }

  if (value >= 0.8) {
    return signalMetricColors.poor;
  }

  return signalMetricColors.critical;
}

function getUploadMetricColor(value: number | null): string {
  if (value === null || value === 0) {
    return signalMetricColors.unknown;
  }

  if (value >= 3.8) {
    return signalMetricColors.excellent;
  }

  if (value >= 2.6) {
    return signalMetricColors.fair;
  }

  if (value >= 1.2) {
    return signalMetricColors.poor;
  }

  return signalMetricColors.critical;
}

function getLatencyMetricColor(value: number | null): string {
  if (value === null) {
    return signalMetricColors.unknown;
  }

  if (value <= 20) {
    return signalMetricColors.excellent;
  }

  if (value <= 50) {
    return signalMetricColors.fair;
  }

  if (value <= 99) {
    return signalMetricColors.poor;
  }

  return signalMetricColors.critical;
}

function getPacketLossMetricColor(value: number | null): string {
  if (value === null) {
    return signalMetricColors.unknown;
  }

  if (value <= 1.9) {
    return signalMetricColors.excellent;
  }

  if (value <= 3.9) {
    return signalMetricColors.fair;
  }

  if (value <= 5.9) {
    return signalMetricColors.poor;
  }

  if (value > 6) {
    return signalMetricColors.critical;
  }

  return signalMetricColors.unknown;
}

function getParticipantTelemetry(analysis: AnalysisResult | null, participantId: number): TelemetryRecord | null {
  if (!analysis) {
    return null;
  }

  if (analysis.mode === "user") {
    return analysis.data.targetUser.userId === participantId ? analysis.data.telemetry : null;
  }

  return analysis.data.analyzedUsers.find((result) => result.participant.userId === participantId)?.telemetry ?? null;
}

function getParticipantLiveTelemetry(participant: MeetingParticipant): TelemetryRecord | null {
  return participant.liveTelemetry;
}

function parseUtcTimestampMs(value: string | null | undefined): number | null {
  if (!value) {
    return null;
  }

  const normalizedValue = /(?:z|[+-]\d{2}:?\d{2})$/i.test(value) ? value : `${value}Z`;
  const timestampMs = Date.parse(normalizedValue);
  return Number.isNaN(timestampMs) ? null : timestampMs;
}

function getAnalysisRemainingMs(participant: MeetingParticipant, nowMs: number): number {
  const expiresAtMs = parseUtcTimestampMs(participant.analysisSessionExpiresAt);
  return expiresAtMs ? Math.max(0, expiresAtMs - nowMs) : 0;
}

function formatAnalysisCountdown(remainingMs: number): string {
  const totalSeconds = Math.max(0, Math.ceil(remainingMs / 1000));
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  if (hours > 0) {
    return `${hours}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  }

  return `${minutes}:${String(seconds).padStart(2, "0")}`;
}

function getParticipantIssues(analysis: AnalysisResult | null, participantId: number): Issue[] {
  if (!analysis) {
    return [];
  }

  if (analysis.mode === "user") {
    return analysis.data.targetUser.userId === participantId ? analysis.data.issues : [];
  }

  return analysis.data.analyzedUsers.find((result) => result.participant.userId === participantId)?.issues ?? [];
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
    hasPassword: readBoolean(record, false, "hasPassword", "HasPassword", "userHasPassword", "UserHasPassword"),
    m365Upn: readString(record, "m365Upn", "M365Upn", "userM365Upn", "UserM365Upn", "user_m365_upn"),
    authProvider: readString(record, "authProvider", "AuthProvider", "userAuthProvider", "UserAuthProvider", "user_auth_provider"),
    credits: readNumber(record, 0, "credits", "Credits", "userCredits", "UserCredits", "user_credits"),
    subscriptionPlan: readString(record, "subscriptionPlan", "SubscriptionPlan", "userStripePlan", "UserStripePlan"),
    activationCode: readString(record, "activationCode", "ActivationCode", "userActivationCode", "UserActivationCode"),
    clientIsActive: readBoolean(record, false, "clientIsActive", "ClientIsActive", "clientIsActivated", "ClientIsActivated"),
  };
}

function normalizeAuthCurrentUser(response: AuthResponse): CurrentUser {
  const nestedUser = normalizeCurrentUser(response.user ?? response.User);
  const topLevelUser = normalizeCurrentUser(response);
  return mergeCurrentUser(nestedUser, topLevelUser);
}

function isProfileRequired(response: AuthResponse, user: CurrentUser): boolean {
  const record = asRecord(response);
  const nested = asRecord(response.user ?? response.User);

  return (
    readBoolean(record, false, "profileRequired", "ProfileRequired") ||
    readBoolean(nested, false, "profileRequired", "ProfileRequired") ||
    (!user.firstName && !user.lastName && readBoolean(record, false, "userCreated", "UserCreated"))
  );
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
    hasPassword: primary.hasPassword || fallback.hasPassword,
    m365Upn: primary.m365Upn ?? fallback.m365Upn,
    authProvider: primary.authProvider ?? fallback.authProvider,
    credits: Number.isFinite(primary.credits) ? primary.credits : fallback.credits,
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
    link: readString(record, "link", "Link", "serviceIncidentLink", "ServiceIncidentLink"),
    uniqueIdentifier: readString(record, "uniqueIdentifier", "UniqueIdentifier", "serviceIncidentUniqueIdentifier", "ServiceIncidentUniqueIdentifier"),
  };
}

function asTelemetryRecord(value: unknown): TelemetryRecord | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }

  const telemetry: TelemetryRecord = {};

  for (const [key, rawValue] of Object.entries(value)) {
    if (typeof rawValue === "string" || typeof rawValue === "number" || typeof rawValue === "boolean" || rawValue === null) {
      telemetry[key] = rawValue;
    }
  }

  return Object.keys(telemetry).length > 0 ? telemetry : null;
}

function isResolvedIncident(incident: ServiceIncident): boolean {
  return incident.status.trim().toLowerCase() === "resolved";
}

function normalizeTeamsServiceHealth(value: unknown): TeamsServiceHealth {
  const record = asRecord(value);
  const unresolvedIncidents = readArray(record, "unresolvedIncidents", "UnresolvedIncidents", "incidents", "Incidents").map(
    normalizeServiceIncident
  );
  const activeIncidents = readArray(record, "activeIncidents", "ActiveIncidents").map(normalizeServiceIncident);
  const recentIncidents = readArray(record, "recentIncidents", "RecentIncidents").map(normalizeServiceIncident);

  return {
    serviceId: 3,
    serviceName: "Microsoftteams",
    serviceDisplayName: readString(record, "serviceDisplayName", "ServiceDisplayName") ?? "Microsoft Teams",
    currentStatus: readString(record, "currentStatus", "CurrentStatus", "serviceCurrentStatus", "ServiceCurrentStatus") ?? "No data",
    unresolvedIncidents,
    activeIncidents: activeIncidents.length > 0 ? activeIncidents : unresolvedIncidents.filter((incident) => incident.status.toLowerCase() === "active"),
    recentIncidents: recentIncidents.filter(isResolvedIncident).slice(0, 10),
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
    meetingRole: readString(record, "meetingRole", "MeetingRole", "participantRole", "ParticipantRole", "role", "Role"),
    joinedAt: readString(record, "joinedAt", "JoinedAt") ?? "",
    lastSeenAt: readString(record, "lastSeenAt", "LastSeenAt") ?? "",
    signalScore: readNumber(record, Number.NaN, "signalScore", "SignalScore"),
    overallStatus: readString(record, "overallStatus", "OverallStatus", "signalOverallStatus", "SignalOverallStatus"),
    deviceStatus: readString(record, "deviceStatus", "DeviceStatus", "signalSystemStatus", "SignalSystemStatus"),
    workspaceStatus: readString(record, "workspaceStatus", "WorkspaceStatus", "signalWifiStatus", "SignalWifiStatus"),
    networkStatus: readString(record, "networkStatus", "NetworkStatus", "signalBandwidthStatus", "SignalBandwidthStatus"),
    liveTelemetry: asTelemetryRecord(record.liveTelemetry ?? record.LiveTelemetry ?? record.telemetry ?? record.Telemetry),
    clientDataStatus: clientDataStatus as ClientDataStatus,
    clientIsActive: readBoolean(record, clientDataStatus === "active", "clientIsActive", "ClientIsActive"),
    analysisSessionExpiresAt: readString(record, "analysisSessionExpiresAt", "AnalysisSessionExpiresAt", "analysis_session_expires_at"),
  };
}

function normalizeSignalScoreTrend(value: unknown): SignalScoreTrendResponse {
  const record = asRecord(value);
  const points = readArray(record, "points", "Points").map((point) => {
    const pointRecord = asRecord(point);

    return {
      timestampUtc: readString(pointRecord, "timestampUtc", "TimestampUtc", "minuteTimestampUtc", "MinuteTimestampUtc") ?? "",
      averageScore: readNumber(pointRecord, 0, "averageScore", "AverageScore", "averageSignalScore", "AverageSignalScore"),
    };
  });

  return {
    userId: readNumber(record, 0, "userId", "UserId", "user_id"),
    deviceId: readString(record, "deviceId", "DeviceId", "signalDeviceId", "SignalDeviceId"),
    points,
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
  const token = response.token ?? response.Token ?? response.sessionToken ?? response.jwt ?? response.jwtToken;

  if (!token) {
    throw new Error(
      "Teams SSO succeeded, but the backend response did not include a SignalTuner session token. Expected token, sessionToken, jwt, or jwtToken."
    );
  }

  return token;
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

function SessionLoadingPage({ message = "Restoring your SignalTuner session." }: { message?: string }) {
  return (
    <main className="pageShell compactShell">
      <section className="centerPanel sessionLoadingPanel" aria-busy="true" aria-live="polite">
        <Spinner />
        <h1>Loading SignalTuner</h1>
        <p>{message}</p>
      </section>
    </main>
  );
}

function SignedInOutsideMeetingPage({ onSignOut }: { onSignOut: () => void }) {
  return (
    <main className="pageShell compactShell">
      <section className="centerPanel">
        <h1>Open SignalTuner in a Teams meeting</h1>
        <p>You are signed in. Open this tab from inside a Microsoft Teams meeting to join a meeting health view.</p>
        <button className="secondaryButton" onClick={onSignOut} type="button">
          Sign out
        </button>
      </section>
    </main>
  );
}

function Spinner() {
  return <span className="spinner" aria-hidden="true" />;
}

function renderMarkdownInline(text: string): React.ReactNode[] {
  const nodes: React.ReactNode[] = [];
  const boldPattern = /\*\*(.+?)\*\*/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = boldPattern.exec(text)) !== null) {
    if (match.index > lastIndex) {
      nodes.push(text.slice(lastIndex, match.index));
    }

    nodes.push(<strong key={`${match.index}-${match[1]}`}>{match[1]}</strong>);
    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < text.length) {
    nodes.push(text.slice(lastIndex));
  }

  return nodes;
}

function renderLegalMarkdown(markdown: string): React.ReactNode[] {
  const blocks: React.ReactNode[] = [];
  const lines = markdown.split(/\r?\n/);
  let paragraphLines: string[] = [];
  let listItems: string[] = [];
  let blockIndex = 0;

  const flushParagraph = () => {
    if (!paragraphLines.length) {
      return;
    }

    blocks.push(<p key={`p-${blockIndex++}`}>{renderMarkdownInline(paragraphLines.join(" "))}</p>);
    paragraphLines = [];
  };

  const flushList = () => {
    if (!listItems.length) {
      return;
    }

    blocks.push(
      <ul className="legalList" key={`ul-${blockIndex++}`}>
        {listItems.map((item, index) => (
        <li key={`${index}-${item}`}>{renderMarkdownInline(item)}</li>
        ))}
      </ul>
    );
    listItems = [];
  };

  for (const line of lines) {
    const trimmedLine = line.trim();

    if (!trimmedLine) {
      flushParagraph();
      flushList();
      continue;
    }

    const heading = /^(#{1,3})\s+(.+)$/.exec(trimmedLine);
    if (heading) {
      flushParagraph();
      flushList();

      const level = heading[1].length;
      const content = renderMarkdownInline(heading[2]);
      const key = `h-${blockIndex++}`;

      if (level === 1) {
        blocks.push(<h1 key={key}>{content}</h1>);
      } else if (level === 2) {
        blocks.push(<h2 key={key}>{content}</h2>);
      } else {
        blocks.push(<h3 key={key}>{content}</h3>);
      }
      continue;
    }

    const listItem = /^-\s+(.+)$/.exec(trimmedLine);
    if (listItem) {
      flushParagraph();
      listItems.push(listItem[1]);
      continue;
    }

    flushList();
    paragraphLines.push(trimmedLine.replace(/\s{2,}$/, ""));
  }

  flushParagraph();
  flushList();
  return blocks;
}

function LegalDocumentPage({ markdown, title }: { markdown: string; title: string }) {
  React.useEffect(() => {
    document.title = title;
  }, [title]);

  const backToSignInLink = (
    <a className="legalBackButton" href="/tabs/home">
      <span className="backArrowIcon" aria-hidden="true" />
      <span>Back to sign in</span>
    </a>
  );

  return (
    <main className="pageShell legalShell">
      <article className="legalPanel">
        <header className="legalHeader">
          <SignalTunerLogo className="legalLogo" />
          {backToSignInLink}
        </header>
        <div className="legalContent">{renderLegalMarkdown(markdown)}</div>
        <footer className="legalFooter">{backToSignInLink}</footer>
      </article>
    </main>
  );
}

function AuthErrorAlert({ message }: { message: string | null }) {
  const alertRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    if (message) {
      alertRef.current?.focus();
    }
  }, [message]);

  if (!message) {
    return null;
  }

  return (
    <div className="authAlert" role="alert" aria-live="assertive" tabIndex={-1} ref={alertRef}>
      <strong>Authentication issue</strong>
      <span>{message}</span>
    </div>
  );
}

function PasswordField({
  error,
  helpText,
  id,
  label,
  onChange,
  placeholder,
  value,
}: {
  error?: string;
  helpText?: string;
  id: string;
  label: string;
  onChange: (value: string) => void;
  placeholder: string;
  value: string;
}) {
  const [isVisible, setIsVisible] = React.useState(false);
  const helpId = `${id}-help`;
  const errorId = `${id}-error`;

  return (
    <div className="fieldGroup">
      <label htmlFor={id}>{label}</label>
      <div className="passwordInputWrap">
        <input
          aria-describedby={[helpText ? helpId : null, error ? errorId : null].filter(Boolean).join(" ") || undefined}
          aria-invalid={Boolean(error)}
          autoComplete={id.includes("new") || id.includes("confirm") ? "new-password" : "current-password"}
          id={id}
          onChange={(event) => onChange(event.target.value)}
          placeholder={placeholder}
          type={isVisible ? "text" : "password"}
          value={value}
        />
        <button
          aria-label={isVisible ? "Hide password" : "Show password"}
          className="iconButton"
          onClick={() => setIsVisible((current) => !current)}
          type="button"
        >
          {isVisible ? "Hide" : "Show"}
        </button>
      </div>
      {helpText && (
        <p className="fieldHelp" id={helpId}>
          {helpText}
        </p>
      )}
      {error && (
        <p className="fieldError" id={errorId}>
          {error}
        </p>
      )}
    </div>
  );
}

function TeamsAuthButton({
  disabled,
  isBusy,
  label,
  onClick,
}: {
  disabled: boolean;
  isBusy: boolean;
  label: string;
  onClick: () => void;
}) {
  return (
    <button
      aria-busy={isBusy}
      className="teamsButton"
      disabled={disabled || isBusy}
      onClick={onClick}
      type="button"
    >
      {isBusy ? <Spinner /> : <img className="teamsGlyph" src={microsoftTeamsLogo} alt="" aria-hidden="true" />}
      <span>{isBusy ? "Signing you in with Microsoft Teams..." : label}</span>
    </button>
  );
}

function SignalTunerLogo({ className }: { className: string }) {
  return (
    <span className={`${className} logoSwap`}>
      <img className="logoSwapLight" src={signalTunerLogo} alt="SignalTuner" />
      <img className="logoSwapDark" src={signalTunerDarkLogo} alt="SignalTuner" />
    </span>
  );
}

function BrandPanel() {
  return (
    <aside className="brandPanel">
      <SignalTunerLogo className="brandLogo" />
      <div className="brandCopy">
        <h1>Better connections.<br />Better meetings.</h1>
        <p>
          Analyze device, workspace, network, and service telemetry to find and fix issues in Microsoft Teams.
        </p>
      </div>
      <div className="benefitList">
        <article>
          <strong>Real-time monitoring</strong>
          <span>Continuously track connectivity quality of all meeting participants.</span>
        </article>
        <article>
          <strong>Actionable insights</strong>
          <span>
            Identify issues impacting meeting quality and get comprehensive analysis reports with prioritized
            recommendations.
          </span>
        </article>
        <article>
          <strong>Teams native</strong>
          <span>A focus on the Microsoft Teams experience, including live incident reports.</span>
        </article>
      </div>
      <img className="meetingPreview meetingPreviewLight" src={loginDashboardPreview} alt="" aria-hidden="true" />
      <img className="meetingPreview meetingPreviewDark" src={analysisPreview} alt="" aria-hidden="true" />
    </aside>
  );
}

function AuthFooter({ onSupport }: { onSupport: () => void }) {
  return (
    <footer className="authFooter">
      <span>&copy; {getCurrentYear()} SignalTuner</span>
      <a href={PRIVACY_POLICY_PATH}>Privacy Policy</a>
      <a href={TERMS_OF_SERVICE_PATH}>Terms of Service</a>
      <button type="button" onClick={onSupport}>
        Support
      </button>
    </footer>
  );
}

function AuthLayout({ children, onSupport }: { children: React.ReactNode; onSupport: () => void }) {
  return (
    <main className="pageShell authShell">
      <div className="authLayout">
        <BrandPanel />
        <section className="authCard" aria-labelledby="auth-title">
          <SignalTunerLogo className="compactAuthLogo" />
          {children}
        </section>
      </div>
      <AuthFooter onSupport={onSupport} />
    </main>
  );
}

function LoginPage({
  error,
  busyState,
  isRunningInTeams,
  meetingContext,
  onCreateAccount,
  onEmailSignIn,
  onForgotPassword,
  onSupport,
  onTeamsSignIn,
}: {
  error: string | null;
  busyState: AuthBusyState;
  isRunningInTeams: boolean;
  meetingContext: TeamsMeetingContext | null;
  onCreateAccount: () => void;
  onEmailSignIn: (email: string, password: string) => Promise<void>;
  onForgotPassword: () => void;
  onSupport: () => void;
  onTeamsSignIn: () => Promise<void>;
}) {
  const [email, setEmail] = React.useState("");
  const [password, setPassword] = React.useState("");
  const [fieldErrors, setFieldErrors] = React.useState<{ email?: string; password?: string }>({});
  const emailRef = React.useRef<HTMLInputElement | null>(null);
  const passwordRef = React.useRef<HTMLInputElement | null>(null);
  const isBusy = busyState !== "idle";
  const teamsBusy = busyState === "teams-sso" || busyState === "auto-sso";

  const submitEmail = async (event: React.FormEvent) => {
    event.preventDefault();

    const nextErrors = {
      email: isValidEmail(email) ? undefined : "Enter a valid email address.",
      password: password ? undefined : "Enter your password.",
    };

    setFieldErrors(nextErrors);

    if (nextErrors.email) {
      emailRef.current?.focus();
      return;
    }

    if (nextErrors.password) {
      passwordRef.current?.focus();
      return;
    }

    await onEmailSignIn(email, password);
  };

  return (
    <AuthLayout onSupport={onSupport}>
      <div className="authHeader">
        <h1 id="auth-title">Welcome back</h1>
        <p>Sign in to your SignalTuner account.</p>
      </div>
      <AuthErrorAlert message={error} />
      {busyState === "auto-sso" && (
        <p className="authStatus" role="status" aria-live="polite">
          <Spinner /> Signing you in with Microsoft Teams...
        </p>
      )}
      {!isRunningInTeams && (
        <p className="authNotice">Teams SSO is available when this app is opened in Microsoft Teams.</p>
      )}
      {isRunningInTeams && !meetingContext && (
        <p className="authNotice">Open SignalTuner from inside a Teams meeting to join a meeting session after sign-in.</p>
      )}
      <TeamsAuthButton
        disabled={!isRunningInTeams || isBusy}
        isBusy={teamsBusy}
        label="Continue with Microsoft Teams"
        onClick={() => void onTeamsSignIn()}
      />
      <div className="authSeparator">
        <span>or</span>
      </div>
      <form className="authForm" aria-busy={busyState === "email-login"} onSubmit={(event) => void submitEmail(event)}>
        <div className="fieldGroup">
          <label htmlFor="login-email">Email</label>
          <input
            aria-describedby={fieldErrors.email ? "login-email-error" : undefined}
            aria-invalid={Boolean(fieldErrors.email)}
            autoComplete="email"
            id="login-email"
            onChange={(event) => setEmail(event.target.value)}
            placeholder="name@company.com"
            ref={emailRef}
            type="email"
            value={email}
          />
          {fieldErrors.email && (
            <p className="fieldError" id="login-email-error">
              {fieldErrors.email}
            </p>
          )}
        </div>
        <PasswordField
          error={fieldErrors.password}
          id="login-password"
          label="Password"
          onChange={setPassword}
          placeholder="Enter your password"
          value={password}
        />
        <button className="forgotLink" disabled={isBusy} onClick={onForgotPassword} type="button">
          Forgot password?
        </button>
        <button className="primaryButton fullWidthButton" disabled={isBusy} type="submit">
          {busyState === "email-login" ? <Spinner /> : null}
          <span>Sign in</span>
        </button>
      </form>
      <div className="authSwitch">
        <span>New to SignalTuner?</span>
        <button className="secondaryButton fullWidthButton" disabled={isBusy} onClick={onCreateAccount} type="button">
          Create account
        </button>
      </div>
    </AuthLayout>
  );
}

function ForgotPasswordPage({
  busyState,
  error,
  onRequestPasswordReset,
  onSignIn,
  onSupport,
}: {
  busyState: AuthBusyState;
  error: string | null;
  onRequestPasswordReset: (email: string) => Promise<boolean>;
  onSignIn: () => void;
  onSupport: () => void;
}) {
  const [email, setEmail] = React.useState("");
  const [fieldErrors, setFieldErrors] = React.useState<{ email?: string }>({});
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);
  const emailRef = React.useRef<HTMLInputElement | null>(null);
  const isBusy = busyState !== "idle";

  const submitResetRequest = async (event: React.FormEvent) => {
    event.preventDefault();
    setSuccessMessage(null);

    const nextErrors = {
      email: isValidEmail(email) ? undefined : "Enter a valid email address.",
    };

    setFieldErrors(nextErrors);

    if (nextErrors.email) {
      emailRef.current?.focus();
      return;
    }

    const wasSent = await onRequestPasswordReset(email);

    if (wasSent) {
      setEmail("");
      setSuccessMessage("If that email is registered, a password reset email has been sent.");
    }
  };

  return (
    <AuthLayout onSupport={onSupport}>
      <div className="authHeader">
        <h1 id="auth-title">Reset password</h1>
        <p>Enter your account email to receive a reset link.</p>
      </div>
      <AuthErrorAlert message={error} />
      {successMessage && (
        <p className="authNotice" role="status" aria-live="polite">
          {successMessage}
        </p>
      )}
      <form className="authForm" aria-busy={busyState === "password-reset"} onSubmit={(event) => void submitResetRequest(event)}>
        <div className="fieldGroup">
          <label htmlFor="forgot-password-email">Email</label>
          <input
            aria-describedby={fieldErrors.email ? "forgot-password-email-error" : undefined}
            aria-invalid={Boolean(fieldErrors.email)}
            autoComplete="email"
            id="forgot-password-email"
            onChange={(event) => setEmail(event.target.value)}
            placeholder="name@company.com"
            ref={emailRef}
            type="email"
            value={email}
          />
          {fieldErrors.email && (
            <p className="fieldError" id="forgot-password-email-error">
              {fieldErrors.email}
            </p>
          )}
        </div>
        <button className="primaryButton fullWidthButton" disabled={isBusy} type="submit">
          {busyState === "password-reset" ? <Spinner /> : null}
          <span>Send reset email</span>
        </button>
      </form>
      <button className="secondaryButton fullWidthButton" disabled={isBusy} onClick={onSignIn} type="button">
        Back to sign in
      </button>
    </AuthLayout>
  );
}

function SupportForm({
  defaultEmail,
  onSubmit,
}: {
  defaultEmail?: string | null;
  onSubmit: (request: { sender: string; subject: string; body: string }) => Promise<string>;
}) {
  const [sender, setSender] = React.useState(defaultEmail ?? "");
  const [subject, setSubject] = React.useState("Support request");
  const [body, setBody] = React.useState("");
  const [fieldErrors, setFieldErrors] = React.useState<{ sender?: string; subject?: string; body?: string }>({});
  const [statusMessage, setStatusMessage] = React.useState<string | null>(null);
  const [isSending, setIsSending] = React.useState(false);
  const senderRef = React.useRef<HTMLInputElement | null>(null);
  const subjectRef = React.useRef<HTMLInputElement | null>(null);
  const bodyRef = React.useRef<HTMLTextAreaElement | null>(null);
  const maxBodyLength = 1000;

  React.useEffect(() => {
    setSender(defaultEmail ?? "");
  }, [defaultEmail]);

  const submitSupportRequest = async (event: React.FormEvent) => {
    event.preventDefault();
    setStatusMessage(null);

    const trimmedSender = sender.trim();
    const trimmedSubject = subject.trim();
    const trimmedBody = body.trim();
    const nextErrors = {
      sender: isValidEmail(trimmedSender) ? undefined : "Enter a valid email address.",
      subject: trimmedSubject ? undefined : "Enter a subject.",
      body: trimmedBody ? undefined : "Enter a message.",
    };

    setFieldErrors(nextErrors);

    if (nextErrors.sender) {
      senderRef.current?.focus();
      return;
    }

    if (nextErrors.subject) {
      subjectRef.current?.focus();
      return;
    }

    if (nextErrors.body) {
      bodyRef.current?.focus();
      return;
    }

    setIsSending(true);

    try {
      const message = await onSubmit({ sender: trimmedSender, subject: trimmedSubject, body: trimmedBody });
      setSubject("Support request");
      setBody("");
      setStatusMessage(message);
    } catch (caught) {
      setStatusMessage(getFriendlyErrorMessage(caught, "SignalTuner support is temporarily unavailable. Please try again."));
    } finally {
      setIsSending(false);
    }
  };

  return (
    <form className="supportForm" aria-busy={isSending} onSubmit={(event) => void submitSupportRequest(event)}>
      <div className="fieldGroup">
        <label htmlFor="support-email">From</label>
        <input
          aria-describedby={fieldErrors.sender ? "support-email-error" : undefined}
          aria-invalid={Boolean(fieldErrors.sender)}
          autoComplete="email"
          id="support-email"
          maxLength={50}
          onChange={(event) => setSender(event.target.value)}
          placeholder="name@company.com"
          ref={senderRef}
          type="email"
          value={sender}
        />
        {fieldErrors.sender && (
          <p className="fieldError" id="support-email-error">
            {fieldErrors.sender}
          </p>
        )}
      </div>
      <div className="fieldGroup">
        <label htmlFor="support-subject">Subject</label>
        <input
          aria-describedby={fieldErrors.subject ? "support-subject-error" : undefined}
          aria-invalid={Boolean(fieldErrors.subject)}
          id="support-subject"
          maxLength={80}
          onChange={(event) => setSubject(event.target.value)}
          ref={subjectRef}
          type="text"
          value={subject}
        />
        {fieldErrors.subject && (
          <p className="fieldError" id="support-subject-error">
            {fieldErrors.subject}
          </p>
        )}
      </div>
      <div className="fieldGroup">
        <label htmlFor="support-message">Message</label>
        <textarea
          aria-describedby={fieldErrors.body ? "support-message-error support-message-count" : "support-message-count"}
          aria-invalid={Boolean(fieldErrors.body)}
          id="support-message"
          maxLength={maxBodyLength}
          onChange={(event) => setBody(event.target.value)}
          placeholder="Hello, I have a question about SignalTuner..."
          ref={bodyRef}
          value={body}
        />
        <span className="characterCount" id="support-message-count">
          {body.length} / {maxBodyLength}
        </span>
        {fieldErrors.body && (
          <p className="fieldError" id="support-message-error">
            {fieldErrors.body}
          </p>
        )}
      </div>
      {statusMessage && (
        <p className="supportStatus" role="status" aria-live="polite">
          {statusMessage}
        </p>
      )}
      <button className="primaryButton fullWidthButton" disabled={isSending} type="submit">
        {isSending ? <Spinner /> : null}
        <span>Send</span>
      </button>
    </form>
  );
}

function SupportAuthPage({
  defaultEmail,
  onSendSupportEmail,
  onSignIn,
  onSupport,
}: {
  defaultEmail?: string | null;
  onSendSupportEmail: (request: { sender: string; subject: string; body: string }) => Promise<string>;
  onSignIn: () => void;
  onSupport: () => void;
}) {
  return (
    <AuthLayout onSupport={onSupport}>
      <div className="authHeader">
        <h1 id="auth-title">Support</h1>
        <p>Send a message to support@signaltuner.com.</p>
      </div>
      <SupportForm defaultEmail={defaultEmail} onSubmit={onSendSupportEmail} />
      <button className="secondaryButton fullWidthButton" onClick={onSignIn} type="button">
        Back to sign in
      </button>
    </AuthLayout>
  );
}

function CreateAccountPage({
  busyState,
  error,
  isRunningInTeams,
  onEmailRegister,
  onSignIn,
  onSupport,
  onTeamsAccountCreate,
}: {
  busyState: AuthBusyState;
  error: string | null;
  isRunningInTeams: boolean;
  onEmailRegister: (email: string, password: string, firstName: string, lastName: string) => Promise<void>;
  onSignIn: () => void;
  onSupport: () => void;
  onTeamsAccountCreate: () => void;
}) {
  const [firstName, setFirstName] = React.useState("");
  const [lastName, setLastName] = React.useState("");
  const [email, setEmail] = React.useState("");
  const [password, setPassword] = React.useState("");
  const [confirmPassword, setConfirmPassword] = React.useState("");
  const [emailTermsAccepted, setEmailTermsAccepted] = React.useState(false);
  const [fieldErrors, setFieldErrors] = React.useState<{
    firstName?: string;
    lastName?: string;
    email?: string;
    password?: string;
    confirmPassword?: string;
    terms?: string;
  }>({});
  const firstNameRef = React.useRef<HTMLInputElement | null>(null);
  const lastNameRef = React.useRef<HTMLInputElement | null>(null);
  const emailRef = React.useRef<HTMLInputElement | null>(null);
  const termsRef = React.useRef<HTMLInputElement | null>(null);
  const isBusy = busyState !== "idle";
  const teamsBusy = busyState === "teams-sso" || busyState === "auto-sso";

  const submitRegistration = async (event: React.FormEvent) => {
    event.preventDefault();

    const nextErrors = {
      firstName: firstName.trim() ? undefined : "Enter your first name.",
      lastName: lastName.trim() ? undefined : "Enter your last name.",
      email: isValidEmail(email) ? undefined : "Enter a valid email address.",
      password: validatePassword(password),
      confirmPassword: password === confirmPassword ? undefined : "Passwords must match.",
      terms: emailTermsAccepted ? undefined : "Accept the Terms of Service and Privacy Policy to create an account.",
    };

    setFieldErrors(nextErrors);

    if (nextErrors.firstName) {
      firstNameRef.current?.focus();
      return;
    }

    if (nextErrors.lastName) {
      lastNameRef.current?.focus();
      return;
    }

    if (nextErrors.email) {
      emailRef.current?.focus();
      return;
    }

    if (nextErrors.terms) {
      termsRef.current?.focus();
      return;
    }

    if (nextErrors.password || nextErrors.confirmPassword) {
      return;
    }

    await onEmailRegister(email, password, firstName, lastName);
  };

  return (
    <AuthLayout onSupport={onSupport}>
      <div className="authHeader">
        <h1 id="auth-title">Create your account</h1>
        <p>Start monitoring and improving your Microsoft Teams meetings.</p>
      </div>
      <AuthErrorAlert message={error} />
      {!isRunningInTeams && (
        <p className="authNotice">Teams account creation is available when this app is opened in Microsoft Teams.</p>
      )}
      <div className="teamsCreateBlock">
        <TeamsAuthButton
          disabled={!isRunningInTeams || isBusy}
          isBusy={teamsBusy}
          label="Create with Microsoft Teams"
          onClick={onTeamsAccountCreate}
        />
        <p>Use the Microsoft account already signed in to Teams.</p>
      </div>
      <div className="authSeparator">
        <span>or create with email</span>
      </div>
      <form className="authForm" aria-busy={busyState === "email-register"} onSubmit={(event) => void submitRegistration(event)}>
        <div className="splitFields">
          <div className="fieldGroup">
            <label htmlFor="register-first-name">First name</label>
            <input
              aria-describedby={fieldErrors.firstName ? "register-first-name-error" : undefined}
              aria-invalid={Boolean(fieldErrors.firstName)}
              autoComplete="given-name"
              id="register-first-name"
              onChange={(event) => setFirstName(event.target.value)}
              placeholder="First name"
              ref={firstNameRef}
              type="text"
              value={firstName}
            />
            {fieldErrors.firstName && (
              <p className="fieldError" id="register-first-name-error">
                {fieldErrors.firstName}
              </p>
            )}
          </div>
          <div className="fieldGroup">
            <label htmlFor="register-last-name">Last name</label>
            <input
              aria-describedby={fieldErrors.lastName ? "register-last-name-error" : undefined}
              aria-invalid={Boolean(fieldErrors.lastName)}
              autoComplete="family-name"
              id="register-last-name"
              onChange={(event) => setLastName(event.target.value)}
              placeholder="Last name"
              ref={lastNameRef}
              type="text"
              value={lastName}
            />
            {fieldErrors.lastName && (
              <p className="fieldError" id="register-last-name-error">
                {fieldErrors.lastName}
              </p>
            )}
          </div>
        </div>
        <div className="fieldGroup">
          <label htmlFor="register-email">Email</label>
          <input
            aria-describedby={fieldErrors.email ? "register-email-error" : undefined}
            aria-invalid={Boolean(fieldErrors.email)}
            autoComplete="email"
            id="register-email"
            onChange={(event) => setEmail(event.target.value)}
            placeholder="name@company.com"
            ref={emailRef}
            type="email"
            value={email}
          />
          {fieldErrors.email && (
            <p className="fieldError" id="register-email-error">
              {fieldErrors.email}
            </p>
          )}
        </div>
        <PasswordField
          error={fieldErrors.password}
          helpText={PASSWORD_REQUIREMENT_TEXT}
          id="new-password"
          label="Password"
          onChange={setPassword}
          placeholder="Enter your password"
          value={password}
        />
        <PasswordField
          error={fieldErrors.confirmPassword}
          id="confirm-new-password"
          label="Confirm password"
          onChange={setConfirmPassword}
          placeholder="Enter your password"
          value={confirmPassword}
        />
        <label className="checkboxRow" htmlFor="email-terms">
          <input
            aria-describedby={fieldErrors.terms ? "email-terms-error" : undefined}
            aria-invalid={Boolean(fieldErrors.terms)}
            checked={emailTermsAccepted}
            id="email-terms"
            onChange={(event) => {
              setEmailTermsAccepted(event.target.checked);
              if (event.target.checked) {
                setFieldErrors((current) => ({ ...current, terms: undefined }));
              }
            }}
            ref={termsRef}
            type="checkbox"
          />
          <span>
            I agree to the <a href={TERMS_OF_SERVICE_PATH}>Terms of Service</a> and{" "}
            <a href={PRIVACY_POLICY_PATH}>Privacy Policy</a>.
          </span>
        </label>
        {fieldErrors.terms && (
          <p className="fieldError" id="email-terms-error">
            {fieldErrors.terms}
          </p>
        )}
        <button className="primaryButton fullWidthButton" disabled={isBusy} type="submit">
          {busyState === "email-register" ? <Spinner /> : null}
          <span>Create account</span>
        </button>
      </form>
      <div className="authSwitch">
        <span>Already have an account?</span>
        <button className="secondaryButton fullWidthButton" disabled={isBusy} onClick={onSignIn} type="button">
          Sign in
        </button>
      </div>
    </AuthLayout>
  );
}

function CompleteProfilePage({
  busyState,
  error,
  requireLegalConsent = false,
  secondaryActionLabel = "Sign out",
  user,
  onSubmit,
  onSupport,
  onSignOut,
}: {
  busyState: AuthBusyState;
  error: string | null;
  requireLegalConsent?: boolean;
  secondaryActionLabel?: string;
  user: CurrentUser;
  onSubmit: (firstName: string, lastName: string) => Promise<void>;
  onSupport: () => void;
  onSignOut: () => void;
}) {
  const [firstName, setFirstName] = React.useState(user.firstName ?? "");
  const [lastName, setLastName] = React.useState(user.lastName ?? "");
  const [termsAccepted, setTermsAccepted] = React.useState(false);
  const [fieldErrors, setFieldErrors] = React.useState<{ firstName?: string; lastName?: string; terms?: string }>({});
  const firstNameRef = React.useRef<HTMLInputElement | null>(null);
  const lastNameRef = React.useRef<HTMLInputElement | null>(null);
  const termsRef = React.useRef<HTMLInputElement | null>(null);
  const isBusy = busyState !== "idle";

  const submitProfile = async (event: React.FormEvent) => {
    event.preventDefault();

    const nextErrors = {
      firstName: firstName.trim() ? undefined : "Enter your first name.",
      lastName: lastName.trim() ? undefined : "Enter your last name.",
      terms: !requireLegalConsent || termsAccepted ? undefined : "Accept the Terms of Service and Privacy Policy to create an account.",
    };

    setFieldErrors(nextErrors);

    if (nextErrors.firstName) {
      firstNameRef.current?.focus();
      return;
    }

    if (nextErrors.lastName) {
      lastNameRef.current?.focus();
      return;
    }

    if (nextErrors.terms) {
      termsRef.current?.focus();
      return;
    }

    await onSubmit(firstName, lastName);
  };

  return (
    <AuthLayout onSupport={onSupport}>
      <div className="authHeader">
        <h1 id="auth-title">Finish your profile</h1>
        <p>Enter your name so teammates can recognize you in SignalTuner dashboards.</p>
      </div>
      <AuthErrorAlert message={error} />
      {user.email && <p className="authNotice">Signed in with Microsoft Teams as {user.email}.</p>}
      <form className="authForm" aria-busy={busyState === "teams-sso"} onSubmit={(event) => void submitProfile(event)}>
        <div className="splitFields">
          <div className="fieldGroup">
            <label htmlFor="profile-first-name">First name</label>
            <input
              aria-describedby={fieldErrors.firstName ? "profile-first-name-error" : undefined}
              aria-invalid={Boolean(fieldErrors.firstName)}
              autoComplete="given-name"
              id="profile-first-name"
              onChange={(event) => setFirstName(event.target.value)}
              placeholder="First name"
              ref={firstNameRef}
              type="text"
              value={firstName}
            />
            {fieldErrors.firstName && (
              <p className="fieldError" id="profile-first-name-error">
                {fieldErrors.firstName}
              </p>
            )}
          </div>
          <div className="fieldGroup">
            <label htmlFor="profile-last-name">Last name</label>
            <input
              aria-describedby={fieldErrors.lastName ? "profile-last-name-error" : undefined}
              aria-invalid={Boolean(fieldErrors.lastName)}
              autoComplete="family-name"
              id="profile-last-name"
              onChange={(event) => setLastName(event.target.value)}
              placeholder="Last name"
              ref={lastNameRef}
              type="text"
              value={lastName}
            />
            {fieldErrors.lastName && (
              <p className="fieldError" id="profile-last-name-error">
                {fieldErrors.lastName}
              </p>
            )}
          </div>
        </div>
        {requireLegalConsent && (
          <>
            <label className="checkboxRow" htmlFor="profile-terms">
              <input
                aria-describedby={fieldErrors.terms ? "profile-terms-error" : undefined}
                aria-invalid={Boolean(fieldErrors.terms)}
                checked={termsAccepted}
                id="profile-terms"
                onChange={(event) => {
                  setTermsAccepted(event.target.checked);
                  if (event.target.checked) {
                    setFieldErrors((current) => ({ ...current, terms: undefined }));
                  }
                }}
                ref={termsRef}
                type="checkbox"
              />
              <span>
                I agree to the <a href={TERMS_OF_SERVICE_PATH}>Terms of Service</a> and{" "}
                <a href={PRIVACY_POLICY_PATH}>Privacy Policy</a>.
              </span>
            </label>
            {fieldErrors.terms && (
              <p className="fieldError" id="profile-terms-error">
                {fieldErrors.terms}
              </p>
            )}
          </>
        )}
        <button className="primaryButton fullWidthButton" disabled={isBusy} type="submit">
          {busyState === "teams-sso" ? <Spinner /> : null}
          <span>Continue</span>
        </button>
      </form>
      <button className="secondaryButton fullWidthButton" disabled={isBusy} onClick={onSignOut} type="button">
        {secondaryActionLabel}
      </button>
    </AuthLayout>
  );
}

function ClientPrompt({
  activationCodeError,
  isLoading,
  onContinue,
  onRefresh,
  onSignOut,
}: {
  activationCodeError: string | null;
  isLoading: boolean;
  onContinue: () => void;
  onRefresh: () => Promise<void>;
  onSignOut: () => void;
}) {
  const downloadUrl = getDownloadUrl();
  const browserLaunchUrl = "https://www.signaltuner.com/Launch.html";

  const [rapidPollingEndsAt, setRapidPollingEndsAt] = React.useState<number | null>(null);
  const onRefreshRef = React.useRef(onRefresh);
  const refreshInFlightRef = React.useRef(false);

  React.useEffect(() => {
    onRefreshRef.current = onRefresh;
  }, [onRefresh]);

  const runRefresh = React.useCallback(async () => {
    if (refreshInFlightRef.current) {
      return;
    }

    refreshInFlightRef.current = true;

    try {
      await onRefreshRef.current();
    } finally {
      refreshInFlightRef.current = false;
    }
  }, []);

  const handleActivationCodeCopied = React.useCallback(() => {
    setRapidPollingEndsAt(Date.now() + CLIENT_PROMPT_COPIED_REFRESH_DURATION_MS);
    void runRefresh();
  }, [runRefresh]);

  const refreshIntervalMs =
    rapidPollingEndsAt === null ? CLIENT_PROMPT_REFRESH_INTERVAL_MS : CLIENT_PROMPT_COPIED_REFRESH_INTERVAL_MS;

  React.useEffect(() => {
    const intervalId = window.setInterval(() => {
      void runRefresh();
    }, refreshIntervalMs);

    return () => window.clearInterval(intervalId);
  }, [refreshIntervalMs, runRefresh]);

  React.useEffect(() => {
    if (rapidPollingEndsAt === null) {
      return;
    }

    const timeoutId = window.setTimeout(
      () => setRapidPollingEndsAt(null),
      Math.max(rapidPollingEndsAt - Date.now(), 0)
    );

    return () => window.clearTimeout(timeoutId);
  }, [rapidPollingEndsAt]);

  const setupSteps = [
    { icon: downloadStepIcon, isImage: true, title: "Download the SignalTuner desktop app", description: "Get the latest version for Windows." },
    {
      icon: personCheckIcon,
      isImage: true,
      title: "Open the app and enter your activation code, or sign in with email and password",
      description: "This links the app to your account.",
    },
    { icon: refreshStepIcon, isImage: true, isFlipped: true, title: "Return here and refresh your status", description: "We'll confirm when your device is connected." },
  ];

  const benefits = [
    "Live network telemetry during meetings",
    "Device insights for better visibility",
    "Better troubleshooting and faster resolutions",
  ];

  return (
    <section className="clientPrompt">
      <div className="clientPromptIntro">
        <SignalTunerLogo className="clientPromptLogo" />
        <h1>Connect the desktop client</h1>
        <p className="clientPromptLead">
          Download and activate the SignalTuner desktop app to share your live Wi-Fi, network, and device telemetry
          during meetings.
        </p>

        <div className="setupSteps">
          {setupSteps.map((step, index) => (
            <div className="setupStep" key={step.title}>
              <div className="setupStepRail" aria-hidden="true">
                <span className="setupStepNumber">{index + 1}</span>
                {index < setupSteps.length - 1 && <span className="setupStepConnector" />}
              </div>
              <div className="setupStepIcon">
                {index === 0 && downloadUrl ? (
                  <a className="setupStepIconButton" href={downloadUrl} aria-label="Download SignalTuner">
                    <img src={step.icon} alt="" />
                  </a>
                ) : index === 1 ? (
                  <a
                    className="setupStepIconButton"
                    href={browserLaunchUrl}
                    target="_blank"
                    rel="noreferrer"
                    aria-label="Open SignalTuner desktop client"
                  >
                    <img src={step.icon} alt="" />
                  </a>
                ) : index === 2 ? (
                  <button
                    className="setupStepIconButton"
                    disabled={isLoading}
                    onClick={() => void runRefresh()}
                    type="button"
                    aria-label="Refresh status"
                  >
                    {isLoading ? <Spinner /> : <img className="setupStepIconFlipped" src={step.icon} alt="" />}
                  </button>
                ) : step.isImage ? (
                  <img src={step.icon} alt="" />
                ) : step.icon}
              </div>
              <div className="setupStepCopy">
                <strong>{step.title}</strong>
                <p>{step.description}</p>
              </div>
            </div>
          ))}
        </div>

        <div className="clientBenefits">
          <h2>Why connect it?</h2>
          {benefits.map((benefit) => (
            <div className="benefitItem" key={benefit}>
              <span className="benefitCheck" aria-hidden="true">✓</span>
              <span>{benefit}</span>
            </div>
          ))}
        </div>

        <div className="clientInfoNote">
          <span className="infoGlyph" aria-hidden="true">i</span>
          <span>You can continue without device data, but meeting insights will be limited until your desktop client is connected.</span>
        </div>
      </div>

      <div className="activationPanel">
        <div className="activationPanelHeader">
          <span className="eyebrow">GET STARTED</span>
          <h2>Your activation code</h2>
          <p>Use this code in the desktop client.</p>
        </div>
        <TemporaryActivationCode onCopied={handleActivationCodeCopied} />
        {activationCodeError && <p className="activationError">{activationCodeError}</p>}

        <div className="desktopPreview" aria-label="Illustration of where to enter the activation code">
          <img className="desktopPreviewImage" src={clientPreview} alt="SignalTuner desktop client sign-in screen" />
        </div>

        <div className="clientActions">
          {downloadUrl ? (
            <a className="primaryButton buttonLink clientDownloadButton" href={downloadUrl}>
              <img className="actionIcon actionIconImage" src={downloadStepIcon} alt="" aria-hidden="true" /> Download SignalTuner
            </a>
          ) : (
            <span className="inlineNote">Desktop downloads are available for Windows and macOS.</span>
          )}
          <button className="secondaryButton clientActionButton" disabled={isLoading} onClick={() => void runRefresh()} type="button">
            {isLoading ? <Spinner /> : <span className="actionIcon" aria-hidden="true">↻</span>} Refresh status
          </button>
          <button className="secondaryButton clientActionButton clientTertiaryButton" disabled={isLoading} onClick={onContinue} type="button">
            Continue without syncing
          </button>
        </div>
        <button className="clientSignOut" disabled={isLoading} onClick={onSignOut} type="button">Sign out</button>
      </div>
    </section>
  );
}

function IncidentReportRow({
  incident,
  onSelect,
  statusClassName,
}: {
  incident: ServiceIncident;
  onSelect: () => void;
  statusClassName: string;
}) {
  const reportDate = incident.startedAt ? new Date(incident.startedAt).toLocaleString() : "No date";
  const iconGlyph = statusClassName === "statusOperational" ? "✓" : "!";

  return (
    <button className="compactIncident" onClick={onSelect} type="button">
      <span className={`incidentIcon ${statusClassName}`} aria-hidden="true">{iconGlyph}</span>
      <strong>{incident.title}</strong>
      <span className={`semanticBadge ${statusClassName}`}>{incident.status}</span>
      <time>{reportDate}</time>
    </button>
  );
}

function IncidentDetailModal({
  incident,
  onClose,
}: {
  incident: ServiceIncident;
  onClose: () => void;
}) {
  const incidentLink = incident.link || "https://status.cloud.microsoft/microsoft-365";

  return (
    <div
      className="incidentModalOverlay"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) {
          onClose();
        }
      }}
    >
      <section aria-labelledby="incident-detail-title" aria-modal="true" className="incidentModal" role="dialog">
        <header className="incidentModalHeader">
          <div>
            <span className="eyebrow">Incident report</span>
            <h2 id="incident-detail-title">{incident.title}</h2>
          </div>
          <button aria-label="Close incident details" className="modalCloseButton" onClick={onClose} type="button">
            Close
          </button>
        </header>
        <div className="incidentModalBody">
          <div className="incidentModalField">
            <span>Description</span>
            <p>{incident.impact || "No incident description is available."}</p>
          </div>
          <div className="incidentModalField">
            <span>Unique identifier</span>
            <p>{incident.uniqueIdentifier || "Not available"}</p>
          </div>
          <a className="incidentReportLink" href={incidentLink} rel="noreferrer" target="_blank">
            Microsoft Teams incident report
          </a>
        </div>
      </section>
    </div>
  );
}

function CreditRequiredModal({
  prompt,
  onClose,
  onOpenAccount,
}: {
  prompt: SubscriptionPrompt;
  onClose: () => void;
  onOpenAccount: () => void;
}) {
  React.useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        onClose();
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [onClose]);

  return (
    <div
      className="incidentModalOverlay"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) {
          onClose();
        }
      }}
    >
      <section aria-labelledby="credit-required-title" aria-modal="true" className="incidentModal creditRequiredModal" role="dialog">
        <header className="incidentModalHeader">
          <div>
            <span className="eyebrow">Account notification</span>
            <h2 id="credit-required-title">More credits required</h2>
          </div>
          <button aria-label="Dismiss credit notification" className="modalCloseButton" onClick={onClose} type="button">
            X
          </button>
        </header>
        <div className="incidentModalBody">
          <p className="creditRequiredMessage">
            This analysis requires {prompt.requiredCredits} credits. You have {prompt.availableCredits} available.
          </p>
          <div className="creditRequiredActions">
            <button className="primaryButton" onClick={onOpenAccount} type="button">
              Open Account
            </button>
            <button className="secondaryButton" onClick={onClose} type="button">
              Dismiss
            </button>
          </div>
        </div>
      </section>
    </div>
  );
}

function CompactSignalScoreDisplay({
  score,
  overallStatus,
  hasData,
  accessibleLabel = "Signal Score",
}: {
  score: number | null;
  overallStatus: AnalysisStatus | null;
  hasData: boolean;
  accessibleLabel?: string;
}) {
  const scoreIsAvailable = hasData && score !== null && score !== undefined && Number.isFinite(score);
  const displayScore = scoreIsAvailable ? Math.round(score) : null;
  const clampedScore = displayScore === null ? 0 : Math.max(0, Math.min(100, displayScore));
  const displayStatus = getSignalOverallStatusLabel(overallStatus, scoreIsAvailable ? score : null);
  const meterColor = getSignalStatusColor(displayStatus);

  return (
    <div
      className={`compactSignalScoreDisplay compactSignalScore-${displayStatus.toLowerCase()}`}
      aria-label={
        scoreIsAvailable
          ? `${accessibleLabel} ${displayScore} percent, overall status ${displayStatus}`
          : `${accessibleLabel} unavailable`
      }
    >
      <svg className="signal-score-meter compactSignalScoreMeter" viewBox="0 0 130 106" aria-hidden="true">
        <path
          className="compactSignalScoreOuter"
          d="M16.6779 103.938C-10.8555 68.9427 5.22696 6.00502 64.9035 3.64307C98.9156 3.64307 126.489 31.2165 126.489 65.2286C126.489 79.8935 121.363 93.3614 112.805 103.938"
        />
        <path
          className="compactSignalScoreTrack"
          d="M18.8746 84.5C17.3302 82.6052 15.9338 80.5851 14.7029 78.4573C10.789 71.6913 8.54889 63.8359 8.54889 55.4573C8.54889 30.0522 29.1438 9.45728 54.5489 9.45728C79.954 9.45728 100.549 30.0522 100.549 55.4573C100.549 63.8359 98.3088 71.6913 94.3949 78.4573C93.238 80.4573 91.9348 82.3621 90.5 84.157"
        />
        {scoreIsAvailable && (
          <path
            className="compactSignalScoreArc"
            d="M18.8746 84.5C17.3302 82.6052 15.9338 80.5851 14.7029 78.4573C10.789 71.6913 8.54889 63.8359 8.54889 55.4573C8.54889 30.0522 29.1438 9.45728 54.5489 9.45728C79.954 9.45728 100.549 30.0522 100.549 55.4573C100.549 63.8359 98.3088 71.6913 94.3949 78.4573C93.238 80.4573 91.9348 82.3621 90.5 84.157"
            pathLength={100}
            stroke={meterColor}
            strokeDasharray={`${clampedScore} 100`}
          />
        )}
      </svg>
      <span className="compactSignalScoreValue">
        {displayScore === null ? (
          "-"
        ) : (
          <>
            <span>{displayScore}</span>
            <span>%</span>
          </>
        )}
      </span>
      <span className="compactSignalScoreStatus">{displayStatus}</span>
    </div>
  );
}

function SignalScoreTrendChart({
  trend,
  isLoading,
  error,
}: {
  trend: SignalScoreTrendResponse | null;
  isLoading: boolean;
  error: string | null;
}) {
  const points = trend?.points.slice(0, 10) ?? [];
  const hasCompleteTrend = points.length === 10;

  const width = 320;
  const height = 142;
  const left = 34;
  const right = 10;
  const top = 12;
  const bottom = 28;
  const plotWidth = width - left - right;
  const plotHeight = height - top - bottom;
  const coordinates = hasCompleteTrend ? points.map((point, index) => {
    const score = Math.max(0, Math.min(100, point.averageScore));
    const color = getSignalStatusColor(getSignalDisplayStatusFromScore(score));

    return {
      x: left + (plotWidth / (points.length - 1)) * index,
      y: top + ((100 - score) / 100) * plotHeight,
      color,
      score,
      timestampUtc: point.timestampUtc,
    };
  }) : [];
  const polylinePoints = coordinates.map((point) => `${point.x},${point.y}`).join(" ");
  const areaPoints = `${left},${top + plotHeight} ${polylinePoints} ${left + plotWidth},${top + plotHeight}`;


  return (
    <div className="signalScoreTrend" aria-label="Last 10 Signal Scores">
      <div className="signalScoreTrendHeader">
        <p>Last 10 Signal Scores</p>
      </div>
      <svg className="signalScoreTrendSvg" role="img" viewBox={`0 0 ${width} ${height}`}>
        <title>Last 10 Signal Scores</title>
        {[0, 50, 100].map((tick) => {
          const y = top + ((100 - tick) / 100) * plotHeight;

          return (
            <g key={tick}>
              <line className="trendGridLine" x1={left} x2={width - right} y1={y} y2={y} />
              <text className="trendYAxisLabel" x={left - 8} y={y + 4}>
                {tick}
              </text>
            </g>
          );
        })}
        {hasCompleteTrend ? <polygon className="trendArea" points={areaPoints} /> : null}
        {hasCompleteTrend ? coordinates.slice(1).map((point, index) => {
          const previousPoint = coordinates[index];

          return (
            <line
              className="trendSegment"
              key={`${previousPoint.timestampUtc}-${point.timestampUtc}`}
              stroke={point.color}
              x1={previousPoint.x}
              x2={point.x}
              y1={previousPoint.y}
              y2={point.y}
            >
              <title>{formatTrendTooltip(point.timestampUtc, point.score)}</title>
            </line>
          );
        }) : null}
        {hasCompleteTrend ? coordinates.map((point) => (
          <circle className="trendPoint" cx={point.x} cy={point.y} fill={point.color} key={point.timestampUtc} r="3.4">
            <title>{formatTrendTooltip(point.timestampUtc, point.score)}</title>
          </circle>
        )) : null}
        {!hasCompleteTrend ? (
          <text className="trendNoDataLabel" textAnchor="middle" x={left + plotWidth / 2} y={top + plotHeight / 2 + 4}>
            No data
          </text>
        ) : null}
      </svg>
      {isLoading && !trend ? <p className="signalScoreTrendState">Loading Signal Score trend.</p> : null}
      {error ? <p className="signalScoreTrendState">{error}</p> : null}
    </div>
  );
}

function AccountPage({
  isAddingTestingCredit,
  onUpdateProfile,
  onUpdateEmail,
  onUpdatePassword,
  onAddTestingCredit,
  user,
}: {
  isAddingTestingCredit: boolean;
  onUpdateProfile: (firstName: string, lastName: string) => Promise<CurrentUser>;
  onUpdateEmail: (email: string) => Promise<CurrentUser>;
  onUpdatePassword: (newPassword: string) => Promise<string>;
  onAddTestingCredit: () => Promise<void>;
  user: CurrentUser;
}) {
  const [firstName, setFirstName] = React.useState(user.firstName ?? "");
  const [lastName, setLastName] = React.useState(user.lastName ?? "");
  const [email, setEmail] = React.useState(user.email ?? "");
  const [newPassword, setNewPassword] = React.useState("");
  const [confirmPassword, setConfirmPassword] = React.useState("");
  const [profileMessage, setProfileMessage] = React.useState<string | null>(null);
  const [profileError, setProfileError] = React.useState<string | null>(null);
  const [emailMessage, setEmailMessage] = React.useState<string | null>(null);
  const [emailError, setEmailError] = React.useState<string | null>(null);
  const [passwordMessage, setPasswordMessage] = React.useState<string | null>(null);
  const [passwordError, setPasswordError] = React.useState<string | null>(null);
  const [isUpdatingProfile, setIsUpdatingProfile] = React.useState(false);
  const [isUpdatingEmail, setIsUpdatingEmail] = React.useState(false);
  const [isUpdatingPassword, setIsUpdatingPassword] = React.useState(false);
  const [profileFieldErrors, setProfileFieldErrors] = React.useState<{ firstName?: string; lastName?: string }>({});
  const firstNameRef = React.useRef<HTMLInputElement | null>(null);
  const lastNameRef = React.useRef<HTMLInputElement | null>(null);
  const emailRef = React.useRef<HTMLInputElement | null>(null);
  const subscriptionPlan = user.subscriptionPlan?.trim() || "Free";

  const hasPassword = Boolean(user.hasPassword);

  React.useEffect(() => {
    setFirstName(user.firstName ?? "");
    setLastName(user.lastName ?? "");
    setEmail(user.email ?? "");
  }, [user.email, user.firstName, user.lastName]);

  const submitProfileUpdate = async (event: React.FormEvent) => {
    event.preventDefault();
    setProfileMessage(null);
    setProfileError(null);

    const nextErrors = {
      firstName: firstName.trim() ? undefined : "Enter your first name.",
      lastName: lastName.trim() ? undefined : "Enter your last name.",
    };

    setProfileFieldErrors(nextErrors);

    if (nextErrors.firstName) {
      firstNameRef.current?.focus();
      return;
    }

    if (nextErrors.lastName) {
      lastNameRef.current?.focus();
      return;
    }

    setIsUpdatingProfile(true);

    try {
      const updatedUser = await onUpdateProfile(firstName, lastName);
      setFirstName(updatedUser.firstName ?? firstName.trim());
      setLastName(updatedUser.lastName ?? lastName.trim());
      setProfileMessage("Display name updated successfully.");
    } catch (caught) {
      setProfileError(getFriendlyErrorMessage(caught, "Failed to update display name."));
    } finally {
      setIsUpdatingProfile(false);
    }
  };

  const submitEmailUpdate = async (event: React.FormEvent) => {
    event.preventDefault();
    setEmailMessage(null);
    setEmailError(null);

    if (!isValidEmail(email)) {
      setEmailError("Enter a valid email address.");
      emailRef.current?.focus();
      return;
    }

    if (!hasPassword) {
      setEmailError("Create a password before updating your email.");
      return;
    }

    setIsUpdatingEmail(true);

    try {
      const updatedUser = await onUpdateEmail(email);
      setEmail(updatedUser.email ?? email.trim());
      setEmailMessage("Email updated successfully.");
    } catch (caught) {
      setEmailError(getFriendlyErrorMessage(caught, "Failed to update email."));
    } finally {
      setIsUpdatingEmail(false);
    }
  };

  const submitPasswordUpdate = async (event: React.FormEvent) => {
    event.preventDefault();
    setPasswordMessage(null);

    const validationError = validatePassword(newPassword);

    if (validationError) {
      setPasswordError(validationError);
      return;
    }

    if (newPassword !== confirmPassword) {
      setPasswordError("Passwords must match.");
      return;
    }

    setIsUpdatingPassword(true);
    setPasswordError(null);

    try {
      const message = await onUpdatePassword(newPassword);
      setNewPassword("");
      setConfirmPassword("");
      setPasswordMessage(message);
    } catch (caught) {
      setPasswordError(getFriendlyErrorMessage(caught, "Failed to update password."));
    } finally {
      setIsUpdatingPassword(false);
    }
  };

  return (
    <section className="panel accountPagePanel">
      <div className="sectionTitleRow">
        <div>
          <h2>Account</h2>
        </div>
      </div>
      <div className="settingsGrid">
        <div className="accountControls">
          <form className="settingsSection" aria-label="Display name" onSubmit={(event) => void submitProfileUpdate(event)}>
            <h3>Display name</h3>
            <div className="splitFields">
              <div className="fieldGroup">
                <label htmlFor="account-first-name">First name</label>
                <input
                  aria-describedby={profileFieldErrors.firstName ? "account-first-name-error" : undefined}
                  aria-invalid={Boolean(profileFieldErrors.firstName)}
                  autoComplete="given-name"
                  id="account-first-name"
                  onChange={(event) => setFirstName(event.target.value)}
                  ref={firstNameRef}
                  type="text"
                  value={firstName}
                />
                {profileFieldErrors.firstName && (
                  <p className="fieldError" id="account-first-name-error">
                    {profileFieldErrors.firstName}
                  </p>
                )}
              </div>
              <div className="fieldGroup">
                <label htmlFor="account-last-name">Last name</label>
                <input
                  aria-describedby={profileFieldErrors.lastName ? "account-last-name-error" : undefined}
                  aria-invalid={Boolean(profileFieldErrors.lastName)}
                  autoComplete="family-name"
                  id="account-last-name"
                  onChange={(event) => setLastName(event.target.value)}
                  ref={lastNameRef}
                  type="text"
                  value={lastName}
                />
                {profileFieldErrors.lastName && (
                  <p className="fieldError" id="account-last-name-error">
                    {profileFieldErrors.lastName}
                  </p>
                )}
              </div>
            </div>
            {profileError && <p className="inlineError">{profileError}</p>}
            {profileMessage && (
              <p className="settingsStatus" role="status" aria-live="polite">
                {profileMessage}
              </p>
            )}
            <button className="primaryButton settingsSaveButton" disabled={isUpdatingProfile} type="submit">
              {isUpdatingProfile ? <Spinner /> : null}
              <span>Update name</span>
            </button>
          </form>

          <form className="settingsSection" aria-label="Account email" onSubmit={(event) => void submitEmailUpdate(event)}>
            <h3>Email</h3>
            <div className="fieldGroup">
              <label htmlFor="account-email">Email</label>
              <input
                aria-describedby={emailError ? "account-email-error" : undefined}
                aria-invalid={Boolean(emailError)}
                autoComplete="email"
                id="account-email"
                onChange={(event) => setEmail(event.target.value)}
                ref={emailRef}
                type="email"
                value={email}
              />
              {emailError && (
                <p className="fieldError" id="account-email-error">
                  {emailError}
                </p>
              )}
            </div>
            {!hasPassword && <p className="inlineNote">Create a password before updating your email.</p>}
            {emailMessage && (
              <p className="settingsStatus" role="status" aria-live="polite">
                {emailMessage}
              </p>
            )}
            <button className="primaryButton settingsSaveButton" disabled={isUpdatingEmail || !hasPassword} type="submit">
              {isUpdatingEmail ? <Spinner /> : null}
              <span>Update email</span>
            </button>
          </form>

          <form className="settingsSection" aria-label="Account password" onSubmit={(event) => void submitPasswordUpdate(event)}>
            <h3>Password</h3>
            <div className="splitFields">
              <PasswordField
                error={passwordError ?? undefined}
                helpText={PASSWORD_REQUIREMENT_TEXT}
                id="account-new-password"
                label={hasPassword ? "New password" : "Create password"}
                onChange={setNewPassword}
                placeholder={hasPassword ? "Enter a new password" : "Create a password"}
                value={newPassword}
              />
              <PasswordField
                error={newPassword && confirmPassword && newPassword !== confirmPassword ? "Passwords must match." : undefined}
                id="account-confirm-password"
                label="Confirm password"
                onChange={setConfirmPassword}
                placeholder="Confirm password"
                value={confirmPassword}
              />
            </div>
            {passwordMessage && (
              <p className="settingsStatus" role="status" aria-live="polite">
                {passwordMessage}
              </p>
            )}
            <button className="primaryButton settingsSaveButton" disabled={isUpdatingPassword} type="submit">
              {isUpdatingPassword ? <Spinner /> : null}
              <span>{hasPassword ? "Update password" : "Create password"}</span>
            </button>
          </form>
        </div>
        <aside className="settingsSection accountMetaSection">
          <div className="activationCodeBlock accountActivationCode">
            <span>Desktop client activation code</span>
            <TemporaryActivationCode />
          </div>
          <div className="fieldGroup">
            <label htmlFor="account-organization">Organization</label>
            <input id="account-organization" readOnly type="text" value="" />
          </div>
          <div className="fieldGroup">
            <label htmlFor="account-subscription">Subscription</label>
            <input id="account-subscription" readOnly type="text" value={subscriptionPlan} />
          </div>
          <div className="fieldGroup">
            <label htmlFor="account-credits">Credits</label>
            <div className="creditControl">
              <input id="account-credits" readOnly type="text" value={String(user.credits)} />
              <button
                className="secondaryButton creditIncrementButton"
                disabled={isAddingTestingCredit}
                onClick={() => void onAddTestingCredit()}
                type="button"
              >
                +1
              </button>
            </div>
          </div>
        </aside>
      </div>
    </section>
  );
}

function SettingsPage({
  themePreference,
  onThemePreferenceChange,
}: {
  themePreference: SignalTunerThemePreference;
  onThemePreferenceChange: (preference: SignalTunerThemePreference) => void;
}) {
  return (
    <section className="panel settingsPagePanel">
      <div className="sectionTitleRow">
        <div>
          <h2>Settings</h2>
        </div>
      </div>
      <div className="settingsGrid">
        <div className="settingsSection themePreferenceControl">
          <strong>Appearance</strong>
          <div className="segmentedControl" role="group" aria-label="SignalTuner appearance">
            <button
              aria-pressed={themePreference === "light"}
              className={themePreference === "light" ? "activeSegment" : ""}
              onClick={() => onThemePreferenceChange("light")}
              type="button"
            >
              Light
            </button>
            <button
              aria-pressed={themePreference === "dark"}
              className={themePreference === "dark" ? "activeSegment" : ""}
              onClick={() => onThemePreferenceChange("dark")}
              type="button"
            >
              Dark
            </button>
          </div>
        </div>
        <div className="settingsSection">
          <h3>Microsoft Teams</h3>
          <label className="settingsToggleRow" htmlFor="setting-auto-add-meetings">
            <input id="setting-auto-add-meetings" type="checkbox" />
            <span>Add SignalTuner to recurring Teams meetings</span>
          </label>
          <label className="settingsToggleRow" htmlFor="setting-post-activation-prompts">
            <input id="setting-post-activation-prompts" type="checkbox" />
            <span>Post activation prompts for participants without client data</span>
          </label>
          <label className="settingsToggleRow" htmlFor="setting-post-analysis-summary">
            <input id="setting-post-analysis-summary" type="checkbox" />
            <span>Post diagnostic summaries to meeting chat after full analysis</span>
          </label>
          <label className="settingsToggleRow" htmlFor="setting-post-service-incidents">
            <input id="setting-post-service-incidents" type="checkbox" />
            <span>Post active Microsoft Teams incident updates to meeting chat</span>
          </label>
          <label className="settingsToggleRow" htmlFor="setting-organizer-approval">
            <input defaultChecked id="setting-organizer-approval" type="checkbox" />
            <span>Require organizer approval before automated chat posts</span>
          </label>
          <button className="primaryButton settingsSaveButton" disabled type="button">
            Save settings
          </button>
        </div>
      </div>
    </section>
  );
}

function SupportPage({
  defaultEmail,
  onSendSupportEmail,
}: {
  defaultEmail?: string | null;
  onSendSupportEmail: (request: { sender: string; subject: string; body: string }) => Promise<string>;
}) {
  return (
    <section className="panel supportPagePanel">
      <div className="sectionTitleRow">
        <div>
          <h2>Support</h2>
          <p>Send a message to support@signaltuner.com.</p>
        </div>
      </div>
      <div className="supportPanelBody">
        <SupportForm defaultEmail={defaultEmail} onSubmit={onSendSupportEmail} />
      </div>
    </section>
  );
}

function Dashboard({
  analysis,
  apiBaseUrl,
  dashboard,
  activePage,
  error,
  isLoading,
  isAddingTestingCredit,
  onAddTestingCredit,
  onAnalyzeAll,
  onAnalyzeUser,
  onInvite,
  onNavigate,
  onDismissSubscriptionPrompt,
  onSendSupportEmail,
  onSignOut,
  onThemePreferenceChange,
  onUpdateProfile,
  onUpdateEmail,
  onUpdatePassword,
  sessionToken,
  subscriptionPrompt,
  themePreference,
}: {
  analysis: AnalysisResult | null;
  apiBaseUrl: string;
  dashboard: DashboardData;
  activePage: InAppPage;
  error: string | null;
  isLoading: boolean;
  isAddingTestingCredit: boolean;
  onAddTestingCredit: () => Promise<void>;
  onAnalyzeAll: () => Promise<void>;
  onAnalyzeUser: (targetUserId: number) => Promise<void>;
  onInvite: () => Promise<void>;
  onNavigate: (page: InAppPage) => void;
  onDismissSubscriptionPrompt: () => void;
  onSendSupportEmail: (request: { sender: string; subject: string; body: string }) => Promise<string>;
  onSignOut: () => void;
  onThemePreferenceChange: (preference: SignalTunerThemePreference) => void;
  onUpdateProfile: (firstName: string, lastName: string) => Promise<CurrentUser>;
  onUpdateEmail: (email: string) => Promise<CurrentUser>;
  onUpdatePassword: (newPassword: string) => Promise<string>;
  sessionToken: string;
  subscriptionPrompt: SubscriptionPrompt | null;
  themePreference: SignalTunerThemePreference;
}) {
  const [accountOpen, setAccountOpen] = React.useState(false);
  const accountMenuRef = React.useRef<HTMLDivElement | null>(null);
  const [recentIncidentsOpen, setRecentIncidentsOpen] = React.useState(false);
  const recentIncidentsRef = React.useRef<HTMLDivElement | null>(null);
  const [selectedIncident, setSelectedIncident] = React.useState<ServiceIncident | null>(null);
  const [expandedParticipantIds, setExpandedParticipantIds] = React.useState<Set<number>>(() => new Set());
  const [signalScoreTrends, setSignalScoreTrends] = React.useState<Record<number, SignalScoreTrendResponse>>({});
  const [signalScoreTrendLoadingIds, setSignalScoreTrendLoadingIds] = React.useState<Set<number>>(() => new Set());
  const [signalScoreTrendErrors, setSignalScoreTrendErrors] = React.useState<Record<number, string>>({});
  const [nowMs, setNowMs] = React.useState(Date.now());
  const user = dashboard.currentUser;
  const participants = React.useMemo(() => sortParticipantsByMeetingRole(dashboard.participants), [dashboard.participants]);
  const activeParticipants = participants.filter((participant) => participant.clientDataStatus === "active");
  const participantsWithActiveScores = activeParticipants.filter(
    (participant) => participant.signalScore !== null && Number.isFinite(participant.signalScore)
  );
  const activeIncidents =
    dashboard.teamsServiceHealth.activeIncidents.length > 0
      ? dashboard.teamsServiceHealth.activeIncidents
      : EXAMPLE_ACTIVE_INCIDENTS;
  const resolvedIncidents = dashboard.teamsServiceHealth.recentIncidents.filter(isResolvedIncident);
  const teamsStatus =
    dashboard.teamsServiceHealth.activeIncidents.length > 0 ? normalizeTeamsServiceStatus(dashboard.teamsServiceHealth) : "activeIncident";
  const teamsStatusMeta = getTeamsStatusMeta(teamsStatus);
  const connectedParticipantCount = activeParticipants.length;
  const totalParticipantCount = participants.length;
  const connectivityIssueCount = participants.filter((participant) => participant.clientDataStatus !== "active").length;
  const activeIncidentCount = dashboard.teamsServiceHealth.activeIncidents.length;
  const meetingScore =
    participantsWithActiveScores.length > 0
      ? Math.round(
          participantsWithActiveScores.reduce((sum, participant) => sum + (participant.signalScore ?? 0), 0) /
            participantsWithActiveScores.length
        )
      : null;
  const expandedAnalyzableParticipantIds = participants
    .filter(
      (participant) =>
        expandedParticipantIds.has(participant.userId) &&
        participant.clientDataStatus === "active" &&
        getAnalysisRemainingMs(participant, nowMs) > 0 &&
        Boolean(getParticipantTelemetry(analysis, participant.userId) || getParticipantLiveTelemetry(participant))
    )
    .map((participant) => participant.userId);
  const expandedAnalyzableParticipantKey = expandedAnalyzableParticipantIds.join(",");

  const toggleParticipantTelemetry = React.useCallback((participantId: number) => {
    setExpandedParticipantIds((current) => {
      const next = new Set(current);
      if (next.has(participantId)) {
        next.delete(participantId);
      } else {
        next.add(participantId);
      }

      return next;
    });
  }, []);

  const handleAnalyzeUser = React.useCallback(
    async (participantId: number) => {
      setExpandedParticipantIds((current) => new Set(current).add(participantId));
      await onAnalyzeUser(participantId);
    },
    [onAnalyzeUser]
  );

  const handleAnalyzeAll = React.useCallback(async () => {
    setExpandedParticipantIds(new Set(participants.map((participant) => participant.userId)));
    await onAnalyzeAll();
  }, [onAnalyzeAll, participants]);

  React.useEffect(() => {
    if (!participants.some((participant) => getAnalysisRemainingMs(participant, Date.now()) > 0)) {
      return;
    }

    const intervalId = window.setInterval(() => setNowMs(Date.now()), 1000);
    return () => window.clearInterval(intervalId);
  }, [participants]);

  React.useEffect(() => {
    if (!accountOpen) {
      return;
    }

    const handlePointerDown = (event: PointerEvent) => {
      if (!accountMenuRef.current?.contains(event.target as Node)) {
        setAccountOpen(false);
      }
    };

    window.addEventListener("pointerdown", handlePointerDown);
    return () => window.removeEventListener("pointerdown", handlePointerDown);
  }, [accountOpen]);

  React.useEffect(() => {
    if (!recentIncidentsOpen) {
      return;
    }

    const handlePointerDown = (event: PointerEvent) => {
      if (!recentIncidentsRef.current?.contains(event.target as Node)) {
        setRecentIncidentsOpen(false);
      }
    };

    window.addEventListener("pointerdown", handlePointerDown);
    return () => window.removeEventListener("pointerdown", handlePointerDown);
  }, [recentIncidentsOpen]);

  React.useEffect(() => {
    if (!selectedIncident) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setSelectedIncident(null);
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [selectedIncident]);

  React.useEffect(() => {
    if (!expandedAnalyzableParticipantKey || !apiBaseUrl || !sessionToken) {
      return;
    }

    let isCancelled = false;
    const participantIds = expandedAnalyzableParticipantKey.split(",").map(Number);

    const loadTrend = async (participantId: number) => {
      setSignalScoreTrendLoadingIds((current) => new Set(current).add(participantId));

      try {
        const data = normalizeSignalScoreTrend(
          await fetchJson<unknown>(
            `${apiBaseUrl}/api/TeamsMeetings/${dashboard.meetingSessionId}/signal-score-trend/${participantId}`,
            {
              headers: buildAuthHeaders(sessionToken),
            }
          )
        );

        if (!isCancelled) {
          setSignalScoreTrends((current) => ({ ...current, [participantId]: data }));
          setSignalScoreTrendErrors((current) => {
            const next = { ...current };
            delete next[participantId];
            return next;
          });
        }
      } catch {
        if (!isCancelled) {
          setSignalScoreTrendErrors((current) => ({
            ...current,
            [participantId]: "Signal Score trend is unavailable.",
          }));
        }
      } finally {
        if (!isCancelled) {
          setSignalScoreTrendLoadingIds((current) => {
            const next = new Set(current);
            next.delete(participantId);
            return next;
          });
        }
      }
    };

    const loadExpandedTrends = () => {
      participantIds.forEach((participantId) => void loadTrend(participantId));
    };

    loadExpandedTrends();
    const intervalId = window.setInterval(loadExpandedTrends, 10000);

    return () => {
      isCancelled = true;
      window.clearInterval(intervalId);
      setSignalScoreTrendLoadingIds((current) => {
        const next = new Set(current);
        participantIds.forEach((participantId) => next.delete(participantId));
        return next;
      });
    };
  }, [
    apiBaseUrl,
    dashboard.meetingSessionId,
    expandedAnalyzableParticipantKey,
    sessionToken,
  ]);

  return (
    <main className="pageShell dashboardShell">
      <header className="appHeader">
        <SignalTunerLogo className="appLogo" />
        <div className="accountMenuWrap" ref={accountMenuRef}>
          <button
            aria-expanded={accountOpen}
            aria-haspopup="menu"
            className="accountButton"
            onClick={() => setAccountOpen((current) => !current)}
            type="button"
          >
            <span className="avatarBubble" aria-hidden="true">
              {getInitials(user.displayName, user.email)}
            </span>
            <span>{user.displayName ?? user.email ?? "SignalTuner account"}</span>
            <span className="chevron chevronDown" aria-hidden="true" />
          </button>
          {accountOpen && (
            <div className="accountDropdown" role="menu">
              <div className="accountDropdownHeader">
                <strong>{user.displayName ?? user.email ?? "SignalTuner account"}</strong>
                <span>{user.email}</span>
              </div>
              <button
                className={activePage === "dashboard" ? "activeMenuItem" : ""}
                role="menuitem"
                type="button"
                onClick={() => {
                  onNavigate("dashboard");
                  setAccountOpen(false);
                }}
              >
                Dashboard
              </button>
              <button
                className={activePage === "account" ? "activeMenuItem" : ""}
                role="menuitem"
                type="button"
                onClick={() => {
                  onNavigate("account");
                  setAccountOpen(false);
                }}
              >
                Account
              </button>
              <button
                className={activePage === "settings" ? "activeMenuItem" : ""}
                role="menuitem"
                type="button"
                onClick={() => {
                  onNavigate("settings");
                  setAccountOpen(false);
                }}
              >
                Settings
              </button>
              <button
                className={activePage === "support" ? "activeMenuItem" : ""}
                role="menuitem"
                type="button"
                onClick={() => {
                  onNavigate("support");
                  setAccountOpen(false);
                }}
              >
                Support
              </button>
              <button role="menuitem" type="button" onClick={onSignOut}>
                Sign out
              </button>
            </div>
          )}
        </div>
      </header>

      {error && <section className="panel panelAlert">{error}</section>}

      {activePage === "account" && (
        <AccountPage
          isAddingTestingCredit={isAddingTestingCredit}
          onAddTestingCredit={onAddTestingCredit}
          onUpdateEmail={onUpdateEmail}
          onUpdatePassword={onUpdatePassword}
          onUpdateProfile={onUpdateProfile}
          user={user}
        />
      )}

      {activePage === "settings" && (
        <SettingsPage themePreference={themePreference} onThemePreferenceChange={onThemePreferenceChange} />
      )}

      {activePage === "support" && (
        <SupportPage defaultEmail={user.email} onSendSupportEmail={onSendSupportEmail} />
      )}

      {activePage === "dashboard" && (
        <>
      <section className="panel meetingHealthPanel">
        <div className="meetingScoreDisplay">
          <span className="meetingScoreLabel">Meeting Score</span>
          <CompactSignalScoreDisplay
            accessibleLabel="Meeting Score"
            score={meetingScore}
            overallStatus={null}
            hasData={meetingScore !== null}
          />
        </div>
        <div className="meetingHealthMetrics">
          <MeetingHealthMetric
            icon={
              <svg className="meetingHealthGroupIcon" viewBox="0 0 24 24" focusable="false">
                <circle cx="12" cy="8.5" r="2.1" />
                <path d="M16.2 16.4a4.2 4.2 0 0 0-8.4 0Z" />
                <circle cx="18.2" cy="8.5" r="2.1" />
                <path d="M17.6 16.4h4.1a4.2 4.2 0 0 0-6.2-3.7" />
                <circle cx="5.8" cy="8.5" r="2.1" />
                <path d="M6.4 16.4H2.3a4.2 4.2 0 0 1 6.2-3.7" />
              </svg>
            }
            label={<>Meeting Participants<br />Connected</>}
            tone={totalParticipantCount > 0 && connectedParticipantCount === totalParticipantCount ? "good" : connectedParticipantCount === 0 ? "poor" : "warning"}
            value={`${connectedParticipantCount} / ${totalParticipantCount}`}
          />
          <MeetingHealthMetric
            icon={
              activeIncidentCount > 0 ? (
                <svg className="meetingHealthIncidentIcon" viewBox="0 0 64 64" focusable="false">
                  <path d="M32.427 7.987c2.183.124 4 1.165 5.096 3.281l17.936 36.208c1.739 3.66-.954 8.585-5.373 8.656H13.967c-4.022-.064-7.322-4.631-5.352-8.696l18.271-36.207c.342-.65.498-.838.793-1.179 1.186-1.375 2.483-2.111 4.748-2.063Zm-.295 3.997c-.687.034-1.316.419-1.659 1.017-6.312 11.979-12.397 24.081-18.301 36.267-.546 1.225.391 2.797 1.762 2.863 12.06.195 24.125.195 36.185 0 1.325-.064 2.321-1.584 1.769-2.85-5.793-12.184-11.765-24.286-17.966-36.267-.366-.651-.903-1.042-1.79-1.03Z" />
                  <path d="M33.631 40.581h-3.348l-.368-16.449h4.1l-.384 16.449Zm-3.828 5.03c0-.609.197-1.113.592-1.514.396-.4.935-.601 1.618-.601.684 0 1.223.201 1.618.601.395.401.593.905.593 1.514 0 .587-.193 1.078-.577 1.473-.385.395-.929.593-1.634.593-.705 0-1.249-.198-1.634-.593-.384-.395-.576-.886-.576-1.473Z" />
                </svg>
              ) : (
                <svg className="meetingHealthIncidentIcon" viewBox="0 0 24 24" focusable="false">
                  <path d="M12 2C6.5 2 2 6.5 2 12s4.5 10 10 10 10-4.5 10-10S17.5 2 12 2Zm0 18c-4.5 0-8-3.5-8-8s3.5-8 8-8 8 3.5 8 8-3.5 8-8 8Z" />
                  <path d="m9.8 16.8-3.7-3.6 1.4-1.5L9.8 14l5.7-6.1L17 9.3l-7.2 7.5Z" />
                </svg>
              )
            }
            label={<>Microsoft Teams<br />Active Incidents</>}
            tone={activeIncidentCount === 0 ? "good" : "warning"}
            value={String(activeIncidentCount)}
          />
          <MeetingHealthMetric
            icon={
              <svg className="meetingHealthConnectivityIcon" viewBox="0 0 56 56" focusable="false">
                {connectivityIssueCount > 0 ? (
                  <path d="M33.7169 50.6051C45.9141 50.6051 56 40.4968 56 28.2994 56 16.1245 45.892 5.9937 33.6944 5.9937 22.418 5.9937 12.9611 14.6419 11.5909 25.5365c.584 0 1.1456.0449 1.7072.1347.6963.0899 1.3927.2471 2.0665.4493 1.0558-9.1873 8.828-16.3081 18.3298-16.3081 10.2654 0 18.4868 8.2439 18.5093 18.487 0 4.7846-1.7743 9.0975-4.6947 12.3771-3.3919-2.8304-8.4684-4.7172-13.8146-4.7172-2.5159 0-5.3686.5391-8.0193 1.4825.1573.8536.2471 1.7297.2471 2.6057 0 2.9427-.921 5.7056-2.4709 8.0193 3.0774 1.5948 6.5816 2.5383 10.2656 2.5383Zm-.0225-18.5095c4.3129 0 7.57-3.7288 7.57-8.4236 0-4.4251-3.3245-8.2663-7.57-8.2663-4.223 0-7.57 3.8412-7.57 8.2663 0 4.6948 3.2796 8.4236 7.57 8.4236ZM11.4112 51.4587c6.2671 0 11.4112-5.1215 11.4112-11.4112 0-6.2447-5.1441-11.4112-11.4112-11.4112C5.1665 28.6363 0 33.8028 0 40.0475c0 6.2897 5.1665 11.4112 11.4112 11.4112Zm0-9.6141c-.7413 0-1.2804-.4942-1.3029-1.2355l-.1797-6.1998c-.0224-.876.5841-1.4825 1.4826-1.4825.9209 0 1.505.6065 1.4825 1.4825l-.1797 6.1998c-.0224.7413-.5391 1.2355-1.3028 1.2355Zm0 5.3686c-1.0109 0-1.842-.8311-1.842-1.842 0-1.0108.8311-1.8419 1.842-1.8419 1.0333 0 1.8419.8311 1.8419 1.8419 0 1.0109-.8086 1.842-1.8419 1.842Z" />
                ) : (
                  <path d="M33.7169 50.6051C45.9141 50.6051 56 40.4968 56 28.2994 56 16.1245 45.892 5.9937 33.6944 5.9937 22.418 5.9937 12.9611 14.6419 11.5909 25.5365c.584 0 1.1456.0449 1.7072.1347.6963.0899 1.3927.2471 2.0665.4493 1.0558-9.1873 8.828-16.3081 18.3298-16.3081 10.2654 0 18.4868 8.2439 18.5093 18.487 0 4.7846-1.7743 9.0975-4.6947 12.3771-3.3919-2.8304-8.4684-4.7172-13.8146-4.7172-2.5159 0-5.3686.5391-8.0193 1.4825.1573.8536.2471 1.7297.2471 2.6057 0 2.9427-.921 5.7056-2.4709 8.0193 3.0774 1.5948 6.5816 2.5383 10.2656 2.5383Zm-.0225-18.5095c4.3129 0 7.57-3.7288 7.57-8.4236 0-4.4251-3.3245-8.2663-7.57-8.2663-4.223 0-7.57 3.8412-7.57 8.2663 0 4.6948 3.2796 8.4236 7.57 8.4236ZM11.4112 51.4587c6.2671 0 11.4112-5.1215 11.4112-11.4112 0-6.2447-5.1441-11.4112-11.4112-11.4112C5.1665 28.6363 0 33.8028 0 40.0475c0 6.2897 5.1665 11.4112 11.4112 11.4112Zm-1.3478-4.7172c-.3594 0-.8087-.1573-1.1007-.4717L4.6498 41.5301c-.1572-.1797-.2695-.5616-.2695-.8761 0-.7637.6065-1.3702 1.3702-1.3702.4493 0 .8087.2022 1.0558.4717l3.1897 3.4593 5.9527-8.2439c.2471-.3594.6514-.6065 1.1456-.6065.7413 0 1.3927.584 1.3927 1.3478 0 .2471-.1123.5391-.3145.8311l-6.9635 9.6815c-.2246.3145-.6739.5167-1.1456.5167Z" />
                )}
              </svg>
            }
            label="Participants experiencing connectivity issues"
            tone={connectivityIssueCount === 0 ? "good" : connectedParticipantCount === 0 ? "poor" : "warning"}
            value={String(connectivityIssueCount)}
          />
        </div>
        <div className="healthActions">
          <button className="primaryButton" disabled={isLoading || activeParticipants.length === 0} onClick={handleAnalyzeAll} type="button">
            Run full analysis
          </button>
          <p>Deep diagnostic analysis of all participants and network path.</p>
        </div>
      </section>

      <section className="panel participantsPanel">
        <div className="sectionTitleRow">
          <div>
            <h2>Meeting Participants</h2>
          </div>
          <div className="buttonRow inlineButtons">
            <button className="secondaryButton" disabled={isLoading} onClick={onInvite} type="button">
              Invite
            </button>
          </div>
        </div>

        <div className="tableWrap">
          <table className="participantTable">
            <colgroup>
              <col className="expandColumn" />
              <col className="participantColumn" />
              <col className="scoreColumn" />
              <col className="signalColumn" />
              <col className="signalColumn" />
              <col className="signalColumn" />
              <col className="actionsColumn" />
            </colgroup>
            <thead>
              <tr>
                <th aria-label="Expand participant telemetry"></th>
                <th>Participant</th>
                <th><span className="scoreHeaderLabel">Signal Score</span></th>
                <th className="centeredSignalColumn">
                  <span className="participantHeaderLabel">
                    <img src={deviceDivIcon} alt="" aria-hidden="true" />
                    <span>Device</span>
                  </span>
                </th>
                <th className="centeredSignalColumn">
                  <span className="participantHeaderLabel">
                    <img src={workspaceDivIcon} alt="" aria-hidden="true" />
                    <span>Workspace</span>
                  </span>
                </th>
                <th className="centeredSignalColumn">
                  <span className="participantHeaderLabel">
                    <img src={networkDivIcon} alt="" aria-hidden="true" />
                    <span>Network</span>
                  </span>
                </th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {participants.map((participant) => {
                const analysisRemainingMs = getAnalysisRemainingMs(participant, nowMs);
                const hasActiveAnalysisSession = analysisRemainingMs > 0;
                const hasData = participant.clientDataStatus === "active";
                const isExpanded = expandedParticipantIds.has(participant.userId);
                const telemetry = hasActiveAnalysisSession
                  ? getParticipantTelemetry(analysis, participant.userId) ?? getParticipantLiveTelemetry(participant)
                  : null;
                const issues = hasActiveAnalysisSession ? getParticipantIssues(analysis, participant.userId) : [];
                const analysisCoversParticipant = Boolean(telemetry);

                return (
                  <React.Fragment key={participant.userId}>
                    <tr>
                      <td className="expandCell">
                        <button
                          aria-expanded={isExpanded}
                          aria-label={`${isExpanded ? "Collapse" : "Expand"} telemetry for ${getParticipantName(participant)}`}
                          className="expandButton"
                          onClick={() => toggleParticipantTelemetry(participant.userId)}
                          type="button"
                        >
                          <span className={`chevron ${isExpanded ? "chevronUp" : "chevronRight"}`} aria-hidden="true" />
                        </button>
                      </td>
                      <td>
                        <div className="participantIdentity">
                          <span className="avatarBubble smallAvatar" aria-hidden="true">
                            {getInitials(participant.displayName, participant.email)}
                          </span>
                          <div>
                            <strong className="participantName" title={participant.email ?? getParticipantName(participant)}>
                              {getParticipantName(participant)}
                            </strong>
                            <span>{formatParticipantMeetingRole(participant.meetingRole)}</span>
                          </div>
                        </div>
                      </td>
                      <td className="scoreCell">
                        <CompactSignalScoreDisplay
                          score={participant.signalScore}
                          overallStatus={participant.overallStatus}
                          hasData={hasData}
                        />
                      </td>
                      <td className="statusLabelCell">
                        <span className={`telemetryStatusLabel telemetryStatus-${getSignalStatusLabel(participant.deviceStatus, hasData ? participant.signalScore : null).toLowerCase()}`}>
                          {getSignalStatusLabel(participant.deviceStatus, hasData ? participant.signalScore : null)}
                        </span>
                      </td>
                      <td className="statusLabelCell">
                        <span className={`telemetryStatusLabel telemetryStatus-${getSignalStatusLabel(participant.workspaceStatus, hasData ? participant.signalScore : null).toLowerCase()}`}>
                          {getSignalStatusLabel(participant.workspaceStatus, hasData ? participant.signalScore : null)}
                        </span>
                      </td>
                      <td className="statusLabelCell">
                        <span className={`telemetryStatusLabel telemetryStatus-${getSignalStatusLabel(participant.networkStatus, hasData ? participant.signalScore : null).toLowerCase()}`}>
                          {getSignalStatusLabel(participant.networkStatus, hasData ? participant.signalScore : null)}
                        </span>
                      </td>
                      <td>
                        {hasActiveAnalysisSession ? (
                          <span className="analysisSessionStatus">
                            <span>Analysis in progress</span>
                            <time>{formatAnalysisCountdown(analysisRemainingMs)}</time>
                          </span>
                        ) : hasData ? (
                          <button className="secondaryButton compactAction" disabled={isLoading} onClick={() => handleAnalyzeUser(participant.userId)} type="button">
                            Analyze
                          </button>
                        ) : (
                          <button className="secondaryButton compactAction" disabled={isLoading} onClick={onInvite} type="button">
                            Prompt
                          </button>
                        )}
                      </td>
                    </tr>
                    {isExpanded && (
                      <ParticipantTelemetryDetail
                        hasActiveAnalysisSession={hasActiveAnalysisSession}
                        hasData={hasData || hasActiveAnalysisSession}
                        issues={issues}
                        signalScoreTrend={analysisCoversParticipant ? signalScoreTrends[participant.userId] ?? null : null}
                        signalScoreTrendError={analysisCoversParticipant ? signalScoreTrendErrors[participant.userId] ?? null : null}
                        signalScoreTrendLoading={analysisCoversParticipant && signalScoreTrendLoadingIds.has(participant.userId)}
                        telemetry={telemetry}
                      />
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </section>

      <section className="panel statusIncidentsPanel" aria-label="Microsoft Teams service health and incidents">
        <div className="serviceStatusCard">
          <img className="teamsMark" src={microsoftTeamsLogo} alt="Microsoft Teams" />
          <div className="serviceStatusContent">
            <div className="cardTitleRow">
              <h2>Microsoft Teams</h2>
              <span className={`semanticBadge ${teamsStatusMeta.className}`}>{teamsStatusMeta.label}</span>
            </div>
            <p>{teamsStatusMeta.description}</p>
            <a className="sourceLink" href="https://www.signaltuner.com/teamsstatus" target="_blank" rel="noreferrer">
              View Microsoft 365 Service Health
            </a>
          </div>
        </div>

        <div className="incidentsCard">
          <div className="cardTitleRow">
            <h2>Incident reports</h2>
          </div>
          {activeIncidents.length > 0 ? (
            <div className={`compactIncidentList activeIncidentList ${activeIncidents.length > 3 ? "scrollableIncidentList" : ""}`}>
              {activeIncidents.map((incident) => (
                <IncidentReportRow
                  incident={incident}
                  key={incident.incidentId}
                  onSelect={() => setSelectedIncident(incident)}
                  statusClassName={teamsStatusMeta.className}
                />
              ))}
            </div>
          ) : (
            <p className="emptyState">No active Microsoft Teams incidents are currently reported.</p>
          )}
          <div className="semanticIncidentDropdown" ref={recentIncidentsRef}>
            <button
              aria-controls="recent-incident-menu"
              aria-expanded={recentIncidentsOpen}
              className="recentIncidentsToggle"
              onClick={() => setRecentIncidentsOpen((current) => !current)}
              type="button"
            >
              <span>Resolved incidents</span>
              <span className={`chevron ${recentIncidentsOpen ? "chevronUp" : "chevronDown"}`} aria-hidden="true" />
            </button>
            {recentIncidentsOpen && (
              <div className="compactIncidentList recentIncidentList semanticIncidentMenu" id="recent-incident-menu" role="menu">
                {resolvedIncidents.length > 0 ? (
                  resolvedIncidents.map((incident) => (
                    <IncidentReportRow
                      incident={incident}
                      key={incident.incidentId}
                      onSelect={() => setSelectedIncident(incident)}
                      statusClassName={incident.status.toLowerCase() === "active" ? "statusIncident" : "statusOperational"}
                    />
                  ))
                ) : (
                  <p className="emptyState">No incident reports are available.</p>
                )}
              </div>
            )}
          </div>
        </div>
      </section>
      {selectedIncident ? (
        <IncidentDetailModal incident={selectedIncident} onClose={() => setSelectedIncident(null)} />
      ) : null}
        </>
      )}

      {activePage === "dashboard" && subscriptionPrompt && (
        <CreditRequiredModal
          prompt={subscriptionPrompt}
          onClose={onDismissSubscriptionPrompt}
          onOpenAccount={() => {
            onDismissSubscriptionPrompt();
            setAccountOpen(false);
            onNavigate("account");
          }}
        />
      )}
    </main>
  );
}

function MeetingHealthMetric({
  icon,
  label,
  tone,
  value,
}: {
  icon: React.ReactNode;
  label: React.ReactNode;
  tone: "good" | "warning" | "poor";
  value: string;
}) {
  return (
    <div className="meetingHealthMetric">
      <span className={`meetingHealthMetricIcon meetingHealthMetric-${tone}`} aria-hidden="true">
        {icon}
      </span>
      <strong className={`meetingHealthMetricValue meetingHealthMetric-${tone}`}>{value}</strong>
      <span className="meetingHealthMetricLabel">{label}</span>
    </div>
  );
}

function ParticipantTelemetryDetail({
  hasActiveAnalysisSession,
  hasData,
  issues,
  signalScoreTrend,
  signalScoreTrendError,
  signalScoreTrendLoading,
  telemetry,
}: {
  hasActiveAnalysisSession: boolean;
  hasData: boolean;
  issues: Issue[];
  signalScoreTrend: SignalScoreTrendResponse | null;
  signalScoreTrendError: string | null;
  signalScoreTrendLoading: boolean;
  telemetry: TelemetryRecord | null;
}) {
  const recommendation = hasActiveAnalysisSession
    ? ""
    : issues[0]?.recommendation ?? (hasData ? "" : "Prompt the participant to activate the local client.");
  const cpu = getTelemetryNumber(telemetry, ["cpu", "cpuUsage", "signal_cpu", "cpu_percent", "SignalCPU"]);
  const memory = getTelemetryNumber(telemetry, ["memory", "memoryUsage", "signal_memory", "memory_percent", "SignalMemory"]);
  const wifiStrength = getTelemetryNumber(telemetry, ["wifiStrength", "wiFiStrength", "signal_wifi_strength", "wifi_strength", "SignalWifiStrength"]);
  const download = getTelemetryNumber(telemetry, ["downloadSpeed", "signal_download_speed", "download_speed", "SignalDownloadSpeed"]);
  const upload = getTelemetryNumber(telemetry, ["uploadSpeed", "signal_upload_speed", "upload_speed", "SignalUploadSpeed"]);
  const latency = getTelemetryNumber(telemetry, ["latency", "ping", "signal_latency", "latencyMs", "SignalPing"]);
  const packetLoss = getTelemetryNumber(telemetry, ["packetLoss", "signal_packet_loss", "packet_loss", "SignalPacketLoss"]);
  const currentNetwork = getTelemetryValue(telemetry, ["currentNetwork", "signal_current_network", "current_network", "SignalCurrentNetwork"]);
  const frequency = getTelemetryValue(telemetry, ["wifiBand", "signal_wifi_band", "wifi_band", "frequency", "networkFrequency", "SignalWifiBand"]);
  const vpn = getTelemetryValue(telemetry, ["vpn", "vpnStatus", "vpnDetected", "signal_vpn", "vpn_status", "signal_vpn_detected", "SignalVpnDetected"]);
  const telemetryInsights = buildTelemetryInsights(telemetry);
  const processor = getTelemetryValue(telemetry, ["processor", "signal_processor", "cpuProcessor", "deviceProcessor", "SignalProcessor"]);
  const cores = getTelemetryValue(telemetry, ["cores", "signal_cores", "cpuCores", "deviceCores", "SignalCores"]);
  const telemetryGroups = [
    {
      title: "Device",
      metrics: [
        { label: "CPU", icon: deviceCpuIcon, value: formatTelemetryMeasurement(telemetry, ["cpu", "cpuUsage", "signal_cpu", "cpu_percent", "SignalCPU"], "%"), color: getCpuMetricColor(cpu) },
        { label: "Memory", icon: deviceMemoryIcon, value: formatTelemetryMeasurement(telemetry, ["memory", "memoryUsage", "signal_memory", "memory_percent", "SignalMemory"], "%"), color: getMemoryMetricColor(memory) },
        { label: "Processor", icon: deviceProcessorIcon, value: processor, color: processor === "No data" ? signalMetricColors.unknown : signalMetricColors.excellent },
        { label: "Cores", icon: deviceCoresIcon, value: cores, color: cores === "No data" ? signalMetricColors.unknown : signalMetricColors.excellent },
      ],
    },
    {
      title: "Workspace",
      metrics: [
        { label: "Wi-Fi Strength", icon: workspaceWifiStrengthIcon, value: formatTelemetryMeasurement(telemetry, ["wifiStrength", "wiFiStrength", "signal_wifi_strength", "wifi_strength", "SignalWifiStrength"], "%"), color: getWifiStrengthMetricColor(wifiStrength) },
        { label: "Current Network", icon: workspaceCurrentNetworkIcon, value: currentNetwork, color: currentNetwork === "No data" ? signalMetricColors.unknown : signalMetricColors.excellent },
        { label: "Frequency", icon: workspaceNetworkFrequencyIcon, value: frequency, color: getWifiBandMetricColor(frequency, wifiStrength) },
        { label: "VPN", icon: workspaceVpnIcon, value: vpn, color: getVpnMetricColor(vpn) },
      ],
    },
    {
      title: "Network",
      metrics: [
        { label: "Download", icon: networkDownloadIcon, value: formatTelemetryMeasurement(telemetry, ["downloadSpeed", "signal_download_speed", "download_speed", "SignalDownloadSpeed"], " Mbps"), color: getDownloadMetricColor(download) },
        { label: "Upload", icon: networkUploadIcon, value: formatTelemetryMeasurement(telemetry, ["uploadSpeed", "signal_upload_speed", "upload_speed", "SignalUploadSpeed"], " Mbps"), color: getUploadMetricColor(upload) },
        { label: "Latency", icon: networkLatencyIcon, value: formatTelemetryMeasurement(telemetry, ["latency", "ping", "signal_latency", "latencyMs", "SignalPing"], " ms"), color: getLatencyMetricColor(latency) },
        { label: "Packet loss", icon: networkPacketLossIcon, value: formatTelemetryMeasurement(telemetry, ["packetLoss", "signal_packet_loss", "packet_loss", "SignalPacketLoss"], "%"), color: getPacketLossMetricColor(packetLoss) },
      ],
    },
  ];

  return (
    <>
      <tr className="telemetryDetailRow">
        <td className="telemetryDetailLead" colSpan={3}>
          {recommendation ? <p>{recommendation}</p> : null}
          <SignalScoreTrendChart
            error={signalScoreTrendError}
            isLoading={signalScoreTrendLoading}
            trend={signalScoreTrend}
          />
        </td>
        {telemetryGroups.map((group) => (
          <td className="telemetryDetailCell" key={group.title}>
            <div className="telemetryGroupMetrics" aria-label={`${group.title} telemetry`}>
              {group.metrics.map((item) => (
                <div className="metricItem" key={item.label}>
                  <span className="metricLabel">
                    <img src={item.icon} alt="" aria-hidden="true" />
                    <span>{item.label}</span>
                  </span>
                  <strong style={{ color: item.color }} title={getTelemetryValueTitle(item.value)}>
                    {truncateTelemetryValue(item.value)}
                  </strong>
                </div>
              ))}
            </div>
          </td>
        ))}
        <td className="telemetryDetailActionCell"></td>
      </tr>
      <tr className="telemetryInsightsRow">
        <td colSpan={6}>
          <section className="telemetryInsights" aria-label="Participant telemetry insights">
            <div className="telemetryInsightsHeader">
              <img className="signal-reccommendation-icon-svg telemetryInsightsIcon" src={signalRecommendationIcon} alt="" aria-hidden="true" />
              <h3>Insights</h3>
            </div>
            {hasActiveAnalysisSession ? (
              <>
                <div className="telemetryInsightOverall">
                  <span>Overall recommendations</span>
                  {telemetryInsights.recommendations.length > 0 ? telemetryInsights.recommendations.map((recommendation) => (
                    <p key={recommendation.text}>- {recommendation.text}</p>
                  )) : <p>{telemetryInsights.overall}</p>}
                </div>
                <div className="telemetryInsightCategories">
                  {telemetryInsights.categories.map((insight) => (
                    <article className="telemetryInsightItem" key={insight.title}>
                      <span style={{ color: getTelemetryInsightColor(insight.severity) }}>{insight.title}</span>
                      <p style={{ color: getTelemetryInsightColor(insight.severity) }}>{insight.text}</p>
                    </article>
                  ))}
                </div>
              </>
            ) : (
              <p className="telemetryInsightsEmpty">Run analysis to gather insights.</p>
            )}
          </section>
        </td>
        <td className="telemetryDetailActionCell"></td>
      </tr>
    </>
  );
}

export default function App() {
  const isConfigPage = window.location.pathname.toLowerCase().startsWith("/tabs/config");
  const isPrivacyPolicyPage =
    window.location.pathname.toLowerCase().startsWith(PRIVACY_POLICY_PATH) ||
    window.location.pathname.toLowerCase().startsWith("/privacy");
  const isTermsOfServicePage =
    window.location.pathname.toLowerCase().startsWith(TERMS_OF_SERVICE_PATH) ||
    window.location.pathname.toLowerCase().startsWith("/terms");
  const apiBaseUrl = normalizeBaseUrl(import.meta.env.VITE_SIGNALTUNER_API_URL);
  const [authPageMode, setAuthPageMode] = React.useState<AuthPageMode>(() => getAuthPageMode());
  const [returnUrl, setReturnUrl] = React.useState(() => getReturnUrl());
  const [sessionToken, setSessionToken] = React.useState<string | null>(() =>
    window.localStorage.getItem(SIGNALTUNER_SESSION_TOKEN_KEY)
  );
  const [dashboard, setDashboard] = React.useState<DashboardData | null>(null);
  const [meetingContext, setMeetingContext] = React.useState<TeamsMeetingContext | null>(null);
  const [analysis, setAnalysis] = React.useState<AnalysisResult | null>(null);
  const [subscriptionPrompt, setSubscriptionPrompt] = React.useState<SubscriptionPrompt | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [isLoading, setIsLoading] = React.useState(false);
  const [isAddingTestingCredit, setIsAddingTestingCredit] = React.useState(false);
  const [busyState, setBusyState] = React.useState<AuthBusyState>("idle");
  const [isRunningInTeams, setIsRunningInTeams] = React.useState(false);
  const [hasCheckedTeamsContext, setHasCheckedTeamsContext] = React.useState(false);
  const [teamsTheme, setTeamsTheme] = React.useState<TeamsTheme>("default");
  const [themePreference, setThemePreference] = React.useState<SignalTunerThemePreference>(() => getStoredThemePreference());
  const [isClientPromptDismissed, setIsClientPromptDismissed] = React.useState(false);
  const [activationCodeError, setActivationCodeError] = React.useState<string | null>(null);
  const [accountUser, setAccountUser] = React.useState<CurrentUser | null>(null);
  const [pendingProfileAuth, setPendingProfileAuth] = React.useState<PendingProfileAuth | null>(null);
  const [pendingTeamsSsoAccountCreation, setPendingTeamsSsoAccountCreation] = React.useState<PendingTeamsSsoAccountCreation | null>(null);
  const [activePage, setActivePage] = React.useState<InAppPage>("dashboard");
  const isMountedRef = React.useRef(true);

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

  React.useEffect(() => {
    applyTeamsTheme(getEffectiveTheme(teamsTheme, themePreference));
    window.localStorage.setItem(SIGNALTUNER_THEME_PREFERENCE_KEY, themePreference);
  }, [teamsTheme, themePreference]);

  React.useEffect(() => {
    isMountedRef.current = true;

    const updateRoute = () => {
      setAuthPageMode(getAuthPageMode());
      setReturnUrl(getReturnUrl());
    };

    window.addEventListener("popstate", updateRoute);

    return () => {
      isMountedRef.current = false;
      window.removeEventListener("popstate", updateRoute);
    };
  }, []);

  const mergeCredits = React.useCallback((credits: number) => {
    setAccountUser((current) =>
      current
        ? {
            ...current,
            credits,
          }
        : current
    );
    setDashboard((current) =>
      current
        ? {
            ...current,
            currentUser: {
              ...current.currentUser,
              credits,
            },
          }
        : current
    );
  }, []);

  const refreshAccountInfo = React.useCallback(
    async (token: string, fallbackUser: CurrentUser | null = null) => {
      const account = mergeCurrentUser(
        normalizeCurrentUser(
          await fetchJson<unknown>(`${apiBaseUrl}/api/auth/me`, {
            headers: buildAuthHeaders(token),
          })
        ),
        fallbackUser
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
      const signalTunerSessionToken = getSignalTunerSessionToken(response);
      const responseUser = normalizeAuthCurrentUser(response);

      if (isProfileRequired(response, responseUser)) {
        setPendingProfileAuth({
          token: signalTunerSessionToken,
          user: responseUser,
        });
        setAccountUser(responseUser);
        return;
      }

      window.localStorage.setItem(SIGNALTUNER_SESSION_TOKEN_KEY, signalTunerSessionToken);
      window.localStorage.removeItem(SIGNALTUNER_EXPLICIT_SIGN_OUT_KEY);
      setPendingProfileAuth(null);
      setSessionToken(signalTunerSessionToken);
      setIsClientPromptDismissed(false);
      setActivationCodeError(null);
      const account = await refreshAccountInfo(signalTunerSessionToken, responseUser);

      if (meetingContext) {
        await joinMeetingSession(signalTunerSessionToken, meetingContext, account);
        restoreReturnUrl(returnUrl);
        return;
      }

      setError("You are signed in. Open SignalTuner from inside a Microsoft Teams meeting to join a meeting health view.");
    },
    [joinMeetingSession, meetingContext, refreshAccountInfo, returnUrl]
  );

  React.useEffect(() => {
    if (isConfigPage || isPrivacyPolicyPage || isTermsOfServicePage) {
      return;
    }

    let isCancelled = false;
    let disposeThemeSubscription: (() => void) | null = null;

    const initializeTeamsContext = async () => {
      setHasCheckedTeamsContext(false);
      setIsLoading(true);
      setBusyState("initializing-teams");

      try {
        const context = await initializeTeams();
        if (isCancelled) {
          return;
        }

        setIsRunningInTeams(true);
        const nextTheme = mapTeamsTheme(context.app.theme);
        setTeamsTheme(nextTheme);
        disposeThemeSubscription = subscribeToTeamsThemeChanges(setTeamsTheme);
        const teamsMeetingId = context.meeting?.id;

        if (!teamsMeetingId) {
          setMeetingContext(null);
          setError(null);
          return;
        }

        const meetingDetails = await getTeamsMeetingDetails();
        const organizer = asRecord(meetingDetails?.organizer);
        const organizerM365ObjectId = normalizeTeamsObjectId(readString(organizer, "id"));
        const organizerTenantId = readString(organizer, "tenantId");
        const currentUserM365ObjectId = normalizeTeamsObjectId(context.user?.id);
        const currentUserTenantId = context.user?.tenant?.id ?? null;
        const currentUserMeetingRole = sameTeamsId(currentUserM365ObjectId, organizerM365ObjectId) ? "Organizer" : null;

        if (import.meta.env.DEV) {
          console.info("SignalTuner meeting role context", {
            hasMeetingDetails: Boolean(meetingDetails),
            hasCurrentUserId: Boolean(currentUserM365ObjectId),
            hasOrganizerId: Boolean(organizerM365ObjectId),
            currentUserIsOrganizer: currentUserMeetingRole === "Organizer",
          });
        }

        setMeetingContext({
          teamsMeetingId,
          teamsConversationId: context.chat?.id ?? context.channel?.id ?? null,
          teamsTenantId: context.user?.tenant?.id ?? null,
          meetingTitle: readString(asRecord(context.meeting), "title", "subject", "displayName") ?? readString(asRecord(meetingDetails?.details), "title"),
          organizerM365ObjectId,
          organizerTenantId,
          currentUserM365ObjectId,
          currentUserTenantId,
          currentUserMeetingRole,
        });
        setError(null);
      } catch {
        if (isCancelled) {
          return;
        }

        setIsRunningInTeams(false);
        setMeetingContext(null);
        setTeamsTheme("default");
        setError(null);
      } finally {
        if (!isCancelled) {
          setHasCheckedTeamsContext(true);
          setIsLoading(false);
          setBusyState("idle");
        }
      }
    };

    void initializeTeamsContext();

    return () => {
      isCancelled = true;
      disposeThemeSubscription?.();
    };
  }, [isConfigPage, isPrivacyPolicyPage, isTermsOfServicePage]);

  React.useEffect(() => {
    if (
      isConfigPage ||
      isPrivacyPolicyPage ||
      isTermsOfServicePage ||
      authPageMode === "create-account" ||
      authPageMode === "support" ||
      !apiBaseUrl ||
      !isRunningInTeams ||
      !meetingContext ||
      sessionToken ||
      dashboard ||
      pendingProfileAuth ||
      pendingTeamsSsoAccountCreation ||
      window.localStorage.getItem(SIGNALTUNER_EXPLICIT_SIGN_OUT_KEY) === "true" ||
      window.sessionStorage.getItem(SIGNALTUNER_AUTO_SSO_FAILED_KEY) === "true"
    ) {
      return;
    }

    let isCancelled = false;

    const attemptAutomaticTeamsSso = async () => {
      setBusyState("auto-sso");
      setIsLoading(true);
      setError(null);

      try {
        const response = await authenticateWithTeamsSso(apiBaseUrl, meetingContext);
        if (!isCancelled && isMountedRef.current) {
          await completeAuth(response);
        }
      } catch (caught) {
        window.sessionStorage.setItem(SIGNALTUNER_AUTO_SSO_FAILED_KEY, "true");
        if (!isCancelled && isMountedRef.current) {
          setError(getFriendlyErrorMessage(caught, "Teams sign-in could not be completed. Try again or sign in with email."));
        }
      } finally {
        if (!isCancelled && isMountedRef.current) {
          setIsLoading(false);
          setBusyState("idle");
        }
      }
    };

    void attemptAutomaticTeamsSso();

    return () => {
      isCancelled = true;
    };
  }, [
    apiBaseUrl,
    authPageMode,
    completeAuth,
    dashboard,
    isConfigPage,
    isPrivacyPolicyPage,
    isRunningInTeams,
    isTermsOfServicePage,
    meetingContext,
    pendingProfileAuth,
    pendingTeamsSsoAccountCreation,
    sessionToken,
  ]);

  React.useEffect(() => {
    if (!sessionToken || !meetingContext || dashboard) {
      return;
    }

    setIsLoading(true);
    setBusyState("checking-session");
    refreshAccountInfo(sessionToken)
      .then((account) => joinMeetingSession(sessionToken, meetingContext, account))
      .catch((caught) => {
        window.localStorage.removeItem(SIGNALTUNER_SESSION_TOKEN_KEY);
        setSessionToken(null);
        setError(sanitizeAuthError(caught, "Your session expired. Sign in again to continue."));
      })
      .finally(() => {
        setIsLoading(false);
        setBusyState("idle");
      });
  }, [dashboard, joinMeetingSession, meetingContext, refreshAccountInfo, sessionToken]);

  React.useEffect(() => {
    if (!dashboard || !sessionToken) {
      return;
    }

    const currentUser = mergeCurrentUser(dashboard.currentUser, accountUser);

    if (!currentUser.clientIsActive && !isClientPromptDismissed) {
      return;
    }

    const intervalId = window.setInterval(() => {
      void refreshDashboard(dashboard.meetingSessionId, sessionToken);
    }, 10000);

    return () => window.clearInterval(intervalId);
  }, [accountUser, dashboard, isClientPromptDismissed, refreshDashboard, sessionToken]);

  const signInWithTeams = React.useCallback(async () => {
    if (!apiBaseUrl) {
      setError("SignalTuner is temporarily unavailable. Please try again.");
      return;
    }

    setIsLoading(true);
    setBusyState("teams-sso");
    setError(null);

    try {
      const response = await authenticateWithTeamsSso(apiBaseUrl, meetingContext);
      await completeAuth(response);
      window.sessionStorage.removeItem(SIGNALTUNER_AUTO_SSO_FAILED_KEY);
    } catch (caught) {
      setError(getFriendlyErrorMessage(caught, "Teams sign-in could not be completed. Try again or sign in with email."));
    } finally {
      setIsLoading(false);
      setBusyState("idle");
    }
  }, [apiBaseUrl, completeAuth, meetingContext]);

  const beginTeamsSsoAccountCreation = React.useCallback(() => {
    if (!isRunningInTeams) {
      setError("Teams account creation is available when this app is opened in Microsoft Teams.");
      return;
    }

    setError(null);
    setPendingTeamsSsoAccountCreation({
      user: {
        userId: 0,
        email: null,
        displayName: null,
        credits: 0,
        clientIsActive: false,
        authProvider: "teams_sso",
      },
    });
  }, [isRunningInTeams]);

  const createAccountWithTeams = React.useCallback(
    async (firstName: string, lastName: string) => {
      if (!apiBaseUrl || !pendingTeamsSsoAccountCreation) {
        setError("SignalTuner is temporarily unavailable. Please try again.");
        return;
      }

      setIsLoading(true);
      setBusyState("teams-sso");
      setError(null);

      try {
        const response = await authenticateWithTeamsSso(apiBaseUrl, meetingContext, {
          allowAccountCreation: true,
          firstName,
          lastName,
          termsOfServiceAccepted: true,
          privacyPolicyAccepted: true,
          termsOfServiceVersion: LEGAL_DOCUMENT_EFFECTIVE_DATE,
          privacyPolicyVersion: LEGAL_DOCUMENT_EFFECTIVE_DATE,
        });
        setPendingTeamsSsoAccountCreation(null);
        await completeAuth(response);
        window.sessionStorage.removeItem(SIGNALTUNER_AUTO_SSO_FAILED_KEY);
      } catch (caught) {
        setError(getFriendlyErrorMessage(caught, "We could not create your account. Review your profile and consent, then try again."));
      } finally {
        setIsLoading(false);
        setBusyState("idle");
      }
    },
    [apiBaseUrl, completeAuth, meetingContext, pendingTeamsSsoAccountCreation]
  );

  const emailSignIn = React.useCallback(
    async (email: string, password: string) => {
      if (!apiBaseUrl) {
        setError("SignalTuner is temporarily unavailable. Please try again.");
        return;
      }

      setIsLoading(true);
      setBusyState("email-login");
      setError(null);

      try {
        const response = await signInWithEmail(apiBaseUrl, email, password);
        response.email = response.email ?? response.Email ?? email;
        await completeAuth(response);
      } catch (caught) {
        setError(sanitizeAuthError(caught, "Your email or password was not recognized."));
      } finally {
        setIsLoading(false);
        setBusyState("idle");
      }
    },
    [apiBaseUrl, completeAuth]
  );

  const emailRegister = React.useCallback(
    async (email: string, password: string, firstName: string, lastName: string) => {
      if (!apiBaseUrl) {
        setError("SignalTuner is temporarily unavailable. Please try again.");
        return;
      }

      setIsLoading(true);
      setBusyState("email-register");
      setError(null);

      try {
        const response = await createAccountWithEmail(apiBaseUrl, { email, password, firstName, lastName });
        await completeAuth(response);
      } catch (caught) {
        setError(sanitizeAuthError(caught, "We could not create your account. Review the information and try again."));
      } finally {
        setIsLoading(false);
        setBusyState("idle");
      }
    },
    [apiBaseUrl, completeAuth]
  );

  const requestPasswordReset = React.useCallback(
    async (email: string): Promise<boolean> => {
      if (!apiBaseUrl) {
        setError("SignalTuner is temporarily unavailable. Please try again.");
        return false;
      }

      setIsLoading(true);
      setBusyState("password-reset");
      setError(null);

      try {
        await requestPasswordResetEmail(apiBaseUrl, email);
        return true;
      } catch (caught) {
        setError(caught instanceof Error ? caught.message : "Failed to send password reset email.");
        return false;
      } finally {
        setIsLoading(false);
        setBusyState("idle");
      }
    },
    [apiBaseUrl]
  );

  const submitProfile = React.useCallback(
    async (firstName: string, lastName: string) => {
      if (!apiBaseUrl || !pendingProfileAuth) {
        setError("SignalTuner is temporarily unavailable. Please try again.");
        return;
      }

      setIsLoading(true);
      setBusyState("teams-sso");
      setError(null);

      try {
        const response = await completeUserProfile(apiBaseUrl, pendingProfileAuth.token, { firstName, lastName });
        await completeAuth(response);
        window.sessionStorage.removeItem(SIGNALTUNER_AUTO_SSO_FAILED_KEY);
      } catch (caught) {
        setError(sanitizeAuthError(caught, "We could not save your profile. Review your name and try again."));
      } finally {
        setIsLoading(false);
        setBusyState("idle");
      }
    },
    [apiBaseUrl, completeAuth, pendingProfileAuth]
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
        mergeCredits(data.remainingCredits);
        await refreshDashboard(dashboard.meetingSessionId, sessionToken);
      } catch (caught) {
        if (!parseCreditError(caught)) {
          setError(caught instanceof Error ? caught.message : String(caught));
        }
      } finally {
        setIsLoading(false);
      }
    },
    [apiBaseUrl, dashboard, mergeCredits, parseCreditError, refreshDashboard, sessionToken]
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
      mergeCredits(data.remainingCredits);
      await refreshDashboard(dashboard.meetingSessionId, sessionToken);
    } catch (caught) {
      if (!parseCreditError(caught)) {
        setError(caught instanceof Error ? caught.message : String(caught));
      }
    } finally {
      setIsLoading(false);
    }
  }, [apiBaseUrl, dashboard, mergeCredits, parseCreditError, refreshDashboard, sessionToken]);

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

  const refreshClientPromptStatus = React.useCallback(async () => {
    if (!dashboard || !sessionToken) {
      return;
    }

    setActivationCodeError(null);

    try {
      await refreshDashboard(dashboard.meetingSessionId, sessionToken);
    } catch (caught) {
      setActivationCodeError(caught instanceof Error ? caught.message : "Unable to refresh desktop telemetry status.");
    }
  }, [dashboard, refreshDashboard, sessionToken]);

  const addTestingCredit = React.useCallback(async () => {
    if (!sessionToken) {
      return;
    }

    setIsAddingTestingCredit(true);
    setError(null);

    try {
      const response = asRecord(
        await fetchJson<unknown>(`${apiBaseUrl}/api/User/credits/testing/add-one`, {
          method: "POST",
          headers: buildAuthHeaders(sessionToken),
        })
      );
      const credits = readNumber(response, mergeCurrentUser(dashboard?.currentUser ?? normalizeCurrentUser(null), accountUser).credits, "credits", "Credits", "userCredits", "UserCredits");
      mergeCredits(credits);
      await refreshAccountInfo(sessionToken);
      if (dashboard) {
        await refreshDashboard(dashboard.meetingSessionId, sessionToken);
      }
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
    } finally {
      setIsAddingTestingCredit(false);
    }
  }, [accountUser, apiBaseUrl, dashboard, mergeCredits, refreshAccountInfo, refreshDashboard, sessionToken]);

  const applyAccountUpdateResponse = React.useCallback((response: AuthResponse) => {
    const nextToken = getSignalTunerSessionToken(response);
    const nextUser = normalizeAuthCurrentUser(response);

    window.localStorage.setItem(SIGNALTUNER_SESSION_TOKEN_KEY, nextToken);
    setSessionToken(nextToken);
    setAccountUser(nextUser);
    setDashboard((current) =>
      current
        ? {
            ...current,
            currentUser: mergeCurrentUser(nextUser, current.currentUser),
          }
        : current
    );

    return { nextToken, nextUser };
  }, []);

  const submitAccountProfileUpdate = React.useCallback(
    async (firstName: string, lastName: string): Promise<CurrentUser> => {
      if (!apiBaseUrl || !sessionToken) {
        throw new Error("SignalTuner is temporarily unavailable. Please try again.");
      }

      setError(null);
      const response = await completeUserProfile(apiBaseUrl, sessionToken, { firstName, lastName });
      const { nextToken, nextUser } = applyAccountUpdateResponse(response);

      if (dashboard) {
        await refreshDashboard(dashboard.meetingSessionId, nextToken);
      }

      return nextUser;
    },
    [apiBaseUrl, applyAccountUpdateResponse, dashboard, refreshDashboard, sessionToken]
  );

  const submitAccountEmailUpdate = React.useCallback(
    async (email: string): Promise<CurrentUser> => {
      if (!apiBaseUrl || !sessionToken) {
        throw new Error("SignalTuner is temporarily unavailable. Please try again.");
      }

      setError(null);
      const response = await updateAccountEmail(apiBaseUrl, sessionToken, email);
      const { nextToken, nextUser } = applyAccountUpdateResponse(response);

      if (dashboard) {
        await refreshDashboard(dashboard.meetingSessionId, nextToken);
      }

      return nextUser;
    },
    [apiBaseUrl, applyAccountUpdateResponse, dashboard, refreshDashboard, sessionToken]
  );

  const submitAccountPasswordUpdate = React.useCallback(
    async (newPassword: string): Promise<string> => {
      const currentUser = mergeCurrentUser(dashboard?.currentUser ?? normalizeCurrentUser(null), accountUser);

      if (!apiBaseUrl || !sessionToken || !currentUser.userId) {
        throw new Error("SignalTuner is temporarily unavailable. Please try again.");
      }

      setError(null);
      const message = await updateAccountPassword(apiBaseUrl, sessionToken, currentUser.userId, newPassword);
      await refreshAccountInfo(sessionToken);
      return message;
    },
    [accountUser, apiBaseUrl, dashboard, refreshAccountInfo, sessionToken]
  );

  const submitSupportRequest = React.useCallback(
    async (request: { sender: string; subject: string; body: string }): Promise<string> => {
      if (!apiBaseUrl) {
        throw new Error("SignalTuner support is temporarily unavailable. Please try again.");
      }

      return sendSupportEmail(apiBaseUrl, request);
    },
    [apiBaseUrl]
  );

  const signOut = React.useCallback(() => {
    window.localStorage.removeItem(SIGNALTUNER_SESSION_TOKEN_KEY);
    window.localStorage.setItem(SIGNALTUNER_EXPLICIT_SIGN_OUT_KEY, "true");
    window.sessionStorage.setItem(SIGNALTUNER_AUTO_SSO_FAILED_KEY, "true");
    setSessionToken(null);
    setDashboard(null);
    setAnalysis(null);
    setSubscriptionPrompt(null);
    setIsClientPromptDismissed(false);
    setActivationCodeError(null);
    setAccountUser(null);
    setPendingProfileAuth(null);
    setPendingTeamsSsoAccountCreation(null);
    setActivePage("dashboard");
  }, []);

  if (isConfigPage) {
    return <ConfigPage />;
  }

  if (isPrivacyPolicyPage) {
    return <LegalDocumentPage markdown={privacyPolicyMarkdown} title="SignalTuner Privacy Policy" />;
  }

  if (isTermsOfServicePage) {
    return <LegalDocumentPage markdown={termsOfServiceMarkdown} title="SignalTuner Terms of Service" />;
  }

  if (sessionToken && !dashboard) {
    if (!hasCheckedTeamsContext || meetingContext || busyState === "checking-session" || busyState === "initializing-teams" || isLoading) {
      return <SessionLoadingPage />;
    }

    return <SignedInOutsideMeetingPage onSignOut={signOut} />;
  }

  if (pendingTeamsSsoAccountCreation && !sessionToken && !dashboard) {
    return (
      <CompleteProfilePage
        busyState={busyState}
        error={error}
        requireLegalConsent
        secondaryActionLabel="Back"
        user={pendingTeamsSsoAccountCreation.user}
        onSubmit={createAccountWithTeams}
        onSupport={() => navigateToSupport(returnUrl)}
        onSignOut={() => setPendingTeamsSsoAccountCreation(null)}
      />
    );
  }

  if (pendingProfileAuth && !sessionToken && !dashboard) {
    return (
      <CompleteProfilePage
        busyState={busyState}
        error={error}
        user={pendingProfileAuth.user}
        onSubmit={submitProfile}
        onSupport={() => navigateToSupport(returnUrl)}
        onSignOut={signOut}
      />
    );
  }

  if (!sessionToken || !dashboard) {
    if (authPageMode === "create-account") {
      return (
        <CreateAccountPage
          busyState={busyState}
          error={error}
          isRunningInTeams={isRunningInTeams}
          onEmailRegister={emailRegister}
          onSignIn={() => navigateAuth("login", returnUrl)}
          onSupport={() => navigateToSupport(returnUrl)}
          onTeamsAccountCreate={beginTeamsSsoAccountCreation}
        />
      );
    }

    if (authPageMode === "forgot-password") {
      return (
        <ForgotPasswordPage
          busyState={busyState}
          error={error}
          onRequestPasswordReset={requestPasswordReset}
          onSignIn={() => navigateAuth("login", returnUrl)}
          onSupport={() => navigateToSupport(returnUrl)}
        />
      );
    }

    if (authPageMode === "support") {
      return (
        <SupportAuthPage
          defaultEmail={accountUser?.email}
          onSendSupportEmail={submitSupportRequest}
          onSignIn={() => navigateAuth("login", returnUrl)}
          onSupport={() => navigateToSupport(returnUrl)}
        />
      );
    }

    return (
      <LoginPage
        busyState={busyState}
        error={error}
        isRunningInTeams={isRunningInTeams}
        meetingContext={meetingContext}
        onCreateAccount={() => navigateAuth("create-account", returnUrl)}
        onEmailSignIn={emailSignIn}
        onForgotPassword={() => navigateAuth("forgot-password", returnUrl)}
        onSupport={() => navigateToSupport(returnUrl)}
        onTeamsSignIn={signInWithTeams}
      />
    );
  }

  const currentUser = mergeCurrentUser(dashboard.currentUser, accountUser);
  const displayDashboard = { ...dashboard, currentUser };
  const shouldShowClientPrompt = !currentUser.clientIsActive && !isClientPromptDismissed;

  return (
    <ActivationContext.Provider value={{ apiBaseUrl, token: sessionToken }}>
      {shouldShowClientPrompt && (
        <main className="pageShell">
          <ClientPrompt
            activationCodeError={activationCodeError}
            isLoading={isLoading}
            onContinue={() => {
              setActivePage("dashboard");
              setIsClientPromptDismissed(true);
            }}
            onRefresh={refreshClientPromptStatus}
            onSignOut={signOut}
          />
        </main>
      )}
      {!shouldShowClientPrompt && (
        <Dashboard
          analysis={analysis}
          apiBaseUrl={apiBaseUrl}
          activePage={activePage}
          dashboard={displayDashboard}
          error={error}
          isLoading={isLoading}
          isAddingTestingCredit={isAddingTestingCredit}
          onAddTestingCredit={addTestingCredit}
          onAnalyzeAll={analyzeAll}
          onAnalyzeUser={analyzeUser}
          onInvite={inviteParticipants}
          onNavigate={setActivePage}
          onDismissSubscriptionPrompt={() => setSubscriptionPrompt(null)}
          onSendSupportEmail={submitSupportRequest}
          onSignOut={signOut}
          onThemePreferenceChange={setThemePreference}
          onUpdateEmail={submitAccountEmailUpdate}
          onUpdatePassword={submitAccountPasswordUpdate}
          onUpdateProfile={submitAccountProfileUpdate}
          sessionToken={sessionToken}
          subscriptionPrompt={subscriptionPrompt}
          themePreference={themePreference}
        />
      )}
    </ActivationContext.Provider>
  );
}
