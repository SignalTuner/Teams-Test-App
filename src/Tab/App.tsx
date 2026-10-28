import ConnectivitySummary from "./ConnectivitySummary";
import ClientPrompt from "./ClientPrompt";
import PersonalDashboard from "./PersonalDashboard";
import DashboardHeader, { SignalTunerLogo, getInitials } from "./DashboardHeader";
import { getSignalDisplayStatusFromScore, CompactSignalScoreDisplay, ParticipantTelemetryDetail } from "./ConnectivityPresentation";
import type { CurrentUser, Issue, TelemetryRecord, SignalScoreTrendResponse } from "./dashboardTypes";
import { ActivationContext } from "./TemporaryActivationCode";
import React from "react";
import * as teamsJs from "@microsoft/teams-js";
import { initializeTeams } from "./teamsHost";
import { PaymentDisclosure, useExternalCommerce } from "./TeamsHostProvider";

import "./App.css";
import "./ConfigPage.css";
import SettingsPage from "./SettingsPage";
import AppearanceSelector from "./AppearanceSelector";
import SupportPage from "./SupportPage";
import { saveAccountProfile } from "./settingsAccount";
import teamsAnalysisPreview from "./assets/teams-analysis-preview.png";
import MeetingServiceStatus from "./MeetingServiceStatus";
import { openMeetingInvitations } from "./meetingInvitations";
import microsoftTeamsLogo from "./assets/microsoft-teams.png";
import RemoteAnalysisButton, { type RequestState } from "./RemoteAnalysisButton";
import MeetingAnalysisButton from "./MeetingAnalysisButton";
import privacyPolicyMarkdown from "./content/privacy-policy.md?raw";
import termsOfServiceMarkdown from "./content/terms-of-service.md?raw";

type AuthProvider = "teams_sso" | "google" | "github" | "email_magic_code";
type ClientDataStatus = "active" | "inactive" | "no_data";
type AnalysisStatus = "Healthy" | "Fair" | "Poor" | "Critical" | "Offline" | "No data" | string;

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

type AuthPageMode = "login" | "create-account" | "forgot-password" | "support";
type InAppPage = "dashboard" | "settings" | "support";
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
// Temporarily paused until public-tenant SSO is ready. Keep backend contracts intact.
const TEAMS_SSO_ENABLED = false;
const SIGNALTUNER_EXPLICIT_SIGN_OUT_KEY = "signaltunerExplicitSignOut";
const SIGNALTUNER_THEME_PREFERENCE_KEY = "signaltunerThemePreference";
const PRIVACY_POLICY_PATH = "/tabs/privacy";
const TERMS_OF_SERVICE_PATH = "/tabs/terms";
const LEGAL_DOCUMENT_EFFECTIVE_DATE = "2026-08-12";
const PASSWORD_REQUIREMENT_TEXT = "Use at least 8 characters, including uppercase, lowercase, number, and symbol.";


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
  if (!TEAMS_SSO_ENABLED) throw new Error("Please sign in with email and password.");
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

function normalizeBaseUrl(baseUrl: string | undefined): string {
  return baseUrl ? baseUrl.replace(/\/+$/, "") : "";
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
    let cancelled = false;
    let disposeThemeSubscription: (() => void) | undefined;
    const updateTheme = (theme: TeamsTheme) => {
      if (cancelled) return;
      const preference = window.localStorage.getItem(SIGNALTUNER_THEME_PREFERENCE_KEY);
      applyTeamsTheme(preference === "light" || preference === "dark" ? getEffectiveTheme(theme, preference) : theme);
    };
    updateTheme("dark");
    initializeTeams().then(context => {
      if (cancelled) return;
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
      disposeThemeSubscription = subscribeToTeamsThemeChanges(updateTheme);
      updateTheme(mapTeamsTheme(context.app.theme));
    }).catch(() => undefined);
    return () => { cancelled = true; disposeThemeSubscription?.(); };
  }, []);

  return (
    <main className="configPage">
      <section className="configOverview" aria-labelledby="config-heading">
        <header>
          <h1 id="config-heading">Add SignalTuner to your meeting</h1>
        </header>
        <ul className="configFeatures">
          <li>
            <span className="configFeatureIcon" aria-hidden="true"><span className="configSignalIcon" /></span>
            <div><h2>Monitor connectivity health</h2><p>View Signal Scores and Device, Workspace, and Network health for connected participants.</p></div>
          </li>
          <li>
            <span className="configFeatureIcon" aria-hidden="true"><span className="configAnalysisIcon" /></span>
            <div><h2>Analyze connection issues</h2><p>Run detailed connectivity analyses to help identify and resolve disruptions.</p></div>
          </li>
          <li>
            <span className="configFeatureIcon" aria-hidden="true"><span className="configServicesIcon" /></span>
            <div><h2>Monitor Microsoft Teams status</h2><p>Stay informed about service incidents and outages that may affect your meeting.</p></div>
          </li>
        </ul>
        <aside className="configInfoNote">
          <span className="infoGlyph" aria-hidden="true">i</span>
          <p>Participants must sign in and connect the SignalTuner desktop app to share their connectivity health.</p>
        </aside>
        <footer className="configFooter"><a href={PRIVACY_POLICY_PATH}>Privacy Policy</a></footer>
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

function BrandPanel() {
  return (
    <aside className="brandPanel">
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
      <img className="meetingPreview" src={teamsAnalysisPreview} alt="" aria-hidden="true" />
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
        <SignalTunerLogo className="brandLogo" />
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
      {TEAMS_SSO_ENABLED && busyState === "auto-sso" && (
        <p className="authStatus" role="status" aria-live="polite">
          <Spinner /> Signing you in with Microsoft Teams...
        </p>
      )}
      {TEAMS_SSO_ENABLED && !isRunningInTeams && (
        <p className="authNotice">Teams SSO is available when this app is opened in Microsoft Teams.</p>
      )}
      {isRunningInTeams && !meetingContext && (
        <p className="authNotice">Sign in to view your connectivity, analyses and previous Teams meetings.</p>
      )}
      {TEAMS_SSO_ENABLED && (
        <>
          <TeamsAuthButton
            disabled={!isRunningInTeams || isBusy}
            isBusy={teamsBusy}
            label="Continue with Microsoft Teams"
            onClick={() => void onTeamsSignIn()}
          />
          <div className="authSeparator">
            <span>or</span>
          </div>
        </>
      )}
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
      {TEAMS_SSO_ENABLED && !isRunningInTeams && (
        <p className="authNotice">Teams account creation is available when this app is opened in Microsoft Teams.</p>
      )}
      {TEAMS_SSO_ENABLED && (
        <>
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
        </>
      )}
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
        <PaymentDisclosure />
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
      {user.email && <p className="authNotice">Signed in as {user.email}.</p>}
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
            <PaymentDisclosure />
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

function CreditRequiredModal({
  prompt,
  onClose,
  onOpenSettings,
}: {
  prompt: SubscriptionPrompt;
  onClose: () => void;
  onOpenSettings: () => void;
}) {
  const allowCommerce = useExternalCommerce();
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
            <h2 id="credit-required-title">Report allowance unavailable</h2>
          </div>
          <button aria-label="Dismiss credit notification" className="modalCloseButton" onClick={onClose} type="button">
            X
          </button>
        </header>
        <div className="incidentModalBody">
          <p className="creditRequiredMessage">
            This analysis needs an allowance of {prompt.requiredCredits} reports. You have {prompt.availableCredits} available.
          </p>
          <p>Test Analysis remains available on supported connected desktops without using report allowance.</p>
          <div className="creditRequiredActions">
            {allowCommerce && <button className="primaryButton" onClick={onOpenSettings} type="button">
              Open Settings
            </button>}
            <button className="secondaryButton" onClick={onClose} type="button">
              Dismiss
            </button>
          </div>
        </div>
      </section>
    </div>
  );
}

// Match the website SignalScoreGauge's center, radius and 250-degree sweep.
function Dashboard({
  analysis,
  apiBaseUrl,
  dashboard,
  activePage,
  error,
  isLoading,
  onInvite,
  onNavigate,
  onDismissSubscriptionPrompt,
  onSendSupportEmail,
  onSignOut,
  onThemePreferenceChange,
  onUpdateProfile,
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
  onInvite: () => Promise<void>;
  onNavigate: (page: InAppPage) => void;
  onDismissSubscriptionPrompt: () => void;
  onSendSupportEmail: (request: { sender: string; subject: string; body: string }) => Promise<string>;
  onSignOut: () => void;
  onThemePreferenceChange: (preference: SignalTunerThemePreference) => void;
  onUpdateProfile: (firstName: string, lastName: string, email: string) => Promise<void>;
  onUpdatePassword: (newPassword: string) => Promise<string>;
  sessionToken: string;
  subscriptionPrompt: SubscriptionPrompt | null;
  themePreference: SignalTunerThemePreference;
}) {
  const [expandedParticipantIds, setExpandedParticipantIds] = React.useState<Set<number>>(() => new Set());
  const [signalScoreTrends, setSignalScoreTrends] = React.useState<Record<number, SignalScoreTrendResponse>>({});
  const [signalScoreTrendLoadingIds, setSignalScoreTrendLoadingIds] = React.useState<Set<number>>(() => new Set());
  const [signalScoreTrendErrors, setSignalScoreTrendErrors] = React.useState<Record<number, string>>({});
  const [nowMs, setNowMs] = React.useState(Date.now());
  const [meetingAnalysisRequests, setMeetingAnalysisRequests] = React.useState<Record<string, RequestState>>({});
  const user = dashboard.currentUser;
  const participants = React.useMemo(() => sortParticipantsByMeetingRole(dashboard.participants), [dashboard.participants]);
  const activeParticipants = participants.filter((participant) => participant.clientDataStatus === "active");
  const participantsWithActiveScores = activeParticipants.filter(
    (participant) => participant.signalScore !== null && Number.isFinite(participant.signalScore)
  );
  const connectedParticipantCount = activeParticipants.length;
  const totalParticipantCount = participants.length;
  const connectivityIssueCount = participantsWithActiveScores.filter((participant) =>
    ["Poor", "Critical"].includes(getSignalDisplayStatusFromScore(participant.signalScore))
  ).length;
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

  React.useEffect(() => {
    if (!participants.some((participant) => getAnalysisRemainingMs(participant, Date.now()) > 0)) {
      return;
    }

    const intervalId = window.setInterval(() => setNowMs(Date.now()), 1000);
    return () => window.clearInterval(intervalId);
  }, [participants]);


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
      <DashboardHeader user={user} activePage={activePage} onNavigate={onNavigate} onSignOut={onSignOut} />

      {error && <section className="panel panelAlert">{error}</section>}

      {activePage === "settings" && (
        <SettingsPage user={user} apiBaseUrl={apiBaseUrl} token={sessionToken} onOpenSupport={() => onNavigate("support")}
          themePreference={themePreference} onThemePreferenceChange={onThemePreferenceChange}
          onUpdateProfile={onUpdateProfile} onUpdatePassword={onUpdatePassword} validatePassword={validatePassword} />
      )}

      {activePage === "support" && (
        <SupportPage defaultName={user.displayName} defaultEmail={user.email} onSendSupportEmail={onSendSupportEmail}
          themePreference={themePreference} onThemePreferenceChange={onThemePreferenceChange} />
      )}

      {activePage === "dashboard" && (
        <>
      <div className="dashboardAppearanceRow">
        <AppearanceSelector value={themePreference} onChange={onThemePreferenceChange} />
      </div>
      <section className="panel meetingHealthPanel" aria-labelledby="meeting-overview-heading">
        <h2 id="meeting-overview-heading" className="meetingOverviewTitle">Meeting Overview</h2>
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
            label={<>Participants experiencing<br />connectivity issues</>}
            tone={connectedParticipantCount === 0 ? "neutral" : connectivityIssueCount === 0 ? "good" : "warning"}
            value={String(connectivityIssueCount)}
          />
        </div>
        <div className="healthActions">
          <MeetingAnalysisButton key={`${user.userId}:${dashboard.meetingSessionId}`}
            apiBaseUrl={apiBaseUrl} token={sessionToken} userId={user.userId}
            meetingSessionId={dashboard.meetingSessionId} participants={participants} disabled={isLoading}
            onAccepted={results => setMeetingAnalysisRequests(current => ({ ...current,
              ...Object.fromEntries(results.filter(result => result.request).map(result => [
                `${user.userId}:${dashboard.meetingSessionId}:${result.targetUserId}`, result.request!,
              ])),
            }))} />
        </div>
      </section>

      <section className="panel participantsPanel" aria-labelledby="meeting-participants-heading">
        <div className="participantsToolbar">
          <div>
            <h2 id="meeting-participants-heading">Meeting Participants</h2>
            <p>{participants.length} {participants.length === 1 ? "participant" : "participants"}</p>
          </div>
          <div className="buttonRow inlineButtons">
            <button className="secondaryButton participantInviteButton" disabled={isLoading} onClick={onInvite} type="button">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
                <rect x="3" y="5" width="18" height="14" rx="2" />
                <path d="m3 7 9 6 9-6" />
              </svg>
              Invite
            </button>
          </div>
        </div>

        {participants.length === 0 && <p className="participantsEmpty">No meeting participants are available yet.</p>}
        <ul className="participantList" aria-label="Meeting participant list">
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
              <RemoteAnalysisButton key={`${user.userId}:${dashboard.meetingSessionId}:${participant.userId}`} apiBaseUrl={apiBaseUrl} token={sessionToken}
                targetUserId={participant.userId} targetName={getParticipantName(participant)}
                backgroundRequest={meetingAnalysisRequests[`${user.userId}:${dashboard.meetingSessionId}:${participant.userId}`]}
                meetingSessionId={dashboard.meetingSessionId}
                expanded={isExpanded} onToggle={() => toggleParticipantTelemetry(participant.userId)}
                onExpand={() => setExpandedParticipantIds(current => new Set(current).add(participant.userId))}
                details={isExpanded && hasActiveAnalysisSession ? <ParticipantTelemetryDetail
                  hasActiveAnalysisSession={hasActiveAnalysisSession}
                  hasData={hasData || hasActiveAnalysisSession}
                  issues={issues}
                  signalScoreTrend={analysisCoversParticipant ? signalScoreTrends[participant.userId] ?? null : null}
                  signalScoreTrendError={analysisCoversParticipant ? signalScoreTrendErrors[participant.userId] ?? null : null}
                  signalScoreTrendLoading={analysisCoversParticipant && signalScoreTrendLoadingIds.has(participant.userId)}
                  telemetry={telemetry}
                /> : null}
                renderInline={({ action, panel }) => <li className="participantListItem">
                <div className="participantRow">
                  <div className="participantIdentity">
                    <span className="participantAvatar" aria-hidden="true">
                      {getInitials(participant.displayName, participant.email)}
                    </span>
                    <div>
                      <h3 className="participantName" title={getParticipantName(participant)}>
                        {getParticipantName(participant)}
                      </h3>
                      {participant.email && participant.email !== getParticipantName(participant) && (
                        <p className="participantEmail" title={participant.email}>{participant.email}</p>
                      )}
                      <span className="participantRole">{formatParticipantMeetingRole(participant.meetingRole)}</span>
                    </div>
                  </div>
                  <ConnectivitySummary name={getParticipantName(participant)} score={participant.signalScore}
                    overallStatus={participant.overallStatus} hasData={hasData} meetingFallback
                    workspaceStatus={participant.workspaceStatus} networkStatus={participant.networkStatus} deviceStatus={participant.deviceStatus} />
                  <div className="participantActions" role="group" aria-label={`Actions for ${getParticipantName(participant)}`}>
                    {hasActiveAnalysisSession ? (
                      <span className="analysisSessionStatus">
                        <span>Analysis in progress</span>
                        <time>{formatAnalysisCountdown(analysisRemainingMs)}</time>
                      </span>
                    ) : !hasData ? (
                      <button className="secondaryButton compactAction" disabled={isLoading} onClick={onInvite} type="button">
                        Prompt
                      </button>
                    ) : null}
                    {action}
                  </div>
                </div>
                {panel}
              </li>} />
            );
          })}
        </ul>
      </section>

      <MeetingServiceStatus apiBaseUrl={apiBaseUrl} token={sessionToken} />
        </>
      )}

      {activePage === "dashboard" && subscriptionPrompt && (
        <CreditRequiredModal
          prompt={subscriptionPrompt}
          onClose={onDismissSubscriptionPrompt}
          onOpenSettings={() => {
            onDismissSubscriptionPrompt();
            onNavigate("settings");
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
  tone: "good" | "warning" | "poor" | "neutral";
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
  const activeTokenRef = React.useRef(sessionToken);
  activeTokenRef.current = sessionToken;
  const [validatedToken, setValidatedToken] = React.useState<string | null>(null);
  const [sessionRestoreError, setSessionRestoreError] = React.useState(false);
  const [sessionRetry, setSessionRetry] = React.useState(0);


  React.useEffect(() => {
    if (isConfigPage) return; // ConfigPage resolves Teams appearance without starting meeting sign-in.
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

  const refreshAccountInfo = React.useCallback(
    async (token: string, fallbackUser: CurrentUser | null = null, signal?: AbortSignal) => {
      const account = mergeCurrentUser(
        normalizeCurrentUser(
          await fetchJson<unknown>(`${apiBaseUrl}/api/auth/me`, {
            signal,
            headers: buildAuthHeaders(token),
          })
        ),
        fallbackUser
      );

      if (activeTokenRef.current !== token || signal?.aborted) return account;
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
      if (activeTokenRef.current !== token) return data;
      setDashboard((current) => ({
        ...data,
        currentUser: mergeCurrentUser(data.currentUser, current?.currentUser ?? accountUser),
      }));
      return data;
    },
    [accountUser, apiBaseUrl]
  );

  const joinMeetingSession = React.useCallback(
    async (token: string, context: TeamsMeetingContext, fallbackUser: CurrentUser | null = null, signal?: AbortSignal) => {
      const data = normalizeDashboardData(await fetchJson<unknown>(`${apiBaseUrl}/api/TeamsMeetings/join`, {
        signal,
        method: "POST",
        headers: buildAuthHeaders(token),
        body: JSON.stringify(context),
      }));
      const mergedData = {
        ...data,
        currentUser: mergeCurrentUser(data.currentUser, fallbackUser),
      };
      if (activeTokenRef.current === token && !signal?.aborted) setDashboard(mergedData);
      return mergedData;
    },
    [apiBaseUrl]
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
      setDashboard(null);
      setAccountUser(null);
      setError(null);
      restoreReturnUrl(returnUrl);
    },
    [returnUrl]
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
      !TEAMS_SSO_ENABLED ||
      isConfigPage ||
      isPrivacyPolicyPage ||
      isTermsOfServicePage ||
      authPageMode === "create-account" ||
      authPageMode === "support" ||
      !apiBaseUrl ||
      !isRunningInTeams ||
      !hasCheckedTeamsContext ||
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
    hasCheckedTeamsContext,
    isTermsOfServicePage,
    meetingContext,
    pendingProfileAuth,
    pendingTeamsSsoAccountCreation,
    sessionToken,
  ]);

  React.useEffect(() => {
    if (!sessionToken || !hasCheckedTeamsContext) return;
    let cancelled = false;
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), 15000);
    setSessionRestoreError(false);
    setIsLoading(true);
    setBusyState("checking-session");
    void refreshAccountInfo(sessionToken, null, controller.signal)
      .then(async account => {
        if (cancelled) return;
        if (meetingContext) await joinMeetingSession(sessionToken, meetingContext, account, controller.signal);
        if (!cancelled) setValidatedToken(sessionToken);
      })
      .catch(caught => {
        if (cancelled) return;
        if ((caught as { status?: number }).status === 401) {
          window.localStorage.removeItem(SIGNALTUNER_SESSION_TOKEN_KEY);
          activeTokenRef.current = null;
          setSessionToken(null); setAccountUser(null); setDashboard(null); setValidatedToken(null);
        } else setSessionRestoreError(true);
        setError("Your session could not be restored. Please retry or sign in again.");
      })
      .finally(() => {
        window.clearTimeout(timeout);
        if (!cancelled) { setIsLoading(false); setBusyState("idle"); }
      });
    return () => { cancelled = true; controller.abort(); window.clearTimeout(timeout); };
  }, [hasCheckedTeamsContext, joinMeetingSession, meetingContext, refreshAccountInfo, sessionToken, sessionRetry]);

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
    if (!TEAMS_SSO_ENABLED) return;
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
    if (!TEAMS_SSO_ENABLED) return;
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
      if (!TEAMS_SSO_ENABLED) return;
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

  const inviteParticipants = React.useCallback(async () => {
    setError(null);
    try {
      openMeetingInvitations(meetingContext?.teamsMeetingId ?? "", setError);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
    }
  }, [meetingContext]);

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

  const submitAccountProfileUpdate = React.useCallback(
    async (firstName: string, lastName: string, email: string): Promise<void> => {
      const currentUser = mergeCurrentUser(dashboard?.currentUser ?? normalizeCurrentUser(null), accountUser);
      if (!apiBaseUrl || !sessionToken || !currentUser.userId) {
        throw new Error("SignalTuner is temporarily unavailable. Please try again.");
      }
      await saveAccountProfile(apiBaseUrl, sessionToken, currentUser.userId, firstName, lastName, email);
      if (activeTokenRef.current !== sessionToken) return;
      const updated = { ...currentUser, firstName: firstName.trim(), lastName: lastName.trim(),
        email: email.trim().toLowerCase(), displayName: `${firstName.trim()} ${lastName.trim()}` };
      setAccountUser(updated);
      setDashboard(current => current ? { ...current, currentUser: updated } : current);
      // The profile save is already complete. A background refresh failure must not
      // report it as failed or invite duplicate saves. Keep the valid session token.
      void refreshAccountInfo(sessionToken, updated).catch(() => {});
    },
    [accountUser, apiBaseUrl, dashboard, refreshAccountInfo, sessionToken]
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
    activeTokenRef.current = null;
    setValidatedToken(null);
    setSessionRestoreError(false);
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

  React.useEffect(() => {
    const syncSession = (event: StorageEvent) => {
      if (event.key !== SIGNALTUNER_SESSION_TOKEN_KEY) return;
      activeTokenRef.current = event.newValue;
      setValidatedToken(null); setDashboard(null); setAccountUser(null); setAnalysis(null);
      setPendingProfileAuth(null); setPendingTeamsSsoAccountCreation(null);
      setSubscriptionPrompt(null); setIsClientPromptDismissed(false); setActivePage("dashboard");
      setSessionToken(event.newValue);
      if (!event.newValue) signOut();
    };
    window.addEventListener("storage", syncSession);
    return () => window.removeEventListener("storage", syncSession);
  }, [signOut]);

  if (isConfigPage) {
    return <ConfigPage />;
  }

  if (isPrivacyPolicyPage) {
    return <LegalDocumentPage markdown={privacyPolicyMarkdown} title="SignalTuner Privacy Policy" />;
  }

  if (isTermsOfServicePage) {
    return <LegalDocumentPage markdown={termsOfServiceMarkdown} title="SignalTuner Terms of Service" />;
  }

  if (sessionToken && (validatedToken !== sessionToken || (meetingContext && !dashboard))) {
    if (sessionRestoreError) return <main className="pageShell"><section className="panel personalPanel">
      <h1>Unable to restore your session</h1><p role="alert">{error}</p>
      <button className="primaryButton" onClick={() => setSessionRetry(value => value + 1)}>Retry</button>
      <button className="secondaryButton" onClick={signOut}>Sign out</button>
    </section></main>;
    return <SessionLoadingPage />;
  }
  if (sessionToken && accountUser && !meetingContext) {
    return <ActivationContext.Provider value={{ apiBaseUrl, token: sessionToken }}>
      <PersonalDashboard key={`${accountUser.userId}:${sessionToken}`} user={accountUser} apiBaseUrl={apiBaseUrl} token={sessionToken}
        onSignOut={signOut} themePreference={themePreference} onThemePreferenceChange={setThemePreference}
        onUpdateProfile={submitAccountProfileUpdate} onUpdatePassword={submitAccountPasswordUpdate}
        onSendSupportEmail={submitSupportRequest} validatePassword={validatePassword} />
    </ActivationContext.Provider>;
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
          onInvite={inviteParticipants}
          onNavigate={setActivePage}
          onDismissSubscriptionPrompt={() => setSubscriptionPrompt(null)}
          onSendSupportEmail={submitSupportRequest}
          onSignOut={signOut}
          onThemePreferenceChange={setThemePreference}
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
