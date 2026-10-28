import { useEffect, useId, useMemo, useRef, useState, type ComponentProps } from "react";
import DashboardHeader, { getInitials } from "./DashboardHeader";
import SettingsPage from "./SettingsPage";
import SupportPage from "./SupportPage";
import AppearanceSelector from "./AppearanceSelector";
import ClientPrompt from "./ClientPrompt";
import MeetingServiceStatus from "./MeetingServiceStatus";
import ConnectivitySummary from "./ConnectivitySummary";
import PersonalAnalysis from "./PersonalAnalysis";
import RemoteAnalysisButton from "./RemoteAnalysisButton";
import { createPersonalApi, signalPresentation, teamsHistory, utcDate, usePersonalResource, type MeetingDetail, type MeetingSummary, type PersonalApi } from "./personalDashboardData";
import type { CurrentUser, TelemetryRecord } from "./dashboardTypes";
import type { ReportAccount } from "./settingsAccount";
import "./PersonalDashboard.css";

type Props = Omit<ComponentProps<typeof SettingsPage>, "user" | "onOpenSupport"> & {
  user: CurrentUser; onSignOut: () => void;
  onSendSupportEmail: (request: { sender: string; subject: string; body: string }) => Promise<string>;
};
export default function PersonalDashboard(props: Props) {
  const { apiBaseUrl, token, onSignOut } = props;
  const api = useMemo(() => createPersonalApi(apiBaseUrl, token, onSignOut), [apiBaseUrl, token, onSignOut]);
  const account = usePersonalResource<CurrentUser>(api, "auth/me");
  const creditAccount = usePersonalResource<ReportAccount>(api, "credits/account");
  const telemetry = usePersonalResource<TelemetryRecord>(api, "User/signal/current", 10000);
  const meetings = usePersonalResource<{ meetings: MeetingSummary[] }>(api, "meetings?includeTeams=true", 60000);
  const [page, setPage] = useState<"dashboard" | "settings" | "support">("dashboard");
  const [selectedMeeting, setSelectedMeeting] = useState<string | null>(null);
  const [meetingsExpanded, setMeetingsExpanded] = useState(false);
  const [accountExpanded, setAccountExpanded] = useState(false);
  const [activationDismissed, setActivationDismissed] = useState(false);
  const [lastConfirmedUser, setLastConfirmedUser] = useState(props.user);
  useEffect(() => { if (account.data) setLastConfirmedUser(account.data); }, [account.data]);
  const user = account.data || lastConfirmedUser;
  const signal = signalPresentation(telemetry.data);
  const connected = Boolean(user.clientIsActive && !signal.offline);
  const balance = creditAccount.data?.availableReports;
  const credits = typeof balance === "number" && Number.isFinite(balance) ? balance : null;
  // Match the in-meeting sign-in gate using the backend's reporting-data flag.
  // A signal-read failure or loading state is not evidence that setup is needed.
  const showActivation = !user.clientIsActive && !activationDismissed;
  const refreshConnection = async () => { await Promise.all([account.refresh(), telemetry.refresh()]); };
  const history = teamsHistory(meetings.data?.meetings || []);
  if (showActivation) return <main className="pageShell">
    <ClientPrompt activationCodeError={account.error || null} isLoading={account.loading}
      onContinue={() => { setPage("dashboard"); setActivationDismissed(true); }}
      onRefresh={refreshConnection} onSignOut={onSignOut} />
  </main>;
  return <main className="pageShell dashboardShell personalDashboard">
    <DashboardHeader user={user} activePage={page} onNavigate={next => { setPage(next); if (next === "dashboard") setSelectedMeeting(null); }} onSignOut={onSignOut} />
    {page === "settings" && <SettingsPage {...props} user={user} onOpenSupport={() => setPage("support")} onUpdateProfile={async (...args) => { await props.onUpdateProfile(...args); account.refresh(); }}
      onUpdatePassword={async value => { const result = await props.onUpdatePassword(value); account.refresh(); return result; }} />}
    {page === "support" && <SupportPage defaultName={user.displayName} defaultEmail={user.email} onSendSupportEmail={props.onSendSupportEmail}
      themePreference={props.themePreference} onThemePreferenceChange={props.onThemePreferenceChange} />}
    {page === "dashboard" && <>
      <div className="dashboardAppearanceRow"><AppearanceSelector value={props.themePreference} onChange={props.onThemePreferenceChange} /></div>
        <section className="panel participantsPanel" aria-labelledby="personal-health-title">
          <div className="participantsToolbar"><div><h2 id="personal-health-title">Your Connectivity Health</h2>
            <p>{account.error ? "Desktop connection unavailable" : account.loading ? "Checking your desktop…" : connected ? "Desktop connected" : "Desktop disconnected"}</p></div>
            <button className="secondaryButton" onClick={() => void refreshConnection()}>Refresh status</button></div>
          {telemetry.loading && <p className="participantsEmpty" role="status">Loading your connectivity…</p>}
          {telemetry.error && <p className="participantsEmpty" role="alert">{telemetry.error}</p>}
          <PersonalAnalysis api={api} userId={user.userId} credits={credits} onCreditsChanged={creditAccount.refresh}>
          <div className="participantIdentity">
            <span className="participantAvatar" aria-hidden="true">{getInitials(user.displayName, user.email)}</span>
            <div><h3 className="participantName">{user.displayName || "You"}</h3>
              {user.email && <p className="participantEmail" title={user.email}>{user.email}</p>}
              {!user.clientIsActive && <button className="textButton" onClick={() => setActivationDismissed(false)}>Connect desktop client</button>}
            </div>
          </div>
          <ConnectivitySummary name={user.displayName || "you"} score={signal.score} overallStatus={connected ? signal.overall : signal.offline ? "Offline" : null}
            hasData={signal.hasData && connected} workspaceStatus={stringMetric(telemetry.data, "SignalWifiStatus")}
            networkStatus={stringMetric(telemetry.data, "SignalBandwidthStatus")} deviceStatus={stringMetric(telemetry.data, "SignalSystemStatus")} />
          </PersonalAnalysis>
        </section>
        <MeetingServiceStatus apiBaseUrl={apiBaseUrl} token={token} />
        <section className="panel personalPanel" aria-labelledby="personal-meetings-title">
          <h2 id="personal-meetings-title" className="personalSectionHeading">
            <button className="personalSectionToggle" type="button" aria-expanded={meetingsExpanded} aria-controls="personal-meetings-content"
              onClick={() => setMeetingsExpanded(value => !value)}>
              <span>Your Meetings</span><span className="teamsServiceDisclosure" aria-hidden="true" />
            </button>
          </h2>
          <div id="personal-meetings-content" hidden={!meetingsExpanded}>
          <div className="personalSectionToolbar"><button className="secondaryButton" onClick={meetings.refresh}>Refresh meetings</button></div>
          <p>Previously joined SignalTuner meetings. These records do not indicate whether a Teams call is active.</p>
          {meetings.loading && <p role="status">Loading meetings…</p>}
          {meetings.error && <p role="alert">{meetings.error}</p>}
          {!meetings.loading && !meetings.error && history.length === 0 && <p>No Teams meetings yet. Open SignalTuner inside a meeting to join its dashboard.</p>}
          <ul className="personalMeetingList">{history.map(meeting => <li key={meeting.meetingId}>
            <div><h3>{meeting.name}</h3><p>{utcDate(meeting.createdAt)?.toLocaleString() || "Date unavailable"} · {meeting.participantCount} authenticated SignalTuner participants</p></div>
            <button className="secondaryButton" onClick={() => setSelectedMeeting(meeting.meetingId)}>View Dashboard</button>
          </li>)}</ul>
          </div>
        </section>
        <section className="panel personalPanel" aria-labelledby="personal-account-title">
          <h2 id="personal-account-title" className="personalSectionHeading">
            <button className="personalSectionToggle" type="button" aria-expanded={accountExpanded} aria-controls="personal-account-content"
              onClick={() => setAccountExpanded(value => !value)}>
              <span>Account Information</span><span className="teamsServiceDisclosure" aria-hidden="true" />
            </button>
          </h2>
          <div id="personal-account-content" hidden={!accountExpanded}>
          <div className="personalSectionToolbar"><button className="secondaryButton" onClick={() => setPage("settings")}>Settings</button></div>
          {account.error && <p role="alert">{account.error} <button className="textButton" onClick={account.refresh}>Retry account</button></p>}
          {creditAccount.error && <p role="alert">{creditAccount.error} <button className="textButton" onClick={creditAccount.refresh}>Retry credit balance</button></p>}
          <dl className="personalAccount"><div><dt>Name</dt><dd>{user.displayName || "Unavailable"}</dd></div><div><dt>Email</dt><dd>{user.email || "Unavailable"}</dd></div>
            <div><dt>Sign-in identity</dt><dd>{({ teams_sso: "Microsoft Teams", google: "Google", github: "GitHub", email_magic_code: "Email", email: "Email" } as Record<string, string>)[user.authProvider || ""] || "SignalTuner account"}</dd></div><div><dt>Available report credits</dt><dd>{creditAccount.loading ? "Loading…" : credits ?? "Unavailable"}</dd></div>
            <div><dt>Plan</dt><dd>{creditAccount.data ? creditAccount.data.subscription?.name || "Pay as you go" : user.subscriptionPlan || "Unavailable"}</dd></div>
            <div><dt>Desktop client</dt><dd>{account.error ? "Unavailable" : connected ? "Connected" : "Disconnected"}</dd></div></dl>
          </div>
        </section>
      {selectedMeeting && <HistoricalMeeting key={selectedMeeting} api={api} id={selectedMeeting} apiBaseUrl={apiBaseUrl} token={token} onClose={() => setSelectedMeeting(null)} />}
    </>}
  </main>;
}
function stringMetric(data: TelemetryRecord | null, key: string) { return typeof data?.[key] === "string" ? data[key] as string : null; }

function HistoricalMeeting({ api, id, apiBaseUrl, token, onClose }: {
  api: PersonalApi; id: string; apiBaseUrl: string; token: string; onClose: () => void;
}) {
  const detail = usePersonalResource<MeetingDetail>(api, `meetings/${encodeURIComponent(id)}`);
  const [expanded, setExpanded] = useState<number | null>(null);
  const dialog = useRef<HTMLDialogElement>(null);
  const titleId = useId();
  useEffect(() => {
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const previousOverflow = document.body.style.overflow;
    if (!dialog.current?.open) dialog.current?.showModal();
    document.body.style.overflow = "hidden";
    return () => { document.body.style.overflow = previousOverflow; if (opener?.isConnected) opener.focus(); };
  }, []);
  return <dialog ref={dialog} className="personalMeetingModal" aria-labelledby={titleId}
    onCancel={event => { if (event.target === event.currentTarget) { event.preventDefault(); event.currentTarget.close(); } }}
    onClose={event => { if (event.target === event.currentTarget) onClose(); }}>
    <header className="personalMeetingModalHeader">
      <h2 id={titleId}>{detail.data?.meeting.name || "Meeting dashboard"}</h2>
      <button autoFocus className="secondaryButton" type="button" aria-label="Close meeting dashboard" onClick={() => dialog.current?.close()}>Close</button>
    </header>
    <div className="personalMeetingModalBody">
    <div className="buttonRow"><button className="secondaryButton" onClick={detail.refresh}>Refresh meeting</button></div>
    <section className="panel personalPanel">
      {detail.loading && <p role="status">Loading meeting…</p>}
      {detail.error && <p role="alert">{detail.error}</p>}
      {detail.data && <><p>Historical meeting · {utcDate(detail.data.meeting.createdAt)?.toLocaleString()}</p>
        <p>These are previously signed-in participants and their latest available connectivity, not live call attendance or telemetry recorded at meeting time.</p>
        <ul className="participantList">{detail.data.participants.map(person => {
          const signal = person.signal;
          const hasData = Boolean(signal && signal.signalOverallStatus && signal.signalOverallStatus !== "Offline");
          const row = (action?: React.ReactNode, panel?: React.ReactNode) => <li className="participantListItem" key={person.participantId}>
            <div className="participantRow"><div className="participantIdentity"><span className="participantAvatar" aria-hidden="true">{getInitials(person.name, person.email)}</span>
              <div><h3 className="participantName">{person.name}</h3><p className="participantEmail">{person.email}</p><span className="participantRole">{person.meetingRole || "Participant"}</span></div></div>
              <ConnectivitySummary name={person.name} score={signal?.signalScore ?? null} overallStatus={signal?.signalOverallStatus ?? null} hasData={hasData}
                workspaceStatus={signal?.signalWifiStatus ?? null} networkStatus={signal?.signalBandwidthStatus ?? null} deviceStatus={signal?.signalSystemStatus ?? null} />
              <div className="participantActions">{action}</div></div>{panel}</li>;
          return detail.data!.canAnalyze ? <RemoteAnalysisButton key={person.participantId} apiBaseUrl={apiBaseUrl} token={token} targetUserId={person.userId} targetName={person.name}
            meetingSessionId={detail.data!.meeting.meetingSessionId} expanded={expanded === person.userId}
            backgroundRequest={detail.data!.requests?.find(request => request.targetUserId === person.userId)}
            onToggle={() => setExpanded(value => value === person.userId ? null : person.userId)} onExpand={() => setExpanded(person.userId)}
            renderInline={({ action, panel }) => row(action, panel)} /> : row();
        })}</ul></>}
    </section>
    </div>
  </dialog>;
}
