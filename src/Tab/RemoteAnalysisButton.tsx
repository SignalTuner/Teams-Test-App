import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import AnalysisSection from "./AnalysisSection";
import AnalysisConfirmationDialog from "./AnalysisConfirmationDialog";
import "./RemoteAnalysisButton.css";
import "./MeetingAnalysisButton.css";
import "./MeetingServiceStatus.css";
import LiveAnalysisCharts, { type LiveTelemetry } from "./LiveAnalysisCharts";
import AnalysisCountdown, { AnalysisCanceled } from "./AnalysisCountdown";
import { activeAnalysisStates, analysisStatusPath, pollAnalysisRequest } from "./analysisPolling";
import { loadReportAccount } from "./settingsAccount";
import { analysisRateLimitMessage } from "./analysisErrors";

export type RequestState = {
  requestId: string; state: string; creditState: string; targetUserId?: number;
  liveTelemetry?: LiveTelemetry | null;
  deliveries: Array<{ recipient: string; state: string }>;
};
type Props = {
  apiBaseUrl: string; token: string; targetUserId: number; targetName: string;
  meetingSessionId?: number; backgroundRequest?: RequestState;
  expanded?: boolean; onToggle?: () => void; onExpand?: () => void; details?: ReactNode;
  renderInline?: (parts: { action: ReactNode; panel: ReactNode }) => ReactNode;
};

const activeStates = activeAnalysisStates;
export const remoteAnalysisLabels: Record<string, string> = {
  Pending: "Waiting for permission on the device", Approved: "Approved — waiting for the device",
  Running: "Five-minute analysis in progress", Completed: "PDF generated",
  Denied: "The device user declined", Cancelled: "Analysis cancelled", Expired: "Request expired",
  Failed: "Analysis could not finish", NoCredits: "No report allowance is available",
  Queued: "Waiting to send", Sending: "Submitting to email provider", Accepted: "Accepted by email provider",
  Rejected: "Email provider rejected this copy", Uncertain: "Email acceptance could not be confirmed",
};

const labels = remoteAnalysisLabels;

export default function RemoteAnalysisButton({ apiBaseUrl, token, targetUserId, targetName, meetingSessionId, backgroundRequest,
  expanded = false, onToggle, onExpand, details, renderInline }: Props) {
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [connected, setConnected] = useState(false);
  const [credits, setCredits] = useState<number | null>(null);
  const [loadingCredits, setLoadingCredits] = useState(true);
  const [creditRefresh, setCreditRefresh] = useState(0);
  const [loaded, setLoaded] = useState(false);
  const [request, setRequest] = useState<RequestState | null>(null);
  const [lastTelemetry, setLastTelemetry] = useState<{ id: string; data: LiveTelemetry } | null>(null);
  const mutationVersion = useRef(0);
  const mutationBusy = useRef(false);
  const pendingCreation = useRef("");
  const handedOff = useRef<string | null>(null);
  const accessLost = useRef(false);
  const requestId = useRef("");
  const seenBackgroundRequest = useRef<string | null>(null);
  useEffect(() => {
    if (!backgroundRequest || busy || seenBackgroundRequest.current === backgroundRequest.requestId) return;
    seenBackgroundRequest.current = backgroundRequest.requestId;
    mutationVersion.current++;
    setRequest(backgroundRequest);
  }, [backgroundRequest, busy]);
  const context = `targetUserId=${targetUserId}${meetingSessionId ? `&meetingSessionId=${meetingSessionId}` : ""}`;
  useEffect(() => {
    if (!open) return;
    const refresh = () => { if (!document.hidden) setCreditRefresh(value => value + 1); };
    const interval = window.setInterval(refresh, 60000);
    window.addEventListener("focus", refresh);
    document.addEventListener("visibilitychange", refresh);
    return () => {
      window.clearInterval(interval);
      window.removeEventListener("focus", refresh);
      document.removeEventListener("visibilitychange", refresh);
    };
  }, [open]);
  useEffect(() => {
    if (!open) return;
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), 8000);
    let current = true;
    setCredits(null); setLoadingCredits(true);
    loadReportAccount(apiBaseUrl.replace(/\/$/, ""), token, controller.signal)
      .then(account => {
        if (!Number.isSafeInteger(account.availableReports) || account.availableReports < 0) throw new Error();
        if (current) setCredits(account.availableReports);
      })
      .catch(() => { if (current) setCredits(null); })
      .finally(() => { window.clearTimeout(timeout); if (current) setLoadingCredits(false); });
    return () => { current = false; window.clearTimeout(timeout); controller.abort(); };
  }, [open, apiBaseUrl, token, creditRefresh, request?.creditState]);
  const api = useCallback(async <T,>(path: string, method = "GET", body?: unknown, signal?: AbortSignal): Promise<T> => {
    if (!token) throw new Error("Sign in to request an analysis.");
    const response = await fetch(`${apiBaseUrl.replace(/\/$/, "")}/api/remote-analysis${path}`, {
      method, signal, headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw Object.assign(new Error(response.status === 401 ? "Sign in again to request an analysis."
      : response.status === 402 ? "No report allowance is available. Test Analysis remains available on supported connected desktops."
      : response.status === 429 ? analysisRateLimitMessage(data.message)
      : response.status >= 500 ? "Remote analysis is unavailable. Try again shortly."
      : data.message || "Remote analysis is unavailable. Try again shortly."), { status: response.status });
    return data;
  }, [apiBaseUrl, token]);

  useEffect(() => {
    if (!open) return;
    accessLost.current = false;
    let cancelled = false;
    setBusy(true); setError(""); setConnected(false); setLoaded(false);
    const version = ++mutationVersion.current;
    requestId.current = pendingCreation.current || crypto.randomUUID().replace(/-/g, "");
    Promise.all([api<{ connected: boolean }>(`/current-device?${context}`), api<{ request: RequestState | null }>(`/latest?${context}`)])
      .then(([available, latest]) => {
        if (cancelled || version !== mutationVersion.current) return;
        if (latest.request?.requestId === pendingCreation.current) {
          pendingCreation.current = "";
          requestId.current = crypto.randomUUID().replace(/-/g, "");
        }
        // Opening a confirmation inspects an existing request; it is not a new
        // submission and must not automatically dismiss the start form.
        handedOff.current = latest.request?.requestId || null;
        setLoaded(true); setConnected(available.connected); setRequest(latest.request);
      })
      .catch(e => { if (!cancelled) setError(e.message); })
      .finally(() => { if (!cancelled) setBusy(false); });
    return () => { cancelled = true; };
  }, [open, api, context]);

  const inlineEnabled = Boolean(renderInline);
  const polledRequestId = request?.requestId;
  const requestState = request?.state;
  const pendingDelivery = Boolean(request?.deliveries?.some(d => ["Queued", "Sending"].includes(d.state)));
  useEffect(() => {
    if (!polledRequestId) return;
    return pollAnalysisRequest({
      read: signal => api<RequestState>(analysisStatusPath(polledRequestId), "GET", undefined, signal),
      onData: updated => { setRequest(updated); setError(""); },
      onError: () => setError("Status could not be refreshed. Your existing request is unchanged."),
      onAccessLost: () => {
        accessLost.current = true;
        mutationVersion.current++;
        setRequest(null); setLastTelemetry(null);
        setError("This analysis is no longer available. Reopen Analyze to check your access.");
      },
      version: () => mutationVersion.current,
      isMutating: () => mutationBusy.current,
    });
  }, [open, api, inlineEnabled, polledRequestId, requestState, pendingDelivery]);

  // Restore this requester's meeting-scoped analysis when a row is opened after reload.
  useEffect(() => {
    if (!expanded || open || polledRequestId || targetUserId <= 0 || accessLost.current) return;
    let cancelled = false;
    const controller = new AbortController();
    const version = mutationVersion.current;
    api<{ request: RequestState | null }>(`/latest?${context}`, "GET", undefined, controller.signal)
      .then(latest => { if (!cancelled && version === mutationVersion.current) setRequest(latest.request); })
      .catch(() => { if (!cancelled) setError("Analysis history could not be loaded. Reopen the section to retry."); });
    return () => { cancelled = true; controller.abort(); };
  }, [expanded, open, polledRequestId, targetUserId, api, context]);

  useEffect(() => {
    if (request?.liveTelemetry?.snapshot) setLastTelemetry({ id: request.requestId, data: request.liveTelemetry });
  }, [request]);
  const telemetry = request?.liveTelemetry ?? (lastTelemetry?.id === polledRequestId ? lastTelemetry?.data : undefined);
  useEffect(() => {
    if (busy || !activeStates.includes(requestState || "") || !polledRequestId || handedOff.current === polledRequestId) return;
    handedOff.current = polledRequestId;
    setOpen(false); onExpand?.();
  }, [busy, polledRequestId, requestState, onExpand]);

  async function submit(path: string, body?: unknown) {
    if (busy || mutationBusy.current) return;
    if (path === "" && request && activeStates.includes(request.state)) return;
    mutationBusy.current = true; mutationVersion.current++;
    setBusy(true); setError("");
    if (path === "") pendingCreation.current = requestId.current;
    try {
      const updated = await api<RequestState>(path, "POST", body);
      setRequest(updated);
      if (path === "") {
        pendingCreation.current = "";
        requestId.current = crypto.randomUUID().replace(/-/g, "");
      }
    }
    catch (e) { setError(e instanceof Error ? e.message : "The request could not be confirmed. Retry the same request."); }
    finally { mutationBusy.current = false; setBusy(false); }
  }

  const isActive = request && activeStates.includes(request.state);
  const waitingForApproval = ["Pending", "Approved"].includes(requestState || "");
  const inlineProgress = inlineEnabled && !open && (waitingForApproval || ["Running", "Cancelled"].includes(requestState || ""));
  const progress = <span className={`remoteAnalysisProgress${waitingForApproval ? " remoteAnalysisWaiting" : requestState === "Running" ? " remoteAnalysisRunning" : ""}`} role="status">
    {waitingForApproval ? "Waiting for approval..." : requestState === "Cancelled" ? <AnalysisCanceled />
      : requestState === "Running" ? <>{labels.Running}<AnalysisCountdown key={polledRequestId} telemetry={telemetry} /></>
      : labels[requestState || ""] || requestState}
  </span>;
  const creditMessages: Record<string, string> = {
    Reserved: "One credit is reserved until the report completes. It is returned if analysis fails or is cancelled.",
    Consumed: "One report credit used. Email retries use no extra report credit.",
    Released: "The reserved credit was returned.",
  };
  const creditMessage = creditMessages[request?.creditState || ""];
  const requestStatus = request && <section aria-live="polite">
        {inlineEnabled && !inlineProgress && <h3>{request.state === "Cancelled" ? <AnalysisCanceled /> : labels[request.state] || request.state}
          {request.state === "Pending" && <span className="remoteAnalysisWaitingDots" aria-hidden="true"><span>.</span><span>.</span><span>.</span></span>}
          {request.state === "Running" && <AnalysisCountdown key={request.requestId} telemetry={telemetry} />}</h3>}
        {creditMessage && <p>{creditMessage}</p>}
        {request.deliveries?.map(d => <p key={d.recipient}><strong>{d.recipient}:</strong> {labels[d.state] || d.state}</p>)}
        {request.deliveries?.some(d => d.state === "Uncertain") && <p>A copy may already have been sent. It will not be sent again automatically.</p>}
        <footer>
          {isActive && <button type="button" disabled={busy} onClick={() => void submit(`/${request.requestId}/cancel`)}>Cancel analysis</button>}
          {request.deliveries?.some(d => d.state === "Rejected") && <button type="button" disabled={busy}
            onClick={() => void submit(`/${request.requestId}/retry-delivery`)}>Retry rejected email</button>}
        </footer>
      </section>;
  const action = <div className="remoteAnalysisAction" onClick={e => e.stopPropagation()}>
    <button type="button" className="meetingAnalyzeButton remoteAnalyzeButton" disabled={targetUserId <= 0}
      onClick={() => setOpen(true)} title={targetUserId <= 0 ? "This participant must sign in to SignalTuner first" : "Request a PDF analysis"}>Analyze</button>
    {!inlineEnabled && !open && (request || error) && <div className="remoteAnalysisCompactStatus">
      <strong>Analysis for {targetName}</strong>{request && progress}
      {error && <span role="alert" className="remoteAnalysisError">{error}</span>}
      {requestStatus}
    </div>}
    <AnalysisConfirmationDialog open={open} name={targetName} credits={credits} loadingCredits={loadingCredits}
      busy={busy} error={error} onClose={() => setOpen(false)}>
      {loaded && <section>
        {isActive && <p>An analysis is already in progress for this user. Its status and cancellation controls are in the participant analysis section.</p>}
        {!connected && <p>The updated SignalTuner desktop app must be running and signed in on this user’s device. Reopen this dialog after they connect.</p>}
          <footer><button type="button" className="meetingAnalyzeButton" disabled={busy || !connected || Boolean(isActive)} onClick={() => void submit("", {
            requestId: requestId.current, targetUserId, meetingSessionId: meetingSessionId || null,
          })}>Request analysis · 1 credit</button></footer>
      </section>}
    </AnalysisConfirmationDialog>
  </div>;
  const panel = <AnalysisSection title={`Analysis for ${targetName}`} expanded={expanded} onToggle={onToggle}
      summary={inlineProgress ? progress : undefined}>
      {!open && error && <p role="alert" className="remoteAnalysisError">{error}</p>}
      {!open && busy && <p role="status">Updating...</p>}
      <LiveAnalysisCharts telemetry={telemetry} running={request?.state === "Running"} requestState={request?.state}
        idleMessage={targetUserId <= 0 ? "This participant must sign in to SignalTuner before analysis is available." : undefined} />
      {!open && requestStatus}
      {!request && details}

  </AnalysisSection>;
  return renderInline ? renderInline({ action, panel }) : action;
}
