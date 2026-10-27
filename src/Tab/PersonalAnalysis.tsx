import { useEffect, useRef, useState, type ReactNode } from "react";
import LiveAnalysisCharts from "./LiveAnalysisCharts";
import AnalysisCountdown from "./AnalysisCountdown";
import AnalysisSection from "./AnalysisSection";
import { remoteAnalysisLabels, type RequestState } from "./RemoteAnalysisButton";
import { activeAnalysisStates, pollAnalysisRequest } from "./analysisPolling";
import { pendingAnalysisKey, readPendingAnalysis, sendSelfAnalysis, usePersonalResource, type PersonalApi, type SelfOperation } from "./personalDashboardData";

type SelfRequest = RequestState & { mode: "Test" | "Report" };
export default function PersonalAnalysis({ api, userId, credits, onCreditsChanged, children }: {
  api: PersonalApi; userId: number; credits: number | null; onCreditsChanged: () => void; children: ReactNode;
}) {
  const desktop = usePersonalResource<{ connected: boolean; supportsTestAnalysis: boolean }>(api, `remote-analysis/current-device?targetUserId=${userId}`);
  const [request, setRequest] = useState<SelfRequest | null>(null);
  const [expanded, setExpanded] = useState(false);
  const [pending, setPending] = useState(() => readPendingAnalysis(sessionStorage, userId));
  const [ready, setReady] = useState(false), [busy, setBusy] = useState(false), [error, setError] = useState("");
  const [revision, setRevision] = useState(0);
  const mounted = useRef(false), mutating = useRef(false), version = useRef(0);
  const controllers = useRef(new Set<AbortController>());
  useEffect(() => {
    mounted.current = true;
    const activeControllers = controllers.current;
    return () => { mounted.current = false; activeControllers.forEach(controller => controller.abort()); };
  }, []);
  useEffect(() => {
    let cancelled = false;
    const controller = new AbortController();
    const generation = version.current;
    setReady(false);
    const timeout = setTimeout(() => controller.abort(), 15000);
    void (async () => {
      try {
        const latest = await api<{ request: SelfRequest | null }>(`remote-analysis/latest?targetUserId=${userId}&mode=Any`, { signal: controller.signal });
        const value = latest.request ? await api<SelfRequest>(`remote-analysis/${latest.request.requestId}?includeLive=true`, { signal: controller.signal }) : null;
        if (!cancelled && generation === version.current) { setRequest(value); setReady(true); setError(""); }
      } catch { if (!cancelled) setError("Your previous analysis could not be checked. Refresh status before starting another."); }
      finally { clearTimeout(timeout); }
    })();
    return () => { cancelled = true; controller.abort(); };
  }, [api, userId, revision]);
  const active = Boolean(request && activeAnalysisStates.includes(request.state));
  useEffect(() => { if (active || pending || error) setExpanded(true); }, [active, request?.requestId, pending, error]);
  const sending = request?.deliveries?.some(delivery => ["Queued", "Sending"].includes(delivery.state));
  useEffect(() => {
    if (!request || (!active && !sending)) return;
    return pollAnalysisRequest({
      read: signal => api<SelfRequest>(`remote-analysis/${request.requestId}?includeLive=true`, { signal }),
      onData: value => { setRequest(value); setError(""); },
      onError: () => setError("Analysis status could not be refreshed. Your request is unchanged."),
      onAccessLost: () => { setRequest(null); setReady(false); setError("Analysis is no longer available. Refresh status to check access."); },
      version: () => version.current, isMutating: () => mutating.current,
    });
  }, [api, request?.requestId, active, sending]);
  useEffect(() => { onCreditsChanged(); }, [request?.creditState, onCreditsChanged]);

  async function submit(mode: SelfOperation["mode"], retry = false) {
    if (mutating.current || (!retry && (!ready || active || pending))) return;
    const operation = retry ? pending : { requestId: crypto.randomUUID().replace(/-/g, ""), targetUserId: userId, meetingSessionId: null, mode };
    if (!operation) return;
    mutating.current = true; version.current++; setBusy(true); setError(""); setPending(operation);
    const controller = new AbortController(); controllers.current.add(controller);
    const timeout = setTimeout(() => controller.abort(), 15000);
    try {
      const value = await sendSelfAnalysis<SelfRequest>(api, sessionStorage, operation, controller.signal);
      if (mounted.current) { setRequest(value); setPending(null); setReady(true); }
    } catch (reason) {
      if (mounted.current) {
        const status = (reason as { status?: number }).status;
        const rejected = status && status < 500 && status !== 408;
        if (rejected) { sessionStorage.removeItem(pendingAnalysisKey(userId)); setPending(null); setReady(false); }
        setError(rejected ? "The request was declined. Refresh status and check desktop availability and credits."
          : "The request could not be confirmed. Retry the same analysis to avoid a duplicate request.");
      }
    } finally {
      clearTimeout(timeout); controllers.current.delete(controller); mutating.current = false;
      if (mounted.current) { setBusy(false); onCreditsChanged(); }
    }
  }
  async function action(path: "cancel" | "retry-delivery") {
    if (!request || mutating.current) return;
    mutating.current = true; version.current++; setBusy(true);
    const controller = new AbortController(); controllers.current.add(controller);
    const timeout = setTimeout(() => controller.abort(), 15000);
    try {
      const value = await api<SelfRequest>(`remote-analysis/${request.requestId}/${path}`, { method: "POST", signal: controller.signal });
      if (mounted.current) { setRequest(value); setError(""); }
    } catch { if (mounted.current) setError("The action could not be confirmed. Refresh status before retrying."); }
    finally { clearTimeout(timeout); controllers.current.delete(controller); mutating.current = false; if (mounted.current) { setBusy(false); onCreditsChanged(); } }
  }
  const blocked = busy || active || Boolean(pending) || !ready || !desktop.data?.connected;
  return <>
    <div className="participantRow">
      {children}
      <div className="participantActions personalAnalysisActions" role="group" aria-label="Your analysis actions">
        <button className="secondaryButton" type="button" title="Free five-minute analysis. No credits, PDF or email."
          disabled={blocked || !desktop.data?.supportsTestAnalysis} onClick={() => void submit("Test")}>Test Analysis</button>
        <span className="remoteAnalysisAction"><button className="meetingAnalyzeButton" type="button" title="One report credit. Your PDF is emailed when ready."
          disabled={blocked || credits === null || credits < 1} onClick={() => void submit("Report")}>Analysis + Report</button></span>
        <p className="personalAnalysisCredits">Available credits: <strong>{credits === null ? "Unavailable" : credits}</strong></p>
        {credits === 0 && <p className="personalAnalysisHint">Test Analysis is free. A report requires one credit.</p>}
        {desktop.loading ? <p role="status">Checking your desktop…</p> : !desktop.data?.connected ? <p className="personalAnalysisHint">Connect your desktop client to start an analysis.</p>
          : !desktop.data.supportsTestAnalysis && <p className="personalAnalysisHint">Update your desktop client to enable Test Analysis.</p>}
        {desktop.error && <p role="alert">{desktop.error} <button className="textButton" onClick={desktop.refresh}>Retry availability</button></p>}
      </div>
    </div>
    <AnalysisSection title="Your analysis telemetry" expanded={expanded} onToggle={() => setExpanded(value => !value)}>
      {error && <p role="alert">{error}</p>}
      <div className="buttonRow">
        <button className="secondaryButton" disabled={busy} onClick={() => { setRevision(value => value + 1); desktop.refresh(); onCreditsChanged(); }}>Refresh analysis status</button>
        {pending && <button className="primaryButton" disabled={busy} onClick={() => void submit(pending.mode, true)}>Retry same analysis</button>}
        {active && <button className="secondaryButton" disabled={busy} onClick={() => void action("cancel")}>Cancel analysis</button>}
      </div>
      {busy && <p role="status">Updating analysis…</p>}
      {request && <div className="personalAnalysisResult" aria-live="polite">
        <h3>{request.mode === "Test" ? "Test Analysis" : "Report analysis"}</h3>
        <p>{request.state === "Completed" && request.mode === "Test" ? "Test complete. No report generated or credit used."
          : ["Approved", "Pending"].includes(request.state) ? "Waiting for your desktop to start."
          : remoteAnalysisLabels[request.state] || request.state}</p>
        {request.state === "Running" && <AnalysisCountdown key={request.requestId} telemetry={request.liveTelemetry} />}
        {request.mode !== "Test" && <><p>Credit status: {request.creditState}</p>
          {request.deliveries?.map((delivery, index) => <p key={index}>Report email: {remoteAnalysisLabels[delivery.state] || delivery.state}</p>)}
          {request.deliveries?.some(delivery => ["Rejected", "Uncertain"].includes(delivery.state)) &&
            <button className="secondaryButton" disabled={busy} onClick={() => void action("retry-delivery")}>Retry report email</button>}</>}
      </div>}
        <LiveAnalysisCharts key={request?.requestId || "idle"} telemetry={request?.liveTelemetry} running={request?.state === "Running"} requestState={request?.state}
        idleMessage="Start Test Analysis or Analysis + Report to see five minutes of live telemetry." />
    </AnalysisSection>
  </>;
}
