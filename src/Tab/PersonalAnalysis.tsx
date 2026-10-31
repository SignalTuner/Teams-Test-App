import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import LiveAnalysisCharts from "./LiveAnalysisCharts";
import AnalysisCountdown, { AnalysisCanceled } from "./AnalysisCountdown";
import AnalysisSection from "./AnalysisSection";
import AnalysisConfirmationDialog from "./AnalysisConfirmationDialog";
import { analysisRateLimitMessage } from "./analysisErrors";
import { remoteAnalysisLabels, type RequestState } from "./RemoteAnalysisButton";
import { activeAnalysisStates, pollAnalysisRequest } from "./analysisPolling";
import { pendingAnalysisKey, readPendingAnalysis, sendSelfAnalysis, usePersonalResource, type PersonalApi, type SelfOperation } from "./personalDashboardData";

type SelfRequest = RequestState & { mode: "Test" | "Report" };
export default function PersonalAnalysis({ api, userId, displayName, credits, loadingCredits, onCreditsChanged, children }: {
  api: PersonalApi; userId: number; displayName: string; credits: number | null; loadingCredits?: boolean; onCreditsChanged: () => void; children: ReactNode;
}) {
  const desktop = usePersonalResource<{ connected: boolean; supportsTestAnalysis: boolean }>(api, `remote-analysis/current-device?targetUserId=${userId}`);
  const [request, setRequest] = useState<SelfRequest | null>(null);
  const [expanded, setExpanded] = useState(false);
  const [pending, setPending] = useState(() => readPendingAnalysis(sessionStorage, userId));
  const [busy, setBusy] = useState(false), [error, setError] = useState("");
  const [reportOpen, setReportOpen] = useState(false);
  const active = Boolean(request && activeAnalysisStates.includes(request.state));
  const mounted = useRef(false), mutating = useRef(false), version = useRef(0);
  const controllers = useRef(new Set<AbortController>());
  const readLatest = useCallback(async (signal: AbortSignal) => {
    const latest = await api<{ request: SelfRequest | null }>(`remote-analysis/latest?targetUserId=${userId}&mode=Any`, { signal });
    return latest.request ? api<SelfRequest>(`remote-analysis/${latest.request.requestId}?includeLive=true`, { signal }) : null;
  }, [api, userId]);
  useEffect(() => {
    mounted.current = true;
    const activeControllers = controllers.current;
    return () => { mounted.current = false; activeControllers.forEach(controller => controller.abort()); };
  }, []);
  useEffect(() => {
    let cancelled = false;
    let controller: AbortController | undefined;
    let fetching = false;
    const refresh = async () => {
      if (active || mutating.current || fetching || document.hidden) return;
      fetching = true;
      controller = new AbortController();
      const generation = version.current;
      const timeout = setTimeout(() => controller?.abort(), 15000);
      try {
        const value = await readLatest(controller.signal);
        if (!cancelled && generation === version.current) setRequest(value);
      } catch { if (!cancelled && generation === version.current) setError("Your previous analysis could not be checked. Select an analysis action to check again."); }
      finally { clearTimeout(timeout); fetching = false; }
    };
    void refresh();
    const timer = setInterval(() => void refresh(), 30000);
    window.addEventListener("focus", refresh);
    document.addEventListener("visibilitychange", refresh);
    return () => { cancelled = true; clearInterval(timer); controller?.abort(); window.removeEventListener("focus", refresh); document.removeEventListener("visibilitychange", refresh); };
  }, [readLatest, active]);
  useEffect(() => { if (active || pending || error) setExpanded(true); }, [active, request?.requestId, pending, error]);
  const sending = request?.deliveries?.some(delivery => ["Queued", "Sending"].includes(delivery.state));
  useEffect(() => {
    if (!request || (!active && !sending)) return;
    return pollAnalysisRequest({
      read: signal => api<SelfRequest>(`remote-analysis/${request.requestId}?includeLive=true`, { signal }),
      onData: value => { setRequest(value); setError(""); },
      onError: () => setError("Analysis status could not be refreshed. Your request is unchanged."),
      onAccessLost: () => { setRequest(null); setError("Analysis is no longer available. Select an analysis action to check access."); },
      version: () => version.current, isMutating: () => mutating.current,
    });
  }, [api, request?.requestId, active, sending]);
  useEffect(() => { onCreditsChanged(); }, [request?.creditState, onCreditsChanged]);

  async function submit(mode: SelfOperation["mode"]) {
    if (mutating.current || active) return;
    mutating.current = true; version.current++; setBusy(true); setError("");
    const controller = new AbortController(); controllers.current.add(controller);
    const timeout = setTimeout(() => controller.abort(), 15000);
    let submitted = false;
    try {
      const latest = await readLatest(controller.signal);
      if (!mounted.current) return;
      setRequest(latest);
      let operation = pending;
      if (latest && latest.requestId === operation?.requestId) {
        sessionStorage.removeItem(pendingAnalysisKey(userId)); setPending(null); operation = null;
      }
      if (latest && activeAnalysisStates.includes(latest.state)) {
        setExpanded(true); setReportOpen(false); return;
      }
      if (operation && operation.mode !== mode) {
        setError(`An earlier ${operation.mode === "Test" ? "Test Analysis" : "Analysis + Report"} request is unconfirmed. Use that action to confirm its status before starting a different analysis.`);
        return;
      }
      // Recover an uncertain request with its original ID even if desktop
      // availability changed. New requests check current capability first.
      if (!operation) {
        const available = await api<{ connected: boolean; supportsTestAnalysis: boolean }>(`remote-analysis/current-device?targetUserId=${userId}`, { signal: controller.signal });
        if (!mounted.current) return;
        if (!available.connected) { setError("Connect your desktop client to start an analysis."); return; }
        if (mode === "Test" && !available.supportsTestAnalysis) { setError("Update your desktop client to run Test Analysis."); return; }
        operation = { requestId: crypto.randomUUID().replace(/-/g, ""), targetUserId: userId, meetingSessionId: null, mode };
      }
      setPending(operation); submitted = true;
      const value = await sendSelfAnalysis<SelfRequest>(api, sessionStorage, operation, controller.signal);
      if (mounted.current) {
        setRequest(value); setPending(null); setExpanded(true);
        if (activeAnalysisStates.includes(value.state)) setReportOpen(false);
        else if (["NoCredits", "Failed", "Denied", "Expired"].includes(value.state)) setError(remoteAnalysisLabels[value.state]);
      }
    } catch (reason) {
      if (mounted.current) {
        const status = (reason as { status?: number }).status;
        const rejected = status && status < 500 && status !== 408;
        if (submitted && rejected) { sessionStorage.removeItem(pendingAnalysisKey(userId)); setPending(null); }
        setError(status === 402 ? "No report allowance is available. Test Analysis remains available on supported connected desktops."
          : status === 429 ? analysisRateLimitMessage(reason instanceof Error ? reason.message : null)
          : rejected ? "The request was declined. Check desktop availability and report allowance, then select an analysis action to try again."
          : "The request could not be confirmed. Select the same analysis action to safely check it again.");
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
    } catch { if (mounted.current) setError("The action could not be confirmed. Status updates automatically; you can try the action again."); }
    finally { clearTimeout(timeout); controllers.current.delete(controller); mutating.current = false; if (mounted.current) { setBusy(false); onCreditsChanged(); } }
  }
  const blocked = busy || active;
  return <>
    <div className="participantRow">
      {children}
      <div className="participantActions personalAnalysisActions" role="group" aria-label="Your analysis actions">
        <button className="secondaryButton" type="button" title="Five-minute analysis without using report allowance. No PDF or email."
          disabled={blocked} onClick={() => void submit("Test")}>Test Analysis</button>
        <span className="remoteAnalysisAction"><button className="meetingAnalyzeButton" type="button" title="One report credit. Your PDF is emailed when ready."
          disabled={blocked} aria-haspopup="dialog" onClick={() => { setError(""); setReportOpen(true); onCreditsChanged(); }}>Analysis + Report</button></span>
        <p className="personalAnalysisCredits">Available credits: <strong>{credits === null ? "Unavailable" : credits}</strong></p>
        {credits === 0 && <p className="personalAnalysisHint">No report allowance is available. Test Analysis remains available on supported connected desktops.</p>}
        {desktop.loading ? <p role="status">Checking your desktop…</p> : !desktop.data?.connected ? <p className="personalAnalysisHint">Connect your desktop client to start an analysis.</p>
          : !desktop.data.supportsTestAnalysis && <p className="personalAnalysisHint">Update your desktop client to run Test Analysis.</p>}
        {desktop.error && <p role="alert" className="remoteAnalysisError">{desktop.error} <button className="textButton" onClick={desktop.refresh}>Retry availability</button></p>}
      </div>
    </div>
    <AnalysisConfirmationDialog open={reportOpen} name={displayName || "you"} credits={credits} loadingCredits={loadingCredits}
      busy={busy} error={error} onClose={() => setReportOpen(false)}>
      <footer><button type="button" className="meetingAnalyzeButton" disabled={blocked} onClick={() => void submit("Report")}>Request analysis · 1 credit</button></footer>
    </AnalysisConfirmationDialog>
    <AnalysisSection title={`Analysis for ${displayName || "you"}`} expanded={expanded} onToggle={() => setExpanded(value => !value)}>
      {request && <p className={["Running", "Cancelled"].includes(request.state) ? "remoteAnalysisProgress" : "personalAnalysisStatus"} aria-live="polite">
        {request.state === "Cancelled" ? <AnalysisCanceled /> : <span>{request.state === "Completed" && request.mode === "Test"
          ? "Test complete. No report generated or credit used."
          : ["Approved", "Pending"].includes(request.state) ? "Waiting for your desktop to start."
          : remoteAnalysisLabels[request.state] || request.state}</span>}
        {request.state === "Running" && <AnalysisCountdown key={request.requestId} telemetry={request.liveTelemetry} />}
      </p>}
      {error && !reportOpen && <p role="alert" className="remoteAnalysisError">{error}</p>}
      {busy && <p role="status">Updating analysis…</p>}
      <LiveAnalysisCharts key={request?.requestId || "idle"} telemetry={request?.liveTelemetry} running={request?.state === "Running"} requestState={request?.state}
        idleMessage="Start Test Analysis or Analysis + Report to see five minutes of live telemetry." />
      <div className="buttonRow">
        {active && <button className="secondaryButton" disabled={busy} onClick={() => void action("cancel")}>Cancel analysis</button>}
      </div>
      {request && request.mode !== "Test" && <div className="personalAnalysisResult" aria-live="polite">
          {request.deliveries?.map((delivery, index) => <p key={index}>Report email: {remoteAnalysisLabels[delivery.state] || delivery.state}</p>)}
          {request.deliveries?.some(delivery => ["Rejected", "Uncertain"].includes(delivery.state)) &&
            <button className="secondaryButton" disabled={busy} onClick={() => void action("retry-delivery")}>Retry report email</button>}
      </div>}
    </AnalysisSection>
  </>;
}
