import { useEffect, useRef, useState } from "react";
import { remoteAnalysisLabels } from "./RemoteAnalysisButton";
import { activeMeetingParticipantIds, readMeetingAnalysisOperation, requestMeetingAnalysis,
  type MeetingAnalysisOperation, type MeetingAnalysisResult } from "./meetingAnalysis";
import "./MeetingAnalysisButton.css";

type Props = {
  apiBaseUrl: string; token: string; userId: number; meetingSessionId: number; disabled?: boolean;
  participants: Array<{ userId: number; displayName: string | null; clientDataStatus: string }>;
  onAccepted: (results: MeetingAnalysisResult[]) => void;
};

export default function MeetingAnalysisButton({ apiBaseUrl, token, userId, meetingSessionId, participants, disabled, onAccepted }: Props) {
  const storageKey = `teamsMeetingAnalysis:${userId}:${meetingSessionId}`;
  const [saved] = useState(() => {
    try { return { operation: readMeetingAnalysisOperation(sessionStorage.getItem(storageKey)), error: "" }; }
    catch { return { operation: null, error: "Saved analysis requests could not be read. Allow tab storage and reload before starting an analysis." }; }
  });
  const [operation, setOperation] = useState<MeetingAnalysisOperation | null>(saved.operation);
  const [targets, setTargets] = useState<number[]>(saved.operation?.participantUserIds || []);
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(saved.error);
  const [results, setResults] = useState<MeetingAnalysisResult[]>([]);
  const [credits, setCredits] = useState<number | null>(null);
  const [loadingCredits, setLoadingCredits] = useState(true);
  const [creditRefresh, setCreditRefresh] = useState(0);
  const dialog = useRef<HTMLDialogElement>(null);
  const inFlight = useRef(false);
  const mounted = useRef(true);
  const controller = useRef<AbortController | null>(null);
  const activeIds = activeMeetingParticipantIds(participants);

  useEffect(() => {
    mounted.current = true;
    return () => { mounted.current = false; controller.current?.abort(); };
  }, []);

  useEffect(() => {
    if (open && !dialog.current?.open) dialog.current?.showModal();
  }, [open]);

  useEffect(() => {
    const refresh = () => {
      if (!document.hidden) setCreditRefresh(value => value + 1);
    };
    const interval = window.setInterval(refresh, 60000);
    window.addEventListener("focus", refresh);
    document.addEventListener("visibilitychange", refresh);
    return () => {
      window.clearInterval(interval);
      window.removeEventListener("focus", refresh);
      document.removeEventListener("visibilitychange", refresh);
    };
  }, []);

  useEffect(() => {
    const abort = new AbortController();
    const timeout = setTimeout(() => abort.abort(), 8000);
    let current = true;
    setCredits(null); setLoadingCredits(true);
    fetch(`${apiBaseUrl.replace(/\/$/, "")}/api/credits/account`, {
      signal: abort.signal, cache: "no-store", headers: { Authorization: `Bearer ${token}` },
    }).then(async response => {
      const data = await response.json();
      if (!response.ok || !Number.isSafeInteger(data.availableReports) || data.availableReports < 0) throw new Error();
      if (current) setCredits(data.availableReports);
    }).catch(() => { if (current) setCredits(null); })
      .finally(() => { clearTimeout(timeout); if (current) setLoadingCredits(false); });
    return () => { current = false; clearTimeout(timeout); abort.abort(); };
  }, [open, apiBaseUrl, token, creditRefresh]);

  async function start() {
    if (inFlight.current || disabled || saved.error || !token || !targets.length || targets.length > 1000) return;
    const next = operation || { requestId: crypto.randomUUID().replace(/-/g, ""), participantUserIds: targets };
    // Persist before any request. Reload/retry must never create a second purchase.
    try { sessionStorage.setItem(storageKey, JSON.stringify(next)); }
    catch { setError("The request could not be saved in this tab. Allow tab storage before starting an analysis."); return; }
    setOperation(next);
    inFlight.current = true; setBusy(true); setError("");
    const abort = new AbortController();
    controller.current = abort;
    const timeout = setTimeout(() => abort.abort(), 60000);
    try {
      const accepted = await requestMeetingAnalysis(apiBaseUrl, token, meetingSessionId, next, abort.signal);
      if (!mounted.current) return;
      setResults(accepted); onAccepted(accepted);
    } catch (reason) {
      if (mounted.current) setError(reason instanceof Error && reason.name !== "AbortError" ? reason.message
        : "The response could not be confirmed. Retry this same meeting analysis; accepted requests keep running.");
    } finally {
      clearTimeout(timeout); inFlight.current = false;
      if (mounted.current) { setBusy(false); setCreditRefresh(value => value + 1); }
    }
  }

  function newAnalysis() {
    try { sessionStorage.removeItem(storageKey); }
    catch { setError("The saved request could not be cleared. Allow tab storage and try again."); return; }
    setOperation(null); setTargets(activeIds); setResults([]); setError("");
  }

  const targetCount = operation?.participantUserIds.length || targets.length;
  const acceptedCount = results.filter(result => result.request).length;
  const allAccepted = results.length > 0 && acceptedCount === results.length;
  return <>
    <div className="meetingAnalysisAction">
      <div className="meetingAnalysisCreditBalance" aria-live="polite" aria-atomic="true">
        Available report allowance: <strong>{loadingCredits ? "Loading..." : credits === null ? "Unavailable" : credits.toLocaleString()}</strong>
      </div>
      <button type="button" className="meetingAnalyzeButton" disabled={disabled || busy || (!operation && activeIds.length === 0)}
        aria-haspopup="dialog" onClick={() => { if (!operation) setTargets(activeIds); setOpen(true); }}>
        Analyze Meeting
      </button>
    </div>
    {open && <dialog ref={dialog} className="remoteAnalysisDialog meetingAnalysisDialog"
      aria-labelledby="meeting-analysis-title" onCancel={event => { if (busy) event.preventDefault(); else setOpen(false); }} onClose={() => setOpen(false)}>
      <header><h2 id="meeting-analysis-title">Analyze Meeting</h2>
        <button type="button" disabled={busy} aria-label="Close meeting analysis" onClick={() => setOpen(false)}>Close</button>
      </header>
      <p>Request a five-minute analysis for {targetCount} active {targetCount === 1 ? "participant" : "participants"}, including you if your desktop is active.</p>
      <p>Each completed report is emailed to you and its participant. This uses one unit of report allowance per report, up to <strong>{targetCount} units</strong>.</p>
      <p>Existing team roles and device approval preferences apply. A participant who disconnects or has a pending analysis may be unavailable.</p>
      <p>Available report allowance: {loadingCredits ? "Loading..." : credits ?? "Unavailable"}</p>
      {credits === 0 && <p>No report allowance is available. Test Analysis remains available from your personal dashboard on supported connected desktops.</p>}
      {targetCount > 1000 && <p role="alert">Meeting analysis supports up to 1000 active participants per run.</p>}
      {operation && <p>This run keeps its original participants. Retrying recovers accepted requests and retries only participants without a request.</p>}
      {error && <p role="alert" className="remoteAnalysisError">{error}</p>}
      {busy && <p role="status">Requesting participant analyses...</p>}
      {results.length > 0 && <>
        <p role="status">{acceptedCount} of {results.length} requests accepted. Open Analyze beside a participant for current progress or cancellation.</p>
        <ul className="meetingAnalysisResults" aria-label="Meeting analysis results">
          {results.map(result => <li key={result.targetUserId}>
            <strong>{participants.find(p => p.userId === result.targetUserId)?.displayName || "Participant"}</strong>
            <span>{result.statusCode === 402 ? "No report allowance is available" : result.error || remoteAnalysisLabels[result.request!.state] || result.request!.state}</span>
          </li>)}
        </ul>
      </>}
      <footer>
        {!allAccepted && <button type="button" className="meetingAnalyzeButton" disabled={busy || disabled || Boolean(saved.error) || !targetCount || targetCount > 1000}
          onClick={() => void start()}>{busy ? "Requesting..." : operation ? "Retry meeting analysis" : "Start meeting analysis"}</button>}
        {results.length > 0 && !error && <button type="button" disabled={busy || activeIds.length === 0} onClick={newAnalysis}>New meeting analysis</button>}
        <button type="button" disabled={busy} onClick={() => setOpen(false)}>{allAccepted ? "Done" : "Close"}</button>
      </footer>
    </dialog>}
  </>;
}
