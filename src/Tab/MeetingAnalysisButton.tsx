import { useEffect, useRef, useState } from "react";
import { activeMeetingParticipantIds, readMeetingAnalysisOperation, requestMeetingAnalysis,
  type MeetingAnalysisOperation, type MeetingAnalysisResult } from "./meetingAnalysis";
import "./MeetingAnalysisButton.css";

type Props = {
  apiBaseUrl: string; token: string; userId: number; meetingSessionId: number; disabled?: boolean;
  participants: Array<{ userId: number; displayName: string | null; email: string | null; clientDataStatus: string }>;
  onAccepted: (results: MeetingAnalysisResult[]) => void;
  onOpenInTab?: () => void;
  openRequest?: string;
  onOpenRequestHandled?: () => void;
};

export default function MeetingAnalysisButton({ apiBaseUrl, token, userId, meetingSessionId, participants, disabled, onAccepted, onOpenInTab, openRequest, onOpenRequestHandled }: Props) {
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
  const [startedIds, setStartedIds] = useState<number[]>([]);
  const [credits, setCredits] = useState<number | null>(null);
  const [loadingCredits, setLoadingCredits] = useState(true);
  const [creditRefresh, setCreditRefresh] = useState(0);
  const dialog = useRef<HTMLDialogElement>(null);
  const inFlight = useRef(false);
  const mounted = useRef(true);
  const controller = useRef<AbortController | null>(null);
  const activeIds = activeMeetingParticipantIds(participants);
  const selectedIds = operation?.participantUserIds ?? targets.filter(id => activeIds.includes(id) && !startedIds.includes(id));
  const rosterIds = [...new Set([...participants.filter(p => Number.isSafeInteger(p.userId) && p.userId > 0).map(p => p.userId), ...selectedIds])];
  const handledOpenRequest = useRef<string | undefined>(undefined);
  useEffect(() => {
    if (!openRequest || onOpenInTab || disabled || handledOpenRequest.current === openRequest) return;
    handledOpenRequest.current = openRequest;
    if (!operation) { setTargets(activeIds); setStartedIds([]); setResults([]); setError(saved.error); }
    setOpen(true);
    onOpenRequestHandled?.();
  }, [openRequest, onOpenInTab, disabled, operation, activeIds, onOpenRequestHandled]);

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
    if (inFlight.current || disabled || saved.error || !token || !selectedIds.length || selectedIds.length > 1000) return;
    const next = operation || { requestId: crypto.randomUUID().replace(/-/g, ""), participantUserIds: selectedIds };
    // Persist before any request. Reload/retry must never create a second purchase.
    try { sessionStorage.setItem(storageKey, JSON.stringify(next)); }
    catch { setError("The request could not be saved in this tab. Allow tab storage before starting an analysis."); return; }
    setOperation(next);
    inFlight.current = true; setBusy(true); setError(""); setResults([]);
    const abort = new AbortController();
    controller.current = abort;
    const timeout = setTimeout(() => abort.abort(), 60000);
    try {
      const accepted = await requestMeetingAnalysis(apiBaseUrl, token, meetingSessionId, next, abort.signal);
      if (!mounted.current) return;
      setResults(accepted); onAccepted(accepted);
      // A complete response resolves the operation. Future explicit starts use
      // a fresh ID; only uncertain submissions retain the saved ID and roster.
      sessionStorage.removeItem(storageKey);
      setOperation(null);
      const newlyStarted = accepted.filter(result => result.request).map(result => result.targetUserId);
      setStartedIds(current => [...new Set([...current, ...newlyStarted])]);
      setTargets(current => current.filter(id => !newlyStarted.includes(id)));
      if (accepted.some(result => !result.request && result.statusCode === 429)) setOpen(true);
      else if (newlyStarted.length) setOpen(false);
      else setError("No participant analyses could be started. Check that the selected desktops are connected before starting again.");
    } catch (reason) {
      if (mounted.current) setError(reason instanceof Error && reason.name !== "AbortError" ? reason.message
        : "The response could not be confirmed. Retry this same meeting analysis; accepted requests keep running.");
    } finally {
      clearTimeout(timeout); inFlight.current = false;
      if (mounted.current) { setBusy(false); setCreditRefresh(value => value + 1); }
    }
  }

  const targetCount = selectedIds.length;
  const acceptedCount = results.filter(result => result.request).length;
  const rateLimited = results.filter(result => !result.request && result.statusCode === 429);
  return <>
    <div className="meetingAnalysisAction">
      <div className="meetingAnalysisCreditBalance" aria-live="polite" aria-atomic="true">
        Available credits: <strong>{loadingCredits ? "Loading..." : credits === null ? "Unavailable" : credits.toLocaleString()}</strong>
      </div>
      <button type="button" className="meetingAnalyzeButton" disabled={disabled || busy || (!operation && activeIds.length === 0)}
        aria-haspopup={onOpenInTab ? undefined : "dialog"} title={onOpenInTab ? "Open the meeting tab to confirm analysis" : undefined}
        onClick={() => { if (onOpenInTab) { onOpenInTab(); return; } if (!operation) { setTargets(activeIds); setStartedIds([]); setResults([]); setError(saved.error); } setOpen(true); }}>
        Analyze Meeting
      </button>
    </div>
    {!open && results.some(result => !result.request) && <p role="status" className="meetingAnalysisPartialResult">
      {acceptedCount} of {results.length} requests accepted. Open Analyze Meeting to select participants and start an analysis.
    </p>}
    {open && <dialog ref={dialog} className="remoteAnalysisDialog meetingAnalysisDialog"
      aria-labelledby="meeting-analysis-title" onCancel={event => { if (busy) event.preventDefault(); else setOpen(false); }} onClose={() => setOpen(false)}>
      <header><h2 id="meeting-analysis-title">Analyze Meeting</h2>
        <button type="button" disabled={busy} aria-label="Close meeting analysis" onClick={() => setOpen(false)}>Close</button>
      </header>
      <p>Connectivity Health Analysis runs for 5 minutes, starting when the user grants permission on the SignalTuner desktop client.</p>
      <p>Upon completion, a Connectivity Health Report will be emailed to you and the analyzed user.</p>
      <p>Each Connectivity Health Report requires 1 credit.</p>
      <p>Analysis disrupted by permission denial, cancellation, or disconnection will not deduct credits.</p>
      <p className="meetingAnalysisAvailable" aria-live="polite" aria-atomic="true">Available credits: <strong>{loadingCredits ? "Loading..." : credits ?? "Unavailable"}</strong></p>
      <fieldset className="meetingAnalysisParticipants" disabled={busy || Boolean(operation)}>
        <legend>Participants</legend>
        {rosterIds.map(id => {
          const participant = participants.find(p => p.userId === id);
          const unavailable = !activeIds.includes(id);
          const alreadyStarted = startedIds.includes(id);
          return <label key={id} className="meetingAnalysisParticipant">
            <span>{participant?.email || participant?.displayName || "Participant unavailable"}
              {alreadyStarted ? <small>Analysis requested</small> : unavailable && !operation && <small>Desktop not connected</small>}
            </span>
            <input type="checkbox" checked={selectedIds.includes(id)} disabled={(unavailable || alreadyStarted) && !operation}
              onChange={event => setTargets(current => event.target.checked ? [...new Set([...current, id])] : current.filter(target => target !== id))} />
          </label>;
        })}
        {!rosterIds.length && <p>No participants are available for analysis.</p>}
      </fieldset>
      <p className="meetingAnalysisRequired" aria-live="polite" aria-atomic="true">Credits required: <strong>{targetCount}</strong></p>
      {credits === 0 && <p>No report allowance is available. Test Analysis remains available from your personal dashboard on supported connected desktops.</p>}
      {targetCount > 1000 && <p role="alert">Meeting analysis supports up to 1000 active participants per run.</p>}
      {operation && <p>The previous submission is unconfirmed. Start meeting analysis will safely confirm that same selection before another analysis can be started.</p>}
      {error && <p role="alert" className="remoteAnalysisError">{error}</p>}
      {rateLimited.length > 0 && <>
        {acceptedCount > 0 && <p role="status">{acceptedCount} participant {acceptedCount === 1 ? "analysis was" : "analyses were"} requested. These participants are marked above and will not be submitted again.</p>}
        <div role="alert" className="remoteAnalysisError">
          The following users cannot be analyzed due to too many requests. Deselect them from the list above or try again later.
          <ul>{rateLimited.map(result => {
            const participant = participants.find(p => p.userId === result.targetUserId);
            return <li key={result.targetUserId}>{participant?.email || participant?.displayName || "Participant unavailable"}</li>;
          })}</ul>
        </div>
      </>}
      {busy && <p role="status">Requesting participant analyses...</p>}
      <footer>
        <button type="button" className="meetingAnalyzeButton" disabled={busy || disabled || Boolean(saved.error) || !targetCount || targetCount > 1000}
          onClick={() => void start()}>{busy ? "Requesting..." : "Start meeting analysis"}</button>
        <button type="button" disabled={busy} onClick={() => setOpen(false)}>Close</button>
      </footer>
    </dialog>}
  </>;
}
