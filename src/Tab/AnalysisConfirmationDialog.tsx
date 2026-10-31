import { useEffect, useRef, type ReactNode } from "react";
import "./RemoteAnalysisButton.css";
import "./MeetingAnalysisButton.css";

export default function AnalysisConfirmationDialog({ open, name, credits, loadingCredits = false, busy, error, onClose, children }: {
  open: boolean; name: string; credits: number | null; loadingCredits?: boolean;
  busy: boolean; error: string; onClose: () => void; children: ReactNode;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  useEffect(() => { if (open && !dialog.current?.open) dialog.current?.showModal(); }, [open]);
  if (!open) return null;
  return <dialog ref={dialog} className="remoteAnalysisDialog meetingAnalysisDialog" aria-label={`Analyze ${name}`}
    onCancel={event => { if (busy) event.preventDefault(); else onClose(); }} onClose={onClose}>
    <header><h2>Analyze {name}</h2><button type="button" disabled={busy} aria-label="Close analysis request" onClick={onClose}>×</button></header>
    <p>Connectivity Health Analysis runs for 5 minutes, starting when the user grants permission on the SignalTuner desktop client.</p>
    <p>Upon completion, a Connectivity Health Report will be emailed to you and the analyzed user.</p>
    <p>Each Connectivity Health Report requires 1 credit.</p>
    <p>Analysis disrupted by permission denial, cancellation, or disconnection will not deduct credits.</p>
    <p className="meetingAnalysisAvailable" aria-live="polite" aria-atomic="true">Available credits: <strong>{loadingCredits ? "Loading..." : credits ?? "Unavailable"}</strong></p>
    {error && <p role="alert" className="remoteAnalysisError">{error}</p>}
    {busy && <p role="status">Updating analysis...</p>}
    {children}
  </dialog>;
}
