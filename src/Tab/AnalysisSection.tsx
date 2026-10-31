import { useId, type ReactNode } from "react";
import "./RemoteAnalysisButton.css";
import "./MeetingServiceStatus.css";

export default function AnalysisSection({ title, expanded, onToggle, summary, children }: {
  title: string; expanded: boolean; onToggle?: () => void; summary?: ReactNode; children: ReactNode;
}) {
  const id = useId();
  return <section className="remoteAnalysisInline" aria-label={title}>
    <button type="button" className="remoteAnalysisToggle" aria-expanded={expanded} aria-controls={id} onClick={onToggle}>
      <span>{title}</span><span className="teamsServiceDisclosure" aria-hidden="true" />
    </button>
    {summary && <div className="remoteAnalysisInlineSummary">{summary}</div>}
    <div id={id} className="remoteAnalysisInlineBody" hidden={!expanded}>{children}</div>
  </section>;
}
