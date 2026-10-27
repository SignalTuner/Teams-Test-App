import { useId, type ReactNode } from "react";
import "./RemoteAnalysisButton.css";
import "./MeetingServiceStatus.css";

export default function AnalysisSection({ title, expanded, onToggle, children }: {
  title: string; expanded: boolean; onToggle?: () => void; children: ReactNode;
}) {
  const id = useId();
  return <section className="remoteAnalysisInline" aria-label={title}>
    <button type="button" className="remoteAnalysisToggle" aria-expanded={expanded} aria-controls={id} onClick={onToggle}>
      <span>{title}</span><span className="teamsServiceDisclosure" aria-hidden="true" />
    </button>
    <div id={id} className="remoteAnalysisInlineBody" hidden={!expanded}>{children}</div>
  </section>;
}
