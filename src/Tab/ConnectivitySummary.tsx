import { CompactSignalScoreDisplay, getSignalStatusLabel } from "./ConnectivityPresentation";

export default function ConnectivitySummary({ name, score, overallStatus, workspaceStatus, networkStatus, deviceStatus, hasData, meetingFallback = false }: {
  name: string; score: number | null; overallStatus: string | null; workspaceStatus: string | null;
  networkStatus: string | null; deviceStatus: string | null; hasData: boolean; meetingFallback?: boolean;
}) {
  return <dl className="participantMetrics" aria-label={`Signal status for ${name}`}>
    <div><dt>Signal Score</dt><dd><CompactSignalScoreDisplay score={score} overallStatus={overallStatus} hasData={hasData} /></dd></div>
    {[["Workspace", workspaceStatus], ["Network", networkStatus], ["Device", deviceStatus]].map(([label, status]) => {
      const text = getSignalStatusLabel(hasData || meetingFallback ? status : null, meetingFallback && hasData ? score : null);
      return <div key={label}><dt>{label}</dt><dd className={`telemetryStatusLabel telemetryStatus-${text.toLowerCase()}`}>{text}</dd></div>;
    })}
  </dl>;
}
