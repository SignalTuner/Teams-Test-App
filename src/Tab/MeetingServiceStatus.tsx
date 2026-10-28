import { useEffect, useId, useState } from "react";
import { ExternalWebsiteLink } from "./TeamsHostProvider";
import microsoftTeamsLogo from "./assets/microsoft-teams.png";
import "./MeetingServiceStatus.css";

type Incident = { IncidentId: number; Title: string; Status: string; PublishDate: string; Description: string; Link?: string };
export type ServiceDetails = { DisplayName: string; CurrentStatus: string; incidents: Incident[] };

// Presentation rules match the website's serviceIncidentPresentation.js.
export function incidentState(status: string) {
  switch (String(status || "").trim().toLowerCase()) {
    case "active": case "active incident": return "active";
    case "outage": return "outage";
    case "resolved": return "resolved";
    default: return "unknown";
  }
}
export function incidentPreview(incidents: Incident[]) {
  let resolved = 0;
  return [...incidents].sort((a, b) =>
    Number(incidentState(a.Status) === "resolved") - Number(incidentState(b.Status) === "resolved")
    || (Date.parse(b.PublishDate) || 0) - (Date.parse(a.PublishDate) || 0)
    || b.IncidentId - a.IncidentId)
    .filter(incident => incidentState(incident.Status) !== "resolved" || resolved++ < 5);
}
export function incidentReportLink(link?: string) {
  try {
    const url = new URL(link || "");
    return ["https:", "http:"].includes(url.protocol) ? url.href : null;
  } catch { return null; }
}
function incidentText(html: string) {
  // React also escapes the fallback used during server rendering.
  if (typeof DOMParser === "undefined") return html || "";
  const document = new DOMParser().parseFromString(html || "", "text/html");
  document.querySelectorAll("script, style").forEach(node => node.remove());
  document.querySelectorAll("br").forEach(node => node.replaceWith("\n"));
  document.querySelectorAll("p, div, li").forEach(node => node.append("\n"));
  return document.body.textContent?.trim() || "";
}

export async function loadTeamsServiceStatus(apiBaseUrl: string, token: string, signal: AbortSignal): Promise<ServiceDetails> {
  const read = async (path: string) => {
    const response = await fetch(`${apiBaseUrl.replace(/\/$/, "")}/api/User/${path}/3`, {
      signal, headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error("Service status unavailable");
    return response.json();
  };
  const [status, incidents] = await Promise.all([read("GetServiceStatus"), read("GetServiceIncidents")]);
  if (!status || typeof status.CurrentStatus !== "string" || !Array.isArray(incidents)) throw new Error("Invalid service status");
  return { DisplayName: status.DisplayName || "Microsoft Teams", CurrentStatus: status.CurrentStatus, incidents: incidentPreview(incidents) };
}

export default function MeetingServiceStatus({ apiBaseUrl, token }: { apiBaseUrl: string; token: string }) {
  const [details, setDetails] = useState<ServiceDetails | null>(null);
  const [error, setError] = useState("");
  useEffect(() => {
    setDetails(null); setError("");
    let cancelled = false, fetching = false;
    let controller: AbortController | undefined;
    async function refresh() {
      if (fetching || document.hidden) return;
      fetching = true; controller = new AbortController();
      const timeout = setTimeout(() => controller?.abort(), 15000);
      try {
        const value = await loadTeamsServiceStatus(apiBaseUrl, token, controller.signal);
        if (!cancelled) { setDetails(value); setError(""); }
      } catch {
        if (!cancelled) { setDetails(null); setError("This service's current status and incidents could not be loaded. Please retry shortly."); }
      } finally { clearTimeout(timeout); fetching = false; }
    }
    void refresh();
    const timer = setInterval(() => void refresh(), 30000);
    document.addEventListener("visibilitychange", refresh);
    return () => { cancelled = true; controller?.abort(); clearInterval(timer); document.removeEventListener("visibilitychange", refresh); };
  }, [apiBaseUrl, token]);
  return <MeetingServiceStatusView details={details} error={error} />;
}

export function MeetingServiceStatusView({ details, error = "" }: { details: ServiceDetails | null; error?: string }) {
  const [expanded, setExpanded] = useState(false);
  const bodyId = useId();
  const status = details?.CurrentStatus || "Status unavailable";
  const style = status === "Operational" ? "healthy" : status === "Outage" ? "outage" : status === "Active incident" ? "incident" : "unknown";
  const activeCount = details?.incidents.filter(incident => incidentState(incident.Status) !== "resolved").length || 0;
  const countStyle = !details ? "unknown" : !activeCount ? "healthy"
    : style === "outage" || details.incidents.some(incident => incidentState(incident.Status) === "outage") ? "outage" : "incident";
  const name = details?.DisplayName || "Microsoft Teams";
  return <section className="panel teamsMeetingService" data-service-status={style} aria-label="Microsoft Teams service status and incidents">
    <button className="teamsServiceRow" type="button" aria-expanded={expanded} aria-controls={bodyId}
      disabled={!details} onClick={() => setExpanded(value => !value)}>
      <span className="teamsServiceIdentity">
        <span className="teamsServiceLogo" aria-hidden="true"><img src={microsoftTeamsLogo} alt="" /></span>
        <span><span className="teamsServiceName">{name}</span><span className={`teamsServiceStatus teamsServiceStatus--${style}`}><span aria-hidden="true" />{!details && !error ? "Loading service status..." : status}</span></span>
      </span>
      <span className={`teamsServiceCount teamsServiceStatus teamsServiceStatus--${countStyle}`}>
        {details ? `${activeCount} active ${activeCount === 1 ? "incident" : "incidents"}` : "Incident count unavailable"}
      </span>
      <span className="teamsServiceToggle">
        {expanded ? "Hide incidents" : "View incidents"}<span className="teamsServiceDisclosure" aria-hidden="true" />
      </span>
    </button>
    {error && <p className="teamsServiceError" role="alert">{error}</p>}
    <div className="teamsServiceIncidents" id={bodyId} hidden={!expanded}>
      {details && <><h4>Incidents for {name}</h4>
        {!details.incidents.length && <p>No incidents to display.</p>}
        {details.incidents.map(incident => {
          const state = incidentState(incident.Status);
          const label = state === "unknown" ? incident.Status || "Unknown" : state[0].toUpperCase() + state.slice(1);
          const link = incidentReportLink(incident.Link);
          return <article key={incident.IncidentId} data-incident-state={state}>
            <div className="teamsServiceIncidentHeading"><h5>{incident.Title}</h5><span className="teamsServiceIncidentLabel">{label}</span></div>
            <p className="teamsServiceDate">Updated {new Date(incident.PublishDate).toLocaleString()}</p>
            <p className="teamsServiceDescription">{incidentText(incident.Description)}</p>
            {link && <ExternalWebsiteLink href={link}>Read provider report on website</ExternalWebsiteLink>}
          </article>;
        })}
        <ExternalWebsiteLink className="teamsServiceHistory" href="https://www.signaltuner.com/teamsstatus">See full report history on website</ExternalWebsiteLink>
      </>}
    </div>
  </section>;
}
