import { useEffect, useMemo, useRef, useState } from "react";
import "./LiveAnalysisCharts.css";

// Match MainWindow.xaml.cs active-analysis windows, scales, and severity bands.
const colors = ["#3BB537", "#F9AE00", "#FC6F20", "#EC2F3E"];
const descending = (value: number, limits: number[]) => colors[limits.findIndex(limit => value >= limit)] || colors[3];
const ascending = (value: number, limits: number[]) => colors[limits.findIndex(limit => value < limit)] || colors[3];
export type MetricKey = "wifiStrength" | "downloadMbps" | "uploadMbps" | "packetLossPercent" | "latencyMs" | "cpuPercent" | "memoryPercent" | "jitterHistoryMs";
export type LiveTelemetry = {
  ageSeconds?: number; receivedAt?: string;
  snapshot?: Partial<Record<MetricKey, Array<number | null> | null>> & {
    elapsedSeconds?: number; jitterMs?: number | null; vpnStatus?: string | null;
  };
};
type Metric = { key: MetricKey; title: string; unit: string; count: number; max: number; bars?: boolean; color: (value: number) => string };
export const jitterColor = (value: number) => colors[value <= 15 ? 0 : value <= 30 ? 1 : value <= 50 ? 2 : 3];
export const liveMetrics: Metric[] = [
  { key: "wifiStrength", title: "Wi-Fi strength", unit: "%", count: 20, max: 100, color: v => descending(v, [61, 46, 31]) },
  { key: "downloadMbps", title: "Download speed", unit: "Mbps", count: 25, max: 4, bars: true, color: v => descending(v, [3, 1.8, 0.8]) },
  { key: "uploadMbps", title: "Upload speed", unit: "Mbps", count: 25, max: 3.8, bars: true, color: v => descending(v, [3.8, 2.6, 1.2]) },
  { key: "packetLossPercent", title: "Packet loss", unit: "%", count: 10, max: 100, color: v => ascending(v, [1.900001, 3.900001, 5.900001]) },
  { key: "latencyMs", title: "Latency", unit: "ms", count: 10, max: 100, color: v => ascending(Math.ceil(v), [21, 51, 100]) },
  { key: "jitterHistoryMs", title: "Jitter", unit: "ms", count: 10, max: 100, color: jitterColor },
  { key: "cpuPercent", title: "CPU usage", unit: "%", count: 20, max: 100, color: v => ascending(v, [80, 90, 95]) },
  { key: "memoryPercent", title: "Memory usage", unit: "%", count: 20, max: 100, color: v => ascending(v, [93, 96, 98]) },
];
const finite = (value: unknown): value is number => typeof value === "number" && Number.isFinite(value) && value >= 0;
const format = (value: unknown, unit: string) => finite(value) ? `${Number(value.toFixed(2))}${unit === "%" ? "" : " "}${unit}` : "Unavailable";

export default function LiveAnalysisCharts({ telemetry, running, requestState, idleMessage = "Click Analyze to begin a 5-minute rapid sampling session. See live connection health data and receive a detailed PDF report." }: {
  telemetry?: LiveTelemetry | null; running: boolean; requestState?: string; idleMessage?: string;
}) {
  const receipt = useMemo(() => ({ time: Date.now(), telemetry }), [telemetry]);
  const [now, setNow] = useState(Date.now);
  useEffect(() => {
    if (!running) return;
    const timer = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(timer);
  }, [running]);
  const snapshot = telemetry?.snapshot;
  // Older producers supply only the scalar. Show that one real point without
  // fabricating a history from repeated polls or latency timeout sentinels.
  const jitterSamples = Array.isArray(snapshot?.jitterHistoryMs) ? snapshot.jitterHistoryMs
    : finite(snapshot?.jitterMs) ? [snapshot.jitterMs] : [];
  const age = Math.max(0, Number(telemetry?.ageSeconds) || 0) + Math.max(0, now - receipt.time) / 1000;
  const stale = snapshot && running && age > 8;
  const elapsed = Math.min(300, Math.max(0, Number(snapshot?.elapsedSeconds) || 0));
  return <section className="live-analysis" aria-label="Live analysis charts">
    <div className="live-analysis-heading"><h3>Analysis telemetry</h3><span>{snapshot ? `${elapsed}s / 300s sampled` : "Waiting for samples"}</span></div>
    <p className={stale ? "live-analysis-stale" : "live-analysis-status"}>
      {!snapshot ? running || ["Pending", "Approved"].includes(requestState || "")
        ? "Waiting for live samples. Live charts require the updated SignalTuner desktop app."
        : requestState ? "No samples were received for this analysis." : idleMessage
        : !running ? "Last reported samples. Analysis is no longer running."
        : stale ? `Live data is delayed. Last update ${Math.floor(age)} seconds ago.` : "Live updates every two seconds"}
    </p>
    <div className="live-analysis-grid">
      {liveMetrics.slice(0, 6).map(metric => <MetricChart key={metric.key} metric={metric}
        samples={metric.key === "jitterHistoryMs" ? jitterSamples : snapshot?.[metric.key]} />)}
      <SummaryCard title="VPN" value={snapshot?.vpnStatus === "Detected" ? "VPN Detected" : snapshot?.vpnStatus === "NotDetected" ? "No VPN Detected" : "Detection Unavailable"}
        color={snapshot?.vpnStatus === "Detected" ? colors[1] : undefined}
        description="Local network indicators identify VPN-like connections. Some VPNs may not be detectable, so no detection does not guarantee that no VPN is active." />
      {liveMetrics.slice(6).map(metric => <MetricChart key={metric.key} metric={metric} samples={snapshot?.[metric.key]} />)}
    </div>
  </section>;
}

function SummaryCard({ title, value, color, description }: { title: string; value: string; color?: string; description: string }) {
  return <div className="live-metric live-metric--summary" role="group" aria-label={title} title={description}>
    <div className="live-metric-heading"><h4>{title}</h4><strong style={{ color: color || "var(--muted)" }}>{value}</strong></div>
  </div>;
}

function MetricChart({ metric, samples }: { metric: Metric; samples?: Array<number | null> | null }) {
  const chart = useRef<SVGSVGElement>(null);
  const [width, setWidth] = useState(324);
  useEffect(() => {
    const element = chart.current;
    if (!element) return;
    // Match the SVG coordinate system to its card so a wide card has a wide plot,
    // without stretching text/points or adding aspect-ratio letterboxing.
    const measure = (value: number) => { if (Number.isFinite(value) && value > 0) setWidth(value); };
    measure(element.getBoundingClientRect().width);
    if (typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(entries => measure(entries[0]?.contentRect.width));
    observer.observe(element);
    return () => observer.disconnect();
  }, []);
  const values = Array.isArray(samples) ? samples.slice(-metric.count).map(v => finite(v) ? v : null) : [];
  const latest = values.length ? values[values.length - 1] : null;
  const maximum = metric.unit === "ms" ? Math.min(1000, Math.max(100, ...values.filter(finite))) : metric.max;
  const right = Math.max(35, width - 14);
  const x = (i: number) => 34 + (metric.count - values.length + i) * (right - 34) / (metric.count - 1);
  const y = (v: number) => 38 - Math.min(maximum, v) / maximum * 30;
  const hasSamples = values.some(finite);
  return <div className={`live-metric live-metric--chart live-metric--${metric.key}`} role="group" aria-label={metric.title}>
    <div className="live-metric-heading"><h4>{metric.title}</h4><strong style={{ color: finite(latest) ? metric.color(latest) : undefined }}>{format(latest, metric.unit)}</strong></div>
    <svg ref={chart} viewBox={`0 0 ${width} 59`} role="img" aria-label={`${metric.title} chart: ${format(latest, metric.unit)}, latest ${metric.count} samples`}>
      <line x1="34" y1="8" x2={right} y2="8" className="live-chart-grid" />
      <line x1="34" y1="23" x2={right} y2="23" className="live-chart-grid" />
      <line x1="34" y1="38" x2={right} y2="38" className="live-chart-grid" />
      <text x="29" y="11" textAnchor="end">{maximum}</text><text x="29" y="41" textAnchor="end">0</text>
      <text x="34" y="54">Recent samples</text><text x={right} y="54" textAnchor="end">Latest</text>
      {!hasSamples && <text x={(34 + right) / 2} y="27" textAnchor="middle">Waiting for samples</text>}
      {values.map((v, i) => !finite(v) ? null : metric.bars
        ? <rect key={i} x={x(i) - 3} y={y(v)} width="6" height={Math.max(1, 38 - y(v))} rx="1" fill={metric.color(v)}><title>{format(v, metric.unit)}</title></rect>
        : <g key={i}>{i > 0 && finite(values[i - 1]) && <line x1={x(i - 1)} y1={y(values[i - 1]!)} x2={x(i)} y2={y(v)} stroke={metric.color(v)} strokeWidth="2" />}
          <circle cx={x(i)} cy={y(v)} r="2.5" fill={metric.color(v)}><title>{format(v, metric.unit)}</title></circle></g>)}
    </svg>
  </div>;
}
