import type { Issue, TelemetryRecord, SignalScoreTrendResponse } from "./dashboardTypes";
import deviceCoresIcon from "./assets/device-cores-icon.svg";
import deviceCpuIcon from "./assets/device-cpu-icon.svg";
import deviceMemoryIcon from "./assets/device-memory-icon.svg";
import deviceProcessorIcon from "./assets/device-processor-icon.svg";
import networkDownloadIcon from "./assets/network-download-icon.svg";
import networkLatencyIcon from "./assets/network-latency-icon.svg";
import networkPacketLossIcon from "./assets/network-packet-loss-icon.svg";
import networkUploadIcon from "./assets/network-upload-icon.svg";
import signalRecommendationIcon from "./assets/signal-reccommendation-icon.svg";
import workspaceCurrentNetworkIcon from "./assets/workspace-current-network-icon.svg";
import workspaceNetworkFrequencyIcon from "./assets/workspace-network-frequency-icon.svg";
import workspaceVpnIcon from "./assets/workspace-vpn-icon.svg";
import workspaceWifiStrengthIcon from "./assets/workspace-wifi-strength-icon.svg";
type AnalysisStatus = string;
type SignalDisplayStatus = "Healthy" | "Fair" | "Poor" | "Critical" | "Offline" | "Unknown";
function formatValue(value: unknown): string {
  if (value === undefined || value === null || value === "") {
    return "No data";
  }

  if (typeof value === "boolean") {
    return value ? "Yes" : "No";
  }

  return String(value);
}

function formatTrendTooltip(timestampUtc: string, score: number): string {
  const timestamp = new Date(timestampUtc);
  const timestampLabel = Number.isNaN(timestamp.getTime())
    ? timestampUtc
    : timestamp.toLocaleString([], { hour: "numeric", minute: "2-digit", second: "2-digit" });

  return `${timestampLabel}: ${Math.round(score)} Signal Score`;
}

export function getSignalDisplayStatusFromScore(score: number | null | undefined): "Healthy" | "Fair" | "Poor" | "Critical" | "Unknown" {
  if (score === null || score === undefined || Number.isNaN(score)) {
    return "Unknown";
  }

  if (score >= 90) {
    return "Healthy";
  }

  if (score >= 80) {
    return "Fair";
  }

  if (score >= 70) {
    return "Poor";
  }

  return "Critical";
}

export function getSignalStatusLabel(status: AnalysisStatus | null | undefined, score: number | null | undefined): "Healthy" | "Fair" | "Poor" | "Critical" | "Unknown" {
  const normalizedStatus = status?.trim().toLowerCase();

  if ((normalizedStatus === "healthy" || normalizedStatus === "excellent")) {
    return "Healthy";
  }

  if (normalizedStatus === "fair") {
    return "Fair";
  }

  if (normalizedStatus === "poor") {
    return "Poor";
  }

  if (normalizedStatus === "critical") {
    return "Critical";
  }

  if (normalizedStatus === "unknown" || normalizedStatus === "disconnected" || normalizedStatus === "offline") return "Unknown";
  return getSignalDisplayStatusFromScore(score);
}

function getSignalOverallStatusLabel(status: AnalysisStatus | null | undefined, score: number | null | undefined): SignalDisplayStatus {
  const normalizedStatus = status?.trim().toLowerCase();

  if ((normalizedStatus === "healthy" || normalizedStatus === "excellent")) {
    return "Healthy";
  }

  if (normalizedStatus === "fair") {
    return "Fair";
  }

  if (normalizedStatus === "poor") {
    return "Poor";
  }

  if (normalizedStatus === "critical") {
    return "Critical";
  }

  if (normalizedStatus === "offline") {
    return "Offline";
  }

  return getSignalStatusLabel(null, score);
}

function getSignalStatusColor(status: SignalDisplayStatus): string {
  switch (status) {
    case "Healthy":
      return "#3BB537";
    case "Fair":
      return "#F9AE00";
    case "Poor":
      return "#FC6F20";
    case "Critical":
      return "#EC2F3E";
    case "Offline":
      return "#61708A";
    default:
      return "#6B96C1";
  }
}

function getTelemetryValue(telemetry: TelemetryRecord | null, keys: string[], fallback = "No data"): string {
  if (!telemetry) {
    return fallback;
  }

  for (const key of keys) {
    const value = telemetry[key];

    if (value !== undefined && value !== null && value !== "") {
      return formatValue(value);
    }
  }

  return fallback;
}

function getTelemetryRawValue(telemetry: TelemetryRecord | null, keys: string[]): string | number | boolean | null {
  if (!telemetry) {
    return null;
  }

  for (const key of keys) {
    const value = telemetry[key];

    if (value !== undefined && value !== null && value !== "") {
      return value;
    }
  }

  return null;
}

function getTelemetryNumber(telemetry: TelemetryRecord | null, keys: string[]): number | null {
  const value = getTelemetryRawValue(telemetry, keys);

  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string") {
    const parsed = Number.parseFloat(value.replace(/,/g, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }

  return null;
}

function formatTelemetryMeasurement(telemetry: TelemetryRecord | null, keys: string[], unit = ""): string {
  const value = getTelemetryRawValue(telemetry, keys);

  if (value === null) {
    return "No data";
  }

  if (typeof value === "number") {
    return `${Number.isInteger(value) ? value : value.toFixed(1)}${unit}`;
  }

  if (typeof value === "boolean") {
    return value ? "Yes" : "No";
  }

  const text = value.trim();
  const isNumericText = text !== "" && Number.isFinite(Number.parseFloat(text.replace(/,/g, "")));
  return isNumericText && unit && !text.toLowerCase().includes(unit.trim().toLowerCase()) ? `${text}${unit}` : text;
}

function truncateTelemetryValue(value: string, maxLength = 12): string {
  return value.length > maxLength ? `${value.slice(0, maxLength - 1)}...` : value;
}

function getTelemetryValueTitle(value: string): string | undefined {
  return value.length > 12 ? value : undefined;
}

const signalMetricColors = {
  unknown: "#6b96c1",
  healthy: "#3BB537",
  fair: "#F9AE00",
  poor: "#FC6F20",
  critical: "#EC2F3E",
} as const;

type TelemetryInsight = {
  title: "Device" | "Workspace" | "Network";
  text: string;
  severity: "Informational" | "Fair" | "Poor" | "Critical";
};

type TelemetryRecommendation = {
  text: string;
  severity: TelemetryInsight["severity"];
  priority: number;
};

function getTelemetryInsightColor(severity: TelemetryInsight["severity"]): string {
  switch (severity) {
    case "Critical":
      return signalMetricColors.critical;
    case "Poor":
      return signalMetricColors.poor;
    case "Fair":
      return signalMetricColors.fair;
    default:
      return signalMetricColors.healthy;
  }
}

function buildTelemetryInsights(telemetry: TelemetryRecord | null): {
  overall: string;
  categories: TelemetryInsight[];
  recommendations: TelemetryRecommendation[];
} {
  const cpu = getTelemetryNumber(telemetry, ["cpu", "cpuUsage", "signal_cpu", "cpu_percent", "SignalCPU"]);
  const memory = getTelemetryNumber(telemetry, ["memory", "memoryUsage", "signal_memory", "memory_percent", "SignalMemory"]);
  const wifiStrength = getTelemetryNumber(telemetry, ["wifiStrength", "wiFiStrength", "signal_wifi_strength", "wifi_strength", "SignalWifiStrength"]);
  const wifiStatus = getTelemetryValue(telemetry, ["wifiStatus", "workspaceStatus", "signal_wifi_status", "SignalWifiStatus"]);
  const wifiColor = getWifiStrengthMetricColor(wifiStrength, wifiStatus);
  const download = getTelemetryNumber(telemetry, ["downloadSpeed", "signal_download_speed", "download_speed", "SignalDownloadSpeed"]);
  const upload = getTelemetryNumber(telemetry, ["uploadSpeed", "signal_upload_speed", "upload_speed", "SignalUploadSpeed"]);
  const latency = getTelemetryNumber(telemetry, ["latency", "ping", "signal_latency", "latencyMs", "SignalPing"]);
  const packetLoss = getTelemetryNumber(telemetry, ["packetLoss", "signal_packet_loss", "packet_loss", "SignalPacketLoss"]);
  const vpn = getTelemetryValue(telemetry, ["vpn", "vpnStatus", "vpnDetected", "signal_vpn", "vpn_status", "signal_vpn_detected", "SignalVpnDetected"]);
  const categories: TelemetryInsight[] = [];
  const recommendations: TelemetryRecommendation[] = [];

  if (cpu !== null && memory !== null && cpu >= 85 && memory >= 85) {
    categories.push({ title: "Device", text: "Your device is under heavy CPU and memory load, which may cause degraded audio/video performance or application responsiveness.", severity: cpu >= 95 || memory >= 95 ? "Critical" : "Poor" });
    recommendations.push({ text: "Close unnecessary or resource-intensive applications to free system resources.", severity: "Poor", priority: 96 });
  } else if (cpu !== null && cpu >= 95) {
    categories.push({ title: "Device", text: "Your CPU is nearly fully utilized and may not have enough processing capacity available for reliable audio or video processing.", severity: "Critical" });
    recommendations.push({ text: "Close CPU-intensive applications before joining or continuing your meeting.", severity: "Critical", priority: 88 });
  } else if (memory !== null && memory >= 95) {
    categories.push({ title: "Device", text: "Available memory is critically low. Windows may begin relying heavily on virtual memory, which can reduce meeting performance.", severity: "Critical" });
    recommendations.push({ text: "Close memory-intensive applications to free RAM.", severity: "Critical", priority: 86 });
  } else if ((cpu !== null && cpu >= 85) || (memory !== null && memory >= 85)) {
    categories.push({ title: "Device", text: cpu !== null && cpu >= 85 ? "CPU utilization is high and may affect meeting performance if demand increases further." : "Memory usage is high and could reduce application responsiveness during a meeting.", severity: "Poor" });
    recommendations.push({ text: cpu !== null && cpu >= 85 ? "Close unnecessary applications to free processing capacity." : "Close unused applications or browser tabs.", severity: "Poor", priority: 60 });
  } else {
    categories.push({ title: "Device", text: "Your device currently has ample processing and memory capacity for a video meeting.", severity: "Informational" });
  }

  if (wifiColor === signalMetricColors.critical) {
    categories.push({ title: "Workspace", text: "Your Wi-Fi signal is critically weak.", severity: "Critical" });
    recommendations.push({ text: "Move closer to your Wi-Fi access point or improve Wi-Fi coverage in this workspace.", severity: "Poor", priority: 88 });
  } else if (wifiColor === signalMetricColors.poor) {
    categories.push({ title: "Workspace", text: "Your Wi-Fi signal is weak. Distance or obstructions between your device and access point may be reducing connection stability.", severity: "Poor" });
    recommendations.push({ text: "Move closer to the access point or switch to a more reliable Wi-Fi band or access point.", severity: "Poor", priority: 84 });
  } else if (wifiColor === signalMetricColors.unknown) {
    categories.push({ title: "Workspace", text: "Wi-Fi signal measurement is unavailable.", severity: "Informational" });
  } else if (wifiColor === signalMetricColors.fair) {
    categories.push({ title: "Workspace", text: "Your Wi-Fi signal is below the healthy range.", severity: "Fair" });
  } else {
    categories.push({ title: "Workspace", text: "Your workspace connection currently provides suitable Wi-Fi coverage for a video meeting.", severity: "Informational" });
  }

  const networkValues = [download, upload, latency, packetLoss];
  if (networkValues.some((value) => value !== null)) {
    const criticalPacketLoss = packetLoss !== null && packetLoss > 5.9;
    const criticalLatency = latency !== null && latency > 99;
    const poorBandwidth = (download !== null && download < 0.8) || (upload !== null && upload < 1.2);
    if (criticalPacketLoss && wifiColor === signalMetricColors.healthy) {
      categories.push({ title: "Network", text: "Your Wi-Fi signal is strong, so wireless coverage is unlikely to be the primary source of your connection problem.", severity: "Critical" });
      recommendations.push({ text: "Check router or ISP performance and temporarily test without the VPN if one is connected.", severity: "Critical", priority: 90 });
    } else if (criticalLatency && vpn === "Connected") {
      categories.push({ title: "Network", text: "Your connection is experiencing critically high latency while routed through a VPN.", severity: "Critical" });
      recommendations.push({ text: "If permitted by your organization, temporarily disconnect the VPN and retest.", severity: "Critical", priority: 94 });
    } else if (criticalPacketLoss || criticalLatency) {
      categories.push({ title: "Network", text: "Your connection is experiencing broader instability that may affect real-time audio and video communication.", severity: "Critical" });
      recommendations.push({ text: "Reduce network congestion or switch to a more stable connection.", severity: "Critical", priority: 96 });
    } else if (poorBandwidth) {
      categories.push({ title: "Network", text: "Your connection is experiencing a bandwidth bottleneck that may affect meeting quality.", severity: "Poor" });
      recommendations.push({ text: "Reduce other network activity or switch to a faster internet connection.", severity: "Poor", priority: 90 });
    } else {
      categories.push({ title: "Network", text: vpn === "Connected" ? "Your network is performing well even with the VPN connected." : "Your internet connection is currently well suited for real-time audio and video communication.", severity: "Informational" });
    }
  } else {
    categories.push({ title: "Network", text: "Network telemetry is not available for this analysis window.", severity: "Informational" });
  }

  const rankedRecommendations = recommendations
    .sort((first, second) => second.priority - first.priority)
    .filter((recommendation, index, all) => all.findIndex((candidate) => candidate.text === recommendation.text) === index)
    .slice(0, 3);
  const overall = rankedRecommendations.length > 0
    ? rankedRecommendations.map((recommendation) => recommendation.text).join(" ")
    : "No immediate recommendations. Your current workspace, network, and device appear healthy.";
  return { overall, categories, recommendations: rankedRecommendations };
}

function getCpuMetricColor(value: number | null): string {
  if (value === null || value === 0) {
    return signalMetricColors.unknown;
  }

  if (value >= 95) {
    return signalMetricColors.critical;
  }

  if (value >= 85) {
    return signalMetricColors.poor;
  }

  if (value >= 70) {
    return signalMetricColors.fair;
  }

  return signalMetricColors.healthy;
}

function getMemoryMetricColor(value: number | null): string {
  if (value === null || value === 0) {
    return signalMetricColors.unknown;
  }

  if (value >= 98) {
    return signalMetricColors.critical;
  }

  if (value >= 96) {
    return signalMetricColors.poor;
  }

  if (value >= 93) {
    return signalMetricColors.fair;
  }

  return signalMetricColors.healthy;
}

function getWifiStrengthMetricColor(value: number | null, status?: string): string {
  const normalized = status?.trim().toLowerCase();
  if (normalized === "excellent" || normalized === "healthy") return signalMetricColors.healthy;
  if (normalized === "fair") return signalMetricColors.fair;
  if (normalized === "poor") return signalMetricColors.poor;
  if (normalized === "critical") return signalMetricColors.critical;
  if (normalized === "unknown" || normalized === "disconnected" || normalized === "offline" || value === null) return signalMetricColors.unknown;
  if (value > 60) return signalMetricColors.healthy;
  if (value > 45) return signalMetricColors.fair;
  if (value > 30) return signalMetricColors.poor;
  return signalMetricColors.critical;
}

function getWifiBandMetricColor(value: string, wifiStrength: number | null, wifiStatus?: string): string {
  if (value === "5 GHz" || value === "6 GHz") {
    return getWifiStrengthMetricColor(wifiStrength, wifiStatus) === signalMetricColors.healthy ? signalMetricColors.healthy : signalMetricColors.fair;
  }

  if (value === "2.4 GHz") {
    return signalMetricColors.healthy;
  }

  return signalMetricColors.unknown;
}

function getVpnMetricColor(value: string): string {
  if (value === "Not Connected") {
    return signalMetricColors.healthy;
  }

  if (value === "Connected") {
    return signalMetricColors.fair;
  }

  return signalMetricColors.unknown;
}

function getDownloadMetricColor(value: number | null): string {
  if (value === null || value === 0) {
    return signalMetricColors.unknown;
  }

  if (value >= 3) {
    return signalMetricColors.healthy;
  }

  if (value >= 1.8) {
    return signalMetricColors.fair;
  }

  if (value >= 0.8) {
    return signalMetricColors.poor;
  }

  return signalMetricColors.critical;
}

function getUploadMetricColor(value: number | null): string {
  if (value === null || value === 0) {
    return signalMetricColors.unknown;
  }

  if (value >= 3.8) {
    return signalMetricColors.healthy;
  }

  if (value >= 2.6) {
    return signalMetricColors.fair;
  }

  if (value >= 1.2) {
    return signalMetricColors.poor;
  }

  return signalMetricColors.critical;
}

function getLatencyMetricColor(value: number | null): string {
  if (value === null) {
    return signalMetricColors.unknown;
  }

  if (value <= 20) {
    return signalMetricColors.healthy;
  }

  if (value <= 50) {
    return signalMetricColors.fair;
  }

  if (value <= 99) {
    return signalMetricColors.poor;
  }

  return signalMetricColors.critical;
}

function getPacketLossMetricColor(value: number | null): string {
  if (value === null) {
    return signalMetricColors.unknown;
  }

  if (value <= 1.9) {
    return signalMetricColors.healthy;
  }

  if (value <= 3.9) {
    return signalMetricColors.fair;
  }

  if (value <= 5.9) {
    return signalMetricColors.poor;
  }

  if (value > 6) {
    return signalMetricColors.critical;
  }

  return signalMetricColors.unknown;
}

function signalScoreArc(sweep: number): string {
  const point = (angle: number) => {
    const radians = angle * Math.PI / 180;
    return `${47.3 + 34.6 * Math.cos(radians)} ${51 + 34.6 * Math.sin(radians)}`;
  };
  return `M ${point(145)} A 34.6 34.6 0 ${sweep > 180 ? 1 : 0} 1 ${point(145 + sweep)}`;
}

export function CompactSignalScoreDisplay({
  score,
  overallStatus,
  hasData,
  accessibleLabel = "Signal Score",
}: {
  score: number | null;
  overallStatus: AnalysisStatus | null;
  hasData: boolean;
  accessibleLabel?: string;
}) {
  const scoreIsAvailable = hasData && score !== null && score !== undefined && Number.isFinite(score);
  const displayScore = scoreIsAvailable ? Math.round(score) : null;
  const clampedScore = displayScore === null ? 0 : Math.max(0, Math.min(100, displayScore));
  const displayStatus = getSignalOverallStatusLabel(overallStatus, scoreIsAvailable ? score : null);

  return (
    <div
      className={`compactSignalScoreDisplay compactSignalScore-${displayStatus.toLowerCase()}`}
      role="img"
      aria-label={
        scoreIsAvailable
          ? `${accessibleLabel} ${displayScore} percent, overall status ${displayStatus}`
          : `${accessibleLabel} unavailable, overall status ${displayStatus}`
      }
    >
      <svg className="compactSignalScoreMeter" viewBox="0 0 95 82" aria-hidden="true" focusable="false">
        <path
          className="compactSignalScoreTrack"
          d={signalScoreArc(250)}
        />
        {scoreIsAvailable && clampedScore > 0 && (
          <path
            className="compactSignalScoreArc"
            d={signalScoreArc(250 * clampedScore / 100)}
          />
        )}
        <text className="compactSignalScoreValue" x="47.3" y="49" textAnchor="middle">
          {displayScore === null ? "-" : `${displayScore}%`}
        </text>
        <text className="compactSignalScoreStatus" x="47.3" y="63" textAnchor="middle">
          {displayStatus}
        </text>
      </svg>
    </div>
  );
}

function SignalScoreTrendChart({
  trend,
  isLoading,
  error,
}: {
  trend: SignalScoreTrendResponse | null;
  isLoading: boolean;
  error: string | null;
}) {
  const points = trend?.points.slice(0, 10) ?? [];
  const hasCompleteTrend = points.length === 10;

  const width = 320;
  const height = 142;
  const left = 34;
  const right = 10;
  const top = 12;
  const bottom = 28;
  const plotWidth = width - left - right;
  const plotHeight = height - top - bottom;
  const coordinates = hasCompleteTrend ? points.map((point, index) => {
    const score = Math.max(0, Math.min(100, point.averageScore));
    const color = getSignalStatusColor(getSignalDisplayStatusFromScore(score));

    return {
      x: left + (plotWidth / (points.length - 1)) * index,
      y: top + ((100 - score) / 100) * plotHeight,
      color,
      score,
      timestampUtc: point.timestampUtc,
    };
  }) : [];
  const polylinePoints = coordinates.map((point) => `${point.x},${point.y}`).join(" ");
  const areaPoints = `${left},${top + plotHeight} ${polylinePoints} ${left + plotWidth},${top + plotHeight}`;


  return (
    <div className="signalScoreTrend" aria-label="Last 10 Signal Scores">
      <div className="signalScoreTrendHeader">
        <p>Last 10 Signal Scores</p>
      </div>
      <svg className="signalScoreTrendSvg" role="img" viewBox={`0 0 ${width} ${height}`}>
        <title>Last 10 Signal Scores</title>
        {[0, 50, 100].map((tick) => {
          const y = top + ((100 - tick) / 100) * plotHeight;

          return (
            <g key={tick}>
              <line className="trendGridLine" x1={left} x2={width - right} y1={y} y2={y} />
              <text className="trendYAxisLabel" x={left - 8} y={y + 4}>
                {tick}
              </text>
            </g>
          );
        })}
        {hasCompleteTrend ? <polygon className="trendArea" points={areaPoints} /> : null}
        {hasCompleteTrend ? coordinates.slice(1).map((point, index) => {
          const previousPoint = coordinates[index];

          return (
            <line
              className="trendSegment"
              key={`${previousPoint.timestampUtc}-${point.timestampUtc}`}
              stroke={point.color}
              x1={previousPoint.x}
              x2={point.x}
              y1={previousPoint.y}
              y2={point.y}
            >
              <title>{formatTrendTooltip(point.timestampUtc, point.score)}</title>
            </line>
          );
        }) : null}
        {hasCompleteTrend ? coordinates.map((point) => (
          <circle className="trendPoint" cx={point.x} cy={point.y} fill={point.color} key={point.timestampUtc} r="3.4">
            <title>{formatTrendTooltip(point.timestampUtc, point.score)}</title>
          </circle>
        )) : null}
        {!hasCompleteTrend ? (
          <text className="trendNoDataLabel" textAnchor="middle" x={left + plotWidth / 2} y={top + plotHeight / 2 + 4}>
            No data
          </text>
        ) : null}
      </svg>
      {isLoading && !trend ? <p className="signalScoreTrendState">Loading Signal Score trend.</p> : null}
      {error ? <p className="signalScoreTrendState">{error}</p> : null}
    </div>
  );
}

export function ParticipantTelemetryDetail({
  hasActiveAnalysisSession,
  hasData,
  issues,
  signalScoreTrend,
  signalScoreTrendError,
  signalScoreTrendLoading,
  telemetry,
}: {
  hasActiveAnalysisSession: boolean;
  hasData: boolean;
  issues: Issue[];
  signalScoreTrend: SignalScoreTrendResponse | null;
  signalScoreTrendError: string | null;
  signalScoreTrendLoading: boolean;
  telemetry: TelemetryRecord | null;
}) {
  const recommendation = hasActiveAnalysisSession
    ? ""
    : issues[0]?.recommendation ?? (hasData ? "" : "Prompt the participant to activate the local client.");
  const cpu = getTelemetryNumber(telemetry, ["cpu", "cpuUsage", "signal_cpu", "cpu_percent", "SignalCPU"]);
  const memory = getTelemetryNumber(telemetry, ["memory", "memoryUsage", "signal_memory", "memory_percent", "SignalMemory"]);
  const wifiStrength = getTelemetryNumber(telemetry, ["wifiStrength", "wiFiStrength", "signal_wifi_strength", "wifi_strength", "SignalWifiStrength"]);
  const wifiStatus = getTelemetryValue(telemetry, ["wifiStatus", "workspaceStatus", "signal_wifi_status", "SignalWifiStatus"]);
  const wifiColor = getWifiStrengthMetricColor(wifiStrength, wifiStatus);
  const download = getTelemetryNumber(telemetry, ["downloadSpeed", "signal_download_speed", "download_speed", "SignalDownloadSpeed"]);
  const upload = getTelemetryNumber(telemetry, ["uploadSpeed", "signal_upload_speed", "upload_speed", "SignalUploadSpeed"]);
  const latency = getTelemetryNumber(telemetry, ["latency", "ping", "signal_latency", "latencyMs", "SignalPing"]);
  const packetLoss = getTelemetryNumber(telemetry, ["packetLoss", "signal_packet_loss", "packet_loss", "SignalPacketLoss"]);
  const currentNetwork = getTelemetryValue(telemetry, ["currentNetwork", "signal_current_network", "current_network", "SignalCurrentNetwork"]);
  const frequency = getTelemetryValue(telemetry, ["wifiBand", "signal_wifi_band", "wifi_band", "frequency", "networkFrequency", "SignalWifiBand"]);
  const vpn = getTelemetryValue(telemetry, ["vpn", "vpnStatus", "vpnDetected", "signal_vpn", "vpn_status", "signal_vpn_detected", "SignalVpnDetected"]);
  const telemetryInsights = buildTelemetryInsights(telemetry);
  const processor = getTelemetryValue(telemetry, ["processor", "signal_processor", "cpuProcessor", "deviceProcessor", "SignalProcessor"]);
  const cores = getTelemetryValue(telemetry, ["cores", "signal_cores", "cpuCores", "deviceCores", "SignalCores"]);
  const telemetryGroups = [
    {
      title: "Device",
      metrics: [
        { label: "CPU", icon: deviceCpuIcon, value: formatTelemetryMeasurement(telemetry, ["cpu", "cpuUsage", "signal_cpu", "cpu_percent", "SignalCPU"], "%"), color: getCpuMetricColor(cpu) },
        { label: "Memory", icon: deviceMemoryIcon, value: formatTelemetryMeasurement(telemetry, ["memory", "memoryUsage", "signal_memory", "memory_percent", "SignalMemory"], "%"), color: getMemoryMetricColor(memory) },
        { label: "Processor", icon: deviceProcessorIcon, value: processor, color: processor === "No data" ? signalMetricColors.unknown : signalMetricColors.healthy },
        { label: "Cores", icon: deviceCoresIcon, value: cores, color: cores === "No data" ? signalMetricColors.unknown : signalMetricColors.healthy },
      ],
    },
    {
      title: "Workspace",
      metrics: [
        { label: "Wi-Fi Strength", icon: workspaceWifiStrengthIcon, value: wifiColor === signalMetricColors.unknown ? "Unknown" : formatTelemetryMeasurement(telemetry, ["wifiStrength", "wiFiStrength", "signal_wifi_strength", "wifi_strength", "SignalWifiStrength"], "%"), color: wifiColor },
        { label: "Current Network", icon: workspaceCurrentNetworkIcon, value: currentNetwork, color: currentNetwork === "No data" ? signalMetricColors.unknown : signalMetricColors.healthy },
        { label: "Frequency", icon: workspaceNetworkFrequencyIcon, value: frequency, color: getWifiBandMetricColor(frequency, wifiStrength, wifiStatus) },
        { label: "VPN", icon: workspaceVpnIcon, value: vpn, color: getVpnMetricColor(vpn) },
      ],
    },
    {
      title: "Network",
      metrics: [
        { label: "Download", icon: networkDownloadIcon, value: formatTelemetryMeasurement(telemetry, ["downloadSpeed", "signal_download_speed", "download_speed", "SignalDownloadSpeed"], " Mbps"), color: getDownloadMetricColor(download) },
        { label: "Upload", icon: networkUploadIcon, value: formatTelemetryMeasurement(telemetry, ["uploadSpeed", "signal_upload_speed", "upload_speed", "SignalUploadSpeed"], " Mbps"), color: getUploadMetricColor(upload) },
        { label: "Latency", icon: networkLatencyIcon, value: formatTelemetryMeasurement(telemetry, ["latency", "ping", "signal_latency", "latencyMs", "SignalPing"], " ms"), color: getLatencyMetricColor(latency) },
        { label: "Packet loss", icon: networkPacketLossIcon, value: formatTelemetryMeasurement(telemetry, ["packetLoss", "signal_packet_loss", "packet_loss", "SignalPacketLoss"], "%"), color: getPacketLossMetricColor(packetLoss) },
      ],
    },
  ];

  return (
    <>
      <div className="telemetryDetailRow">
        <div className="telemetryDetailLead">
          {recommendation ? <p>{recommendation}</p> : null}
          <SignalScoreTrendChart
            error={signalScoreTrendError}
            isLoading={signalScoreTrendLoading}
            trend={signalScoreTrend}
          />
        </div>
        {telemetryGroups.map((group) => (
          <section className="telemetryDetailCell" key={group.title} aria-label={`${group.title} telemetry`}>
            <h4>{group.title}</h4>
            <div className="telemetryGroupMetrics" aria-label={`${group.title} telemetry`}>
              {group.metrics.map((item) => (
                <div className="metricItem" key={item.label}>
                  <span className="metricLabel">
                    <img src={item.icon} alt="" aria-hidden="true" />
                    <span>{item.label}</span>
                  </span>
                  <strong style={{ color: item.color }} title={getTelemetryValueTitle(item.value)}>
                    {truncateTelemetryValue(item.value)}
                  </strong>
                </div>
              ))}
            </div>
          </section>
        ))}
      </div>
      <section className="telemetryInsights" aria-label="Participant telemetry insights">
        <div className="telemetryInsightsHeader">
          <img className="signal-reccommendation-icon-svg telemetryInsightsIcon" src={signalRecommendationIcon} alt="" aria-hidden="true" />
          <h3>Insights</h3>
        </div>
        {hasActiveAnalysisSession ? (
          <>
            <div className="telemetryInsightOverall">
              <span>Overall recommendations</span>
              {telemetryInsights.recommendations.length > 0 ? telemetryInsights.recommendations.map((recommendation) => (
                <p key={recommendation.text}>- {recommendation.text}</p>
              )) : <p>{telemetryInsights.overall}</p>}
            </div>
            <div className="telemetryInsightCategories">
              {telemetryInsights.categories.map((insight) => (
                <article className="telemetryInsightItem" key={insight.title}>
                  <span style={{ color: getTelemetryInsightColor(insight.severity) }}>{insight.title}</span>
                  <p style={{ color: getTelemetryInsightColor(insight.severity) }}>{insight.text}</p>
                </article>
              ))}
            </div>
          </>
        ) : (
          <p className="telemetryInsightsEmpty">Run analysis to gather insights.</p>
        )}
      </section>
    </>
  );
}

