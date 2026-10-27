import { useEffect, useState } from "react";
import "./AnalysisCountdown.css";
import type { LiveTelemetry } from "./LiveAnalysisCharts";

export function AnalysisCanceled() {
  return <span className="analysis-canceled" role="status">Canceled</span>;
}

export default function AnalysisCountdown({ telemetry }: { telemetry?: LiveTelemetry | null }) {
  const [deadline, setDeadline] = useState<number | null>(null);
  const [now, setNow] = useState(Date.now);
  const elapsed = telemetry?.snapshot?.elapsedSeconds;
  const age = typeof telemetry?.ageSeconds === "number" && Number.isFinite(telemetry.ageSeconds) ? Math.max(0, telemetry.ageSeconds) : 0;

  useEffect(() => {
    if (typeof elapsed !== "number" || !Number.isFinite(elapsed) || elapsed < 0 || elapsed > 300) return;
    const received = Date.now();
    const end = received + Math.max(0, 300 - elapsed - age) * 1000;
    // Repeated or delayed snapshots must never restart the countdown.
    setDeadline(previous => previous === null ? end : Math.min(previous, end));
    setNow(received);
  }, [elapsed, age]);

  useEffect(() => {
    if (deadline === null) return;
    const update = () => setNow(Date.now());
    const timer = setInterval(update, 1000);
    document.addEventListener("visibilitychange", update);
    return () => { clearInterval(timer); document.removeEventListener("visibilitychange", update); };
  }, [deadline]);

  const remaining = deadline === null ? 300 : Math.max(0, Math.ceil((deadline - now) / 1000));
  const time = `${Math.floor(remaining / 60)}:${String(remaining % 60).padStart(2, "0")}`;
  return <span className="analysis-countdown" role="timer" aria-label="Analysis time remaining" aria-live="off"
    title={deadline === null ? "Countdown starts when analysis data arrives." : "Estimated time remaining"}>{time} remaining</span>;
}
