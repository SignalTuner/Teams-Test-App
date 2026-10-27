import { useCallback, useEffect, useRef, useState } from "react";
import type { TelemetryRecord } from "./dashboardTypes";

export type PersonalApi = <T>(path: string, options?: RequestInit) => Promise<T>;
export function createPersonalApi(base: string, token: string, expired: () => void): PersonalApi {
  return async <T,>(path: string, options: RequestInit = {}): Promise<T> => {
    const response = await fetch(`${base.replace(/\/+$/, "")}/api/${path}`, {
      ...options, cache: "no-store", headers: { ...options.headers, Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    });
    if (options.signal?.aborted) throw new DOMException("Request cancelled", "AbortError");
    if (response.status === 401) expired();
    if (!response.ok) throw Object.assign(new Error(response.status === 401 ? "Your session has expired. Sign in again."
      : response.status === 404 ? "No data is available."
      : response.status === 403 ? "You no longer have access to this information."
      : "This information is temporarily unavailable. Please retry."), { status: response.status });
    return response.json();
  };
}

// Independent, cancellable resources; hidden tabs stop polling and refresh on return.
export function usePersonalResource<T>(api: PersonalApi, path: string | null, interval = 30000) {
  const [data, setData] = useState<T | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);
  const readRef = useRef<() => Promise<void>>(async () => {});
  const refresh = useCallback(() => readRef.current(), []);
  useEffect(() => {
    let stopped = false, fetching = false;
    let controller: AbortController | undefined;
    setData(null); setError(""); setLoading(Boolean(path));
    const read = async () => {
      if (!path || stopped || fetching || document.hidden) return;
      fetching = true; controller = new AbortController();
      const timeout = setTimeout(() => controller?.abort(), 15000);
      try {
        const value = await api<T>(path, { signal: controller.signal });
        if (!stopped) { setData(value); setError(""); }
      } catch (reason) {
        if (!stopped) { setData(null); setError(reason instanceof Error ? reason.message : "Unable to refresh. Please retry."); }
      } finally { clearTimeout(timeout); fetching = false; if (!stopped) setLoading(false); }
    };
    readRef.current = read;
    void read();
    const timer = setInterval(() => void read(), interval);
    document.addEventListener("visibilitychange", read);
    return () => { stopped = true; controller?.abort(); clearInterval(timer); document.removeEventListener("visibilitychange", read); };
  }, [api, path, interval]);
  return { data, error, loading, refresh };
}

export type MeetingSummary = { meetingId: string; name: string; source: string; meetingSessionId: number;
  createdAt: string; participantCount: number };
export type MeetingDetail = { meeting: MeetingSummary; canAnalyze: boolean; participants: Array<{
  participantId: string; userId: number; name: string; email: string | null; meetingRole: string | null;
  signal: { signalScore: number | null; signalWifiStatus: string | null; signalBandwidthStatus: string | null;
    signalSystemStatus: string | null; signalOverallStatus: string | null } | null;
}>; requests: import("./RemoteAnalysisButton").RequestState[] };

export function teamsHistory(meetings: MeetingSummary[]) {
  return meetings.filter(meeting => meeting.source === "Teams" && /^teams-[1-9]\d*$/.test(meeting.meetingId))
    .sort((a, b) => Date.parse(b.createdAt) - Date.parse(a.createdAt));
}
export function utcDate(value: unknown): Date | null {
  if (typeof value !== "string" || !value) return null;
  const date = new Date(/[zZ]$|[+-]\d\d:\d\d$/.test(value) ? value : `${value}Z`);
  return Number.isFinite(date.getTime()) ? date : null;
}
export function signalPresentation(signal: TelemetryRecord | null, now = Date.now()) {
  const updated = utcDate(signal?.LastUpdated ?? signal?.Timestamp);
  const stale = Boolean(updated && now - updated.getTime() > 120000);
  const offline = signal?.SignalOverallStatus === "Offline" || stale;
  const hasData = Boolean(signal && !offline && typeof signal.SignalScore === "number");
  return { updated, stale, offline, hasData, score: hasData ? signal!.SignalScore as number : null,
    overall: offline ? "Offline" : typeof signal?.SignalOverallStatus === "string" ? signal.SignalOverallStatus : null };
}

export type SelfOperation = { requestId: string; targetUserId: number; meetingSessionId: null; mode: "Test" | "Report" };
export const pendingAnalysisKey = (userId: number) => `signaltunerPersonalAnalysis:${userId}`;
export function readPendingAnalysis(storage: Pick<Storage, "getItem">, userId: number): SelfOperation | null {
  try {
    const value = JSON.parse(storage.getItem(pendingAnalysisKey(userId)) || "null");
    return value?.targetUserId === userId && value.meetingSessionId === null && /^[a-f0-9]{32}$/.test(value.requestId)
      && ["Test", "Report"].includes(value.mode) ? value : null;
  } catch { return null; }
}

// Persist the exact operation before any network request. A storage failure must prevent submission.
export async function sendSelfAnalysis<T>(api: PersonalApi, storage: Pick<Storage, "setItem" | "removeItem">, operation: SelfOperation, signal: AbortSignal): Promise<T> {
  const key = pendingAnalysisKey(operation.targetUserId);
  storage.setItem(key, JSON.stringify(operation));
  const result = await api<T>("remote-analysis", { method: "POST", body: JSON.stringify(operation), signal });
  storage.removeItem(key);
  return result;
}
