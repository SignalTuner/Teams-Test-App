import type { RequestState } from "./RemoteAnalysisButton";

export type MeetingAnalysisOperation = { requestId: string; participantUserIds: number[] };
export type MeetingAnalysisResult = { targetUserId: number; request?: RequestState; error?: string; statusCode?: number };
const requestIdPattern = /^[a-f0-9]{32}$/;

export function activeMeetingParticipantIds(participants: Array<{ userId: number; clientDataStatus: string }>): number[] {
  return [...new Set(participants.filter(p => p.clientDataStatus === "active" && Number.isSafeInteger(p.userId) && p.userId > 0).map(p => p.userId))];
}

export function readMeetingAnalysisOperation(saved: string | null): MeetingAnalysisOperation | null {
  if (!saved) return null;
  const value = JSON.parse(saved) as MeetingAnalysisOperation;
  if (!value || !requestIdPattern.test(value.requestId) || !Array.isArray(value.participantUserIds)
    || value.participantUserIds.length < 1 || value.participantUserIds.length > 1000
    || value.participantUserIds.some(id => !Number.isSafeInteger(id) || id <= 0)
    || new Set(value.participantUserIds).size !== value.participantUserIds.length)
    throw new Error("The saved meeting analysis could not be read. Reopen the original tab to check its requests.");
  return { requestId: value.requestId, participantUserIds: value.participantUserIds };
}

export function validateMeetingAnalysisResults(value: unknown, operation: MeetingAnalysisOperation, meetingSessionId: number): MeetingAnalysisResult[] {
  const data = value as { batchId?: string; meetingSessionId?: number; results?: MeetingAnalysisResult[] } | null;
  if (!data || data.batchId !== operation.requestId || data.meetingSessionId !== meetingSessionId
    || !Array.isArray(data.results) || data.results.length !== operation.participantUserIds.length
    || new Set(data.results.map(result => result?.targetUserId)).size !== operation.participantUserIds.length
    || data.results.some(result => !result || !operation.participantUserIds.includes(result.targetUserId)
      || (result.request ? !requestIdPattern.test(result.request.requestId) || typeof result.request.state !== "string"
        || result.request.targetUserId !== result.targetUserId || Boolean(result.error)
        : typeof result.error !== "string" || !result.error.trim())))
    throw new Error("The response could not be confirmed. Retry this same meeting analysis to recover its results.");
  return data.results;
}

export async function requestMeetingAnalysis(apiBaseUrl: string, token: string, meetingSessionId: number,
  operation: MeetingAnalysisOperation, signal: AbortSignal): Promise<MeetingAnalysisResult[]> {
  const response = await fetch(`${apiBaseUrl.replace(/\/$/, "")}/api/remote-analysis/teams-meeting`, {
    method: "POST", signal,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ ...operation, meetingSessionId }),
  });
  if (!response.ok) {
    throw new Error(response.status === 401 ? "Sign in again to request a meeting analysis."
      : response.status === 403 ? "Your access to this meeting could not be confirmed."
      : response.status === 404 ? "Meeting analysis is not available on this server yet."
      : response.status === 429 ? "Please wait before retrying this meeting analysis."
      : "The request could not be confirmed. Retry this same meeting analysis; accepted requests keep running.");
  }
  let data: unknown;
  try { data = await response.json(); }
  catch { throw new Error("The response could not be confirmed. Retry this same meeting analysis to recover its results."); }
  return validateMeetingAnalysisResults(data, operation, meetingSessionId);
}
