export const deviceAnalysisLimitMessage = "This device has received too many requests. Try again later.";
const requesterAnalysisLimitMessage = "Please wait before requesting another analysis.";

// Only known public messages may cross from an error response into the UI.
export function analysisRateLimitMessage(message: unknown): string {
  return message === deviceAnalysisLimitMessage ? deviceAnalysisLimitMessage : requesterAnalysisLimitMessage;
}
