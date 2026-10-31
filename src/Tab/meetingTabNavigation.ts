// Navigation hints only: authorization and paid analysis still require the normal
// authenticated dashboard and an explicit confirmation in MeetingAnalysisButton.
export type MeetingTabContext = {
  app?: { appId?: { toString(): string } | string };
  page?: { id?: string; frameContext?: string; subPageId?: string };
  chat?: { id?: string };
  meeting?: { id?: string };
  user?: { tenant?: { id?: string } };
};

const confirmationPrefix = "confirm-meeting-analysis:";
export function meetingAnalysisOpenRequest(context: MeetingTabContext | null): string | undefined {
  const value = context?.page?.subPageId;
  return context?.page?.frameContext === "content" && context.meeting?.id && context.chat?.id
    && value && /^confirm-meeting-analysis:[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value)
    ? value : undefined;
}

export function buildMeetingTabLink(context: MeetingTabContext | null, confirmationId?: string): string {
  const appId = context?.app?.appId?.toString();
  const pageId = context?.page?.id;
  const chatId = context?.chat?.id;
  if (!appId || !pageId || !chatId || !context?.meeting?.id) {
    throw new Error("Open SignalTuner from this meeting's chat tab. Teams has not provided the information needed to open it here.");
  }
  const link = new URL(`https://teams.microsoft.com/l/entity/${encodeURIComponent(appId)}/${encodeURIComponent(pageId)}`);
  link.searchParams.set("context", JSON.stringify({ chatId, contextType: "chat",
    ...(confirmationId ? { subEntityId: confirmationPrefix + confirmationId } : {}) }));
  link.searchParams.set("openInMeeting", "false");
  if (context.user?.tenant?.id) link.searchParams.set("tenantId", context.user.tenant.id);
  return link.toString();
}
