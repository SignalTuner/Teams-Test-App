import * as teamsJs from "@microsoft/teams-js";

// The existing bot in appPackage/manifest.json; this is not the Teams app ID.
export const invitationBotId = "e32ebb3f-81e9-4ca2-85f1-a0a2d7b74999";

export function invitationStartCard(meetingId: string, requestId: string) {
  return {
    type: "AdaptiveCard", version: "1.4",
    body: [
      { type: "TextBlock", text: "Invite meeting participants", weight: "Bolder", size: "Medium", wrap: true },
      { type: "TextBlock", text: "Load the participants from this meeting’s Teams chat, then choose who to invite to open SignalTuner and sign in.", wrap: true },
      { type: "TextBlock", text: "Your invitation will mention the selected people in the meeting chat and be visible to everyone in that chat. Loading participants does not send a message.", wrap: true },
    ],
    actions: [{ type: "Action.Submit", title: "Load participants", data: {
      action: "signaltuner.invite.load", meetingId, requestId,
    } }],
  };
}

export function openMeetingInvitations(meetingId: string, onError: (message: string) => void): void {
  if (!meetingId || !teamsJs.dialog.adaptiveCard.bot.isSupported()) {
    throw new Error("Open SignalTuner inside the meeting in Teams desktop or web to invite participants.");
  }
  teamsJs.dialog.adaptiveCard.bot.open({
    title: "Invite meeting participants",
    size: { height: 520, width: 480 },
    completionBotId: invitationBotId,
    card: JSON.stringify(invitationStartCard(meetingId, crypto.randomUUID())),
  }, ({ err }) => {
    // Closing a native Teams dialog is not an invitation failure.
    if (err && !/cancel|closed by|user closed/i.test(err)) {
      onError("Teams could not complete the invitation dialog. Check that the SignalTuner bot is installed in this meeting, then try again.");
    }
  });
}
