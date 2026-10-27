const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");
const Module = require("node:module");
const filename = path.resolve(__dirname, "../src/Tab/meetingInvitations.ts");
let supported = true;
let opened;
const sdk = { dialog: { adaptiveCard: { bot: {
  isSupported: () => supported,
  open: (info, callback) => { opened = { info, callback }; },
} } } };
const loaded = new Module(filename, module);
loaded.filename = filename;
loaded.paths = module.paths;
const originalRequire = loaded.require.bind(loaded);
loaded.require = name => name === "@microsoft/teams-js" ? sdk : originalRequire(name);
loaded._compile(ts.transpileModule(fs.readFileSync(filename, "utf8"), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
}).outputText, filename);
const { openMeetingInvitations, invitationBotId } = loaded.exports;

test("Invite opens a bot dialog for this meeting without sending an invitation", () => {
  supported = true;
  openMeetingInvitations("synthetic-meeting", assert.fail);
  const manifest = JSON.parse(fs.readFileSync(path.resolve(__dirname, "../appPackage/manifest.json"), "utf8"));
  assert.equal(invitationBotId, manifest.bots[0].botId);
  assert.equal(invitationBotId, manifest.webApplicationInfo.id, "Meeting RSC permissions must belong to the bot");
  assert.equal(manifest.webApplicationInfo.resource, "${{TEAMS_SSO_RESOURCE}}");
  assert.ok(manifest.authorization.permissions.resourceSpecific.some(permission =>
    permission.name === "OnlineMeeting.ReadBasic.Chat" && permission.type === "Application"));
  assert.equal(opened.info.completionBotId, invitationBotId);
  const card = JSON.parse(opened.info.card);
  assert.equal(card.actions.length, 1);
  assert.equal(card.actions[0].data.action, "signaltuner.invite.load");
  assert.equal(card.actions[0].data.meetingId, "synthetic-meeting");
  assert.match(card.actions[0].data.requestId, /^[0-9a-f-]{36}$/i);
  assert.ok(card.body.some(block => block.text.includes("visible to everyone")));
});

test("missing meeting or unsupported hosts fail before opening the dialog", () => {
  opened = undefined;
  assert.throws(() => openMeetingInvitations("", assert.fail), /inside the meeting/);
  supported = false;
  assert.throws(() => openMeetingInvitations("meeting", assert.fail), /inside the meeting/);
  assert.equal(opened, undefined);
  supported = true;
});

test("cancel is quiet and a bot failure shows actionable feedback", () => {
  const errors = [];
  openMeetingInvitations("meeting", message => errors.push(message));
  opened.callback({ err: "User canceled or closed the dialog." });
  opened.callback({ result: {} });
  assert.equal(errors.length, 0);
  opened.callback({ err: "Invalid app ID." });
  assert.equal(errors.length, 1);
  assert.match(errors[0], /bot is installed in this meeting/);
});
