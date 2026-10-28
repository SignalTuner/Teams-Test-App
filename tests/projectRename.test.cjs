const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

test("rename helper moves one checkout, updates every map/workspace, preserves identity and can resume", t => {
  const prefix = path.join(os.tmpdir(), "signaltuner-rename-fixture-");
  const fixture = fs.mkdtempSync(prefix);
  const fixtureReal = fs.realpathSync(fixture);
  t.after(() => {
    // Only this test's resolved temporary fixture may be removed recursively.
    assert.equal(fs.realpathSync(fixture), fixtureReal);
    assert.ok(path.resolve(fixture).startsWith(path.resolve(prefix)));
    fs.rmSync(fixture, { recursive: true, force: true });
  });
  const old = path.join(fixture, "Teams Test App"), target = path.join(fixture, "SignalTuner Teams App");
  const projects = ["SignalTuner", "signaltuner-back", "signaltuner-front", "is-service-down", "Teams Test App"];
  for (const project of projects) {
    fs.mkdirSync(path.join(fixture, project), { recursive: true });
    fs.writeFileSync(path.join(fixture, project, "AGENTS.md"), `Project map: ${old}\nUnrelated text preserved.\n`);
  }
  fs.mkdirSync(path.join(old, "scripts"));
  fs.copyFileSync(path.join(__dirname, "../scripts/complete-project-rename.ps1"), path.join(old, "scripts/complete-project-rename.ps1"));
  fs.mkdirSync(path.join(old, ".git")); fs.writeFileSync(path.join(old, ".git/config"), "synthetic-remote-must-not-change");
  fs.mkdirSync(path.join(old, "appPackage")); fs.writeFileSync(path.join(old, "appPackage/manifest.json"), '{"id":"synthetic-preserved-id"}');
  const workspace = path.join(fixture, "SignalTuner.code-workspace");
  fs.writeFileSync(workspace, '{"folders":[{"name":"SignalTuner — Microsoft Teams app", "path": "Teams Test App"}]}');
  const run = (root, ...args) => spawnSync("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", path.join(root, "scripts/complete-project-rename.ps1"), ...args], { cwd: fixture, encoding: "utf8" });
  let result = run(old, "-ValidateOnly"); assert.equal(result.status, 0, result.stderr); assert.ok(fs.existsSync(old)); assert.ok(!fs.existsSync(target));
  // Two checkouts must never be merged, and the workspace must remain unchanged on refusal.
  fs.mkdirSync(target);
  result = run(old); assert.notEqual(result.status, 0); assert.match(fs.readFileSync(workspace, "utf8"), /Teams Test App/);
  fs.rmdirSync(target); // Empty directory created immediately above, no recursive deletion.
  result = run(old); assert.equal(result.status, 0, result.stderr); assert.ok(!fs.existsSync(old));
  for (const project of projects.map(p => p === "Teams Test App" ? "SignalTuner Teams App" : p)) {
    const doc = fs.readFileSync(path.join(fixture, project, "AGENTS.md"), "utf8");
    assert.ok(doc.includes(target)); assert.ok(!doc.includes(old)); assert.ok(doc.includes("Unrelated text preserved."));
  }
  assert.equal(fs.readFileSync(path.join(target, ".git/config"), "utf8"), "synthetic-remote-must-not-change");
  assert.equal(fs.readFileSync(path.join(target, "appPackage/manifest.json"), "utf8"), '{"id":"synthetic-preserved-id"}');
  assert.deepEqual(JSON.parse(fs.readFileSync(workspace, "utf8")).folders[0], { name: "SignalTuner Teams App", path: "SignalTuner Teams App" });
  result = run(target); assert.equal(result.status, 0, result.stderr);
});
