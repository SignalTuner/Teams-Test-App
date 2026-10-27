// Run with the existing localhost Vite preview and Playwright environment overrides.
const { test } = require("node:test");
const assert = require("node:assert/strict");
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "playwright");
const origin = process.env.PERSONAL_TEST_ORIGIN || "http://127.0.0.1:4317";

async function configPage(t, preference) {
  const browser = await chromium.launch({ headless: true, ...(process.env.BROWSER_EXECUTABLE ? { executablePath: process.env.BROWSER_EXECUTABLE } : {}) });
  t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 500, height: 560 } });
  page.setDefaultTimeout(10000);
  const errors = [], apiCalls = [];
  page.on("pageerror", error => errors.push(error.message));
  await page.addInitScript(preference => {
    if (preference) localStorage.setItem("signaltunerThemePreference", preference);
  }, preference);
  await page.route("**/api/**", route => { apiCalls.push(route.request().url()); return route.abort(); });
  await page.route("**/@microsoft_teams-js.js*", route => route.fulfill({ contentType: "application/javascript", body: `
    window.configHarness = { valid: false, config: null, result: null, rejectSave: false };
    export const app = { initialize: async () => {}, getContext: async () => ({app:{theme:"dark"}}),
      registerOnThemeChangeHandler: callback => { window.configHarness.theme = callback; } };
    export const pages = { config: {
      setValidityState: valid => { window.configHarness.valid = valid; },
      registerOnSaveHandler: callback => { window.configHarness.save = callback; },
      setConfig: async config => { if (window.configHarness.rejectSave) throw new Error("Synthetic save failure"); window.configHarness.config = config; }
    } };
    export const authentication = {}; export const meeting = {};` }));
  // Vite's base is /tabs/home; serve its unchanged entry document at the real config route.
  await page.route(`${origin}/tabs/config`, async route => route.fulfill({ response: await route.fetch({ url: `${origin}/tabs/home` }) }));
  await page.goto(`${origin}/tabs/config`);
  await page.getByRole("heading", { name: "Add SignalTuner to your meeting" }).waitFor();
  await page.waitForFunction(() => window.configHarness?.valid && window.configHarness?.theme);
  return { page, errors, apiCalls };
}

test("configuration overview fits the dialog, follows Teams themes and preserves native Save", async t => {
  const { page, errors, apiCalls } = await configPage(t);
  assert.equal(await page.locator(".configFeatures li").count(), 3);
  assert.equal(await page.getByRole("button").count(), 0);
  assert.equal(await page.getByRole("checkbox").count(), 0);
  assert.equal(await page.getByRole("link", { name: "Privacy Policy" }).getAttribute("href"), "/tabs/privacy");
  for (const theme of ["dark", "default", "contrast"]) {
    await page.evaluate(theme => window.configHarness.theme(theme), theme);
    assert.equal(await page.locator("html").getAttribute("data-teams-theme"), theme);
    for (const width of [500, 390]) {
      await page.setViewportSize({ width, height: width === 500 ? 320 : 360 });
      assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), `${width}px horizontal overflow`);
      assert.ok(await page.evaluate(() => document.documentElement.scrollHeight <= innerHeight), `${width}px dialog needs no scrolling`);
      if (process.env.CONFIG_TEST_SCREENSHOTS) {
        require("node:fs").mkdirSync(process.env.CONFIG_TEST_SCREENSHOTS, { recursive: true });
        await page.screenshot({ path: require("node:path").join(process.env.CONFIG_TEST_SCREENSHOTS, `config-${theme}-${width}.png`), fullPage: true });
      }
    }
  }
  const save = () => page.evaluate(async () => {
    const harness = window.configHarness;
    harness.result = null;
    harness.save({ notifySuccess: () => { harness.result = "success"; }, notifyFailure: message => { harness.result = message; } });
  });
  await save();
  await page.waitForFunction(() => window.configHarness.result === "success");
  assert.deepEqual(await page.evaluate(() => window.configHarness.config), {
    entityId: "home", suggestedDisplayName: "SignalTuner", contentUrl: `${origin}/tabs/home`, websiteUrl: `${origin}/tabs/home`,
  });
  await page.evaluate(() => { window.configHarness.rejectSave = true; });
  await save();
  await page.waitForFunction(() => window.configHarness.result === "Synthetic save failure");
  assert.deepEqual(errors, []);
  assert.deepEqual(apiCalls, []);
});

test("configuration honors saved appearance while allowing Teams high contrast", async t => {
  const { page } = await configPage(t, "light");
  assert.equal(await page.locator("html").getAttribute("data-teams-theme"), "default");
  await page.evaluate(() => window.configHarness.theme("contrast"));
  assert.equal(await page.locator("html").getAttribute("data-teams-theme"), "contrast");
  await page.evaluate(() => window.configHarness.theme("dark"));
  assert.equal(await page.locator("html").getAttribute("data-teams-theme"), "default");
});
