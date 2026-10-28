const { test } = require("node:test");
const assert = require("node:assert/strict");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const load = require("./loadTabModule.cjs");
const { allowsExternalCommerce } = load("teamsHost.ts");
const { ExternalWebsiteLink, PaymentDisclosure } = load("TeamsHostProvider.tsx");

test("only recognized Teams desktop/web hosts expose external commerce", () => {
  for (const name of ["Teams", "TeamsModern"]) {
    for (const clientType of ["desktop", "web", "macos"])
      assert.equal(allowsExternalCommerce({ app: { host: { name, clientType } } }), true);
    for (const clientType of ["ios", "android", "ipados", "teamsPhones", "teamsDisplays", "surfaceHub", "rigel", "unknown", "", undefined])
      assert.equal(allowsExternalCommerce({ app: { host: { name, clientType } } }), false, clientType);
  }
  for (const context of [null, {}, { app: {} }, { app: { host: { clientType: "web" } } }, { app: { host: { name: "Outlook", clientType: "web" } } }])
    assert.equal(allowsExternalCommerce(context), false);
});

test("initial rendering cannot flash purchase links or payment prompts before host discovery", () => {
  assert.equal(renderToStaticMarkup(React.createElement(ExternalWebsiteLink, { href: "https://example.invalid/plans" }, "Buy credits")), "");
  assert.equal(renderToStaticMarkup(React.createElement(PaymentDisclosure)), "");
});
