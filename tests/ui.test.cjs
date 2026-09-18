const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { createRequire } = require("node:module");
const { JSDOM } = require("jsdom");
const { officeMock } = require("./helpers.cjs");
const { createAnalytics } = require("../src/shared/analytics");
const tick = () => new Promise((resolve) => setImmediate(resolve));
async function pane(t, lists, settings = {}) {
  const dom = new JSDOM(
    fs.readFileSync(path.join(__dirname, "../src/taskpane/taskpane.html"), "utf8"),
    { url: "https://clearsend.vercel.app/taskpane.html" }
  );
  t.after(() => dom.window.close());
  const mock = officeMock(lists, settings);
  const events = [];
  const file = path.join(__dirname, "../src/taskpane/taskpane.js");
  const localRequire = createRequire(file);
  const context = {
    Office: mock.office,
    document: dom.window.document,
    window: dom.window,
    navigator: dom.window.navigator,
    location: dom.window.location,
    __ANALYTICS_ORIGIN__: "https://clearsend.vercel.app",
    structuredClone,
    Blob,
    URL,
    setTimeout: (fn, ms) => dom.window.setTimeout(fn, ms),
    clearTimeout: dom.window.clearTimeout.bind(dom.window),
    setInterval: dom.window.setInterval.bind(dom.window),
    clearInterval: dom.window.clearInterval.bind(dom.window),
    require: (name) =>
      name === "../shared/analytics"
        ? {
            createAnalytics: (options) =>
              createAnalytics({
                ...options,
                location: dom.window.location,
                navigator: dom.window.navigator,
                fetch: async (_url, { body }) => {
                  events.push(JSON.parse(body).event);
                },
              }),
          }
        : localRequire(name),
  };
  vm.runInNewContext(fs.readFileSync(file, "utf8"), context);
  await tick();
  return {
    ...mock,
    events,
    document: dom.window.document,
    click: async (id) => {
      dom.window.document.getElementById(id).click();
      await tick();
    },
  };
}
test("new installs count opening and processing; opting out survives restore and reopening", async (t) => {
  const lists = { to: ["a@example.com"], cc: [], bcc: [] };
  const ui = await pane(t, lists);
  assert.equal(ui.document.getElementById("analyticsCheck").checked, true);
  assert.deepEqual(ui.events, ["pane_open"]);
  await ui.click("checkCleanBtn");
  assert.equal(ui.events.filter((event) => event === "process_click").length, 1);
  await ui.click("analyticsCheck");
  assert.equal(ui.values.clearSendSettings.analyticsEnabled, false);
  const count = ui.events.length;
  await ui.click("checkCleanBtn");
  await ui.click("restoreDefaultsBtn");
  assert.equal(ui.events.length, count);
  assert.equal(ui.values.clearSendSettings.analyticsEnabled, false);
  assert.equal(ui.document.getElementById("analyticsCheck").checked, false);
  assert.match(ui.document.getElementById("usageCountsNotice").textContent, /counts are off/);
  const reopened = await pane(t, lists, ui.values.clearSendSettings);
  await reopened.click("checkCleanBtn");
  assert.deepEqual(reopened.events, []);
});
test("restore preserves enabled counting and existing false preferences never emit an opening", async (t) => {
  const lists = { to: [], cc: [], bcc: [] };
  const enabled = await pane(t, lists);
  await enabled.click("restoreDefaultsBtn");
  assert.equal(enabled.values.clearSendSettings.analyticsEnabled, true);
  await enabled.click("checkCleanBtn");
  assert.ok(enabled.events.includes("process_click"));
  const disabled = await pane(t, lists, { analyticsEnabled: false });
  assert.deepEqual(disabled.events, []);
  await disabled.click("analyticsCheck");
  assert.equal(disabled.values.clearSendSettings.analyticsEnabled, true);
  await disabled.click("checkCleanBtn");
  assert.ok(disabled.events.includes("process_click"));
});
test("single click processes once, ordering changes can be undone", async (t) => {
  const original = { to: ["Z <z@example.com>", "A <a@example.com>"], cc: [], bcc: [] };
  const ui = await pane(t, original);
  assert.equal(ui.document.getElementById("checkCleanBtn").disabled, false);
  await ui.click("checkCleanBtn");
  assert.deepEqual(ui.lists.to, ["A <a@example.com>", "Z <z@example.com>"]);
  assert.equal(ui.events.filter((e) => e === "process_click").length, 1);
  assert.equal(ui.document.getElementById("undoBtn").disabled, false);
  await ui.click("undoBtn");
  assert.deepEqual(ui.lists, original);
});
test("list toggle has one handler and mailbox strings remain text", async (t) => {
  const address = "<img src=x onerror=alert(1)> <a@example.com>";
  const ui = await pane(t, { to: [address], cc: [], bcc: [] });
  assert.equal(ui.document.querySelector(".recipient-email").textContent, address);
  assert.equal(ui.document.querySelector("#toContent img"), null);
  await ui.click("toggleToBtn");
  assert.equal(ui.document.getElementById("toContent").style.display, "block");
  await ui.click("toggleToBtn");
  assert.equal(ui.document.getElementById("toContent").style.display, "none");
});
test("panel blocks invalid recipients, counts duplicates independently of display name", async (t) => {
  const ui = await pane(t, {
    to: ["Name <a@example.com>", "a@example.com", "bad"],
    cc: [],
    bcc: [],
  });
  assert.equal(ui.document.getElementById("duplicatedCount").textContent, "1");
  await ui.click("checkCleanBtn");
  assert.deepEqual(ui.writes, []);
  assert.ok(ui.events.includes("process_blocked"));
});
test("deleting a saved invalid address does not write mailbox recipients", async (t) => {
  const ui = await pane(
    t,
    { to: ["bad"], cc: [], bcc: [] },
    { enabledSteps: ["validate"], keepInvalid: true }
  );
  await ui.click("checkCleanBtn");
  ui.document.querySelector("#savedInvalidContent .recipient-delete").click();
  await tick();
  assert.deepEqual(ui.writes, []);
  assert.equal(ui.values.savedInvalidAddresses, undefined);
});
test("ribbon uses the same validation and always completes even if notifications fail", async () => {
  const mock = officeMock({ to: ["bad"], cc: [], bcc: [] });
  let command;
  let completed = 0;
  mock.office.actions.associate = (name, fn) => {
    if (name === "quickClean") command = fn;
  };
  mock.office.context.mailbox.item.notificationMessages.replaceAsync = () => {
    throw new Error("notification failed");
  };
  const file = path.join(__dirname, "../src/commands/commands.js");
  vm.runInNewContext(fs.readFileSync(file, "utf8"), {
    require: createRequire(file),
    Office: mock.office,
    __ANALYTICS_ORIGIN__: "",
  });
  await command({
    completed() {
      completed++;
    },
  });
  assert.equal(completed, 1);
  assert.deepEqual(mock.writes, []);
});
