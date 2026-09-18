const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { createRequire } = require("node:module");
const { JSDOM } = require("jsdom");
const { officeMock } = require("./helpers.cjs");
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
            createAnalytics: () => ({
              setEnabled() {},
              track(event) {
                events.push(event);
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
