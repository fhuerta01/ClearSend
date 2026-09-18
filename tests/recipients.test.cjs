const test = require("node:test");
const assert = require("node:assert/strict");
const { readRecipients, writeRecipients } = require("../src/shared/recipients");
const { officeMock } = require("./helpers.cjs");
const before = { to: ["a@example.com"], cc: ["b@example.com"], bcc: [] };
const after = { to: ["z@example.com"], cc: [], bcc: ["c@example.com"] };
test("rejects an oversized write before touching any field", async () => {
  const mock = officeMock(before);
  await assert.rejects(
    writeRecipients(mock.office, before, { ...after, cc: Array(101).fill("a@example.com") }),
    /100 recipients/
  );
  assert.deepEqual(mock.writes, []);
});
test("updates only changed fields and supports clearing a field", async () => {
  const mock = officeMock(before);
  await writeRecipients(mock.office, before, { ...before, cc: [] });
  assert.deepEqual(mock.writes, ["cc"]);
  assert.deepEqual(mock.lists.cc, []);
});
test("partial failures restore previously written fields", async () => {
  const mock = officeMock(before);
  mock.fail("cc");
  await assert.rejects(writeRecipients(mock.office, before, after), /Previous lists were restored/);
  assert.deepEqual(mock.lists, before);
});
test("rollback failure is explicit and never reported as success", async () => {
  const mock = officeMock(before);
  mock.fail("cc", true);
  await assert.rejects(writeRecipients(mock.office, before, after), /could not restore all/);
});
test("detects edits since the snapshot before writing or undoing", async () => {
  const mock = officeMock(before);
  mock.lists.to.push("new@example.com");
  await assert.rejects(writeRecipients(mock.office, before, after), /Recipients changed/);
  assert.deepEqual(mock.writes, []);
});
test("unresolved recipients are never silently dropped", async () => {
  const mock = officeMock(before);
  mock.office.context.mailbox.item.to.getAsync = (cb) =>
    cb({ status: "succeeded", value: [{ displayName: "Unresolved" }] });
  await assert.rejects(readRecipients(mock.office), /Resolve every recipient/);
});
