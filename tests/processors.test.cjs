const test = require("node:test");
const assert = require("node:assert/strict");
const {
  processRecipients,
  validateEmailFormat,
  checkForTypos,
  extractEmail,
} = require("../src/taskpane/processors");
const { normalizeSettings, boundedInvalid } = require("../src/shared/settings");
const { csvCell, toCsv } = require("../src/shared/csv");
function process(to, steps, extra = {}) {
  return processRecipients({
    to,
    cc: [],
    bcc: [],
    userSettings: { enabledSteps: steps, stepOrder: steps, ...extra },
  });
}
test("dedupe uses addresses across fields, preserving To then CC then BCC", () => {
  const result = processRecipients({
    to: ["Alice <a@example.com>", "A@EXAMPLE.COM"],
    cc: ["a@example.com", "b@example.com"],
    bcc: ["b@example.com", "c@example.com"],
    userSettings: { enabledSteps: ["dedupe"] },
  });
  assert.deepEqual(result.result, {
    to: ["Alice <a@example.com>"],
    cc: ["b@example.com"],
    bcc: ["c@example.com"],
  });
});
test("validation blocks before any ordered destructive operation", () => {
  const input = ["broken", "outside@example.org"];
  const result = process(input, ["removeExternal", "validate"], {
    internalDomains: ["example.com"],
  });
  assert.equal(result.success, false);
  assert.deepEqual(result.result.to, input);
  assert.equal(result.actions.length, 0);
});
test("sort leaves input intact; disabling validation does not filter recipients", () => {
  const input = ["z@example.com", "broken", "a@example.com"];
  assert.deepEqual(process(input, ["sort"]).result.to, [
    "a@example.com",
    "broken",
    "z@example.com",
  ]);
  assert.deepEqual(input, ["z@example.com", "broken", "a@example.com"]);
});
test("step order is applied, including sort after internal prioritization", () => {
  const input = ["Z <z@internal.com>", "A <a@external.com>"];
  assert.equal(
    process(input, ["prioritizeInternal", "sort"], { internalDomains: ["internal.com"] }).result
      .to[0],
    input[1]
  );
  assert.equal(
    process(input, ["sort", "prioritizeInternal"], { internalDomains: ["internal.com"] }).result
      .to[0],
    input[0]
  );
});
test("domain matching accepts subdomains but rejects suffix impersonation", () => {
  const result = process(
    ["a@EXAMPLE.COM", "b@sub.example.com", "c@notexample.com", "d@example.com.evil.org"],
    ["removeExternal"],
    { internalDomains: ["example.com"] }
  );
  assert.deepEqual(result.result.to, ["a@EXAMPLE.COM", "b@sub.example.com"]);
});
test("empty domains cannot remove all recipients", () =>
  assert.deepEqual(process(["a@example.com"], ["removeExternal"]).result.to, ["a@example.com"]));
test("format edge cases and conservative SMTP length checks", () => {
  for (const email of ["-tag@example.com", "o'brien@example.com", "a+b@example.com"])
    assert.equal(validateEmailFormat(email).isValid, true, email);
  for (const email of [
    "a.@example.com",
    ".a@example.com",
    "a..b@example.com",
    "a@-bad.com",
    "a@bad-.com",
    "x@local",
    "a".repeat(65) + "@example.com",
  ])
    assert.equal(validateEmailFormat(email).isValid, false, email);
});
test("typo hints never flag exact common domains or alter local parts", () => {
  assert.equal(checkForTypos("me@gmail.com").hasTypo, false);
  assert.equal(checkForTypos("gmial.com@gmial.com").suggestion, "gmial.com@gmail.com");
  assert.equal(checkForTypos("a@constructor").hasTypo, false);
  assert.equal(extractEmail("Name <A@example.com>  "), "A@example.com");
});
test("malformed settings are normalized and analytics defaults off", () => {
  const settings = normalizeSettings({
    enabledSteps: "removeExternal",
    stepOrder: ["unknown", "sort", "sort"],
    internalDomains: ['"><img>', null, " Example.COM "],
    analyticsEnabled: "true",
  });
  assert.deepEqual(settings.internalDomains, ["example.com"]);
  assert.equal(settings.analyticsEnabled, false);
  assert.equal(new Set(settings.stepOrder).size, settings.stepOrder.length);
});
test("saved invalid lists are deduplicated and size bounded", () => {
  assert.deepEqual(boundedInvalid(["bad", "BAD", null]), ["bad"]);
  assert.ok(
    boundedInvalid(Array.from({ length: 1000 }, (_, i) => String(i) + "a".repeat(500))).length < 100
  );
});
test("CSV cells escape delimiters, quotes and formula triggers", () => {
  assert.equal(csvCell('=HYPERLINK("x")'), '"\'=HYPERLINK(""x"")"');
  assert.equal(csvCell(" \t@SUM(A1)"), '"\' \t@SUM(A1)"');
  assert.equal(
    toCsv([["A;B <a@b.com>", 'x"y'], ["line\nbreak"]]),
    '"A;B <a@b.com>";"x""y"\r\n"line\nbreak"'
  );
});
