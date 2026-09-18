const fs = require("node:fs");
const path = require("node:path");
const assert = require("node:assert/strict");
const { JSDOM } = require("jsdom");
for (const page of ["taskpane.html", "commands.html"]) {
  const document = new JSDOM(fs.readFileSync(path.join("dist", page), "utf8")).window.document;
  const scripts = [];
  for (const node of document.querySelectorAll("script[src], link[href]")) {
    const url = node.getAttribute("src") || node.getAttribute("href");
    if (/^https:/.test(url)) continue;
    assert.ok(fs.existsSync(path.join("dist", url)), `Missing asset in ${page}: ${url}`);
    if (node.tagName === "SCRIPT") scripts.push(url);
  }
  assert.equal(scripts.length, new Set(scripts).size, "Scripts must not initialize twice");
  assert.equal(document.querySelectorAll("[onclick]").length, 0, "Inline handlers are disallowed");
}
