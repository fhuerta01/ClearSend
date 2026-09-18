// Quoting prevents delimiter injection; the apostrophe prevents spreadsheet formulas.
function csvCell(value) {
  let text = String(value);
  // eslint-disable-next-line no-control-regex -- spreadsheet formula prefixes include control characters.
  if (/^[\s\u0000-\u001f]*[=+\-@]/.test(text) || /^[\t\r\n]/.test(text)) text = "'" + text;
  return '"' + text.replace(/"/g, '""') + '"';
}
function toCsv(lists) {
  return lists.map((list) => list.map(csvCell).join(";")).join("\r\n");
}
module.exports = { csvCell, toCsv };
