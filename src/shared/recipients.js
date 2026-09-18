// Office.js is the only recipient I/O boundary. No recipient network API exists.
const FIELDS = Object.freeze(["to", "cc", "bcc"]);
const LIMIT = 100; // Conservative setAsync limit supported across target Outlook clients.
function fieldCall(office, field, method, value) {
  return new Promise((resolve, reject) => {
    if (!field || typeof field[method] !== "function") {
      reject(new Error("Open ClearSend while composing an email."));
      return;
    }
    const callback = (result) =>
      result.status === office.AsyncResultStatus.Succeeded
        ? resolve(result.value)
        : reject(new Error("Outlook could not read or update recipients."));
    if (method === "getAsync") field[method](callback);
    else field[method](value, callback);
  });
}
function recipientString(recipient) {
  if (!recipient || typeof recipient.emailAddress !== "string" || !recipient.emailAddress.trim()) {
    // Never silently drop unresolved recipients from the next write.
    throw new Error("Resolve every recipient in Outlook before processing.");
  }
  const email = recipient.emailAddress.trim();
  return recipient.displayName ? `${recipient.displayName} <${email}>` : email;
}
function officeRecipients(values) {
  return values.map((value) => {
    const match = value.trim().match(/^(.*?)\s*<([^<>]*)>$/);
    return match
      ? { displayName: match[1].trim(), emailAddress: match[2].trim() }
      : { displayName: "", emailAddress: value.trim() };
  });
}
async function readRecipients(office, item = office.context.mailbox.item) {
  const lists = await Promise.all(
    FIELDS.map((field) => fieldCall(office, item?.[field], "getAsync"))
  );
  return Object.fromEntries(FIELDS.map((field, i) => [field, lists[i].map(recipientString)]));
}
function sameRecipients(a, b) {
  return FIELDS.every((field) => JSON.stringify(a[field]) === JSON.stringify(b[field]));
}
async function writeRecipients(office, before, after, item = office.context.mailbox.item) {
  const changed = FIELDS.filter(
    (field) => JSON.stringify(before[field]) !== JSON.stringify(after[field])
  );
  if (!changed.length) return;
  for (const field of changed) {
    if (
      !Array.isArray(after[field]) ||
      after[field].length > LIMIT ||
      before[field].length > LIMIT
    ) {
      throw new Error(
        "This action supports up to 100 recipients per changed field. No changes were applied."
      );
    }
  }
  const current = await readRecipients(office, item);
  if (!sameRecipients(current, before))
    throw new Error("Recipients changed in Outlook. Refresh and try again.");
  const attempted = [];
  try {
    for (const field of changed) {
      attempted.push(field);
      await fieldCall(office, item[field], "setAsync", officeRecipients(after[field]));
    }
  } catch (_error) {
    let restored = true;
    for (const field of attempted.reverse()) {
      try {
        await fieldCall(office, item[field], "setAsync", officeRecipients(before[field]));
      } catch (_rollbackError) {
        restored = false;
      }
    }
    // eslint-disable-next-line preserve-caught-error -- Office errors may contain recipient data.
    throw new Error(
      restored
        ? "Outlook could not apply the changes. Previous lists were restored."
        : "Outlook could not restore all recipients. Review To, CC and BCC before sending."
    );
  }
}
module.exports = {
  FIELDS,
  LIMIT,
  readRecipients,
  writeRecipients,
  sameRecipients,
  officeRecipients,
};
