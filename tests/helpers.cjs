const { officeRecipients } = require("../src/shared/recipients");
function officeMock(initial = { to: [], cc: [], bcc: [] }, settings = {}) {
  const lists = structuredClone(initial);
  const writes = [];
  const values = { clearSendSettings: settings };
  let failField = null;
  let alwaysFail = false;
  const item = {
    itemType: "message",
    notificationMessages: { replaceAsync() {} },
    addHandlerAsync(_event, handler, cb) {
      item.changed = handler;
      cb({ status: "succeeded" });
    },
    removeHandlerAsync() {},
  };
  for (const field of ["to", "cc", "bcc"]) {
    item[field] = {
      getAsync(cb) {
        cb({ status: "succeeded", value: officeRecipients(lists[field]) });
      },
      setAsync(value, cb) {
        writes.push(field);
        if (field === failField) {
          if (!alwaysFail) failField = null;
          cb({ status: "failed" });
          return;
        }
        lists[field] = value.map((r) =>
          r.displayName ? `${r.displayName} <${r.emailAddress}>` : r.emailAddress
        );
        cb({ status: "succeeded" });
      },
    };
  }
  const office = {
    AsyncResultStatus: { Succeeded: "succeeded" },
    HostType: { Outlook: "outlook" },
    EventType: { RecipientsChanged: "recipientsChanged" },
    MailboxEnums: {
      ItemNotificationMessageType: { InformationalMessage: "info", ErrorMessage: "error" },
    },
    context: {
      mailbox: { item },
      roamingSettings: {
        get(key) {
          return values[key];
        },
        set(key, value) {
          values[key] = value;
        },
        remove(key) {
          delete values[key];
        },
        saveAsync(cb) {
          cb({ status: "succeeded" });
        },
      },
    },
    onReady(cb) {
      cb({ host: "outlook" });
    },
    actions: { associate() {} },
  };
  return {
    office,
    lists,
    writes,
    values,
    fail(field, persistent = false) {
      failField = field;
      alwaysFail = persistent;
    },
  };
}
module.exports = { officeMock };
