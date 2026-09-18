// Ribbon actions share the task pane's processing and Office I/O; no recipient telemetry.
const { processRecipients } = require("../taskpane/processors");
const { readRecipients, writeRecipients } = require("../shared/recipients");
const {
  normalizeSettings,
  orderedSteps,
  boundedInvalid,
  saveSettings,
} = require("../shared/settings");
const { createAnalytics } = require("../shared/analytics");
let busy = false;
function notify(message, error = false) {
  try {
    const notification = error
      ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message }
      : {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message,
          icon: "Icon.80x80",
          persistent: false,
        };
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "ClearSendNotification",
      notification
    );
  } catch (_error) {
    /* Completion must still reach Outlook. */
  }
}
async function quickClean(event) {
  if (busy) {
    event.completed();
    return;
  }
  busy = true;
  let analytics;
  try {
    const settings = normalizeSettings(Office.context.roamingSettings.get("clearSendSettings"));
    analytics = createAnalytics({
      enabled: settings.analyticsEnabled,
      origin: __ANALYTICS_ORIGIN__,
    });
    analytics.track("quick_clean_click");
    if (!orderedSteps(settings).length) {
      analytics.track("quick_clean_blocked");
      notify("Enable processing options in the ClearSend panel first.", true);
      return;
    }
    const item = Office.context.mailbox.item;
    const before = await readRecipients(Office, item);
    const result = processRecipients({ ...before, userSettings: settings });
    if (settings.keepInvalid) {
      const saved = Office.context.roamingSettings.get("savedInvalidAddresses");
      await saveSettings(
        Office,
        settings,
        boundedInvalid([...(Array.isArray(saved) ? saved : []), ...result.invalid])
      );
    }
    if (!result.success) {
      analytics.track("quick_clean_blocked");
      notify("Address format check blocked processing. No recipients changed.", true);
      return;
    }
    await writeRecipients(Office, before, result.result, item);
    analytics.track("quick_clean_success");
    notify("Recipients processed. Review To, CC and BCC before sending.");
  } catch (error) {
    analytics?.track("quick_clean_error");
    notify(error.message || "Quick clean failed. Review recipients.", true);
  } finally {
    busy = false;
    event.completed();
  }
}
Office.actions.associate("quickClean", quickClean);
Office.actions.associate("action", quickClean);
