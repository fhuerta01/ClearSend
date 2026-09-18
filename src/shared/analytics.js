const { EVENTS } = require("./analytics-events");
function createAnalytics({
  enabled = false,
  origin = "",
  location = globalThis.location,
  navigator = globalThis.navigator,
  fetch = globalThis.fetch,
} = {}) {
  // Stay inactive until the caller has loaded the saved preference.
  let countingEnabled = enabled === true;
  return {
    setEnabled(value) {
      countingEnabled = value === true;
    },
    track(event) {
      // Respect opt-out, exact production origin, DNT/GPC; no identifiers.
      if (
        !countingEnabled ||
        !origin ||
        location?.origin !== origin ||
        !origin.startsWith("https://") ||
        navigator?.doNotTrack === "1" ||
        navigator?.globalPrivacyControl === true ||
        !EVENTS.includes(event) ||
        typeof fetch !== "function"
      )
        return;
      try {
        Promise.resolve(
          fetch("/api/events", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ event }),
            credentials: "omit",
            referrerPolicy: "no-referrer",
            keepalive: true,
            cache: "no-store",
          })
        ).catch(() => {}); // Never retry: an uncertain delivery must not be counted twice.
      } catch (_error) {
        /* Analytics must never interrupt recipient handling. */
      }
    },
  };
}
module.exports = { createAnalytics };
