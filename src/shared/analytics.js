const { EVENTS } = require("./analytics-events");
function createAnalytics({
  enabled = false,
  origin = "",
  location = globalThis.location,
  navigator = globalThis.navigator,
  fetch = globalThis.fetch,
} = {}) {
  let consent = enabled === true;
  return {
    setEnabled(value) {
      consent = value === true;
    },
    track(event) {
      // Opt in, exact production origin, no previews, no DNT/GPC, no identifiers.
      if (
        !consent ||
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
