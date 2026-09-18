// Preferences stay in the user's Microsoft mailbox. No preferences enter analytics.
const STEPS = Object.freeze([
  "sort",
  "dedupe",
  "validate",
  "prioritizeInternal",
  "removeExternal",
  "keepInvalid",
]);
const DEFAULT_STEPS = ["sort", "dedupe", "validate"];
function validDomain(value) {
  return (
    typeof value === "string" &&
    value.length <= 253 &&
    /^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/i.test(value)
  );
}
function normalizeSettings(value = {}) {
  if (!value || typeof value !== "object") value = {};
  const enabledSteps = Array.isArray(value.enabledSteps)
    ? [...new Set(value.enabledSteps.filter((step) => STEPS.includes(step)))]
    : [...DEFAULT_STEPS];
  const order = Array.isArray(value.stepOrder)
    ? value.stepOrder.filter((step) => STEPS.includes(step))
    : [];
  const internalDomains = Array.isArray(value.internalDomains)
    ? [
        ...new Set(
          value.internalDomains
            .filter((domain) => typeof domain === "string")
            .map((domain) => domain.trim().toLowerCase())
            .filter(
              (domain) =>
                domain !== "mydomain.com" && domain !== "newdomain.com" && validDomain(domain)
            )
        ),
      ].slice(0, 3)
    : [];
  return {
    enabledSteps: enabledSteps.filter(
      (step) => internalDomains.length || !["prioritizeInternal", "removeExternal"].includes(step)
    ),
    stepOrder: [...new Set([...order, ...STEPS])],
    internalDomains,
    keepInvalid: value.keepInvalid === true,
    analyticsEnabled: value.analyticsEnabled === true,
  };
}
function orderedSteps(settings) {
  return settings.stepOrder.filter(
    (step) => settings.enabledSteps.includes(step) && step !== "keepInvalid"
  );
}
function saveSettings(office, settings, savedInvalid) {
  return new Promise((resolve, reject) => {
    const store = office.context.roamingSettings;
    store.set("clearSendSettings", normalizeSettings(settings));
    if (savedInvalid !== undefined) {
      if (savedInvalid.length) store.set("savedInvalidAddresses", savedInvalid);
      else store.remove("savedInvalidAddresses");
    }
    store.saveAsync((result) =>
      result.status === office.AsyncResultStatus.Succeeded
        ? resolve()
        : reject(new Error("Settings could not be saved to your Microsoft mailbox."))
    );
  });
}
// Bound mailbox storage (RoamingSettings has a shared 32 KB limit).
function boundedInvalid(values) {
  const result = [];
  const seen = new Set();
  let bytes = 0;
  for (const value of values) {
    if (typeof value !== "string" || value.length > 512 || seen.has(value.toLowerCase())) continue;
    const size = new TextEncoder().encode(JSON.stringify(value)).length + 1;
    if (result.length >= 100 || bytes + size > 12000) break;
    seen.add(value.toLowerCase());
    result.push(value);
    bytes += size;
  }
  return result;
}
module.exports = {
  STEPS,
  normalizeSettings,
  orderedSteps,
  validDomain,
  saveSettings,
  boundedInvalid,
};
