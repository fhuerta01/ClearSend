/**
 * Recipient processing stays inside Outlook. Preferences and optional saved invalid
 * addresses use the user's Microsoft mailbox. Analytics receives fixed event names only.
 */
const {
  processRecipients,
  extractEmail,
  validateEmailFormat,
  getDomainIndex,
} = require("./processors");
const { FIELDS, readRecipients, writeRecipients, sameRecipients } = require("../shared/recipients");
const {
  normalizeSettings,
  orderedSteps,
  validDomain,
  saveSettings,
  boundedInvalid,
} = require("../shared/settings");
const { createAnalytics } = require("../shared/analytics");
const { toCsv } = require("../shared/csv");

const state = {
  settings: normalizeSettings(),
  savedInvalid: [],
  invalid: [],
  recipients: null,
  undo: null,
  busy: false,
  refreshing: false,
  stopped: false,
  poll: null,
  debounce: null,
  saveQueue: Promise.resolve(),
  item: null,
};
const analytics = createAnalytics({ origin: __ANALYTICS_ORIGIN__ });
const $ = (id) => document.getElementById(id);
const labels = {
  sort: "Sort alphabetically",
  dedupe: "Remove duplicates",
  validate: "Check address format",
  prioritizeInternal: "Internal domains first",
  removeExternal: "Remove external",
  keepInvalid: "Save invalid addresses",
};

function toast(message, type = "info") {
  const node = document.createElement("div");
  node.className = `toast ${type}`;
  node.textContent = message;
  $("toastContainer").appendChild(node);
  setTimeout(() => node.remove(), type === "error" ? 10000 : 4000);
}
function actionMessage(message) {
  $("lastActionText").textContent = message;
}
function setBusy(value) {
  state.busy = value;
  $("checkCleanBtn").disabled = value;
  $("undoBtn").disabled = value || !state.undo;
  document.querySelectorAll(".recipient-delete").forEach((button) => {
    button.disabled = value;
  });
}
function persist() {
  const settings = structuredClone(state.settings);
  const saved = [...state.savedInvalid];
  state.saveQueue = state.saveQueue
    .catch(() => {})
    .then(() => saveSettings(Office, settings, saved));
  state.saveQueue.catch(() =>
    toast("Settings were not saved. They may reset when you reopen ClearSend.", "error")
  );
  return state.saveQueue;
}
function syncSettingsUI() {
  for (const step of state.settings.stepOrder) {
    const checkbox = $(step + "Check");
    checkbox.checked =
      step === "keepInvalid"
        ? state.settings.keepInvalid
        : state.settings.enabledSteps.includes(step);
    const domainDependent = ["prioritizeInternal", "removeExternal"].includes(step);
    checkbox.disabled = domainDependent && !state.settings.internalDomains.length;
    checkbox.closest(".feature-item").classList.toggle("disabled", checkbox.disabled);
    $("featureGrid").appendChild(checkbox.closest(".feature-item"));
  }
  $("analyticsCheck").checked = state.settings.analyticsEnabled;
  $("analyticsCheck").disabled = !__ANALYTICS_ORIGIN__ || location.origin !== __ANALYTICS_ORIGIN__;
  $("analyticsAvailability").textContent = $("analyticsCheck").disabled
    ? "Usage counting is disabled for this installation."
    : "Optional: share action counts only. No Outlook data or user identifiers.";
  $("processingOptionsSummary").textContent =
    state.settings.stepOrder
      .filter((step) =>
        step === "keepInvalid"
          ? state.settings.keepInvalid
          : state.settings.enabledSteps.includes(step)
      )
      .map((step) => labels[step])
      .join(", ") || "None";
}
function updateSettings() {
  const items = [...document.querySelectorAll(".feature-item")];
  state.settings = normalizeSettings({
    ...state.settings,
    stepOrder: items.map((item) => item.dataset.step),
    enabledSteps: items
      .filter((item) => item.querySelector("input").checked)
      .map((item) => item.dataset.step),
    keepInvalid: $("keepInvalidCheck").checked,
  });
  syncSettingsUI();
  void persist();
}
function renderDomains() {
  $("domainsList").replaceChildren();
  const domains = state.settings.internalDomains.length ? state.settings.internalDomains : [""];
  domains.forEach((domain, index) => {
    const row = document.createElement("div");
    row.className = "domain-item";
    const input = document.createElement("input");
    input.className = "domain-input";
    input.value = domain;
    input.placeholder = "example.com";
    input.setAttribute("aria-label", `Internal domain ${index + 1}`);
    input.maxLength = 253;
    input.addEventListener("change", () => {
      const next = input.value.trim().toLowerCase();
      if (next && !validDomain(next)) {
        toast("Enter a domain such as example.com.", "error");
        input.value = domain;
        return;
      }
      const updated = [...domains];
      updated[index] = next;
      state.settings = normalizeSettings({ ...state.settings, internalDomains: updated });
      renderDomains();
      syncSettingsUI();
      void persist();
      void refresh();
    });
    const remove = button("−", "Remove domain", () => {
      state.settings.internalDomains.splice(index, 1);
      state.settings = normalizeSettings(state.settings);
      renderDomains();
      syncSettingsUI();
      void persist();
      void refresh();
    });
    remove.className = "domain-btn";
    row.append(input, remove);
    $("domainsList").appendChild(row);
  });
  if (domains.length < 3 && domains.every(Boolean)) {
    const add = button("+ Add domain", "Add domain", () => {
      // An empty editable row is never treated as an internal domain.
      state.settings.internalDomains.push("");
      renderDomains();
    });
    add.className = "config-item-btn";
    $("domainsList").appendChild(add);
  }
}
function button(text, title, handler) {
  const node = document.createElement("button");
  node.type = "button";
  node.textContent = text;
  node.title = title;
  node.setAttribute("aria-label", title);
  node.addEventListener("click", () => {
    Promise.resolve()
      .then(handler)
      .catch((error) => toast(error.message, "error"));
  });
  return node;
}
function emailStatus(address) {
  const email = extractEmail(address);
  if (!validateEmailFormat(email).isValid)
    return { text: "Unrecognized address format", color: "red" };
  if (!state.settings.internalDomains.length)
    return { text: "Format checked; internal domains not configured", color: null };
  return getDomainIndex(email, state.settings.internalDomains) >= 0
    ? { text: "Internal domain; format checked", color: null }
    : { text: "External domain; format checked", color: "orange" };
}
function renderList(field, addresses, counts) {
  const container = $(field + "Content");
  container.replaceChildren();
  const toggle = $("toggle" + field[0].toUpperCase() + field.slice(1) + "Btn");
  toggle.disabled = !addresses.length;
  if (!addresses.length) {
    container.style.display = "none";
    toggle.setAttribute("aria-expanded", "false");
  }
  addresses.forEach((address, index) => {
    const row = document.createElement("div");
    row.className = "recipient-item";
    const status = emailStatus(address);
    const duplicate = counts.get(extractEmail(address).toLowerCase()) > 1;
    const dot = document.createElement("span");
    dot.className = status.color
      ? `recipient-circle ${status.color}`
      : "recipient-circle-placeholder";
    const label = document.createElement("span");
    label.className = "recipient-email";
    label.textContent = address;
    label.title = `${address} — ${status.text}${duplicate ? "; Duplicate" : ""}`;
    const copy = button("Copy", "Copy email address", async () => {
      analytics.track("copy_click");
      await navigator.clipboard.writeText(extractEmail(address));
      toast("Email copied to clipboard.");
    });
    copy.className = "recipient-copy";
    row.append(dot, label, copy);
    if (FIELDS.includes(field) || field === "savedInvalid") {
      const remove = button(
        "×",
        field === "savedInvalid" ? "Delete saved address" : "Remove recipient",
        async () => {
          if (state.busy) return;
          if (field === "savedInvalid") {
            state.savedInvalid.splice(index, 1);
            await persist();
            await refresh();
            return;
          }
          analytics.track("remove_click");
          await mutate(async () => {
            const before = await readRecipients(Office, state.item);
            if (!sameRecipients(before, state.recipients))
              throw new Error("Recipients changed. Refresh and try again.");
            const after = structuredClone(before);
            after[field].splice(index, 1);
            await writeRecipients(Office, before, after, state.item);
            state.undo = { before, after };
            actionMessage("Recipient removed. You can undo this change.");
          });
        }
      );
      remove.className = "recipient-delete";
      remove.disabled = state.busy;
      row.appendChild(remove);
    }
    container.appendChild(row);
  });
}
async function refresh() {
  if (state.refreshing || state.stopped || state.busy) return;
  state.refreshing = true;
  try {
    const recipients = await readRecipients(Office, state.item);
    state.recipients = recipients;
    const all = FIELDS.flatMap((field) => recipients[field]);
    const counts = new Map();
    for (const address of all) {
      const key = extractEmail(address).toLowerCase();
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    const unique = [
      ...new Map(all.map((value) => [extractEmail(value).toLowerCase(), value])).values(),
    ];
    state.invalid = unique.filter((value) => !validateEmailFormat(extractEmail(value)).isValid);
    $("totalDestinations").textContent = all.length;
    $("duplicatedCount").textContent = [...counts.values()].reduce(
      (sum, count) => sum + Math.max(0, count - 1),
      0
    );
    $("externalCount").textContent = state.settings.internalDomains.length
      ? unique.filter(
          (value) =>
            validateEmailFormat(extractEmail(value)).isValid &&
            getDomainIndex(extractEmail(value), state.settings.internalDomains) < 0
        ).length
      : "—";
    $("invalidCount").textContent = state.invalid.length;
    for (const field of FIELDS) {
      $(field + "Count").textContent = `(${recipients[field].length})`;
      renderList(field, recipients[field], counts);
    }
    $("invalidAddressCount").textContent = state.invalid.length;
    $("savedInvalidAddressCount").textContent = state.savedInvalid.length;
    renderList("invalid", state.invalid, counts);
    renderList("savedInvalid", state.savedInvalid, counts);
    $("downloadBtn").disabled = !all.length;
    $("downloadInvalidBtn").disabled = !state.invalid.length && !state.savedInvalid.length;
  } finally {
    state.refreshing = false;
  }
}
async function mutate(operation) {
  if (state.busy) return;
  setBusy(true);
  try {
    await operation();
  } catch (error) {
    actionMessage(error.message);
    throw error;
  } finally {
    setBusy(false);
    try {
      await refresh();
    } catch (_error) {
      toast("Could not refresh the display. Review recipients in Outlook.", "error");
    }
  }
}
async function clean(override) {
  if (state.busy) return;
  analytics.track("process_click");
  await mutate(async () => {
    try {
      const settings = override
        ? normalizeSettings({ ...state.settings, enabledSteps: override })
        : state.settings;
      if (!orderedSteps(settings).length) {
        analytics.track("process_blocked");
        toast("Enable a processing option first.", "warning");
        return;
      }
      const before = await readRecipients(Office, state.item);
      const result = processRecipients({ ...before, userSettings: settings });
      if (state.settings.keepInvalid) {
        state.savedInvalid = boundedInvalid([...state.savedInvalid, ...result.invalid]);
        await persist();
      }
      if (!result.success) {
        analytics.track("process_blocked");
        actionMessage("No changes applied: review unrecognized address formats in Outlook.");
        toast("Address format check blocked processing. No recipients were removed.", "warning");
        return;
      }
      const after = result.result;
      if (sameRecipients(before, after)) {
        actionMessage("No changes needed.");
      } else {
        await writeRecipients(Office, before, after, state.item);
        state.undo = { before, after };
        actionMessage("Recipient lists updated. Review them before sending; Undo is available.");
      }
      analytics.track("process_success");
      toast("Processing completed. Review the destination fields before sending.", "success");
    } catch (error) {
      analytics.track("process_error");
      throw error;
    }
  });
}
function bind(id, event, handler) {
  $(id).addEventListener(event, () => {
    Promise.resolve()
      .then(handler)
      .catch((error) => toast(error.message, "error"));
  });
}
function showTab(name) {
  for (const tab of ["details", "config"]) {
    $(tab + "Content").classList.toggle("active", tab === name);
    $(tab + "Footer").style.display = tab === name ? "flex" : "none";
  }
}
function download(name, lists) {
  const blob = new Blob(["\ufeff" + toCsv(lists)], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function setupHandlers() {
  bind("checkCleanBtn", "click", () => clean());
  bind("summaryRefreshBtn", "click", () => {
    analytics.track("refresh_click");
    return refresh();
  });
  bind("processingOptionsBtn", "click", () => {
    analytics.track("settings_click");
    showTab("config");
  });
  bind("configBackBtn", "click", () => showTab("details"));
  bind("undoBtn", "click", async () => {
    if (!state.undo || state.busy) return;
    analytics.track("undo_click");
    await mutate(async () => {
      await writeRecipients(Office, state.undo.after, state.undo.before, state.item);
      state.undo = null;
      actionMessage("Previous recipient lists restored.");
    });
  });
  bind("downloadBtn", "click", async () => {
    analytics.track("export_click");
    const recipients = await readRecipients(Office, state.item);
    download(
      "recipients.csv",
      FIELDS.map((field) => recipients[field])
    );
  });
  bind("downloadInvalidBtn", "click", () => {
    analytics.track("export_invalid_click");
    download("invalid_addresses.csv", [state.invalid, state.savedInvalid]);
  });
  for (const field of [...FIELDS, "invalid", "savedInvalid"]) {
    const id = "toggle" + field[0].toUpperCase() + field.slice(1) + "Btn";
    bind(id, "click", () => {
      const open = $(field + "Content").style.display === "none";
      $(field + "Content").style.display = open ? "block" : "none";
      $(id).setAttribute("aria-expanded", String(open));
    });
  }
  for (const step of state.settings.stepOrder) bind(step + "Check", "change", updateSettings);
  bind("analyticsCheck", "change", async () => {
    state.settings.analyticsEnabled = $("analyticsCheck").checked;
    analytics.setEnabled(false);
    try {
      await persist();
      analytics.setEnabled(state.settings.analyticsEnabled);
    } catch (_error) {
      state.settings.analyticsEnabled = false;
      $("analyticsCheck").checked = false;
    }
  });
  bind("clearSavedInvalidBtn", "click", async () => {
    state.savedInvalid = [];
    await persist();
    await refresh();
  });
  bind("restoreDefaultsBtn", "click", async () => {
    state.settings = normalizeSettings();
    state.savedInvalid = [];
    analytics.setEnabled(false);
    await persist();
    syncSettingsUI();
    renderDomains();
    await refresh();
    toast("Defaults restored; saved invalid addresses deleted.");
  });
  document.addEventListener("keydown", (event) => {
    if (!event.ctrlKey || !event.altKey || event.repeat) return;
    const shortcuts = {
      KeyQ: () => clean(),
      KeyS: () => clean(["sort"]),
      KeyD: () => clean(["dedupe"]),
      KeyV: async () => {
        await refresh();
        toast(`${state.invalid.length} unrecognized address formats. No recipients changed.`);
      },
    };
    if (shortcuts[event.code]) {
      event.preventDefault();
      Promise.resolve()
        .then(shortcuts[event.code])
        .catch((error) => toast(error.message, "error"));
    }
  });
  let dragged = null;
  document.querySelectorAll(".feature-item").forEach((item) => {
    item.draggable = true;
    const handle = item.querySelector(".drag-handle");
    handle.setAttribute("aria-label", "Move option up (Shift: move down)");
    handle.addEventListener("click", (event) => {
      const sibling = event.shiftKey ? item.nextElementSibling : item.previousElementSibling;
      if (sibling) {
        if (event.shiftKey) sibling.after(item);
        else sibling.before(item);
        updateSettings();
      }
    });
    item.addEventListener("dragstart", (event) => {
      dragged = item;
      event.dataTransfer.setData("text/plain", item.dataset.step);
      item.classList.add("dragging");
    });
    item.addEventListener("dragover", (event) => {
      if (!dragged || dragged === item) return;
      event.preventDefault();
      const box = item.getBoundingClientRect();
      if (event.clientY < box.top + box.height / 2) item.before(dragged);
      else item.after(dragged);
    });
    item.addEventListener("dragend", () => {
      item.classList.remove("dragging");
      dragged = null;
      updateSettings();
    });
  });
}
function recipientChanged() {
  clearTimeout(state.debounce);
  state.debounce = setTimeout(() => {
    refresh().catch(() => {});
  }, 300);
}
function startPolling() {
  if (state.stopped || state.poll) return;
  state.poll = setInterval(() => {
    if (!document.hidden) refresh().catch(() => {});
  }, 2000);
}
async function initialize() {
  if (state.item) return;
  state.item = Office.context.mailbox.item;
  if (!state.item?.to?.setAsync) throw new Error("Open ClearSend while composing an email.");
  state.settings = normalizeSettings(Office.context.roamingSettings.get("clearSendSettings"));
  const saved = Office.context.roamingSettings.get("savedInvalidAddresses");
  state.savedInvalid = boundedInvalid(Array.isArray(saved) ? saved : []);
  analytics.setEnabled(state.settings.analyticsEnabled);
  setupHandlers();
  syncSettingsUI();
  renderDomains();
  await refresh();
  setBusy(false);
  analytics.track("pane_open");
  if (state.item.addHandlerAsync && Office.EventType.RecipientsChanged) {
    state.item.addHandlerAsync(Office.EventType.RecipientsChanged, recipientChanged, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) startPolling();
    });
  } else startPolling();
}
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook)
    initialize().catch((error) => {
      $("checkCleanBtn").disabled = true;
      actionMessage(error.message);
      toast(error.message, "error");
    });
});
window.addEventListener("pagehide", () => {
  state.stopped = true;
  clearInterval(state.poll);
  clearTimeout(state.debounce);
  if (state.item?.removeHandlerAsync && Office.EventType.RecipientsChanged) {
    state.item.removeHandlerAsync(Office.EventType.RecipientsChanged, {
      handler: recipientChanged,
    });
  }
});
