/**
 * ClearSend Task Pane JavaScript
 *
 * PRIVACY GUARANTEE: All processing happens locally in your Outlook client.
 * Your email addresses NEVER leave your device. No servers process your data.
 *
 * This file only uses Office.js API to read/write recipients locally.
 * No network requests transmit any email or recipient data.
 */

/* global Office, document, window, setTimeout, setInterval, clearTimeout, clearInterval, Blob, URL */

/**
 * Configuration Constants
 * Centralized configuration to avoid magic numbers and improve maintainability
 */
const CONFIG = {
  // Timing constants
  RECIPIENT_DEBOUNCE_MS: 500, // Delay before processing recipient changes
  RECIPIENT_POLLING_INTERVAL_MS: 2000, // How often to poll for recipient changes
  TOAST_DURATION_MS: 3000, // How long to show toast notifications
  REFRESH_DELAY_MS: 300, // Small delay for Outlook to process changes
  PROGRESS_ANIMATION_DELAY_MS: 100, // Progress bar animation delay

  // Network constants
  NETWORK_TIMEOUT_MS: 10000, // Network request timeout (10 seconds)
  MAX_RETRY_ATTEMPTS: 3, // Maximum number of retry attempts
  RETRY_BASE_DELAY_MS: 1000, // Base delay for exponential backoff

  // Limits
  MAX_INTERNAL_DOMAINS: 3, // Maximum number of internal domains
  MIN_INTERNAL_DOMAINS: 1, // Minimum number of internal domains
  MAX_RECIPIENTS_PER_FIELD: 500, // Maximum recipients allowed per field (To/CC/BCC) by Outlook

  // Validation
  MIN_EMAIL_LENGTH: 3, // Minimum email length to validate
  EMAIL_REGEX:
    /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/,
};

/**
 * Global toggle function for collapsible sections
 * Available immediately for HTML onclick handlers
 * @param {string} contentId - ID of the content element to toggle
 * @param {HTMLElement} headerElement - Header element containing the collapse button
 */
window.toggleSection = function (contentId, headerElement) {
  const content = document.getElementById(contentId);
  const btn = headerElement?.querySelector(".collapse-btn");

  if (content) {
    const currentDisplay = window.getComputedStyle(content).display;
    const hasCollapsedClass = content.classList.contains("collapsed");
    const isHidden = currentDisplay === "none" || hasCollapsedClass;

    if (isHidden) {
      content.style.display = "block";
      content.classList.remove("collapsed");
      if (btn)
        btn.innerHTML =
          '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="12" height="12" fill="currentColor"><path d="M233.4 105.4c12.5-12.5 32.8-12.5 45.3 0l192 192c12.5 12.5 12.5 32.8 0 45.3s-32.8 12.5-45.3 0L256 173.3 86.6 342.6c-12.5 12.5-32.8 12.5-45.3 0s-12.5-32.8 0-45.3l192-192z"/></svg>';
    } else {
      content.style.display = "none";
      content.classList.add("collapsed");
      if (btn)
        btn.innerHTML =
          '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="12" height="12" fill="currentColor"><path d="M233.4 406.6c12.5 12.5 32.8 12.5 45.3 0l192-192c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0L256 338.7 86.6 169.4c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3l192 192z"/></svg>';
    }
  } else {
  }
};

/**
 * ClearSend Application State
 * Global state object managing the add-in's runtime state
 */
const ClearSend = {
  isInitialized: false,
  currentUser: null,
  settings: {
    enabledSteps: ["sort", "dedupe", "validate", "prioritizeInternal"],
    stepOrder: [
      "sort",
      "dedupe",
      "validate",
      "prioritizeInternal",
      "removeExternal",
      "keepInvalid",
    ],
    orgDomain: "",
    internalDomains: ["mydomain.com"],
    keepInvalid: false,
  },
  cache: new Map(),
  recipientChangeTimeout: null, // Timeout handle for debouncing recipient changes
  recipientPollingInterval: null, // Interval handle for polling recipient changes
  lastRecipientState: null, // Saved state for undo functionality
  lastActionMessage: null, // Message describing the last action performed
  isUpdatingDisplay: false, // Flag to prevent concurrent display updates (race condition protection)
  eventHandlers: new Map(), // Store event handler references for proper cleanup
  invalidAddresses: [], // Store detected invalid addresses (current)
  savedInvalidAddresses: [], // Store saved invalid addresses (persisted)
};

// Initialize ClearSend when Office is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initializeClearSend();
    // Setup cleanup handler to prevent memory leaks
    setupCleanupHandlers();
  }
});

/**
 * Setup cleanup handlers to prevent memory leaks
 * Clears all intervals, timeouts, and event listeners when the window unloads
 */
function setupCleanupHandlers() {
  const cleanup = () => {
    try {
      // Clear polling interval
      if (ClearSend.recipientPollingInterval) {
        clearInterval(ClearSend.recipientPollingInterval);
        ClearSend.recipientPollingInterval = null;
      }

      // Clear debounce timeout
      if (ClearSend.recipientChangeTimeout) {
        clearTimeout(ClearSend.recipientChangeTimeout);
        ClearSend.recipientChangeTimeout = null;
      }

      // Remove all tracked event handlers
      ClearSend.eventHandlers.forEach((handler, element) => {
        try {
          if (element && handler) {
            element.removeEventListener(handler.event, handler.callback);
          }
        } catch (error) {}
      });
      ClearSend.eventHandlers.clear();
    } catch (error) {}
  };

  // Listen for window unload events
  window.addEventListener("beforeunload", cleanup);
  window.addEventListener("unload", cleanup);
}

/**
 * Initialize ClearSend application
 * Sets up event handlers, loads settings, and initializes the UI
 */
async function initializeClearSend() {
  try {
    // Load user settings (this will also render domains if they exist)
    await loadUserSettings();

    // Setup event handlers
    setupEventHandlers();

    // Setup Office.js recipient change listeners
    setupRecipientChangeListeners();

    // Check if we're in compose mode
    if (
      Office.context.mailbox.item &&
      Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message
    ) {
      await updateRecipientDisplay();
    }

    // Update processing options summary
    updateProcessingOptionsSummary();

    // Update domain-dependent features state
    updateDomainDependentFeatures();

    ClearSend.isInitialized = true;
  } catch (error) {
    showToast("Failed to initialize ClearSend", "error");
  }
}

/**
 * Safe event handler wrapper
 * Wraps event handlers to catch and log errors gracefully
 * @param {Function} handler - The event handler function
 * @param {string} handlerName - Name of the handler for logging
 * @returns {Function} Wrapped handler function
 */
function safeEventHandler(handler, handlerName = "handler") {
  return async function (...args) {
    try {
      await handler.apply(this, args);
    } catch (error) {
      showToast(`An error occurred in ${handlerName}`, "error");
    }
  };
}

/**
 * Office.js Recipient Change Listeners Setup
 * Attempts to use event-based listeners with fallback to polling
 */
function setupRecipientChangeListeners() {
  try {
    if (!Office.context.mailbox.item) {
      return;
    }

    // Try event-based listeners first
    const item = Office.context.mailbox.item;
    let eventListenersSetup = false;

    try {
      // Check if event listeners are supported
      if (
        item.to &&
        item.to.addHandlerAsync &&
        Office.EventType &&
        Office.EventType.RecipientsChanged
      ) {
        // Listen for To recipients changes
        item.to.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          onRecipientsChanged,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              eventListenersSetup = true;
            } else {
              startRecipientPolling();
            }
          }
        );
        // Listen for CC recipients changes
        item.cc.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          onRecipientsChanged,
          (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              startRecipientPolling();
            }
          }
        );
        // Listen for BCC recipients changes
        item.bcc.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          onRecipientsChanged,
          (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              startRecipientPolling();
            }
          }
        );
      }
    } catch (eventError) {
      startRecipientPolling();
    }

    // Fallback to polling-based approach if event listeners not setup after a delay
    setTimeout(() => {
      if (!eventListenersSetup) {
        startRecipientPolling();
      }
    }, 1000);
  } catch (error) {
    // Fallback to polling
    startRecipientPolling();
  }
}

/**
 * Polling-based recipient change detection
 * Used as fallback when event-based listeners are not supported
 */
function startRecipientPolling() {
  // Clear existing interval if any
  if (ClearSend.recipientPollingInterval) {
    clearInterval(ClearSend.recipientPollingInterval);
    ClearSend.recipientPollingInterval = null;
  }

  let lastRecipientHash = "";

  const checkForChanges = async () => {
    try {
      const recipients = await getCurrentRecipients();
      const currentHash = JSON.stringify({
        to: recipients.to.sort(),
        cc: recipients.cc.sort(),
        bcc: recipients.bcc.sort(),
      });

      // Only trigger update if hash changed (and we have a baseline)
      if (currentHash !== lastRecipientHash && lastRecipientHash !== "") {
        onRecipientsChanged({ source: "polling" });
      }

      lastRecipientHash = currentHash;
    } catch (error) {}
  };

  // Check every CONFIG.RECIPIENT_POLLING_INTERVAL_MS
  ClearSend.recipientPollingInterval = setInterval(
    checkForChanges,
    CONFIG.RECIPIENT_POLLING_INTERVAL_MS
  );

  // Initial check to set baseline
  checkForChanges();
}

/**
 * Handle recipient changes with debouncing and race condition protection
 * @param {Object} eventArgs - Event arguments from Office.js or polling
 */
async function onRecipientsChanged(eventArgs) {
  // Debounce rapid changes - clear previous timeout
  if (ClearSend.recipientChangeTimeout) {
    clearTimeout(ClearSend.recipientChangeTimeout);
    ClearSend.recipientChangeTimeout = null;
  }

  // Set new timeout using CONFIG constant
  ClearSend.recipientChangeTimeout = setTimeout(async () => {
    // Race condition protection - check if already updating
    if (ClearSend.isUpdatingDisplay) {
      return;
    }

    try {
      await updateRecipientDisplay();
    } catch (error) {}
  }, CONFIG.RECIPIENT_DEBOUNCE_MS);
}

// Event Handlers Setup
function setupEventHandlers() {
  // Tab navigation (only if elements exist - may not exist if tab navigation is hidden)
  const detailsTab = document.getElementById("detailsTab");
  const configTab = document.getElementById("configTab");
  if (detailsTab) detailsTab.addEventListener("click", () => switchTab("details"));
  if (configTab) configTab.addEventListener("click", () => switchTab("config"));

  // Action buttons
  document.getElementById("summaryRefreshBtn").addEventListener("click", handleRefresh);
  document.getElementById("undoBtn").addEventListener("click", handleUndo);
  document.getElementById("downloadBtn").addEventListener("click", handleDownloadCSV);

  // Field toggle buttons
  document.getElementById("toggleToBtn").addEventListener("click", () => toggleField("to"));
  document.getElementById("toggleCcBtn").addEventListener("click", () => toggleField("cc"));
  document.getElementById("toggleBccBtn").addEventListener("click", () => toggleField("bcc"));

  // Collapsible sections will be set up in initializeClearSend

  // Settings
  document.getElementById("sortCheck").addEventListener("change", updateSettings);
  document.getElementById("dedupeCheck").addEventListener("change", updateSettings);
  document.getElementById("validateCheck").addEventListener("change", updateSettings);
  document.getElementById("prioritizeInternalCheck").addEventListener("change", updateSettings);
  document.getElementById("removeExternalCheck").addEventListener("change", updateSettings);
  document.getElementById("keepInvalidCheck").addEventListener("change", updateSettings);

  // Initialize drag-and-drop for feature items
  initializeDragAndDrop();

  // Download invalid addresses
  document.getElementById("downloadInvalidBtn").addEventListener("click", handleDownloadInvalidCSV);

  // Configuration - Restore defaults button
  document.getElementById("restoreDefaultsBtn").addEventListener("click", handleRestoreDefaults);

  // Footer - Check and clean button
  document.getElementById("checkCleanBtn").addEventListener("click", handleClean);

  // Keyboard shortcuts
  document.addEventListener("keydown", handleKeyboardShortcuts);
}

// Keyboard Shortcuts Handler
function handleKeyboardShortcuts(event) {
  if (event.ctrlKey && event.altKey) {
    switch (event.code) {
      case "KeyS":
        event.preventDefault();
        handleSort();
        break;
      case "KeyD":
        event.preventDefault();
        handleDedupe();
        break;
      case "KeyV":
        event.preventDefault();
        handleValidate();
        break;
      case "KeyC":
        event.preventDefault();
        // Toggle task pane (handled by Office.js)
        break;
    }
  }
}

// Core Action Handlers
async function handleSort() {
  try {
    showProgress("Sorting recipients...");

    const recipients = await getCurrentRecipients();

    const result = await callOrchestrator(recipients, ["sort"]);

    if (result.status === "success") {
      try {
        const originalCount = result.to.length + result.cc.length + result.bcc.length;
        await updateRecipients(result);

        // Check if recipients were filtered by getting current count
        const currentRecipients = await getCurrentRecipients();
        const currentCount =
          currentRecipients.to.length + currentRecipients.cc.length + currentRecipients.bcc.length;
        const filteredByOutlook = originalCount - currentCount;

        const sortAction = result.actions.find((a) => a.type === "sort");

        if (filteredByOutlook > 0) {
          showToast(
            `Sorted ${sortAction?.processed || 0} recipients (${filteredByOutlook} invalid entries removed by Outlook)`,
            "warning"
          );
        } else {
          showToast(`Sorted ${sortAction?.processed || 0} recipients`, "success");
        }

        // Update recipient analysis display
        await updateRecipientDisplay();
      } catch (updateError) {
        showToast(
          "Sort completed but failed to update recipients: " + updateError.message,
          "warning"
        );

        // Still update the display with what we have
        try {
          await updateRecipientDisplay();
        } catch (displayError) {}
      }
    } else {
      showToast("Sort failed: " + result.error, "error");
    }
  } catch (error) {
    showToast("Sort failed: " + error.message, "error");
  } finally {
    hideProgress();
  }
}

async function handleDedupe() {
  try {
    showProgress("Removing duplicates...");

    const recipients = await getCurrentRecipients();

    const result = await callOrchestrator(recipients, ["dedupe"]);

    if (result.status === "success") {
      try {
        const originalCount = result.to.length + result.cc.length + result.bcc.length;
        await updateRecipients(result);

        // Check if recipients were filtered by getting current count
        const currentRecipients = await getCurrentRecipients();
        const currentCount =
          currentRecipients.to.length + currentRecipients.cc.length + currentRecipients.bcc.length;
        const filteredByOutlook = originalCount - currentCount;

        const dedupeAction = result.actions.find((a) => a.type === "dedupe");

        if (filteredByOutlook > 0) {
          showToast(
            `Removed ${dedupeAction?.duplicatesFound || 0} duplicates (${filteredByOutlook} invalid entries removed by Outlook)`,
            "warning"
          );
        } else {
          showToast(`Removed ${dedupeAction?.duplicatesFound || 0} duplicates`, "success");
        }

        // Update recipient analysis display
        await updateRecipientDisplay();
      } catch (updateError) {
        showToast(
          "Deduplication completed but failed to update recipients: " + updateError.message,
          "warning"
        );

        try {
          await updateRecipientDisplay();
        } catch (displayError) {}
      }
    } else {
      showToast("Deduplication failed: " + result.error, "error");
    }
  } catch (error) {
    showToast("Deduplication failed: " + error.message, "error");
  } finally {
    hideProgress();
  }
}

async function handleValidate() {
  try {
    showProgress("Validating emails...");

    const recipients = await getCurrentRecipients();

    const result = await callOrchestrator(recipients, ["validate"]);

    if (result.status === "success") {
      try {
        const originalCount = result.to.length + result.cc.length + result.bcc.length;
        const validateAction = result.actions.find((a) => a.type === "validate");

        // Update recipients in Outlook with the validated list
        await updateRecipients(result);

        // Check if recipients were filtered by getting current count
        const currentRecipients = await getCurrentRecipients();
        const currentCount =
          currentRecipients.to.length + currentRecipients.cc.length + currentRecipients.bcc.length;
        const filteredByOutlook = originalCount - currentCount;

        // Note: displayValidationResults() creates a different UI format
        // We'll rely on updateRecipientDisplay() for consistent formatting
        // But still update status to show validation info
        updateValidationStatusOnly(validateAction);

        const removedByValidation = validateAction?.errorCount || 0;

        if (filteredByOutlook > 0) {
          showToast(
            `Validated ${validateAction?.processed || 0} recipients: removed ${removedByValidation} invalid (${filteredByOutlook} additional entries removed by Outlook)`,
            "warning"
          );
        } else {
          showToast(
            `Validated ${validateAction?.processed || 0} recipients: removed ${removedByValidation} invalid`,
            "success"
          );
        }

        // Update recipient analysis display
        await updateRecipientDisplay();
      } catch (updateError) {
        showToast(
          "Validation completed but failed to update recipients: " + updateError.message,
          "warning"
        );

        // Still update validation status
        try {
          const validateAction = result.actions.find((a) => a.type === "validate");
          updateValidationStatusOnly(validateAction);
        } catch (displayError) {}
      }
    } else {
      showToast("Validation failed: " + result.error, "error");
    }
  } catch (error) {
    showToast("Validation failed: " + error.message, "error");
  } finally {
    hideProgress();
  }
}

async function handleClean() {
  try {
    // Disable the clean button to prevent multiple clicks
    const cleanBtn = document.getElementById("cleanBtn");
    if (cleanBtn) {
      cleanBtn.disabled = true;
    }

    // Get current recipients
    const recipients = await getCurrentRecipients();

    // Save current state for undo
    saveRecipientState(recipients);
    const totalOriginal = recipients.to.length + recipients.cc.length + recipients.bcc.length;

    // Convert to Office.js format for processing
    const toRecipients = convertToOfficeFormat(recipients.to);
    const ccRecipients = convertToOfficeFormat(recipients.cc);
    const bccRecipients = convertToOfficeFormat(recipients.bcc);

    let processedTo = toRecipients;
    let processedCc = ccRecipients;
    let processedBcc = bccRecipients;

    const enabledSteps = getEnabledSteps();

    if (enabledSteps.length === 0) {
      showToast(
        "No cleaning steps are enabled. Please enable some features in Configuration tab.",
        "warning"
      );
      if (cleanBtn) cleanBtn.disabled = false;
      return;
    }

    let invalidCount = 0;
    let duplicateCount = 0;
    let externalRemovedCount = 0;

    const internalDomains = getValidInternalDomains();
    const sortAlphabetically = enabledSteps.includes("sort");

    // Clear invalid addresses array at start
    ClearSend.invalidAddresses = [];

    // Save invalid addresses if keepInvalid is enabled (do this BEFORE any processing)
    if (ClearSend.settings.keepInvalid) {
      const allRecipients = [...processedTo, ...processedCc, ...processedBcc];
      const invalidRecipients = allRecipients.filter(
        (recipient) => !isValidEmail(recipient.emailAddress || "")
      );

      if (invalidRecipients.length > 0) {
        const formattedInvalids = invalidRecipients.map((r) => {
          if (r.displayName && r.displayName !== r.emailAddress) {
            return `${r.displayName} <${r.emailAddress}>`;
          }
          return r.emailAddress;
        });

        // Add to saved list, avoiding duplicates
        formattedInvalids.forEach((invalid) => {
          const normalized = invalid.toLowerCase().trim();
          if (
            !ClearSend.savedInvalidAddresses.some(
              (saved) => saved.toLowerCase().trim() === normalized
            )
          ) {
            ClearSend.savedInvalidAddresses.push(invalid);
          }
        });

        // Save to roaming storage
        saveSavedInvalidAddresses();
      }
    }

    // Step 1: Validate (if enabled) - this prevents processing if invalids are found
    if (enabledSteps.includes("validate")) {
      // Detect invalid addresses
      const allRecipients = [...processedTo, ...processedCc, ...processedBcc];
      const invalidRecipients = allRecipients.filter(
        (recipient) => !isValidEmail(recipient.emailAddress || "")
      );

      // Abort processing if there are invalid addresses (prevent invalids processing)
      if (invalidRecipients.length > 0) {
        invalidCount = invalidRecipients.length;

        // Show toast and abort processing
        showToast(
          "Processing of addresses disabled due to invalid addresses in the lists",
          "error"
        );
        updateRecipientDisplay(recipients);
        if (cleanBtn) cleanBtn.disabled = false;
        return;
      }

      // No invalid addresses found, proceed with validation (this will be a no-op since all are valid)
      const validatedTo = validateRecipients(processedTo);
      const validatedCc = validateRecipients(processedCc);
      const validatedBcc = validateRecipients(processedBcc);

      invalidCount =
        processedTo.length +
        processedCc.length +
        processedBcc.length -
        (validatedTo.length + validatedCc.length + validatedBcc.length);

      processedTo = validatedTo;
      processedCc = validatedCc;
      processedBcc = validatedBcc;
    }

    // Step 2: Deduplicate (if enabled)
    if (enabledSteps.includes("dedupe")) {
      const beforeDedupe = processedTo.length + processedCc.length + processedBcc.length;
      const deduped = deduplicateRecipients(processedTo, processedCc, processedBcc);

      processedTo = deduped.to;
      processedCc = deduped.cc;
      processedBcc = deduped.bcc;

      const afterDedupe = processedTo.length + processedCc.length + processedBcc.length;
      duplicateCount = beforeDedupe - afterDedupe;
    }

    // Step 3: Remove External (if enabled)
    if (enabledSteps.includes("removeExternal")) {
      const toResult = removeExternalRecipients(processedTo, internalDomains);
      const ccResult = removeExternalRecipients(processedCc, internalDomains);
      const bccResult = removeExternalRecipients(processedBcc, internalDomains);

      processedTo = toResult.filtered;
      processedCc = ccResult.filtered;
      processedBcc = bccResult.filtered;

      externalRemovedCount = toResult.removed + ccResult.removed + bccResult.removed;
    }

    // Step 4: Prioritize Internal (if enabled)
    if (enabledSteps.includes("prioritizeInternal")) {
      processedTo = prioritizeInternalRecipients(processedTo, internalDomains, sortAlphabetically);
      processedCc = prioritizeInternalRecipients(processedCc, internalDomains, sortAlphabetically);
      processedBcc = prioritizeInternalRecipients(
        processedBcc,
        internalDomains,
        sortAlphabetically
      );
    } else if (enabledSteps.includes("sort")) {
      // Step 5: Sort (if enabled and prioritize is not enabled)
      processedTo = sortRecipients(processedTo);
      processedCc = sortRecipients(processedCc);
      processedBcc = sortRecipients(processedBcc);
    }

    // Update recipients in Outlook
    await updateRecipientsDirectly(processedTo, processedCc, processedBcc);

    const totalFinal = processedTo.length + processedCc.length + processedBcc.length;

    // Show success toast with processed count
    showToast(`Success. ${totalOriginal} addresses processed.`, "success");

    // Build last action message (without "Cleaned: X recipients" prefix)
    let lastActionMessage = "";
    const changes = [];
    if (invalidCount > 0) changes.push(`${invalidCount} invalid removed`);
    if (duplicateCount > 0) changes.push(`${duplicateCount} duplicates removed`);
    if (externalRemovedCount > 0) changes.push(`${externalRemovedCount} external removed`);

    if (changes.length > 0) {
      lastActionMessage = changes.join(", ");
      // Update last action message and enable undo
      updateLastAction(lastActionMessage);
    } else {
      lastActionMessage = "No changes applied";
      // Update last action message but disable undo (no changes to revert)
      updateLastAction(lastActionMessage, false);
    }

    // Update recipient analysis display
    await updateRecipientDisplay();
  } catch (error) {
    // Check if error is due to recipient limit
    if (error.message === "RECIPIENT_LIMIT_EXCEEDED") {
      showToast("Exceeded the configured limit of destination addresses in a field", "error");
    } else {
      showToast("Clean failed: " + error.message, "error");
    }
  } finally {
    // Re-enable the clean button
    const cleanBtn = document.getElementById("cleanBtn");
    if (cleanBtn) {
      cleanBtn.disabled = false;
    }
  }
}

async function handleOnlyInternal() {
  showToast("Only Internal feature not yet configured", "warning");
  // TODO: Implement domain filtering once domain configuration is available
}

async function handleRefresh() {
  try {
    // Add a small delay to allow Outlook to process any recent changes
    await new Promise((resolve) => setTimeout(resolve, 300));

    await updateRecipientDisplay();

    // Show success toast
    showToast("Recipients information refreshed", "success");
  } catch (error) {
    showToast("Failed to refresh: " + error.message, "error");
  }
}

function handleCopyRecipient(event) {
  try {
    const button = event.currentTarget;
    const address = decodeURIComponent(button.getAttribute("data-address"));

    // Extract just the email address if it's in format "Name <email@domain.com>"
    let emailToCopy = address;
    const match = address.match(/<(.+)>/);
    if (match) {
      emailToCopy = match[1];
    }

    // Use the fallback method (works in all browsers and environments)
    const textArea = document.createElement("textarea");
    textArea.value = emailToCopy;
    textArea.style.position = "fixed";
    textArea.style.left = "-999999px";
    textArea.style.top = "0";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();

    try {
      const successful = document.execCommand("copy");
      if (successful) {
        showToast("Email copied to clipboard", "success");
      } else {
        showToast("Failed to copy email", "error");
      }
    } catch (err) {
      showToast("Failed to copy email", "error");
    } finally {
      document.body.removeChild(textArea);
    }
  } catch (error) {
    showToast("Failed to copy: " + error.message, "error");
  }
}

async function handleRemoveRecipient(event) {
  try {
    const button = event.currentTarget;
    const field = button.getAttribute("data-field");
    const addressToRemove = decodeURIComponent(button.getAttribute("data-address"));

    // Get current recipients and save state for undo
    const recipients = await getCurrentRecipients();
    ClearSend.lastRecipientState = {
      to: [...recipients.to],
      cc: [...recipients.cc],
      bcc: [...recipients.bcc],
    };

    // Convert to Office.js format
    const toRecipients = convertToOfficeFormat(recipients.to);
    const ccRecipients = convertToOfficeFormat(recipients.cc);
    const bccRecipients = convertToOfficeFormat(recipients.bcc);

    // Remove only the first occurrence of the recipient from the appropriate field
    let updatedTo = toRecipients;
    let updatedCc = ccRecipients;
    let updatedBcc = bccRecipients;

    // Helper function to remove first occurrence only
    const removeFirstOccurrence = (recipients, targetAddress) => {
      let found = false;
      return recipients.filter((r) => {
        if (found) return true; // Keep all after first match

        const displayEmail = r.displayName
          ? `${r.displayName} <${r.emailAddress}>`
          : r.emailAddress;
        const matches = displayEmail === targetAddress || r.emailAddress === targetAddress;

        if (matches) {
          found = true;
          return false; // Remove this one
        }
        return true; // Keep this one
      });
    };

    if (field === "to") {
      updatedTo = removeFirstOccurrence(toRecipients, addressToRemove);
    } else if (field === "cc") {
      updatedCc = removeFirstOccurrence(ccRecipients, addressToRemove);
    } else if (field === "bcc") {
      updatedBcc = removeFirstOccurrence(bccRecipients, addressToRemove);
    }

    // Update recipients in Outlook
    await updateRecipientsDirectly(updatedTo, updatedCc, updatedBcc);

    // Update last action message with undo enabled
    updateLastAction(`Removed ${addressToRemove} manually`, true);

    // Refresh the display
    await updateRecipientDisplay();

    showToast("Recipient removed", "success");
  } catch (error) {
    showToast("Failed to remove recipient: " + error.message, "error");
  }
}

/**
 * Helper Functions for Local Processing
 */

/**
 * Convert recipient strings to Office.js format
 * Handles both "Display Name <email@domain.com>" and "email@domain.com" formats
 * @param {Array} recipients - Array of recipient strings
 * @returns {Array} Array of recipient objects with displayName and emailAddress
 */
function convertToOfficeFormat(recipients) {
  if (!Array.isArray(recipients)) {
    return [];
  }

  return recipients.map((r) => {
    if (typeof r === "string") {
      const match = r.match(/^(.+?)\s*<(.+)>$/);
      if (match) {
        return {
          displayName: match[1].trim(),
          emailAddress: match[2].trim(),
        };
      }
      return {
        displayName: "",
        emailAddress: r.trim(),
      };
    }
    // Already in Office format
    return r;
  });
}

/**
 * Validate recipients using email format rules
 * Filters out invalid email addresses
 * @param {Array} recipients - Array of recipient objects
 * @returns {Array} Array of valid recipients only
 */
function isValidEmail(email) {
  // Basic validation checks
  if (!email || typeof email !== "string") return false;
  if (!email.includes("@")) return false;
  if (email.split("@").length !== 2) return false;

  const [localPart, domainPart] = email.split("@");
  if (!localPart || !domainPart) return false;
  if (localPart.length === 0 || domainPart.length === 0) return false;
  if (!domainPart.includes(".")) return false;
  if (email.includes("..")) return false; // No consecutive dots

  // Use CONFIG email regex for validation
  return CONFIG.EMAIL_REGEX.test(email);
}

function validateRecipients(recipients) {
  if (!Array.isArray(recipients)) {
    return [];
  }

  return recipients.filter((recipient) => {
    const email = recipient.emailAddress || "";
    return isValidEmail(email);
  });
}

function deduplicateRecipients(toRecipients, ccRecipients, bccRecipients) {
  const uniqueTo = deduplicateArray(toRecipients);
  const uniqueCc = deduplicateArray(ccRecipients);
  const uniqueBcc = deduplicateArray(bccRecipients);

  const allEmails = new Set();
  const finalTo = [];
  const finalCc = [];
  const finalBcc = [];

  uniqueTo.forEach((recipient) => {
    const email = (recipient.emailAddress || "").toLowerCase();
    if (!allEmails.has(email)) {
      allEmails.add(email);
      finalTo.push(recipient);
    }
  });

  uniqueCc.forEach((recipient) => {
    const email = (recipient.emailAddress || "").toLowerCase();
    if (!allEmails.has(email)) {
      allEmails.add(email);
      finalCc.push(recipient);
    }
  });

  uniqueBcc.forEach((recipient) => {
    const email = (recipient.emailAddress || "").toLowerCase();
    if (!allEmails.has(email)) {
      allEmails.add(email);
      finalBcc.push(recipient);
    }
  });

  return { to: finalTo, cc: finalCc, bcc: finalBcc };
}

function deduplicateArray(recipients) {
  const seen = new Set();
  return recipients.filter((recipient) => {
    const email = (recipient.emailAddress || "").toLowerCase();
    if (seen.has(email)) return false;
    seen.add(email);
    return true;
  });
}

function sortRecipients(recipients) {
  return recipients.sort((a, b) => {
    const nameA = (a.displayName || a.emailAddress || "").toLowerCase();
    const nameB = (b.displayName || b.emailAddress || "").toLowerCase();
    return nameA.localeCompare(nameB);
  });
}

/**
 * Escape HTML special characters to prevent XSS
 * @param {string} unsafe - Unsafe string that may contain HTML
 * @returns {string} Escaped safe string
 */
function escapeHtml(unsafe) {
  if (!unsafe || typeof unsafe !== "string") {
    return "";
  }
  return unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * Get the domain index for an email address
 * Returns the index of the matching internal domain, or -1 if external
 * @param {string} email - Email address to check
 * @param {Array} internalDomains - Array of internal domain strings
 * @returns {number} Index of matching domain, or -1 if not found
 */
function getDomainIndex(email, internalDomains) {
  // Input validation
  if (!email || typeof email !== "string" || !email.includes("@")) {
    return -1;
  }
  if (!Array.isArray(internalDomains) || internalDomains.length === 0) {
    return -1;
  }

  const parts = email.split("@");
  if (parts.length !== 2) {
    return -1;
  }

  const domain = parts[1].toLowerCase().trim();
  if (!domain) {
    return -1;
  }

  for (let i = 0; i < internalDomains.length; i++) {
    const internalDomain = (internalDomains[i] || "").toLowerCase().trim();
    if (!internalDomain) {
      continue;
    }
    // Exact match or subdomain match
    if (domain === internalDomain || domain.endsWith("." + internalDomain)) {
      return i;
    }
  }

  return -1;
}

function prioritizeInternalRecipients(recipients, internalDomains, sortAlphabetically = false) {
  if (!recipients || recipients.length === 0) {
    return recipients;
  }

  if (!internalDomains || internalDomains.length === 0) {
    return recipients;
  }

  // Separate internal and external recipients
  const internal = [];
  const external = [];

  recipients.forEach((recipient) => {
    const email = recipient.emailAddress || "";
    const domainIndex = getDomainIndex(email, internalDomains);

    if (domainIndex >= 0) {
      internal.push({ recipient, domainIndex, email });
    } else {
      external.push({ recipient, email });
    }
  });

  // Sort internal by domain index
  internal.sort((a, b) => {
    if (a.domainIndex !== b.domainIndex) {
      return a.domainIndex - b.domainIndex;
    }

    // If same domain and alphabetical sorting is enabled, sort alphabetically
    if (sortAlphabetically) {
      const nameA = (a.recipient.displayName || a.email).toLowerCase();
      const nameB = (b.recipient.displayName || b.email).toLowerCase();
      return nameA.localeCompare(nameB);
    }

    return 0;
  });

  // Sort external alphabetically if enabled
  if (sortAlphabetically) {
    external.sort((a, b) => {
      const nameA = (a.recipient.displayName || a.email).toLowerCase();
      const nameB = (b.recipient.displayName || b.email).toLowerCase();
      return nameA.localeCompare(nameB);
    });
  }

  // Combine: internal first, then external
  return [...internal.map((item) => item.recipient), ...external.map((item) => item.recipient)];
}

function removeExternalRecipients(recipients, internalDomains) {
  if (!recipients || recipients.length === 0) {
    return { filtered: [], removed: 0 };
  }

  if (!internalDomains || internalDomains.length === 0) {
    return { filtered: [...recipients], removed: 0 };
  }

  const filtered = [];
  let removedCount = 0;

  recipients.forEach((recipient) => {
    const email = recipient.emailAddress || "";
    const domainIndex = getDomainIndex(email, internalDomains);

    if (domainIndex >= 0) {
      // Internal domain - keep it
      filtered.push(recipient);
    } else {
      // External domain - remove it
      removedCount++;
    }
  });

  return { filtered, removed: removedCount };
}

/**
 * Promisified helper to set recipients for a specific field
 * @param {Object} field - Office.js recipient field (to, cc, or bcc)
 * @param {Array} recipients - Array of recipient objects
 * @param {string} fieldName - Name of the field for error messages
 * @returns {Promise<void>}
 */
function setRecipientsAsync(field, recipients, fieldName) {
  return new Promise((resolve, reject) => {
    if (!field || !field.setAsync) {
      reject(new Error(`Invalid field: ${fieldName}`));
      return;
    }

    // Validate input
    if (!Array.isArray(recipients)) {
      reject(new Error(`Recipients for ${fieldName} must be an array`));
      return;
    }

    field.setAsync(recipients, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        const errorMsg = result.error?.message || "Unknown error";

        // Check if error is due to recipient limit
        if (
          errorMsg.includes("ArgumentOutOfRangeException") ||
          errorMsg.includes("out of the range") ||
          errorMsg.includes("Specified argument was out of the range")
        ) {
          reject(new Error("RECIPIENT_LIMIT_EXCEEDED"));
          return;
        }

        reject(
          new Error(
            `Failed to update ${fieldName} recipients (${recipients.length} addresses): ${errorMsg}`
          )
        );
        return;
      }
      resolve();
    });
  });
}

/**
 * Update recipients directly using Office.js
 * Uses Promise.all for better performance and error handling
 * @param {Array} toRecipients - Array of To recipient objects
 * @param {Array} ccRecipients - Array of CC recipient objects
 * @param {Array} bccRecipients - Array of BCC recipient objects
 * @returns {Promise<void>}
 */
async function updateRecipientsDirectly(toRecipients, ccRecipients, bccRecipients) {
  try {
    const item = Office.context.mailbox.item;

    if (!item) {
      throw new Error("No mailbox item available");
    }

    // Validate inputs
    if (
      !Array.isArray(toRecipients) ||
      !Array.isArray(ccRecipients) ||
      !Array.isArray(bccRecipients)
    ) {
      throw new Error("All recipient parameters must be arrays");
    }

    // Update all fields in parallel for better performance
    await Promise.all([
      setRecipientsAsync(item.to, toRecipients, "TO"),
      setRecipientsAsync(item.cc, ccRecipients, "CC"),
      setRecipientsAsync(item.bcc, bccRecipients, "BCC"),
    ]);
  } catch (error) {
    throw error;
  }
}

/**
 * Promisified helper to get recipients from a specific field
 * @param {Object} field - Office.js recipient field (to, cc, or bcc)
 * @param {string} fieldName - Name of the field for error messages
 * @returns {Promise<Array>} Array of recipient strings
 */
function getRecipientsAsync(field, fieldName) {
  return new Promise((resolve, reject) => {
    if (!field || !field.getAsync) {
      reject(new Error(`Invalid field: ${fieldName}`));
      return;
    }

    field.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(
          new Error(
            `Failed to get ${fieldName} recipients: ${result.error?.message || "Unknown error"}`
          )
        );
        return;
      }

      // Safely convert recipients with robust error handling
      const safeConvertRecipients = (recipientList) => {
        if (!Array.isArray(recipientList)) {
          return [];
        }

        return recipientList
          .map((r, index) => {
            try {
              if (!r || typeof r !== "object") {
                return null;
              }

              const displayName = (r.displayName || "").trim();
              const emailAddress = (r.emailAddress || "").trim();

              // Skip completely empty entries
              if (!emailAddress) {
                return null;
              }

              return displayName ? `${displayName} <${emailAddress}>` : emailAddress;
            } catch (error) {
              return null;
            }
          })
          .filter((r) => r !== null);
      };

      resolve(safeConvertRecipients(result.value));
    });
  });
}

/**
 * Get current recipients from all fields (To, CC, BCC)
 * Uses Promise.all for better performance and error handling
 * @returns {Promise<Object>} Object containing to, cc, and bcc arrays
 */
async function getCurrentRecipients() {
  try {
    // Check if Office.js is available
    if (!Office || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
      // Return test data for debugging
      return {
        to: ["John Doe <john@example.com>", "alice@company.com", "john@example.com"],
        cc: ["peter@gmail.com", "mary@invalid"],
        bcc: ["tom@example.com", "@invalid.com"],
      };
    }

    const item = Office.context.mailbox.item;

    // Use Promise.all to get all recipients in parallel for better performance
    const [to, cc, bcc] = await Promise.all([
      getRecipientsAsync(item.to, "TO"),
      getRecipientsAsync(item.cc, "CC"),
      getRecipientsAsync(item.bcc, "BCC"),
    ]);

    return { to, cc, bcc };
  } catch (error) {
    throw error;
  }
}

async function updateRecipients(result) {
  try {
    const item = Office.context.mailbox.item;

    // Convert back to Office.js format with robust validation
    const convertToRecipients = (recipients) => {
      return recipients
        .filter((r) => r && typeof r === "string" && r.trim().length > 0) // Filter out empty/invalid
        .map((r) => {
          try {
            r = r.trim();

            // Handle display name + email format: "Name <email>"
            const match = r.match(/^(.+?)\s*<(.+)>$/);
            if (match) {
              const displayName = match[1].trim();
              const emailAddress = match[2].trim();

              // Validate email has basic structure
              if (emailAddress && emailAddress.length > 0) {
                return {
                  displayName: displayName,
                  emailAddress: emailAddress,
                };
              }
            }

            // Handle cases like "SJK:" or other malformed inputs
            // If it doesn't look like an email, try to sanitize or skip
            if (r.includes("@") || r.length > 3) {
              return {
                displayName: "",
                emailAddress: r,
              };
            }

            // Skip obviously malformed entries like "SJK:"

            return null;
          } catch (error) {
            return null;
          }
        })
        .filter((r) => r !== null); // Remove null entries
    };

    return new Promise((resolve, reject) => {
      try {
        const toRecipients = convertToRecipients(result.to);
        const ccRecipients = convertToRecipients(result.cc);
        const bccRecipients = convertToRecipients(result.bcc);

        // Count how many entries might be filtered by Outlook
        const originalCount = result.to.length + result.cc.length + result.bcc.length;
        const convertedCount = toRecipients.length + ccRecipients.length + bccRecipients.length;
        const filteredCount = originalCount - convertedCount;

        // Update To recipients
        item.to.setAsync(toRecipients, (toResult) => {
          if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
            reject(
              new Error(
                "Failed to update To recipients: " + (toResult.error?.message || "Unknown error")
              )
            );
            return;
          }

          // Update CC recipients
          item.cc.setAsync(ccRecipients, (ccResult) => {
            if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
              reject(
                new Error(
                  "Failed to update CC recipients: " + (ccResult.error?.message || "Unknown error")
                )
              );
              return;
            }

            // Update BCC recipients
            item.bcc.setAsync(bccRecipients, (bccResult) => {
              if (bccResult.status !== Office.AsyncResultStatus.Succeeded) {
                reject(
                  new Error(
                    "Failed to update BCC recipients: " +
                      (bccResult.error?.message || "Unknown error")
                  )
                );
                return;
              }

              resolve();
            });
          });
        });
      } catch (error) {
        reject(new Error("Failed to convert recipients: " + error.message));
      }
    });
  } catch (error) {
    throw new Error("Failed to update recipients: " + error.message);
  }
}

/**
 * API Integration - Call orchestrator with timeout and retry logic
 * Implements exponential backoff retry strategy and AbortController timeout
 * @param {Object} recipients - Recipients object with to, cc, bcc arrays
 * @param {Array} enabledSteps - Array of enabled processing steps
 * @param {number} attempt - Current retry attempt (used internally for recursion)
 * @returns {Promise<Object>} Processing result from API
 */
async function callOrchestrator(recipients, enabledSteps, attempt = 1) {
  // Validate inputs
  if (!recipients || typeof recipients !== "object") {
    throw new Error("Invalid recipients parameter");
  }
  if (
    !Array.isArray(recipients.to) ||
    !Array.isArray(recipients.cc) ||
    !Array.isArray(recipients.bcc)
  ) {
    throw new Error("Recipients must contain to, cc, and bcc arrays");
  }
  if (!Array.isArray(enabledSteps)) {
    throw new Error("Enabled steps must be an array");
  }

  const payload = {
    to: recipients.to,
    cc: recipients.cc,
    bcc: recipients.bcc,
    userSettings: {
      enabledSteps: enabledSteps,
      internalDomains: getValidInternalDomains(),
      orgDomain: ClearSend.settings.orgDomain || "",
    },
  };

  try {
    // Process recipients client-side using the processors library
    const result = window.ClearSendProcessors.processRecipients(payload);
    return result;
  } catch (error) {
    throw new Error(`Processing failed: ${error.message}`);
  }
}

/**
 * UI Helper Functions
 */

/**
 * Show progress indicator with animated bar
 * @param {string} message - Message to display
 */
function showProgress(message) {
  const container = document.getElementById("progressContainer");
  const text = document.getElementById("progressText");
  const fill = document.getElementById("progressFill");

  if (!container || !text || !fill) {
    return;
  }

  text.textContent = message;
  fill.style.width = "0%";
  container.style.display = "block";

  // Animate progress bar using CONFIG constant
  setTimeout(() => (fill.style.width = "100%"), CONFIG.PROGRESS_ANIMATION_DELAY_MS);
}

/**
 * Hide progress indicator
 */
function hideProgress() {
  const container = document.getElementById("progressContainer");
  if (container) {
    container.style.display = "none";
  }
}

/**
 * Show loading overlay with message
 * @param {string} message - Loading message to display
 */
function showLoadingOverlay(message) {
  const overlay = document.getElementById("loadingOverlay");
  if (!overlay) {
    return;
  }

  const text = overlay.querySelector(".loading-text");
  if (text) {
    text.textContent = message;
  }
  overlay.style.display = "flex";
}

/**
 * Hide loading overlay
 */
function hideLoadingOverlay() {
  const overlay = document.getElementById("loadingOverlay");
  if (overlay) {
    overlay.style.display = "none";
  }
}

/**
 * Show toast notification
 * Prevents duplicate messages from appearing simultaneously
 * @param {string} message - Message to display
 * @param {string} type - Toast type: 'info', 'success', 'warning', or 'error'
 */
function showToast(message, type = "info") {
  const container = document.getElementById("toastContainer");
  if (!container) {
    return;
  }

  // Sanitize message to prevent XSS
  const sanitizedMessage = escapeHtml(message);

  // Check for duplicate toasts with the same message
  const existingToasts = container.querySelectorAll(".toast");
  for (const existingToast of existingToasts) {
    if (existingToast.textContent === sanitizedMessage) {
      return; // Don't show duplicate
    }
  }

  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = sanitizedMessage;

  container.appendChild(toast);

  // Auto remove after CONFIG.TOAST_DURATION_MS
  setTimeout(() => {
    if (toast.parentNode) {
      toast.parentNode.removeChild(toast);
    }
  }, CONFIG.TOAST_DURATION_MS);
}

// Settings Management
async function loadUserSettings() {
  try {
    // Office.js roamingSettings.get() is synchronous, not async
    const savedSettings = Office.context.roamingSettings.get("clearSendSettings");
    const savedInvalidAddresses = Office.context.roamingSettings.get("savedInvalidAddresses");

    if (savedSettings && typeof savedSettings === "object") {
      // Merge saved settings with defaults
      ClearSend.settings = {
        ...ClearSend.settings,
        ...savedSettings,
      };

      // Update UI to reflect loaded settings
      updateSettingsUI();
    } else {
      // Save default settings
      saveSettings();
    }

    // Load saved invalid addresses if they exist
    if (savedInvalidAddresses && Array.isArray(savedInvalidAddresses)) {
      ClearSend.savedInvalidAddresses = savedInvalidAddresses;
    } else {
      // Initialize as empty array if not found
      ClearSend.savedInvalidAddresses = [];
    }

    // Always render internal domains (either from loaded settings or defaults)
    renderInternalDomains();
  } catch (error) {
    // Reset to defaults on any error
    ClearSend.savedInvalidAddresses = [];
    ClearSend.settings = {
      enabledSteps: ["sort", "dedupe", "validate", "prioritizeInternal"],
      stepOrder: [
        "sort",
        "dedupe",
        "validate",
        "prioritizeInternal",
        "removeExternal",
        "keepInvalid",
      ],
      orgDomain: "",
      internalDomains: ["mydomain.com"],
      keepInvalid: false,
    };
    renderInternalDomains();
    saveSettings();
  }
}

function updateSettings() {
  const enabledSteps = [];

  // Get the order from the DOM (which reflects user's drag-and-drop arrangement)
  const featureItems = document.querySelectorAll(".feature-item");
  const stepOrder = [];

  featureItems.forEach((item) => {
    const step = item.getAttribute("data-step");
    if (step) {
      stepOrder.push(step);
      const checkbox = item.querySelector('input[type="checkbox"]');
      if (checkbox && checkbox.checked) {
        enabledSteps.push(step);
      }
    }
  });

  ClearSend.settings.enabledSteps = enabledSteps;
  ClearSend.settings.stepOrder = stepOrder;
  ClearSend.settings.keepInvalid = document.getElementById("keepInvalidCheck").checked;

  // Save all settings to roaming storage
  saveSettings();
}

function updateSettingsUI() {
  // Restore checkbox states
  document.getElementById("sortCheck").checked = ClearSend.settings.enabledSteps.includes("sort");
  document.getElementById("dedupeCheck").checked =
    ClearSend.settings.enabledSteps.includes("dedupe");
  document.getElementById("validateCheck").checked =
    ClearSend.settings.enabledSteps.includes("validate");
  document.getElementById("prioritizeInternalCheck").checked =
    ClearSend.settings.enabledSteps.includes("prioritizeInternal");
  document.getElementById("removeExternalCheck").checked =
    ClearSend.settings.enabledSteps.includes("removeExternal");
  document.getElementById("keepInvalidCheck").checked = ClearSend.settings.keepInvalid || false;

  // Restore the order of feature items if saved
  if (ClearSend.settings.stepOrder && ClearSend.settings.stepOrder.length > 0) {
    restoreFeatureOrder();
  }
}

function restoreFeatureOrder() {
  const featureGrid = document.getElementById("featureGrid");
  if (!featureGrid) return;

  const stepOrder = ClearSend.settings.stepOrder;
  const featureItems = Array.from(featureGrid.querySelectorAll(".feature-item"));

  // Create a map of step -> element
  const itemMap = new Map();
  featureItems.forEach((item) => {
    const step = item.getAttribute("data-step");
    if (step) {
      itemMap.set(step, item);
    }
  });

  // Reorder based on saved order
  stepOrder.forEach((step) => {
    const item = itemMap.get(step);
    if (item) {
      featureGrid.appendChild(item);
    }
  });
}

// Drag and Drop functionality
function initializeDragAndDrop() {
  const featureGrid = document.getElementById("featureGrid");
  if (!featureGrid) return;

  const featureItems = featureGrid.querySelectorAll(".feature-item");
  let draggedItem = null;

  featureItems.forEach((item) => {
    const checkbox = item.querySelector('input[type="checkbox"]');

    // Prevent checkbox from initiating drag
    checkbox.addEventListener("mousedown", (e) => {
      e.stopPropagation();
    });

    checkbox.addEventListener("click", (e) => {
      e.stopPropagation();
    });

    // Make entire item draggable except checkbox
    item.setAttribute("draggable", "true");

    item.addEventListener("dragstart", (e) => {
      // Prevent drag if clicking on checkbox
      if (e.target === checkbox) {
        e.preventDefault();
        return;
      }

      draggedItem = item;
      item.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/html", item.innerHTML);
    });

    item.addEventListener("dragend", () => {
      item.classList.remove("dragging");

      // Remove drag-over class from all items
      featureItems.forEach((i) => i.classList.remove("drag-over"));

      // Save the new order
      updateSettings();
    });

    item.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";

      const afterElement = getDragAfterElement(featureGrid, e.clientY);
      if (afterElement == null) {
        featureGrid.appendChild(draggedItem);
      } else {
        featureGrid.insertBefore(draggedItem, afterElement);
      }
    });

    item.addEventListener("dragenter", (e) => {
      if (item !== draggedItem) {
        item.classList.add("drag-over");
      }
    });

    item.addEventListener("dragleave", () => {
      item.classList.remove("drag-over");
    });
  });
}

function getDragAfterElement(container, y) {
  const draggableElements = [...container.querySelectorAll(".feature-item:not(.dragging)")];

  return draggableElements.reduce(
    (closest, child) => {
      const box = child.getBoundingClientRect();
      const offset = y - box.top - box.height / 2;

      if (offset < 0 && offset > closest.offset) {
        return { offset: offset, element: child };
      } else {
        return closest;
      }
    },
    { offset: Number.NEGATIVE_INFINITY }
  ).element;
}

// Tab Management with slide animations
function switchTab(tabName) {
  const detailsContent = document.getElementById("detailsContent");
  const configContent = document.getElementById("configContent");
  const detailsFooter = document.getElementById("detailsFooter");
  const configFooter = document.getElementById("configFooter");

  // Determine current active tab
  const currentTab = detailsContent.classList.contains("active") ? "details" : "config";

  // Don't animate if already on the target tab
  if (currentTab === tabName) return;

  // Update tab buttons (only if they exist)
  const tabBtn = document.getElementById(tabName + "Tab");
  if (tabBtn) {
    document.querySelectorAll(".tab-btn").forEach((btn) => {
      btn.classList.remove("active");
    });
    tabBtn.classList.add("active");
  }

  // Determine animation direction
  if (tabName === "config") {
    // Going to config: details slides out left, config slides in from right
    detailsContent.classList.add("slide-out-left");
    configContent.classList.add("slide-in-left");
    configContent.style.display = "block";

    // After animation completes
    setTimeout(() => {
      detailsContent.classList.remove("active", "slide-out-left");
      detailsContent.style.display = "none";
      configContent.classList.add("active");
      configContent.classList.remove("slide-in-left");

      // Update footers
      if (detailsFooter) detailsFooter.style.display = "none";
      if (configFooter) configFooter.style.display = "flex";
    }, 300);
  } else if (tabName === "details") {
    // Going to details: config slides out right, details slides in from left
    configContent.classList.add("slide-out-right");
    detailsContent.classList.add("slide-in-right");
    detailsContent.style.display = "block";

    // After animation completes
    setTimeout(() => {
      configContent.classList.remove("active", "slide-out-right");
      configContent.style.display = "none";
      detailsContent.classList.add("active");
      detailsContent.classList.remove("slide-in-right");

      // Update footers
      if (detailsFooter) detailsFooter.style.display = "flex";
      if (configFooter) configFooter.style.display = "none";
    }, 300);
  }
}

// Collapsible Sections Management (DISABLED - using HTML onclick instead)
function setupCollapsibleSections() {
  // This function is disabled since we're using HTML onclick handlers
}

function expandSection(sectionName) {
  const content = document.getElementById(sectionName + "Content");
  const btn = content.previousElementSibling.querySelector(".collapse-btn");

  if (content && content.classList.contains("collapsed")) {
    content.classList.remove("collapsed");
    btn.innerHTML =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="12" height="12" fill="currentColor"><path d="M233.4 105.4c12.5-12.5 32.8-12.5 45.3 0l192 192c12.5 12.5 12.5 32.8 0 45.3s-32.8 12.5-45.3 0L256 173.3 86.6 342.6c-12.5 12.5-32.8 12.5-45.3 0s-12.5-32.8 0-45.3l192-192z"/></svg>';
  }
}

function collapseSection(sectionName) {
  const content = document.getElementById(sectionName + "Content");
  const btn = content.previousElementSibling.querySelector(".collapse-btn");

  if (content && !content.classList.contains("collapsed")) {
    content.classList.add("collapsed");
    btn.innerHTML =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="12" height="12" fill="currentColor"><path d="M233.4 406.6c12.5 12.5 32.8 12.5 45.3 0l192-192c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0L256 338.7 86.6 169.4c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3l192 192z"/></svg>';
  }
}

// Utility Functions
function getEnabledSteps() {
  // Use stepOrder if available, otherwise fall back to enabledSteps
  const orderedSteps = ClearSend.settings.stepOrder || ClearSend.settings.enabledSteps;

  // Filter to only include steps that are enabled (checked)
  const result = orderedSteps.filter((step) => {
    const checkboxId = step + "Check";
    const checkbox = document.getElementById(checkboxId);
    return checkbox && checkbox.checked;
  });

  return result;
}

/**
 * Update recipient display with race condition protection
 * Shows recipient counts, statistics, and lists
 */
async function updateRecipientDisplay() {
  // Race condition protection - check if already updating
  if (ClearSend.isUpdatingDisplay) {
    return;
  }

  // Set lock
  ClearSend.isUpdatingDisplay = true;

  try {
    const recipients = await getCurrentRecipients();

    const list = document.getElementById("recipientList");

    const totalRecipients = recipients.to.length + recipients.cc.length + recipients.bcc.length;

    // Calculate summary statistics
    const allRecipients = [...recipients.to, ...recipients.cc, ...recipients.bcc];
    let duplicatedCount = 0;
    let externalCount = 0;
    let invalidCount = 0;

    // Count unique addresses that have duplicates
    const emailCounts = {};
    allRecipients.forEach((address) => {
      const normalizedEmail = address.toLowerCase().trim();
      emailCounts[normalizedEmail] = (emailCounts[normalizedEmail] || 0) + 1;
    });
    duplicatedCount = Object.values(emailCounts).filter((count) => count > 1).length;

    // Count unique external and invalid addresses
    const uniqueExternal = new Set();
    const uniqueInvalid = new Set();
    const invalidAddressesList = [];
    allRecipients.forEach((address) => {
      const normalizedEmail = address.toLowerCase().trim();
      const status = getEmailStatus(address);
      if (status.status === "invalid") {
        uniqueInvalid.add(normalizedEmail);
        // Only add unique invalids to the list
        if (!invalidAddressesList.some((addr) => addr.toLowerCase().trim() === normalizedEmail)) {
          invalidAddressesList.push(address);
        }
      } else if (status.status === "external") {
        uniqueExternal.add(normalizedEmail);
      }
    });
    externalCount = uniqueExternal.size;
    invalidCount = uniqueInvalid.size;

    // Update invalid addresses array for the Invalid destinations section
    ClearSend.invalidAddresses = invalidAddressesList;

    // Update total destinations
    document.getElementById("totalDestinations").textContent = totalRecipients;

    // Update summary stats
    document.getElementById("duplicatedCount").textContent = duplicatedCount;
    document.getElementById("externalCount").textContent = externalCount;
    document.getElementById("invalidCount").textContent = invalidCount;

    // Update stat labels with conditional circles
    const duplicatedLabel = document.querySelector("#duplicatedCount + .stat-label");
    const externalLabel = document.querySelector("#externalCount + .stat-label");
    const invalidLabel = document.querySelector("#invalidCount + .stat-label");

    // Orange circle for duplicated if count > 0
    if (duplicatedLabel) {
      if (duplicatedCount > 0) {
        duplicatedLabel.innerHTML = `<span class="stat-circle orange"></span>Duplicated`;
      } else {
        duplicatedLabel.textContent = "Duplicated";
      }
    }

    // Red/orange circle for external based on removeExternal setting
    if (externalLabel) {
      const removeExternalEnabled = document.getElementById("removeExternal")?.checked || false;
      if (removeExternalEnabled) {
        externalLabel.innerHTML = `<span class="stat-circle red"></span>External`;
      } else if (externalCount > 0) {
        externalLabel.innerHTML = `<span class="stat-circle orange"></span>External`;
      } else {
        externalLabel.textContent = "External";
      }
    }

    // Red circle for invalid if count > 0
    if (invalidLabel) {
      if (invalidCount > 0) {
        invalidLabel.innerHTML = `<span class="stat-circle red"></span>Invalid`;
      } else {
        invalidLabel.textContent = "Invalid";
      }
    }

    // Update field counts
    document.getElementById("toCount").textContent = `(${recipients.to.length})`;
    document.getElementById("ccCount").textContent = `(${recipients.cc.length})`;
    document.getElementById("bccCount").textContent = `(${recipients.bcc.length})`;

    // Enable/disable toggle buttons based on whether field has addresses
    document.getElementById("toggleToBtn").disabled = recipients.to.length === 0;
    document.getElementById("toggleCcBtn").disabled = recipients.cc.length === 0;
    document.getElementById("toggleBccBtn").disabled = recipients.bcc.length === 0;

    // Enable/disable download button
    const downloadBtn = document.getElementById("downloadBtn");
    if (downloadBtn) {
      downloadBtn.disabled = totalRecipients === 0;
    }

    if (totalRecipients === 0) {
      document.getElementById("totalDestinations").textContent = "0";
      document.getElementById("duplicatedCount").textContent = "0";
      document.getElementById("externalCount").textContent = "0";
      document.getElementById("invalidCount").textContent = "0";
      document.getElementById("toCount").textContent = "(0)";
      document.getElementById("ccCount").textContent = "(0)";
      document.getElementById("bccCount").textContent = "(0)";
      document.getElementById("toggleToBtn").disabled = true;
      document.getElementById("toggleCcBtn").disabled = true;
      document.getElementById("toggleBccBtn").disabled = true;
    }

    // Populate field content areas (always call, even if empty, to clear stale data)
    populateFieldContent("to", recipients.to);
    populateFieldContent("cc", recipients.cc);
    populateFieldContent("bcc", recipients.bcc);

    // Update invalid addresses display
    const invalidAddressCount =
      (ClearSend.invalidAddresses && ClearSend.invalidAddresses.length) || 0;
    const savedInvalidAddressCount =
      (ClearSend.savedInvalidAddresses && ClearSend.savedInvalidAddresses.length) || 0;

    document.getElementById("invalidAddressCount").textContent = invalidAddressCount;
    document.getElementById("toggleInvalidBtn").disabled = invalidAddressCount === 0;

    // Download button enabled if either list has items
    document.getElementById("downloadInvalidBtn").disabled =
      invalidAddressCount === 0 && savedInvalidAddressCount === 0;

    if (invalidAddressCount > 0) {
      populateFieldContent("invalid", ClearSend.invalidAddresses);
    } else {
      const invalidContent = document.getElementById("invalidContent");
      if (invalidContent) {
        invalidContent.innerHTML = "";
        invalidContent.style.display = "none";
      }
    }

    // Update saved invalid addresses display
    document.getElementById("savedInvalidAddressCount").textContent = savedInvalidAddressCount;
    document.getElementById("toggleSavedInvalidBtn").disabled = savedInvalidAddressCount === 0;

    if (savedInvalidAddressCount > 0) {
      populateFieldContent("savedInvalid", ClearSend.savedInvalidAddresses);
    } else {
      const savedInvalidContent = document.getElementById("savedInvalidContent");
      if (savedInvalidContent) {
        savedInvalidContent.innerHTML = "";
        savedInvalidContent.style.display = "none";
      }
    }

    // Add click event listeners to all buttons
    document.querySelectorAll(".recipient-copy").forEach((button) => {
      button.addEventListener("click", handleCopyRecipient);
    });
    document.querySelectorAll(".recipient-delete").forEach((button) => {
      button.addEventListener("click", handleRemoveRecipient);
    });
  } catch (error) {
    showToast("Failed to update display", "error");
  } finally {
    // Always release lock
    ClearSend.isUpdatingDisplay = false;
  }
}

function updateStatusIndicator(count) {
  const indicator = document.getElementById("statusText");
  if (count === 0) {
    indicator.textContent = "No recipients";
  } else {
    indicator.textContent = `${count} recipient${count !== 1 ? "s" : ""}`;
  }
}

function updateStatusIndicatorWithValidation(total, valid, warnings) {
  const statusIcon = document.querySelector(".status-icon");
  const statusText = document.getElementById("statusText");

  if (total === 0) {
    statusIcon.textContent = "🟢";
    statusText.textContent = "No recipients";
  } else if (warnings > 0) {
    statusIcon.textContent = "⚠️";
    statusText.textContent = `${warnings} warning${warnings !== 1 ? "s" : ""}`;
  } else {
    statusIcon.textContent = "🟢";
    statusText.textContent = `${total} recipient${total !== 1 ? "s" : ""}`;
  }
}

async function displayValidationResults(validateAction) {
  const list = document.getElementById("recipientList");
  list.innerHTML = "";

  if (validateAction.validationResults && validateAction.validationResults.length > 0) {
    validateAction.validationResults.forEach((result) => {
      const item = createRecipientItem(result);
      list.appendChild(item);
    });
  }

  // Update status indicator with validation results
  const errorCount = validateAction.errorCount || 0;
  const warningCount = validateAction.warningCount || 0;
  const validCount = validateAction.validCount || 0;

  updateValidationStatus(validCount, warningCount, errorCount);
}

function updateValidationStatus(validCount, warningCount, errorCount) {
  const statusIcon = document.querySelector(".status-icon");
  const statusText = document.getElementById("statusText");

  if (errorCount > 0) {
    statusIcon.textContent = "🔴";
    statusText.textContent = `${errorCount} invalid removed`;
  } else if (warningCount > 0) {
    statusIcon.textContent = "⚠️";
    statusText.textContent = `${warningCount} warnings`;
  } else {
    statusIcon.textContent = "🟢";
    statusText.textContent = `${validCount} valid`;
  }
}

function updateValidationStatusOnly(validateAction) {
  if (!validateAction) return;

  const errorCount = validateAction.errorCount || 0;
  const warningCount = validateAction.warningCount || 0;
  const validCount = validateAction.validCount || 0;

  updateValidationStatus(validCount, warningCount, errorCount);
}

function createRecipientItem(result) {
  const item = document.createElement("div");
  item.className = "recipient-item";

  const statusIcon = result.status === "valid" ? "🟢" : result.status === "warning" ? "⚠️" : "🔴";

  item.innerHTML = `
        <span class="recipient-status">${statusIcon}</span>
        <span class="recipient-email">${result.address}</span>
        <div class="recipient-actions">
            ${result.status !== "valid" ? '<button class="recipient-action">Edit ✎</button>' : ""}
            <button class="recipient-action">Remove –</button>
        </div>
    `;

  return item;
}

async function displayAllResults(result) {
  const list = document.getElementById("recipientList");
  list.innerHTML = "";

  // Create summary of all actions
  const summaryDiv = document.createElement("div");
  summaryDiv.className = "processing-summary";
  summaryDiv.innerHTML = "<h4>Processing Summary</h4>";

  result.actions.forEach((action) => {
    const actionDiv = document.createElement("div");
    actionDiv.className = "action-summary";

    switch (action.type) {
      case "validate":
        actionDiv.innerHTML = `
                    <strong>📋 Validation:</strong>
                    <ul>
                        <li>Processed: ${action.processed} recipients</li>
                        <li>Valid: ${action.validCount}</li>
                        <li>Warnings: ${action.warningCount}</li>
                        <li>Removed invalid: ${action.removedInvalid}</li>
                    </ul>
                `;
        break;
      case "dedupe":
        actionDiv.innerHTML = `
                    <strong>🔗 Deduplication:</strong>
                    <ul>
                        <li>Processed: ${action.processed} recipients</li>
                        <li>Duplicates found: ${action.duplicatesFound || 0}</li>
                        <li>Removed: ${action.removed ? action.removed.length : 0}</li>
                    </ul>
                `;
        break;
      case "sort":
        actionDiv.innerHTML = `
                    <strong>📶 Sorting:</strong>
                    <ul>
                        <li>Processed: ${action.processed} recipients</li>
                        <li>Sorted alphabetically by name</li>
                    </ul>
                `;
        break;
      case "flagExt":
        actionDiv.innerHTML = `
                    <strong>🏢 External Detection:</strong>
                    <ul>
                        <li>Total: ${action.summary.totalRecipients} recipients</li>
                        <li>Internal: ${action.summary.internalCount}</li>
                        <li>External: ${action.summary.externalCount}</li>
                        <li>External domains: ${action.summary.uniqueExternalDomains}</li>
                    </ul>
                `;
        break;
    }

    summaryDiv.appendChild(actionDiv);
  });

  list.appendChild(summaryDiv);

  // Show validation results if available
  const validateAction = result.actions.find((a) => a.type === "validate");
  if (validateAction && validateAction.validationResults) {
    const validationDiv = document.createElement("div");
    validationDiv.className = "validation-details";
    validationDiv.innerHTML = "<h4>Validation Details</h4>";

    validateAction.validationResults.forEach((vResult) => {
      if (vResult.status !== "valid") {
        const item = createRecipientItem(vResult);
        validationDiv.appendChild(item);
      }
    });

    if (validateAction.validationResults.some((r) => r.status !== "valid")) {
      list.appendChild(validationDiv);
    }
  }

  // Update overall status
  const totalErrors = result.actions.reduce((sum, action) => sum + (action.errorCount || 0), 0);
  const totalWarnings = result.actions.reduce((sum, action) => sum + (action.warningCount || 0), 0);
  const finalCount = result.to.length + result.cc.length + result.bcc.length;

  updateValidationStatus(finalCount, totalWarnings, totalErrors);
}

// Internal Domains Management
function renderInternalDomains() {
  const domainsList = document.getElementById("domainsList");
  domainsList.innerHTML = "";

  ClearSend.settings.internalDomains.forEach((domain, index) => {
    const domainItem = document.createElement("div");
    domainItem.className = "domain-item";
    const isPlaceholder = domain === "mydomain.com";
    domainItem.innerHTML = `
            <input type="text" class="domain-input${isPlaceholder ? " placeholder" : ""}" value="${domain}" data-index="${index}" />
            <button class="domain-btn add-btn" data-index="${index}" title="Add domain">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" width="14" height="14" fill="currentColor"><path d="M256 80c0-17.7-14.3-32-32-32s-32 14.3-32 32l0 144L48 224c-17.7 0-32 14.3-32 32s14.3 32 32 32l144 0 0 144c0 17.7 14.3 32 32 32s32-14.3 32-32l0-144 144 0c17.7 0 32-14.3 32-32s-14.3-32-32-32l-144 0 0-144z"/></svg>
            </button>
            <button class="domain-btn remove-btn" data-index="${index}" title="Remove domain">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" width="14" height="14" fill="currentColor"><path d="M432 256c0 17.7-14.3 32-32 32L48 288c-17.7 0-32-14.3-32-32s14.3-32 32-32l352 0c17.7 0 32 14.3 32 32z"/></svg>
            </button>
        `;
    domainsList.appendChild(domainItem);
  });

  // Add event listeners
  document.querySelectorAll(".domain-input").forEach((input) => {
    input.addEventListener("change", handleDomainChange);
    input.addEventListener("focus", handleDomainFocus);
    input.addEventListener("blur", handleDomainBlur);
  });
  document.querySelectorAll(".remove-btn").forEach((btn) => {
    btn.addEventListener("click", handleRemoveDomain);
  });
  document.querySelectorAll(".add-btn").forEach((btn) => {
    btn.addEventListener("click", handleAddDomain);
  });

  // Update domain-dependent features state
  updateDomainDependentFeatures();
}

/**
 * Handle domain input focus - clear placeholder
 */
function handleDomainFocus(event) {
  if (event.target.classList.contains("placeholder")) {
    event.target.value = "";
    event.target.classList.remove("placeholder");
  }
}

/**
 * Handle domain input blur - restore placeholder if empty
 */
function handleDomainBlur(event) {
  const index = parseInt(event.target.getAttribute("data-index"));
  const value = event.target.value.trim();

  if (!value) {
    event.target.value = "mydomain.com";
    event.target.classList.add("placeholder");
    ClearSend.settings.internalDomains[index] = "mydomain.com";
    saveSettings();
    updateDomainDependentFeatures();
  }
}

/**
 * Handle domain input change
 * Validates domain format before saving
 * @param {Event} event - Change event from domain input
 */
function handleDomainChange(event) {
  const index = parseInt(event.target.getAttribute("data-index"));
  const newValue = event.target.value.trim();

  // Validate domain format
  if (!newValue) {
    showToast("Domain cannot be empty", "warning");
    return;
  }

  // Basic domain validation (alphanumeric, dots, hyphens)
  const domainRegex = /^[a-zA-Z0-9][a-zA-Z0-9-]*(\.[a-zA-Z0-9][a-zA-Z0-9-]*)*\.[a-zA-Z]{2,}$/;
  if (!domainRegex.test(newValue)) {
    showToast("Invalid domain format", "error");
    event.target.value = ClearSend.settings.internalDomains[index] || "";
    return;
  }

  // Remove placeholder styling if valid domain entered
  event.target.classList.remove("placeholder");

  // Save validated domain
  ClearSend.settings.internalDomains[index] = newValue;
  saveSettings();
  updateDomainDependentFeatures();
  showToast("Domain updated", "success");
}

/**
 * Check if there are any valid (non-placeholder) domains defined
 */
function hasValidDomains() {
  return ClearSend.settings.internalDomains.some(
    (domain) => domain && domain !== "mydomain.com" && domain.trim() !== ""
  );
}

/**
 * Get only valid (non-placeholder) internal domains
 */
function getValidInternalDomains() {
  return (ClearSend.settings.internalDomains || []).filter(
    (domain) => domain && domain !== "mydomain.com" && domain.trim() !== ""
  );
}

/**
 * Update domain-dependent features (prioritize internal, remove external)
 * Disable them if no valid domains are defined
 */
function updateDomainDependentFeatures() {
  const hasValid = hasValidDomains();

  // Get the feature items for domain-dependent options
  const prioritizeInternalItem = document.querySelector('[data-step="prioritizeInternal"]');
  const removeExternalItem = document.querySelector('[data-step="removeExternal"]');

  const prioritizeInternalCheck = document.getElementById("prioritizeInternalCheck");
  const removeExternalCheck = document.getElementById("removeExternalCheck");

  if (!hasValid) {
    // Disable domain-dependent features
    if (prioritizeInternalItem) prioritizeInternalItem.classList.add("disabled");
    if (removeExternalItem) removeExternalItem.classList.add("disabled");

    if (prioritizeInternalCheck) {
      prioritizeInternalCheck.checked = false;
      prioritizeInternalCheck.disabled = true;
    }
    if (removeExternalCheck) {
      removeExternalCheck.checked = false;
      removeExternalCheck.disabled = true;
    }

    // Remove from enabled steps
    ClearSend.settings.enabledSteps = ClearSend.settings.enabledSteps.filter(
      (step) => step !== "prioritizeInternal" && step !== "removeExternal"
    );
    saveSettings();
  } else {
    // Enable domain-dependent features
    if (prioritizeInternalItem) prioritizeInternalItem.classList.remove("disabled");
    if (removeExternalItem) removeExternalItem.classList.remove("disabled");

    if (prioritizeInternalCheck) prioritizeInternalCheck.disabled = false;
    if (removeExternalCheck) removeExternalCheck.disabled = false;
  }
}

function handleRemoveDomain(event) {
  const index = parseInt(event.currentTarget.getAttribute("data-index"));
  if (ClearSend.settings.internalDomains.length > 1) {
    ClearSend.settings.internalDomains.splice(index, 1);
    saveSettings();
    renderInternalDomains();
    showToast("Domain removed", "success");
  } else {
    showToast("At least one domain must remain", "warning");
  }
}

/**
 * Handle adding a new domain
 * @param {Event} event - Click event from add button
 */
function handleAddDomain(event) {
  if (ClearSend.settings.internalDomains.length >= CONFIG.MAX_INTERNAL_DOMAINS) {
    showToast(`Maximum ${CONFIG.MAX_INTERNAL_DOMAINS} internal domains allowed`, "error");
    return;
  }
  ClearSend.settings.internalDomains.push("newdomain.com");
  saveSettings();
  renderInternalDomains();
  showToast("Domain added - please update with your domain", "success");
}

function saveSettings() {
  try {
    // Set the settings in roaming storage
    Office.context.roamingSettings.set("clearSendSettings", ClearSend.settings);

    // Save asynchronously with callback
    Office.context.roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        updateProcessingOptionsSummary();
      } else {
      }
    });
  } catch (error) {}
}

/**
 * Save saved invalid addresses to roaming storage
 */
function saveSavedInvalidAddresses() {
  try {
    // Set the saved invalid addresses in roaming storage
    Office.context.roamingSettings.set("savedInvalidAddresses", ClearSend.savedInvalidAddresses);

    // Save asynchronously with callback
    Office.context.roamingSettings.saveAsync((result) => {
      // Silent save - no action needed
    });
  } catch (error) {
    // Silent error - no action needed
  }
}

function updateProcessingOptionsSummary() {
  const summary = document.getElementById("processingOptionsSummary");
  if (!summary) return;

  const enabledSteps = ClearSend.settings.enabledSteps || [];
  const stepOrder = ClearSend.settings.stepOrder || [
    "sort",
    "dedupe",
    "validate",
    "prioritizeInternal",
    "removeExternal",
    "keepInvalid",
  ];

  // Map step IDs to their short descriptions
  const stepLabels = {
    sort: "Alphabetical sorting",
    dedupe: "Remove duplicates",
    validate: "Prevent invalids processing",
    prioritizeInternal: "Internal domains first",
    removeExternal: "Remove externals",
    keepInvalid: "Save invalid addresses",
  };

  // Build options list in the order defined by stepOrder
  const options = [];
  stepOrder.forEach((step) => {
    if (enabledSteps.includes(step) && stepLabels[step]) {
      options.push(stepLabels[step]);
    }
  });

  summary.textContent = options.length > 0 ? options.join(", ") : "None";
}

/**
 * Email Validation and Status Check
 * Determines if an email is valid, internal, or external
 * @param {string} address - Email address to check (may include display name)
 * @returns {Object} Status object with status, circle color, and label
 */
function getEmailStatus(address) {
  // Input validation
  if (!address || typeof address !== "string") {
    return { status: "invalid", circle: "red", label: "Invalid" };
  }

  // Extract email from format "Name <email@domain.com>" or just "email@domain.com"
  let email = address.trim();
  const match = email.match(/<(.+)>/);
  if (match) {
    email = match[1].trim();
  }

  // Use isValidEmail for consistent validation
  if (!isValidEmail(email)) {
    return { status: "invalid", circle: "red", label: "Invalid" };
  }

  // Email is valid, now check if internal or external
  const parts = email.split("@");
  const domain = parts[1].toLowerCase().trim();

  // Check against internal domains
  const internalDomains = getValidInternalDomains();
  const isInternal = internalDomains.some((internalDomain) => {
    const normalizedInternal = (internalDomain || "").toLowerCase().trim();
    if (!normalizedInternal) return false;
    return domain === normalizedInternal || domain.endsWith("." + normalizedInternal);
  });

  if (isInternal) {
    return { status: "internal", circle: null, label: "Valid & Internal" };
  } else {
    // External emails are always flagged with warning (orange)
    return { status: "external", circle: "orange", label: "Valid & External" };
  }
}

// Undo functionality
function saveRecipientState(recipients) {
  ClearSend.lastRecipientState = {
    to: [...recipients.to],
    cc: [...recipients.cc],
    bcc: [...recipients.bcc],
  };
}

function updateLastAction(message, enableUndo = true) {
  ClearSend.lastActionMessage = message;
  const lastActionText = document.getElementById("lastActionText");
  const undoBtn = document.getElementById("undoBtn");

  if (lastActionText) {
    lastActionText.textContent = message;
  }

  if (undoBtn) {
    undoBtn.disabled = !enableUndo;
  }
}

async function handleUndo() {
  try {
    if (!ClearSend.lastRecipientState) {
      return; // Silently do nothing if no state to undo
    }

    showProgress("Restoring previous lists...");

    const toRecipients = convertToOfficeFormat(ClearSend.lastRecipientState.to);
    const ccRecipients = convertToOfficeFormat(ClearSend.lastRecipientState.cc);
    const bccRecipients = convertToOfficeFormat(ClearSend.lastRecipientState.bcc);

    await updateRecipientsDirectly(toRecipients, ccRecipients, bccRecipients);

    // Clear the saved state and update UI
    ClearSend.lastRecipientState = null;
    updateLastAction("Previous lists restored");

    const undoBtn = document.getElementById("undoBtn");
    if (undoBtn) {
      undoBtn.disabled = true;
    }

    await updateRecipientDisplay();

    // No toast on successful undo
  } catch (error) {
    showToast("Failed to restore previous lists", "error");
  } finally {
    hideProgress();
  }
}

// Populate field content with recipient list
function populateFieldContent(field, addresses) {
  const contentId = field + "Content";
  const content = document.getElementById(contentId);
  const btnId = "toggle" + field.charAt(0).toUpperCase() + field.slice(1) + "Btn";
  const btn = document.getElementById(btnId);

  if (!content) {
    return;
  }

  // Validate addresses parameter
  if (!addresses || !Array.isArray(addresses)) {
    addresses = [];
  }

  // Clear existing content
  content.innerHTML = "";

  // If field is now empty, hide it and reset button to eye icon
  if (addresses.length === 0) {
    content.style.display = "none";
    content.innerHTML = '<div class="empty-state">No recipients</div>';

    // Reset button to eye icon (not eye-slash)
    if (btn) {
      btn.innerHTML =
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512" width="14" height="14" fill="currentColor"><path d="M288 32c-80.8 0-145.5 36.8-192.6 80.6C48.6 156 17.3 208 2.5 243.7c-3.3 7.9-3.3 16.7 0 24.6C17.3 304 48.6 356 95.4 399.4C142.5 443.2 207.2 480 288 480s145.5-36.8 192.6-80.6c46.8-43.5 78.1-95.4 93-131.1c3.3-7.9 3.3-16.7 0-24.6c-14.9-35.7-46.2-87.7-93-131.1C433.5 68.8 368.8 32 288 32zM144 256a144 144 0 1 1 288 0 144 144 0 1 1 -288 0zm144-64c0 35.3-28.7 64-64 64c-7.1 0-13.9-1.2-20.3-3.3c-5.5-1.8-11.9 1.6-11.7 7.4c.3 6.9 1.3 13.8 3.2 20.7c13.7 51.2 66.4 81.6 117.6 67.9s81.6-66.4 67.9-117.6c-11.1-41.5-47.8-69.4-88.6-71.1c-5.8-.2-9.2 6.1-7.4 11.7c2.1 6.4 3.3 13.2 3.3 20.3z"/></svg>';
      btn.title = `Show ${field.toUpperCase()} list`;
    }
    return;
  }

  // If field has addresses, keep current display state (visible lists stay visible)
  const currentDisplay = content.style.display;
  content.style.display = currentDisplay || "none";

  // Count duplicates for tooltip
  const addressCounts = {};
  addresses.forEach((addr) => {
    const normalized = addr.toLowerCase().trim();
    addressCounts[normalized] = (addressCounts[normalized] || 0) + 1;
  });

  // Add each recipient
  addresses.forEach((address) => {
    const emailStatus = getEmailStatus(address);
    const normalizedAddress = address.toLowerCase().trim();
    const isDuplicate = addressCounts[normalizedAddress] > 1;

    const item = document.createElement("div");
    item.className = "recipient-item";

    // Build tooltip reasons
    const reasons = [];
    if (emailStatus.status === "invalid") reasons.push("Invalid");
    if (emailStatus.status === "external") reasons.push("External");
    if (isDuplicate) reasons.push("Duplicated");

    const tooltipText = reasons.length > 0 ? reasons.join(", ") : "";

    // Create circle HTML if needed
    const circleHTML = emailStatus.circle
      ? `<span class="recipient-circle ${emailStatus.circle}" title="${tooltipText}"></span>`
      : '<span class="recipient-circle-placeholder"></span>';

    item.innerHTML = `
            ${circleHTML}
            <span class="recipient-email" title="${escapeHtml(address)}" data-full-email="${escapeHtml(address)}">${escapeHtml(address)}</span>
            <button class="recipient-copy" data-address="${encodeURIComponent(address)}" title="Copy to clipboard">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" width="14" height="14" fill="currentColor"><path d="M384 336H192c-8.8 0-16-7.2-16-16V64c0-8.8 7.2-16 16-16l140.1 0L400 115.9V320c0 8.8-7.2 16-16 16zM192 384H384c35.3 0 64-28.7 64-64V115.9c0-12.7-5.1-24.9-14.1-33.9L366.1 14.1c-9-9-21.2-14.1-33.9-14.1H192c-35.3 0-64 28.7-64 64V320c0 35.3 28.7 64 64 64zM64 128c-35.3 0-64 28.7-64 64V448c0 35.3 28.7 64 64 64H256c35.3 0 64-28.7 64-64V416H272v32c0 8.8-7.2 16-16 16H64c-8.8 0-16-7.2-16-16V192c0-8.8 7.2-16 16-16H96V128H64z"/></svg>
            </button>
            <button class="recipient-delete" data-field="${escapeHtml(field)}" data-address="${encodeURIComponent(address)}" title="Remove recipient">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" width="14" height="14" fill="currentColor"><path d="M135.2 17.7L128 32H32C14.3 32 0 46.3 0 64S14.3 96 32 96H416c17.7 0 32-14.3 32-32s-14.3-32-32-32H320l-7.2-14.3C307.4 6.8 296.3 0 284.2 0H163.8c-12.1 0-23.2 6.8-28.6 17.7zM416 128H32L53.2 467c1.6 25.3 22.6 45 47.9 45H346.9c25.3 0 46.3-19.7 47.9-45L416 128z"/></svg>
            </button>
        `;
    content.appendChild(item);
  });
}

// Field visibility toggle (To, CC, BCC)
window.toggleField = function (field) {
  const content = document.getElementById(field + "Content");
  const btn = document.getElementById(
    "toggle" + field.charAt(0).toUpperCase() + field.slice(1) + "Btn"
  );

  if (!content || !btn) {
    return;
  }

  // Simply toggle display style
  if (content.style.display === "none") {
    content.style.display = "block";
    // Change to eye-slash icon
    btn.innerHTML =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 512" width="14" height="14" fill="currentColor"><path d="M38.8 5.1C28.4-3.1 13.3-1.2 5.1 9.2S-1.2 34.7 9.2 42.9l592 464c10.4 8.2 25.5 6.3 33.7-4.1s6.3-25.5-4.1-33.7L525.6 386.7c39.6-40.6 66.4-86.1 79.9-118.4c3.3-7.9 3.3-16.7 0-24.6c-14.9-35.7-46.2-87.7-93-131.1C465.5 68.8 400.8 32 320 32c-68.2 0-125 26.3-169.3 60.8L38.8 5.1zM223.1 149.5C248.6 126.2 282.7 112 320 112c79.5 0 144 64.5 144 144c0 24.9-6.3 48.3-17.4 68.7L408 294.5c8.4-19.3 10.6-41.4 4.8-63.3c-11.1-41.5-47.8-69.4-88.6-71.1c-5.8-.2-9.2 6.1-7.4 11.7c2.1 6.4 3.3 13.2 3.3 20.3c0 10.2-2.4 19.8-6.6 28.3l-90.3-70.8zM373 389.9c-16.4 6.5-34.3 10.1-53 10.1c-79.5 0-144-64.5-144-144c0-6.9 .5-13.6 1.4-20.2L83.1 161.5C60.3 191.2 44 220.8 34.5 243.7c-3.3 7.9-3.3 16.7 0 24.6c14.9 35.7 46.2 87.7 93 131.1C174.5 443.2 239.2 480 320 480c47.8 0 89.9-12.9 126.2-32.5L373 389.9z"/></svg>';
    btn.title = `Hide ${field.toUpperCase()} list`;
  } else {
    content.style.display = "none";
    // Change to eye icon
    btn.innerHTML =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512" width="14" height="14" fill="currentColor"><path d="M288 32c-80.8 0-145.5 36.8-192.6 80.6C48.6 156 17.3 208 2.5 243.7c-3.3 7.9-3.3 16.7 0 24.6C17.3 304 48.6 356 95.4 399.4C142.5 443.2 207.2 480 288 480s145.5-36.8 192.6-80.6c46.8-43.5 78.1-95.4 93-131.1c3.3-7.9 3.3-16.7 0-24.6c-14.9-35.7-46.2-87.7-93-131.1C433.5 68.8 368.8 32 288 32zM144 256a144 144 0 1 1 288 0 144 144 0 1 1 -288 0zm144-64c0 35.3-28.7 64-64 64c-7.1 0-13.9-1.2-20.3-3.3c-5.5-1.8-11.9 1.6-11.7 7.4c.3 6.9 1.3 13.8 3.2 20.7c13.7 51.2 66.4 81.6 117.6 67.9s81.6-66.4 67.9-117.6c-11.1-41.5-47.8-69.4-88.6-71.1c-5.8-.2-9.2 6.1-7.4 11.7c2.1 6.4 3.3 13.2 3.3 20.3z"/></svg>';
    btn.title = `Show ${field.toUpperCase()} list`;
  }
};

// Switch to config tab (accessible from HTML onclick)
window.switchToConfigTab = function () {
  switchTab("config");
};

// Switch to details tab (accessible from HTML onclick)
window.switchToDetailsTab = function () {
  switchTab("details");
};

// Download CSV
async function handleDownloadCSV() {
  try {
    const recipients = await getCurrentRecipients();
    const totalRecipients = recipients.to.length + recipients.cc.length + recipients.bcc.length;

    if (totalRecipients === 0) {
      showToast("No recipients to download", "warning");
      return;
    }

    // Create CSV content: 3 lines, each with semicolon-separated emails
    const toLine = recipients.to.join(";");
    const ccLine = recipients.cc.join(";");
    const bccLine = recipients.bcc.join(";");
    const csvContent = `${toLine}\n${ccLine}\n${bccLine}`;

    // Create blob and download
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    link.setAttribute("download", "recipients.csv");
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showToast("CSV downloaded successfully", "success");
  } catch (error) {
    showToast("Failed to download CSV", "error");
  }
}

// Download Invalid Addresses CSV
async function handleDownloadInvalidCSV() {
  try {
    const hasInvalid = ClearSend.invalidAddresses && ClearSend.invalidAddresses.length > 0;
    const hasSavedInvalid =
      ClearSend.savedInvalidAddresses && ClearSend.savedInvalidAddresses.length > 0;

    if (!hasInvalid && !hasSavedInvalid) {
      showToast("No invalid addresses to download", "warning");
      return;
    }

    // Create CSV content: two rows
    // Row 1: Invalid addresses (current)
    // Row 2: Saved invalid addresses
    const invalidLine = hasInvalid ? ClearSend.invalidAddresses.join(";") : "";
    const savedInvalidLine = hasSavedInvalid ? ClearSend.savedInvalidAddresses.join(";") : "";
    const csvContent = `${invalidLine}\n${savedInvalidLine}`;

    // Create blob and download
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    link.setAttribute("download", "invalid_addresses.csv");
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showToast("Invalid addresses CSV downloaded", "success");
  } catch (error) {
    showToast("Failed to download invalid addresses", "error");
  }
}

/**
 * Handle restore default settings
 */
async function handleRestoreDefaults() {
  try {
    // Clear all roaming settings
    Office.context.roamingSettings.remove("clearSendSettings");
    Office.context.roamingSettings.remove("savedInvalidAddresses");

    // Save the cleared settings
    await new Promise((resolve, reject) => {
      Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error("Failed to save settings"));
        }
      });
    });

    // Reset in-memory state to defaults
    ClearSend.settings = {
      enabledSteps: ["sort", "dedupe", "validate", "prioritizeInternal"],
      stepOrder: [
        "sort",
        "dedupe",
        "validate",
        "prioritizeInternal",
        "removeExternal",
        "keepInvalid",
      ],
      orgDomain: "",
      internalDomains: ["mydomain.com"],
      keepInvalid: false,
    };

    // Clear saved invalid addresses
    ClearSend.savedInvalidAddresses = [];

    // Reset UI checkboxes to defaults
    document.getElementById("sortCheck").checked = true;
    document.getElementById("dedupeCheck").checked = true;
    document.getElementById("validateCheck").checked = true;
    document.getElementById("prioritizeInternalCheck").checked = true;
    document.getElementById("removeExternalCheck").checked = false;
    document.getElementById("keepInvalidCheck").checked = false;

    // Reset internal domains to default
    renderInternalDomains();

    // Reset feature order to default
    restoreFeatureOrder();

    // Update processing options summary
    updateProcessingOptionsSummary();

    // Update domain-dependent features state
    updateDomainDependentFeatures();

    // Refresh recipient display
    await updateRecipientDisplay();

    showToast("Default settings restored", "success");
  } catch (error) {
    showToast("Failed to restore default settings", "error");
  }
}
