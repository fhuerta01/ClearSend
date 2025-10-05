/* ClearSend Commands - Quick Actions from Ribbon */

/* global Office, window, setTimeout */

Office.onReady(() => {
  // Commands ready
});

/**
 * Quick Clean function - processes recipients using the same logic as the task pane
 * Respects user's enabled settings from configuration
 * @param event {Office.AddinCommands.Event}
 */
async function quickClean(event) {
  try {
    // Load user settings
    const settings = Office.context.roamingSettings.get("clearSendSettings") || {
      enabledSteps: ["sort", "dedupe", "validate", "prioritizeInternal"],
      stepOrder: ["sort", "dedupe", "validate", "prioritizeInternal", "removeExternal"],
      orgDomain: "",
      internalDomains: [],
    };

    // Get steps in the user's preferred order (stepOrder defines the execution order)
    // Filter to only include steps that are enabled
    const orderedSteps = settings.stepOrder || settings.enabledSteps;
    const enabledStepsInOrder = orderedSteps.filter((step) => settings.enabledSteps.includes(step));

    // Get current recipients
    const recipients = await getAllRecipients();
    const totalOriginal = recipients.to.length + recipients.cc.length + recipients.bcc.length;

    // Convert to string format for processing
    const toStrings = recipients.to.map((r) =>
      r.displayName ? `${r.displayName} <${r.emailAddress}>` : r.emailAddress
    );
    const ccStrings = recipients.cc.map((r) =>
      r.displayName ? `${r.displayName} <${r.emailAddress}>` : r.emailAddress
    );
    const bccStrings = recipients.bcc.map((r) =>
      r.displayName ? `${r.displayName} <${r.emailAddress}>` : r.emailAddress
    );

    // Call the processors library if available
    if (window.ClearSendProcessors && window.ClearSendProcessors.processRecipients) {
      const result = window.ClearSendProcessors.processRecipients({
        to: toStrings,
        cc: ccStrings,
        bcc: bccStrings,
        userSettings: {
          enabledSteps: enabledStepsInOrder,
          internalDomains: (settings.internalDomains || []).filter(
            (d) => d && d !== "mydomain.com" && d.trim() !== ""
          ),
          orgDomain: settings.orgDomain || "",
        },
      });

      // Convert processed results back to Office format
      const processedTo = convertToOfficeFormat(result.result.to);
      const processedCc = convertToOfficeFormat(result.result.cc);
      const processedBcc = convertToOfficeFormat(result.result.bcc);

      // Update recipients
      await updateAllRecipients(processedTo, processedCc, processedBcc);

      // Show success notification
      const message = `Success. ${totalOriginal} addresses processed.`;
      showNotification(
        message,
        Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage
      );
    } else {
      // Fallback: use local processing if processors library not loaded
      let processedTo = recipients.to;
      let processedCc = recipients.cc;
      let processedBcc = recipients.bcc;

      // Validate if enabled
      if (settings.enabledSteps.includes("validate")) {
        const validatedTo = validateAndFilterRecipients(processedTo);
        const validatedCc = validateAndFilterRecipients(processedCc);
        const validatedBcc = validateAndFilterRecipients(processedBcc);

        processedTo = validatedTo;
        processedCc = validatedCc;
        processedBcc = validatedBcc;
      }

      // Deduplicate if enabled
      if (settings.enabledSteps.includes("dedupe")) {
        const deduped = deduplicateAllRecipients(processedTo, processedCc, processedBcc);
        processedTo = deduped.to;
        processedCc = deduped.cc;
        processedBcc = deduped.bcc;
      }

      // Sort if enabled
      if (settings.enabledSteps.includes("sort")) {
        processedTo = sortRecipientsAlphabetically(processedTo);
        processedCc = sortRecipientsAlphabetically(processedCc);
        processedBcc = sortRecipientsAlphabetically(processedBcc);
      }

      // Update recipients
      await updateAllRecipients(processedTo, processedCc, processedBcc);

      // Show success notification
      const message = `Success. ${totalOriginal} addresses processed.`;
      showNotification(
        message,
        Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage
      );
    }
  } catch (_error) {
    showNotification(
      "Quick clean failed. Please try again.",
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
    );
  }

  event.completed();
}

/**
 * Convert strings to Office format
 */
function convertToOfficeFormat(recipients) {
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
    return r;
  });
}

/**
 * Get all recipients from the current email
 */
function getAllRecipients() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;

    item.to.getAsync((toResult) => {
      if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error("Failed to get To recipients"));
        return;
      }

      item.cc.getAsync((ccResult) => {
        if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Failed to get CC recipients"));
          return;
        }

        item.bcc.getAsync((bccResult) => {
          if (bccResult.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error("Failed to get BCC recipients"));
            return;
          }

          resolve({
            to: toResult.value || [],
            cc: ccResult.value || [],
            bcc: bccResult.value || [],
          });
        });
      });
    });
  });
}

/**
 * Validate and filter recipients (remove invalid emails)
 */
function validateAndFilterRecipients(recipients) {
  return recipients.filter((recipient) => {
    const email = recipient.emailAddress || "";

    // Basic email validation
    if (!email || typeof email !== "string") return false;
    if (!email.includes("@")) return false;
    if (email.split("@").length !== 2) return false;

    const [localPart, domainPart] = email.split("@");
    if (!localPart || !domainPart) return false;
    if (!domainPart.includes(".")) return false;
    if (email.includes("..")) return false;

    // Basic regex check
    const emailRegex =
      /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;
    return emailRegex.test(email);
  });
}

/**
 * Remove duplicates across all recipient fields (To has priority over CC over BCC)
 */
function deduplicateAllRecipients(toRecipients, ccRecipients, bccRecipients) {
  // First dedupe within each array
  const uniqueTo = deduplicateWithinArray(toRecipients);
  const uniqueCc = deduplicateWithinArray(ccRecipients);
  const uniqueBcc = deduplicateWithinArray(bccRecipients);

  // Then dedupe across arrays with priority
  const allEmails = new Set();
  const finalTo = [];
  const finalCc = [];
  const finalBcc = [];

  // Process To recipients first (highest priority)
  uniqueTo.forEach((recipient) => {
    const email = (recipient.emailAddress || "").toLowerCase();
    if (!allEmails.has(email)) {
      allEmails.add(email);
      finalTo.push(recipient);
    }
  });

  // Process CC recipients (medium priority)
  uniqueCc.forEach((recipient) => {
    const email = (recipient.emailAddress || "").toLowerCase();
    if (!allEmails.has(email)) {
      allEmails.add(email);
      finalCc.push(recipient);
    }
  });

  // Process BCC recipients (lowest priority)
  uniqueBcc.forEach((recipient) => {
    const email = (recipient.emailAddress || "").toLowerCase();
    if (!allEmails.has(email)) {
      allEmails.add(email);
      finalBcc.push(recipient);
    }
  });

  return {
    to: finalTo,
    cc: finalCc,
    bcc: finalBcc,
  };
}

/**
 * Remove duplicates within a single array
 */
function deduplicateWithinArray(recipients) {
  const seen = new Set();
  return recipients.filter((recipient) => {
    const email = (recipient.emailAddress || "").toLowerCase();
    if (seen.has(email)) {
      return false;
    }
    seen.add(email);
    return true;
  });
}

/**
 * Sort recipients alphabetically
 */
function sortRecipientsAlphabetically(recipients) {
  return recipients.sort((a, b) => {
    const nameA = (a.displayName || a.emailAddress || "").toLowerCase();
    const nameB = (b.displayName || b.emailAddress || "").toLowerCase();
    return nameA.localeCompare(nameB);
  });
}

/**
 * Update all recipient fields
 */
function updateAllRecipients(toRecipients, ccRecipients, bccRecipients) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;

    item.to.setAsync(toRecipients, (toResult) => {
      if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error("Failed to update To recipients"));
        return;
      }

      item.cc.setAsync(ccRecipients, (ccResult) => {
        if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Failed to update CC recipients"));
          return;
        }

        item.bcc.setAsync(bccRecipients, (bccResult) => {
          if (bccResult.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error("Failed to update BCC recipients"));
            return;
          }

          resolve();
        });
      });
    });
  });
}

/**
 * Show notification message
 */
function showNotification(message, type) {
  const notification = {
    type: type,
    message: message,
    icon: "Icon.80x80",
    persistent: false,
  };

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ClearSendNotification",
    notification
  );

  // Auto-remove after 3 seconds
  setTimeout(() => {
    Office.context.mailbox.item.notificationMessages.removeAsync("ClearSendNotification");
  }, 3000);
}

/**
 * Legacy action function for compatibility
 */
function action(event) {
  quickClean(event);
}

// Register functions with Office
Office.actions.associate("quickClean", quickClean);
Office.actions.associate("action", action);
