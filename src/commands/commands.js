/* ClearSend Commands - Quick Actions from Ribbon */

/* global Office */

Office.onReady(() => {
    
});

/**
 * Quick Clean function - validates, deduplicates, and sorts recipients from ribbon button
 * @param event {Office.AddinCommands.Event}
 */
async function quickClean(event) {
    try {
        const item = Office.context.mailbox.item;

        // Get current recipients
        const recipients = await getAllRecipients();
        const totalOriginal = recipients.to.length + recipients.cc.length + recipients.bcc.length;

        // Step 1: Validate and remove invalid emails
        const validatedTo = validateAndFilterRecipients(recipients.to);
        const validatedCc = validateAndFilterRecipients(recipients.cc);
        const validatedBcc = validateAndFilterRecipients(recipients.bcc);

        // Step 2: Remove duplicates across all fields (priority: To > CC > BCC)
        const dedupedRecipients = deduplicateAllRecipients(validatedTo, validatedCc, validatedBcc);

        // Step 3: Sort each recipient list alphabetically
        const sortedTo = sortRecipientsAlphabetically(dedupedRecipients.to);
        const sortedCc = sortRecipientsAlphabetically(dedupedRecipients.cc);
        const sortedBcc = sortRecipientsAlphabetically(dedupedRecipients.bcc);

        // Update recipients
        await updateAllRecipients(sortedTo, sortedCc, sortedBcc);

        // Calculate results
        const totalFinal = sortedTo.length + sortedCc.length + sortedBcc.length;
        const invalidCount = totalOriginal - (validatedTo.length + validatedCc.length + validatedBcc.length);
        const duplicateCount = (validatedTo.length + validatedCc.length + validatedBcc.length) - totalFinal;
        const sortedCount = totalFinal;

        // Show success notification
        showNotification(
            `ClearSend completed: ${invalidCount} invalid, ${duplicateCount} duplicated, ${sortedCount} sorted addresses.`,
            Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage
        );

    } catch (error) {
        
        showNotification(
            'Quick clean failed. Please try again.',
            Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        );
    }

    event.completed();
}

/**
 * Get all recipients from the current email
 */
function getAllRecipients() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;

        item.to.getAsync((toResult) => {
            if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error('Failed to get To recipients'));
                return;
            }

            item.cc.getAsync((ccResult) => {
                if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error('Failed to get CC recipients'));
                    return;
                }

                item.bcc.getAsync((bccResult) => {
                    if (bccResult.status !== Office.AsyncResultStatus.Succeeded) {
                        reject(new Error('Failed to get BCC recipients'));
                        return;
                    }

                    resolve({
                        to: toResult.value || [],
                        cc: ccResult.value || [],
                        bcc: bccResult.value || []
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
    return recipients.filter(recipient => {
        const email = recipient.emailAddress || '';

        // Basic email validation
        if (!email || typeof email !== 'string') return false;
        if (!email.includes('@')) return false;
        if (email.split('@').length !== 2) return false;

        const [localPart, domainPart] = email.split('@');
        if (!localPart || !domainPart) return false;
        if (!domainPart.includes('.')) return false;
        if (email.includes('..')) return false;

        // Basic regex check
        const emailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;
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
    uniqueTo.forEach(recipient => {
        const email = (recipient.emailAddress || '').toLowerCase();
        if (!allEmails.has(email)) {
            allEmails.add(email);
            finalTo.push(recipient);
        }
    });

    // Process CC recipients (medium priority)
    uniqueCc.forEach(recipient => {
        const email = (recipient.emailAddress || '').toLowerCase();
        if (!allEmails.has(email)) {
            allEmails.add(email);
            finalCc.push(recipient);
        }
    });

    // Process BCC recipients (lowest priority)
    uniqueBcc.forEach(recipient => {
        const email = (recipient.emailAddress || '').toLowerCase();
        if (!allEmails.has(email)) {
            allEmails.add(email);
            finalBcc.push(recipient);
        }
    });

    return {
        to: finalTo,
        cc: finalCc,
        bcc: finalBcc
    };
}

/**
 * Remove duplicates within a single array
 */
function deduplicateWithinArray(recipients) {
    const seen = new Set();
    return recipients.filter(recipient => {
        const email = (recipient.emailAddress || '').toLowerCase();
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
        const nameA = (a.displayName || a.emailAddress || '').toLowerCase();
        const nameB = (b.displayName || b.emailAddress || '').toLowerCase();
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
                reject(new Error('Failed to update To recipients'));
                return;
            }

            item.cc.setAsync(ccRecipients, (ccResult) => {
                if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error('Failed to update CC recipients'));
                    return;
                }

                item.bcc.setAsync(bccRecipients, (bccResult) => {
                    if (bccResult.status !== Office.AsyncResultStatus.Succeeded) {
                        reject(new Error('Failed to update BCC recipients'));
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
