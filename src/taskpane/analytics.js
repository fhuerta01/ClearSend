/**
 * ClearSend Minimal Privacy-Safe Analytics
 *
 * PRIVACY GUARANTEE:
 * This module sends ONLY a single "ping" when the add-in loads.
 * - No user identification
 * - No email addresses
 * - No recipient data
 * - No personal information
 * - No IP addresses (beyond standard HTTP request)
 * - No tracking cookies
 * - No session IDs
 * - No timestamps
 * - No device fingerprinting
 *
 * The ping simply counts "add-in was loaded" - nothing more.
 * All email processing remains 100% local in your Outlook client.
 */

/**
 * Configuration
 */
const ANALYTICS_CONFIG = {
  ENABLED: true, // Set to false to completely disable analytics
  ENDPOINT: '/api/ping', // Vercel serverless function endpoint
  TIMEOUT_MS: 3000, // Request timeout (fail silently after 3 seconds)
};

/**
 * Send a privacy-safe usage ping
 * This function is called ONCE when the taskpane loads
 * Fails silently - never blocks or disrupts the add-in
 */
export async function sendUsagePing() {
  // Check if analytics is enabled
  if (!ANALYTICS_CONFIG.ENABLED) {
    return;
  }

  try {
    // Create abort controller for timeout
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), ANALYTICS_CONFIG.TIMEOUT_MS);

    // Send minimal ping (no data in body)
    // Using fetch with no-cors to avoid CORS issues
    await fetch(ANALYTICS_CONFIG.ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({}), // Empty body - no data sent
      signal: controller.signal,
      // Fire-and-forget: we don't care about the response
    }).finally(() => {
      clearTimeout(timeoutId);
    });

    // Intentionally ignore response
    // We don't want to know the result or any data back
  } catch (error) {
    // Fail silently - never disrupt the add-in
    // Don't log errors to avoid console pollution
  }
}

/**
 * Initialize analytics (call this once on taskpane load)
 */
export function initAnalytics() {
  // Send ping asynchronously (non-blocking)
  sendUsagePing().catch(() => {
    // Silently ignore failures
  });
}
