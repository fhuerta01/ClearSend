/**
 * ClearSend Minimal Usage Ping
 *
 * PRIVACY GUARANTEE:
 * - No user identification
 * - No IP address logging
 * - No timestamps stored
 * - No request metadata captured
 * - Only increments a simple counter
 *
 * This endpoint exists ONLY to count add-in usage (how many times it's loaded).
 * Zero personally identifiable information is collected.
 */

let usageCount = 0; // In-memory counter (resets on cold start)

export default async function handler(req, res) {
  // Only accept POST requests
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // Increment counter (no data logging)
    usageCount++;

    // Return minimal response (no counter value shared)
    // This prevents clients from knowing usage statistics
    return res.status(200).json({ ok: true });

  } catch (error) {
    // Generic error (no details leaked)
    return res.status(500).json({ error: 'Server error' });
  }
}
