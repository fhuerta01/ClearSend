/**
 * ClearSend Minimal Usage Ping (Persistent Version)
 *
 * PRIVACY GUARANTEE:
 * - No user identification
 * - No IP address logging
 * - No timestamps stored
 * - No request metadata captured
 * - Only increments a simple counter in Vercel KV
 *
 * SETUP REQUIRED:
 * 1. Install Vercel KV: npm install @vercel/kv
 * 2. Create KV database in Vercel dashboard (free tier available)
 * 3. KV automatically connects via environment variables
 *
 * This endpoint exists ONLY to count add-in usage (how many times it's loaded).
 * Zero personally identifiable information is collected.
 */

import { kv } from '@vercel/kv';

const USAGE_KEY = 'clearsend:usage:count';

export default async function handler(req, res) {
  // Only accept POST requests
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // Increment counter in Vercel KV (atomic operation)
    await kv.incr(USAGE_KEY);

    // Return minimal response (no counter value shared)
    // This prevents clients from knowing usage statistics
    return res.status(200).json({ ok: true });

  } catch (error) {
    // Generic error (no details leaked)
    // Don't log error details to avoid leaking information
    return res.status(500).json({ error: 'Server error' });
  }
}
