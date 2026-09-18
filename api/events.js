// The only public API accepts one fixed event name, never Outlook data.
const { EVENTS } = require("../src/shared/analytics-events");
const MAX_BYTES = 80;
module.exports = async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("X-Content-Type-Options", "nosniff");
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).end();
  }
  const origin = process.env.ANALYTICS_ALLOWED_ORIGIN;
  if (!origin || req.headers.origin !== origin || req.headers["sec-fetch-site"] === "cross-site")
    return res.status(403).end();
  if (!/^application\/json(?:\s*;|$)/i.test(req.headers["content-type"] || ""))
    return res.status(415).end();
  if (Number(req.headers["content-length"]) > MAX_BYTES) return res.status(413).end();
  let body = req.body;
  try {
    if (typeof body === "string") {
      if (Buffer.byteLength(body) > MAX_BYTES) return res.status(413).end();
      body = JSON.parse(body);
    }
    if (
      !body ||
      typeof body !== "object" ||
      Array.isArray(body) ||
      Object.keys(body).length !== 1 ||
      !Object.prototype.hasOwnProperty.call(body, "event") ||
      !EVENTS.includes(body.event)
    )
      return res.status(400).end();
  } catch (_error) {
    return res.status(400).end();
  }
  if (
    process.env.ANALYTICS_ENABLED !== "true" ||
    (process.env.VERCEL_ENV && process.env.VERCEL_ENV !== "production")
  )
    return res.status(204).end();
  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  // Dedicated Supabase project only. No arbitrary remote endpoints or query strings.
  if (!/^https:\/\/[a-z0-9-]+\.supabase\.co$/.test(url || "") || !key) return res.status(503).end();
  try {
    const result = await fetch(`${url}/rest/v1/rpc/increment_usage_counter`, {
      method: "POST",
      headers: { "Content-Type": "application/json", apikey: key, Authorization: `Bearer ${key}` },
      body: JSON.stringify({ event_name: body.event }),
      signal: AbortSignal.timeout(3000),
      redirect: "error",
    });
    // No headers, IPs, user agents, referrers, error bodies or raw request logs are forwarded.
    return res.status(result.ok ? 204 : 503).end();
  } catch (_error) {
    return res.status(503).end();
  }
};
