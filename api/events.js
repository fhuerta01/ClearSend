// The only public API accepts one fixed event name, never Outlook data.
const { EVENTS } = require("../src/shared/analytics-events");
const MAX_BYTES = 80;
function unavailable(res, code) {
  // Only fixed diagnostic codes: never log requests, configuration or exceptions.
  // eslint-disable-next-line no-console
  console.warn(`[ClearSend analytics] ${code}`);
  res.setHeader("X-ClearSend-Analytics", code);
  return res.status(503).end();
}
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
  ) {
    res.setHeader("X-ClearSend-Analytics", "disabled");
    return res.status(204).end();
  }
  const url = (process.env.SUPABASE_URL || "").trim().replace(/\/+$/, "");
  const key = (process.env.SUPABASE_SERVICE_ROLE_KEY || "").trim();
  // Dedicated Supabase project only. No arbitrary remote endpoints or query strings.
  if (!url) return unavailable(res, "supabase_url_missing");
  if (!/^https:\/\/[a-z0-9-]+\.supabase\.co$/.test(url))
    return unavailable(res, "supabase_url_invalid");
  if (!key) return unavailable(res, "supabase_key_missing");
  if (/[^\x21-\x7e]/.test(key)) return unavailable(res, "supabase_key_invalid");
  try {
    const result = await fetch(`${url}/rest/v1/rpc/increment_usage_counter`, {
      method: "POST",
      headers: { "Content-Type": "application/json", apikey: key, Authorization: `Bearer ${key}` },
      body: JSON.stringify({ event_name: body.event }),
      signal: AbortSignal.timeout(3000),
      redirect: "error",
    });
    // No headers, IPs, user agents, referrers, error bodies or raw request logs are forwarded.
    if (!result.ok) {
      if (result.status === 401 || result.status === 403)
        return unavailable(res, "supabase_auth_rejected");
      if (result.status === 404) return unavailable(res, "supabase_rpc_unavailable");
      return unavailable(res, "supabase_write_rejected");
    }
    res.setHeader("X-ClearSend-Analytics", "stored");
    return res.status(204).end();
  } catch (error) {
    return unavailable(
      res,
      error?.name === "TimeoutError" ? "supabase_timeout" : "supabase_connection_failed"
    );
  }
};
