const test = require("node:test");
const assert = require("node:assert/strict");
const { createAnalytics } = require("../src/shared/analytics");
const handler = require("../api/events");
const origin = "https://clearsend.vercel.app";
function client(extra = {}) {
  const calls = [];
  return {
    calls,
    analytics: createAnalytics({
      enabled: true,
      origin,
      location: { origin },
      navigator: {},
      fetch: (...args) => {
        calls.push(args);
        return Promise.resolve();
      },
      ...extra,
    }),
  };
}
test("analytics sends only the event name, with no cookies or referrer", () => {
  const { analytics, calls } = client();
  analytics.track("process_click", { email: "secret@example.com" });
  assert.equal(calls.length, 1);
  assert.equal(calls[0][0], "/api/events");
  assert.equal(calls[0][1].body, '{"event":"process_click"}');
  assert.equal(calls[0][1].credentials, "omit");
  assert.equal(calls[0][1].referrerPolicy, "no-referrer");
});
test("no events when disabled or on local/preview origins or DNT/GPC", () => {
  for (const extra of [
    { enabled: false },
    { origin: "" },
    { location: { origin: "http://localhost:3000" } },
    { location: { origin: "https://preview.vercel.app" } },
    { navigator: { doNotTrack: "1" } },
    { navigator: { globalPrivacyControl: true } },
  ]) {
    const { analytics, calls } = client(extra);
    analytics.track("pane_open");
    assert.equal(calls.length, 0);
  }
});
test("unknown events and opt out are blocked", () => {
  const { analytics, calls } = client();
  analytics.track("secret@example.com");
  analytics.setEnabled(false);
  analytics.track("process_click");
  assert.equal(calls.length, 0);
});
test("network errors are swallowed and not retried", async () => {
  let calls = 0;
  const { analytics } = client({
    fetch: () => {
      calls++;
      return Promise.reject(new Error("offline"));
    },
  });
  analytics.track("pane_open");
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(calls, 1);
});
async function request(req = {}) {
  const result = { statusCode: 0, headers: {} };
  await handler(
    {
      method: "POST",
      headers: { origin, "content-type": "application/json" },
      body: { event: "process_click" },
      ...req,
    },
    {
      setHeader(k, v) {
        result.headers[k] = v;
      },
      status(code) {
        result.statusCode = code;
        return this;
      },
      end() {},
    }
  );
  return result;
}
test("endpoint rejects invalid inputs and forwards only the fixed event", async (t) => {
  t.mock.method(console, "warn", () => {});
  const env = { ...process.env };
  const originalFetch = global.fetch;
  const calls = [];
  t.after(() => {
    process.env = env;
    global.fetch = originalFetch;
  });
  Object.assign(process.env, {
    ANALYTICS_ALLOWED_ORIGIN: origin,
    ANALYTICS_ENABLED: "true",
    VERCEL_ENV: "production",
    SUPABASE_URL: "https://example.supabase.co",
    SUPABASE_SERVICE_ROLE_KEY: "test-secret",
  });
  global.fetch = async (...args) => {
    calls.push(args);
    return { ok: true };
  };
  for (const [req, expected] of [
    [{ method: "GET" }, 405],
    [{ headers: { origin: "https://evil.example", "content-type": "application/json" } }, 403],
    [{ headers: { origin, "content-type": "text/plain" } }, 415],
    [{ headers: { origin, "content-type": "application/json", "content-length": "9999" } }, 413],
    [{ body: { event: "process_click", email: "private@example.com" } }, 400],
    [{ body: { event: "private@example.com" } }, 400],
    [{ body: "bad json" }, 400],
    [{ body: [] }, 400],
    [{ body: null }, 400],
  ])
    assert.equal((await request(req)).statusCode, expected);
  assert.equal(calls.length, 0);
  assert.equal((await request()).statusCode, 204);
  assert.equal(calls.length, 1);
  assert.equal(calls[0][1].body, '{"event_name":"process_click"}');
  assert.equal(calls[0][1].headers["x-forwarded-for"], undefined);
  process.env.VERCEL_ENV = "preview";
  assert.equal((await request()).statusCode, 204);
  assert.equal(calls.length, 1);
  process.env.VERCEL_ENV = "production";
  global.fetch = async () => {
    throw new Error("db failed");
  };
  assert.equal((await request()).statusCode, 503);
});

test("storage diagnostics distinguish configuration and upstream failures without leaking data", async (t) => {
  const env = { ...process.env };
  const originalFetch = global.fetch;
  const logs = [];
  const calls = [];
  t.mock.method(console, "warn", (...args) => logs.push(args));
  t.after(() => {
    process.env = env;
    global.fetch = originalFetch;
  });
  const config = {
    ANALYTICS_ALLOWED_ORIGIN: origin,
    ANALYTICS_ENABLED: "true",
    VERCEL_ENV: "production",
    SUPABASE_URL: "https://example.supabase.co",
    SUPABASE_SERVICE_ROLE_KEY: "private-test-key",
  };
  global.fetch = async (...args) => {
    calls.push(args);
    return { ok: true };
  };
  for (const [setting, value, expected] of [
    ["SUPABASE_URL", "", "supabase_url_missing"],
    ["SUPABASE_URL", " \n", "supabase_url_missing"],
    ["SUPABASE_URL", "https://supabase.com/dashboard/project/example", "supabase_url_invalid"],
    ["SUPABASE_URL", "https://example.supabase.co/rest/v1", "supabase_url_invalid"],
    ["SUPABASE_URL", "https://example.supabase.co.evil.test", "supabase_url_invalid"],
    ["SUPABASE_URL", "https://example.supabase.co?secret=private", "supabase_url_invalid"],
    ["SUPABASE_URL", "http://example.supabase.co", "supabase_url_invalid"],
    ["SUPABASE_SERVICE_ROLE_KEY", " \n", "supabase_key_missing"],
    ["SUPABASE_SERVICE_ROLE_KEY", "private\r\nkey", "supabase_key_invalid"],
  ]) {
    Object.assign(process.env, config, { [setting]: value });
    const response = await request();
    assert.equal(response.statusCode, 503);
    assert.equal(response.headers["X-ClearSend-Analytics"], expected);
    assert.deepEqual(logs.at(-1), [`[ClearSend analytics] ${expected}`]);
  }
  assert.equal(calls.length, 0, "invalid configuration must not contact any remote host");

  Object.assign(process.env, config, {
    SUPABASE_URL: " \nhttps://example.supabase.co/ \n",
    SUPABASE_SERVICE_ROLE_KEY: " private-test-key\n",
  });
  const success = await request();
  assert.equal(success.statusCode, 204);
  assert.equal(success.headers["X-ClearSend-Analytics"], "stored");
  assert.equal(calls.length, 1);
  assert.equal(calls[0][0], "https://example.supabase.co/rest/v1/rpc/increment_usage_counter");
  assert.equal(calls[0][1].headers.Authorization, "Bearer private-test-key");

  for (const [status, expected] of [
    [401, "supabase_auth_rejected"],
    [403, "supabase_auth_rejected"],
    [404, "supabase_rpc_unavailable"],
    [500, "supabase_write_rejected"],
  ]) {
    global.fetch = async () => ({
      ok: false,
      status,
      text() {
        assert.fail("Upstream error bodies must not be read or exposed");
      },
    });
    const response = await request();
    assert.equal(response.statusCode, 503);
    assert.equal(response.headers["X-ClearSend-Analytics"], expected);
    assert.deepEqual(logs.at(-1), [`[ClearSend analytics] ${expected}`]);
  }
  for (const [name, expected] of [
    ["TimeoutError", "supabase_timeout"],
    ["TypeError", "supabase_connection_failed"],
  ]) {
    let attempts = 0;
    global.fetch = async () => {
      attempts++;
      const error = new Error("private-test-key private@example.com");
      error.name = name;
      throw error;
    };
    const response = await request();
    assert.equal(response.statusCode, 503);
    assert.equal(response.headers["X-ClearSend-Analytics"], expected);
    assert.deepEqual(logs.at(-1), [`[ClearSend analytics] ${expected}`]);
    assert.equal(attempts, 1, "uncertain writes must never be retried");
  }
  const logCount = logs.length;
  process.env.ANALYTICS_ENABLED = "false";
  global.fetch = () => assert.fail("Disabled analytics must not contact Supabase");
  const disabled = await request();
  assert.equal(disabled.statusCode, 204);
  assert.equal(disabled.headers["X-ClearSend-Analytics"], "disabled");
  assert.equal(logs.length, logCount);
  assert.doesNotMatch(JSON.stringify(logs), /private|process_click|https:/);
});
