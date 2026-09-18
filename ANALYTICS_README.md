# Aggregate usage counting

This is the implementation's single setup guide. ClearSend uses a Vercel function plus Supabase **aggregate counters**, not Vercel Web Analytics. It is disabled by default; no Supabase project or production deployment is created by building the repository.

## What each metric means

| Event | Meaning |
| --- | --- |
| `pane_open` | One successful panel initialization, if the user had already opted in. |
| `process_click` | A panel processing activation accepted while idle, including the documented keyboard shortcuts. |
| `process_success` | Processing completed successfully, including a no-op when no changes were needed. |
| `process_blocked` | No enabled processing steps, or the address-format precondition stopped processing. |
| `process_error` | Processing or a required Office/settings operation failed; no error text is sent. |
| `quick_clean_click/success/blocked/error` | The equivalent events for the ribbon command. |
| `undo_click` | Undo requested while available and idle, whether or not it succeeds. |
| `export_click`, `export_invalid_click` | CSV export requested, not proof of a saved file. |
| `remove_click`, `copy_click` | Recipient removal or clipboard copy requested. No recipient properties. |
| `refresh_click`, `settings_click` | Manual refresh or Configuration opened. |

Do not add outcome counters together as extra clicks. `process_click + quick_clean_click` is the total of process activations received. These are activations, not unique users or literal mouse-only clicks. A blocked in-flight double click is ignored. No events come from polling, rendering or opening recipient lists.

## Deploy with Supabase

1. Create or choose a **dedicated ClearSend analytics Supabase project**. Run `supabase/migrations/202609180001_usage_counters.sql` once through your migration process or Supabase SQL editor as the project administrator. It creates the counter table, a restricted writer role and an RPC function. Do not run the test role-creation commands in an existing Supabase project.
2. Inspect the table and RPC grants. `anon` and `authenticated` must have neither table access nor RPC execution. Only the server-side `service_role` key calls `increment_usage_counter`. Never expose that key in a browser variable or a `VITE_`/`NEXT_PUBLIC_` variable.
3. In the Vercel project's **Production** environment, set:

   ```text
   CLEARSEND_ANALYTICS_ORIGIN=https://clearsend.vercel.app
   ANALYTICS_ALLOWED_ORIGIN=https://clearsend.vercel.app
   ANALYTICS_ENABLED=true
   SUPABASE_URL=https://YOUR_PROJECT_REF.supabase.co
   SUPABASE_SERVICE_ROLE_KEY=<server-only service_role JWT>
   ```

   Use the exact HTTPS origin, with no trailing slash, query or path. A different host requires changing both origin values and the manifests/host settings. `CLEARSEND_ANALYTICS_ORIGIN` is public build configuration; the other values are read only by the server. `.env.example` is documentation, not automatically loaded by webpack.
4. Use Vercel's Node 22 or 24 runtime, `npm run build`, output directory `dist`. Keep `api/events.js` at repository root. A static-only host does not run the counter API. Do not enable Vercel Web Analytics or a separate tracking script for this design.
5. Review provider logging, backups, region and access policies. Keep raw request bodies out of logs. Add platform-level abuse controls for `/api/events`; a strict Origin check is useful for browsers, **not authentication against scripts**. Account for any provider metadata processing in your published policy.
6. Deploy the reviewed branch, then verify with synthetic addresses in Outlook. This repository change alone does not enable production collection.
7. Opt in under Configuration. Check the network request is exactly `POST /api/events` with `{"event":"process_click"}` and has no Cookie or Referer header. Successful writes return 204; storage failure returns 503. Confirm the database counter increment. A disabled endpoint also returns 204 without incrementing, so check the table as well.

The client requires an exact matching production origin, explicit user opt-in, and no DNT/GPC. Development builds and Vercel previews disable the public origin. The server separately rejects writes from preview deployments. No public secret is used as a pretend authentication mechanism.

## Query the counts

Run queries in the Supabase SQL editor as an administrator; no public dashboard endpoint is provided.

```sql
-- Received process activations per UTC day; panel and ribbon combined.
select day, sum(count) as process_activations
from public.usage_counters
where event in ('process_click', 'quick_clean_click')
group by day order by day;

-- Inspect outcomes separately from activations.
select event, sum(count) as total
from public.usage_counters
group by event order by event;

-- Daily maintenance, including after analytics has been turned off.
delete from public.usage_counters
where day < (now() at time zone 'UTC')::date - 89;
```

Incoming events trigger the same retention cleanup. Schedule that deletion in your database operations if you require expiry even during inactivity. Provider backups may retain historical aggregates longer, according to their configuration.

## Accuracy and anonymity boundaries

An atomic SQL upsert prevents lost increments under concurrency and survives Vercel cold starts. The store has only `day`, `event`, `count`: there is no row for a person or a single event. It contains no IP, user-agent, recipient data, identifiers, recipient hashes or precise timestamps.

It is **not an exact count of all use**: opt-outs, blockers, offline clients and requests lost during closing reduce totals. Requests can also be automated or spoofed. No retry is performed after uncertain delivery, avoiding retry-based double counting. Success and click events are independent and may be delivered on different UTC days or one may be lost. Ratios are estimates, not per-user funnels. An in-memory server counter would reset on cold starts and is intentionally not used.

Network providers necessarily receive connection metadata. “Anonymous” describes the application's aggregate storage, not a guarantee that no infrastructure provider sees an IP address. See [PRIVACY.md](PRIVACY.md).

## Disable or remove

- User: uncheck **Share anonymous action counts**. Defaults and Restore also turn it off.
- Operator: set `ANALYTICS_ENABLED=false` to stop database increments; redeploy the server configuration. Clear `CLEARSEND_ANALYTICS_ORIGIN` and rebuild to stop client requests too.
- Local/self-hosted: leave all analytics configuration empty.

## Verification

`npm test` checks the client gates, closed event vocabulary, payload exclusions and API failure behavior. CI runs `tests/database.sql` against a disposable PostgreSQL service to check permissions, incrementing, retention and rejected event names. Use the [manual release checklist](docs/RELEASE_CHECKLIST.md) for a real Outlook/Vercel deployment.

References: [Supabase functions and execution permissions](https://supabase.com/docs/guides/database/functions), [row-level security](https://supabase.com/docs/guides/database/postgres/row-level-security). Vercel Web Analytics uses visitor hashing and additional visit dimensions, which is why it is not used for this aggregate-only contract: [Vercel's description](https://vercel.com/docs/analytics/privacy-policy).
