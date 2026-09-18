# Aggregate usage counting

This is the implementation's single setup guide. ClearSend uses a Vercel function plus Supabase **aggregate counters**, not Vercel Web Analytics. An unconfigured deployment sends no events. Once the operator enables production counting, the user preference defaults on if absent; saved opt-outs remain off. No Supabase project or production deployment is created by building the repository.

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

1. Create or choose a **dedicated ClearSend analytics Supabase project** (PostgreSQL 16 or newer). Run `supabase/migrations/202609180001_usage_counters.sql` once through your migration process or Supabase SQL editor as the project administrator (`postgres`). It creates the counter table, a restricted writer role and an RPC function. Do not run the test role-creation commands in an existing Supabase project. If the original script failed with `must be able to SET ROLE "clearsend_counter_writer"`, run `ROLLBACK;` and then execute the entire updated migration. The failed transaction does not leave its table or role behind. An already successful installation does not need to rerun this migration.
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
7. With no saved preference, counting starts automatically. Existing saved `false` preferences stay off; use **Configuration → Share aggregate usage counts** to change your choice. Check the network request is exactly `POST /api/events` with `{"event":"process_click"}` and has no Cookie or Referer header. Successful writes return 204; storage failure returns 503. Confirm the database counter increment. A disabled endpoint also returns 204 without incrementing, so check the table as well.

The client requires an exact matching production origin, an enabled usage preference (on by default only if absent), and no DNT/GPC. Development builds and Vercel previews disable the public origin. The server separately rejects writes from preview deployments. No public secret is used as a pretend authentication mechanism.

### Troubleshoot a 503 response

Open the invocation in Vercel Runtime Logs and find `[ClearSend analytics]`, or inspect the `X-ClearSend-Analytics` response header. These diagnostics use fixed codes without logging secrets, event names, request details or Supabase error bodies.

| Code | Check |
| --- | --- |
| `supabase_url_missing` | Set `SUPABASE_URL` in the deployed Production environment. |
| `supabase_url_invalid` | Use `https://PROJECT_REF.supabase.co`, not the dashboard URL, a database connection string or an API path. |
| `supabase_key_missing` | Set the exact variable name `SUPABASE_SERVICE_ROLE_KEY` in Production. |
| `supabase_key_invalid` | Copy the key as a single line without embedded spaces or control characters. |
| `supabase_auth_rejected` | Check the Legacy `service_role` JWT belongs to the same project as the URL; inspect RPC execution permissions. |
| `supabase_rpc_unavailable` | Verify the migration succeeded in that project and the public RPC is exposed through the Data API. |
| `supabase_write_rejected` | Inspect Supabase database/API logs and the migration's function/table permissions. |
| `supabase_timeout` / `supabase_connection_failed` | Check the project is running and reachable from Vercel. |

Surrounding whitespace in the Supabase URL/key and trailing slashes on the URL are normalized. Other malformed URLs remain rejected. Redeploy after changing environment variables. A 204 response with `stored` means Supabase accepted the increment; `disabled` means no write was attempted. Do not paste credentials into logs or support messages.

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

- User: uncheck **Share aggregate usage counts**. This stops subsequent events; Restore preserves the current choice. Saved opt-outs from earlier versions remain off because they cannot be distinguished from an explicitly disabled preference.
- Operator: set `ANALYTICS_ENABLED=false` to stop database increments; redeploy the server configuration. Clear `CLEARSEND_ANALYTICS_ORIGIN` and rebuild to stop client requests too.
- Local/self-hosted: leave all analytics configuration empty.

## Verification

`npm test` checks the client gates, closed event vocabulary, payload exclusions and API failure behavior. CI applies the migration as a non-superuser with role-creation privileges on disposable PostgreSQL 16 and 17 services, then runs `tests/database.sql` to check restricted ownership, removal of temporary privileges, RLS, permissions, incrementing, retention and rejected event names. Use the [manual release checklist](docs/RELEASE_CHECKLIST.md) for a real Outlook/Vercel deployment.

References: [Supabase functions and execution permissions](https://supabase.com/docs/guides/database/functions), [row-level security](https://supabase.com/docs/guides/database/postgres/row-level-security). Vercel Web Analytics uses visitor hashing and additional visit dimensions, which is why it is not used for this aggregate-only contract: [Vercel's description](https://vercel.com/docs/analytics/privacy-policy).
