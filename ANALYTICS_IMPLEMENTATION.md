# Analytics implementation

See [ANALYTICS_README.md](ANALYTICS_README.md) for the current setup, exact event definitions, queries and limitations.

The implementation is `src/shared/analytics.js` → `api/events.js` → `supabase/migrations/202609180001_usage_counters.sql`. Only fixed event names are accepted; PostgreSQL stores one count per UTC day/action. Vercel Web Analytics, Vercel KV and in-memory usage counters are not used.

Earlier instructions for `api/ping.js`, `api/ping-persistent.js` and automatic page-view tracking described removed or nonexistent files and do not apply. Analytics is already integrated; do not add a second initialization call or click listener.
