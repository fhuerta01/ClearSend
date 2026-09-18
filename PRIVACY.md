# ClearSend privacy policy

Updated: 18 September 2026. This policy describes the source implementation in this revision; a hosted deployment must use the same revision and configuration to match it.

## Core promise

**ClearSend processes recipient lists inside Outlook and does not transmit Outlook data to ClearSend hosting or analytics services.** This includes recipient addresses, display names, message content, internal domains and mailbox/account identifiers.

“Inside your environment” includes Outlook, your Microsoft mailbox and exports you request. It does not mean that Outlook is offline or that Microsoft never synchronizes your data.

## Data handled by the add-in

| Data | Purpose and location | Retention/control |
| --- | --- | --- |
| To, CC and BCC | Read and modify the active draft via Office.js; processing takes place in the add-in runtime. | Outlook manages the draft. ClearSend keeps display/Undo state in memory until the panel closes. |
| Preferences and internal domains | Office.js RoamingSettings in your Microsoft mailbox; Microsoft can sync these between clients. | Restore default settings in Configuration. |
| Saved invalid addresses | Optional, off by default; stored with Microsoft roaming settings, never in ClearSend/Supabase. | Up to 100 entries and 12 KB JSON; delete individual entries, use Delete saved invalid addresses, or Restore. Turning the feature off does not delete existing entries. |
| CSV and clipboard | Explicit user export/copy. | Controlled by your device and your subsequent use of the file/clipboard. |
| Aggregate usage | Enabled by default when no preference is saved on a configured production deployment; fixed event names sent to the app's own `/api/events` endpoint and counted in Supabase. | Only UTC day, event name and count; no event-level history. See below. |

Office requires `ReadWriteItem` permission to edit recipients. That permission is broader than the functions ClearSend uses. The implementation does not read message bodies/attachments or obtain access tokens. RoamingSettings is not a secrets vault; it is part of your Microsoft environment. [Microsoft reference](https://learn.microsoft.com/en-us/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview).

## Optional aggregate analytics

The deployment owner must configure and enable counting. On that production deployment, usage counts are enabled by default when no preference is saved. Turn them off under **Configuration → Share aggregate usage counts**. Existing saved `false` preferences remain off, including preferences saved by earlier versions; Restore preserves this choice. This is an opt-out setting, not a consent prompt. Counting respects Do Not Track and Global Privacy Control and is disabled in local development and preview deployments. Opting out stops subsequent events; already received aggregate counts cannot be tied back to you or individually deleted.

The browser sends only a fixed action name such as `process_click`. It does not send Outlook-derived properties, URLs, referrers, recipient counts, recipient hashes, account data or error text. Requests omit cookies and referrers. There are no analytics user IDs, session IDs, persistent device identifiers or fingerprint hashes. The application does not load Vercel Web Analytics.

The server validates an allowlist and forwards only the event name to a restricted database function. Supabase stores **UTC day + action + count**, incremented atomically. No raw event record is stored. Rows older than 90 UTC calendar days are deleted when new events arrive; if collection stops, the owner must run the cleanup query in the analytics guide. There is no public dashboard or public database read access.

This measures aggregate actions, not people. Blockers, opt-outs, network failures and automated traffic affect accuracy. No retries or persistent event queues are used.

## Hosting and technical metadata

The hosted app downloads code and assets from Vercel and Office.js from Microsoft. When analytics is enabled, the browser also contacts the app's own endpoint. HTTP infrastructure necessarily receives connection metadata such as IP addresses and user-agent headers, and providers may keep security/access logs under their policies. ClearSend's code does not log these values or copy them into counters; it cannot guarantee that providers never retain them.

The browser does not contact Supabase directly. Supabase receives a server-to-server counter request, not the user's connection headers. The deployment owner must review Vercel/Supabase logging, data location, access, retention and backups before enabling analytics. Aggregate database storage does **not** justify a blanket claim of total anonymity for network transport or automatic legal compliance.

When counter storage fails, ClearSend writes a fixed technical error code to server logs (for example, `supabase_url_missing`) and returns that code in a response header. These application diagnostics contain no event names, Outlook data, identifiers, credentials, configuration values or upstream error bodies. The hosting provider may associate the log with its own request metadata as described above.

[Vercel privacy notice](https://vercel.com/legal/privacy-policy) · [Supabase privacy policy](https://supabase.com/privacy) · [Microsoft privacy statement](https://privacy.microsoft.com/privacystatement).

## Your choices

- Turn usage counting off at any time in Configuration. Saved disabled preferences remain off after upgrades and Restore.
- Disable saving invalid addresses and delete previously saved entries separately.
- Use a self-hosted build with analytics configuration empty for no ClearSend analytics requests.
- Audit the open-source code. Hosted administrators control future code updates, so the deployed revision is part of the trust boundary.

For questions, use the contact in [CONTRIBUTING.md](CONTRIBUTING.md). Do not include real recipients, mailbox exports or credentials in public issues.
