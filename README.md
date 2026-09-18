# ClearSend — review Outlook recipients before you send

[![CI](https://github.com/fhuerta01/ClearSend/actions/workflows/ci.yml/badge.svg)](https://github.com/fhuerta01/ClearSend/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**Clean up long To, CC and BCC lists without sending your recipient data to another service.** ClearSend is a free, open-source Outlook add-in for people who regularly email teams, clients or mixed internal/external groups.

Sort recipients, remove duplicates, review address formats and apply your own internal-domain rules before sending. ClearSend processes lists inside Outlook; it does not upload recipients, message content or internal domains to ClearSend hosting or analytics services.

## What it helps you do

| Task | Behavior |
| --- | --- |
| Remove duplicates | Compare email addresses regardless of display name or letter case; retain the first occurrence in To, then CC, then BCC. Never move BCC-only recipients into visible fields. |
| Make lists easier to review | Sort by display name/email; optionally prioritize up to three internal domains and their subdomains. |
| Check address formats | Block processing when unrecognized formats are present. This does **not** check mailbox existence or delivery. |
| Keep an email internal | Explicitly enable removal of addresses outside your configured domains. Review the result before sending. |
| Control processing | Reorder steps by dragging or with the move button (Shift moves down). The format check always runs before changes. |
| Recover or export | Undo the last panel edit while the panel is open; export CSV locally. Saved invalid-address lists are optional, bounded and deletable. |
| Run a quick action | Use the ribbon's Quick clean to apply the same preferences. Use the panel when you need Undo. |

ClearSend is a recipient-review helper. It does not send emails, intercept Send, verify identities or replace organizational data-loss prevention policies.

## Install

Use a current Outlook desktop or web client with a Microsoft 365/Exchange mailbox and a modern webview. Availability of custom add-ins depends on your account and administrator. Mobile and legacy Internet Explorer webviews are not supported.

1. Download [manifest.prod.xml](https://raw.githubusercontent.com/fhuerta01/ClearSend/main/manifest.prod.xml).
2. In Outlook, open the add-in management page and choose **My add-ins → Add a custom add-in → Add from file**, where available.
3. Select the manifest, compose a message and open **ClearSend** from the Apps/add-ins menu or ribbon.
4. Review **Configuration**, set internal domains if needed, then select **Process destination fields**.

The production manifest loads the hosted application from `https://clearsend.vercel.app`. A manifest from a development branch does not publish that branch's code: the maintainer must deploy it first. Hosted updates take effect without downloading the complete repository.

See Microsoft's [sideloading instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing) if your Outlook menus differ.

## Privacy: your Outlook data stays in your environment

- **Recipient processing:** inside the Outlook add-in. ClearSend does not read message bodies or attachments, request mailbox tokens, or send recipient data to Vercel/Supabase.
- **Microsoft mailbox:** Outlook manages drafts and recipients. Preferences, internal domains and optional saved invalid-address lists use Office.js RoamingSettings and can sync through your Microsoft account. This is **not device-only storage**.
- **Local exports:** CSV downloads and clipboard copies happen only when requested. You control the resulting files and clipboard.
- **Optional usage counts:** disabled by default. When the deployment supports it and you opt in, only fixed action names are sent to a same-origin endpoint. Supabase stores a count per action per UTC day—no addresses, names, domains, user IDs, session IDs, cookies or fingerprint hashes.
- **Hosting:** the hosted app and Microsoft Office.js still require network requests. Providers necessarily receive technical connection metadata. ClearSend does not promise that the device makes no network requests or that provider logs contain no IP addresses.

Read the complete [privacy policy](PRIVACY.md), [security boundaries](SECURITY.md) and [analytics setup](ANALYTICS_README.md). Review the source or self-host for control over the deployed application.

## Develop or self-host

Use Node.js **22.22.2+ or 24 LTS** and npm. `.nvmrc` selects 24.

```sh
git clone https://github.com/fhuerta01/ClearSend.git
cd ClearSend
npm ci
npm run check
npm run dev-server
```

The development server uses `https://localhost:3000`. Follow the Office development certificate prompt, then sideload **manifest.xml** for local testing. `npm start` is an optional Office debugging helper; platform/account support varies. Local development never emits ClearSend analytics.

For your own HTTPS host, build with `npm run build`, serve `dist/`, and replace the production host in `webpack.config.js`, both manifests and any analytics origin configuration. The build copies the CSS, icons and manifests and inserts each script once. The Office.js library is loaded from Microsoft; self-hosting is not a guarantee of offline operation.

### Commands

| Command | Purpose |
| --- | --- |
| `npm run check` | Lint, regression tests, production build and both manifest validations. |
| `npm test` | Processing, Office failure recovery, privacy contract and UI regression tests. |
| `npm audit` | Check dependencies, including development tooling. |
| `npm run dev-server` | Local HTTPS development server. |
| `npm run build` | Static production bundle; Vercel serves `api/events.js` separately. |

With focus inside the panel: **Ctrl+Alt+Q** processes configured steps, **Ctrl+Alt+S** sorts, **Ctrl+Alt+D** deduplicates and **Ctrl+Alt+V** checks formats without changing recipients. These shortcuts do not open the panel globally.

## Behavior and limits

- At most **100 recipients per changed field**, checked before writes, for portability across Outlook clients. Unchanged fields are not rewritten.
- If a field update fails, ClearSend attempts to restore all attempted fields. Outlook offers no transaction across To/CC/BCC; a failed recovery is shown prominently. Always review recipients before sending.
- Changes made in Outlook after a snapshot cause processing/Undo to stop instead of knowingly overwriting that snapshot. Edits during Office's asynchronous writes remain a host limitation.
- Address checks intentionally support common SMTP formats. Exchange aliases, distribution lists, internationalized or quoted addresses may need resolving in Outlook or disabling the format check. No DNS or delivery queries are made.
- Saved invalid addresses are limited to 100 entries and a 12 KB JSON budget. Turning off saving stops additions; **Delete saved invalid addresses** or **Restore** deletes the existing list from the add-in's Microsoft settings.
- Undo covers the last successful panel modification only, and is lost when the panel closes. Ribbon Quick clean has no persistent undo history.

## Usage measurement

GitHub's Traffic page measures repository visits and clones, **not add-in users or clicks**. The optional aggregate counter distinguishes process activations, successful processing, blocked operations and errors. No user identity is available, so it cannot report unique users or individual conversion journeys. Counts can be reduced by opt-outs, blockers and failed delivery, or inflated by automated requests; they are product usage estimates, not billing records.

See [ANALYTICS_README.md](ANALYTICS_README.md) for the exact event vocabulary, SQL queries and deployment requirements. Vercel Web Analytics is not loaded by this implementation.

## Architecture and contributing

- `src/taskpane/processors.js`: pure recipient processing shared by panel and ribbon.
- `src/shared/recipients.js`: Office reads, bounded writes, conflict checks and rollback.
- `src/shared/settings.js`: normalized Microsoft roaming preferences.
- `src/shared/analytics.js` → `api/events.js` → Supabase RPC: optional aggregate-only usage counts.
- `tests/`: mocked Outlook/UI regressions and PostgreSQL permissions/counter checks.

[Contributions](CONTRIBUTING.md), reproducible bug reports and documentation improvements are welcome. Use synthetic addresses in issues and screenshots. [MIT license](LICENSE).
