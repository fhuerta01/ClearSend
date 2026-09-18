# Release verification

This branch has automated coverage with mocked Office.js and local PostgreSQL. Real Outlook, production Vercel and a hosted Supabase project must also be checked before release.

- Run `npm ci`, `npm run check`, `npm audit` on Node 22 or 24.
- Sideload the local manifest into Outlook web and the target Windows/macOS client. Confirm the add-in is offered only for compose, and that CSS and commands load once without CSP errors.
- Test the production CSP as well (the development server does not apply Vercel headers). MicrosoftAjax must finish initializing `Sys.CultureInfo.InvariantCulture` and `Sys.Res` without `cannotDeserializeInvalidJson`. The SDK currently requires the documented `unsafe-eval` compatibility exception. Browser `unload` policy warnings from Office libraries are separate: verify recipient processing and storage diagnostics rather than treating those warnings as an analytics failure.
- With synthetic addresses, test sorting, name-independent deduplication, internal subdomains, external removal, empty lists and unrecognized/Exchange addresses. Verify BCC-only recipients never move into To or CC.
- Check format blocking in both panel and ribbon. Check that configured step order is respected. Confirm no message is sent by ClearSend.
- Verify panel Undo, changes made manually before Undo, field limits and visible partial-write/recovery errors. Review list contents in Outlook, not just the panel.
- Save preferences, reopen, then test saved-invalid deletion and Restore. Microsoft settings may sync between clients; there is no local-only guarantee.
- Test panel at 320 px and keyboard navigation. Verify drag reordering also has button/keyboard alternatives.
- Deploy to a preview with analytics disabled. Inspect requests to confirm no Vercel Analytics script or counter events.
- Before enabling production counting: apply the reviewed migration to the dedicated Supabase project, check grants, configure server-only secrets, set exact origins, review provider logging/backups and add platform abuse controls.
- On production with synthetic addresses: test opt-in, opt-out, DNT/GPC and network failure. Inspect payloads; only a fixed event name is allowed. Confirm the SQL count increments and no recipient/user metadata is stored.
- Schedule aggregate expiry if retention must run during inactivity. Confirm README, privacy text and hosted revision agree.

Do not infer live Outlook compatibility from XML schema validation alone. Do not infer live analytics activation from a passing local test or a 204 from a disabled endpoint.
