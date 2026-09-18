# Security boundaries and reporting

Report vulnerabilities privately to the maintainer contact in [CONTRIBUTING.md](CONTRIBUTING.md). Do not put real recipients or credentials in public issues.

## Boundaries

- Outlook's `ReadWriteItem` permission is needed for recipient writes; ClearSend uses no body, attachment or mailbox-token APIs.
- Recipient transformations are shared pure functions. Only `src/shared/recipients.js` reads/writes recipient fields through Office.js.
- Mailbox-derived values render as text or input values, not HTML. There are no inline JavaScript event attributes. CSV cells are quoted and formula-leading values are prefixed with an apostrophe.
- The analytics API accepts only a known event name. The browser cannot provide a date, count, user ID, domain or arbitrary metadata. Secrets remain server-side. The database denies public reads/writes and RPC execution.
- Analytics defaults on when no preference is saved on a configured production deployment, with opt-out and DNT/GPC respected. It remains separate from recipient processing. Infrastructure connection metadata is outside the aggregate counter's data model; see [PRIVACY.md](PRIVACY.md).
- Security headers limit resource sources without blocking Outlook's embedded frame. Do not add `X-Frame-Options: DENY` or `SAMEORIGIN`: this app must run in an Outlook iframe. Test CSP in supported Outlook clients before a production release.
- CSP permits `'unsafe-eval'` because the MicrosoftAjax library loaded by Office.js uses `eval` to initialize `Sys.CultureInfo`. Blocking it reproduces `cannotDeserializeInvalidJson` during Office startup. This exception allows string evaluation by all permitted scripts; CSP cannot restrict it to Microsoft alone. Script sources remain limited to this origin and Microsoft's Office CDN, and inline JavaScript remains blocked. ClearSend does not evaluate mailbox content as code. Revisit this exception when the Office SDK no longer needs it.

## Limits

Office does not provide a transaction across recipient fields. ClearSend validates sizes, checks that the snapshot is current and attempts rollback after partial failure. It cannot prevent users or another add-in from editing fields during the asynchronous write. Review lists before sending, especially after an error.

The hosted application's publisher can change future JavaScript, and Office.js is supplied by Microsoft. Open source makes a revision auditable; it does not eliminate supply-chain or future deployment trust. The manifests do not provide a send-time security gate.

The anonymous public event endpoint is not authenticated. A script can spoof an Origin header. Deployment-level abuse controls are needed before interpreting counts or incurring database usage. Do not add IP-based identity or tracking to the database to make counts look exact.

## Dependency maintenance

Use `npm ci`, `npm run check` and `npm audit`. CI treats lint, test, build, manifest and audit failures as failures. Node 22 and 24 are the supported development runtimes. Build/debug packages are development dependencies, not browser runtime packages.

`adm-zip` is overridden to `^0.6.1` because Office development tooling otherwise resolves vulnerable versions. Recheck upstream compatibility and the override on tool upgrades. Automated validation does not certify the absence of all vulnerabilities.
