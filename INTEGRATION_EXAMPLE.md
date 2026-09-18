# Analytics integration contract

Panel and ribbon integration is already included. Setup is documented in [ANALYTICS_README.md](ANALYTICS_README.md).

```js
const { createAnalytics } = require('./src/shared/analytics');
const analytics = createAnalytics({
  enabled: false, // Keep inactive until normalizeSettings loads the saved preference.
  origin: '', // Public, exact production origin from the build configuration.
});
analytics.track('process_click');
```

The tracker accepts only event names from `src/shared/analytics-events.js`. Never pass an event object, mailbox state, recipient data, error messages, URLs or arbitrary properties. Do not instrument the DOM globally or derive names from button text. Unknown events are discarded. Opt-out, DNT/GPC, local builds and preview deployments send no events.

This example does nothing with the defaults above. The application loads the saved settings and calls `setEnabled(settings.analyticsEnabled)` before tracking. `normalizeSettings` defaults an absent preference to true and preserves explicit false values. Production configuration and DNT/GPC gates still apply; see the setup guide.
