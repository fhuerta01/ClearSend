# Analytics Integration Example

This file shows exactly how to integrate the privacy-safe analytics into your taskpane.

## Step 1: Import the analytics module

At the top of `src/taskpane/taskpane.js`, add this import:

### Before:
```javascript
/**
 * ClearSend Task Pane JavaScript
 *
 * PRIVACY GUARANTEE: All processing happens locally in your Outlook client.
 * Your email addresses NEVER leave your device. No servers process your data.
 *
 * This file only uses Office.js API to read/write recipients locally.
 * No network requests transmit any email or recipient data.
 */

/* global Office, document, window, setTimeout, setInterval, clearTimeout, clearInterval, Blob, URL */
```

### After:
```javascript
/**
 * ClearSend Task Pane JavaScript
 *
 * PRIVACY GUARANTEE: All processing happens locally in your Outlook client.
 * Your email addresses NEVER leave your device. No servers process your data.
 *
 * This file only uses Office.js API to read/write recipients locally.
 * No network requests transmit any email or recipient data.
 */

/* global Office, document, window, setTimeout, setInterval, clearTimeout, clearInterval, Blob, URL */

// Privacy-safe analytics (optional - only counts add-in loads)
import { initAnalytics } from './analytics.js';
```

---

## Step 2: Call analytics on Office.onReady

Find the `Office.onReady()` function (around line 106) and add the analytics initialization:

### Before:
```javascript
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialize the task pane
    initializeTaskPane();
  }
});
```

### After:
```javascript
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialize privacy-safe analytics (optional - just counts usage)
    initAnalytics();

    // Initialize the task pane
    initializeTaskPane();
  }
});
```

---

## Complete Example

Here's the complete modified section:

```javascript
/**
 * ClearSend Task Pane JavaScript
 *
 * PRIVACY GUARANTEE: All processing happens locally in your Outlook client.
 * Your email addresses NEVER leave your device. No servers process your data.
 *
 * This file only uses Office.js API to read/write recipients locally.
 * No network requests transmit any email or recipient data.
 */

/* global Office, document, window, setTimeout, setInterval, clearTimeout, clearInterval, Blob, URL */

// Privacy-safe analytics (optional - only counts add-in loads)
import { initAnalytics } from './analytics.js';

/**
 * Configuration Constants
 * ... rest of your code ...
 */

// ... all your existing code ...

/**
 * Office.onReady - Entry point when Office.js is loaded
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialize privacy-safe analytics (optional - just counts usage)
    // This sends a single anonymous ping - no user data collected
    initAnalytics();

    // Initialize the task pane
    initializeTaskPane();
  }
});

// ... rest of your code ...
```

---

## That's It!

Just two small changes:
1. ✅ Import the analytics module
2. ✅ Call `initAnalytics()` in Office.onReady

The analytics will now:
- Send a single POST request to `/api/ping` when the taskpane loads
- Not block or slow down your add-in (fire-and-forget)
- Fail silently if the request fails
- Not collect any personal data

---

## Testing

### 1. Build the project
```bash
npm run build
```

### 2. Run locally
```bash
npm run dev-server
```

### 3. Check the network tab
- Open browser DevTools → Network tab
- Load the taskpane
- Look for: `POST /api/ping`
- Response: `{ "ok": true }`

---

## Disabling Analytics

### Option 1: Don't call it
Just remove the `initAnalytics()` line. Done.

### Option 2: Use the config flag
In `src/taskpane/analytics.js`:

```javascript
const ANALYTICS_CONFIG = {
  ENABLED: false, // Set to false to disable
  // ...
};
```

---

## Privacy Compliance

This implementation:
- ✅ **GDPR compliant** - No personal data collected
- ✅ **CCPA compliant** - No personal information sold or shared
- ✅ **No cookies** - No tracking cookies used
- ✅ **No consent required** - Anonymous usage counting (not personal data)
- ✅ **Open source** - Fully auditable code

---

## Summary

**What you're adding**:
- 1 import statement
- 1 function call
- 0 user data collected

**What you get**:
- A count of how many times your add-in has been loaded
- No personal information
- No tracking

**Privacy impact**:
- Minimal - just a counter
- Fully disclosed in PRIVACY.md
- Can be disabled with 1 line change
