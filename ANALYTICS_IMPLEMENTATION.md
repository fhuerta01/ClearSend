# ClearSend Privacy-First Analytics Implementation Guide

## Overview

This document explains how to add **minimal, privacy-safe usage tracking** to ClearSend while maintaining your strong privacy commitments.

## Privacy Guarantees

✅ **No user identification** - No user IDs, emails, or names
✅ **No IP address logging** - Standard HTTP headers only (not stored)
✅ **No email data** - Email addresses never leave the client
✅ **No tracking cookies** - No cookies or session storage
✅ **No timestamps** - No usage patterns tracked
✅ **No metadata** - No browser, device, or location data
✅ **Just a counter** - Only counts "add-in loaded"

**What you get**: A simple count of how many times the add-in has been loaded.

---

## Implementation Options

You have **three options** (choose one):

### Option 1: Vercel Web Analytics (Easiest - Recommended)

**What it is**: Vercel's built-in privacy-friendly analytics

**Pros**:
- ✅ Zero code changes needed
- ✅ Free on hobby tier
- ✅ GDPR/CCPA compliant
- ✅ No cookies
- ✅ Privacy-focused by design
- ✅ Beautiful dashboard

**Cons**:
- ⚠️ Tracks page views (not usage events)
- ⚠️ Less control over what's measured

**How to enable**:
1. Go to your Vercel project dashboard
2. Click **Analytics** tab
3. Click **Enable Web Analytics**
4. Done!

**Cost**: Free on hobby tier

---

### Option 2: Simple In-Memory Counter (Simplest Code)

**What it is**: A Vercel serverless function with in-memory counter

**Pros**:
- ✅ Extremely simple (12 lines of code)
- ✅ No dependencies
- ✅ No database needed
- ✅ Absolutely zero tracking

**Cons**:
- ⚠️ Counter resets on cold starts (~every 5-15 minutes)
- ⚠️ Only useful for "order of magnitude" estimates
- ⚠️ Not accurate for real usage counts

**Files created**:
- `api/ping.js` - The serverless function

**Setup**:
1. Files already created (see below)
2. Import analytics in taskpane.js (see below)
3. Deploy to Vercel

**Cost**: Free (no dependencies)

---

### Option 3: Persistent Counter with Vercel KV (Most Accurate)

**What it is**: A Vercel serverless function with Redis-backed persistent counter

**Pros**:
- ✅ Accurate usage counts (never resets)
- ✅ Still zero user tracking
- ✅ Simple atomic increment
- ✅ View count anytime via Vercel dashboard

**Cons**:
- ⚠️ Requires Vercel KV setup
- ⚠️ One additional dependency (@vercel/kv)

**Files created**:
- `api/ping-persistent.js` - The serverless function with KV

**Setup**:

1. **Install Vercel KV dependency**:
   ```bash
   npm install @vercel/kv
   ```

2. **Create KV database**:
   - Go to Vercel dashboard → Storage → Create Database
   - Select "KV" (Redis)
   - Name it "clearsend-analytics" (or anything)
   - Select your project
   - Click "Create"
   - Environment variables are auto-configured

3. **Rename the API file**:
   ```bash
   mv api/ping-persistent.js api/ping.js
   ```

4. **Import analytics in taskpane.js** (see below)

5. **Deploy to Vercel**

6. **View your usage count** (optional):
   - Vercel Dashboard → Storage → Your KV Database
   - Browse data → View `clearsend:usage:count`

**Cost**:
- Free tier: 256 MB storage, 100,000 reads/month, 100,000 writes/month
- More than enough for usage tracking

---

## Client-Side Integration

Regardless of which option you choose (2 or 3), you need to add the analytics call to your taskpane:

### Step 1: Import the analytics module

Add this import at the top of `src/taskpane/taskpane.js`:

```javascript
import { initAnalytics } from './analytics.js';
```

### Step 2: Call it on initialization

Find the `Office.onReady()` block (around line 106) and add the analytics call:

```javascript
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {

    // Initialize privacy-safe analytics (optional - just counts usage)
    initAnalytics();

    // ... rest of your initialization code
    initializeTaskPane();
  }
});
```

### Step 3: Build and test

```bash
npm run build
```

That's it! The analytics will now send a single ping when the taskpane loads.

---

## Disabling Analytics

If you want to disable analytics entirely:

### Option A: Remove the code

Just don't add the `initAnalytics()` call.

### Option B: Use the config flag

In `src/taskpane/analytics.js`, change:

```javascript
const ANALYTICS_CONFIG = {
  ENABLED: false, // Changed from true to false
  // ...
};
```

---

## Testing Analytics

### Test the ping is sent:

1. **Development (localhost)**:
   ```bash
   npm run dev-server
   ```
   - Open browser DevTools → Network tab
   - Load the taskpane
   - Look for POST request to `/api/ping`
   - Should return `{ "ok": true }`

2. **Production (Vercel)**:
   - Deploy to Vercel: `vercel --prod`
   - Load the add-in
   - Check Vercel logs: Dashboard → Your Project → Logs
   - Should see POST /api/ping requests

### View usage count (Option 3 only):

```bash
# Install Vercel CLI if needed
npm i -g vercel

# Login to Vercel
vercel login

# Pull KV data (requires Vercel CLI 33+)
vercel kv get clearsend:usage:count
```

Or view in Vercel dashboard: Storage → Your KV → Browse Data

---

## Privacy Policy Update

You MUST update your PRIVACY.md to disclose this minimal tracking. See the updated PRIVACY.md file (next step).

Key points to include:
- What is collected (just a count)
- What is NOT collected (no personal data)
- Purpose (understand usage volume)
- How to disable (remove the call or set ENABLED: false)

---

## Recommended Choice

**For your use case, I recommend**:

1. **Start with Option 1 (Vercel Web Analytics)** - Zero effort, immediate value
2. **Add Option 3 (Persistent Counter)** if you want to track "Process" button clicks or other specific events

You can use both together - Vercel Analytics for page views, and your custom counter for specific usage events.

---

## Files Created

The following files have been created for you:

```
api/
  ├── ping.js                    # Simple in-memory counter (Option 2)
  └── ping-persistent.js         # Persistent KV counter (Option 3)

src/taskpane/
  └── analytics.js               # Client-side analytics utility
```

**Next steps**:
1. Choose your option (1, 2, or 3)
2. Follow the setup steps above
3. Update PRIVACY.md (see next commit)
4. Test locally
5. Deploy to Vercel

---

## Example Usage Stats

With Option 3 (Persistent Counter), you can query your usage:

```bash
# How many times has the add-in been loaded?
vercel kv get clearsend:usage:count
# Returns: 1547
```

That's all you get - a number. No who, when, where, or what. Just "how many times."

**This is the most privacy-respecting analytics possible while still getting useful data.**
