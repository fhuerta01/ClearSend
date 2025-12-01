# ClearSend Analytics - Quick Start Guide

## 🎯 Goal

Add **minimal, privacy-safe usage tracking** to understand how many people use ClearSend, without collecting any personal data.

## 📊 What You'll Track

Just one number: **"How many times has the add-in been loaded?"**

That's it. No names, emails, IPs, timestamps, or any identifying information.

---

## 🚀 Quick Decision Tree

**Choose your approach:**

```
Do you want analytics?
│
├─ No → Do nothing. Your add-in is 100% private already.
│
└─ Yes → Which level of effort?
    │
    ├─ Zero effort → Option 1: Vercel Web Analytics (just enable in dashboard)
    │
    ├─ Minimal effort → Option 2: Simple counter (resets periodically)
    │
    └─ Best tracking → Option 3: Persistent counter (accurate counts)
```

---

## Option 1: Vercel Web Analytics ⭐ Recommended for Beginners

### What it is
Vercel's built-in privacy-friendly page view tracking.

### Pros & Cons
✅ Zero code changes
✅ Free on hobby tier
✅ GDPR compliant
✅ Beautiful dashboard
⚠️ Tracks page views (not usage events)

### How to set up
1. Go to [Vercel Dashboard](https://vercel.com/dashboard)
2. Select your project
3. Click **Analytics** → **Enable**
4. Done!

### Cost
**Free** on hobby tier

### Privacy
- No cookies
- No personal data
- Compliant with GDPR/CCPA

---

## Option 2: Simple In-Memory Counter

### What it is
A tiny serverless function that increments a number each time the add-in loads.

### Pros & Cons
✅ Super simple (no database)
✅ Absolutely zero user data
⚠️ Resets every 5-15 minutes (cold starts)
⚠️ Only gives "order of magnitude" estimates

### How to set up

1. **Files are already created** in `api/ping.js` and `src/taskpane/analytics.js`

2. **Integrate into taskpane** (see `INTEGRATION_EXAMPLE.md`):
   ```javascript
   // Add to top of taskpane.js
   import { initAnalytics } from './analytics.js';

   // Add to Office.onReady()
   Office.onReady((info) => {
     if (info.host === Office.HostType.Outlook) {
       initAnalytics(); // ← Add this line
       initializeTaskPane();
     }
   });
   ```

3. **Build and deploy**:
   ```bash
   npm run build
   vercel --prod
   ```

4. **Done!** The add-in will now send a ping each time it loads.

### Cost
**Free** (no dependencies or services)

### Privacy
- Zero personal data
- No IP logging
- Just a simple counter

---

## Option 3: Persistent Counter (KV Database) ⭐ Recommended for Production

### What it is
Same as Option 2, but uses Vercel KV (Redis) to persist the counter permanently.

### Pros & Cons
✅ Accurate counts (never resets)
✅ Still zero user tracking
✅ Can view count anytime
⚠️ Requires Vercel KV setup
⚠️ One dependency to install

### How to set up

1. **Install Vercel KV**:
   ```bash
   npm install @vercel/kv
   ```

2. **Create KV database**:
   - Go to [Vercel Dashboard](https://vercel.com/dashboard) → Storage
   - Click **Create Database** → Select **KV**
   - Name it "clearsend-analytics"
   - Select your project
   - Click **Create**
   - Environment variables auto-configured ✓

3. **Use the persistent API**:
   ```bash
   # Delete the simple version
   rm api/ping.js

   # Rename persistent version
   mv api/ping-persistent.js api/ping.js
   ```

4. **Integrate into taskpane** (same as Option 2):
   ```javascript
   // Add to top of taskpane.js
   import { initAnalytics } from './analytics.js';

   // Add to Office.onReady()
   Office.onReady((info) => {
     if (info.host === Office.HostType.Outlook) {
       initAnalytics(); // ← Add this line
       initializeTaskPane();
     }
   });
   ```

5. **Build and deploy**:
   ```bash
   npm run build
   vercel --prod
   ```

6. **View your usage count**:
   ```bash
   vercel kv get clearsend:usage:count
   ```
   Or: Vercel Dashboard → Storage → Your KV → Browse Data

### Cost
**Free tier**:
- 256 MB storage
- 100,000 reads/month
- 100,000 writes/month

(Way more than enough for usage tracking)

### Privacy
- Zero personal data
- No IP logging
- Just a persistent counter

---

## 📋 Summary Comparison

| Feature | Option 1 (Vercel) | Option 2 (Simple) | Option 3 (Persistent) |
|---------|------------------|-------------------|----------------------|
| **Code changes** | None | Minimal (2 lines) | Minimal (2 lines) |
| **Dependencies** | None | None | @vercel/kv |
| **Setup time** | 30 seconds | 5 minutes | 10 minutes |
| **Accuracy** | Page views | Approximate | Exact |
| **Data persistence** | Permanent | Temporary | Permanent |
| **Cost (hobby)** | Free | Free | Free |
| **Privacy** | High | Maximum | Maximum |
| **Recommended for** | Quick start | Minimal effort | Production use |

---

## 🔐 Privacy Guarantee

**All options collect ZERO personal data:**

❌ No user names or emails
❌ No IP addresses stored
❌ No timestamps
❌ No browser/device info
❌ No location data
❌ No cookies
❌ No session tracking

✅ Just a count: "The add-in was loaded X times"

---

## 🧪 Testing

After integration:

1. **Build**:
   ```bash
   npm run build
   ```

2. **Test locally**:
   ```bash
   npm run dev-server
   ```

3. **Check DevTools**:
   - Open browser DevTools → Network tab
   - Load taskpane
   - Look for `POST /api/ping`
   - Should return `{"ok": true}`

4. **Deploy**:
   ```bash
   vercel --prod
   ```

---

## ❌ Disabling Analytics

### Permanently remove:
Don't call `initAnalytics()` in your code.

### Temporary disable:
In `src/taskpane/analytics.js`:
```javascript
const ANALYTICS_CONFIG = {
  ENABLED: false, // ← Change to false
  // ...
};
```

---

## 📚 Documentation Files

- `ANALYTICS_IMPLEMENTATION.md` - Detailed implementation guide
- `INTEGRATION_EXAMPLE.md` - Code examples
- `PRIVACY.md` - Updated privacy policy (includes analytics disclosure)
- `README.md` - Updated main README

---

## ❓ FAQ

### Do I need to disclose this to users?
**Yes.** Already done - see updated `PRIVACY.md`.

### Is this GDPR compliant?
**Yes.** No personal data = no consent needed (anonymous counting).

### Can I track specific events (like button clicks)?
**Yes!** Use the same `sendUsagePing()` function from `analytics.js` anywhere in your code:
```javascript
import { sendUsagePing } from './analytics.js';

// On button click
document.getElementById('myButton').addEventListener('click', () => {
  sendUsagePing(); // Fire-and-forget ping
});
```

### What if Vercel KV isn't free anymore?
Switch to Option 2 (simple counter) or Option 1 (Vercel Web Analytics).

### Can I see who is using the add-in?
**No.** By design. Zero identification.

---

## 🎉 Recommended Path

**For your use case:**

1. ✅ Start with **Option 1** (Vercel Web Analytics) - Enable it now in 30 seconds
2. ✅ Add **Option 3** (Persistent Counter) for accurate counts - Takes 10 minutes
3. ✅ Use both together - Vercel for page views, KV counter for specific events

**Next steps:**
1. Choose your option
2. Follow the setup steps above
3. Build and deploy
4. Verify it works (check Network tab)
5. Done!

---

## 📞 Questions?

- Check `ANALYTICS_IMPLEMENTATION.md` for detailed instructions
- Check `INTEGRATION_EXAMPLE.md` for code examples
- Open an issue on GitHub

---

**Privacy-first analytics. Simple. Transparent. Respectful.**
