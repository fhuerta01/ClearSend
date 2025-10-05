# Privacy Policy

**Last Updated: October 2025**

## Our Privacy Commitment

**ClearSend collects ZERO data. Your email addresses NEVER leave your device.**

ClearSend is designed with privacy as its foundation. All email processing happens **entirely within your Outlook application** (desktop or web browser). No servers, no cloud, no transmission.

## What We Collect

**Absolutely Nothing.**

- ❌ No email addresses
- ❌ No recipient data
- ❌ No personal information
- ❌ No usage analytics
- ❌ No tracking cookies
- ❌ No telemetry data
- ❌ No diagnostics
- ❌ No logs

**We cannot access your data because it never reaches us.**

## How It Works

### Local Processing Only - How Your Data Stays Private

**Every operation runs in your device's memory:**

1. **Data Read**: ClearSend reads email recipients using Office.js API (Microsoft's official API)
2. **Local Processing**: JavaScript executes sorting, validation, deduplication **entirely in your browser/Outlook**
3. **Local Update**: Processed recipients are written back to Outlook using Office.js API
4. **Zero Transmission**: At no point does any data leave your Outlook application

**Technical Details:**
- All JavaScript code runs in your browser's sandboxed environment
- No `fetch()`, `XMLHttpRequest`, or network calls to external services
- Email addresses remain in memory only - never serialized for transmission
- Source code is open for verification: every line of code is reviewable

### Settings Storage

- Your preferences (enabled features, internal domains) are stored locally in Office.js roaming settings
- This data is managed by Microsoft Office and syncs across your Office installations
- We do not have access to this data

## Third-Party Services

### Vercel (Hosting)

- We use Vercel **only** to host static files (HTML, CSS, JavaScript)
- Your browser downloads these files once when the add-in loads
- **Your email data NEVER touches Vercel servers** - processing happens locally in your browser
- Vercel may collect standard web access logs (IP address, browser type) when downloading the add-in files
- No email addresses, recipient data, or personal information is ever sent to Vercel
- Read Vercel's privacy policy: https://vercel.com/legal/privacy-policy

### Microsoft Office

- ClearSend uses Office.js API to interact with Outlook
- Email data is accessed only within your local Outlook session
- Microsoft's privacy policy applies to Office 365 data: https://privacy.microsoft.com

## Data Security

**The most secure data is data that's never transmitted.**

Since all processing is 100% client-side:
- ✅ **Zero Data Transmission** - Your email addresses never leave your device
- ✅ **No Backend Servers** - We don't operate any servers that process or store email data
- ✅ **No Databases** - No databases exist to store your information
- ✅ **No User Accounts** - No registration, login, or authentication required
- ✅ **No Cookies** - No tracking cookies or session management
- ✅ **No Network Calls** - Email data is never sent over the network

**Your data stays in your Outlook application, under your control, always.**

## Open Source

ClearSend is open source (MIT License). You can:
- Review the complete source code on GitHub
- Verify that no data is transmitted externally
- Build and host your own version

## Changes to Privacy Policy

We will update this policy if our practices change. Check the "Last Updated" date above.

## Contact

Questions about privacy? Open an issue on GitHub or email privacy@clearsend.com

## Your Rights

You have complete control over your data because we never access it. You can:
- Uninstall the add-in at any time
- Clear local settings through Office settings
- Review all source code to verify our claims

---

## Privacy Summary

**Three Core Guarantees:**

1. 🔒 **Your email addresses NEVER leave your device** - All processing is 100% local
2. 🚫 **We have NO servers processing your data** - Only static file hosting
3. ✅ **You can verify everything** - Complete source code is publicly available

**Your privacy is not a feature - it's our architecture.**
