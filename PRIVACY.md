# Privacy Policy

**Last Updated: October 2025**

## Our Privacy Commitment

**ClearSend collects ZERO data. Your email addresses NEVER leave your device.**

ClearSend is designed with privacy as its foundation. All email processing happens **entirely within your Outlook application** (desktop or web browser). No servers, no cloud, no transmission.

## What We Collect

### Email and Personal Data

**Absolutely Nothing.**

- ❌ No email addresses
- ❌ No recipient data
- ❌ No personal information
- ❌ No tracking cookies
- ❌ No diagnostics
- ❌ No logs

**We cannot access your email data because it never reaches us.**

### Minimal Usage Analytics

**For Vercel-hosted version only**: We use Vercel Analytics to track anonymous page views.

**What IS collected** (Vercel-hosted version only):
- ✅ Anonymous page view counts
- ✅ No cookies (on Vercel's free tier)
- ✅ Privacy-friendly by design

**What is NOT collected**:
- ❌ No user identification (no names, emails, IDs)
- ❌ No email addresses or recipient data
- ❌ No cookies or persistent tracking (on free tier)
- ❌ No personal information

**Why**: This helps us understand approximate usage volume (e.g., "~1,000 people use this") for development prioritization. Nothing more.

**How to avoid**: Clone the repository and self-host locally. Local installations (manifest.xml) are NOT tracked by Vercel Analytics - only the production Vercel-hosted version (manifest.prod.xml) sends anonymous page views.

**Technical details**: Vercel Analytics uses `@vercel/analytics` package to send anonymous page view events. No custom tracking code, no databases, no API endpoints.

**Open source**: You can review the implementation:
- Client-side: `src/taskpane/taskpane.js` - Contains `import { inject } from '@vercel/analytics'`
- Vercel Analytics documentation: https://vercel.com/docs/analytics

**What we get**: Anonymous page view counts visible in Vercel dashboard. We cannot know:
- Who visited (no user identification)
- Email addresses or personal data (never transmitted)
- Specific user behavior (just page loads)

**Comparison to typical analytics**: Most analytics tools (Google Analytics, Mixpanel, etc.) collect:
- User IDs and sessions
- IP addresses and geolocation
- Browser fingerprints
- Page visit timestamps
- User behavior flows
- Device and browser details

**Vercel Analytics on free tier**: Anonymous page views. No cookies. Privacy-friendly.

## How It Works

### Local Processing Only - How Your Data Stays Private

**Every operation runs in your device's memory:**

1. **Data Read**: ClearSend reads email recipients using Office.js API (Microsoft's official API)
2. **Local Processing**: JavaScript executes sorting, validation, deduplication **entirely in your browser/Outlook**
3. **Local Update**: Processed recipients are written back to Outlook using Office.js API
4. **Zero Transmission**: At no point does any data leave your Outlook application

**Technical Details:**
- All JavaScript code runs in your client's sandboxed environment
- No `fetch()`, `XMLHttpRequest`, or network calls to external services
- Email addresses remain in memory only during processing - never serialized for transmission
- Source code is open for verification: every line of code is reviewable

### Local Storage

ClearSend stores configuration data **locally on your device only** using Office.js roaming settings:

**What is stored locally:**
- Processing options (which features are enabled/disabled)
- Internal domains configuration (your organization's domains)
- Invalid email addresses list (if "Save invalid addresses" feature is enabled)
- Processing order preferences (drag-and-drop configuration)

**How it's stored:**
- All data is stored using Office.js roaming settings API
- This data is managed entirely by Microsoft Office and syncs across your Office installations via Microsoft's infrastructure
- Storage is local to your Microsoft account and Office environment
- **We do not have access to this data** - it never reaches our servers
- Only you and Microsoft Office can access your roaming settings

**What is NOT stored:**
- Email addresses from recipient fields (discarded after processing)
- Email content or message bodies
- Contact information
- Any personally identifiable information beyond what you explicitly configure (internal domains)

## Third-Party Services

### Vercel (Hosting and Analytics)

**Static Hosting:**
- We use Vercel to host static files (HTML, CSS, JavaScript)
- Your browser downloads these files once when the add-in loads
- **Your email data NEVER touches Vercel servers** - processing happens locally in your browser
- Vercel may collect standard web access logs (IP address, browser type) when downloading the add-in files
- No email addresses, recipient data, or personal information is ever sent to Vercel

**Vercel Analytics:**
- Vercel Analytics tracks anonymous page views for the hosted version
- No cookies on free tier - privacy-friendly by design
- No personal data, user IDs, or identifying information collected
- Only aggregate page view counts visible in Vercel dashboard
- Self-hosted/local installations are NOT tracked

Read Vercel's privacy policy: https://vercel.com/legal/privacy-policy
Read Vercel Analytics documentation: https://vercel.com/docs/analytics/privacy-policy

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

Questions about privacy? Contact us at clear_send@outlook.com or open an issue on GitHub

## Your Rights

You have complete control over your data because we never access it. You can:
- Uninstall the add-in at any time
- Clear local settings through Office settings
- Review all source code to verify our claims
- Disable analytics (if enabled) by modifying the source code

---

## Frequently Asked Questions

### Q: Do I need to consent to analytics?
**A:** Analytics are enabled on the Vercel-hosted version (manifest.prod.xml). Since Vercel Analytics collects no personal data (just anonymous page views, no cookies on free tier), no formal consent is required under GDPR/CCPA.

### Q: Can you see who I am if I use the add-in?
**A:** No. Vercel Analytics only tracks anonymous page views. We cannot identify you, your organization, your location, or any personal details.

### Q: What if I don't want any analytics at all?
**A:** Clone the repository and self-host locally using manifest.xml. Local installations are NOT tracked by Vercel Analytics - only the production Vercel-hosted version sends anonymous page views.

### Q: Is this GDPR compliant?
**A:** Yes. We collect no personal data, so GDPR's strict requirements don't apply. Anonymous usage counting is explicitly permitted under GDPR.

### Q: Is this CCPA compliant?
**A:** Yes. CCPA regulates "personal information." We collect no personal information, so CCPA doesn't apply.

### Q: What about Vercel's HTTP logs?
**A:** Like any web service, Vercel may log IP addresses in their standard HTTP access logs when you load the add-in files or ping the analytics endpoint. However:
- We don't have access to Vercel's HTTP logs
- We don't store or process IP addresses
- Vercel's logs are temporary and used only for infrastructure purposes
- This is standard for any web service (unavoidable when downloading files)

### Q: How can I verify these privacy claims?
**A:** Review the source code on GitHub:
- `src/taskpane/taskpane.js` - See the Vercel Analytics integration (just `inject()` call)
- Search the codebase for any network requests - you'll see none that transmit email data
- Vercel Analytics source code is open: https://github.com/vercel/analytics

### Q: What if I find a privacy violation?
**A:** Please report it immediately:
- Email: clear_send@outlook.com
- GitHub Issues: https://github.com/fhuerta01/ClearSend/issues
- We take privacy seriously and will address any concerns promptly

### Q: Can analytics be used to track my email recipients?
**A:** No. Absolutely not. Vercel Analytics only tracks page views. Email recipients are NEVER transmitted anywhere - all email processing happens 100% locally in your browser.

---

## Privacy Summary

**Three Core Guarantees:**

1. 🔒 **Your email addresses NEVER leave your device** - All processing is 100% local
2. 🚫 **We have NO servers processing your data** - Only static file hosting
3. ✅ **You can verify everything** - Complete source code is publicly available

**Your privacy is not a feature - it's our architecture.**
