# ClearSend - Email Recipient Management for Outlook

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Outlook-orange.svg)
![Privacy](https://img.shields.io/badge/privacy-first-green.svg)

ClearSend is a free, open-source Microsoft Outlook add-in that helps you manage email recipients with powerful automation features. **100% client-side processing** - your email addresses never leave your device. No servers, no cloud, no tracking.

## ✨ Features

- **📋 Sort Recipients** - Alphabetically organize recipients by name or email
- **🔄 Remove Duplicates** - Cross-field deduplication across To, CC, and BCC
- **✅ Prevent Invalids Processing** - Stop processing if invalid email addresses are detected
- **💾 Keep Invalid Addresses** - Save invalid addresses across sessions for tracking
- **🏢 Prioritize Internal** - Move internal domain recipients to the top of the list
- **🚫 Remove External** - Filter out external recipients for internal-only emails
- **⚡ Quick Clean** - One-click recipient cleaning with keyboard shortcut (Ctrl+Alt+Q)
- **↩️ Undo Support** - Revert to previous recipient lists
- **📊 Recipient Analysis** - Real-time statistics for destinations, duplicates, and invalid addresses
- **💾 Export to CSV** - Download recipient lists and invalid addresses for analysis
- **⚙️ Customizable Order** - Drag-and-drop to reorder processing steps
- **🔧 Restore Defaults** - One-click reset to default settings

## 🔒 Privacy First - Your Data Stays With You

**Absolute Privacy Guarantee:**
- ✅ **100% Client-Side Processing** - All operations run locally in Outlook (desktop or web browser)
- ✅ **Zero Data Transmission** - Your email addresses NEVER leave your device
- ✅ **No Servers** - No backend servers process or store your data
- ✅ **No Cloud Storage** - Nothing is uploaded or synchronized to any cloud service
- ✅ **No Tracking** - No analytics, cookies, or telemetry
- ✅ **Open Source** - Verify our claims by reviewing the complete source code

**Your email addresses remain exclusively in your Outlook application. Period.**

## 🚀 Getting Started

### Installation Options

ClearSend offers two installation methods:

#### Option 1: Quick Install (Recommended) - Using Vercel Deployment

This method uses our hosted version on Vercel. Perfect for most users.

1. Download only the **manifest.prod.xml** file from the [Releases page](https://github.com/fhuerta01/ClearSend/releases)
2. Open Outlook (Desktop or Web)
3. Go to **Get Add-ins** → **My Add-ins** → **Add from File**
4. Select the downloaded `manifest.prod.xml` file
5. Click **Install**

**Benefits:**
- Smallest download (just the manifest file)
- Always up-to-date with latest version
- No local server required
- Faster installation

#### Option 2: Local Installation - Self-Hosted

This method runs ClearSend entirely from your local machine. Ideal for offline use or corporate environments.

**Prerequisites:**
- Node.js 14+ and npm
- Microsoft Outlook (Desktop or Web)

**Steps:**

1. Download the complete source code from the [Releases page](https://github.com/fhuerta01/ClearSend/releases) or clone the repository:
   ```bash
   git clone https://github.com/fhuerta01/ClearSend.git
   cd ClearSend
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Build the project:
   ```bash
   npm run build
   ```

4. Start the local server:
   ```bash
   npm start
   ```

5. In Outlook:
   - Go to **Get Add-ins** → **My Add-ins** → **Add from File**
   - Select the `manifest.xml` file (not manifest.prod.xml)
   - Click **Install**


### Quick Start

1. Compose a new email in Outlook
2. Click the **ClearSend** button in the ribbon (or use Ctrl+Alt+C)
3. Configure your preferences in the Configuration tab
4. Click **Process destination fields** to clean your recipients

## 🏗️ Architecture

### Client-Side Processing - Privacy by Design

All email processing logic runs **entirely in your browser/Outlook client** using the `processors.js` library. No data ever leaves your device:

- **Sort Module** - Alphabetical and domain-based sorting (local only)
- **Dedupe Module** - Cross-field duplicate detection (local only)
- **Validation Module** - Email format validation (local only)
- **Internal Prioritization** - Internal domain identification (local only)
- **External Filtering** - External recipient removal (local only)
- **Invalid Tracking** - Saved invalid addresses storage (local roaming settings only)

**Technical Implementation:**
- Pure JavaScript functions execute in your browser's memory
- No network requests to external APIs
- No data serialization or transmission
- Email addresses remain in Outlook's context only
- Settings stored in Office.js roaming settings (synced by Microsoft across your devices)


## 📋 Project Structure

```
ClearSend/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html          # Main UI
│   │   ├── taskpane.js            # UI logic and Office.js integration
│   │   ├── processors.js          # Client-side processing library
│   │   └── clearsend.css          # Fluent UI styles
│   └── commands/
│       ├── commands.html          # Command function UI
│       └── commands.js            # Quick Clean ribbon action
├── assets/                        # Icons and images
├── api/                           # Vercel serverless functions (static hosting only)
├── manifest.xml                   # Development manifest (localhost)
├── manifest.prod.xml              # Production manifest (Vercel)
├── webpack.config.js              # Build configuration
├── vercel.json                    # Vercel deployment config
└── package.json                   # Dependencies
```


## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

### Disclaimer

**THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.**

By using ClearSend, you acknowledge that:
- You use this software at your own risk
- The authors and contributors are not responsible for any data loss, email delivery issues, or other problems that may arise from using this software
- This is free, open-source software provided with no guarantees or warranties
- You are responsible for testing and verifying the software meets your needs before relying on it for critical operations


## 📧 Support

- **Issues**: [GitHub Issues](https://github.com/fhuerta01/ClearSend/issues)
- **Discussions**: [GitHub Discussions](https://github.com/fhuerta01/ClearSend/discussions)
- **Email**: clear_send@outlook.com
- **Privacy Policy**: [PRIVACY.md](PRIVACY.md)

---

Made with ❤️ for privacy-conscious distribution lists owners