# ClearSend - Email Recipient Management for Outlook

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Outlook-orange.svg)
![Privacy](https://img.shields.io/badge/privacy-first-green.svg)

ClearSend is a free, open-source Microsoft Outlook add-in that helps you manage email recipients with powerful automation features. **100% client-side processing** - your email addresses never leave your device. No servers, no cloud, no tracking.

## ✨ Features

- **📋 Sort Recipients** - Alphabetically organize recipients by name or email
- **🔄 Remove Duplicates** - Cross-field deduplication across To, CC, and BCC
- **✅ Validate Emails** - Format validation with common typo detection
- **🏢 Prioritize Internal** - Move internal domain recipients to the top
- **🚫 Remove External** - Filter out external recipients for internal-only emails
- **⚡ Quick Clean** - One-click recipient cleaning with keyboard shortcut (Ctrl+Alt+Q)
- **↩️ Undo Support** - Revert to previous recipient lists
- **💾 Export to CSV** - Download recipient lists for analysis

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

### For Users

#### Installation

1. Download the latest release from the [Releases page](https://github.com/fhuerta01/ClearSend/releases)
2. Open Outlook (Desktop or Web)
3. Go to **Get Add-ins** → **My Add-ins** → **Add from File**
4. Select the downloaded `manifest.xml` file
5. Click **Install**

#### Quick Start

1. Compose a new email in Outlook
2. Click the **ClearSend** button in the ribbon
3. Configure your preferences in the settings
4. Click **Process destination fields** to clean your recipients

### For Developers

#### Prerequisites

- Node.js 14+ and npm
- Microsoft Outlook (Desktop or Web)
- Office Add-ins Developer Certificate (auto-generated)

#### Local Development

```bash
# Clone the repository
git clone https://github.com/fhuerta01/ClearSend.git
cd ClearSend

# Install dependencies
npm install

# Start development server
npm run dev-server

# In another terminal, sideload the add-in
npm start
```

#### Build for Production

```bash
# Build optimized bundle
npm run build

# Validate manifest
npm run validate:prod
```

## 🏗️ Architecture

### Client-Side Processing - Privacy by Design

All email processing logic runs **entirely in your browser/Outlook client** using the `processors.js` library. No data ever leaves your device:

- **Sort Module** - Alphabetical and domain-based sorting (local only)
- **Dedupe Module** - Cross-field duplicate detection (local only)
- **Validation Module** - Email format validation and typo detection (local only)
- **Internal Prioritization** - Internal domain identification (local only)
- **External Filtering** - External recipient removal (local only)

**Technical Implementation:**
- Pure JavaScript functions execute in your browser's memory
- No network requests to external APIs
- No data serialization or transmission
- Email addresses remain in Outlook's context only

### Technology Stack

- **Frontend**: Vanilla JavaScript (ES6+)
- **UI Framework**: Microsoft Fluent UI Core
- **Office Integration**: Office.js API
- **Build Tool**: Webpack 5
- **Deployment**: Vercel (static hosting)

## 📖 User Guide

### Configuration

1. Click the **Settings** icon (⚙️) to configure processing options
2. Enable/disable features:
   - Sort recipients alphabetically
   - Remove duplicates
   - Validate email addresses
   - Prioritize internal domains
   - Remove external recipients
3. Add your organization's internal domains

### Processing Options

#### Sort Recipients
Organizes recipients alphabetically by display name or email address.

#### Remove Duplicates
Eliminates duplicate email addresses across To, CC, and BCC fields. Priority: To > CC > BCC.

#### Validate Emails
Checks email format and detects common typos in domains (e.g., gmial.com → gmail.com).

#### Prioritize Internal
Moves internal domain recipients to the top of the list while maintaining alphabetical order.

#### Remove External
Filters out all external recipients, keeping only internal domain addresses.

### Keyboard Shortcuts

- `Ctrl+Alt+C` - Open ClearSend panel
- `Ctrl+Alt+Q` - Quick clean (one-click processing)

## 🤝 Contributing

We welcome contributions! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

### Development Workflow

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Run tests and linting (`npm run lint`)
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## 📋 Project Structure

```
ClearSend/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html          # Main UI
│   │   ├── taskpane.js            # UI logic
│   │   ├── processors.js          # Processing library
│   │   └── clearsend.css          # Styles
│   └── commands/
│       ├── commands.html          # Command functions
│       └── commands.js            # Quick actions
├── assets/                        # Icons and images
├── manifest.xml                   # Development manifest
├── manifest.prod.xml              # Production manifest
├── webpack.config.js              # Build configuration
└── package.json                   # Dependencies
```

## 🐛 Troubleshooting

### Add-in Not Loading

1. Clear browser cache (Ctrl+Shift+R)
2. Restart Outlook
3. Check the browser console for errors
4. Verify manifest URLs are correct

### Recipients Not Updating

1. Ensure you have write permissions for the email
2. Check that the email is in compose mode
3. Try the Refresh button (🔄)

### Common Issues

- **Issue**: "Failed to update recipients"
  **Solution**: Check Office.js permissions in manifest

- **Issue**: Add-in shows blank screen
  **Solution**: Verify all assets are loaded correctly

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Microsoft Office Add-ins team for the excellent documentation
- Fluent UI team for the design system
- All contributors who have helped improve ClearSend

## 📧 Support

- **Issues**: [GitHub Issues](https://github.com/fhuerta01/ClearSend/issues)
- **Discussions**: [GitHub Discussions](https://github.com/fhuerta01/ClearSend/discussions)
- **Email**: [support@clearsend.com](mailto:support@clearsend.com)

## 🗺️ Roadmap

- [ ] Advanced filtering rules
- [ ] Recipient analytics and insights
- [ ] Integration with contact management
- [ ] Team collaboration features
- [ ] Custom domain rules

---

Made with ❤️ by the ClearSend team
