# PayTM Safe Send - Outlook Add-in

Get warned before sending emails to external (non-PayTM) recipients.

## ✨ Features

- ⚠️ **Real-time Warnings** - Get alerted before sending to external recipients
- 🛡️ **Domain Protection** - Only paytm.com is trusted, all others trigger warning
- 🔒 **User Confirmation** - You must confirm before sending to external addresses
- 📧 **Full Visibility** - See which recipients are external in the warning dialog
- 🖥️ **Cross-Platform** - Works on Mac and Windows Outlook
- ☁️ **Cloud Support** - Works with Gmail, Office 365, and Exchange accounts

## 📋 Requirements

- **Mac or Windows Computer**
- **Outlook for Mac 2016+** or **Outlook for Windows 2016+**
- **Gmail, Office 365, or Exchange account** configured in Outlook
- **Internet Connection**

## 🚀 Installation

### Step 1: Get the Manifest URL

The manifest file is hosted on GitHub:
```
https://raw.githubusercontent.com/vikash3pandey-sys/outlook-alerts/main/manifest.xml
```

### Step 2: Add to Outlook

**For Mac:**
1. Open Outlook
2. Click **Outlook** menu (top left) → **Settings**
3. Click **Manage Add-ins**
4. Click **"Upload My Add-in"** or **"+"**
5. Choose **"Add from URL"**
6. Paste the manifest URL above
7. Click **Add**

**For Windows:**
1. Open Outlook
2. Click **File** → **Options**
3. Click **Trust Center** → **Trust Center Settings**
4. Click **Trusted Add-in Catalogs**
5. Paste the manifest URL
6. Click **Add Catalog**

### Step 3: Restart Outlook

Close Outlook completely and reopen it. The add-in will load automatically.

## 🧪 Testing

1. **Compose a new email** (Cmd+N on Mac, Ctrl+N on Windows)
2. **Add a recipient** from external domain (e.g., test@gmail.com)
3. **Click Send**
4. **Warning dialog appears** showing the external recipient
5. **Choose:**
   - ✓ **Continue** to send anyway
   - ✗ **Cancel** to edit recipients

## 📁 File Structure

```
outlook-alerts/
├── manifest.xml              # Add-in configuration (this is what you upload)
├── onSend.html              # ItemSend event handler (blocks/warns external recipients)
├── taskpane.html            # Settings and info panel
├── README.md                # This file
└── assets/
    ├── icon16.png
    ├── icon32.png
    └── icon80.png
```

## ⚙️ How It Works

### ItemSend Event Handler (onSend.html)

When you click Send:
1. **onSend.html** is triggered (before email is sent)
2. It checks all recipients in To, Cc, Bcc fields
3. Compares domain against trusted list (paytm.com)
4. If external found → Shows warning dialog
5. User confirms or cancels
6. Email either sends or is blocked

### Taskpane (taskpane.html)

Shows settings and information about the add-in.

## 🔧 Configuration

### Trusted Domain

Currently configured to trust: **paytm.com**

To add more trusted domains, edit `onSend.html`:
```javascript
const TRUSTED_DOMAIN = 'paytm.com';
```

### External Recipients

Any email NOT from the trusted domain will trigger a warning.

## 🛡️ Security

- ✓ No email content is scanned
- ✓ No data is sent to external servers
- ✓ All processing happens locally in Outlook
- ✓ User has full control (can confirm or cancel)
- ✓ Open source - code is visible on GitHub

## 🐛 Troubleshooting

### "Installation failed" error

**Solution:**
1. Make sure the manifest URL is accessible in browser
2. Copy the exact URL: `https://raw.githubusercontent.com/vikash3pandey-sys/outlook-alerts/main/manifest.xml`
3. Try removing any old version first
4. Restart Outlook completely (close all windows)
5. Try installing again

### Warning not appearing when sending

**Solution:**
1. Make sure you're composing a NEW email
2. Add external recipient (test@gmail.com)
3. Click Send (not Ctrl+Enter)
4. Wait 1-2 seconds for dialog
5. Restart Outlook if still not working

### Add-in appears but then disappears

**This should NOT happen with this version** - it uses ItemSend event which is stable.

If it does:
1. Remove the add-in from Outlook
2. Restart Outlook
3. Reinstall using the manifest URL
4. Wait 15 seconds after Outlook opens

## 📝 Manifest Details

The manifest (manifest.xml) includes:

- **ItemSend Event** - Handles send events synchronously
- **onSend.html** - Function execution file
- **Permissions** - ReadWriteMailbox (needed for recipient access)
- **Rules** - Applies to Message and Appointment editing
- **Icons** - Embedded SVG icons

This is based on a proven, production-tested structure similar to Safeguard Send.

## 🔄 Updates

To update to a new version:
1. Update files in GitHub
2. Outlook will automatically use latest version
3. No need to reinstall (URLs stay the same)

## 📞 Support

- **GitHub Issues:** https://github.com/vikash3pandey-sys/outlook-alerts/issues
- **Email:** vikash3.pandey-sys@email.com

## 📜 License

Open source. Free to use and modify.

## ✅ Version History

**v1.0.0.0** (Current)
- Initial release
- ItemSend event handler
- PayTM domain protection
- Works on Mac and Windows
- Supports Gmail, Office 365, Exchange

---

**Made with ❤️ for PayTM - Keep your emails safe!**
