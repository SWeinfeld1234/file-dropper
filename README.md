# Azure File Attacher — Outlook Add-in

An Outlook Add-in that lets you attach files to emails by:
1. **Dragging files** from Salesforce (or your desktop) into the add-in's task pane
2. **Browsing Azure storage** directly from within Outlook and clicking "Attach"

## Architecture

```
┌─────────────────────┐     drag & drop      ┌──────────────────────┐
│  Salesforce LWC      │ ──── (DownloadURL) ──→│  Outlook Add-in      │
│  (Azure File Browser)│      to desktop,     │  Task Pane           │
│                      │      then to pane    │                      │
└─────────────────────┘                       │  ┌────────────────┐  │
                                              │  │ Drop Zone      │  │
                                              │  │ reads file     │  │
                                              │  │ as base64      │  │
                                              │  └───────┬────────┘  │
                                              │          │           │
                                              │          ▼           │
                                              │  addFileAttachment   │
                                              │  FromBase64Async()   │
                                              │          │           │
                                              │          ▼           │
                                              │  ┌────────────────┐  │
                                              │  │ Email gets     │  │
                                              │  │ attachment ✓   │  │
                                              │  └────────────────┘  │
                                              │                      │
                                              │  ┌────────────────┐  │
                                              │  │ Browse Azure   │──┼──→ Salesforce Apex REST
                                              │  │ (optional)     │←─┼──  (AzureFileStorageRest)
                                              │  └────────────────┘  │
                                              └──────────────────────┘
```

## Project Structure

```
outlook-addin/
├── manifest.xml                     # Office Add-in manifest (sideload this)
├── package.json                     # Dev dependencies & scripts
├── src/
│   ├── taskpane.html                # Main task pane UI
│   ├── taskpane.css                 # Styles (Fluent UI inspired)
│   ├── taskpane.js                  # Core logic (drop, browse, attach)
│   ├── function-file.html           # Required by manifest (empty)
│   └── assets/
│       ├── icon-16.png              # Ribbon icon 16x16
│       ├── icon-32.png              # Ribbon icon 32x32
│       └── icon-80.png              # Ribbon icon 80x80
└── salesforce/
    └── classes/
        ├── AzureFileStorageRest.cls          # Apex REST endpoint
        └── AzureFileStorageRest.cls-meta.xml
```

## Prerequisites

- Node.js 18+ (for the dev server)
- An Office 365 / Microsoft 365 account with Outlook
- Your existing Salesforce org with `AzureFileStorageController` deployed
- SSL certificates for localhost (Office Add-ins require HTTPS)

## Setup & Development

### 1. Install dependencies

```bash
cd outlook-addin
npm install
```

### 2. Generate SSL certificates

Office Add-ins must be served over HTTPS, even in development.

```bash
# Option A: Use the Office Add-in dev certs tool
npx office-addin-dev-certs install

# Option B: Use mkcert
npm run generate-certs
```

### 3. Start the dev server

```bash
npm run dev
```

This serves `src/` on `https://localhost:3000`.

### 4. Sideload into Outlook

#### Outlook on the Web (easiest for testing)

1. Open https://outlook.office.com
2. Create a new email (Compose mode)
3. Click the **"…" (More actions)** button in the toolbar
4. Click **"Get Add-ins"**
5. Click **"My add-ins"** → **"Add a custom add-in"** → **"Add from file…"**
6. Upload `manifest.xml`
7. The "Attach Files" button appears in the ribbon

#### New Outlook on Windows

Same process as Outlook on the Web — the new Outlook uses the same add-in framework.

#### Classic Outlook on Windows

```bash
npm run sideload
```

Or manually:
1. In Outlook, go to **File → Manage Add-ins**
2. Click **"Add from File…"**
3. Select `manifest.xml`

### 5. Deploy the Apex REST endpoint (for Browse tab)

If you want the "Browse Azure" tab to work (fetching files directly from Azure storage via Salesforce), deploy the Apex class:

```bash
cd salesforce
sfdx force:source:deploy -p classes/AzureFileStorageRest.cls -u YourOrg
```

Then configure the connection in the add-in's Settings panel with your Salesforce instance URL and a session token.

## How It Works

### Drop Zone (Tab 1)

1. User drags a file from **Salesforce** (your LWC uses Chrome's `DownloadURL` to create a temp file) or from the desktop / file explorer.
2. The file lands in the add-in's drop zone (HTML5 `drop` event).
3. The add-in reads the file as base64 using `FileReader.readAsDataURL()`.
4. It calls `Office.context.mailbox.item.addFileAttachmentFromBase64Async()` to attach.
5. The file appears as an attachment on the email being composed.

### Browse Azure (Tab 2)

1. User configures their Salesforce connection (instance URL + session token).
2. The add-in calls `AzureFileStorageRest` (Apex) to list directories.
3. User navigates folders and clicks "Attach" on any file.
4. The add-in fetches the file content (base64) via Apex and attaches it.

## Salesforce Connection Options

The Browse tab needs to call your Salesforce Apex endpoint. There are several ways to handle auth:

### Option A: Session Token (simplest, for internal use)

Paste a Salesforce session ID into the Settings panel. You can get this from:
- Setup → Session Management
- Developer Console → `UserInfo.getSessionId()`
- Your LWC could pass it via URL parameter when opening the add-in

### Option B: Connected App + OAuth (recommended for production)

1. Create a Connected App in Salesforce
2. Configure OAuth callback URL to your add-in's domain
3. Implement the OAuth flow in the add-in
4. Store the refresh token securely

### Option C: Middleware Proxy

Deploy a small proxy (Azure Function, Heroku, etc.) that:
- Handles CORS (Salesforce REST API doesn't support it from arbitrary origins)
- Stores Salesforce credentials securely
- Forwards requests to the Apex REST endpoint

## CORS Considerations

Salesforce REST API blocks cross-origin requests from non-Salesforce domains by default. The add-in runs on your domain (e.g., `localhost:3000` or `your-addin.azurewebsites.net`), so direct calls to Salesforce will be blocked by CORS.

**Solutions:**
1. **Salesforce CORS Allowlist**: Setup → CORS → add your add-in's origin
2. **Proxy**: Route through a middleware that handles CORS headers
3. **Salesforce Site**: Host the Apex REST endpoint on a Salesforce Site with CORS enabled

## Production Deployment

1. **Host the add-in files** on a web server with HTTPS (Azure App Service, Vercel, Netlify, etc.)
2. **Update URLs** in `manifest.xml` to point to your production domain
3. **Submit to your org's admin** for centralized deployment via Microsoft 365 Admin Center
4. Replace placeholder icons with your branded icons

## Compatibility

| Client                    | Drop Zone | Browse Azure |
|---------------------------|-----------|--------------|
| Outlook on the Web        | ✓         | ✓            |
| New Outlook (Windows)     | ✓         | ✓            |
| Classic Outlook (Windows) | ✓         | ✓            |
| Outlook for Mac           | ✓         | ✓            |
| Outlook Mobile            | ✗         | ✗            |

## Troubleshooting

**"Office.js not ready" error**
- Make sure the Office.js script loads before your code
- Ensure you're in Compose mode (not reading an email)

**Drop zone not accepting files**
- Verify the add-in task pane is open and visible
- Try a smaller file first (attachment size limits apply)

**CORS errors on Browse tab**
- Add your add-in's origin to Salesforce CORS allowlist
- Or use a proxy middleware

**addFileAttachmentFromBase64Async fails**
- Check file size (Exchange has attachment limits, typically 25-150 MB depending on config)
- Ensure the base64 string doesn't include the `data:` prefix
- Verify you're in Compose mode with `Mailbox 1.8` requirement set

## License

MIT — Logicfold
