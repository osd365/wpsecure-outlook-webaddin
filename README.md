# WPSecure Signature Inserter — Outlook Web Add-in

> **Centrally managed, automatically inserted email signatures for Microsoft 365 — delivered directly from your own tenant. No email routing. No third parties. No connectors.**

![Version](https://img.shields.io/badge/version-1.0.0-1F4E79?style=flat-square)
![Platform](https://img.shields.io/badge/platform-New%20Outlook%20%7C%20Outlook%20on%20the%20Web-2E75B6?style=flat-square)
![Runtime](https://img.shields.io/badge/runtime-.NET%2010%20LTS-512bd4?style=flat-square)
![License](https://img.shields.io/badge/license-Proprietary-a4262c?style=flat-square)

---

## What Is This Repository?

This repository contains the deployment files for the **WPSecure Signature Inserter Outlook Web Add-in**. If you have arrived here from [wpsecure.shop](https://www.wpsecure.shop), you are in the right place.

If you are not yet a WPSecure customer, visit [www.wpsecure.shop](https://www.wpsecure.shop) for licensing and pricing.

> ⚠️ **Deployment requires an experienced Azure administrator** with permissions to create Azure App Services, Storage Accounts, and Entra ID App Registrations. Please read the full [Installation Guide](https://www.wpsecure.shop/documentation-set-signature-using-outlook-web-add-in-v1/) before proceeding.

---

## What's in This Repository

| File | What it is | Where it goes |
|---|---|---|
| `wwwroot/wpsecure.js` | Add-in logic | Azure Web App — `wwwroot/` |
| `wwwroot/wpsecure.html` | Add-in UI and task pane | Azure Web App — `wwwroot/` |
| `wwwroot/wpsecure_manifest.xml` | Office Add-in manifest | Microsoft 365 Admin Center |
| `wwwroot/config.json` | Backend URL configuration | Azure Web App — `wwwroot/` |
| `SignaturesController.cs` | Signature delivery API | Azure Web App — backend |
| `TokenController.cs` | Authentication API | Azure Web App — backend |
| `Program.cs` | Application startup | Azure Web App — backend |
| `ImageService.cs` | Image processing service | Azure Web App — backend |
| `SecurityHelpers.cs` | Token validation | Azure Web App — backend |
| `OboModels.cs` | Authentication models | Azure Web App — backend |
| `Limits.cs` | Size limit constants | Azure Web App — backend |

---

## Prerequisites

Before you begin, ensure the following are in place. Full instructions for each item are in the [Installation Guide](https://www.wpsecure.shop/documentation-set-signature-using-outlook-web-add-in-v1/).

- [ ] **WPSecure device licenses** — purchased at [www.wpsecure.shop](https://www.wpsecure.shop)
- [ ] **Branding package deployed** — signature files must exist in the user's OneDrive under `/z__WPSECURE__SYSTEM_DO_NOT_TOUCH/` before the add-in will deliver signatures
- [ ] **Microsoft 365 tenant** — Global Administrator or Exchange Administrator access
- [ ] **Azure subscription** — permission to create App Services, Storage Accounts, and App Registrations
- [ ] **Azure Web App** provisioned — Runtime stack must be **.NET 10 (LTS)**. Recommended naming: `wpsecure-outlook-webaddin-XXXX`
- [ ] **Azure Blob Storage** — with a dedicated container (recommended name: `signatures`)
- [ ] **Entra ID App Registration** — named `wpsecure-outlook-webaddin`, single tenant, with the following API permissions granted and **admin consent given**:

| Permission | Type |
|---|---|
| `User.Read` | Delegated |
| `Files.Read` | Delegated |
| `GroupMember.Read.All` | Delegated |
| `profile` | Delegated |
| `openid` | Delegated |

> After granting admin consent you should see **five items** in the API permissions list. If you see fewer, a step was missed.

---

## Deployment Overview

The deployment involves four phases that require switching between the Azure Portal and a Notepad file where you collect values as you go. **Follow the [Installation Guide](https://www.wpsecure.shop/documentation-set-signature-using-outlook-web-add-in-v1/) step-by-step** — it includes screenshots for every step.

```
Phase 1 → Create Azure Web App (Part A) — note the Default Domain
Phase 2 → Create Entra ID App Registration — note Client ID, Tenant ID, Secret, App ID URI
Phase 3 → Create Azure Blob Storage — note Connection String and Container Name
Phase 4 → Configure environment variables (Part B) + prepare and deploy files
```

---

## Values to Collect Into Notepad

As you work through each phase, save these values. You will need them all before you are done.

| Value | Where to find it |
|---|---|
| **Default Domain** | Azure Web App → Overview (e.g. `wpsecure-outlook-webaddin-7000.azurewebsites.net`) |
| **Application (Client) ID** | Entra ID → App Registrations → your app → Overview |
| **Directory (Tenant) ID** | Entra ID → App Registrations → your app → Overview |
| **Client Secret value** | Entra ID → App Registrations → Certificates & Secrets |
| **Application ID URI** | Entra ID → App Registrations → Expose an API (format: `api://<default-domain>/<client-id>`) |
| **Blob Connection String** | Storage Account → Security + Networking → Access Keys |
| **Blob Container Name** | The name you gave your container (e.g. `signatures`) |

> ⚠️ **Delete the Notepad file** once setup is complete. It contains sensitive credentials.

---

## Preparing the Files

### Step 1 — Download this release

Go to the [Releases](../../releases/latest) page and download the latest ZIP. Extract it to a folder on your machine.

### Step 2 — Edit `config.json`

Open `wwwroot/config.json` and replace the placeholder URL with your **Default Domain**:

```json
{
  "BACKEND_BASE_URL": "https://your-web-app-domain.azurewebsites.net"
}
```

Use `https://`. No trailing slash.

### Step 3 — Edit `wpsecure_manifest.xml`

Open `wwwroot/wpsecure_manifest.xml` and make the following replacements.

**Replacement 1 — Web App URL (20 occurrences)**

Find:
```
your-web-app-domain.azurewebsites.net
```
Replace all with your **Default Domain** (without `https://`). Use **Replace All**. After replacing, search again to confirm no instances remain.

**Replacement 2 — App Registration details**

Locate the `<WebApplicationInfo>` section near the bottom of the file:

```xml
<WebApplicationInfo>
    <Id>f9de811a-1855-4ad7-b618-1126d6d87b8c</Id>
    <Resource>api://your-web-app-domain.azurewebsites.net/f9de811a-1855-4ad7-b618-1126d6d87b8c</Resource>
    <Scopes>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

- Replace the `<Id>` value with your **Application (Client) ID**
- Replace the `<Resource>` value with your **Application ID URI**

After replacement:

```xml
<WebApplicationInfo>
    <Id>your-application-client-id</Id>
    <Resource>api://your-web-app-domain.azurewebsites.net/your-application-client-id</Resource>
    <Scopes>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

> The Application ID URI must end with your Application (Client) ID and must have no trailing slash.

### Step 4 — ZIP for deployment

Navigate to the parent folder (containing `wwwroot/` and the `.cs` files). Select all files and folders and create a ZIP archive named:

```
wpsecure-outlook-webaddin.zip
```

---

## Deploying to Azure Web App

1. Open your Azure Web App in the Azure Portal
2. Navigate to **Deployments → Deployment Center**
3. Under **Source**, select **Public files (new)**
4. Click **Browse** and select your `wpsecure-outlook-webaddin.zip`
5. Click **Save** — the Web App will restart automatically
6. Open the **Logs** tab to confirm the upload completed successfully

**Verify:** Open a browser and navigate to `https://your-web-app-domain.azurewebsites.net/wpsecure_manifest.xml` — it should load without authentication.

---

## Environment Variables Reference

Configure in your Azure Web App under **Configuration → Application Settings**. Click **Save** after adding all values — the Web App will restart automatically.

### Required

| Variable | Description |
|---|---|
| `TENANTID` | Directory (Tenant) ID — GUID from your Notepad file |
| `CLIENTID` | Application (Client) ID — GUID from your Notepad file |
| `CLIENTSECRET` | Client Secret value — treat as a password, never commit to source control |
| `AUDIENCE_APPIDURI` | Application ID URI — from your Notepad file |
| `BLOB_CONNECTION_STRING` | Azure Blob Storage connection string — from your Notepad file |
| `BLOB_CONTAINER_NAME` | Blob container name (e.g. `signatures`) |
| `SUPPORTED_OUTLOOK_SURFACE` | Permitted Outlook surfaces — recommended: `newOutlookWindows,OutlookWebApp` |

### Optional — Azure Environment (Global Azure defaults are pre-set)

| Variable | Global Azure | US Government | China (21Vianet) |
|---|---|---|---|
| `AUTHORITYHOST` | `https://login.microsoftonline.com` | `https://login.microsoftonline.us` | `https://login.chinacloudapi.cn` |
| `GRAPH_BASEURL` | `https://graph.microsoft.com` | `https://graph.microsoft.us` | `https://microsoftgraph.chinacloudapi.cn` |
| `GRAPH_SCOPES` | `https://graph.microsoft.com/.default` | `https://graph.microsoft.us/.default` | `https://microsoftgraph.chinacloudapi.cn/.default` |

> If your organisation uses Global Azure, these three variables do not need to be set.

### Optional — Signature Behaviour

| Variable | Description | Recommended | Default |
|---|---|---|---|
| `SIGNATURE_REFRESH_PERIOD_IN_MINUTES` | How often clients refresh their cached signature | `240` | `240` |
| `MAX_TOTAL_SIGNATURE_SIZE_KB` | Max combined size of signature HTML + images | `150` | `150` (hard cap: 250) |
| `MAX_SIGNATURE_ITEMS` | Max signature files processed per user | `100` | `100` |
| `SIGNATURE_BLANK_LINES_ON_TOP` | Blank lines inserted above the signature | `4` | `0` (range: 0–10) |
| `CID_PLATFORMS` | Platforms receiving images as inline attachments | `newOutlookWindows,OutlookWebApp` | Not set |
| `USE_OFFICERUNTIME_STORAGE_ON_DESKTOP` | Use OfficeRuntime.Storage as primary cache on desktop | `false` | `false` |

### Administrator Controls

| Variable | Description | Default |
|---|---|---|
| `SIGNATURES_DISABLED` | Set to `true` to immediately disable signature delivery for all users | `false` |
| `DISABLED_ENTRA_GROUP_ID` | Object ID of an Entra ID group whose members will not receive signatures. Requires `GroupMember.Read.All` with admin consent. | Not set |

> ⚠️ After changing any environment variable, click **Save** and restart the App Service.

---

## Installing the Manifest in Microsoft 365

1. Log into the [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Go to **Settings → Integrated apps → Upload custom apps**
3. Upload your updated `wpsecure_manifest.xml`
4. Assign the add-in to the relevant users or groups
5. Allow up to 24 hours for the add-in to appear in Outlook for all assigned users

---

## Supported Outlook Surfaces

| Value | Surface | Status |
|---|---|---|
| `newOutlookWindows` | New Outlook for Windows | ✅ Supported |
| `OutlookWebApp` | Outlook on the Web (OWA) | ✅ Supported |
| `classicWin32` | Classic Outlook for Windows | ⛔ Not supported |
| `classicMac` | Classic Outlook for Mac | ⛔ Not supported |
| `newOutlookMac` | New Outlook for Mac | ⛔ Not supported |
| `OutlookIOS` | Outlook for iOS | ⛔ Not supported |
| `OutlookAndroid` | Outlook for Android | ⛔ Not supported |

---

## Signature Files — OneDrive Structure

Signature files are generated and deployed by the **WPSecure Signature Engine** (Windows Branding Tool / Personalization Packager). Do not proceed until you can confirm these files exist in at least one test user's OneDrive.

```
/z__WPSECURE__SYSTEM_DO_NOT_TOUCH/
  wpsecure_cloud_new.htm           ← Required — HTML new message signature
  wpsecure_cloud_new_files/        ← Folder containing images for new messages
  wpsecure_cloud_reply.htm         ← Required — HTML reply/forward signature
  wpsecure_cloud_reply_files/      ← Folder containing images for replies
  wpsecure_cloud_new.txt           ← Optional — plain text new message signature
  wpsecure_cloud_reply.txt         ← Optional — plain text reply/forward signature
```

HTM files are required. TXT files are optional. You may deploy new message signatures without reply signatures and vice versa.

---

## How It Works

1. When a user opens Outlook, the add-in loads silently in the background.
2. It contacts your Azure Web App, which retrieves the user's signature files from their own OneDrive.
3. Images are processed — hosted on your Azure Blob Storage (HTTPS) or embedded inline as CID attachments, depending on the platform.
4. The signature is cached locally on the user's device for up to 4 hours.
5. When the user opens a compose window, the signature is inserted automatically — no user action required.
6. The cache refreshes silently every 4 hours. Updates appear within 4 hours, or immediately when the user clicks **Refresh Signature** in the task pane.

**Nothing routes through third-party servers. All data stays within your Microsoft 365 tenant and your own Azure infrastructure.**

---

## Upgrading to a New Release

1. Download the new release ZIP from the [Releases](../../releases/latest) page
2. Re-apply your `config.json` and `wpsecure_manifest.xml` edits
3. Re-ZIP and redeploy via Deployment Center
4. Check the release notes — add any new environment variables before restarting

---

## Troubleshooting

For diagnostic guidance, Azure log stream reference, and common issue resolution, refer to the [Administrator Troubleshooting Guide](https://www.wpsecure.shop/troubleshooting-guide) on our website.

---

## Support

Visit [www.wpsecure.shop/contact](https://www.wpsecure.shop/contact) or email [support@wpsecure.shop](mailto:support@wpsecure.shop).

---

## License

**WPSecure Signature Inserter Outlook Web Add-in and its component parts — Proprietary Software**

Copyright © 2019 OSD365 Limited. All rights reserved.

This software is licensed on a per-device subscription basis. Unauthorised use, copying, or distribution is strictly prohibited and constitutes a violation of the WPSecure License Agreement.

Visit [www.wpsecure.shop](https://www.wpsecure.shop) to purchase or renew device licenses, or for full licensing terms.
