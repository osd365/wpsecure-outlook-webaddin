WPSecure — One-shot Package (README)
====================================

This package contains a minimal Outlook Web add-in that inserts approved signatures from the user's OneDrive `/.wpsecure` folder. It is fully client-side (no middle tier) and uses MSAL.js with Nested App Authentication (NAA) to call Microsoft Graph directly.

Files
-----
- wpsecure.manifest.xml
- wpsecure.html
- wpsecure.js
- assets/icon-16.png, icon-32.png, icon-80.png

What you must configure (once per tenant/customer)
--------------------------------------------------
1) **Azure App Registration (SPA)**
   - Platform: *Single-page application*
   - **Redirect URI (NAA):** `brk-multihub://<your-$web-origin>` (origin only, e.g., `brk-multihub://wpsecure.blob.core.windows.net`)
   - **Scopes (delegated):** `Files.Read` (minimal for reading templates from OneDrive)
   - (Optional) `User.Read` if you later want to show the user's display name.

2) **Update `wpsecure.js`**
   - Set `MSAL_CLIENT_ID` to the app registration's *Application (client) ID*.
   - Set `MSAL_TENANT_ID` to `common` (multi-tenant) or your tenant ID (single-tenant).

3) **Upload to Azure Blob Static Website**
   - Enable Static website on the Storage Account.
   - Upload all files to the `$web` container, keeping the `assets/` folder structure.

4) **Sideload the Manifest**
   - Replace the icon URLs and HTML/JS URLs in the manifest with your `$web` origin if it differs from the placeholder.
   - Distribute or sideload `wpsecure.manifest.xml` via Integrated Apps or Exchange admin methods.

Usage Model
-----------
- **OnMessageCompose (messages only):** silently preloads from OneDrive and inserts if the matching-format template exists. No toasts; logs go to the browser console (F12).
- **WPSecure button:** tries again on demand (messages and appointments). Shows concise toasts on success/failure. After 10 unsuccessful attempts this session, it stops trying and shows the special "Tried 10 times" toast.

Strict Format Policy
--------------------
- **HTML compose:** requires `*.htm` template; otherwise aborts.
- **Text compose:** requires `*.txt` template; otherwise aborts.

OneDrive Paths (by file)
------------------------
- `/.wpsecure/wpsecure_cloud_new.htm`
- `/.wpsecure/wpsecure_cloud_new.txt`
- `/.wpsecure/wpsecure_cloud_reply.htm`
- `/.wpsecure/wpsecure_cloud_reply.txt`

Logging & Troubleshooting
-------------------------
- Console prefix: `[WPSecure]` with levels info/warn/error (enable debug: `localStorage.setItem('wpsecure_log_level','debug')`).
- Check retry counters in localStorage: `wpsecure_retry_event_*`, `wpsecure_retry_button_*`.
- Ensure your app registration includes the exact NAA redirect `brk-multihub://<origin>`.

Security & Privacy
------------------
- Tokens handled by MSAL in browser; we do not persist tokens beyond MSAL.
- localStorage only holds template content and small counters; no message content.

References
----------
- Outlook event-based activation & signatures (`setSignatureAsync`) require *Mailbox 1.10+*; see Microsoft docs.
- Office.js CDN for Outlook add-ins.
- getComposeTypeAsync to detect new/reply/forward for messages (Mailbox 1.10+).

