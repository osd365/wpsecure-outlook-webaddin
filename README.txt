WPSecure — Command-only (External config.json)
=============================================

This package keeps a single **WPSecure** ribbon button and the **OnMessageCompose** event, with **no task pane UI**. All tenant/environment-specific values (MSAL client/tenant, template paths, retry caps, toasts, etc.) are stored in **config.json**.

What’s included
---------------
- wpsecure.manifest.xml — command-only manifest (no task pane). Requires **Mailbox 1.10+** for event-based activation & setSignatureAsync. (Older clients degrade to setAsync.)
- wpsecure.html — small runtime host page (loads Office.js, MSAL, and wpsecure.js).
- wpsecure.js — single-file logic; loads **config.json**; strict format policy; per-user/session cache; 10-try caps; button toasts.
- config.json — external configuration (edit per-tenant without JS changes).
- assets/icon-*.png — placeholder icons for the ribbon button.

Configure
---------
1) Upload all files (including **config.json**) to your Storage Account’s **$web** container (same-origin).
2) Edit **config.json**: set `MSAL_CLIENT_ID`, `MSAL_TENANT_ID` (or `common`), and confirm `TEMPLATE_PATHS`.
3) Register an **Entra ID SPA app** and add the NAA redirect: `brk-multihub://<your-$web-origin>` (origin only, no path), then assign **Files.Read** delegated scope. 
4) Update **wpsecure.manifest.xml** URLs: replace `<your-origin>` with your storage account hostname (e.g., `contosoaddins.blob.core.windows.net`) and sideload the manifest.

Notes & limitations
-------------------
- **Manifest references cannot be read from config.json.** Outlook resolves the manifest **before** any web code runs; resource URLs must be **absolute** in the manifest. Use per-tenant manifests (or a single shared origin) to vary URLs. 
- Use the Office.js CDN URL in production (required by Marketplace submissions). 

Behavior
--------
- **Messages**: On `OnMessageCompose`, runtime silently preloads templates from OneDrive and inserts if the matching format exists (**no toasts**; logs only). 
- **Button** (messages & appointments): reinserts using the same rules; shows concise toasts. Appointments insert **NEW** only (no event-based insertion).
- **Strict format**: HTML requires `*.htm`; Text requires `*.txt`. No cross-format fallback.

Logging
-------
- Set default level in **config.json** (`DEFAULT_LOG_LEVEL`) and adjust live via DevTools: `localStorage.setItem('wpsecure_log_level','debug')`.

Security
--------
- Client ID and tenant ID are **not secrets** in SPAs (public clients). Do **not** store secrets client-side.
