WPSecure — Command-only (No Task Pane)
======================================

This variant keeps a single **WPSecure** ribbon button and the **OnMessageCompose** event, but removes the task pane UI entirely. The HTML file is a small stub that only hosts the runtime and single JS file.

What changed vs the original
----------------------------
- **Manifest**: Removed <FormSettings>. Only **ExecuteFunction** buttons and the **LaunchEvent** remain. A **FunctionFile** still points to `wpsecure.html` (stub) so Outlook can load the runtime and JS for commands/events.
- **HTML**: now a stub (no DOM controls). It simply loads Office.js, MSAL, and `wpsecure.js`.
- **JS**: removed any task-pane UI wiring; kept command/event handlers and strict insertion policy.

Configure MSAL (NAA)
--------------------
1) Register an SPA app in Entra ID and add the NAA redirect: `brk-multihub://<your-$web-origin>` (origin only).
2) In `wpsecure.js`, set `MSAL_CLIENT_ID` and `MSAL_TENANT_ID` (or `common`).
3) Minimal Graph delegated scope: `Files.Read`.

Upload & Sideload
-----------------
- Upload all files (keep `assets/` folder) to your Storage Account’s `$web` container.
- Update URLs in the manifest to match your origin if needed and sideload the XML.

Behavior
--------
- **OnMessageCompose (messages)**: silent preload and strict insert (HTML needs *.htm; Text needs *.txt). No toasts; console logs only.
- **WPSecure button**: re-injects on demand (messages & appointments). Toasters on success/missing/10 tries.

Logging
-------
- Console prefix: `[WPSecure]`. Enable debug via `localStorage.setItem('wpsecure_log_level','debug')`.
