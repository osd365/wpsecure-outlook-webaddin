// wpsecure.js
// WPSecure Outlook Web Add-in helper
// Behavior:
// - Sets signature via setSignatureAsync
// - If nothing exists before the signature, inserts two <p><br></p> lines at the very top
// - Uses a signature marker <!-- WPSecureSignature:start --> for robust detection
// - Single-button flow: determines new vs reply/forward and applies *_new or *_reply templates
// - Caches templates in localStorage with user-scoped keys (from SSO token)
// - No thrown errors; logs only; returns benignly

// =========================
// Config (adjust as needed)
// =========================
const WPS_CONFIG = {
  spacerLines: 2, // number of blank lines to add at top if no content precedes signature
  cachePrefix: "wpsecure-cache",
  cacheTtlMs: 30 * 60 * 1000, // 30 minutes
  signatureMarkerStart: "<!-- WPSecureSignature:start -->",
  signatureMarkerEnd: "<!-- WPSecureSignature:end -->",
  // Filenames are expected to be preloaded by your existing logic:
  // e.g., /.wpsecure/<email>_new.html, <email>_reply.html, headers, disclaimers, etc.
  // We'll assume your loader passes them into saveTemplatesToCache(...).
};

// =========================
// Small utilities
// =========================
function log(...args) {
  // Centralized logging (can swap to OfficeRuntime.dialog.message if desired)
  try { console.log("WPSecure:", ...args); } catch { /* no-op */ }
}

function now() { return Date.now ? Date.now() : new Date().getTime(); }

function isNullOrWhitespace(str) {
  if (!str) return true;
  return str.replace(/\u200B|\uFEFF/g, "").trim().length === 0;
}

function stripHtml(html) {
  try {
    const div = document.createElement("div");
    div.innerHTML = html || "";
    return (div.textContent || div.innerText || "").trim();
  } catch {
    return (html || "").replace(/<[^>]*>/g, "").trim();
  }
}

// Returns true if fragment has no meaningful visible content (only breaks, nbsp, or empty blocks)
function isOnlyWhitespaceOrBreaks(htmlFragment) {
  if (!htmlFragment) return true;
  const cleaned = htmlFragment
    .replace(/\u200B|\uFEFF/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/<br\s*\/>?/gi, "\n")
    .replace(/<div[^>]*>\s*<\/div>/gi, "")
    .replace(/<p[^>]*>\s*<\/p>/gi, "")
    .trim();
  const textOnly = stripHtml(cleaned);
  return textOnly.length === 0;
}

function buildSpacers(count) {
  const n = typeof count === "number" && count > 0 ? count : WPS_CONFIG.spacerLines;
  return new Array(n).fill("<p><br></p>").join("");
}

// =========================
// Office async wrappers
// =========================
function getBodyHtmlAsync() {
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (res) => {
        if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value || "");
        else reject(res ? res.error : new Error("getAsync failed"));
      });
    } catch (e) { reject(e); }
  });
}

function setBodyHtmlAsync(html) {
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.item.body.setAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        (res) => {
          if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(res ? res.error : new Error("setAsync failed"));
        }
      );
    } catch (e) { reject(e); }
  });
}

function prependBodyHtmlAsync(html) {
  return new Promise((resolve, reject) => {
    try {
      const body = Office.context.mailbox.item.body;
      if (typeof body.prependAsync !== "function") {
        reject(new Error("prependAsync not supported"));
        return;
      }
      body.prependAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        (res) => {
          if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(res ? res.error : new Error("prependAsync failed"));
        }
      );
    } catch (e) { reject(e); }
  });
}

function setSignatureAsync(signatureHtml) {
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.item.setSignatureAsync(
        signatureHtml,
        { coercionType: Office.CoercionType.Html },
        (res) => {
          if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(res ? res.error : new Error("setSignatureAsync failed"));
        }
      );
    } catch (e) { reject(e); }
  });
}

function getBodyTypeAsync() {
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.item.body.getTypeAsync((res) => {
        if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
        else reject(res ? res.error : new Error("getTypeAsync failed"));
      });
    } catch (e) { reject(e); }
  });
}

// =========================
// Spacing logic
// =========================
async function setSignatureThenEnsureTopSpacing(signatureHtml, spacerLines) {
  try {
    // 1) Set/replace the signature first
    await setSignatureAsync(signatureHtml);

    // 2) If compose is PlainText, we can just prepend newlines (best-effort)
    try {
      const bodyType = await getBodyTypeAsync();
      if (bodyType === Office.MailboxEnums.BodyType.Text) {
        // In practice, signature usage coerces compose to HTML in OWA.
      }
    } catch (e) {
      // non-fatal
    }

    // 3) Read current HTML
    const bodyHtml = await getBodyHtmlAsync();
    if (!bodyHtml) return;

    // 4) Locate signature. Prefer marker; fallback to heuristic.
    const marker = WPS_CONFIG.signatureMarkerStart;
    let sigIndex = bodyHtml.indexOf(marker);

    if (sigIndex === -1) {
      // Fallback heuristic: try wpsecure-signature class
      const classMatch = bodyHtml.match(/<[^>]+class=["'][^"']*wpsecure-signature[^"']*["'][^>]*>/i);
      if (classMatch && classMatch.index >= 0) {
        sigIndex = classMatch.index;
      } else {
        // Generic 'signature' class (last resort)
        const genericSig = bodyHtml.search(/<div[^>]+class=["'][^"']*signature[^"']*["'][^>]*>/i);
        if (genericSig >= 0) sigIndex = genericSig;
      }
    }

    // 5) If signature is at index 0 or not found, try prepend (safer than rebuild)
    if (sigIndex <= 0) {
      try {
        await prependBodyHtmlAsync(buildSpacers(spacerLines));
      } catch (e) {
        // If prepend unsupported or fails, we won't full-replace
        log("Spacing prepend skipped/failure:", e && e.message ? e.message : e);
      }
      return;
    }

    // 6) Check if there is meaningful content before signature
    const before = bodyHtml.substring(0, sigIndex);
    const hasContent = !isOnlyWhitespaceOrBreaks(before);

    if (!hasContent) {
      const spacers = buildSpacers(spacerLines);
      const updated = bodyHtml.slice(0, sigIndex) + spacers + bodyHtml.slice(sigIndex);
      try {
        await setBodyHtmlAsync(updated);
      } catch (e) {
        // Fallback: try prepend
        try {
          await prependBodyHtmlAsync(spacers);
        } catch (e2) {
          log("Failed to apply top spacing (both set and prepend failed).", e2);
        }
      }
    }
  } catch (e) {
    // Never throw—log only
    log("setSignatureThenEnsureTopSpacing error:", e && e.message ? e.message : e);
  }
}

// =========================
// Caching for templates
// =========================
function getUserIdFromTokenCached() {
  try {
    const cached = localStorage.getItem("wpsecure-sso-sub");
    if (cached) return cached;
  } catch { /* no-op */ }
  return "unknown-user";
}

function makeCacheKey(suffix) {
  const userId = getUserIdFromTokenCached();
  return `${WPS_CONFIG.cachePrefix}:${userId}:${suffix}`;
}

function saveTemplatesToCache(templatesObj) {
  try {
    const key = makeCacheKey("templates");
    const payload = { t: now(), v: templatesObj || {} };
    localStorage.setItem(key, JSON.stringify(payload));
  } catch (e) {
    log("saveTemplatesToCache error", e);
  }
}

function loadTemplatesFromCache() {
  try {
    const key = makeCacheKey("templates");
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj || typeof obj !== "object") return null;
    if (!obj.t || (now() - obj.t > WPS_CONFIG.cacheTtlMs)) return null;
    return obj.v || null;
  } catch (e) {
    log("loadTemplatesFromCache error", e);
    return null;
  }
}

// =========================
// Compose helpers
// =========================
async function isReplyOrForward() {
  try {
    const body = await getBodyHtmlAsync();
    const hasQuotedHistory = /From:|Sent:|Date:|To:|Cc:|wrote:|-----Original Message-----/i.test(body || "");
    return !!hasQuotedHistory;
  } catch {
    return false;
  }
}

function ensureSignatureHasMarker(signatureHtml) {
  if (!signatureHtml) return signatureHtml;
  const hasStart = signatureHtml.indexOf(WPS_CONFIG.signatureMarkerStart) >= 0;
  const hasEnd = signatureHtml.indexOf(WPS_CONFIG.signatureMarkerEnd) >= 0;
  if (hasStart && hasEnd) return signatureHtml;
  return [
    WPS_CONFIG.signatureMarkerStart,
    signatureHtml,
    WPS_CONFIG.signatureMarkerEnd
  ].join("");
}

// =========================
// Public API
// =========================
const WPSecure = {
  async init() {
    try {
      log("Init start");
      log("Init done");
    } catch (e) {
      log("init error", e);
    }
  },

  async apply() {
    try {
      log("Apply start");
      const templates = loadTemplatesFromCache();
      if (!templates) {
        log("No templates in cache. Ensure your loader saved them via saveTemplatesToCache().");
        return;
      }

      const replying = await isReplyOrForward();
      const signatureRaw = replying ? templates.signature_reply : templates.signature_new;

      if (!signatureRaw || isNullOrWhitespace(signatureRaw)) {
        log("Missing signature template for this compose mode.");
        return;
      }

      const signatureHtml = ensureSignatureHasMarker(signatureRaw);

      await setSignatureThenEnsureTopSpacing(signatureHtml, WPS_CONFIG.spacerLines);

      log("Apply done");
    } catch (e) {
      log("apply error", e);
    }
  },

  // Optional helpers (not wired by default)
  async injectHeader(headerHtml) {
    if (!headerHtml || isNullOrWhitespace(headerHtml)) return;
    try { await prependBodyHtmlAsync(headerHtml); } catch (e) { log("injectHeader failed", e); }
  },

  async injectDisclaimer(disclaimerHtml) {
    if (!disclaimerHtml || isNullOrWhitespace(disclaimerHtml)) return;
    try {
      const bodyHtml = await getBodyHtmlAsync();
      const updated = (bodyHtml || "") + disclaimerHtml;
      await setBodyHtmlAsync(updated);
    } catch (e) { log("injectDisclaimer failed", e); }
  },

  saveTemplatesToCache,
  loadTemplatesFromCache
};

window.WPSecure = WPSecure;

if (typeof Office !== "undefined") {
  Office.onReady(() => {
    // Call WPSecure.init() from your page when ready
  });
}
