/* WPSecure Outlook Add-in Runtime — SSO+OBO Silent Cache Edition
 * Build: 2026-03-01T00:00Z
 * Changes in this build:
 * - TXT (button, New & Reply): insert at CURSOR using setSelectedDataAsync (non-destructive). If body is empty, add 3 leading newlines; otherwise none.
 * - TXT (event): unchanged (default-aware + reply cleanup; New event remains top-only).
 * - HTML (message, New & Reply): when body is effectively empty, inject a 3-line spacer *inside* the signature HTML (after <body> if present; otherwise prefix fragment); avoid placing <div> before <head>.
 * - HTML (appointment): spacer kept (safe merge).
 * - Retry policy: unchanged.
 */
(function(){
  // ---------------- Logger ----------------
  const LEVELS = ['off','error','warn','info','debug'];
  function mkLogger(defaultLevel){
    const lvl = (localStorage.getItem('wpsecure_log_level') || defaultLevel || 'info').toLowerCase();
    const idx = LEVELS.indexOf(lvl);
    return function log(level,msg,meta){
      const lid = LEVELS.indexOf(level);
      if(lvl==='off' || lid<0 || idx<0 || lid>idx) return;
      const p = `[WPSecure][${level}]`;
      if(level==='error') console.error(p,msg,meta||'');
      else if(level==='warn') console.warn(p,msg,meta||'');
      else if(level==='debug') console.debug(p,msg,meta||'');
      else console.log(p,msg,meta||'');
    }
  }

  // ---------------- Config ----------------
  let CONFIG = { BACKEND_BASE_URL: '' }
  async function loadMinimalConfig(){
	try{
		const r = await fetch('./config.json', { cache: 'no-store' });
		if (!r.ok) return;
		const json = await r.json();
		if (json && typeof json === 'object') {
		  const base = (json.BACKEND_BASE_URL || '').trim();
		  if (base) {
			// Normalize to a consistent trailing-slash form
			CONFIG.BACKEND_BASE_URL = base.endsWith('/') ? base : base + '/';
		  }
		}
	  }catch(e){ /* keep defaults */ }
	}

	function computeBootstrapUrl(){
	  try{
		// If BACKEND_BASE_URL is set, build the absolute URL.
		if (CONFIG.BACKEND_BASE_URL) {
		  const baseNoSlash = CONFIG.BACKEND_BASE_URL.endsWith('/')
			? CONFIG.BACKEND_BASE_URL.slice(0, -1)
			: CONFIG.BACKEND_BASE_URL;
		  return `${baseNoSlash}/api/signatures/bootstrap`;
		}
		// Fallback (relative) if config is missing — preserves current safe behavior
		return '/api/signatures/bootstrap';
	  }catch(_){
		return '/api/signatures/bootstrap';
	  }
	}

  // ---------------- Helpers ----------------
  let log = ()=>{};
  function sessionId(){ try{ return (Office && Office.context && Office.context.sessionId) || 'global'; }catch{ return 'global'; } }
  function userFrag(){ try{ const up=Office.context.mailbox.userProfile; return (up && up.emailAddress || 'user').replace(/[^a-z0-9]/gi,'').slice(0,10) || 'user'; }catch{ return 'user'; } }
  function makeKey(kind, fmt, scenario){
    const u = userFrag(); const s = sessionId();
    if(kind==='appointment') return `wpsecure_sig_${u}_${s}_appointment_${fmt}_new`;
    return `wpsecure_sig_${u}_${s}_message_${fmt}_${scenario}`;
  }
  function cacheSet(k,v){ try{ localStorage.setItem(k,v); }catch(_){ } }
  function cacheGet(k){ try{ return localStorage.getItem(k); }catch(_){ return null; } }
  function sleep(ms){ return new Promise(r=>setTimeout(r,ms)); }

  function getBodyFormat(){
    return new Promise(resolve=>{
      try{
        Office.context.mailbox.item.body.getTypeAsync(res=>{
          if(res.status===Office.AsyncResultStatus.Succeeded){
            resolve(res.value===Office.MailboxEnums.BodyType.Html ? 'HTML' : 'txt');
          } else { resolve('HTML'); }
        });
      }catch(_){ resolve('HTML'); }
    });
  }
  function getComposeScenario(){
    return new Promise(resolve=>{
      try{
        const fn = Office.context.mailbox.item.getComposeTypeAsync;
        if(!fn) return resolve('new');
        fn(ar=>{
          if(ar.status!==Office.AsyncResultStatus.Succeeded) return resolve('new');
          const ct = ar.value && ar.value.composeType;
          resolve((ct==='reply' || ct==='forward') ? 'reply' : 'new');
        });
      }catch(_){ resolve('new'); }
    });
  }

  // ---------------- Signature Helpers ----------------
  function disableClientSignature() {
    return new Promise(resolve => {
      try { Office.context.mailbox.item.body.disableClientSignatureAsync(() => resolve()); } catch(_) { resolve(); }
    });
  }
  function getBodyText() {
    return new Promise(resolve => {
      try { Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, ar => { resolve(ar.status === Office.AsyncResultStatus.Succeeded ? (ar.value||'') : ''); }); } catch(_) { resolve(''); }
    });
  }
  function setBodyText(text) {
    return new Promise(resolve => {
      try { Office.context.mailbox.item.body.setAsync(text, { coercionType: Office.CoercionType.Text }, r => { resolve(r.status === Office.AsyncResultStatus.Succeeded); }); } catch(_) { resolve(false); }
    });
  }
  function getBodyHtml() {
    return new Promise(resolve => {
      try { Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, ar => { resolve(ar.status === Office.AsyncResultStatus.Succeeded ? (ar.value||'') : ''); }); } catch(_) { resolve(''); }
    });
  }
  function setBodyHtml(html) {
    return new Promise(resolve => {
      try { Office.context.mailbox.item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, r => { resolve(r.status === Office.AsyncResultStatus.Succeeded); }); } catch(_) { resolve(false); }
    });
  }

  // sessionData for previous TXT signature (used only in event flows now)
  function getSessionTextSig(){
    return new Promise(resolve => {
      try { Office.context.mailbox.item.sessionData.getAsync('wpsecure_txt_sig', ar => { resolve(ar.status === Office.AsyncResultStatus.Succeeded ? (ar.value||'') : ''); }); } catch(_) { resolve(''); }
    });
  }
  function setSessionTextSig(val){
    return new Promise(resolve => {
      try { Office.context.mailbox.item.sessionData.setAsync('wpsecure_txt_sig', val||'', _ => resolve()); } catch(_) { resolve(); }
    });
  }

  // Normalize odd encodings in TEXT
  function normalizeWeirdText(t){ if(!t) return ''; t = t.replace(/[\u0A0D\u0D0A]/g,'\n'); t = t.replace(/[\u200B-\u200D\uFEFF]/g,''); return t; }
  function diffRemoved(before, after){ const b=normalizeWeirdText(before), a=normalizeWeirdText(after); const bl=b.length, al=a.length; let p=0; while(p<bl&&p<al&&b.charCodeAt(p)===a.charCodeAt(p)) p++; let s=0; while(s<bl-p&&s<al-p&&b.charCodeAt(bl-1-s)===a.charCodeAt(al-1-s)) s++; const removed=b.slice(p, bl-s); return { removed, prefixLen:p, normBefore:b, normAfter:a }; }

  // --- HTML spacer application for MESSAGE signatures ---
  function applyHtmlSpacerIfEmptyBody(sigHtml, currentBodyHtml){
    const textOnly = (currentBodyHtml||'').replace(/<style[\s\S]*?<\/style>/gi,'').replace(/<script[\s\S]*?<\/script>/gi,'').replace(/<[^>]+>/g,'').trim();
    if(textOnly.length>0) return sigHtml; // body has content → no spacer
    const spacer = '<div><br></div><div><br></div><div><br></div>';
    const lower = sigHtml.trim().toLowerCase();
    if(lower.startsWith('<!doctype') || lower.startsWith('<html') || lower.startsWith('<head') || lower.includes('<body')){
      // try to inject right after <body>
      return sigHtml.replace(/<body[^>]*>/i, m => m + spacer);
    }
    // treat as fragment
    return spacer + sigHtml;
  }

  // ---------------- Retry bootstrap until cached ----------------
  async function retryBootstrapUntilCached(kind, fmt, scenario, attempts, delayMs){
    const key = makeKey(kind, fmt, scenario);
    let payload = cacheGet(key);
    for(let i=0;i<attempts && !payload;i++){
      await bootstrapSignaturesSilently();
      payload = cacheGet(key);
      if(!payload && delayMs>0) await sleep(delayMs);
    }
    return payload;
  }

  // ---------------- Insertion (message HTML: replace; appt HTML safe merge; TEXT: event vs button) ----------------
  async function insertSignature(content, fmt, scenario, contextKind){
    const isMessage = (function(){ try{ return Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message; }catch(_){ return true; } })();
    const isAppointment = !isMessage;

    if(fmt==='HTML'){
      if(isAppointment){
        try{
          const current = await getBodyHtml();
          const wrapperStart = '<div data-wpsecure="sig-appt">';
          const wrapperEnd = '</div>';
          const spacer = '<div><br></div><div><br></div><div><br></div>';
          const wrapped = `${wrapperStart}${content}${wrapperEnd}`;
          let nextHtml;
          if(current && current.indexOf(wrapperStart)>=0){
            const re = new RegExp(`${wrapperStart}[\\s\\S]*?${wrapperEnd}`,'i');
            nextHtml = current.replace(re, wrapped);
            nextHtml = nextHtml.replace(wrapped, spacer + wrapped);
          } else {
            nextHtml = (current||'') + spacer + wrapped;
          }
          try{ await disableClientSignature(); }catch(_){ }
          return await setBodyHtml(nextHtml);
        }catch(_){ return false; }
      }
      // MESSAGE HTML: possibly apply spacer inside signature when body empty
      try{
        const currentHtml = await getBodyHtml();
        const sigMaybeSpaced = applyHtmlSpacerIfEmptyBody(content, currentHtml);
        // We only disable default when actually inserting
        return await new Promise(resolve=>{
          const opts = { coercionType: Office.CoercionType.Html };
          try{
            if(Office.context.mailbox.item.body.setSignatureAsync){
              disableClientSignature().then(()=>{
                Office.context.mailbox.item.body.setSignatureAsync(sigMaybeSpaced, opts, r=> resolve(r.status===Office.AsyncResultStatus.Succeeded));
              }).catch(()=>{
                Office.context.mailbox.item.body.setSignatureAsync(sigMaybeSpaced, opts, r=> resolve(r.status===Office.AsyncResultStatus.Succeeded));
              });
            } else { resolve(false); }
          }catch(_){ resolve(false); }
        });
      }catch(_){ return false; }
    }

    // TEXT path
    if(contextKind==='button'){
      // Non-destructive: insert at cursor using setSelectedDataAsync.
      // If the WHOLE body is empty, prefix with 3 newlines; otherwise no space.
      const bodyNow = normalizeWeirdText(await getBodyText());
      const prefix = bodyNow.trim().length===0 ? '\n\n\n' : '';
      return await new Promise(resolve=>{
        try{
          Office.context.mailbox.item.body.setSelectedDataAsync(prefix + content, { coercionType: Office.CoercionType.Text }, res => {
            resolve(res.status===Office.AsyncResultStatus.Succeeded);
          });
        }catch(_){ resolve(false); }
      });
    }

    // Event flows (existing behavior): default-aware for NEW; reply cleanup
    try{
      const beforeRaw = await getBodyText();
      const before = normalizeWeirdText(beforeRaw);
      try{ await disableClientSignature(); }catch(_){ }
      const afterRaw = await getBodyText();
      const after = normalizeWeirdText(afterRaw);
      const { removed, prefixLen, normBefore, normAfter } = diffRemoved(before, after);
      const prefixHasText = /\S/.test(normBefore.slice(0, prefixLen));

      let baseline = normAfter;
      const prev = await getSessionTextSig();
      if(prev && baseline && baseline.indexOf(prev)>=0){ baseline = baseline.replace(prev,''); }

      let nextBody;
      if(scenario==='reply'){
        // Prepend signature then collapse down to header
        const pre = content + (baseline ? "\n\n" + baseline : '');
        // minimal cleanup: remove everything up to first From: after our sig
        const t = pre; const sigIdx = t.indexOf(content); const rest = t.slice(sigIdx+content.length);
        const m = /(\n|^)From\s*:\s*.*$/mi.exec(rest);
        if(m){ const hdrPos = sigIdx+content.length + m.index + (m[1]?m[1].length:0); nextBody = t.slice(0, sigIdx+content.length).replace(/\s*$/, '') + "\n\n" + t.slice(hdrPos); }
        else { nextBody = pre; }
      } else {
        // NEW event: top-only
        nextBody = "\n\n\n" + content;
      }
      const ok = await setBodyText(nextBody);
      if(ok) await setSessionTextSig(content);
      return ok;
    }catch(_){ return false; }
  }

  // ---------------- SSO + Bootstrap (silent) ----------------
  async function getSsoIdTokenSilent(){ try{ return await Office.auth.getAccessToken({ allowSignInPrompt:false, allowConsentPrompt:false }); }catch(_){ return null; } }
  async function bootstrapSignaturesSilently(){
    try{
      log('info','bootstrap:start');
      const url = computeBootstrapUrl(); if(!url){ log('warn','bootstrap:missing-endpoint'); return; }
      const idTok = await getSsoIdTokenSilent(); if(!idTok){ log('warn','bootstrap:no-idtoken'); return; }
      const res = await fetch(url, { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ id_token:idTok }) });
      if(!res.ok){ log('error','bootstrap:http',{status:res.status}); return; }
      const payload = await res.json();
      const files = (payload&&payload.files)||{};
      if(files.newHtml) cacheSet(makeKey('message','HTML','new'), files.newHtml);
      if(files.replyHtml) cacheSet(makeKey('message','HTML','reply'), files.replyHtml);
      if(files.newText) cacheSet(makeKey('message','txt','new'), files.newText);
      if(files.replyText) cacheSet(makeKey('message','txt','reply'), files.replyText);
      if(files.apptNewHtml) cacheSet(makeKey('appointment','HTML','new'), files.apptNewHtml);
      if(files.apptNewText) cacheSet(makeKey('appointment','txt','new'), files.apptNewText);
    }catch(_){ log('error','bootstrap:error'); }
  }

  // ---------------- Orchestrator ----------------
  async function doInsertCurrentContext(maxAttempts, contextKind){
    const itemType = (function(){ try{ return Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message ? 'message':'appointment'; }catch(_){ return 'message'; } })();
    const fmt = await getBodyFormat();
    const scenario = itemType==='message' ? await getComposeScenario() : 'new';

    const key = makeKey(itemType, fmt, scenario);
    let payload = cacheGet(key);

    if(!payload){
      const attempts = Math.max(0, maxAttempts|0);
      payload = await retryBootstrapUntilCached(itemType, fmt, scenario, attempts, 250);
      if(!payload && itemType==='appointment'){
        const fb = makeKey('message', fmt, 'new');
        payload = cacheGet(fb) || await retryBootstrapUntilCached('message', fmt, 'new', attempts, 250);
      }
    }

    if(!payload) return false;
    return await insertSignature(payload, fmt, scenario, contextKind);
  }

  // ---------------- Handlers ----------------
  const RETRY_EVENT=2, RETRY_BUTTON=3;
  let __eventInsertedOnce=false;
  async function onMessageCompose(event){ try{ await Office.onReady(); if(!__eventInsertedOnce){ __eventInsertedOnce=true; await doInsertCurrentContext(RETRY_EVENT,'event'); } } finally { try{event&&event.completed&&event.completed();}catch(_){}} }
  async function onReinsert(event){ try{ await Office.onReady(); await doInsertCurrentContext(RETRY_BUTTON,'button'); } finally { try{event&&event.completed&&event.completed();}catch(_){}} }

  // ---------------- Bootstrap ----------------
  (async function(){ await loadMinimalConfig(); log = mkLogger('info'); try{ await Office.onReady(); }catch(_){ } bootstrapSignaturesSilently(); try{ Office.actions&&Office.actions.associate&& ( Office.actions.associate('WPSecure_OnMessageCompose', onMessageCompose), Office.actions.associate('WPSecure_Reinsert', onReinsert) ); }catch(_){ } })();
})();
