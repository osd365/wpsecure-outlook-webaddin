/* WPSecure Outlook Add-in Runtime — SSO+OBO Silent Cache Edition
 * Build: 2026-03-01T00:00Z (message HTML unchanged; TEXT no visible markers via sessionData; SAFE appointment HTML)
 * Policy:
 * - Silent UX (no toasts). Console logs only (no PII, no tokens).
 * - Cache-only insertion on button/event; do nothing if cache miss.
 * - Bootstrap once per session: fetch up to 4 templates via backend (OBO), cache non-null.
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
  let CONFIG = { BACKEND_OBO_ENDPOINT: '', BACKEND_BOOTSTRAP_ENDPOINT: '' };
  async function loadMinimalConfig(){
    try{
      const r = await fetch('./config.json', { cache: 'no-store' });
      if(r.ok){
        const json = await r.json();
        if(json && typeof json==='object'){
          if(json.BACKEND_OBO_ENDPOINT) CONFIG.BACKEND_OBO_ENDPOINT = json.BACKEND_OBO_ENDPOINT;
          if(json.BACKEND_BOOTSTRAP_ENDPOINT) CONFIG.BACKEND_BOOTSTRAP_ENDPOINT = json.BACKEND_BOOTSTRAP_ENDPOINT;
        }
      }
    }catch(e){ /* keep defaults */ }
  }
  function computeBootstrapUrl(){
    if(CONFIG.BACKEND_BOOTSTRAP_ENDPOINT) return CONFIG.BACKEND_BOOTSTRAP_ENDPOINT;
    try{
      if(!CONFIG.BACKEND_OBO_ENDPOINT) return '/api/signatures/bootstrap';
      const u = new URL(CONFIG.BACKEND_OBO_ENDPOINT, location.origin);
      u.pathname = '/api/signatures/bootstrap';
      return u.pathname;
    }catch(_){ return '/api/signatures/bootstrap'; }
  }

  // ---------------- Helpers ----------------
  let log = ()=>{};
  function sessionId(){ try{ return (Office && Office.context && Office.context.sessionId) || 'global'; }catch{ return 'global'; } }
  function userFrag(){ try{ const up=Office.context.mailbox.userProfile; return (up && up.emailAddress || 'user').replace(/[^a-z0-9]/gi,'').slice(0,10) || 'user'; }catch{ return 'user'; } }
  function makeKey(kind, fmt, scenario){ // kind: 'message' | 'appointment'; fmt: 'HTML'|'txt'; scenario: 'new'|'reply'
    const u = userFrag(); const s = sessionId();
    if(kind==='appointment') return `wpsecure_sig_${u}_${s}_appointment_${fmt}_new`;
    return `wpsecure_sig_${u}_${s}_message_${fmt}_${scenario}`;
  }
  function cacheSet(k,v){ try{ localStorage.setItem(k,v); }catch(_){ /* ignore */ } }
  function cacheGet(k){ try{ return localStorage.getItem(k); }catch(_){ return null; } }
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
      try { Office.context.mailbox.item.body.disableClientSignatureAsync(() => resolve()); }
      catch(_) { resolve(); }
    });
  }
  function getBodyText() {
    return new Promise(resolve => {
      try {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, ar => {
          if (ar.status === Office.AsyncResultStatus.Succeeded) resolve(ar.value || '');
          else resolve('');
        });
      } catch(_) { resolve(''); }
    });
  }
  function setBodyText(text) {
    return new Promise(resolve => {
      try {
        Office.context.mailbox.item.body.setAsync(text, { coercionType: Office.CoercionType.Text }, r => {
          resolve(r.status === Office.AsyncResultStatus.Succeeded);
        });
      } catch(_) { resolve(false); }
    });
  }
  function getBodyHtml() {
    return new Promise(resolve => {
      try {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, ar => {
          if (ar.status === Office.AsyncResultStatus.Succeeded) resolve(ar.value || '');
          else resolve('');
        });
      } catch(_) { resolve(''); }
    });
  }
  function setBodyHtml(html) {
    return new Promise(resolve => {
      try {
        Office.context.mailbox.item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, r => {
          resolve(r.status === Office.AsyncResultStatus.Succeeded);
        });
      } catch(_) { resolve(false); }
    });
  }

  // ---------------- TEXT: sessionData-based replacement (no visible markers) ----------------
  function getSessionTextSig(){
    return new Promise(resolve => {
      try {
        Office.context.mailbox.item.sessionData.getAsync('wpsecure_txt_sig', ar => {
          resolve(ar.status === Office.AsyncResultStatus.Succeeded ? (ar.value || '') : '');
        });
      } catch(_) { resolve(''); }
    });
  }
  function setSessionTextSig(val){
    return new Promise(resolve => {
      try {
        Office.context.mailbox.item.sessionData.setAsync('wpsecure_txt_sig', val || '', ar => resolve());
      } catch(_) { resolve(); }
    });
  }

  // ---------------- Insertion (message HTML unchanged; TEXT uses sessionData; appointment HTML safe merge) ----------------
  async function insertSignature(content, fmt){
    // Determine item type inline
    const isMessage = (function(){ try{ return Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message; }catch(_){ return true; } })();
    const isAppointment = !isMessage;

    if (fmt === 'HTML') {
      if (isAppointment) {
        // SAFE appointment HTML path: never wipe user content, add spacer, replace our block only
        try {
          const current = await getBodyHtml();
          const wrapperStart = '<div data-wpsecure="sig-appt">';
          const wrapperEnd   = '</div>';
          const wrappedSig   = `${wrapperStart}${content}${wrapperEnd}`;
          let nextHtml;

          if (current && current.indexOf(wrapperStart) >= 0) {
            const re = new RegExp(`${wrapperStart}[\\s\\S]*?${wrapperEnd}`,'i');
            nextHtml = current.replace(re, wrappedSig);
          } else {
            const spacer = '<div><br></div>';
            nextHtml = (current || '') + spacer + wrappedSig;
          }

          // We are about to insert -> disable client signature first
          try { await disableClientSignature(); } catch(_) {}
          const ok = await setBodyHtml(nextHtml);
          return ok;
        } catch(_) { return false; }
      }

      // HTML Message path (UNCHANGED): replace via setSignatureAsync
      return new Promise(resolve => {
        const opts = { coercionType: Office.CoercionType.Html };
        try {
          if (Office.context.mailbox.item.body.setSignatureAsync) {
            // We are about to insert -> disable client signature first
            disableClientSignature().then(()=>{
              Office.context.mailbox.item.body.setSignatureAsync(content, opts, r => {
                resolve(r.status === Office.AsyncResultStatus.Succeeded);
              });
            }).catch(()=>{
              Office.context.mailbox.item.body.setSignatureAsync(content, opts, r => {
                resolve(r.status === Office.AsyncResultStatus.Succeeded);
              });
            });
          } else {
            resolve(false);
          }
        } catch(_) { resolve(false); }
      });
    }

    // TEXT path: emulate REPLACE without visible markers using sessionData
    try {
      const current = await getBodyText();
      const prev = await getSessionTextSig();
      let baseline = current;

      // If the previous WPSecure text signature is known, strip it first (idempotent)
      if (prev && current && current.indexOf(prev) >= 0) {
        baseline = current.replace(prev, '');
      }

      const nextBody = (baseline || '') + (baseline ? "\n\n" : "") + content;

      // We are about to insert -> disable client signature first
      try { await disableClientSignature(); } catch(_) {}

      const ok = await setBodyText(nextBody);
      if (ok) await setSessionTextSig(content);
      return ok;
    } catch(_) { return false; }
  }

  // ---------------- SSO + Bootstrap (silent) ----------------
  async function getSsoIdTokenSilent(){
    try{ return await Office.auth.getAccessToken({ allowSignInPrompt:false, allowConsentPrompt:false }); }
    catch(_){ return null; }
  }
  async function bootstrapSignaturesSilently(){
    try{
      log('info','bootstrap:start');
      const url = computeBootstrapUrl();
      if(!url){ log('warn','bootstrap:missing-endpoint'); return; }
      const idTok = await getSsoIdTokenSilent();
      if(!idTok){ log('warn','bootstrap:no-idtoken'); return; }
      const res = await fetch(url, {
        method: 'POST', headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ id_token: idTok })
      });
      if(!res.ok){ log('error','bootstrap:http', { status: res.status }); return; }
      const payload = await res.json();
      const files = (payload && payload.files) || {};
      if(files.newHtml) cacheSet(makeKey('message','HTML','new'), files.newHtml);
      if(files.replyHtml)cacheSet(makeKey('message','HTML','reply'), files.replyHtml);
      if(files.newText) cacheSet(makeKey('message','txt','new'), files.newText);
      if(files.replyText)cacheSet(makeKey('message','txt','reply'), files.replyText);
      if(files.apptNewHtml) cacheSet(makeKey('appointment','HTML','new'), files.apptNewHtml);
      if(files.apptNewText) cacheSet(makeKey('appointment','txt','new'), files.apptNewText);
      log('info','bootstrap:files', {
        newHtml: !!files.newHtml, replyHtml: !!files.replyHtml,
        newText: !!files.newText, replyText: !!files.replyText,
        apptNewHtml: !!files.apptNewHtml, apptNewText: !!files.apptNewText
      });
    }catch(_){ log('error','bootstrap:error'); }
  }

  // ---------------- Insertion Orchestrator (cache-only) ----------------
  async function doInsertCurrentContext(){
    const itemType = (function(){ try{
      return Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message ? 'message' : 'appointment';
    }catch(_){ return 'message'; } })();
    const fmt = await getBodyFormat();
    const scenario = itemType==='message' ? await getComposeScenario() : 'new';
    log('info','insert:scenario', { item:itemType, fmt, compose:scenario });

    const key = makeKey(itemType, fmt, scenario);
    let payload = cacheGet(key);

    // Appointment fallback to message-new if appointment cache is empty
    if(!payload && itemType === 'appointment') {
      const fallbackKey = makeKey('message', fmt, 'new');
      payload = cacheGet(fallbackKey);
      if (payload) log('info','insert:appt-fallback', { from: key, to: fallbackKey });
    }

    if(!payload){ log('warn','insert:cache-miss', { key }); return false; }
    const ok = await insertSignature(payload, fmt);
    log('info', ok ? 'insert:done' : 'insert:failed');
    return ok;
  }

  // ---------------- Event & Button handlers ----------------
  let __eventInsertedOnce = false;
  async function onMessageCompose(event){
    try{
      await Office.onReady();
      if(!__eventInsertedOnce){
        __eventInsertedOnce = true;
        await doInsertCurrentContext();
      } else {
        log('debug','event:skipped-second-run');
      }
    } catch(_) { }
    finally { try{ event && event.completed && event.completed(); }catch(_){} }
  }
  async function onReinsert(event){
    try{ await Office.onReady(); await doInsertCurrentContext(); }
    catch(_){ }
    finally{ try{ event && event.completed && event.completed(); }catch(_){} }
  }

  // ---------------- Bootstrap & wire-up ----------------
  (async function bootstrap(){
    await loadMinimalConfig();
    log = mkLogger('info');
    try{ await Office.onReady(); }catch(_){ }
    bootstrapSignaturesSilently();
    try{
      Office.actions && Office.actions.associate && (
        Office.actions.associate('WPSecure_OnMessageCompose', onMessageCompose),
        Office.actions.associate('WPSecure_Reinsert', onReinsert)
      );
      log('info','wireup:associated');
    }catch(_){ log('warn','wireup:failed'); }
  })();

  // Optional debug surface (no secrets)
  window.WPSecure_Debug = {
    hasConfig: () => !!CONFIG.BACKEND_OBO_ENDPOINT || !!CONFIG.BACKEND_BOOTSTRAP_ENDPOINT,
    cacheKeys: () => Object.keys(localStorage).filter(k=>k.startsWith('wpsecure_sig_'))
  };
})();
