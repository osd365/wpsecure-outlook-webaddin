/* WPSecure (command-only) with external config.json */
(function(){
  // ---------------- Logging ----------------
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

  // --------------- External config loader ---------------
  let CONFIG = null;
  async function loadConfig(){
    try{
      const r = await fetch('./config.json', { cache: 'no-store' });
      if(!r.ok) throw new Error(`config ${r.status}`);
      CONFIG = await r.json();
    }catch(e){
      console.error('[WPSecure] Failed to load config.json', e);
      CONFIG = { MSAL_CLIENT_ID:'unset', MSAL_TENANT_ID:'common', GRAPH_SCOPES:['Files.Read'], DEFAULT_LOG_LEVEL:'info',
        TEMPLATE_PATHS:{ newHtml:'/.wpsecure/wpsecure_cloud_new.htm', replyHtml:'/.wpsecure/wpsecure_cloud_reply.htm', newText:'/.wpsecure/wpsecure_cloud_new.txt', replyText:'/.wpsecure/wpsecure_cloud_reply.txt' },
        STRICT_FORMAT:true, RETRY_CAP_EVENT:10, RETRY_CAP_BUTTON:10, ENABLE_APPOINTMENTS:true,
        TOASTS:{ success:'WPSecure: Signature inserted.', missingHtml:"WPSecure: HTML signature for this scenario isn’t available.", missingText:"WPSecure: Plain-text signature for this scenario isn’t available.", tenTries:'WPSecure: Tried 10 times. Please refresh Outlook or restart your session, then try again.', systemError:'WPSecure: Couldn’t insert signature due to a system error. Please try again.'},
        CACHE_KEY_PREFIX:'wpsecure_sig_', COUNTER_KEY_PREFIX:'wpsecure_retry_', SESSION_SCOPED:true
      };
    }
  }

  // Will be initialized after config is loaded
  let log = ()=>{};

  // --------------- MSAL (NAA) ---------------
  function redirectOrigin(){
    if(CONFIG && CONFIG.REDIRECT_ORIGIN) return CONFIG.REDIRECT_ORIGIN; // full origin host if provided
    return location.host; // origin-only host for brk-multihub
  }

  let msalApp = null, account = null;
  async function initMsal(){
    const msalConfig = {
      auth: {
        clientId: CONFIG.MSAL_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${CONFIG.MSAL_TENANT_ID || 'common'}`,
        redirectUri: `brk-multihub://${redirectOrigin()}`
      },
      system: { loggerOptions: { loggerCallback: (_l,m)=>{ if(localStorage.getItem('wpsecure_log_level')==='debug') console.debug('[MSAL]',m); } } },
      cache: { cacheLocation: 'sessionStorage' }
    };
    msalApp = new msal.PublicClientApplication(msalConfig);
  }

  async function getAccessToken({interactive}){
    try{
      const accts = msalApp.getAllAccounts();
      if(accts.length>0) account = accts[0];
      if(!account && !interactive) return null;
      if(!account && interactive){
        const loginResp = await msalApp.loginPopup({ scopes: CONFIG.GRAPH_SCOPES||['Files.Read'] });
        account = loginResp.account;
      }
      if(!account) return null;
      const result = await msalApp.acquireTokenSilent({ account, scopes: CONFIG.GRAPH_SCOPES||['Files.Read'] }).catch(async (e)=>{
        if(interactive) return msalApp.acquireTokenPopup({ account, scopes: CONFIG.GRAPH_SCOPES||['Files.Read'] });
        throw e;
      });
      return result && result.accessToken || null;
    }catch(e){ log('warn','token:failed',{message:e.message}); return null; }
  }

  function userFrag(){
    try{
      const accts = msalApp.getAllAccounts();
      const acc = accts && accts[0];
      const oid = acc && acc.idTokenClaims && (acc.idTokenClaims.oid || acc.idTokenClaims.sub || 'u');
      return String(oid||'u').replace(/-/g,'').slice(0,6);
    }catch{ return 'user'; }
  }

  // --------------- Counters & cache ---------------
  function sessionId(){ return (CONFIG.SESSION_SCOPED!==false && Office && Office.context && Office.context.sessionId) ? Office.context.sessionId : 'global'; }
  function cKey(kind){ return `${CONFIG.COUNTER_KEY_PREFIX||'wpsecure_retry_'}${kind}_${userFrag()}_${sessionId()}`; }
  function getCounter(kind){ return parseInt(localStorage.getItem(cKey(kind))||'0')||0; }
  function incCounter(kind){ const k=cKey(kind); const v=getCounter(kind)+1; localStorage.setItem(k,String(v)); return v; }
  function capReached(kind){ const cap = kind==='button'? (CONFIG.RETRY_CAP_BUTTON||10) : (CONFIG.RETRY_CAP_EVENT||10); return getCounter(kind) >= cap; }

  function kMsg(fmt, scenario){ return `${CONFIG.CACHE_KEY_PREFIX||'wpsecure_sig_'}${userFrag()}_${sessionId()}_message_${fmt}_${scenario}`; }
  function kApt(fmt){ return `${CONFIG.CACHE_KEY_PREFIX||'wpsecure_sig_'}${userFrag()}_${sessionId()}_appointment_${fmt}_new`; }

  function currentUserEmail() {
  // Office.js exposes the SMTP address of the mailbox user
  return Office.context.mailbox.userProfile.emailAddress || '';
}

function applyPlaceholders(relPath) {
  return relPath.replace(/{email}/g, currentUserEmail());
}

function paths(kind, fmt, scenario){
  const p = CONFIG.TEMPLATE_PATHS || {};
  let rel =
    (kind === 'msg')
      ? (fmt === 'HTML'
          ? (scenario === 'reply' ? p.replyHtml : p.newHtml)
          : (scenario === 'reply' ? p.replyText : p.newText))
      : (fmt === 'HTML' ? p.newHtml : p.newText); // appointments use NEW only
  return applyPlaceholders(rel || '');
}

  async function fetchTemplate(token, path){
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURI(path.replace(/^\/+/,''))}:/content`;
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` }});
    if(!r.ok) throw new Error(`graph ${r.status}`);
    return await r.text();
  }

  function cacheSet(k,v){ try{ localStorage.setItem(k,v); }catch(e){ log('warn','cache:set failed',{k}); } }
  function cacheGet(k){ try{ return localStorage.getItem(k); }catch{ return null; } }

  async function ensureCached(kind, fmt, scenario, isInteractive){
    const key = (kind==='msg') ? kMsg(fmt,scenario) : kApt(fmt);
    if(cacheGet(key)) { log('info','cache:hit',{key}); return true; }

    const counterKind = isInteractive ? 'button' : 'event';
    if(capReached(counterKind)) { log('info','retry cap reached',{counterKind}); return false; }

    const token = await getAccessToken({interactive:isInteractive});
    if(!token){ incCounter(counterKind); log('info','token unavailable; abort fetch',{counterKind}); return false; }

    const relPath = paths(kind, fmt, scenario);
    log('info','graph:fetch start',{relPath});
    try{
      const text = await fetchTemplate(token, relPath);
      cacheSet(key, text);
      log('info','graph:fetch ok',{key});
      return true;
    }catch(e){ incCounter(counterKind); log('info','graph:fetch failed',{message:e.message, relPath}); return false; }
  }

  function getBodyFormat(){
    return new Promise(resolve=>{
      Office.context.mailbox.item.body.getTypeAsync(res=>{
        if(res.status===Office.AsyncResultStatus.Succeeded){
          resolve(res.value===Office.MailboxEnums.BodyType.Html ? 'HTML' : 'txt');
        } else { resolve('HTML'); }
      });
    });
  }

  function getComposeScenario(){
    return new Promise(resolve=>{
      if(!Office.context.mailbox.item.getComposeTypeAsync) return resolve('new');
      Office.context.mailbox.item.getComposeTypeAsync(ar=>{
        if(ar.status!==Office.AsyncResultStatus.Succeeded) return resolve('new');
        const ct = ar.value && ar.value.composeType; // newMail|reply|forward
        resolve(ct==='reply' || ct==='forward' ? 'reply' : 'new');
      });
    });
  }

  async function insertSignature(content, fmt){
    return new Promise(resolve=>{
      const opts = { coercionType: fmt==='HTML' ? Office.CoercionType.Html : Office.CoercionType.Text };
      if(Office.context.mailbox.item.body.setSignatureAsync){
        Office.context.mailbox.item.body.setSignatureAsync(content, opts, r=>{
          if(r.status===Office.AsyncResultStatus.Succeeded) return resolve(true);
          log('warn','setSignatureAsync failed; trying setAsync',{code:r.error&&r.error.code,message:r.error&&r.error.message});
          Office.context.mailbox.item.body.setAsync(content, opts, r2=>{
            if(r2.status===Office.AsyncResultStatus.Succeeded) return resolve(true);
            log('error','setAsync failed',{code:r2.error&&r2.error.code,message:r2.error&&r2.error.message});
            resolve(false);
          });
        });
      } else {
        Office.context.mailbox.item.body.setAsync(content, opts, r2=>{
          if(r2.status===Office.AsyncResultStatus.Succeeded) return resolve(true);
          log('error','setAsync failed',{code:r2.error&&r2.error.code,message:r2.error&&r2.error.message});
          resolve(false);
        });
      }
    });
  }

  async function doInsert({interactive, itemType}){
    const fmt = await getBodyFormat();
    const scenario = itemType==='message' ? await getComposeScenario() : 'new';
    log('info','compose context',{itemType, fmt, scenario});

    // Strict policy: require matching format
    if(CONFIG.STRICT_FORMAT!==false){ /* always strict */ }

    const kind = itemType==='message' ? 'msg' : 'apt';
    const ok = await ensureCached(kind, fmt, scenario, interactive);
    if(!ok) return { inserted:false, reason:'missing' };

    const key = (kind==='msg') ? kMsg(fmt, scenario) : kApt(fmt);
    const payload = cacheGet(key);
    if(!payload) return { inserted:false, reason:'missing' };

    const done = await insertSignature(payload, fmt);
    return { inserted: !!done, reason: done? 'ok':'error' };
  }

  function toast(kind, message){
    try{
      const id='wpsecure-toast';
      Office.context.mailbox.item.notificationMessages.replaceAsync(id,{type:kind,message,icon:'icon-16',persistent:false});
    }catch(e){ /* ignore */ }
  }

  async function onMessageCompose(event){
    try{
      await Office.onReady();
      const res = await doInsert({ interactive:false, itemType:'message' });
      log('info','event insert result', res);
    }catch(e){ log('error','onMessageCompose error',{message:e.message}); }
    finally { event.completed(); }
  }

  async function onReinsert(event){
    try{
      await Office.onReady();
      const itemType = Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message ? 'message' : 'appointment';
      if(itemType==='appointment' && CONFIG.ENABLE_APPOINTMENTS===false){ toast('errorMessage','WPSecure: Appointments disabled.'); event.completed&&event.completed(); return; }
      const res = await doInsert({ interactive:true, itemType });
      if(res.inserted){ toast('informationalMessage', CONFIG.TOASTS && CONFIG.TOASTS.success || 'WPSecure: Signature inserted.'); }
      else if(res.reason==='missing'){
        const fmt = await getBodyFormat();
        const label = fmt==='HTML' ? 'HTML' : 'Plain-text';
        const capped = (getCounter('button') >= (CONFIG.RETRY_CAP_BUTTON||10));
        if(capped) toast('errorMessage', (CONFIG.TOASTS && CONFIG.TOASTS.tenTries) || 'WPSecure: Tried 10 times. Please refresh Outlook or restart your session, then try again.');
        else toast('errorMessage', (fmt==='HTML' ? (CONFIG.TOASTS && CONFIG.TOASTS.missingHtml) : (CONFIG.TOASTS && CONFIG.TOASTS.missingText)) || 'Signature not available.');
      } else {
        toast('errorMessage', (CONFIG.TOASTS && CONFIG.TOASTS.systemError) || 'WPSecure: Couldn’t insert signature due to a system error.');
      }
    }catch(e){ toast('errorMessage', (CONFIG.TOASTS && CONFIG.TOASTS.systemError) || 'WPSecure: Unexpected error.'); log('error','onReinsert error',{message:e.message}); }
    finally { event.completed && event.completed(); }
  }

  // Bootstrap: load config → init logger → init MSAL → wire handlers
  (async function bootstrap(){
    await loadConfig();
    log = mkLogger(CONFIG.DEFAULT_LOG_LEVEL||'info');
    await initMsal();
    Office.actions.associate('WPSecure_OnMessageCompose', onMessageCompose);
    Office.actions.associate('WPSecure_Reinsert', onReinsert);
  })();

})();
