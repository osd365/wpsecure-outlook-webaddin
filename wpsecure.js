/* WPSecure: single-file logic (NAA + Graph + strict format insertion) */
(function(){
  const LOG_KEY = 'wpsecure_log_level';
  const DEFAULT_LOG_LEVEL = (localStorage.getItem(LOG_KEY) || 'info').toLowerCase();
  const levels = ['off','error','warn','info','debug'];
  function log(level, msg, meta){
    const idx = levels.indexOf(DEFAULT_LOG_LEVEL);
    const lid = levels.indexOf(level);
    if(idx === -1 || lid === -1 || lid>idx || DEFAULT_LOG_LEVEL==='off') return;
    const prefix = `[WPSecure][${level}]`;
    if(level==='error') console.error(prefix, msg, meta||'');
    else if(level==='warn') console.warn(prefix, msg, meta||'');
    else if(level==='debug') console.debug(prefix, msg, meta||'');
    else console.log(prefix, msg, meta||'');
    const st = document.getElementById('status');
    if(st && level!=='debug') st.textContent = `${new Date().toISOString()} ${prefix} ${msg} ${meta?JSON.stringify(meta):''}`;
  }

  // --- Configuration placeholders (update after app registration) ---
  const MSAL_CLIENT_ID = 'YOUR-CLIENT-ID';
  const MSAL_TENANT_ID = 'common'; // or your tenant GUID
  const ORIGIN_HOST = location.host; // used for NAA redirect
  const GRAPH_SCOPES = ['Files.Read'];

  // NAA redirect must be: brk-multihub://<origin>
  const msalConfig = {
    auth: {
      clientId: MSAL_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${MSAL_TENANT_ID}`,
      redirectUri: `brk-multihub://${ORIGIN_HOST}`
    },
    system: {
      loggerOptions: { loggerCallback: (level, message)=>{ if(DEFAULT_LOG_LEVEL==='debug') console.debug('[MSAL]', message); } }
    },
    cache: { cacheLocation: 'sessionStorage' }
  };

  const msalApp = new msal.PublicClientApplication(msalConfig);
  let account = null; // set after login / silent

  // Avoid prompting during events; only prompt on button click.
  async function getAccessToken({interactive}){
    try {
      const accts = msalApp.getAllAccounts();
      if(accts.length>0) account = accts[0];
      if(!account && !interactive) return null;
      if(!account && interactive){
        const loginResp = await msalApp.loginPopup({ scopes: GRAPH_SCOPES });
        account = loginResp.account;
      }
      if(!account) return null;
      const result = await msalApp.acquireTokenSilent({ account, scopes: GRAPH_SCOPES }).catch(async (e)=>{
        if(interactive) return msalApp.acquireTokenPopup({ account, scopes: GRAPH_SCOPES });
        throw e;
      });
      return result && result.accessToken || null;
    } catch(e){ log('warn','token:failed', { message: e.message }); return null; }
  }

  function getUserIdFragment(){
    try {
      const accts = msalApp.getAllAccounts();
      const acc = accts && accts[0];
      const oid = acc && acc.idTokenClaims && (acc.idTokenClaims.oid || acc.idTokenClaims.sub || 'u');
      return String(oid||'u').replace(/-/g,'').slice(0,6);
    } catch{ return 'user'; }
  }

  // Counters (per session)
  function counterKey(kind){
    return `wpsecure_retry_${kind}_${getUserIdFragment()}_${Office.context.sessionId}`;
  }
  function getCounter(kind){ return parseInt(localStorage.getItem(counterKey(kind)||'0'))||0; }
  function incCounter(kind){ const k = counterKey(kind); const v = getCounter(kind)+1; localStorage.setItem(k, String(v)); return v; }
  function capReached(kind){ return getCounter(kind) >= 10; }

  // Cache keys
  function keyMsg(fmt, scenario){ return `wpsecure_sig_${getUserIdFragment()}_${Office.context.sessionId}_message_${fmt}_${scenario}`; }
  function keyApt(fmt){ return `wpsecure_sig_${getUserIdFragment()}_${Office.context.sessionId}_appointment_${fmt}_new`; }

  // OneDrive (Graph) paths
  const PATHS = {
    msg: {
      HTML: { new: '/me/drive/root:/\.wpsecure/wpsecure_cloud_new.htm:/content', reply: '/me/drive/root:/\.wpsecure/wpsecure_cloud_reply.htm:/content' },
      txt:  { new: '/me/drive/root:/\.wpsecure/wpsecure_cloud_new.txt:/content', reply: '/me/drive/root:/\.wpsecure/wpsecure_cloud_reply.txt:/content' }
    },
    apt: {
      HTML: { new: '/me/drive/root:/\.wpsecure/wpsecure_cloud_new.htm:/content' },
      txt:  { new: '/me/drive/root:/\.wpsecure/wpsecure_cloud_new.txt:/content' }
    }
  };

  async function fetchTemplate(graphToken, url){
    const r = await fetch(`https://graph.microsoft.com/v1.0${url}`, { headers: { Authorization: `Bearer ${graphToken}` }});
    if(!r.ok){ throw new Error(`graph ${r.status}`); }
    return await r.text();
  }

  function cacheSet(k,v){ try{ localStorage.setItem(k, v); }catch(e){ log('warn','cache:set failed',{k}); } }
  function cacheGet(k){ try{ return localStorage.getItem(k); }catch{ return null; } }

  async function ensureCached(kind, fmt, scenario, isInteractive){
    // kind: 'msg'|'apt'
    // fmt: 'HTML'|'txt'
    // scenario: 'new'|'reply'
    const key = kind==='msg' ? keyMsg(fmt, scenario) : keyApt(fmt);
    const existing = cacheGet(key);
    if(existing) { log('info','cache:hit',{key}); return true; }

    // Decide which counter to use
    const counterKind = isInteractive ? 'button' : 'event';
    if(capReached(counterKind)) { log('info','retry cap reached', {counterKind}); return false; }

    // Attempt fetch
    const token = await getAccessToken({interactive: isInteractive});
    if(!token){
      // no token (silent failed, or user cancelled); count attempt only if we intended to fetch
      incCounter(counterKind);
      log('info','token unavailable; aborting fetch',{counterKind});
      return false;
    }

    const path = kind==='msg' ? PATHS.msg[fmt][scenario] : PATHS.apt[fmt].new;
    log('info','graph:fetch start',{path});
    try{
      const text = await fetchTemplate(token, path);
      cacheSet(key, text);
      log('info','graph:fetch ok',{key});
      return true;
    }catch(e){
      incCounter(counterKind);
      log('info','graph:fetch missing or failed',{message:e.message, path});
      return false;
    }
  }

  function getBodyFormat(){
    return new Promise(resolve=>{
      Office.context.mailbox.item.body.getTypeAsync(res=>{
        if(res.status===Office.AsyncResultStatus.Succeeded){
          resolve(res.value===Office.MailboxEnums.BodyType.Html ? 'HTML':'txt');
        } else { resolve('HTML'); }
      });
    });
  }

  function getComposeScenario(){
    return new Promise(resolve=>{
      if(!Office.context.mailbox.item.getComposeTypeAsync){
        return resolve('new'); // older clients: treat as new
      }
      Office.context.mailbox.item.getComposeTypeAsync(ar=>{
        if(ar.status!==Office.AsyncResultStatus.Succeeded) return resolve('new');
        const ct = ar.value && ar.value.composeType; // newMail|reply|forward
        resolve(ct==='reply' || ct==='forward' ? 'reply' : 'new');
      });
    });
  }

  async function insertSignatureStrict(content, fmt){
    return new Promise(resolve=>{
      const opts = { coercionType: fmt==='HTML' ? Office.CoercionType.Html : Office.CoercionType.Text };
      if(Office.context.mailbox.item.body.setSignatureAsync){
        Office.context.mailbox.item.body.setSignatureAsync(content, opts, r=>{
          if(r.status===Office.AsyncResultStatus.Succeeded) return resolve(true);
          log('warn','setSignatureAsync failed; trying setAsync', { code: r.error && r.error.code, message: r.error && r.error.message });
          Office.context.mailbox.item.body.setAsync(content, opts, r2=>{
            if(r2.status===Office.AsyncResultStatus.Succeeded) return resolve(true);
            log('error','setAsync failed', { code: r2.error && r2.error.code, message: r2.error && r2.error.message });
            resolve(false);
          });
        });
      } else {
        Office.context.mailbox.item.body.setAsync(content, opts, r2=>{
          if(r2.status===Office.AsyncResultStatus.Succeeded) return resolve(true);
          log('error','setAsync failed', { code: r2.error && r2.error.code, message: r2.error && r2.error.message });
          resolve(false);
        });
      }
    });
  }

  async function doInsert({interactive, itemType}){
    const fmt = await getBodyFormat();
    const scenario = itemType==='message' ? await getComposeScenario() : 'new';
    log('info','compose context',{itemType, fmt, scenario});

    // Strict policy: require matching format entry
    const kind = itemType==='message' ? 'msg' : 'apt';
    const ok = await ensureCached(kind, fmt, scenario, interactive);
    if(!ok){ return { inserted:false, reason:'missing' }; }

    const key = itemType==='message' ? keyMsg(fmt, scenario) : keyApt(fmt);
    const payload = cacheGet(key);
    if(!payload){ return { inserted:false, reason:'missing' }; }

    const done = await insertSignatureStrict(payload, fmt);
    return { inserted: !!done, reason: done? 'ok':'error' };
  }

  // Notifications (toasts) only on button flow
  function toast(kind, message){
    try{
      const id = 'wpsecure-toast';
      Office.context.mailbox.item.notificationMessages.replaceAsync(id, { type: kind, message, icon: 'icon-16', persistent: false });
    }catch(e){ /* ignore */ }
  }

  // ----- Command functions -----
  async function onMessageCompose(event){
    try{
      await Office.onReady();
      // Silent attempt: no prompting, no toasts
      const res = await doInsert({ interactive:false, itemType:'message' });
      log('info','event insert result', res);
    } catch(e){ log('error','onMessageCompose error', { message:e.message }); }
    finally { event.completed(); }
  }

  async function onReinsert(event){
    try{
      await Office.onReady();
      const itemType = Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message ? 'message':'appointment';
      const res = await doInsert({ interactive:true, itemType });
      if(res.inserted){ toast('informationalMessage', 'WPSecure: Signature inserted.'); }
      else if(res.reason==='missing'){
        const fmt = await getBodyFormat();
        const label = fmt==='HTML' ? 'HTML' : 'Plain-text';
        // Button counter handling for special 10-tries toast
        const cBefore = getCounter('button');
        const capped = capReached('button');
        if(capped || cBefore>=10){
          toast('errorMessage', 'WPSecure: Tried 10 times. Please refresh Outlook or restart your session, then try again.');
        } else {
          toast('errorMessage', `WPSecure: ${label} signature for this scenario isn’t available.`);
        }
      } else {
        toast('errorMessage', 'WPSecure: Couldn’t insert signature due to a system error. Please try again.');
      }
    } catch(e){ toast('errorMessage','WPSecure: Unexpected error.'); log('error','onReinsert error',{message:e.message}); }
    finally { event.completed && event.completed(); }
  }

  // Expose for HTML button too
  window.WPSecure_Reinsert_Action = async ()=>{
    // Faux-event path for in-pane button
    await onReinsert({ completed: ()=>{} });
  };

  // Associate command handlers
  Office.actions.associate('WPSecure_OnMessageCompose', onMessageCompose);
  Office.actions.associate('WPSecure_Reinsert', onReinsert);

})();
