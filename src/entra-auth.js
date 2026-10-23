import {PublicClientApplication} from "@azure/msal-browser";
let memory=null,pending=null;
const settings=()=>window.RBD_CONFIG?.secureApi||{};
const authError=message=>Object.assign(Error(message),{code:"AUTH_REQUIRED"});
function configuration(cacheLocation="memoryStorage"){
  const c=settings();if(!/^[a-f\d-]{36}$/i.test(c.tenantId)||!/^[a-f\d-]{36}$/i.test(c.clientId)||!c.scope)throw authError("Не налаштовано корпоративний вхід РБД");
  return {auth:{clientId:c.clientId,authority:`https://login.microsoftonline.com/${c.tenantId}`,redirectUri:new URL("./rbd-auth.html",location.href).href.split("?")[0],navigateToLoginRequestUrl:false},cache:{cacheLocation}};
}
function remember(result){memory={token:result.accessToken,expires:result.expiresOn instanceof Date?result.expiresOn.getTime():Number(result.expires)};if(!memory.token||!Number.isFinite(memory.expires))throw authError("Microsoft не повернув токен API");return memory.token;}
export async function getAccessToken(){if(memory&&memory.expires>Date.now()+60000)return memory.token;throw authError("Для доступу до приватних листів увійдіть через Microsoft у меню ⋮.");}
export function clear(){memory=null;}
export async function signIn(){
  if(pending)return pending;
  pending=(async()=>{
    if(window.Office?.context?.ui?.displayDialogAsync){
      return new Promise((resolve,reject)=>{
        const nonce=crypto.randomUUID();const url=new URL("./rbd-auth.html",location.href);url.searchParams.set("nonce",nonce);
        Office.context.ui.displayDialogAsync(url.href,{height:65,width:40,displayInIframe:false},result=>{
          if(result.status!==Office.AsyncResultStatus.Succeeded){reject(authError("Не вдалося відкрити корпоративний вхід: "+result.error?.message));return;}
          const dialog=result.value;let done=false;
          const finish=(error,data)=>{if(done)return;done=true;clearTimeout(timer);dialog.close();error?reject(error):resolve(remember(data));};
          const timer=setTimeout(()=>finish(authError("Час входу вичерпано")),5*60*1000);
          dialog.addEventHandler(Office.EventType.DialogMessageReceived,arg=>{
            if(arg.origin!==location.origin)return;
            try{const data=JSON.parse(arg.message);if(data.nonce!==nonce)return;finish(data.error?authError(data.error):null,data);}catch{finish(authError("Некоректна відповідь корпоративного входу"));}
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived,()=>finish(authError("Вхід скасовано")));
        });
      });
    }
    const app=new PublicClientApplication(configuration());await app.initialize();const result=await app.acquireTokenPopup({scopes:[settings().scope]});return remember(result);
  })();
  try{return await pending;}finally{pending=null;}
}
export async function runDialog(){
  const text=document.getElementById("authStatus");
  const nonce=new URL(location.href).searchParams.get("nonce");if(nonce)sessionStorage.setItem("rbd-auth-nonce",nonce);
  const officeReady=window.Office?.onReady?Office.onReady():Promise.resolve();
  async function deliver(data){await officeReady;const value={...data,nonce:sessionStorage.getItem("rbd-auth-nonce")};if(window.Office?.context?.ui?.messageParent){Office.context.ui.messageParent(JSON.stringify(value),{targetOrigin:location.origin});sessionStorage.removeItem("rbd-auth-nonce");}else{text.textContent=data.error||"Вхід завершено. Поверніться до РБД.";}}
  try{
    const app=new PublicClientApplication(configuration("sessionStorage"));await app.initialize();const result=await app.handleRedirectPromise();
    if(result){await deliver({accessToken:result.accessToken,expires:result.expiresOn?.getTime()});await app.clearCache();return;}
    // Popup redirects are handled by MSAL in the opener; only an Office dialog initiates redirect.
    if(!sessionStorage.getItem("rbd-auth-nonce"))return;
    text.textContent="Вхід через корпоративний обліковий запис Microsoft…";
    await app.acquireTokenRedirect({scopes:[settings().scope]});
  }catch(e){text.textContent="Помилка входу: "+e.message;await deliver({error:e.message});}
}
