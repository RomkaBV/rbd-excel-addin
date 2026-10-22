/* RBD SharePoint auto upload. No credentials or client secrets are stored here.
   Configured per tenant in rbd-sharepoint-config.js. */
(function(global){
"use strict";
const GRAPH="https://graph.microsoft.com/v1.0";
let clientPromise=null;
function config(){
  const cfg=global.RBD_SHAREPOINT_CONFIG||{};
  return {
    clientId:String(cfg.clientId||"").trim(),
    tenantId:String(cfg.tenantId||"").trim(),
    driveId:String(cfg.driveId||"").trim(),
    rootFolder:String(cfg.rootFolder||"RBD").trim()
  };
}
function configured(){
  const c=config();
  return !!(c.clientId&&c.tenantId&&c.driveId);
}
async function getToken(){
  const c=config();
  if(!configured())throw new Error("SharePoint ще не підключено: ІТ має вказати Client ID, Tenant ID та ID бібліотеки SharePoint у конфігурації RBD.");
  const supported=typeof Office!=="undefined"&&!!Office.context?.requirements?.isSetSupported?.("NestedAppAuth","1.1");
  if(!supported)throw new Error("Поточний клієнт Excel не підтримує Microsoft Nested App Authentication (NAA). Використайте сумісну версію Excel або зверніться до ІТ.");
  if(!clientPromise){
    clientPromise=(async()=>{
      const mod=await import("https://cdn.jsdelivr.net/npm/@azure/msal-browser@4/+esm");
      if(typeof mod.createNestablePublicClientApplication!=="function")throw new Error("Не вдалося завантажити авторизацію Microsoft.");
      return mod.createNestablePublicClientApplication({
        auth:{clientId:c.clientId,authority:"https://login.microsoftonline.com/"+encodeURIComponent(c.tenantId)},
        cache:{cacheLocation:"sessionStorage"}
      });
    })();
  }
  let client;
  try{client=await clientPromise;}catch(e){clientPromise=null;throw e;}
  let loginHint="";
  try{
    const ctx=await Office.auth?.getAuthContext?.();
    loginHint=ctx?.userPrincipalName||"";
  }catch(_){}
  const scopes=["Files.ReadWrite","Sites.Read.All"];
  let existing=client.getActiveAccount?.()||client.getAllAccounts?.()[0];
  const request={scopes,...(existing?{account:existing}:(loginHint?{loginHint}:{}))};
  try{
    const token=await client.acquireTokenSilent(request);
    if(token?.accessToken)return token.accessToken;
  }catch(e){
    // An interactive sign-in or delegated permission consent is needed.
  }
  const token=await client.acquireTokenPopup(request);
  if(!token?.accessToken)throw new Error("Microsoft не видав токен для SharePoint.");
  return token.accessToken;
}
async function graph(path,token,opts={}){
  const url=GRAPH+path;
  const headers={Authorization:"Bearer "+token,...(opts.headers||{})};
  let resp;
  try{resp=await fetch(url,{method:opts.method||"GET",headers,body:opts.body});}
  catch(e){throw new Error("Немає з'єднання з Microsoft Graph: "+(e?.message||e));}
  if(!resp.ok){
    let hint="";
    try{const b=await resp.json();hint=b?.error?.message||b?.error?.code||"";}catch(_){}
    const err=new Error("Помилка SharePoint ("+resp.status+"): "+(hint||resp.statusText));
    err.status=resp.status;
    throw err;
  }
  return resp.status===204?{}:resp.json();
}
function encodeName(s){
  return encodeURIComponent(String(s).replace(/[\\/:*?"<>|#%]/g,"_").replace(/[. ]+$/,"").slice(0,110)||"Файл");
}
async function ensureFolder(driveId,parentId,folder,token){
  const encoded=encodeName(folder);
  const endpoint="/drives/"+encodeURIComponent(driveId)+"/items/"+encodeURIComponent(parentId)+":/"+encoded;
  try{
    const found=await graph(endpoint,token);
    if(!found.folder)throw new Error("У SharePoint об'єкт «"+folder+"» уже існує, але це не папка.");
    return found;
  }catch(e){
    if(e.status!==404)throw e;
  }
  try{
    return await graph("/drives/"+encodeURIComponent(driveId)+"/items/"+encodeURIComponent(parentId)+"/children",token,{
      method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({name:String(folder).replace(/[\\/:*?"<>|]/g,"_").slice(0,110),folder:{},
      "@microsoft.graph.conflictBehavior":"fail"})
    });
  }catch(e){
    if(e.status===409)return graph(endpoint,token);
    throw e;
  }
}
function validateRequestId(requestId){
  const v=String(requestId||"").trim();
  if(!/^[a-zA-ZА-ЯІЇЄҐа-яіїєґ0-9_-]{3,80}$/.test(v))throw new Error("Некоректний номер заявки для папки SharePoint.");
  return v;
}
async function uploadFiles({requestId,messageId,files,onProgress}){
  const c=config();
  if(!configured())throw new Error("Автозавантаження не налаштоване: потрібні параметри Microsoft Graph від ІТ.");
  const id=validateRequestId(requestId);
  const mid=validateRequestId(messageId);
  const uploads=Array.from(files||[]);
  if(!uploads.length)return [];
  if(uploads.length>20)throw new Error("За один раз можна завантажити не більш ніж 20 файлів.");
  for(const f of uploads){
    if(!f || !f.name || typeof f.size!=="number")throw new Error("Не вдалося прочитати файл.");
    if(f.size>250*1024*1024)throw new Error("Файл «"+f.name+"» перевищує ліміт 250 МБ.");
    if(f.size===0)throw new Error("Файл «"+f.name+"» порожній.");
  }
  const token=await getToken();
  const root=await graph("/drives/"+encodeURIComponent(c.driveId)+"/root",token);
  let folder=root;
  for(const name of [c.rootFolder,id,"Листування",mid]){
    folder=await ensureFolder(c.driveId,folder.id,name,token);
  }
  const results=[];
  for(let i=0;i<uploads.length;i++){
    const f=uploads[i];
    const safeName=String(i+1).padStart(2,"0")+"_"+String(f.name).replace(/[\\/:*?"<>|#%]/g,"_").slice(0,120);
    if(onProgress)onProgress(i,uploads.length,f.name);
    const uploaded=await graph("/drives/"+encodeURIComponent(c.driveId)+"/items/"+encodeURIComponent(folder.id)+":/"+encodeName(safeName)+":/content",token,{
      method:"PUT",headers:{"Content-Type":f.type||"application/octet-stream"},
      body:f
    });
    if(!uploaded.webUrl)throw new Error("SharePoint не повернув посилання на файл "+f.name);
    results.push({originalName:f.name,name:uploaded.name||f.name,mime:f.type||"application/octet-stream",
      size:f.size,url:uploaded.webUrl,driveItemId:uploaded.id||"",folderUrl:folder.webUrl||""});
  }
  if(onProgress)onProgress(uploads.length,uploads.length,"");
  return results;
}
global.RBDSharePoint={configured,config,uploadFiles};
})(window);
