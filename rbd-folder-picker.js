/* RBD: user-selected Microsoft 365 folder. No credentials stored in source. */
(function () {
"use strict";
const GRAPH = "https://graph.microsoft.com/v1.0";
const KEY = "rbd.selectedFolder.v1";
let clientPromise;
let session = null;
const esc = s => String(s || "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const cfg = () => window.RBD_SHAREPOINT_CONFIG || {};
function message(s, error) {
  const el = document.getElementById("rbd-picker-msg");
  if (el) { el.textContent = s; el.style.color = error ? "#b3261e" : "#255b7d"; }
}
async function token() {
  const c = cfg();
  if (!c.clientId || !c.tenantId) throw Error("Для першого входу ІТ має налаштувати Client ID та Tenant ID в rbd-sharepoint-config.js. Пароль не потрібен.");
  if (!window.Office?.context?.requirements?.isSetSupported?.("NestedAppAuth", "1.1"))
    throw Error("Цей клієнт Excel не підтримує Microsoft Nested App Authentication. Потрібен сумісний Excel.");
  if (!clientPromise) clientPromise = import("https://cdn.jsdelivr.net/npm/@azure/msal-browser@4/+esm").then(m =>
    m.createNestablePublicClientApplication({auth:{clientId:c.clientId,authority:"https://login.microsoftonline.com/"+encodeURIComponent(c.tenantId)}}));
  const client = await clientPromise;
  const acc = client.getActiveAccount() || client.getAllAccounts()[0];
  const request = {scopes:["Files.ReadWrite","Sites.Read.All"],...(acc?{account:acc}:{})};
  try { const result = await client.acquireTokenSilent(request); return result.accessToken; }
  catch (_) { return (await client.acquireTokenPopup(request)).accessToken; }
}
async function api(path, opt={}) {
  if (!session) session = await token();
  const req = async bearer => fetch(GRAPH + path, {method:opt.method||"GET",headers:{Authorization:"Bearer "+bearer,...(opt.headers||{})},body:opt.body});
  let resp = await req(session);
  if (resp.status === 401) { session = await token(); resp = await req(session); }
  if (!resp.ok) { let b={}; try {b=await resp.json();} catch (_) {} throw Error(b.error?.message || "Microsoft Graph: "+resp.status); }
  return resp.status===204 ? {} : resp.json();
}
function encodeShareUrl(url) {
  const bytes = new TextEncoder().encode(url);
  let str = ""; for (const b of bytes) str += String.fromCharCode(b);
  return "u!" + btoa(str).replace(/=/g,"").replace(/\+/g,"-").replace(/\//g,"_");
}
function saveSelected(folder) {
  const saved = {driveId:folder.parentReference?.driveId || folder.driveId,id:folder.id,name:folder.name,webUrl:folder.webUrl};
  if (!saved.driveId || !saved.id || !saved.folder) { /* folder verification is done before calling */ }
  localStorage.setItem(KEY,JSON.stringify(saved));
  message("Обрано: "+saved.name+" — збережено для цього браузера Excel.");
  showChosen();
}
function chosen() { try {return JSON.parse(localStorage.getItem(KEY)||"null");} catch (_){return null;} }
function showChosen() {
  const c=chosen(), el=document.getElementById("rbd-picker-selected");
  if (!el) return;
  el.textContent = c ? "Поточна папка: "+c.name+" ("+c.webUrl+")" : "Папку ще не обрано";
}
async function view(folder) {
  if (!folder.folder) throw Error("Це не папка");
  const driveId = folder.parentReference?.driveId || folder.driveId;
  if (!driveId) throw Error("Microsoft Graph не повернув driveId для папки");
  const list=await api("/drives/"+encodeURIComponent(driveId)+"/items/"+encodeURIComponent(folder.id)+"/children?$select=id,name,folder,webUrl,parentReference&$top=200");
  const container = document.getElementById("rbd-picker-folders");
  const dirs=(list.value||[]).filter(x=>x.folder);
  container.replaceChildren();
  document.getElementById("rbd-picker-path").textContent=folder.webUrl||folder.name;
  const use=document.createElement("button"); use.textContent="✓ Обрати цю папку"; use.className="rbdp-primary";
  use.onclick=()=>{saveSelected({...folder,driveId});message("Папку вибрано. Тепер можна завантажити тестовий лист.");};
  container.appendChild(use);
  for (const dir of dirs) {
    const b=document.createElement("button"); b.textContent="📁 "+dir.name; b.className="rbdp-folder";
    b.onclick=()=>run(()=>view({...dir,driveId}));container.appendChild(b);
  }
  if (list["@odata.nextLink"]) {
    const note=document.createElement("div");note.textContent="Показано перші 200 папок. Для великої бібліотеки вибери ближчу папку."; container.appendChild(note);
  }
}
async function myDrive() {
  const root=await api("/me/drive/root?$select=id,name,folder,webUrl,parentReference");
  const me=await api("/me/drive?$select=id");
  await view({...root,driveId:me.id});
}
async function sharedLink() {
  const url=document.getElementById("rbd-picker-link").value.trim();
  if (!/^https:\/\//i.test(url)) throw Error("Вкажи повне HTTPS-посилання на спільну папку.");
  const item=await api("/shares/"+encodeShareUrl(url)+"/driveItem?$select=id,name,folder,webUrl,parentReference");
  await view(item);
}
function safeRequestId(v) {
  v=String(v||"").trim();
  if(!/^[a-zA-ZА-ЯІЇЄҐа-яіїєґ0-9_-]{3,80}$/.test(v))throw Error("Введи коректний номер тестової заявки.");
  return v;
}
async function uploadTest() {
  const folder=chosen(); if (!folder) throw Error("Спочатку обери папку.");
  const id=safeRequestId(document.getElementById("rbd-picker-request").value);
  // Non-destructive: create one uniquely named .eml in a dedicated test folder.
  const path="/drives/"+encodeURIComponent(folder.driveId)+"/items/"+encodeURIComponent(folder.id)+"/children";
  const children=await api(path+"?$select=id,name,folder&$top=200");
  let target=(children.value||[]).find(x=>x.name==="RBD_TEST"&&x.folder);
  if(!target) target=await api(path,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name:"RBD_TEST",folder:{}, "@microsoft.graph.conflictBehavior":"fail"})});
  const name=id+"_"+Date.now()+".eml";
  const mail="From: test@rbd.invalid\r\nTo: test@rbd.invalid\r\nSubject: RBD SharePoint connectivity test\r\nDate: "+new Date().toUTCString()+"\r\nContent-Type: text/plain; charset=UTF-8\r\n\r\nTest message generated by RBD folder picker. Not an actual customer email.\r\n";
  const item=await api("/drives/"+encodeURIComponent(folder.driveId)+"/items/"+encodeURIComponent(target.id)+":/"+encodeURIComponent(name)+":/content",{method:"PUT",headers:{"Content-Type":"message/rfc822"},body:new Blob([mail],{type:"message/rfc822"})});
  message("Тестовий лист завантажено: "+item.webUrl);
  const link=document.getElementById("rbd-picker-result");link.textContent="Відкрити тестовий лист";link.href=item.webUrl;link.hidden=false;
}
async function run(fn){try{message("Виконується запит до Microsoft 365…");await fn();}catch(e){message(e.message||String(e),true);}}
function init() {
  const launch=document.createElement("button"); launch.id="rbd-picker-launch";launch.type="button";launch.textContent="☁ Підключити SharePoint";
  Object.assign(launch.style,{position:"fixed",right:"16px",bottom:"16px",zIndex:"2500",background:"#0b72d9",color:"#fff",border:"none",padding:"11px 16px",borderRadius:"9px",fontWeight:"bold",boxShadow:"0 4px 18px #0003"});
  document.body.appendChild(launch);
  const wrap=document.createElement("div");wrap.id="rbd-folder-picker";
  wrap.innerHTML=`<style>
    #rbd-folder-picker{display:none;position:fixed;inset:0;z-index:2600;background:#092b4566;align-items:center;justify-content:center;font:13px Segoe UI,Arial,sans-serif}
    #rbd-folder-picker .rbdp-panel{width:620px;max-width:95vw;max-height:90vh;overflow:auto;padding:23px;background:white;border-radius:12px;color:#23445d;box-shadow:0 20px 60px #0004}
    #rbd-folder-picker button{margin:5px;padding:9px 12px;border:1px solid #ccdbe7;background:white;border-radius:7px;cursor:pointer;color:#16446e}
    #rbd-folder-picker .rbdp-primary{background:#0b72d9;color:white;border-color:#0b72d9}
    #rbd-folder-picker .rbdp-folder{display:block;width:100%;text-align:left}
    #rbd-folder-picker input{padding:9px;border:1px solid #ccdbe7;border-radius:7px;width:100%;margin:6px 0}
    #rbd-picker-folders{max-height:240px;overflow:auto;border:1px solid #e0eaf2;border-radius:8px;padding:8px;margin:10px 0}
    #rbd-picker-msg{min-height:28px;word-break:break-word}
  </style><div class="rbdp-panel">
  <div style="display:flex;justify-content:space-between;align-items:center"><h2>☁ SharePoint / OneDrive — папка РБД</h2><button id="rbd-picker-close">✕</button></div>
  <p>Увійди через корпоративний Microsoft 365. Обери свою папку або відкрий спільну за посиланням.</p>
  <button id="rbd-picker-own" class="rbdp-primary">Увійти та відкрити мій OneDrive</button>
  <label>Спільна папка SharePoint / OneDrive</label>
  <input id="rbd-picker-link" type="url" placeholder="https://...sharepoint.com/...">
  <button id="rbd-picker-open-link">Відкрити папку за посиланням</button>
  <div id="rbd-picker-path"></div><div id="rbd-picker-folders"></div>
  <div id="rbd-picker-selected"></div><hr>
  <strong>Тест запису (тільки після вибору папки)</strong>
  <p>Створить папку RBD_TEST і один тестовий .eml, не зачіпаючи робочі заявки.</p>
  <input id="rbd-picker-request" value="RBD_TEST_001" aria-label="Номер тестової заявки">
  <button id="rbd-picker-upload">Завантажити тестовий лист</button>
  <a id="rbd-picker-result" hidden target="_blank" rel="noopener noreferrer">Відкрити тестовий лист</a>
  <p id="rbd-picker-msg" role="status"></p>
  </div>`;
  document.body.appendChild(wrap);
  launch.onclick=()=>{wrap.style.display="flex";showChosen();};
  document.getElementById("rbd-picker-close").onclick=()=>wrap.style.display="none";
  document.getElementById("rbd-picker-own").onclick=()=>run(myDrive);
  document.getElementById("rbd-picker-open-link").onclick=()=>run(sharedLink);
  document.getElementById("rbd-picker-upload").onclick=()=>run(uploadTest);
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init);else init();
})();