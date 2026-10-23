/* Request-scoped correspondence controller. No layout mutation or demo records. */
(function(){
"use strict";
const el=id=>document.getElementById(id),states=new Map();
let active="",generation=0,openGeneration=0;
const escape=s=>String(s??"").replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const bridge=()=>window.RBD_GOOGLE_BRIDGE;
function state(id){if(!states.has(id))states.set(id,{local:[],remote:[],registry:[],selected:"",loaded:false,loading:false,error:"",notice:"",errorCode:""});return states.get(id);}
function status(text,error=false){const out=el("rbdMailUploadStatus");out.textContent=text;out.classList.toggle("error",error);}
function letters(s){
  const list=new Map();
  for(const item of [...s.registry,...s.remote,...s.local])list.set(item.id,item);
  return [...list.values()].sort((a,b)=>(Date.parse(b.date)||0)-(Date.parse(a.date)||0));
}
function formatDate(value){const date=new Date(value);return Number.isFinite(date.getTime())?new Intl.DateTimeFormat("uk-UA",{dateStyle:"short",timeStyle:"short",timeZone:window.RBD_CONFIG?.workCalendar?.timeZone||"Europe/Kyiv"}).format(date):"Дата не вказана";}
function render(){
  if(!active)return;const s=state(active),items=letters(s);el("historySideCount").textContent=`(${items.length})`;
  el("messageThreadCards").innerHTML=items.length?items.map(item=>`<article class="rbd-letter-row ${s.selected===item.id?"selected":""}" data-letter="${escape(item.id)}"><div class="rbd-letter-icon">${/\.msg$/i.test(item.name)?"MSG":"EML"}</div><div class="rbd-letter-content"><div class="rbd-letter-name">${escape(item.name)}</div><div class="rbd-letter-info">Від: ${escape(item.info?.from||item.from||"—")}</div><div class="rbd-letter-info">Тема: ${escape(item.info?.subject||item.subject||"Відкрийте лист для читання")}</div><div class="rbd-letter-info">${escape(formatDate(item.info?.date||item.date))} · ${item.size?Math.ceil(item.size/1024)+" КБ":"Розмір невідомий"}</div>${item.file?'<div class="rbd-letter-source">'+escape(item.uploading?"Зберігаємо у Drive…":item.uploadError?"Не збережено: "+item.uploadError:"Локальний файл · ще не збережений у Drive")+'</div>':""}</div><div class="rbd-letter-actions"><button type="button" class="btn primary" data-mail-action="open">Відкрити</button><button type="button" class="btn" data-mail-action="download">↓ Завантажити</button>${item.file&&bridge()?.ready&&!item.uploading?'<button type="button" class="btn" data-mail-action="upload">У Drive</button>':""}</div></article>`).join(""):'<div class="history-side-empty">'+(s.loading?"Завантаження листів…":s.error?"Список листів не завантажено.":"Для цієї заявки листів поки немає.")+'</div>';
  status(s.error||s.notice||(s.loading?"Завантаження приватних листів…":!bridge()?.ready?"Архів Drive ще не підключено. Додані файли доступні лише до перезавантаження сторінки.":""),!!s.error);
  const connect=el("rbdConnectLetters");connect.hidden=!bridge()?.ready||!(["AUTH_REQUIRED","AUTH_EXPIRED"].includes(s.errorCode));
}
async function refresh(interactive=false){
  if(!active)return;const id=active,s=state(id),ticket=++generation;
  if(!bridge()?.ready){s.loading=false;s.loaded=true;render();return;}
  s.loading=true;s.error="";render();
  try{
    if(interactive)await bridge().signIn();
    const result=await bridge().listLetters(id);
    if(ticket!==generation||active!==id)return;
    if(!Array.isArray(result.letters))throw Error("Некоректна відповідь сервера листування");
    s.remote=result.letters.map(x=>({...x,id:String(x.id),source:"drive"}));
    // The authorized server is authoritative; stale Excel registry entries must not outlive it.
    s.registry=[];s.loaded=true;s.error="";s.errorCode="";
    if(s.selected&&!letters(s).some(x=>x.id===s.selected)){s.selected="";window.RBD_MAIL_VIEWER.reset();}
  }catch(e){
    if(ticket!==generation||active!==id)return;
    s.error=e.message;s.errorCode=e.code||"";s.loaded=false;
    if(e.status===401||e.status===403){s.remote=[];s.registry=[];s.selected="";openGeneration++;window.RBD_MAIL_VIEWER.reset("Немає підтвердженого доступу до архіву цієї заявки.");}
  }finally{if(ticket===generation&&active===id){s.loading=false;render();}}
}
function show(requestId,data={}){
  const changed=active!==requestId;if(changed){generation++;openGeneration++;active=requestId;window.RBD_MAIL_VIEWER.reset();}
  const s=state(requestId);
  if(!bridge()?.ready)s.registry=(data.registry||[]).filter(x=>/\.(msg|eml)$/i.test(x["Назва файла"]||"")).map(x=>({id:String(x["ID файла"]||""),name:String(x["Назва файла"]),size:Number(x["Розмір, байт"])||0,date:x["Дата створення"],source:"registry"})).filter(x=>/^[A-Za-z0-9_-]{10,200}$/.test(x.id));
  render();
  if(changed||(!s.loaded&&!s.loading&&!s.error))refresh();
}
async function getFile(item,id){
  if(item.file)return item.file;
  if(!bridge()?.ready)throw Error("Для читання цього приватного файла необхідно підключити корпоративний архів РБД. Завантажений з Drive оригінал можна додати в поле вище.");
  const blob=await bridge().getLetter(id,item.id);return new File([blob],item.name,{type:blob.type});
}
async function act(action,item){
  const id=active,s=state(id),ticket=++openGeneration;
  try{
    if(action==="upload")return upload(id,item);
    if(action==="open"){s.selected=item.id;window.RBD_MAIL_VIEWER.reset("Завантаження оригіналу…");render();}
    const file=await getFile(item,id);if(active!==id||ticket!==openGeneration)return;
    if(action==="download"){window.RBD_MAIL_VIEWER.download(file,item.name);return;}
    const info=await window.RBD_MAIL_VIEWER.open(file);if(active!==id||ticket!==openGeneration)return;
    if(info){item.info=info;render();}
  }catch(e){if(active!==id||ticket!==openGeneration)return;s.error=e.message;if(e.status===401||e.status===403){s.remote=[];s.registry=[];s.selected="";}window.RBD_MAIL_VIEWER.reset("Не вдалося завантажити лист. "+e.message);render();}
}
async function upload(id,item){
  const s=state(id);if(item.uploading)return;item.uploading=true;item.uploadError="";if(active===id)render();
  try{
    const result=await bridge().uploadLetter(id,item.file,item.uploadId);
    if(!result.letter?.id)throw Error("Сервер не підтвердив збереження файла");
    s.local=s.local.filter(x=>x!==item);s.remote=s.remote.filter(x=>x.id!==result.letter.id);s.remote.push({...result.letter,source:"drive",info:item.info});
    if(s.selected===item.id)s.selected=result.letter.id;s.notice="Лист збережено у приватному архіві Drive.";
  }catch(e){item.uploadError=e.message;s.error=e.message;s.errorCode=e.code||"";}
  finally{item.uploading=false;if(active===id)render();}
}
async function accept(files){
  if(!active)return;const id=active,s=state(id),rejected=[];let first=null;
  for(const file of Array.from(files)){
    if(!/\.(eml|msg)$/i.test(file.name)||!file.size||file.size>4*1024*1024){rejected.push(file.name+": потрібен .eml/.msg від 1 байта до 4 МБ");continue;}
    if(s.local.reduce((n,x)=>n+x.size,0)+file.size>32*1024*1024){rejected.push("Ліміт локальних файлів заявки — 32 МБ");continue;}
    if(s.local.some(x=>x.name===file.name&&x.size===file.size&&x.date===new Date(file.lastModified).toISOString()))continue;
    const item={id:"local-"+crypto.randomUUID(),uploadId:crypto.randomUUID(),file,name:file.name,size:file.size,date:new Date(file.lastModified).toISOString()};s.local.push(item);if(!first)first=item;
    if(bridge()?.ready)upload(id,item);
  }
  s.error=rejected.join("; ");if(active===id){render();if(first)act("open",first);}
}
function close(){generation++;openGeneration++;active="";window.RBD_MAIL_VIEWER.reset();}
function init(){
  const drop=el("rbdMailDrop"),input=el("rbdMailFileInput");
  drop.addEventListener("click",()=>input.click());drop.addEventListener("keydown",e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();input.click();}});
  input.addEventListener("change",()=>{accept(input.files||[]);input.value="";});
  drop.addEventListener("dragover",e=>{e.preventDefault();drop.classList.add("dragover");});drop.addEventListener("dragleave",()=>drop.classList.remove("dragover"));drop.addEventListener("drop",e=>{e.preventDefault();drop.classList.remove("dragover");accept(e.dataTransfer?.files||[]);});
  el("messageThreadCards").addEventListener("click",e=>{const b=e.target.closest("[data-mail-action]"),row=b?.closest("[data-letter]");if(!row||!active)return;const item=letters(state(active)).find(x=>x.id===row.dataset.letter);if(item)act(b.dataset.mailAction,item);});
  el("rbdMailMenuToggle").addEventListener("click",()=>{el("rbdMailMenu").hidden=!el("rbdMailMenu").hidden;el("rbdMailMenuToggle").setAttribute("aria-expanded",String(!el("rbdMailMenu").hidden));});
  document.addEventListener("click",e=>{if(!e.target.closest(".rbd-mail-actions")){el("rbdMailMenu").hidden=true;el("rbdMailMenuToggle").setAttribute("aria-expanded","false");}});
  el("rbdRefreshLetters").addEventListener("click",()=>refresh());el("rbdConnectLetters").addEventListener("click",()=>refresh(true));
}
window.RBD_CORRESPONDENCE=Object.freeze({show,refresh,close});
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init);else init();
})();
