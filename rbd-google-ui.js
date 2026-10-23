/* RBD Google pilot controls and right panel collapse. Does not alter Excel data. */
(function(){
"use strict";
function init(){
 const modal=document.getElementById("requestModal"),btn=document.getElementById("rbdToggleRightPanel");
 if(modal&&btn){
  btn.addEventListener("click",()=>{
   const collapsed=modal.classList.toggle("rbd-right-collapsed");
   btn.setAttribute("aria-expanded",String(!collapsed));
   btn.title=collapsed?"Розгорнути листування":"Згорнути листування";
   btn.setAttribute("aria-label",btn.title);
  });
 }
 const left=document.querySelector("#requestModal .history-main-column");
 const heading=document.querySelector("#requestModal .history-events-heading");
 if(modal&&left&&heading&&!document.getElementById("rbdToggleLeftPanel")){
  const toggle=document.createElement("button");
  toggle.id="rbdToggleLeftPanel";toggle.type="button";toggle.textContent="◀";
  toggle.title="Згорнути історію зліва";toggle.setAttribute("aria-label",toggle.title);toggle.setAttribute("aria-expanded","true");
  heading.appendChild(toggle);
  const rail=document.createElement("div");rail.id="rbdLeftCollapsedRail";
  const reopen=document.createElement("button");reopen.type="button";reopen.textContent="▶";reopen.title="Розгорнути історію";
  rail.appendChild(reopen);left.insertBefore(rail,left.firstChild);
  const run=()=>{
   const state=modal.classList.toggle("rbd-left-collapsed");
   if(state)modal.classList.remove("rbd-right-collapsed");
   toggle.setAttribute("aria-expanded",String(!state));
   toggle.textContent=state?"▶":"◀";
   toggle.title=state?"Розгорнути історію":"Згорнути історію зліва";
   const right=document.getElementById("rbdToggleRightPanel");
   if(right)right.setAttribute("aria-expanded","true");
  };
  toggle.addEventListener("click",run);reopen.addEventListener("click",run);
 }
 const title=document.querySelector("#requestModal .history-side .drawer-tool-title");
 if(!title||!title.parentNode)return;
 const row=document.createElement("div");
 row.style.cssText="padding:7px 0;font-size:11px;color:#426781";
 const cfg=window.RBD_GOOGLE_CONFIG||{};
 row.textContent=cfg.webAppUrl?"Google тест: адресу бекенду задано (доступ ще потребує перевірки).":"Google Drive / Gmail: тестовий режим — Apps Script ще не підключено. Файли залишаються локально, поки не налаштований сервер.";
 title.parentNode.insertBefore(row,title.nextSibling);
 // Local staging only. Never claim files were uploaded before a protected API exists.
 const drop=document.getElementById("rbdMailDrop");
 const input=document.getElementById("rbdMailFileInput");
 const list=document.getElementById("rbdMailFileList");
 const status=document.getElementById("rbdMailUploadStatus");
 if(drop&&input&&list&&status){
  const staged=[];
  const validName=name=>/\.(eml|msg)$/i.test(name);
  function paint(){
   list.replaceChildren();
   for(const file of staged){
    const item=document.createElement("div");
    item.style.cssText="display:flex;gap:7px;align-items:center;padding:5px 0";
    const label=document.createElement("span");
    label.textContent=file.name+" ("+Math.ceil(file.size/1024)+" КБ)";
    label.style.flex="1";
    item.appendChild(label);
    const preview=document.createElement("button");
    preview.type="button";preview.textContent="Відкрити";
    preview.className="btn";
    preview.addEventListener("click",async()=>{
      try{await window.RBD_MAIL_VIEWER.open(file);}
      catch(e){status.textContent="Перегляд: "+(e.message||e);}
    });
    item.appendChild(preview);
    list.appendChild(item);
   }
   status.textContent=staged.length
      ? "Вибрано "+staged.length+" файл(ів). Лише локальний перегляд: у Google Drive файли ще НЕ завантажено."
      : "Завантаження до Google Drive ще не підключене.";
  }
  function accept(files){
    for(const file of files){
      if(!validName(file.name)){status.textContent="Підтримуються тільки файли .eml та .msg";continue;}
      if(!file.size||file.size>4*1024*1024){status.textContent="Файл порожній або перевищує 4 МБ";continue;}
      staged.push(file);
    }
    paint();
  }
  drop.addEventListener("click",event=>{if(event.target!==input)input.click();});
  drop.addEventListener("keydown",event=>{if(event.key==="Enter"||event.key===" "){event.preventDefault();input.click();}});
  input.addEventListener("change",()=>{accept(input.files||[]);input.value="";});
  drop.addEventListener("dragover",event=>{event.preventDefault();drop.classList.add("dragover");});
  drop.addEventListener("dragleave",()=>drop.classList.remove("dragover"));
  drop.addEventListener("drop",event=>{event.preventDefault();drop.classList.remove("dragover");accept(event.dataTransfer?.files||[]);});
  const exportBtn=document.createElement("button");
  exportBtn.type="button";
  exportBtn.className="btn";
  exportBtn.textContent="↓ Підготувати листи для Google Drive";
  exportBtn.style.cssText="margin-top:7px";
  exportBtn.addEventListener("click",()=>{
    const requestId=(typeof STATE!=="undefined" && STATE.selectedRequest && STATE.selectedRequest.id)||"";
    if(!/^[A-Za-z0-9_-]{3,80}$/.test(requestId)){
      status.textContent="Відкрийте конкретну заявку перед підготовкою файлів";return;
    }
    if(!staged.length){status.textContent="Спочатку виберіть файл .eml або .msg";return;}
    staged.forEach((file,index)=>{
      const safe=file.name.replace(/[^a-zA-Z0-9._-]/g,"_").slice(0,100);
      const url=URL.createObjectURL(file);
      const a=document.createElement("a");
      a.href=url;
      a.download=requestId+"__"+String(Date.now())+"_"+index+"__"+safe;
      document.body.appendChild(a);a.click();a.remove();
      setTimeout(()=>URL.revokeObjectURL(url),60000);
    });
    status.textContent="Підготовлено "+staged.length+" файл(ів) для заявки "+requestId+". Завантажте їх у Google Drive → RBD → Імпорт листування, потім запустіть importRbdLetters. Це ручний імпорт.";
  });
  list.insertAdjacentElement("afterend",exportBtn);
  paint();
 }

}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init);else init();
})();