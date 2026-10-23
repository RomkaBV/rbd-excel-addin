/* RBD Google pilot controls and right panel collapse. Does not alter Excel data. */
(function(){
"use strict";
function init(){
 const modal=document.getElementById("requestModal"),btn=document.getElementById("rbdToggleRightPanel");
 if(false&&modal&&btn){
  btn.addEventListener("click",()=>{
   const collapsed=modal.classList.toggle("rbd-right-collapsed");
   btn.setAttribute("aria-expanded",String(!collapsed));
   btn.title=collapsed?"Розгорнути листування":"Згорнути листування";
   btn.setAttribute("aria-label",btn.title);
  });
 }
 const left=document.querySelector("#requestModal .history-main-column");
 const heading=document.querySelector("#requestModal .history-events-heading");
 if(false&&modal&&left&&heading&&!document.getElementById("rbdToggleLeftPanel")){
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
 const compose=document.getElementById("rbdComposeToggle");
 if(modal&&compose){
  compose.addEventListener("click",()=>{
   const opened=modal.classList.toggle("rbd-compose-open");
   modal.classList.remove("rbd-reminder-open");
   compose.textContent=opened?"✕ Закрити редактор":"✉ Написати в Outlook";
   compose.setAttribute("aria-expanded",String(opened));
  });
 }
 for(const id of ["reminderToggleButton","messageToggleButton"]){
  const b=document.getElementById(id);
  if(b&&modal)b.addEventListener("click",()=>{
    modal.classList.toggle("rbd-reminder-open",id==="reminderToggleButton");
    modal.classList.toggle("rbd-compose-open",id==="messageToggleButton");
    if(compose){compose.textContent="✕ Закрити редактор";compose.setAttribute("aria-expanded","true");}
  });
 }
 // Fixed layout: move existing action controls instead of duplicating event handlers.
 if(modal){
  modal.classList.remove("rbd-left-collapsed","rbd-right-collapsed","rbd-compose-open","rbd-reminder-open");
  const noteSlot=document.getElementById("rbdNoteActionSlot");
  const reminderSlot=document.getElementById("rbdReminderSlot");
  const noteBtn=document.getElementById("chatSendButton");
  const reminderBtn=document.getElementById("reminderToggleButton");
  if(noteSlot&&noteBtn)noteSlot.appendChild(noteBtn);
  if(reminderSlot&&reminderBtn)reminderSlot.appendChild(reminderBtn);
  const oldMail=document.getElementById("messageToggleButton");
  if(oldMail)oldMail.hidden=true;
  if(compose){
   compose.addEventListener("click",()=>{
    modal.classList.remove("rbd-reminder-open");
   });
  }
 }
 const title=document.querySelector("#requestModal .history-side .drawer-tool-title");
 if(!title||!title.parentNode)return;
 const row=document.createElement("div");
 row.style.cssText="padding:7px 0;font-size:11px;color:#426781";
 const cfg=window.RBD_GOOGLE_CONFIG||{};
 row.textContent="Файли з Drive відкриваються через Google. Перегляд без Google-входу ще не підключений.";
 title.parentNode.insertBefore(row,title.nextSibling);
 // Local staging only. Never claim files were uploaded before a protected API exists.
 const drop=document.getElementById("rbdMailDrop");
 const input=document.getElementById("rbdMailFileInput");
 const list=document.getElementById("rbdMailFileList");
 const status=document.getElementById("rbdMailUploadStatus");
 if(drop&&input&&list&&status){
  const staged=[];
  const validName=name=>/\.(eml|msg)$/i.test(name);
  let currentPreview=null;
  const reader=document.getElementById("rbdInlineMailReader");
  const closeReader=document.getElementById("rbdInlineMailClose");
  if(closeReader)closeReader.addEventListener("click",()=>{if(reader)reader.hidden=true;});
  const downloadReader=document.getElementById("rbdInlineMailDownload");
  if(downloadReader)downloadReader.addEventListener("click",()=>{
    if(!currentPreview)return;
    const url=URL.createObjectURL(currentPreview),a=document.createElement("a");a.href=url;a.download=currentPreview.name;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);
  });
  async function previewInline(file){
    if(!reader)return window.RBD_MAIL_VIEWER.open(file);
    currentPreview=file;
    document.getElementById("rbdInlineMailTitle").textContent="Перегляд листа — "+file.name;
    reader.hidden=false;
    const meta=document.getElementById("rbdInlineMailMeta"),body=document.getElementById("rbdInlineMailBody");
    if(/\\.msg$/i.test(file.name)){
      meta.textContent="Формат Outlook MSG";
      body.textContent="Оригінал можна завантажити й відкрити в Outlook. Для перегляду змісту натисніть «Відкрити» у локальному переглядачі.";
      try{await window.RBD_MAIL_VIEWER.open(file)}catch(e){body.textContent="Не вдалося прочитати MSG: "+String(e.message||e)}
      return;
    }
    const raw=await file.text();
    const idx=raw.search(/\\r?\\n\\r?\\n/),head=idx>=0?raw.slice(0,idx):raw;
    const get=(name)=>{const m=new RegExp("(?:^|\\\\r?\\\\n)"+name+":\\\\s*([^\\\\r\\\\n]*)","i").exec(head);return m?m[1]:"—"};
    meta.textContent="Від: "+get("From")+"\nКому: "+get("To")+"\nТема: "+get("Subject")+"\nДата: "+get("Date");
    body.textContent=(idx>=0?raw.slice(idx).replace(/^\\r?\\n\\r?\\n/,""):raw).slice(0,80000);
  }
  let exportBtn;
  function paint(){
   list.replaceChildren();
   if(exportBtn)exportBtn.hidden=staged.length===0;
   for(const file of staged){
    const item=document.createElement("div");
    item.className="rbd-local-letter-card";
    const icon=document.createElement("span");
    icon.className="rbd-local-letter-icon";icon.textContent="✉";
    const info=document.createElement("div");info.className="rbd-local-letter-info";
    const name=document.createElement("strong");name.textContent=file.name;
    const details=document.createElement("small");
    details.textContent=Math.ceil(file.size/1024)+" КБ · локальний файл";
    info.append(name,details);
    const actions=document.createElement("div");actions.className="rbd-local-letter-actions";
    const preview=document.createElement("button");
    preview.type="button";preview.textContent="Відкрити";preview.className="btn";
    preview.addEventListener("click",async()=>{
      try{await previewInline(file);}
      catch(e){status.textContent="Помилка перегляду: "+(e.message||e);}
    });
    const download=document.createElement("button");download.type="button";
    download.textContent="↓ Завантажити";download.className="btn rbd-download-letter";
    download.addEventListener("click",()=>{
      const url=URL.createObjectURL(file),a=document.createElement("a");
      a.href=url;a.download=file.name;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);
    });
    actions.append(preview,download);item.append(icon,info,actions);list.appendChild(item);
   }
   status.textContent=staged.length
      ? "Вибрано "+staged.length+" файл(ів). Лише локальний перегляд: у Google Drive файли ще НЕ завантажено."
      : "Виберіть .msg або .eml. Збереження в Drive поки через ручний імпорт.";
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
  exportBtn=document.createElement("button");
  exportBtn.type="button";
  exportBtn.className="btn rbd-manual-export";
  exportBtn.textContent="↓ Зберегти через Google Drive";
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