/* One persistent inline reader. Original bytes never leave the page during parsing. */
(function(){
"use strict";
const el=id=>document.getElementById(id),safe=s=>String(s||"").replace(/[<>:"/\\|?*\x00-\x1f]/g,"_").slice(0,180)||"лист.eml";
let selection=0,current=null,tab="message",worker=null,rejectWorker=null;
function cancel(){if(worker){worker.terminate();worker=null;rejectWorker?.(Error("Перегляд скасовано"));rejectWorker=null;}}
async function parse(file){
  const buffer=await file.arrayBuffer();
  return new Promise((resolve,reject)=>{
    const instance=new Worker("./rbd-mail-worker.js?v=3.4.9"),timer=setTimeout(()=>finish(Error("Перевищено час читання листа. Завантажте оригінал для Outlook.")),15000);
    function finish(error,info){clearTimeout(timer);instance.terminate();if(worker===instance){worker=null;rejectWorker=null;}error?reject(error):resolve(info);}
    worker=instance;rejectWorker=error=>finish(error);
    instance.onmessage=e=>finish(e.data.error?Error(e.data.error):null,e.data.info);
    instance.onerror=()=>finish(Error("Не вдалося запустити локальний парсер листів"));
    instance.postMessage({buffer,name:file.name},[buffer]);
  });
}
function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=safe(name);document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
function reset(message="Виберіть лист зі списку для перегляду."){
  selection++;cancel();current=null;tab="message";
  el("rbdInlineMailTitle").textContent="Перегляд листа";
  el("rbdInlineMailMeta").hidden=true;el("rbdReaderActions").hidden=true;el("rbdReaderTabs").hidden=true;
  el("rbdInlineMailBody").textContent=message;
}
function render(){
  if(!current)return;
  const body=el("rbdInlineMailBody"),info=current.info;body.replaceChildren();
  el("rbdReaderMessage").setAttribute("aria-selected",String(tab==="message"));
  el("rbdReaderAttachments").setAttribute("aria-selected",String(tab==="attachments"));
  if(tab==="attachments"){
    if(!info.attachments.length){body.textContent="У листі немає вкладень.";return;}
    for(const a of info.attachments){const row=document.createElement("div"),name=document.createElement("span"),button=document.createElement("button");row.className="rbd-attachment";name.textContent=a.name+(a.error?" — "+a.error:"");button.type="button";button.className="btn";button.textContent="Завантажити";button.disabled=!!a.error;button.onclick=()=>download(new Blob([a.content],{type:a.type||"application/octet-stream"}),a.name);row.append(name,button);body.append(row);}
  }else if(info.html){
    const frame=document.createElement("iframe");frame.className="rbd-mail-frame";frame.title="Текст листа";frame.setAttribute("sandbox","");frame.referrerPolicy="no-referrer";
    frame.srcdoc='<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src \'none\'; base-uri \'none\'; form-action \'none\'"><style>body{font:13px/1.5 Segoe UI,Arial,sans-serif;color:#173652;margin:12px;overflow-wrap:anywhere}table{max-width:100%;border-collapse:collapse}pre{white-space:pre-wrap}a{color:#146aca}</style></head><body>'+window.RBD_MAIL_PARSERS.sanitize(info.html)+'</body></html>';body.append(frame);
  }else body.textContent=info.text||"Текстовий вміст відсутній або лист використовує захищений / RTF-формат. Завантажте оригінал для Outlook.";
}
async function open(file){
  reset("Читання листа…");const id=selection;
  current={file,info:{attachments:[]}};
  el("rbdInlineMailTitle").textContent="Перегляд листа — "+file.name;
  el("rbdReaderActions").hidden=false;
  try{
    const info=await parse(file);if(id!==selection)return null;current={file,info};
    const meta=el("rbdInlineMailMeta");meta.replaceChildren();
    const date=info.date&&Number.isFinite(Date.parse(info.date))?new Intl.DateTimeFormat("uk-UA",{dateStyle:"short",timeStyle:"short",timeZone:window.RBD_CONFIG?.workCalendar?.timeZone||"Europe/Kyiv"}).format(new Date(info.date)):info.date;
    for(const [label,value] of [["Від:",info.from],["Кому:",info.to],["Тема:",info.subject],["Дата:",date]]){const dt=document.createElement("dt"),dd=document.createElement("dd");dt.textContent=label;dd.textContent=value||"—";meta.append(dt,dd);}
    meta.hidden=false;el("rbdReaderTabs").hidden=false;el("rbdReaderAttachments").textContent=`Вкладення (${info.attachments.length})`;render();return info;
  }catch(e){if(id===selection)el("rbdInlineMailBody").textContent="Не вдалося прочитати лист: "+e.message+" Оригінал доступний для завантаження.";return null;}
}
function init(){
  el("rbdInlineMailClose")?.addEventListener("click",()=>reset());
  el("rbdInlineMailDownload")?.addEventListener("click",()=>{if(current)download(current.file,current.file.name);});
  el("rbdInlineMailOutlook")?.addEventListener("click",()=>{if(!current)return;download(current.file,current.file.name);const status=el("rbdMailUploadStatus");status.classList.remove("error");status.textContent="Оригінал завантажено. Відкрийте його із завантажень у Outlook. Браузер не може гарантовано запустити локальний Outlook.";});
  el("rbdReaderMessage")?.addEventListener("click",()=>{tab="message";render();});
  el("rbdReaderAttachments")?.addEventListener("click",()=>{tab="attachments";render();});
}
window.RBD_MAIL_VIEWER=Object.freeze({open,reset,download});
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init);else init();
})();
