/* RBD local mail viewer: never uploads the selected mail or shares Google links. */
(function(){
"use strict";
function field(head,name){const match=new RegExp("(?:^|\\r?\\n)"+name+":\\s*([^\\r\\n]*)","i").exec(head);return match?match[1].trim():"—";}
function decodeBody(part){return part.replace(/\r\n/g,"\n").trim();}
function show(file,info){
 let overlay=document.getElementById("rbdMailPreviewOverlay");
 if(!overlay){
  overlay=document.createElement("div");overlay.id="rbdMailPreviewOverlay";
  overlay.style.cssText="position:fixed;inset:0;background:rgba(10,26,45,.72);z-index:99999;display:flex;align-items:center;justify-content:center;padding:24px";
  const panel=document.createElement("section");
  panel.style.cssText="background:#fff;color:#173652;width:min(850px,96vw);max-height:90vh;display:flex;flex-direction:column;border-radius:12px;overflow:hidden";
  panel.innerHTML='<header style="display:flex;align-items:center;gap:12px;padding:15px;border-bottom:1px solid #ddd"><strong id="rbdMailPreviewTitle" style="flex:1"></strong><button type="button" id="rbdMailPreviewSave" class="btn">Завантажити оригінал</button><button type="button" id="rbdMailPreviewClose" class="btn">Закрити</button></header><div id="rbdMailPreviewMeta" style="padding:15px;border-bottom:1px solid #ddd;white-space:pre-wrap"></div><pre id="rbdMailPreviewBody" style="padding:16px;overflow:auto;flex:1;white-space:pre-wrap;overflow-wrap:anywhere;font:13px/1.5 Segoe UI,Arial,sans-serif"></pre>';
  overlay.appendChild(panel);document.body.appendChild(overlay);
  overlay.querySelector("#rbdMailPreviewClose").onclick=()=>overlay.remove();
  overlay.addEventListener("click",e=>{if(e.target===overlay)overlay.remove();});
 }
 overlay.querySelector("#rbdMailPreviewTitle").textContent=file.name;
 overlay.querySelector("#rbdMailPreviewMeta").textContent="Від: "+(info.from||"—")+"\nКому: "+(info.to||"—")+"\nТема: "+(info.subject||"—")+"\nДата: "+(info.date||"—");
 overlay.querySelector("#rbdMailPreviewBody").textContent=info.body||"Текст листа відсутній або зберігається в іншому форматі.";
 overlay.querySelector("#rbdMailPreviewSave").onclick=()=>{
  const url=URL.createObjectURL(file),a=document.createElement("a");a.href=url;a.download=file.name;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);
 };
}
async function open(file){
 if(/\.eml$/i.test(file.name)){
  const raw=await file.text();
  const split=raw.search(/\r?\n\r?\n/),head=split>=0?raw.slice(0,split):raw;
  let body=split>=0?raw.slice(split).replace(/^\r?\n\r?\n/,""):"";
  const transfer=field(head,"Content-Transfer-Encoding").toLowerCase();
  if(transfer==="base64"){
   try{const bytes=Uint8Array.from(atob(body.replace(/\s/g,"")),c=>c.charCodeAt(0));body=new TextDecoder("utf-8").decode(bytes);}catch(_){}
  }else if(transfer==="quoted-printable"){
   body=body.replace(/=\r?\n/g,"").replace(/=([0-9A-F]{2})/gi,(_,hex)=>String.fromCharCode(parseInt(hex,16)));
  }
  show(file,{from:field(head,"From"),to:field(head,"To"),date:field(head,"Date"),subject:field(head,"Subject"),body:decodeBody(body)});
  return;
 }
 if(/\.msg$/i.test(file.name)){
  try{
   const pkg=await import("https://esm.sh/@npeersab/msgreader@5.2.0");
   const MsgReader=pkg.default||pkg.MsgReader;
   if(typeof MsgReader!=="function")throw Error("MSG parser unavailable");
   const info=new MsgReader(await file.arrayBuffer()).getFileData();
   show(file,{from:info.senderEmail||info.senderSmtpAddress||info.senderName||"—",to:info.displayTo||"—",date:info.messageDeliveryTime||info.clientSubmitTime||"—",subject:info.subject||"—",body:info.body||info.bodyText||"Текст відсутній. Можна завантажити оригінал."});
  }catch(e){throw Error("Не вдалося прочитати MSG: "+(e.message||e));}
 }
}
window.RBD_MAIL_VIEWER=Object.freeze({open});
})();