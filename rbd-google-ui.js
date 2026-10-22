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
 const title=document.querySelector("#requestModal .history-side .drawer-tool-title");
 if(!title||!title.parentNode)return;
 const row=document.createElement("div");
 row.style.cssText="padding:7px 0;font-size:11px;color:#426781";
 const cfg=window.RBD_GOOGLE_CONFIG||{};
 row.textContent=cfg.webAppUrl?"Google тест: адресу бекенду задано (доступ ще потребує перевірки).":"Google Drive / Gmail: тестовий режим — Apps Script ще не підключено. Файли залишаються локально, поки не налаштований сервер.";
 title.parentNode.insertBefore(row,title.nextSibling);
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init);else init();
})();