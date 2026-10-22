/* RBD secure API client. No secrets, no anonymous Google Apps Script calls.
 * API must verify Microsoft identity and per-request access on every operation.
 */
(function(){
 "use strict";
 const cfg=window.RBD_SECURE_API||{};
 const base=String(cfg.baseUrl||"").replace(/\/+$/,"");
 const ready=/^https:\/\//i.test(base)&&typeof cfg.getAccessToken==="function";
 async function call(path,opts){
  if(!ready)throw Error("Захищений сервер РБД ще не підключено");
  const token=await cfg.getAccessToken();
  if(!token||typeof token!=="string")throw Error("Не вдалося підтвердити корпоративний вхід");
  const res=await fetch(base+path,{...opts,headers:{"Authorization":"Bearer "+token,...opts?.headers},credentials:"omit",cache:"no-store"});
  if(!res.ok)throw Error("Сервер РБД повернув помилку "+res.status);
  return res;
 }
 function pathId(id){if(!/^[\p{L}\p{N}_-]{3,80}$/u.test(String(id||"")))throw Error("Некоректний ID заявки");return encodeURIComponent(id);}
 async function uploadLetter(requestId,file){
  if(!file||!(/\.(eml|msg)$/i.test(file.name))||file.size>4*1024*1024||file.size===0)throw Error("Дозволено .eml/.msg до 4 МБ");
  const form=new FormData();form.append("file",file,file.name);
  const res=await call("/api/requests/"+pathId(requestId)+"/letters",{method:"POST",body:form});
  return res.json();
 }
 async function listLetters(requestId){
  const res=await call("/api/requests/"+pathId(requestId)+"/letters",{method:"GET"});
  return res.json();
 }
 async function getLetter(requestId,fileId){
  const res=await call("/api/requests/"+pathId(requestId)+"/letters/"+encodeURIComponent(fileId),{method:"GET"});
  return res.blob();
 }
 window.RBD_GOOGLE_BRIDGE=Object.freeze({ready,uploadLetter,listLetters,getLetter});
})();