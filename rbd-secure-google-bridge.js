/* Authenticated API only. Tokens live in memory and are never sent to Apps Script. */
(function(){
"use strict";
const cfg=window.RBD_SECURE_API||window.RBD_CONFIG?.secureApi||{},base=String(cfg.baseUrl||"").replace(/\/+$/,"");
const ready=/^https:\/\//.test(base)&&(typeof cfg.getAccessToken==="function"||!!(cfg.tenantId&&cfg.clientId&&cfg.scope));
function fail(message,code,status){return Object.assign(Error(message),{code,status});}
async function call(path,opts={}){
  if(!ready)throw fail("Захищений архів РБД ще не підключено","NOT_CONFIGURED",503);
  const token=await (cfg.getAccessToken?cfg.getAccessToken():window.RBD_ENTRA.getAccessToken());
  if(!token||typeof token!=="string")throw fail("Потрібен корпоративний вхід Microsoft","AUTH_REQUIRED",401);
  const abort=new AbortController(),timer=setTimeout(()=>abort.abort(),30000);
  try{
    const res=await fetch(base+path,{...opts,headers:{...opts.headers,Authorization:"Bearer "+token},credentials:"omit",cache:"no-store",redirect:"error",signal:abort.signal});
    if(!res.ok){const messages={401:"Сеанс Microsoft завершився. Увійдіть повторно через меню ⋮.",403:"Недостатньо прав на листи цієї заявки.",404:"Файл або заявку не знайдено.",413:"Файл перевищує 4 МБ.",429:"Забагато запитів. Спробуйте трохи пізніше.",503:"Архів тимчасово недоступний або не налаштований."};if(res.status===401)window.RBD_ENTRA?.clear();throw fail(messages[res.status]||"Помилка архіву: "+res.status,res.status===401?"AUTH_EXPIRED":"API_ERROR",res.status);}
    return res;
  }catch(e){if(e.name==="AbortError")throw fail("Сервер не відповів вчасно. Повторіть завантаження.","TIMEOUT",0);throw e;}finally{clearTimeout(timer);}
}
function pathId(id){if(!/^RBD-[A-Za-z0-9_-]{1,76}$/.test(String(id||"")))throw Error("Некоректний ID заявки");return encodeURIComponent(id);}
async function uploadLetter(requestId,file,uploadId){
  if(!file||!(/\.(eml|msg)$/i.test(file.name))||file.size>4*1024*1024||file.size===0)throw Error("Дозволено .eml/.msg до 4 МБ");
  const form=new FormData();form.append("file",file,file.name);
  return (await call("/api/requests/"+pathId(requestId)+"/letters",{method:"POST",headers:{"Idempotency-Key":uploadId||crypto.randomUUID()},body:form})).json();
}
async function listLetters(requestId){return (await call("/api/requests/"+pathId(requestId)+"/letters")).json();}
async function getLetter(requestId,fileId){if(!/^[A-Za-z0-9_-]{10,200}$/.test(fileId))throw Error("Некоректний ID файла");return (await call("/api/requests/"+pathId(requestId)+"/letters/"+encodeURIComponent(fileId))).blob();}
async function signIn(){return window.RBD_ENTRA.signIn();}
window.RBD_GOOGLE_BRIDGE=Object.freeze({ready,uploadLetter,listLetters,getLetter,signIn});
})();
