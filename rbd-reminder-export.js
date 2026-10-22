/* RBD: manual secure test export of pending reminder queue as JSON.
 * Staff cannot submit directly to owner's Google account without protected API.
 */
(function(){
"use strict";
async function exportReminders(){
 const btn=document.getElementById("rbdExportReminders");
 try{
  if(btn)btn.disabled=true;
  const rows=await Excel.run(async ctx=>{
   const t=ctx.workbook.tables.getItemOrNullObject("tblRBDReminderQueue2026");
   t.load("isNullObject");await ctx.sync();
   if(t.isNullObject)throw Error("Спочатку створіть нагадування у РБД");
   const data=t.getRange();data.load("values");await ctx.sync();return data.values;
  });
  const header=rows[0].map(v=>String(v||""));
  const col=name=>header.indexOf(name);
  const fields=["ID нагадування","ID заявки","Email","Дата нагадування","Текст","Стан синхронізації"];
  if(fields.some(name=>col(name)<0))throw Error("У таблиці черги бракує потрібних колонок");
  const items=rows.slice(1).filter(r=>String(r[col("Стан синхронізації")]).includes("Очікує")).map(r=>{
   const due=String(r[col("Дата нагадування")]||"");
   const date=new Date(due);
   if(Number.isNaN(date.getTime()))throw Error("Некоректна дата: "+due);
   return {id:String(r[col("ID нагадування")]),requestId:String(r[col("ID заявки")]),originalEmail:String(r[col("Email")]),dueISO:date.toISOString(),body:String(r[col("Текст")])};
  });
  if(!items.length)throw Error("Немає нагадувань, які очікують синхронізації");
  const blob=new Blob([JSON.stringify({format:"rbd-reminders-v1",testMode:true,items},null,2)],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const link=document.createElement("a");link.href=url;link.download="RBD_reminders_"+new Date().toISOString().slice(0,10)+".json";
  document.body.appendChild(link);link.click();link.remove();
  setTimeout(()=>URL.revokeObjectURL(url),60000);
  document.getElementById("rbdExportStatus").textContent="Експортовано "+items.length+" записів. Завантаж JSON до Google Drive → RBD → Імпорт нагадувань. У Gmail з Excel автоматично ще нічого не надіслано.";
 }catch(e){document.getElementById("rbdExportStatus").textContent="Експорт: "+(e.message||e);}
 finally{if(btn)btn.disabled=false;}
}
function init(){
 const anchor=document.getElementById("reminderAddButton");if(!anchor||document.getElementById("rbdExportReminders"))return;
 const b=document.createElement("button");b.id="rbdExportReminders";b.type="button";b.className="btn";b.textContent="↓ Експорт у Google (тест)";
 b.addEventListener("click",exportReminders);anchor.insertAdjacentElement("afterend",b);
 const status=document.createElement("div");status.id="rbdExportStatus";status.style.cssText="font-size:10px;color:#3f5e79;margin-top:6px;overflow-wrap:anywhere";
 anchor.parentElement.insertAdjacentElement("afterend",status);
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init);else init();
})();