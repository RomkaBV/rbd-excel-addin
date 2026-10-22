/* Import private Google letters registry CSV/XLSX exported by authorized owner into RBD workbook. */
(function(){
"use strict";
async function loadRegistry(file){
 if(!window.XLSX)throw Error("Модуль Excel XLSX недоступний");
 const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:false});
 const sheet=wb.Sheets[wb.SheetNames[0]];
 const data=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:false});
 const headers=(data[0]||[]).map(x=>String(x).trim());
 const expected=["ID заявки","ID файла","Назва файла","Тип","Розмір, байт","Дата створення","Посилання Google Drive"];
 if(expected.some(h=>!headers.includes(h)))throw Error("Неправильні колонки реєстру листів");
 const rows=data.slice(1).filter(r=>r.some(Boolean)).map(row=>expected.map(h=>String(row[headers.indexOf(h)]||"")));
 for(const row of rows){
  if(!/^RBD-[A-Za-z0-9_-]+$/.test(row[0])||!/^[A-Za-z0-9_-]{10,200}$/.test(row[1])||!/(eml|msg)$/i.test(row[3]))throw Error("Некоректний запис реєстру");
 }
 await Excel.run(async ctx=>{
  const name="tblRBDLetterRegistry2026";
  let table=ctx.workbook.tables.getItemOrNullObject(name);
  table.load("isNullObject");await ctx.sync();
  if(table.isNullObject){
   let ws=ctx.workbook.worksheets.getItemOrNullObject("Реєстр листів");ws.load("isNullObject");await ctx.sync();
   if(ws.isNullObject)ws=ctx.workbook.worksheets.add("Реєстр листів");
   ws.getRangeByIndexes(0,0,1,expected.length).values=[expected];
   table=ws.tables.add(ws.getRangeByIndexes(0,0,1,expected.length),true);table.name=name;await ctx.sync();
  }
  const body=table.getRange();body.load("values");await ctx.sync();
  const existing=new Set(body.values.slice(1).map(r=>String(r[1]||"")));
  const fresh=rows.filter(row=>!existing.has(row[1]));
  if(fresh.length)table.rows.add(null,fresh);
  await ctx.sync();
 });
 return rows.length;
}
function init(){
 const head=document.querySelector("#requestModal .history-side-head");if(!head||document.getElementById("rbdImportLetterRegistry"))return;
 const button=document.createElement("button");button.id="rbdImportLetterRegistry";button.type="button";button.className="btn";button.textContent="↑ Реєстр листів";button.title="Імпортувати CSV або Excel з приватного Google реєстру";
 const input=document.createElement("input");input.type="file";input.accept=".csv,.xlsx,.xls";input.hidden=true;
 button.addEventListener("click",()=>input.click());
 input.addEventListener("change",async()=>{
  const file=input.files?.[0];if(!file)return;
  button.disabled=true;
  try{
   const n=await loadRegistry(file);
   button.textContent="✓ "+n+" лист(ів)";
   if(typeof loadBusinessData==="function")await loadBusinessData();
   if(typeof STATE!=="undefined"&&STATE.selectedRequest&&typeof renderHistorySide==="function")renderHistorySide(STATE.selectedRequest.id);
  }catch(e){button.textContent="Помилка імпорту";window.alert(String(e.message||e));}
  finally{button.disabled=false;input.value="";}
 });
 head.appendChild(button);head.appendChild(input);
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init);else init();
})();