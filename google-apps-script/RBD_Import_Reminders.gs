/* Install alongside the existing RBD Google Code.gs. Manual JSON import from Drive.
   Test ONLY: all imported reminders go to a single test mailbox.
   Uses existing setupRBD Script Property RBD_SHEET_ID and Reminders sheet.
   No unauthenticated web endpoints. */
const RBD_IMPORT_ROOT = "18hQb_ym1aXPD2sDIlvEBOIlntcfXhEw4";
const RBD_TEST_RECIPIENT = "roman.barzo@stv.group";
function importRbdReminderFiles(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw new Error("Імпорт уже запущено");
  try {
    const root=DriveApp.getFolderById(RBD_IMPORT_ROOT);
    const folders=root.getFoldersByName("Імпорт нагадувань");
    const folder=folders.hasNext()?folders.next():root.createFolder("Імпорт нагадувань");
    const bookId=PropertiesService.getScriptProperties().getProperty("RBD_SHEET_ID");
    if(!bookId)throw new Error("Спочатку запусти setupRBD — відсутній RBD_SHEET_ID");
    const sheet=SpreadsheetApp.openById(bookId).getSheetByName("Reminders");
    if(!sheet)throw new Error("Немає аркуша Reminders");
    const values=sheet.getDataRange().getValues();
    const known=new Set(values.slice(1).map(r=>String(r[0]).trim()).filter(Boolean));
    const files=folder.getFilesByType(MimeType.PLAIN_TEXT);
    let added=0,duplicates=0,processed=0;
    // Google Drive may classify JSON as application/json, so scan all files.
    const all=folder.getFiles();
    while(all.hasNext()){
      const file=all.next(), name=file.getName();
      if(!name.endsWith(".json")||!name.startsWith("RBD_reminders_"))continue;
      const obj=JSON.parse(file.getBlob().getDataAsString("UTF-8"));
      if(obj.format!=="rbd-reminders-v1"||obj.testMode!==true||!Array.isArray(obj.items))throw new Error("Некоректний формат "+name);
      const append=[];
      for(const item of obj.items){
        const id=String(item.id||"").trim(),req=String(item.requestId||"").trim();
        if(!/^[A-Za-z0-9_-]{4,100}$/.test(id)||!/^[A-Za-z0-9_-]{3,80}$/.test(req))throw new Error("Некоректний ID");
        if(known.has(id)){duplicates++;continue;}
        const due=new Date(item.dueISO);
        if(isNaN(due.getTime()))throw new Error("Некоректна дата");
        const body=String(item.body||"").slice(0,5000);
        if(!body)throw new Error("Порожнє нагадування");
        append.push([id,req,RBD_TEST_RECIPIENT,due,"Нагадування РБД: "+req,body,"Pending","",""]);
        known.add(id);
      }
      if(append.length){sheet.getRange(sheet.getLastRow()+1,1,append.length,9).setValues(append);added+=append.length;}
      file.setName(name+".imported");processed++;
    }
    Logger.log("Оброблено файлів: "+processed+"; додано: "+added+"; дублікатів: "+duplicates);
    return {processed,added,duplicates};
  }finally{lock.releaseLock();}
}
