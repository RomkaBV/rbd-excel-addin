/** RBD Google pilot backend. Deploy as a separate Apps Script in the SAME Google account as catalog photos.
 * The spreadsheet stores test reminders only. Files are private; never call setSharing(ANYONE_WITH_LINK).
 * Setup: set Script Properties RBD_FOLDER_ID and RBD_SHEET_ID; run setupRbdGooglePilot(), then installRbdReminderTrigger().
 * Restrict access to the Web App. Do NOT publish a web app with "Anyone" + "execute as me".
 */
const RBD_REMINDER_HEADERS=['id','requestId','email','dueAt','subject','body','status','sentAt','lastError'];
function rbdProps_(){return PropertiesService.getScriptProperties();}
function rbdFolder_(){
 const id=rbdProps_().getProperty('RBD_FOLDER_ID');
 if(!id)throw new Error('Set RBD_FOLDER_ID in Script Properties');
 return DriveApp.getFolderById(id);
}
function rbdSheet_(){
 const id=rbdProps_().getProperty('RBD_SHEET_ID');
 if(!id)throw new Error('Set RBD_SHEET_ID in Script Properties');
 const book=SpreadsheetApp.openById(id);
 return book.getSheetByName('RBD_Reminders')||book.insertSheet('RBD_Reminders');
}
function setupRbdGooglePilot(){
 rbdFolder_();
 const sheet=rbdSheet_();
 if(sheet.getLastRow()===0)sheet.appendRow(RBD_REMINDER_HEADERS);
 return 'Pilot ready';
}
function rbdSafe_(name){
 const s=String(name||'').trim();
 if(!/^[A-Za-z0-9А-Яа-яІіЇїЄєҐґ_-]{3,80}$/.test(s))throw new Error('Invalid request number');
 return s;
}
function rbdSubfolder_(parent,name){
 const iter=parent.getFoldersByName(name);
 return iter.hasNext()?iter.next():parent.createFolder(name);
}
/** Call from authenticated Apps Script only; do not expose raw base64 to an unauthenticated public web app. */
function saveRbdLetter(requestId,fileName,base64){
 const id=rbdSafe_(requestId), name=String(fileName||'');
 if(!/\.(eml|msg)$/i.test(name))throw new Error('Only .eml/.msg are allowed');
 const bytes=Utilities.base64Decode(base64||'');
 if(!bytes.length||bytes.length>4*1024*1024)throw new Error('Invalid file size');
 const dir=rbdSubfolder_(rbdSubfolder_(rbdFolder_(),id),'Листування');
 const safeName=new Date().toISOString().replace(/[:.]/g,'-')+'_'+name.replace(/[\\/:*?"<>|#%]/g,'_');
 const blob=Utilities.newBlob(bytes,'application/octet-stream',safeName);
 const file=dir.createFile(blob);
 // Private file: permissions inherited from folder, not public link-sharing.
 return {id:file.getId(),name:file.getName(),url:file.getUrl()};
}
function addRbdTestReminder(requestId,email,isoDate,text){
 const id=rbdSafe_(requestId);
 if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new Error('Invalid recipient email');
 const when=new Date(isoDate);
 if(isNaN(when.getTime()))throw new Error('Invalid date');
 const sheet=rbdSheet_();
 const uuid=Utilities.getUuid();
 sheet.appendRow([uuid,id,email,when,'Нагадування за заявкою '+id,String(text||''),'Pending','','']);
 return uuid;
}
function runRbdReminders(){
 const lock=LockService.getScriptLock();
 if(!lock.tryLock(1000))return;
 try{
  const sheet=rbdSheet_();const rows=sheet.getDataRange().getValues();const now=new Date();
  for(let i=1;i<rows.length;i++){
   const row=rows[i];if(row[6]!=='Pending'||!(row[3] instanceof Date)||row[3]>now)continue;
   try{
    GmailApp.sendEmail(String(row[2]),String(row[4]),String(row[5]));
    sheet.getRange(i+1,7,1,3).setValues([['Sent',new Date(),'']]);
   }catch(e){sheet.getRange(i+1,9).setValue(String(e).slice(0,500));}
  }
 }finally{lock.releaseLock();}
}
function installRbdReminderTrigger(){
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='runRbdReminders').forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger('runRbdReminders').timeBased().everyMinutes(5).create();
}
