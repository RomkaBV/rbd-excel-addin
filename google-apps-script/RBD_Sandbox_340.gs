/**
 * RBD 3.4.0 DEMO — ISOLATED TEST DATA ONLY.
 * Paste into a NEW Google Apps Script project, not the production RBD Google script.
 * Web app public endpoint is deliberately restricted to fixed synthetic content,
 * demo request ID, tiny quota, and one dedicated sandbox folder.
 * Never paste confidential company emails into this demo.
 */
const RBD_DEMO_FOLDER_ID = "1Wy_9rvTuCfv0pQ6BQXsLOSbk6yab1uwW";
const RBD_DEMO_ID = "RBD-DEMO-001";
const RBD_DEMO_MAX_PER_DAY = 10;
const RBD_DEMO_EML = [
 "From: demo@example.invalid",
 "To: demo@example.invalid",
 "Subject: RBD synthetic test letter",
 "MIME-Version: 1.0",
 "Content-Type: text/plain; charset=UTF-8",
 "",
 "This is a synthetic test file. No real company correspondence."
].join("\r\n");

function rbdDemoFolder_(){
 return DriveApp.getFolderById(RBD_DEMO_FOLDER_ID);
}
function rbdDemoJson_(obj){
 return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e){
 // No public listing, download or read endpoint; sandbox test status only.
 return rbdDemoJson_({demo:true,version:"3.4.0-test",message:"Sandbox only. No production files accessible."});
}
function doPost(e){
 const lock=LockService.getScriptLock();
 if(!lock.tryLock(3000))return rbdDemoJson_({ok:false,error:"busy"});
 try {
  const data=JSON.parse(e.postData?.contents||"{}");
  if(data.action!=="createSyntheticLetter"||data.requestId!==RBD_DEMO_ID)
   return rbdDemoJson_({ok:false,error:"demo-action-only"});
  const p=PropertiesService.getScriptProperties();
  const day=Utilities.formatDate(new Date(),"UTC","yyyyMMdd");
  const key="DEMO_WRITE_"+day;
  const count=Number(p.getProperty(key)||"0");
  if(count>=RBD_DEMO_MAX_PER_DAY)return rbdDemoJson_({ok:false,error:"daily-test-limit"});
  const root=rbdDemoFolder_();
  const dirs=root.getFoldersByName(RBD_DEMO_ID);
  const dir=dirs.hasNext()?dirs.next():root.createFolder(RBD_DEMO_ID);
  const mails=dir.getFoldersByName("Листування");
  const maildir=mails.hasNext()?mails.next():dir.createFolder("Листування");
  const file=maildir.createFile("DEMO_"+Utilities.getUuid()+".eml",RBD_DEMO_EML,MimeType.PLAIN_TEXT);
  p.setProperty(key,String(count+1));
  return rbdDemoJson_({ok:true,requestId:RBD_DEMO_ID,fileName:file.getName(),demo:true});
 }catch(err){return rbdDemoJson_({ok:false,error:"invalid-demo-request"});}
 finally{lock.releaseLock();}
}
function verifyRbdDemoSetup(){
 const folder=rbdDemoFolder_();
 Logger.log("Demo sandbox: "+folder.getName()+" "+folder.getUrl());
 Logger.log("Synthetic demo only; no email sending");
}

function testRbdDemoWrite(){
 const result=doPost({postData:{contents:JSON.stringify({action:"createSyntheticLetter",requestId:"RBD-DEMO-001"})}});
 Logger.log(result.getContent());
}
