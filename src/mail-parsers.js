import PostalMime from "postal-mime";
import MsgReader from "@kenjiuno/msgreader";
import DOMPurify from "dompurify";

const address=x=>!x?"":x.group?x.group.map(address).join(", "):[x.name,x.address?`<${x.address}>`:""].filter(Boolean).join(" ");
export async function parse(buffer,name){
  if(!buffer.byteLength||buffer.byteLength>4*1024*1024)throw Error("Дозволено непорожні листи до 4 МБ");
  if(/\.eml$/i.test(name)){
    const email=await PostalMime.parse(buffer,{maxNestingDepth:40,maxHeadersSize:256*1024,maxRfc822NestingDepth:5,maxPartCount:500});
    if(!email.headers.length)throw Error("Файл не містить заголовків EML");
    return {from:address(email.from),to:(email.to||[]).map(address).join(", "),subject:email.subject||"",date:email.date||"",text:email.text||"",html:email.html||"",attachments:email.attachments.map((a,i)=>({name:a.filename||`Вкладення-${i+1}`,type:a.mimeType,content:a.content,contentId:a.contentId||""}))};
  }
  if(!/\.msg$/i.test(name))throw Error("Підтримуються .eml та .msg");
  const magic=new Uint8Array(buffer,0,Math.min(buffer.byteLength,8));
  if([...magic].join(",")!=="208,207,17,224,161,177,26,225")throw Error("Некоректний формат MSG");
  const reader=new MsgReader(buffer),data=reader.getFileData();
  if(data.error)throw Error(data.error);
  const headers=data.headers?await PostalMime.parse(data.headers+"\r\n\r\n"):{};
  const html=data.bodyHtml||(data.html?new TextDecoder("utf-8").decode(data.html):"");
  return {from:address({name:data.senderName,address:data.senderSmtpAddress||data.senderEmail})||address(headers.from),to:data.recipients?.filter(r=>r.recipType!=="cc"&&r.recipType!=="bcc").map(r=>address({name:r.name,address:r.smtpAddress||r.email})).join(", ")||data.displayTo||(headers.to||[]).map(address).join(", "),subject:data.subject||headers.subject||"",date:data.messageDeliveryTime||data.clientSubmitTime||headers.date||"",text:data.body||"",html,attachments:(data.attachments||[]).map((a,i)=>{try{const item=reader.getAttachment(a);return {name:item.fileName||a.fileName||`Вкладення-${i+1}`,type:a.attachMimeTag||"application/octet-stream",content:item.content,contentId:a.pidContentId||""};}catch{return {name:a.fileName||`Вкладення-${i+1}`,error:"Вкладення доступне лише в оригіналі Outlook"};}})};
}
export function sanitize(html){
  return DOMPurify.sanitize(html,{USE_PROFILES:{html:true},FORBID_TAGS:["form","input","button","textarea","select","style","link","meta","base","iframe","object","embed","video","audio"],FORBID_ATTR:["style","srcset","background","action","formaction","ping","src"]});
}
