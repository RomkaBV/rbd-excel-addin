/* Pure calendar and visit calculations, shared by the UI and regression tests. */
(function(root,factory){const api=factory();if(typeof module==="object"&&module.exports)module.exports=api;else root.RBD_SLA=api;})(globalThis,function(){
"use strict";
const norm=s=>String(s||"").trim().toLocaleLowerCase("uk-UA");
const number=v=>Number(String(v??"").replace(/\s/g,"").replace(",","."))||0;
const DEFAULT={timeZone:"Europe/Kyiv",weekdays:[1,2,3,4,5],intervals:[["09:00","18:00"]],holidays:[]};
function calendar(value={}){
  const c={...DEFAULT,...value};
  const minutes=s=>{if(!/^\d{2}:\d{2}$/.test(s))throw Error("Некоректний час робочого графіка");const [h,m]=s.split(":").map(Number);if(h>23||m>59)throw Error("Некоректний час робочого графіка");return h*60+m;};
  c.spans=c.intervals.map(([a,b])=>[minutes(a),minutes(b)]).sort((a,b)=>a[0]-b[0]);
  if(!c.spans.length||c.spans.some(([a,b],i)=>b<=a||(i>0&&a<c.spans[i-1][1])))throw Error("Робочі інтервали мають бути непорожні та не перетинатися");
  c.hoursPerDay=c.spans.reduce((n,[a,b])=>n+(b-a)/60,0);
  c.formatter=new Intl.DateTimeFormat("en-CA",{timeZone:c.timeZone,year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit",second:"2-digit",hourCycle:"h23"});
  return c;
}
// Excel serials and zone-less workbook timestamps are business wall time.
function moment(value,c){
  if(value===null||value===undefined||value==="")return NaN;
  if(typeof value==="number")return (value-25569)*86400000;
  if(!(value instanceof Date)){
    const s=String(value).trim();
    const plain=/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):?(\d{2})(?::?(\d{2})(?:\.\d+)?)?)?$/.exec(s);
    const uk=/^(\d{2})\.(\d{2})\.(\d{4})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
    const m=plain|| (uk?[uk[0],uk[3],uk[2],uk[1],uk[4],uk[5],uk[6]]:null);
    if(m){const t=Date.UTC(+m[1],+m[2]-1,+m[3],+(m[4]||0),+(m[5]||0),+(m[6]||0));const d=new Date(t);return d.getUTCMonth()===+m[2]-1&&d.getUTCDate()===+m[3]?t:NaN;}
  }
  const d=value instanceof Date?value:new Date(value);if(!Number.isFinite(d.getTime()))return NaN;
  const p=Object.fromEntries(c.formatter.formatToParts(d).map(x=>[x.type,x.value]));
  return Date.UTC(+p.year,+p.month-1,+p.day,+p.hour,+p.minute,+p.second);
}
function hoursBetween(start,end,config={}){
  const c=config.spans?config:calendar(config),a=moment(start,c),b=moment(end,c);
  if(!Number.isFinite(a)||!Number.isFinite(b)||b<a)return null;
  let total=0;
  for(let day=Math.floor(a/86400000)*86400000;day<b;day+=86400000){
    const d=new Date(day);if(!c.weekdays.includes(d.getUTCDay())||c.holidays.includes(d.toISOString().slice(0,10)))continue;
    for(const [from,to] of c.spans)total+=Math.max(0,Math.min(b,day+to*60000)-Math.max(a,day+from*60000));
  }
  return total/3600000;
}
function calculate(request,events,statuses,config={},now=new Date()){
  const c=calendar(config),key=norm,at=v=>moment(v,c),valid=v=>Number.isFinite(at(v));
  const definition=name=>statuses.find(s=>key(s["Назва статусу"])===key(name));
  const final=name=>["так","true","1","yes"].includes(key(definition(name)?.["Завершальний"]))||(!definition(name)&&key(name)==="закрита");
  const source=events.filter(e=>String(e["ID заявки"])===String(request.id));
  const seen=new Set();
  const changes=source.filter(e=>e["Тип"]==="Статус"||e["Дія"]==="Змінено статус").filter(e=>{
    const k=[e["Дата події"],e["Було"],e["Стало"],e["Початок статусу"],e["Кінець статусу"]].join("|");
    if(seen.has(k)||!valid(e["Дата події"])||at(e["Дата події"])>at(now)||key(e["Було"])===key(e["Стало"]))return false;seen.add(k);return true;
  }).sort((a,b)=>at(a["Дата події"])-at(b["Дата події"]));
  const creation=source.find(e=>e["Тип"]==="Створення");
  let lastMoment=request.createdAt||creation?.["Дата події"]||"",lastName=changes[0]?.["Було"]||creation?.["Стало"]||request.status;
  const visits=[];
  function add(name,start,end,current=false){
    if(!name)return;
    let hours=hoursBetween(start,end,c);
    if(final(name)&&hours!==null)hours=0; // closed time never ages a reopened request
    const sla=number(definition(name)?.["SLA, год"]);
    visits.push({name,start:valid(start)?start:"",end:current&&!final(name)?"":end,hours,current,ongoing:current&&!final(name),slaHours:sla>0?sla:null,overdueHours:sla>0&&hours!==null?Math.max(0,hours-sla):null});
  }
  for(const e of changes){
    const name=e["Було"]||lastName,end=e["Кінець статусу"]||e["Дата події"];
    let start=e["Початок статусу"]||((key(name)===key(lastName))?lastMoment:"");
    // A stale explicit start must not overlap a previous completed visit.
    if(valid(start)&&valid(lastMoment)&&at(start)<at(lastMoment))start=lastMoment;
    add(name,start,end);lastMoment=end;lastName=e["Стало"]||"";
  }
  let start=request.statusSince||(!changes.length?lastMoment:"");
  if(changes.length&&key(lastName)===key(request.status))start=valid(start)&&at(start)>at(lastMoment)?start:lastMoment;
  // Unknown entry times stay unknown; no invented duration for incomplete history.
  add(request.status,start,final(request.status)?start:now,true);
  const names=statuses.slice().sort((a,b)=>number(a["Порядок"])-number(b["Порядок"])).map(s=>s["Назва статусу"]).filter(Boolean);
  for(const v of visits)if(!names.some(n=>key(n)===key(v.name)))names.push(v.name);
  const stages=names.map(name=>{
    const entries=visits.filter(v=>key(v.name)===key(name)),latest=entries.at(-1),current=key(name)===key(request.status);
    return {name,count:entries.length,current,start:latest?.start||"",end:latest?.end||"",visits:entries,completed:!current&&entries.length>0};
  });
  const timed=visits.filter(v=>v.slaHours!==null);
  const unknown=timed.some(v=>v.hours===null);
  const current=visits.find(v=>v.current);
  const working=visits.filter(v=>!final(v.name)&&key(v.name)!=="призупинена");
  return {stages,visits,current,hoursPerDay:c.hoursPerDay,currentOverdue:current?.overdueHours===null?null:current.overdueHours/c.hoursPerDay,totalOverdue:unknown?null:timed.reduce((n,v)=>n+(v.overdueHours||0),0)/c.hoursPerDay,totalWorkHours:working.some(v=>v.hours===null)?null:working.reduce((n,v)=>n+v.hours,0),hasSla:timed.length>0};
}
return {calendar,hoursBetween,calculate};
});
