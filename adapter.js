/* РБД 2.0: адаптер до погодженого інтерфейсу 2026-09-25.
   Підключається ПІСЛЯ оригінального JS. Жодних демонстраційних заявок.
   Запис — у чинну чергу команд кабінету. Виконання — основним Office Script. */
const RBD2 = {
  users: {
    'Іванченко':'Іванченко В.М.','Войцехівський':'Войцехівський Г.В.',
    'Ридванський':'Ридванський П.С.','Галько':'Галько А.І.',
    'Трунов':'Трунов Ю.О.','Желясков':'Желясков Д.О.',
    'Слепущенко':'Слепущенко О.О.','Сергеєв':'Сергеєв П.А.',
    'Туровський':'Туровський В.О.'
  },
  tableName:'', schema:'', locations:[], priorities:[], customers:[], statuses:[]
};
function r2Get(headers,row,...keys){
  for(const key of keys){const i=headers.indexOf(key);if(i>=0)return row[i]??'';}return '';
}
function r2Unique(a){return [...new Set(a.map(x=>String(x||'').trim()).filter(Boolean))];}
async function r2TryTable(context,names){
  for(const name of names){const r=await readTable(context,name,false);if(r.headers.length)return {...r,name};}
  return {headers:[],rows:[],name:''};
}
async function r2ScanTables(context,test){
  context.workbook.tables.load('items/name');await context.sync();
  for(const item of context.workbook.tables.items){const r=await readTable(context,item.name,false);if(test(r.headers))return {...r,name:item.name};}
  return {headers:[],rows:[],name:''};
}
async function detectMode(){
  const tab=await Excel.run(async ctx=>{const s=ctx.workbook.worksheets.getActiveWorksheet();s.load('name');await ctx.sync();return s.name;});
  STATE.activeSheet=tab;
  const plain=tab.replace(/^РБД_/,'').trim();
  if(['Кабінет керівника','РБД_Керівник'].includes(tab)){
    STATE.employeeMode=false;STATE.employee='';STATE.actor=RBD2.users['Іванченко'];return;
  }
  if(RBD2.users[plain]){STATE.employeeMode=true;STATE.employee=RBD2.users[plain];STATE.actor=STATE.employee;return;}
  if(RBD2.users[tab]){STATE.employeeMode=true;STATE.employee=RBD2.users[tab];STATE.actor=STATE.employee;return;}
  if(STATE.actor)return;
  throw new Error('Відкрийте «Кабінет керівника» або персональну вкладку працівника (за прізвищем).');
}
async function loadDictionaries(){
  const data=await Excel.run(async c=>{
    const pairs={};
    for(const [k,names,head] of [
      ['employees',['RBD_Users','RBD_Employees'],['ПІБ','Display_Name','Прізвище']],
      ['statuses',['RBD_Statuses'],['Назва статусу','Status']],
      ['cities',['RBD_Cities'],['Назва','City','Місто']],
      ['categories',['RBD_Categories'],['Назва','Category','Категорія']],
      ['customers',['RBD_Customers'],['Назва','Customer','Замовник']],
      ['priorities',['RBD_Priorities'],['Назва','Priority','Пріоритет']],
      ['locations',['RBD_Locations'],['Назва','Location','Локація']]
    ]){
      let x=await r2TryTable(c,names);
      if(!x.headers.length && ['locations'].includes(k))x=await r2ScanTables(c,h=>h.includes('ID локації')&&h.includes('Місто')); 
      pairs[k]={headers:x.headers,rows:x.rows};
    }
    return pairs;
  });
  const vals=(name,headers)=>{
    const d=data[name];return r2Unique(d.rows.filter(r=>!d.headers.includes('Активний')||String(r[d.headers.indexOf('Активний')]).trim().toLowerCase()!=='ні').map(r=>r2Get(d.headers,r,...headers)));
  };
  STATE.employees=vals('employees',['ПІБ','Display_Name','Прізвище']);
  if(!STATE.employees.length)STATE.employees=Object.values(RBD2.users);
  STATE.statuses=vals('statuses',['Назва статусу','Status']);
  if(!STATE.statuses.length)STATE.statuses=FALLBACK_STATUSES.slice();
  STATE.cities=vals('cities',['Назва','City','Місто']);
  STATE.categories=vals('categories',['Назва','Category','Категорія']);
  RBD2.customers=vals('customers',['Назва','Customer','Замовник']);
  RBD2.priorities=vals('priorities',['Назва','Priority','Пріоритет']);
  RBD2.locations=data.locations.rows.map(r=>({city:String(r2Get(data.locations.headers,r,'Місто','City')).trim(),name:String(r2Get(data.locations.headers,r,'Назва','Location','Локація')).trim()})).filter(v=>v.name);
  [['filterStatus',STATE.statuses,'Усі статуси'],['filterExecutor',STATE.employees,'Усі виконавці'],
   ['filterCity',STATE.cities,'Усі міста'],['filterCategory',STATE.categories,'Усі категорії'],
   ['createCategory',STATE.categories,'Оберіть категорію'],['createCity',STATE.cities,'Оберіть місто'],
   ['createExecutor',STATE.employees,'Оберіть виконавця'],['editCategory',STATE.categories,'Оберіть категорію'],
   ['editCity',STATE.cities,'Оберіть місто']].forEach(a=>fillSelect(...a));
  // Старе поле адреси залишається, але має підказки з локацій відповідного міста.
  for(const kind of ['create','edit']){
    const city=document.getElementById(kind+'City');const addr=document.getElementById(kind+'Address');
    if(!city||!addr)continue;
    let list=document.getElementById('rbd2-'+kind+'-locations');
    if(!list){list=document.createElement('datalist');list.id='rbd2-'+kind+'-locations';document.body.append(list);}
    addr.setAttribute('list',list.id);
    const refresh=()=>{list.replaceChildren();RBD2.locations.filter(v=>!v.city||!city.value||v.city===city.value).forEach(v=>{const op=document.createElement('option');op.value=v.name;list.append(op);});};
    city.addEventListener('change',refresh);refresh();
  }
}
function rowToRequest(headers,row){
  const g=(...keys)=>r2Get(headers,row,...keys);
  return {
    id:cleanText(g('ID заявки','RBD_ID','ID')),
    sd:cleanText(g('ID ERP','ERP_ID','Номер заявки SD')),
    category:cleanText(g('Категорія','Category')),
    description:cleanText(g('Опис заявки','Description','Опис','Subject','Тема заявки')),
    city:cleanText(g('Місто','City')),
    address:cleanText(g('Локація','Location','Адреса / локація')),
    created:g('Дата створення','Created_At'),
    amount:toNumber(g('Сума витрат','Amount_UAH')),
    customer:cleanText(g('Замовник','Customer')),
    status:cleanText(g('Статус','Status')),
    executor:cleanText(g('Виконавець','Executor')),
    source:cleanText(g('Джерело','Source_System','Джерело заявки')),
    statusSince:g('Дата зміни статусу','Status_Since','Поточний статус з'),
    previousStatus:cleanText(g('Попередній статус')),
    sla:toNumber(g('SLA, год')),
    storedTimeInStatus:toNumber(g('Час у статусі, год')),
    slaOverdue:cleanText(g('Прострочено SLA')),
    plannedDate:g('Планова дата завершення','Planned_Due'),
    dateOverdue:cleanText(g('Прострочено по даті')),
    pauseReason:cleanText(g('Причина призупинення')),
    comment:cleanText(g('Коментар')),
    closedDate:g('Дата завершення','Completed_At','Дата закриття'),
    version:g('Версія','Version'),subject:cleanText(g('Тема заявки','Subject'))
  };
}
async function loadDashboard(){
  showLoading();
  try{
    const r=await Excel.run(async c=>{
      let t=await r2TryTable(c,['RBD_Requests','tbl_RBD_Base']);
      if(!t.headers.length)t=await r2ScanTables(c,h=>h.includes('ID заявки')&&h.includes('Статус')&&h.includes('Виконавець'));
      const a=await r2TryTable(c,['tbl_RBD_Archive']);return {t,a};
    });
    if(!r.t.headers.length)throw new Error('Не знайдено реєстр заявок. Перевірте таблицю «Заявки» / RBD_Requests.');
    RBD2.tableName=r.t.name;RBD2.schema=r.t.headers.includes('ID заявки')?'uk':r.t.headers.includes('RBD_ID')?'rbd':'legacy';
    STATE.allRequests=r.t.rows.map(row=>rowToRequest(r.t.headers,row)).filter(x=>x.id);
    STATE.archive=r.a.rows.map(row=>rowToRequest(r.a.headers,row)).filter(x=>x.id);
    STATE.requests=STATE.employeeMode?STATE.allRequests.filter(x=>normalizeName(x.executor)===normalizeName(STATE.employee)):STATE.allRequests.slice();
    STATE.expandedStatusId='';updateKpis();applyFilters();updateLastUpdated();
  }catch(e){console.error(e);showTableError(getErrorText(e));}
}
function getRowClass(request){
  if(isOverdue(request))return 'row-overdue';
  if(['Закрито','Закрита','Виконана','Виконано'].includes(request.status))return 'row-rbd2-done';
  if(['Нова'].includes(request.status))return 'row-rbd2-new';
  if(['Призупинена','Скасована','Скасовано'].includes(request.status))return 'row-rbd2-paused';
  if(['Контрактація','На погодженні'].includes(request.status))return 'row-rbd2-contract';
  return 'row-rbd2-wait';
}
function getStatusClass(status){return ['Закрита','Закрито','Виконано','Виконана'].includes(status)?'status-closed':status==='Нова'?'status-new':'status-accepted';}
function getStatusButtonClass(status){return ['Закрита','Закрито','Виконано','Виконана'].includes(status)?'btn-closed':'btn-accepted';}
// Команди записуються до A5:P у персональному кабінеті. Стару tbl_RBD_Queue не використовуємо.
async function r2WriteCabinetCommand(action,event){
  const sheetName=STATE.employeeMode?STATE.activeSheet:(['Кабінет керівника','РБД_Керівник'].includes(STATE.activeSheet)?STATE.activeSheet:'Кабінет керівника');
  const id='CMD-'+Date.now()+'-'+Math.random().toString(36).slice(2,7).toUpperCase();
  await Excel.run(async c=>{
    const sh=c.workbook.worksheets.getItemOrNullObject(sheetName);sh.load('isNullObject');await c.sync();
    if(sh.isNullObject)throw new Error('Не знайдено кабінет для запису команди: '+sheetName);
    sh.tables.load('items/name');await c.sync();
    let target=null,headers=[];
    for(const t of sh.tables.items){const h=t.getHeaderRowRange();h.load('values');await c.sync();const a=h.values[0].map(String);if(a.includes('ID команди')&&a.includes('Дія')){target=t;headers=a;break;}}
    if(!target)throw new Error('Не знайдено таблицю команд у кабінеті '+sheetName+' (потрібні колонки «ID команди», «Дія»).');
    const value={
      'ID команди':id,'Дія':action,'ID заявки':event.requestId||'',
      'Тема заявки':event.subject||(event.description||'').split('\n')[0].slice(0,160),
      'Опис заявки':event.description||'', 'Виконавець':event.newExecutor||event.executor||'',
      'Статус':event.newStatus||'', 'Місто':event.city||'', 'Локація':event.address||'',
      'Категорія':event.category||'', 'Замовник':event.customer||'',
      'Пріоритет':event.priority||'Середній','Коментар':event.comment||'',
      'Очікувана версія':event.version||'','Стан обробки':'Нова','Результат':''
    };
    target.rows.add(null,[headers.map(k=>value[k]??'')]);await c.sync();
  });
  return id;
}
async function enqueue(event){
  const op=event.operation;
  const action=op==='CREATE'?'Створити':(op==='CLOSE'||(op==='STATUS_CHANGE'&&['Закрита','Закрито'].includes(event.newStatus)))?'Закрити':op==='STATUS_CHANGE'?'Змінити статус':op==='COMMENT'?'Коментар':op==='ASSIGN'?'Призначити':null;
  if(!action)throw new Error('Редагування всіх полів старої форми (UPDATE) ще не підтримується процесором команд РБД 2.0. Дані не змінено. Для призначення, статусів і коментарів використовуй відповідні команди.');
  return r2WriteCabinetCommand(action,event);
}
async function waitForStatusUpdate(requestId,expectedStatus){
  // Не підміняємо стан. Старі 45-секундні опитування не відповідають ручному Office Script.
  await loadDashboard();
  const r=STATE.allRequests.find(x=>x.id===requestId);
  return !!r&&r.status===expectedStatus;
}
async function r2FetchEvents(){
  return Excel.run(async c=>{
    let ev=await r2TryTable(c,['RBD_Events']);if(!ev.headers.length)ev=await r2ScanTables(c,h=>h.includes('ID події')&&h.includes('ID заявки'));const rem=await r2TryTable(c,['RBD_Reminders']);
    const old=await r2TryTable(c,['tbl_RBD_Activity']);return {ev,rem,old};
  });
}
async function loadActivity(requestId,showLoading=true){
  const body=document.getElementById('activityBody');if(showLoading&&body)body.textContent='Завантаження історії…';
  try{
    const {ev,rem,old}=await r2FetchEvents();const items=[];
    for(const r of ev.rows){if(cleanText(r2Get(ev.headers,r,'RBD_ID','ID заявки'))!==requestId)continue;
      const action=cleanText(r2Get(ev.headers,r,'Event_Type','Дія','Тип'));
      items.push({id:cleanText(r2Get(ev.headers,r,'Event_ID','ID події')),requestId,
        eventDate:r2Get(ev.headers,r,'Event_At','Дата події'),author:cleanText(r2Get(ev.headers,r,'Actor','Користувач')),
        type:action==='STATUS'?'STATUS':'NOTE',text:[action,r2Get(ev.headers,r,'Field','Змінене поле'),r2Get(ev.headers,r,'Old_Value','Було'),r2Get(ev.headers,r,'New_Value','Стало'),r2Get(ev.headers,r,'Comment','Коментар')].filter(Boolean).join(' · '),reminderDate:'',reminderStatus:'',completedDate:''});
    }
    for(const r of rem.rows){if(cleanText(r2Get(rem.headers,r,'RBD_ID','ID заявки'))!==requestId)continue;
      items.push({id:cleanText(r2Get(rem.headers,r,'Reminder_ID')),requestId,eventDate:r2Get(rem.headers,r,'Created_At'),author:cleanText(r2Get(rem.headers,r,'Created_By')),
        type:'REMINDER',text:cleanText(r2Get(rem.headers,r,'Reminder_Text')),reminderDate:r2Get(rem.headers,r,'Due_At'),reminderStatus:cleanText(r2Get(rem.headers,r,'Status')),completedDate:r2Get(rem.headers,r,'Completed_At')});
    }
    for(const r of old.rows){if(cleanText(r2Get(old.headers,r,'Request ID'))!==requestId)continue;items.push(rowToActivity(old.headers,r));}
    STATE.activities=items.sort((a,b)=>dateToMilliseconds(b.eventDate)-dateToMilliseconds(a.eventDate));renderActivity();
  }catch(e){if(body)body.textContent='Помилка історії: '+getErrorText(e);}
}
async function appendActivity(requestId,type,description,reminderDate='',reminderStatus=''){
  await Excel.run(async c=>{
    const targetName=type==='REMINDER'?'RBD_Reminders':'RBD_Events';
    const t=c.workbook.tables.getItemOrNullObject(targetName);t.load('isNullObject');await c.sync();
    if(t.isNullObject)throw new Error('Немає таблиці '+targetName+'. Зміни не записано.');
    const h=t.getHeaderRowRange();h.load('values');await c.sync();
    const fields=type==='REMINDER'?{
      Reminder_ID:createId('REM'),RBD_ID:requestId,Created_At:localTimestamp(),Created_By:STATE.actor,
      Recipient:STATE.actor,Reminder_Text:description,Due_At:reminderDate,Status:reminderStatus||'NEW'
    }:{Event_ID:createId('EV'),RBD_ID:requestId,Event_At:localTimestamp(),Recorded_At:localTimestamp(),Actor:STATE.actor,
      Source:'РБД 2.0 надбудова',Event_Type:type,Field:'Коментар',Old_Value:'',New_Value:description,Comment:description};
    t.rows.add(null,[h.values[0].map(x=>fields[x]??'')]);await c.sync();
  });
}
async function markReminderDone(activityId){
  try{await Excel.run(async c=>{const t=c.workbook.tables.getItemOrNullObject('RBD_Reminders');t.load('isNullObject');await c.sync();if(t.isNullObject)throw new Error('Немає RBD_Reminders');
    const h=t.getHeaderRowRange(),d=t.getDataBodyRange();h.load('values');d.load('values');await c.sync();const ix=h.values[0].indexOf('Reminder_ID'),st=h.values[0].indexOf('Status'),done=h.values[0].indexOf('Completed_At');const row=d.values.findIndex(x=>String(x[ix])===activityId);if(row<0)throw new Error('Нагадування не знайдено');
    if(st>=0)d.getCell(row,st).values=[['DONE']];if(done>=0)d.getCell(row,done).values=[[localTimestamp()]];await c.sync();});await loadActivity(STATE.activityRequestId,false);
  }catch(e){showActivityMessage(getErrorText(e),true);}
}
// Версія RBD 2.0 може не містити Sum/PlannedDate/SLA — ці поля показуються без вигаданих чисел.
