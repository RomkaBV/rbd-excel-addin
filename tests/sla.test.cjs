const {test}=require('node:test'),assert=require('node:assert/strict'),sla=require('../rbd-sla.js');
const c={timeZone:'Europe/Kyiv',weekdays:[1,2,3,4,5],intervals:[['09:00','13:00'],['14:00','18:00']],holidays:[]};
const statuses=['Нова','В роботі','Пошук підрядника','Призупинена','Закрита'].map((n,i)=>({'Назва статусу':n,'Порядок':i,'SLA, год':n==='В роботі'?8:0,'Завершальний':n==='Закрита'?'Так':'Ні'}));
function event(from,to,date,start){return {'ID заявки':'RBD-1','Тип':'Статус','Було':from,'Стало':to,'Дата події':date,'Початок статусу':start};}
test('work calendar excludes weekends, lunch, holidays and keeps fractions',()=>{
 assert.equal(sla.hoursBetween('2026-10-23T12:30','2026-10-26T10:30',c),6);
 assert.equal(sla.hoursBetween('2026-10-23T13:00','2026-10-23T14:00',c),0);
 assert.equal(sla.hoursBetween('2026-10-23T09:00','2026-10-26T18:00',{...c,holidays:['2026-10-26']}),8);
 assert.equal(sla.hoursBetween('2026-10-23T09:00','2026-10-23T09:01',c),1/60);
});
test('ISO instants and workbook wall time use Kyiv calendar',()=>assert.equal(sla.hoursBetween('2026-10-23T06:00:00Z','2026-10-23T15:00:00Z',c),8));
test('closed → suspended → work retains passed stages and ignores stale completedAt/statusSince',()=>{
 const events=[event('Нова','В роботі','2026-10-19T09:00'),event('В роботі','Закрита','2026-10-20T18:00'),event('Закрита','Призупинена','2026-10-22T10:00'),event('Призупинена','В роботі','2026-10-23T09:00')];
 const x=sla.calculate({id:'RBD-1',createdAt:'2026-10-19T09:00',status:'В роботі',statusSince:'2026-10-19T09:00',completedAt:'2026-10-20T18:00'},events,statuses,c,new Date('2026-10-23T09:00:00Z'));
 assert.equal(x.stages.find(s=>s.name==='Закрита').completed,true);assert.equal(x.stages.find(s=>s.name==='Пошук підрядника').count,0);
 assert.equal(x.stages.find(s=>s.name==='В роботі').count,2);assert.equal(x.current.start,'2026-10-23T09:00');assert.equal(x.current.hours,3);assert.equal(x.currentOverdue,0);assert.equal(x.totalOverdue,1);assert.equal(x.totalWorkHours,19);
});
test('each repeated visit has its own SLA; early visits never offset late visits',()=>{
 const events=[event('В роботі','Призупинена','2026-10-20T10:00'),event('Призупинена','В роботі','2026-10-21T09:00'),event('В роботі','Призупинена','2026-10-22T11:00'),event('Призупинена','В роботі','2026-10-23T09:00')];
 const x=sla.calculate({id:'RBD-1',createdAt:'2026-10-19T09:00',status:'В роботі'},events,statuses,c,new Date('2026-10-23T08:00:00Z'));
 assert.equal(x.currentOverdue,0);assert.equal(x.totalOverdue,3/8);assert.equal(x.visits.length,5);
});
test('duplicate transitions do not double duration and unknown times stay unknown',()=>{
 const e=event('Нова','В роботі','2026-10-23T09:00');const x=sla.calculate({id:'RBD-1',status:'В роботі'},[e,e],statuses,c,new Date('2026-10-23T10:00:00Z'));
 assert.equal(x.visits.length,2);assert.equal(x.visits[0].hours,null);assert.equal(x.current.hours,4);
 const y=sla.calculate({id:'RBD-1',status:'В роботі'},[],statuses,c);assert.equal(y.currentOverdue,null);assert.equal(y.totalOverdue,null);
});
test('no SLA, invalid dates, final status and backward bounds',()=>{
 const x=sla.calculate({id:'RBD-1',createdAt:'2026-10-23T09:00',status:'Закрита',statusSince:'2026-10-23T10:00'},[],statuses,c,new Date('2026-10-23T11:00:00Z'));assert.equal(x.current.hours,0);assert.equal(x.currentOverdue,null);assert.equal(x.hasSla,false);
 assert.equal(sla.hoursBetween('bad','2026-10-23',c),null);assert.equal(sla.hoursBetween('2026-10-24','2026-10-23',c),null);assert.throws(()=>sla.calendar({...c,intervals:[['09:00','14:00'],['13:00','18:00']]}));
});
