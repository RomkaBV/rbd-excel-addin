const {test,expect}=require('@playwright/test'),fs=require('node:fs');
const eml=Buffer.from('From: =?UTF-8?B?0J7Qu9GM0LPQsA==?= <olha@example.test>\r\nTo: buyer@example.test\r\nSubject: =?UTF-8?B?0KLQtdGB0YIg0LvQuNGB0YLQsA==?=\r\nDate: Fri, 23 Oct 2026 12:00:00 +0300\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary="mix"\r\n\r\n--mix\r\nContent-Type: text/plain; charset=UTF-8\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n=D0=92=D1=96=D1=82=D0=B0=D1=8E\r\n--mix\r\nContent-Type: text/plain; name="cost.txt"\r\nContent-Disposition: attachment; filename="cost.txt"\r\nContent-Transfer-Encoding: base64\r\n\r\nMTIz\r\n--mix--\r\n');
async function setup(page,api=false){
 await page.route('https://appsforoffice.microsoft.com/**',r=>r.fulfill({contentType:'application/javascript',body:'window.Office={HostType:{Excel:"Excel"},onReady:fn=>{setTimeout(()=>fn({host:"Browser"}),0)}};'}));
 await page.route('https://cdn.jsdelivr.net/**',r=>r.fulfill({contentType:'application/javascript',body:''}));
 if(api)await page.addInitScript(()=>window.RBD_SECURE_API={baseUrl:'https://rbd-api.example.test',getAccessToken:async()=>'test-token'});
 await page.goto('/taskpane.html');
 await page.evaluate(()=>{
  STATE.mode='manager';STATE.actor='Іванченко Олександр';STATE.users=[{'ПІБ':'Іванченко Олександр','Електронна пошта':'ivanchenko@example.test'}];
  STATE.statuses=['Нова','Призначена','В роботі','Пошук підрядника','Вибір постачальника','Контрактація','Закрита'].map((s,i)=>({'Назва статусу':s,'Порядок':i,'SLA, год':s==='В роботі'?9:0,'Завершальний':s==='Закрита'?'Так':'Ні'}));
  const r={id:'RBD-000003',subject:'Заміна обладнання на складі',category:'Інженерні',executor:'Іванченко Олександр',createdAt:'2026-10-20T15:42',statusSince:'2026-10-22T12:21',budget:102450,status:'Пошук підрядника'};
  STATE.requestsAll=[r,{...r,id:'RBD-000004',subject:'Інша заявка'}];
  STATE.timeline=[['Нова','Призначена','2026-10-20T16:10'],['Призначена','В роботі','2026-10-21T09:14'],['В роботі','Пошук підрядника','2026-10-22T12:21']].map(([a,b,d],i)=>({'ID заявки':r.id,'ID події':'EV-'+i,'Тип':'Статус','Дія':'Змінено статус','Було':a,'Стало':b,'Дата події':d,'Автор':'Іванченко Олександр'}));
  for(let i=0;i<8;i++)STATE.timeline.push({'ID заявки':r.id,'ID події':'NOTE-'+i,'Тип':'Коментар','Дія':'Примітка','Текст':'Уточнено специфікацію та строки виконання робіт.','Дата події':'2026-10-23T10:00','Автор':'Іванченко Олександр'});
  fillReferenceSelects();openRequest(r);
 });
}
test('fixed columns, four filters, independent editors and stable reader',async({page})=>{
 const errors=[];page.on('pageerror',e=>errors.push(e.message));await setup(page);
 await expect(page.locator('#historyWorkspaceTitle')).toHaveText('RBD-000003');await expect(page.locator('#rbdInlineMailReader')).toBeVisible();await expect(page.locator('[data-history-filter]')).toHaveCount(4);
 await page.locator('#rbdMailFileInput').setInputFiles({name:'Тест.eml',mimeType:'message/rfc822',buffer:eml});await expect(page.locator('#rbdInlineMailMeta')).toContainText('Тест листа');await expect(page.locator('#rbdInlineMailBody')).toContainText('Вітаю');
 await page.locator('[data-history-filter="status"]').click();await expect(page.locator('#chatList tr')).toHaveCount(3);await expect(page.locator('#rbdInlineMailMeta')).toContainText('Тест листа');await expect(page.locator('#historySideCount')).toHaveText('(1)');
 await page.locator('#reminderToggleButton').click();await expect(page.locator('#rbdReminderDialog')).toBeVisible();await expect(page.locator('#rbdInlineMailMeta')).toContainText('Тест листа');await page.keyboard.press('Escape');await expect(page.locator('#rbdReminderDialog')).not.toBeVisible();
 await page.locator('#rbdComposeToggle').click();await expect(page.locator('#rbdMessageDialog')).toBeVisible();await page.keyboard.press('Escape');
 await page.locator('#rbdReaderAttachments').click();await expect(page.locator('#rbdInlineMailBody')).toContainText('cost.txt');const download=page.waitForEvent('download');await page.locator('.rbd-attachment button').click();expect((await download).suggestedFilename()).toBe('cost.txt');
 await page.locator('#rbdReaderMessage').click();
 const boxes=await page.locator('.history-workspace-grid').evaluate(e=>[...e.children].map(c=>({x:c.getBoundingClientRect().x,y:c.getBoundingClientRect().y,w:c.clientWidth})));expect(boxes[0].y).toBe(boxes[1].y);expect(Math.abs(boxes[0].w-boxes[1].w)).toBeLessThan(3);
 await page.screenshot({path:'test-results/workspace-desktop.png'});expect(errors).toEqual([]);
});
test('files never bleed across requests and invalid uploads do not hide errors',async({page})=>{
 await setup(page);await page.locator('#rbdMailFileInput').setInputFiles({name:'one.eml',mimeType:'message/rfc822',buffer:eml});await expect(page.locator('#rbdInlineMailMeta')).toBeVisible();
 await page.evaluate(()=>openRequest(STATE.requestsAll[1]));await expect(page.locator('#historySideCount')).toHaveText('(0)');await expect(page.locator('#rbdInlineMailBody')).toContainText('Виберіть лист');
 await page.locator('#rbdMailFileInput').setInputFiles({name:'bad.pdf',mimeType:'application/pdf',buffer:Buffer.from('bad')});await expect(page.locator('#rbdMailUploadStatus')).toContainText('потрібен .eml/.msg');
 await page.evaluate(()=>openRequest(STATE.requestsAll[0]));await expect(page.locator('#historySideCount')).toHaveText('(1)');
});
test('MSG parser and attachment work in browser without runtime CDN',async({page})=>{
 await setup(page);const bytes=Buffer.from(fs.readFileSync('tests/fixtures/test2.msg.base64','utf8'),'base64');
 await page.locator('#rbdMailFileInput').setInputFiles({name:'test2.msg',mimeType:'application/vnd.ms-outlook',buffer:bytes});await expect(page.locator('#rbdInlineMailMeta')).toBeVisible();await expect(page.locator('#rbdReaderAttachments')).toContainText('(1)');await page.locator('#rbdReaderAttachments').click();await expect(page.locator('#rbdInlineMailBody')).toContainText('A.txt');
});
test('malformed MSG can still be downloaded',async({page})=>{
 await setup(page);await page.locator('#rbdMailFileInput').setInputFiles({name:'broken.msg',mimeType:'application/vnd.ms-outlook',buffer:Buffer.from('not a MSG')});await expect(page.locator('#rbdInlineMailBody')).toContainText('Некоректний формат MSG');await expect(page.locator('#rbdInlineMailDownload')).toBeVisible();
});
test('private server errors, retry, and late response isolation',async({page})=>{
 let slow;await page.route('https://rbd-api.example.test/**',async route=>{if(route.request().url().includes('000004'))return route.fulfill({status:403,contentType:'application/json',body:'{}'});await new Promise(r=>slow=r);return route.fulfill({contentType:'application/json',body:JSON.stringify({letters:[{id:'1234567890abc',name:'private.eml',size:42}]})});});
 await setup(page,true);await expect(page.locator('#rbdMailUploadStatus')).toContainText('Завантаження');await page.evaluate(()=>openRequest(STATE.requestsAll[1]));await expect(page.locator('#rbdMailUploadStatus')).toContainText('Недостатньо прав');slow();await expect(page.locator('#historySideCount')).toHaveText('(0)');
});
test('narrow viewport preserves both panels and notes',async({page})=>{
 await setup(page);await page.setViewportSize({width:760,height:800});await expect(page.locator('.history-main-column')).toBeVisible();await page.locator('.history-side').scrollIntoViewIfNeeded();await expect(page.locator('#rbdInlineMailReader')).toBeVisible();await page.screenshot({path:'test-results/workspace-narrow.png',fullPage:true});
});
