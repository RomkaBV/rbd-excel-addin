const TABLES = {
  requests: "tblRBDRequests2026",
  users: "tblRBDUsers2026",
  statuses: "tblRBDStatuses2026",
  transitions: "tblRBDTransitions2026",
  cities: "tblRBDCities2026",
  locations: "tblRBDLocations2026",
  categories: "tblRBDCategories2026",
  customers: "tblRBDCustomers2026",
  priorities: "tblRBDPriorities2026",
  commands: "tblRBDCommands2026",
  timeline: "tblRBDTimeline2026",
  notifications: "tblRBDNotifications2026",
  erpStage: "tblRBDERPStage2026",
  erpImports: "tblRBDERPImports2026",
  legacyEvents: "tblRBDEvents2026",
  assignments: "tblRBDAssignments2026"
};

const STATE = {
  activeSheet: "",
  mode: "unknown", // manager | employee | unknown
  actor: "",
  employeeId: "",
  role: "",
  users: [],
  statuses: [],
  transitions: [],
  cities: [],
  locations: [],
  categories: [],
  customers: [],
  priorities: [],
  requestsAll: [],
  requests: [],
  filtered: [],
  timeline: [],
  notifications: [],
  legacyEvents: [],
  assignments: [],
  selectedRequest: null,
  expandedStatusId: "",
  currentView: "requests",
  erpPreview: [],
  erpFileName: ""
};

Office.onReady(async info => {
  bindStaticEvents();
  if (info.host !== Office.HostType.Excel) return;
  try {
    await loadReferenceData();
    await detectCabinet();
    configureMode();
    await loadBusinessData();
    renderAll();
  } catch (e) {
    console.error(e);
    showCabinetNotice(`Помилка запуску: ${errorText(e)}`, true);
  }
});

function bindStaticEvents() {
  document.querySelectorAll("[data-view]").forEach(btn => btn.addEventListener("click", () => switchView(btn.dataset.view)));
  document.querySelectorAll("[data-close]").forEach(btn => btn.addEventListener("click", () => closeModal(btn.dataset.close)));
  ["filterStatus", "filterExecutor", "filterCity", "filterPriority"].forEach(id => byId(id)?.addEventListener("change", applyFilters));
  byId("searchInput")?.addEventListener("input", applyFilters);
  byId("refreshButton")?.addEventListener("click", refreshAll);
  byId("newRequestButton")?.addEventListener("click", openCreateModal);
  byId("saveCreateButton")?.addEventListener("click", saveCreate);
  byId("createCity")?.addEventListener("change", () => fillLocationSelect("createLocation", value("createCity"), "Оберіть локацію"));
  byId("saveAssignButton")?.addEventListener("click", saveAssignment);
  byId("chatSendButton")?.addEventListener("click", sendChatComment);
  byId("erpChooseButton")?.addEventListener("click", () => byId("erpFile")?.click());
  byId("erpFile")?.addEventListener("change", onErpFileSelected);
  byId("erpImportButton")?.addEventListener("click", importErpPreview);
  document.querySelectorAll(".dict-add").forEach(btn => btn.addEventListener("click", () => openDictionaryModal(btn.dataset.dict)));
  byId("saveDictButton")?.addEventListener("click", saveDictionaryValue);
}

async function refreshAll() {
  try {
    setBusy("refreshButton", true, "Оновлення...");
    await loadReferenceData();
    await detectCabinet();
    configureMode();
    await loadBusinessData();
    renderAll();
  } catch (e) {
    showCabinetNotice(`Помилка оновлення: ${errorText(e)}`, true);
  } finally {
    setBusy("refreshButton", false, "↻ Оновити");
  }
}

async function loadReferenceData() {
  const data = await Excel.run(async context => ({
    users: await readTable(context, TABLES.users, true),
    statuses: await readTable(context, TABLES.statuses, true),
    transitions: await readTable(context, TABLES.transitions, true),
    cities: await readTable(context, TABLES.cities, false),
    locations: await readTable(context, TABLES.locations, false),
    categories: await readTable(context, TABLES.categories, false),
    customers: await readTable(context, TABLES.customers, false),
    priorities: await readTable(context, TABLES.priorities, false)
  }));

  STATE.users = rowsToObjects(data.users).filter(x => isYes(x["Активний"]));
  STATE.statuses = rowsToObjects(data.statuses).filter(x => isYes(x["Активний"])).sort((a,b) => num(a["Порядок"]) - num(b["Порядок"]));
  STATE.transitions = rowsToObjects(data.transitions).filter(x => isYes(x["Активний"]));
  STATE.cities = rowsToObjects(data.cities).filter(x => isYes(x["Активний"]));
  STATE.locations = rowsToObjects(data.locations).filter(x => isYes(x["Активний"]));
  STATE.categories = rowsToObjects(data.categories).filter(x => isYes(x["Активний"]));
  STATE.customers = rowsToObjects(data.customers).filter(x => isYes(x["Активний"]));
  STATE.priorities = rowsToObjects(data.priorities).filter(x => isYes(x["Активний"])).sort((a,b) => num(a["Порядок"]) - num(b["Порядок"]));
  fillReferenceSelects();
}

async function detectCabinet() {
  const sheetName = await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();
    return sheet.name;
  });

  STATE.activeSheet = clean(sheetName);
  STATE.mode = "unknown";
  STATE.actor = "";
  STATE.employeeId = "";
  STATE.role = "";

  const user = STATE.users.find(x => norm(surname(x["ПІБ"])) === norm(STATE.activeSheet));
  if (!user) return;

  STATE.actor = clean(user["ПІБ"]);
  STATE.employeeId = clean(user["ID працівника"]);
  STATE.role = clean(user["Роль"]);
  STATE.mode = STATE.role.includes("Керівник") ? "manager" : "employee";
}

function configureMode() {
  const notice = byId("cabinetNotice");
  if (STATE.mode === "manager") {
    setText("pageTitle", `РБД 3.0 — ${STATE.actor}`);
    setText("pageSubtitle", "Кабінет керівника • усі заявки • ERP • контроль");
    setText("currentUserName", STATE.actor);
    setText("currentUserRole", STATE.role || "Керівник");
    setText("userAvatar", initials(STATE.actor));
    show(byId("filterExecutorField"), true);
    show(byId("createExecutorField"), true);
    show(byId("erpManagerNotice"), false);
    show(byId("dictionaryManagerNotice"), false);
    document.querySelectorAll(".dict-add").forEach(x => x.disabled = false);
    if (notice) notice.style.display = "none";
    if (byId("erpImportButton")) byId("erpImportButton").disabled = STATE.erpPreview.filter(x => x.state === "new").length === 0;
  } else if (STATE.mode === "employee") {
    setText("pageTitle", `РБД 3.0 — ${STATE.actor}`);
    setText("pageSubtitle", "Мої заявки • опрацювання • хронологія");
    setText("currentUserName", STATE.actor);
    setText("currentUserRole", STATE.role || "Виконавець");
    setText("userAvatar", initials(STATE.actor));
    show(byId("filterExecutorField"), false);
    show(byId("createExecutorField"), false);
    show(byId("erpManagerNotice"), true);
    show(byId("dictionaryManagerNotice"), true);
    document.querySelectorAll(".dict-add").forEach(x => x.disabled = true);
    if (byId("erpImportButton")) byId("erpImportButton").disabled = true;
    if (notice) notice.style.display = "none";
  } else {
    setText("pageTitle", "РБД 3.0 — Центр заявок");
    setText("pageSubtitle", "Відкрийте персональну вкладку працівника");
    setText("currentUserName", "РБД");
    setText("currentUserRole", "Кабінет не визначено");
    setText("userAvatar", "Р");
    show(byId("filterExecutorField"), false);
    show(byId("createExecutorField"), false);
    show(byId("erpManagerNotice"), true);
    show(byId("dictionaryManagerNotice"), true);
    document.querySelectorAll(".dict-add").forEach(x => x.disabled = true);
    if (byId("erpImportButton")) byId("erpImportButton").disabled = true;
    showCabinetNotice(`Активна вкладка «${STATE.activeSheet}» не відповідає жодному активному працівнику. Назва персональної вкладки має дорівнювати прізвищу працівника з довідника «Працівники».`, false);
  }
}

async function loadBusinessData() {
  const data = await Excel.run(async context => ({
    requests: await readTable(context, TABLES.requests, true),
    timeline: await readTable(context, TABLES.timeline, false),
    notifications: await readTable(context, TABLES.notifications, false),
    legacyEvents: await readTable(context, TABLES.legacyEvents, false),
    assignments: await readTable(context, TABLES.assignments, false)
  }));

  STATE.requestsAll = rowsToObjects(data.requests).map(requestFromObject).filter(x => x.id);
  STATE.timeline = rowsToObjects(data.timeline);
  STATE.notifications = rowsToObjects(data.notifications);
  STATE.legacyEvents = rowsToObjects(data.legacyEvents);
  STATE.assignments = rowsToObjects(data.assignments);

  if (STATE.mode === "manager") STATE.requests = STATE.requestsAll.slice();
  else if (STATE.mode === "employee") STATE.requests = STATE.requestsAll.filter(r => sameName(r.executor, STATE.actor));
  else STATE.requests = [];

  STATE.expandedStatusId = "";
  if (STATE.selectedRequest) STATE.selectedRequest = STATE.requestsAll.find(r => r.id === STATE.selectedRequest.id) || STATE.selectedRequest;
  updateLastUpdated();
}

async function readTable(context, name, required) {
  const table = context.workbook.tables.getItemOrNullObject(name);
  table.load("isNullObject");
  await context.sync();
  if (table.isNullObject) {
    if (required) throw new Error(`Не знайдено таблицю ${name}. Запустіть 01_RBD_Setup_New.`);
    return {headers: [], rows: []};
  }
  const header = table.getHeaderRowRange();
  header.load("values");
  table.rows.load("items");
  await context.sync();
  const headers = header.values[0] || [];
  if (!table.rows.items.length) return {headers, rows: []};
  const body = table.getDataBodyRange();
  body.load("values");
  await context.sync();
  return {headers, rows: body.values};
}

function rowsToObjects(data) {
  if (!data?.headers?.length) return [];
  return data.rows.map(row => {
    const o = {};
    data.headers.forEach((h,i) => o[String(h)] = row[i] ?? "");
    return o;
  });
}

function requestFromObject(x) {
  return {
    id: clean(x["ID заявки"]), erpId: clean(x["ID ERP"]), subject: clean(x["Тема заявки"]), description: clean(x["Опис заявки"]),
    customer: clean(x["Замовник"]), city: clean(x["Місто"]), location: clean(x["Локація"]), category: clean(x["Категорія"]), priority: clean(x["Пріоритет"]),
    status: clean(x["Статус"]), executor: clean(x["Виконавець"]), createdBy: clean(x["Створив"]), createdAt: x["Дата створення"], updatedAt: x["Дата оновлення"],
    version: num(x["Версія"]), source: clean(x["Джерело"]), statusSince: x["Дата зміни статусу"], completedAt: x["Дата завершення"]
  };
}

function renderAll() {
  updateKpis();
  applyFilters();
  renderEmployees();
  renderAnalytics();
  renderDictionaries();
  renderErpPreview();
  if (STATE.selectedRequest && byId("requestModal")?.classList.contains("open")) renderRequestModal(STATE.selectedRequest);
}

function fillReferenceSelects() {
  fillSelect("filterStatus", STATE.statuses.map(x => clean(x["Назва статусу"])), "Усі статуси");
  fillSelect("filterExecutor", STATE.users.map(x => clean(x["ПІБ"])), "Усі виконавці");
  fillSelect("filterCity", STATE.cities.map(x => clean(x["Назва"])), "Усі міста");
  fillSelect("filterPriority", STATE.priorities.map(x => clean(x["Назва"])), "Усі пріоритети");
  fillSelect("createCity", STATE.cities.map(x => clean(x["Назва"])), "Оберіть місто");
  fillSelect("createCategory", STATE.categories.map(x => clean(x["Назва"])), "Оберіть категорію");
  fillSelect("createCustomer", STATE.customers.map(x => clean(x["Назва"])), "Оберіть замовника");
  fillSelect("createPriority", STATE.priorities.map(x => clean(x["Назва"])), "Оберіть пріоритет");
  fillSelect("createExecutor", STATE.users.map(x => clean(x["ПІБ"])), "Без виконавця");
  fillSelect("assignExecutor", STATE.users.map(x => clean(x["ПІБ"])), "Оберіть виконавця");
  fillLocationSelect("createLocation", value("createCity"), "Оберіть локацію");
}

function fillLocationSelect(id, city, placeholder) {
  const list = STATE.locations.filter(x => !city || sameName(x["Місто"], city)).map(x => clean(x["Назва"]));
  fillSelect(id, list, placeholder);
}

function applyFilters() {
  const search = norm(value("searchInput"));
  const status = value("filterStatus");
  const executor = value("filterExecutor");
  const city = value("filterCity");
  const priority = value("filterPriority");

  STATE.filtered = STATE.requests.filter(r => {
    if (status && r.status !== status) return false;
    if (STATE.mode === "manager" && executor && !sameName(r.executor, executor)) return false;
    if (city && r.city !== city) return false;
    if (priority && r.priority !== priority) return false;
    if (search && !norm([r.id,r.erpId,r.subject,r.description,r.customer,r.city,r.location,r.category,r.priority,r.status,r.executor].join(" ")).includes(search)) return false;
    return true;
  });
  renderRequests();
  setText("requestsSummary", `Показано ${STATE.filtered.length} із ${STATE.requests.length} заявок`);
}

function updateKpis() {
  const active = STATE.requests.filter(r => !isFinalStatus(r.status));
  const overdue = active.filter(isOverdue);
  const work = active.filter(r => !["Нова","Розподілена","Призупинена"].includes(r.status));
  const cutoff = Date.now() - 30*86400000;
  const closed = STATE.requests.filter(r => isFinalStatus(r.status) && dateMs(r.completedAt) >= cutoff);
  setText("kpiActive", active.length); setText("kpiActiveSub", `${STATE.requests.filter(r=>!r.executor).length} без виконавця`);
  setText("kpiOverdue", overdue.length); setText("kpiOverdueSub", overdue.length ? "Потребують уваги" : "Прострочень немає");
  setText("kpiWork", work.length); setText("kpiWorkSub", "Активне виконання");
  setText("kpiClosed", closed.length); setText("kpiClosedSub", "За останні 30 днів");
}

function renderRequests() {
  const body = byId("requestsBody"); if (!body) return;
  if (!STATE.filtered.length) { body.innerHTML = `<tr><td colspan="13"><div class="empty">Заявок за обраними умовами немає</div></td></tr>`; return; }
  let html = "";
  for (const r of STATE.filtered) {
    const st = getStatus(r.status); const bg = validHex(st?.["Колір фону"]) ? st["Колір фону"] : "#ffffff"; const fg = validHex(st?.["Колір тексту"]) ? st["Колір тексту"] : "#173652";
    const hours = currentStatusHours(r); const sla = num(st?.["SLA, год"]); const overdue = isOverdue(r);
    html += `<tr class="request-row" style="background:${overdue?"#ffb5ba":bg}22">
      <td><div class="request-id">${esc(r.id)}</div></td><td>${esc(r.erpId||"—")}</td><td><div class="ellipsis" title="${esc(r.subject)}">${esc(r.subject||"—")}</div></td>
      <td>${esc(r.category||"—")}</td><td>${esc(r.city||"—")}</td><td><div class="ellipsis" title="${esc(r.customer)}">${esc(r.customer||"—")}</div></td><td>${esc(r.executor||"—")}</td>
      <td>${priorityPill(r.priority)}</td><td><button class="status-pill" data-status-open="${esc(r.id)}" style="background:${bg};color:${fg};border:0"><span class="dot"></span>${esc(r.status||"—")}</button></td>
      <td class="${overdue?"overdue":""}">${hours===null?"—":`${hours} год.`}</td><td>${sla?`${sla} год.`:"—"}</td><td>${r.version||0}</td>
      <td><div class="action-row"><button class="icon-btn" data-open-request="${esc(r.id)}" title="Деталі та історія">◉</button>${STATE.mode==="manager"?`<button class="icon-btn" data-assign="${esc(r.id)}" title="Призначити">♙</button>`:""}</div></td></tr>`;
    if (STATE.expandedStatusId === r.id) html += statusPanelHtml(r);
  }
  body.innerHTML = html;
  bindRequestEvents();
}

function statusPanelHtml(r) {
  const allowed = allowedNextStatuses(r.status);
  if (!allowed.length) return `<tr class="status-panel-row"><td colspan="13"><div class="status-panel"><span class="muted">Для статусу «${esc(r.status)}» немає дозволених наступних переходів.</span></div></td></tr>`;
  const buttons = allowed.map(s => { const st=getStatus(s);const bg=validHex(st?.["Колір фону"])?st["Колір фону"]:"#e9eef2";const fg=validHex(st?.["Колір тексту"])?st["Колір тексту"]:"#173652";return `<button class="status-option" data-status-change="${esc(s)}" data-request-id="${esc(r.id)}" style="background:${bg};color:${fg}">${esc(s)}</button>`; }).join("");
  return `<tr class="status-panel-row"><td colspan="13"><div class="status-panel"><div style="font-size:9px;font-weight:750;margin-bottom:7px">Змінити статус: ${esc(r.status)}</div><div class="status-options">${buttons}<button class="status-option" data-status-close="1">× Закрити панель</button></div></div></td></tr>`;
}

function bindRequestEvents() {
  document.querySelectorAll("[data-open-request]").forEach(b => b.addEventListener("click", () => openRequest(findRequest(b.dataset.openRequest))));
  document.querySelectorAll("[data-assign]").forEach(b => b.addEventListener("click", () => openAssignment(findRequest(b.dataset.assign))));
  document.querySelectorAll("[data-status-open]").forEach(b => b.addEventListener("click", () => { const id=b.dataset.statusOpen; STATE.expandedStatusId = STATE.expandedStatusId===id ? "" : id; renderRequests(); }));
  document.querySelectorAll("[data-status-close]").forEach(b => b.addEventListener("click", () => { STATE.expandedStatusId=""; renderRequests(); }));
  document.querySelectorAll("[data-status-change]").forEach(b => b.addEventListener("click", async () => {
    const r=findRequest(b.dataset.requestId); const next=b.dataset.statusChange; if(!r)return;
    let comment="";
    if (next === "Призупинена") { comment = window.prompt("Причина призупинення:", "") || ""; if(!comment.trim()) return toast("Для призупинення потрібна причина"); }
    if (next === "Закрита" && !window.confirm(`Закрити заявку ${r.id}?`)) return;
    try { await sendCommand({action: next==="Закрита"?"CLOSE":"STATUS_CHANGE", requestId:r.id, status:next, comment, expectedVersion:r.version}); toast(`Зміна статусу ${r.status} → ${next} передана в чергу`); STATE.expandedStatusId=""; renderRequests(); }
    catch(e){toast(`Помилка: ${errorText(e)}`);}
  }));
}

function openCreateModal() {
  if (STATE.mode === "unknown") return toast("Спочатку відкрийте персональну вкладку працівника");
  ["createSubject","createDescription","createComment"].forEach(id => setInput(id,""));
  setInput("createCity",""); fillLocationSelect("createLocation","","Оберіть локацію"); setInput("createCategory",""); setInput("createCustomer","");
  setInput("createPriority", STATE.priorities.find(x=>clean(x["Назва"])==="Середній") ? "Середній" : "");
  setInput("createExecutor", STATE.mode === "employee" ? STATE.actor : ""); clearMessage("createMessage"); openModal("createModal");
}

async function saveCreate() {
  const subject=value("createSubject"); if(!subject) return showMessage("createMessage","error","Вкажіть тему заявки.");
  const executor=STATE.mode==="employee"?STATE.actor:value("createExecutor");
  try {
    setBusy("saveCreateButton",true,"Створюємо...");
    await sendCommand({action:"CREATE",subject,description:value("createDescription"),executor,city:value("createCity"),location:value("createLocation"),category:value("createCategory"),customer:value("createCustomer"),priority:value("createPriority"),comment:value("createComment")});
    showMessage("createMessage","success","Заявку передано в центральну чергу команд."); toast("Заявку передано на обробку");
  } catch(e){showMessage("createMessage","error",errorText(e));} finally{setBusy("saveCreateButton",false,"Створити");}
}

function openAssignment(r) {
  if (!r) return; if (STATE.mode !== "manager") return toast("Призначення іншого виконавця доступне керівнику");
  STATE.selectedRequest=r; setText("assignTitle",`Призначення — ${r.id}`); setInput("assignExecutor",r.executor||""); setInput("assignComment",""); clearMessage("assignMessage"); openModal("assignModal");
}

async function saveAssignment() {
  const r=STATE.selectedRequest;if(!r)return;const executor=value("assignExecutor");if(!executor)return showMessage("assignMessage","error","Оберіть виконавця.");
  try{setBusy("saveAssignButton",true,"Призначаємо...");await sendCommand({action:"ASSIGN",requestId:r.id,executor,comment:value("assignComment"),expectedVersion:r.version});showMessage("assignMessage","success","Команду передано на обробку.");toast("Призначення передано на обробку");}
  catch(e){showMessage("assignMessage","error",errorText(e));}finally{setBusy("saveAssignButton",false,"Призначити");}
}

async function sendCommand(p) {
  if (STATE.mode === "unknown") throw new Error("Не визначено кабінет користувача");
  const row = {
    "ID команди": createId("CMD"), "Дія": p.action||"", "ID заявки": p.requestId||"", "ID ERP": p.erpId||"", "ID імпорту": p.importId||"",
    "Тема заявки": p.subject||"", "Опис заявки": p.description||"", "Виконавець": p.executor||"", "Статус": p.status||"", "Місто": p.city||"", "Локація": p.location||"", "Категорія": p.category||"", "Замовник": p.customer||"", "Пріоритет": p.priority||"", "Коментар": p.comment||"", "Очікувана версія": p.expectedVersion??"",
    "Користувач": STATE.actor, "Джерело": p.source||"Надбудова", "Дата команди": localTimestamp(), "Стан обробки": "Нова", "Результат": ""
  };
  await appendObjectRow(TABLES.commands,row,true);
  return row["ID команди"];
}

function allowedNextStatuses(currentName) {
  const current=getStatus(currentName); if(!current)return[]; const fromId=clean(current["ID статусу"]); const ids=new Set(STATE.transitions.filter(x=>clean(x["Із статусу (ID)"])===fromId).map(x=>clean(x["До статусу (ID)"]))); return STATE.statuses.filter(x=>ids.has(clean(x["ID статусу"]))).map(x=>clean(x["Назва статусу"]));
}
function getStatus(name){return STATE.statuses.find(x=>clean(x["Назва статусу"])===clean(name));}
function isFinalStatus(name){const s=getStatus(name);return s?isYes(s["Завершальний"]):["Закрита","Скасована","Виконана"].includes(name);}
function currentStatusHours(r){if(isFinalStatus(r.status)||r.status==="Призупинена")return null;const ms=dateMs(r.statusSince);if(!ms)return null;return Math.max(0,Math.round(((Date.now()-ms)/3600000)*10)/10);}
function isOverdue(r){if(isFinalStatus(r.status)||r.status==="Призупинена")return false;const sla=num(getStatus(r.status)?.["SLA, год"]);const h=currentStatusHours(r);return sla>0&&h!==null&&h>sla;}

function openRequest(r) { if(!r)return; STATE.selectedRequest=r; renderRequestModal(r); openModal("requestModal"); }
function renderRequestModal(r) {
  setText("requestModalTitle", `${r.id} — ${r.subject||"Заявка"}`);
  const h=currentStatusHours(r); const sla=num(getStatus(r.status)?.["SLA, год"]);
  const rows=[["ID ERP",r.erpId],["Тема",r.subject],["Опис",r.description],["Замовник",r.customer],["Місто",r.city],["Локація",r.location],["Категорія",r.category],["Пріоритет",r.priority],["Статус",r.status],["Виконавець",r.executor],["Створив",r.createdBy],["Створено",fmtDateTime(r.createdAt)],["Час у статусі",h===null?"—":`${h} год.`],["SLA",sla?`${sla} год.`:"—"],["Версія",r.version],["Джерело",r.source],["Завершено",fmtDateTime(r.completedAt)]];
  byId("requestSummary").innerHTML=rows.map(x=>`<div class="summary-label">${esc(x[0])}</div><div class="summary-value">${esc(String(x[1]||"—"))}</div>`).join("");
  renderChat(r.id); setInput("chatInput",""); clearMessage("chatMessage");
}

function buildConversation(requestId) {
  const items=[];
  for (const x of STATE.timeline.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:clean(x["ID події"]),date:x["Дата події"],author:clean(x["Автор"]),type:clean(x["Тип"]),action:clean(x["Дія"]),text:clean(x["Текст"]),before:clean(x["Було"]),after:clean(x["Стало"]),channel:clean(x["Канал"]),recipient:clean(x["Отримувач"]),delivery:clean(x["Статус доставки"]),source:clean(x["Джерело"])});
  for (const x of STATE.legacyEvents.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:`LEG-${clean(x["ID події"]||x["ID команди"])}`,date:x["Дата події"],author:clean(x["Користувач"]),type:"Подія",action:clean(x["Дія"]),text:clean(x["Коментар"]),before:clean(x["Було"]),after:clean(x["Стало"]),source:clean(x["Джерело"])});
  for (const x of STATE.assignments.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:`AS-${clean(x["ID призначення"])}`,date:x["Дата призначення"],author:clean(x["Призначив"]),type:"Призначення",action:"Зміна виконавця",text:clean(x["Причина"]),before:clean(x["Попередній виконавець"]),after:clean(x["Новий виконавець"]),source:"Excel"});
  for (const x of STATE.notifications.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:`NT-${clean(x["ID сповіщення"])}`,date:x["Дата відправки"]||x["Дата створення"],author:"Система",type:"Сповіщення",action:clean(x["Канал"]||"Повідомлення"),text:clean(x["Текст"]||x["Тема"]),channel:clean(x["Канал"]),recipient:clean(x["Отримувач"]),delivery:clean(x["Статус"]),source:"Система"});
  const seen=new Set(); return items.filter(x=>{const k=x.id||`${x.date}|${x.author}|${x.action}|${x.text}`;if(seen.has(k))return false;seen.add(k);return true;}).sort((a,b)=>dateMs(a.date)-dateMs(b.date));
}

function renderChat(requestId) {
  const list=byId("chatList"); if(!list)return; const items=buildConversation(requestId);
  if(!items.length){list.innerHTML=`<div class="empty">Історія заявки ще порожня</div>`;return;}
  list.innerHTML=items.map(x=>{
    const isComment=x.type==="Коментар"||x.action==="Коментар"; const mine=sameName(x.author,STATE.actor); const change=(x.before||x.after)?`<div class="chat-status">${x.before?esc(x.before):""}${x.before||x.after?" → ":""}${x.after?esc(x.after):""}</div>`:""; const delivery=x.type==="Сповіщення"?`<div class="chat-status">${esc(x.channel||"Повідомлення")} • ${esc(x.recipient||"—")} • ${esc(x.delivery||"Очікує")}</div>`:"";
    if(!isComment) return `<div class="chat-system"><strong>${esc(x.action||x.type||"Подія")}</strong>${x.author?` • ${esc(x.author)}`:""}${x.text?`<br>${esc(x.text)}`:""}${change}${delivery}<div style="margin-top:3px">${esc(fmtDateTime(x.date))}</div></div>`;
    return `<div class="chat-item ${mine?"me":""}"><div class="bubble"><div class="chat-meta">${esc(x.author||"Користувач")} • ${esc(fmtDateTime(x.date))}</div><div class="chat-text">${esc(x.text||x.after||"Коментар")}</div>${delivery}</div></div>`;
  }).join("");
  list.scrollTop=list.scrollHeight;
}

async function sendChatComment() {
  const r=STATE.selectedRequest;if(!r)return;const comment=value("chatInput");if(!comment)return;
  try{setBusy("chatSendButton",true,"...");await sendCommand({action:"COMMENT",requestId:r.id,comment,expectedVersion:r.version});showMessage("chatMessage","success","Коментар передано в чергу. Після обробки він з’явиться у хронології.");setInput("chatInput","");toast("Коментар передано на обробку");}
  catch(e){showMessage("chatMessage","error",errorText(e));}finally{setBusy("chatSendButton",false,"Надіслати");}
}

function renderEmployees(){const body=byId("employeesBody");if(!body)return;let users=STATE.users;if(STATE.mode==="employee")users=users.filter(x=>sameName(x["ПІБ"],STATE.actor));const cutoff=Date.now()-30*86400000;body.innerHTML=users.length?users.map(u=>{const n=clean(u["ПІБ"]);const req=STATE.requestsAll.filter(r=>sameName(r.executor,n));const active=req.filter(r=>!isFinalStatus(r.status));const work=active.filter(r=>!["Нова","Розподілена","Призупинена"].includes(r.status));const overdue=active.filter(isOverdue);const closed=req.filter(r=>isFinalStatus(r.status)&&dateMs(r.completedAt)>=cutoff);return `<tr><td class="bold">${esc(n)}</td><td>${esc(u["Роль"]||"—")}</td><td>${active.length}</td><td>${work.length}</td><td class="${overdue.length?"overdue":""}">${overdue.length}</td><td>${closed.length}</td></tr>`;}).join(""):`<tr><td colspan="6"><div class="empty">Немає даних</div></td></tr>`;}
function renderAnalytics(){const req=STATE.requests;setText("analyticsTotal",req.length);setText("analyticsUnassigned",req.filter(r=>!r.executor).length);setText("analyticsOverdue",req.filter(isOverdue).length);setText("analyticsErp",req.filter(r=>norm(r.source)==="erp"||r.erpId).length);renderBars("statusAnalytics",groupCount(req,r=>r.status||"Без статусу"),STATE.statuses.map(x=>clean(x["Назва статусу"])));renderBars("priorityAnalytics",groupCount(req,r=>r.priority||"Без пріоритету"),STATE.priorities.map(x=>clean(x["Назва"])));}
function renderBars(id,map,order){const el=byId(id);if(!el)return;const entries=[];for(const key of order)if(map[key])entries.push([key,map[key]]);for(const [k,v] of Object.entries(map))if(!entries.some(x=>x[0]===k))entries.push([k,v]);const max=Math.max(1,...entries.map(x=>x[1]));el.innerHTML=entries.length?entries.map(([k,v])=>`<div class="bar-row"><div class="ellipsis" title="${esc(k)}">${esc(k)}</div><div class="bar-track"><div class="bar-fill" style="width:${Math.round(v/max*100)}%"></div></div><div class="right bold">${v}</div></div>`).join(""):`<div class="empty">Немає даних</div>`;}

function renderDictionaries(){renderSimpleDictionary("citiesBody",STATE.cities,["ID міста","Назва","Активний"]);renderSimpleDictionary("categoriesBody",STATE.categories,["ID категорії","Назва","Активний"]);renderSimpleDictionary("customersBody",STATE.customers,["ID замовника","Назва","Електронна пошта"]);renderSimpleDictionary("locationsBody",STATE.locations,["ID локації","Місто","Назва","Адреса"]);}
function renderSimpleDictionary(id,rows,fields){const body=byId(id);if(!body)return;body.innerHTML=rows.length?rows.map(r=>`<tr>${fields.map(f=>`<td><div class="ellipsis">${esc(r[f]||"—")}</div></td>`).join("")}</tr>`).join(""):`<tr><td colspan="${fields.length}"><div class="empty">Немає даних</div></td></tr>`;}

let DICT_TYPE="";
function openDictionaryModal(type){if(STATE.mode!=="manager")return toast("Додавати довідники може лише керівник");DICT_TYPE=type;clearMessage("dictMessage");const titles={city:"Нове місто",category:"Нова категорія",customer:"Новий замовник",location:"Нова локація"};setText("dictTitle",titles[type]||"Нове значення");const box=byId("dictFields");if(type==="location")box.innerHTML=`<div class="field"><label>Місто</label><select id="dictCity"></select></div><div class="field" style="margin-top:10px"><label>Назва локації</label><input id="dictName"></div><div class="field" style="margin-top:10px"><label>Адреса</label><input id="dictAddress"></div>`;else if(type==="customer")box.innerHTML=`<div class="field"><label>Назва</label><input id="dictName"></div><div class="field" style="margin-top:10px"><label>Email</label><input id="dictEmail"></div>`;else box.innerHTML=`<div class="field"><label>Назва</label><input id="dictName"></div>`;if(type==="location")fillSelect("dictCity",STATE.cities.map(x=>clean(x["Назва"])),"Оберіть місто");openModal("dictModal");}
async function saveDictionaryValue(){if(STATE.mode!=="manager")return;const name=value("dictName");if(!name)return showMessage("dictMessage","error","Вкажіть назву.");const action={city:"ADD_CITY",category:"ADD_CATEGORY",customer:"ADD_CUSTOMER",location:"ADD_LOCATION"}[DICT_TYPE];const p={action};if(DICT_TYPE==="city")p.city=name;if(DICT_TYPE==="category")p.category=name;if(DICT_TYPE==="customer"){p.customer=name;p.comment=value("dictEmail");}if(DICT_TYPE==="location"){p.city=value("dictCity");p.location=name;p.description=value("dictAddress");}try{setBusy("saveDictButton",true,"Додаємо...");await sendCommand(p);showMessage("dictMessage","success","Команду передано на обробку.");toast("Значення передано на обробку");}catch(e){showMessage("dictMessage","error",errorText(e));}finally{setBusy("saveDictButton",false,"Додати");}}

async function onErpFileSelected(event){const file=event.target.files?.[0];if(!file)return;STATE.erpFileName=file.name;setText("erpFileName",file.name);try{const data=await file.arrayBuffer();const wb=XLSX.read(data,{type:"array",cellDates:false});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});STATE.erpPreview=mapErpRows(rows);renderErpPreview();configureMode();}catch(e){STATE.erpPreview=[];renderErpPreview();toast(`Не вдалося прочитати файл: ${errorText(e)}`);}}
function mapErpRows(rows){if(!rows?.length)return[];const headers=rows[0].map(x=>normHeader(x));const getIndex=names=>{for(const n of names){const i=headers.indexOf(normHeader(n));if(i>=0)return i;}return-1;};const idx={erp:getIndex(["ID ERP","ERP ID","Зовнішній ID","External ID","ID","Номер заявки","Заявка"]),subject:getIndex(["Тема заявки","Тема","Назва","Предмет"]),description:getIndex(["Опис заявки","Опис"]),customer:getIndex(["Замовник","Ініціатор","Автор"]),city:getIndex(["Місто"]),location:getIndex(["Локація","Адреса","Склад","Об'єкт"]),category:getIndex(["Категорія"]),priority:getIndex(["Пріоритет"])};const val=(row,i)=>i>=0?clean(row[i]):"";const existing=new Set(STATE.requestsAll.map(r=>norm(r.erpId)).filter(Boolean));const local=new Set();const out=[];for(let i=1;i<rows.length;i++){const row=rows[i];if(!row.some(x=>clean(x)))continue;const r={erpId:val(row,idx.erp),subject:val(row,idx.subject),description:val(row,idx.description),customer:val(row,idx.customer),city:val(row,idx.city),location:val(row,idx.location),category:val(row,idx.category),priority:val(row,idx.priority)};let state="new",result="Нова";if(!r.erpId){state="error";result="Немає ID ERP";}else if(existing.has(norm(r.erpId))||local.has(norm(r.erpId))){state="duplicate";result="Вже існує";}else {local.add(norm(r.erpId));const dictError=validateErpDictionaries(r);if(dictError){state="error";result=dictError;}}out.push({...r,state,result});}return out;}
function validateErpDictionaries(r){if(r.city&&!STATE.cities.some(x=>sameName(x["Назва"],r.city)))return"Невідоме місто";if(r.category&&!STATE.categories.some(x=>sameName(x["Назва"],r.category)))return"Невідома категорія";if(r.customer&&!STATE.customers.some(x=>sameName(x["Назва"],r.customer)))return"Невідомий замовник";if(r.priority&&!STATE.priorities.some(x=>sameName(x["Назва"],r.priority)))return"Невідомий пріоритет";if(r.location&&!STATE.locations.some(x=>sameName(x["Назва"],r.location)&&(!r.city||sameName(x["Місто"],r.city))))return"Локація не відповідає місту";return"";}
function renderErpPreview(){const body=byId("erpPreviewBody");if(!body)return;const total=STATE.erpPreview.length,newCount=STATE.erpPreview.filter(x=>x.state==="new").length,dup=STATE.erpPreview.filter(x=>x.state==="duplicate").length,err=STATE.erpPreview.filter(x=>x.state==="error").length;setText("erpRowsCount",total);setText("erpNewCount",newCount);setText("erpDuplicateCount",dup);setText("erpErrorCount",err);setText("erpPreviewSub",STATE.erpFileName?`${STATE.erpFileName} • ${total} рядків`:"Оберіть файл");if(!total){body.innerHTML=`<tr><td colspan="9"><div class="empty">Файл ще не вибрано</div></td></tr>`;}else body.innerHTML=STATE.erpPreview.map(r=>`<tr><td>${esc(r.erpId||"—")}</td><td><div class="ellipsis">${esc(r.subject||"—")}</div></td><td><div class="ellipsis">${esc(r.description||"—")}</div></td><td>${esc(r.customer||"—")}</td><td>${esc(r.city||"—")}</td><td>${esc(r.location||"—")}</td><td>${esc(r.category||"—")}</td><td>${esc(r.priority||"—")}</td><td class="${r.state==="error"?"overdue":r.state==="duplicate"?"muted":"bold"}">${esc(r.result)}</td></tr>`).join("");if(byId("erpImportButton"))byId("erpImportButton").disabled=STATE.mode!=="manager"||newCount===0;}
async function importErpPreview(){if(STATE.mode!=="manager")return toast("ERP-імпорт доступний лише керівнику");const rows=STATE.erpPreview;if(!rows.length)return;const importId=createId("IMP");const timestamp=localTimestamp();const newRows=rows.filter(x=>x.state==="new");try{setBusy("erpImportButton",true,"Імпорт...");await Excel.run(async context=>{const stage=context.workbook.tables.getItem(TABLES.erpStage),commands=context.workbook.tables.getItem(TABLES.commands),imports=context.workbook.tables.getItem(TABLES.erpImports);const hs=stage.getHeaderRowRange(),hc=commands.getHeaderRowRange(),hi=imports.getHeaderRowRange();hs.load("values");hc.load("values");hi.load("values");await context.sync();const sh=hs.values[0],ch=hc.values[0],ih=hi.values[0];const stageRows=rows.map(r=>{const o={"ID імпорту":importId,"Зовнішній ID":r.erpId,"Джерело":STATE.erpFileName||"ERP","Тема заявки":r.subject,"Опис заявки":r.description,"Результат перевірки":r.state==="new"?"Передано в чергу":r.result,"ID заявки РБД":""};return sh.map(h=>o[String(h)]??"");});if(stageRows.length)stage.rows.add(null,stageRows);const cmdRows=newRows.map(r=>{const o={"ID команди":createId("CMD"),"Дія":"CREATE","ID заявки":"","ID ERP":r.erpId,"ID імпорту":importId,"Тема заявки":r.subject||`ERP ${r.erpId}`,"Опис заявки":r.description,"Виконавець":"","Статус":"","Місто":r.city,"Локація":r.location,"Категорія":r.category,"Замовник":r.customer,"Пріоритет":r.priority||"Середній","Коментар":`Імпорт з ${STATE.erpFileName}`,"Очікувана версія":"","Користувач":STATE.actor,"Джерело":"ERP","Дата команди":timestamp,"Стан обробки":"Нова","Результат":""};return ch.map(h=>o[String(h)]??"");});if(cmdRows.length)commands.rows.add(null,cmdRows);const log={"ID імпорту":importId,"Дата імпорту":timestamp,"Імпортував":STATE.actor,"Назва файлу":STATE.erpFileName,"Нових":newRows.length,"Оновлених":0,"Помилок":rows.filter(x=>x.state==="error").length};imports.rows.add(null,[ih.map(h=>log[String(h)]??"")]);await context.sync();});toast(`ERP: ${newRows.length} нових заявок передано в центральну чергу`);clearErpPreview();}catch(e){console.error(e);toast(`Помилка ERP-імпорту: ${errorText(e)}`);}finally{setBusy("erpImportButton",false,"Імпортувати в РБД");configureMode();}}
function clearErpPreview(){STATE.erpPreview=[];STATE.erpFileName="";if(byId("erpFile"))byId("erpFile").value="";setText("erpFileName","Файл не вибрано");renderErpPreview();}

async function appendObjectRow(tableName,obj,required){await Excel.run(async context=>{const table=context.workbook.tables.getItemOrNullObject(tableName);table.load("isNullObject");await context.sync();if(table.isNullObject){if(required)throw new Error(`Не знайдено таблицю ${tableName}`);return;}const h=table.getHeaderRowRange();h.load("values");await context.sync();const headers=h.values[0];table.rows.add(null,[headers.map(x=>obj[String(x)]??"")]);await context.sync();});}

function switchView(view){STATE.currentView=view;document.querySelectorAll(".view").forEach(x=>x.classList.toggle("active",x.id===`view-${view}`));document.querySelectorAll("[data-view]").forEach(x=>x.classList.toggle("active",x.dataset.view===view));if(view==="erp"&&STATE.mode!=="manager")show(byId("erpManagerNotice"),true);}
function openModal(id){byId(id)?.classList.add("open");}function closeModal(id){byId(id)?.classList.remove("open");}
function showCabinetNotice(text,isError){const n=byId("cabinetNotice");if(!n)return;n.textContent=text;n.className=`notice ${isError?"error":"warn"}`;n.style.display="block";}
function findRequest(id){return STATE.requestsAll.find(r=>r.id===id);}
function priorityPill(name){const colors={"Низький":["#D5F2DD","#17653B"],"Середній":["#FFF0B8","#785300"],"Високий":["#FFD4D7","#9C2330"],"Критичний":["#F5ABB2","#7D1420"]};const c=colors[name]||["#e8eef2","#53697a"];return `<span class="priority-pill" style="background:${c[0]};color:${c[1]}">${esc(name||"—")}</span>`;}
function groupCount(rows,keyFn){const m={};for(const r of rows){const k=keyFn(r);m[k]=(m[k]||0)+1;}return m;}
function fillSelect(id,items,placeholder){const el=byId(id);if(!el)return;const old=el.value;el.innerHTML="";const first=document.createElement("option");first.value="";first.textContent=placeholder;el.appendChild(first);for(const x of items.filter(Boolean)){const o=document.createElement("option");o.value=x;o.textContent=x;el.appendChild(o);}if(items.includes(old))el.value=old;}
function byId(id){return document.getElementById(id);}function value(id){return clean(byId(id)?.value??"");}function setInput(id,v){const e=byId(id);if(e)e.value=v??"";}function setText(id,v){const e=byId(id);if(e)e.textContent=String(v??"");}function show(el,yes){if(el)el.style.display=yes?"":"none";}function setBusy(id,busy,text){const b=byId(id);if(!b)return;b.disabled=busy;b.textContent=text;}function showMessage(id,type,text){const e=byId(id);if(!e)return;e.className=`message ${type}`;e.textContent=text;}function clearMessage(id){const e=byId(id);if(!e)return;e.className="message";e.textContent="";}
function toast(text){const t=byId("toast");if(!t)return;t.textContent=text;t.classList.add("show");setTimeout(()=>t.classList.remove("show"),3000);}
function clean(v){return String(v??"").trim();}function norm(v){return clean(v).replace(/\u00A0/g," ").replace(/\s+/g," ").toLocaleLowerCase("uk-UA");}function normHeader(v){return norm(v).replace(/[._-]+/g," ");}function sameName(a,b){return norm(a)===norm(b);}function surname(name){return clean(name).split(/\s+/)[0]||"";}function isYes(v){return ["так","yes","true","1"].includes(norm(v));}function num(v){const n=Number(String(v??"").replace(/\s/g,"").replace(",","."));return Number.isFinite(n)?n:0;}function validHex(v){return /^#[0-9a-fA-F]{6}$/.test(clean(v));}
function initials(name){const p=clean(name).replace(/\./g,"").split(/\s+/).filter(Boolean);return ((p[0]?.[0]||"Р")+(p[1]?.[0]||"")).toUpperCase();}
function esc(v){return String(v??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#039;");}
function createId(prefix){return `${prefix}-${crypto?.randomUUID?crypto.randomUUID():`${Date.now()}-${Math.random().toString(16).slice(2)}`}`;}
function localTimestamp(){const d=new Date(),p=n=>String(n).padStart(2,"0");return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;}
function dateMs(v){if(typeof v==="number")return (v-25569)*86400000;if(!v)return 0;const s=clean(v);const m=s.match(/^(\d{4})-(\d{2})-(\d{2})(?:T|\s)?(\d{2})?:?(\d{2})?:?(\d{2})?/);if(m)return new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),Number(m[4]||0),Number(m[5]||0),Number(m[6]||0)).getTime();const n=Date.parse(s);return Number.isFinite(n)?n:0;}
function fmtDateTime(v){const ms=dateMs(v);if(!ms)return "—";const d=new Date(ms),p=n=>String(n).padStart(2,"0");return `${p(d.getDate())}.${p(d.getMonth()+1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;}
function updateLastUpdated(){const d=new Date(),p=n=>String(n).padStart(2,"0");setText("lastUpdated",`↻ ${p(d.getDate())}.${p(d.getMonth()+1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`);}
function errorText(e){return e?.message||String(e);}
