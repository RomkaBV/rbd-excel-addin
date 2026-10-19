const TABLES = {
  requests: "tblRBDRequests2026",
  events: "tblRBDEvents2026",
  assignments: "tblRBDAssignments2026",
  users: "tblRBDUsers2026",
  statuses: "tblRBDStatuses2026",
  transitions: "tblRBDTransitions2026",
  cities: "tblRBDCities2026",
  locations: "tblRBDLocations2026",
  categories: "tblRBDCategories2026",
  customers: "tblRBDCustomers2026",
  priorities: "tblRBDPriorities2026",
  erpStage: "tblRBDERPStage2026",
  erpImports: "tblRBDERPImports2026",
  managerInput: "tblRBDInputManager2026"
};

const STATE = {
  activeSheet: "",
  mode: "unknown", // manager | employee | unknown
  actor: "",
  employeeId: "",
  managerName: "",
  users: [], statuses: [], transitions: [], cities: [], locations: [], categories: [], customers: [], priorities: [],
  requestsAll: [], requests: [], filtered: [], events: [],
  expandedStatusId: "",
  selectedRequest: null,
  erpPreview: [],
  erpFileName: "",
  currentView: "requests"
};

Office.onReady(async info => {
  bindStaticEvents();
  if (info.host !== Office.HostType.Excel) return;
  try {
    await loadReferenceData();
    await detectMode();
    configureMode();
    await loadBusinessData();
    renderAll();
  } catch (error) {
    console.error(error);
    showRequestsError(errorText(error));
  }
});

function bindStaticEvents() {
  document.querySelectorAll("[data-view]").forEach(btn => btn.addEventListener("click", () => switchView(btn.dataset.view)));
  document.querySelectorAll("[data-close]").forEach(btn => btn.addEventListener("click", () => closeModal(btn.dataset.close)));

  byId("refreshButton")?.addEventListener("click", refreshAll);
  byId("historyRefreshButton")?.addEventListener("click", refreshAll);
  byId("newRequestButton")?.addEventListener("click", openCreateModal);
  byId("saveCreateButton")?.addEventListener("click", createRequest);
  byId("saveAssignButton")?.addEventListener("click", saveAssignment);
  byId("saveCommentButton")?.addEventListener("click", saveComment);

  ["filterStatus", "filterExecutor", "filterCity", "filterCategory"].forEach(id => byId(id)?.addEventListener("change", applyFilters));
  byId("searchInput")?.addEventListener("input", applyFilters);
  byId("historySearch")?.addEventListener("input", renderHistory);
  byId("historyAction")?.addEventListener("change", renderHistory);

  byId("createCity")?.addEventListener("change", () => fillLocationSelect("createLocation", byId("createCity")?.value || ""));
  byId("erpFileInput")?.addEventListener("change", onErpFileSelected);
  byId("erpImportButton")?.addEventListener("click", importErpPreview);
  byId("erpClearButton")?.addEventListener("click", clearErpPreview);

  document.querySelectorAll(".dict-add").forEach(btn => btn.addEventListener("click", () => addDictionaryValue(btn.dataset.dict)));
}

async function refreshAll() {
  try {
    setBusy("refreshButton", true, "Оновлення...");
    await loadReferenceData();
    await detectMode();
    configureMode();
    await loadBusinessData();
    renderAll();
    toast("Дані оновлено");
  } catch (error) {
    showRequestsError(errorText(error));
  } finally {
    setBusy("refreshButton", false, "↻ Оновити");
  }
}

async function loadReferenceData() {
  const data = await Excel.run(async context => {
    const result = {};
    for (const key of ["users", "statuses", "transitions", "cities", "locations", "categories", "customers", "priorities"]) {
      result[key] = await readTable(context, TABLES[key], false);
    }
    return result;
  });

  STATE.users = rowsToObjects(data.users).filter(x => isYes(x["Активний"]));
  STATE.statuses = rowsToObjects(data.statuses).filter(x => isYes(x["Активний"])).sort((a,b) => num(a["Порядок"]) - num(b["Порядок"]));
  STATE.transitions = rowsToObjects(data.transitions).filter(x => isYes(x["Активний"]));
  STATE.cities = rowsToObjects(data.cities).filter(x => isYes(x["Активний"]));
  STATE.locations = rowsToObjects(data.locations).filter(x => isYes(x["Активний"]));
  STATE.categories = rowsToObjects(data.categories).filter(x => isYes(x["Активний"]));
  STATE.customers = rowsToObjects(data.customers).filter(x => isYes(x["Активний"]));
  STATE.priorities = rowsToObjects(data.priorities).filter(x => isYes(x["Активний"])).sort((a,b)=>num(a["Порядок"])-num(b["Порядок"]));

  const manager = STATE.users.find(x => String(x["Роль"] || "").includes("Керівник"));
  STATE.managerName = manager?.["ПІБ"] || "Іванченко В.М.";
  fillReferenceSelects();
}

async function detectMode() {
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

  if (norm(STATE.activeSheet) === norm("Кабінет керівника")) {
    STATE.mode = "manager";
    STATE.actor = STATE.managerName;
    const manager = STATE.users.find(x => sameName(x["ПІБ"], STATE.managerName));
    STATE.employeeId = manager?.["ID працівника"] || "";
    return;
  }

  const user = STATE.users.find(x => norm(surname(x["ПІБ"])) === norm(STATE.activeSheet));
  if (user) {
    STATE.mode = "employee";
    STATE.actor = user["ПІБ"];
    STATE.employeeId = user["ID працівника"];
  }
}

function configureMode() {
  if (STATE.mode === "manager") {
    setText("pageTitle", "РБД 2.0 — Кабінет керівника");
    setText("pageSubtitle", "Усі заявки • Усі виконавці • ERP • Контроль");
    setText("currentUserName", STATE.actor);
    setText("currentUserRole", "Керівник РБД");
    setText("userAvatar", initials(STATE.actor));
    show(byId("filterExecutorField"), true);
    show(byId("createExecutorField"), true);
    show(byId("erpManagerNotice"), false);
    show(byId("dictionaryManagerNotice"), false);
    document.querySelectorAll(".dict-add").forEach(x => x.disabled = false);
    byId("erpImportButton") && (byId("erpImportButton").disabled = STATE.erpPreview.length === 0);
  } else if (STATE.mode === "employee") {
    setText("pageTitle", `РБД 2.0 — ${STATE.actor}`);
    setText("pageSubtitle", "Мої заявки • Опрацювання • Контроль");
    setText("currentUserName", STATE.actor);
    setText("currentUserRole", "Виконавець");
    setText("userAvatar", initials(STATE.actor));
    show(byId("filterExecutorField"), false);
    show(byId("createExecutorField"), false);
    show(byId("erpManagerNotice"), true);
    show(byId("dictionaryManagerNotice"), true);
    document.querySelectorAll(".dict-add").forEach(x => x.disabled = true);
    if (byId("erpImportButton")) byId("erpImportButton").disabled = true;
  } else {
    setText("pageTitle", "РБД 2.0 — Центр заявок");
    setText("pageSubtitle", "Відкрийте вкладку працівника або «Кабінет керівника»");
    setText("currentUserName", "РБД");
    setText("currentUserRole", "Кабінет не визначено");
    setText("userAvatar", "Р");
    show(byId("filterExecutorField"), false);
    show(byId("erpManagerNotice"), true);
    show(byId("dictionaryManagerNotice"), true);
    document.querySelectorAll(".dict-add").forEach(x => x.disabled = true);
  }
}

async function loadBusinessData() {
  const data = await Excel.run(async context => ({
    requests: await readTable(context, TABLES.requests, true),
    events: await readTable(context, TABLES.events, false)
  }));

  STATE.requestsAll = rowsToObjects(data.requests).map(normalizeRequest).filter(x => x.id);
  STATE.events = rowsToObjects(data.events);

  if (STATE.mode === "manager") STATE.requests = STATE.requestsAll.slice();
  else if (STATE.mode === "employee") STATE.requests = STATE.requestsAll.filter(x => sameName(x.executor, STATE.actor));
  else STATE.requests = [];

  STATE.expandedStatusId = "";
  applyFilters();
  updateTimestamp();
}

async function readTable(context, tableName, required = true) {
  const table = context.workbook.tables.getItemOrNullObject(tableName);
  table.load("isNullObject");
  await context.sync();
  if (table.isNullObject) {
    if (required) throw new Error(`Не знайдено таблицю ${tableName}`);
    return { headers: [], rows: [] };
  }

  const header = table.getHeaderRowRange();
  header.load("values");
  table.rows.load("items");
  await context.sync();
  const headers = header.values[0] || [];
  if (!table.rows.items.length) return { headers, rows: [] };
  const body = table.getDataBodyRange();
  body.load("values");
  await context.sync();
  return { headers, rows: body.values };
}

function rowsToObjects(data) {
  if (!data?.headers?.length) return [];
  return data.rows.map(row => Object.fromEntries(data.headers.map((h,i)=>[String(h), row[i] ?? ""])));
}

function normalizeRequest(x) {
  return {
    id: clean(x["ID заявки"]), erpId: clean(x["ID ERP"]), subject: clean(x["Тема заявки"]), description: clean(x["Опис заявки"]),
    customer: clean(x["Замовник"]), city: clean(x["Місто"]), location: clean(x["Локація"]), category: clean(x["Категорія"]),
    priority: clean(x["Пріоритет"]), status: clean(x["Статус"]), executor: clean(x["Виконавець"]), createdBy: clean(x["Створив"]),
    createdAt: x["Дата створення"], updatedAt: x["Дата оновлення"], version: num(x["Версія"]), source: clean(x["Джерело"]),
    statusSince: x["Дата зміни статусу"], completedAt: x["Дата завершення"]
  };
}

function fillReferenceSelects() {
  const names = STATE.users.map(x=>x["ПІБ"]).filter(Boolean);
  const statuses = STATE.statuses.map(x=>x["Назва статусу"]).filter(Boolean);
  const cities = STATE.cities.map(x=>x["Назва"]).filter(Boolean);
  const categories = STATE.categories.map(x=>x["Назва"]).filter(Boolean);
  const customers = STATE.customers.map(x=>x["Назва"]).filter(Boolean);
  const priorities = STATE.priorities.map(x=>x["Назва"]).filter(Boolean);
  fillSelect("filterExecutor", names, "Усі виконавці");
  fillSelect("filterStatus", statuses, "Усі статуси");
  fillSelect("filterCity", cities, "Усі міста");
  fillSelect("filterCategory", categories, "Усі категорії");
  fillSelect("createExecutor", names, "— Без виконавця —");
  fillSelect("assignExecutor", names, "Оберіть виконавця");
  fillSelect("createCity", cities, "Оберіть місто");
  fillSelect("createCategory", categories, "Оберіть категорію");
  fillSelect("createCustomer", customers, "Оберіть замовника");
  fillSelect("createPriority", priorities, "Оберіть пріоритет");
  fillSelect("newLocationCity", cities, "Оберіть місто");
  fillLocationSelect("createLocation", "");
  if (byId("createPriority") && priorities.includes("Середній")) byId("createPriority").value = "Середній";
}

function fillLocationSelect(id, city) {
  const values = STATE.locations.filter(x => !city || sameText(x["Місто"], city)).map(x=>x["Назва"]).filter(Boolean);
  fillSelect(id, values, "Оберіть локацію");
}

function applyFilters() {
  const search = norm(value("searchInput"));
  const status = value("filterStatus");
  const executor = value("filterExecutor");
  const city = value("filterCity");
  const category = value("filterCategory");

  STATE.filtered = STATE.requests.filter(r => {
    if (status && r.status !== status) return false;
    if (STATE.mode === "manager" && executor && !sameName(r.executor, executor)) return false;
    if (city && r.city !== city) return false;
    if (category && r.category !== category) return false;
    if (search) {
      const hay = norm([r.id,r.erpId,r.subject,r.description,r.customer,r.city,r.location,r.category,r.priority,r.status,r.executor].join(" "));
      if (!hay.includes(search)) return false;
    }
    return true;
  });
  renderRequests();
  renderKpis();
  renderAnalytics();
  renderEmployees();
  renderHistory();
}

function renderAll() {
  fillReferenceSelects();
  applyFilters();
  renderDictionaries();
  renderHistoryActionOptions();
  configureMode();
}

function renderKpis() {
  const active = STATE.requests.filter(r => !isFinalStatus(r.status));
  const overdue = active.filter(isOverdue);
  const work = active.filter(r => !["Нова","Розподілена","Призупинена"].includes(r.status));
  const thirty = Date.now() - 30*86400000;
  const closed = STATE.requests.filter(r => isFinalStatus(r.status) && dateMs(r.completedAt) >= thirty);
  setText("kpiActive", active.length); setText("kpiOverdue", overdue.length); setText("kpiWork", work.length); setText("kpiClosed", closed.length);
  setText("requestSummary", `Показано ${STATE.filtered.length} із ${STATE.requests.length} заявок`);
}

function renderRequests() {
  const body = byId("requestsBody"); if (!body) return;
  if (!STATE.filtered.length) { body.innerHTML = `<tr><td colspan="13"><div class="empty">Заявок за обраними умовами немає</div></td></tr>`; return; }
  let html = "";
  for (const r of STATE.filtered) {
    html += requestRow(r);
    if (STATE.expandedStatusId === r.id) html += statusPanel(r);
  }
  body.innerHTML = html;
  bindRequestEvents();
}

function requestRow(r) {
  const status = getStatus(r.status);
  const bg = status?.["Колір фону"] || "#EAF4FF";
  const fg = status?.["Колір тексту"] || "#17382B";
  const rowBg = hexToRgba(bg, isOverdue(r) ? .34 : .18);
  const hours = currentStatusHours(r);
  const sla = num(status?.["SLA, год"]);
  const priority = getPriorityStyle(r.priority);
  return `<tr class="request-row" style="background:${rowBg}">
    <td><div class="request-id">${esc(r.id)}</div></td>
    <td>${esc(r.erpId || "—")}</td>
    <td><div class="ellipsis" title="${esc(r.subject)}">${esc(r.subject || "—")}</div></td>
    <td><div class="ellipsis" title="${esc(r.category)}">${esc(r.category || "—")}</div></td>
    <td>${esc(r.city || "—")}</td>
    <td><div class="ellipsis" title="${esc(r.customer)}">${esc(r.customer || "—")}</div></td>
    <td><div class="ellipsis" title="${esc(r.executor)}">${esc(r.executor || "—")}</div></td>
    <td><span class="priority-pill" style="background:${priority.bg};color:${priority.fg}">${esc(r.priority || "—")}</span></td>
    <td><button class="status-pill" data-status-open="${esc(r.id)}" style="background:${bg};color:${fg}"><span class="dot"></span>${esc(r.status || "—")}</button></td>
    <td class="${isOverdue(r)?"overdue":""}">${hours === null ? "—" : `${hours} год.`}</td>
    <td>${sla > 0 ? `${sla} год.` : "—"}</td>
    <td>${r.version || 0}</td>
    <td><div class="action-row"><button class="icon-btn" data-view-request="${esc(r.id)}" title="Деталі">◉</button>${STATE.mode==="manager"?`<button class="icon-btn" data-assign-request="${esc(r.id)}" title="Призначити">♙</button>`:""}<button class="icon-btn" data-comment-request="${esc(r.id)}" title="Коментар">✎</button></div></td>
  </tr>`;
}

function statusPanel(r) {
  const allowed = allowedNextStatuses(r.status);
  const buttons = allowed.map(name => {
    const s = getStatus(name); const bg=s?.["Колір фону"]||"#edf3f7"; const fg=s?.["Колір тексту"]||"#17382B";
    return `<button class="status-choice" data-change-status="${esc(name)}" data-request-id="${esc(r.id)}" style="background:${bg};color:${fg}">${esc(name)}</button>`;
  }).join("");
  return `<tr class="status-panel-row"><td colspan="13"><div class="status-panel">
    <div class="status-panel-head"><div class="status-panel-title">${esc(r.id)} • ${esc(r.status)} → оберіть наступний статус</div><button class="status-close" data-status-close>×</button></div>
    <div class="status-options">${buttons || `<span class="muted">Немає дозволених переходів зі статусу «${esc(r.status)}»</span>`}</div>
    <textarea id="statusComment_${domId(r.id)}" class="status-comment" placeholder="Коментар. Для статусу «Призупинена» причина обов'язкова."></textarea>
    <div class="panel-actions"><span id="statusMsg_${domId(r.id)}" class="inline-msg"></span></div>
  </div></td></tr>`;
}

function bindRequestEvents() {
  document.querySelectorAll("[data-status-open]").forEach(btn => btn.addEventListener("click", () => { STATE.expandedStatusId = STATE.expandedStatusId === btn.dataset.statusOpen ? "" : btn.dataset.statusOpen; renderRequests(); }));
  document.querySelector("[data-status-close]")?.addEventListener("click", () => { STATE.expandedStatusId=""; renderRequests(); });
  document.querySelectorAll("[data-view-request]").forEach(btn => btn.addEventListener("click", () => openDetails(findRequest(btn.dataset.viewRequest))));
  document.querySelectorAll("[data-assign-request]").forEach(btn => btn.addEventListener("click", () => openAssign(findRequest(btn.dataset.assignRequest))));
  document.querySelectorAll("[data-comment-request]").forEach(btn => btn.addEventListener("click", () => openComment(findRequest(btn.dataset.commentRequest))));
  document.querySelectorAll("[data-change-status]").forEach(btn => btn.addEventListener("click", () => changeStatus(findRequest(btn.dataset.requestId), btn.dataset.changeStatus)));
}

async function changeStatus(r, next) {
  if (!r) return;
  const comment = value(`statusComment_${domId(r.id)}`);
  if (next === "Призупинена" && !comment) { setText(`statusMsg_${domId(r.id)}`, "Вкажіть причину призупинення."); return; }
  if (next === "Закрита" && !confirm(`Закрити заявку ${r.id}?`)) return;
  try {
    setText(`statusMsg_${domId(r.id)}`, "Передаємо команду...");
    await sendCommand({
      action: next === "Закрита" ? "CLOSE" : "STATUS_CHANGE",
      requestId: r.id,
      status: next,
      comment,
      expectedVersion: r.version
    });
    setText(`statusMsg_${domId(r.id)}`, `Команду «${r.status} → ${next}» передано на обробку.`);
    toast("Зміну статусу передано на обробку");
  } catch (error) { setText(`statusMsg_${domId(r.id)}`, `Помилка: ${errorText(error)}`); }
}

function openCreateModal() {
  if (STATE.mode === "unknown") return toast("Відкрийте кабінет працівника або керівника");
  ["createSubject","createDescription","createComment"].forEach(id=>setInput(id,""));
  ["createCity","createLocation","createCategory","createCustomer","createExecutor"].forEach(id=>setInput(id,""));
  if (byId("createPriority") && STATE.priorities.some(x=>x["Назва"]==="Середній")) byId("createPriority").value="Середній";
  clearMessage("createMessage");
  openModal("createModal");
}

async function createRequest() {
  const subject = value("createSubject");
  if (!subject) return showMessage("createMessage","error","Вкажіть тему заявки.");
  const executor = STATE.mode === "employee" ? STATE.actor : value("createExecutor");
  try {
    setBusy("saveCreateButton",true,"Створюємо...");
    await sendCommand({
      action:"CREATE", requestId:"", subject, description:value("createDescription"), executor,
      status:"", city:value("createCity"), location:value("createLocation"), category:value("createCategory"),
      customer:value("createCustomer"), priority:value("createPriority") || "Середній", comment:value("createComment"), expectedVersion:""
    });
    showMessage("createMessage","success","Команду на створення заявки передано. Після виконання сценарію натисніть «Оновити».");
    toast("Нову заявку передано на обробку");
  } catch (error) { showMessage("createMessage","error",errorText(error)); }
  finally { setBusy("saveCreateButton",false,"Створити заявку"); }
}

function openDetails(r) {
  if (!r) return;
  STATE.selectedRequest = r;
  setText("detailsTitle", r.id);
  const status = getStatus(r.status); const sla=num(status?.["SLA, год"]); const hours=currentStatusHours(r);
  const rows = [["ID ERP",r.erpId],["Тема",r.subject],["Опис",r.description],["Замовник",r.customer],["Місто",r.city],["Локація",r.location],["Категорія",r.category],["Пріоритет",r.priority],["Статус",r.status],["Виконавець",r.executor],["Створив",r.createdBy],["Дата створення",fmtDateTime(r.createdAt)],["Дата зміни статусу",fmtDateTime(r.statusSince)],["Час у статусі",hours===null?"—":`${hours} год.`],["SLA",sla?`${sla} год.`:"—"],["Версія",r.version],["Джерело",r.source],["Дата завершення",fmtDateTime(r.completedAt)]];
  byId("detailsContent").innerHTML = rows.map(([a,b])=>`<div class="details-label">${esc(a)}</div><div class="details-value">${esc(String(b||"—"))}</div>`).join("");
  openModal("detailsModal");
}

function openAssign(r) {
  if (!r || STATE.mode !== "manager") return;
  STATE.selectedRequest = r;
  setText("assignTitle", `Призначення — ${r.id}`);
  setInput("assignExecutor",r.executor||""); setInput("assignComment",""); clearMessage("assignMessage"); openModal("assignModal");
}

async function saveAssignment() {
  const r=STATE.selectedRequest; if(!r)return;
  const executor=value("assignExecutor"); if(!executor)return showMessage("assignMessage","error","Оберіть виконавця.");
  try{setBusy("saveAssignButton",true,"Призначаємо...");await sendCommand({action:"ASSIGN",requestId:r.id,executor,comment:value("assignComment"),expectedVersion:r.version});showMessage("assignMessage","success","Команду передано на обробку.");toast("Призначення передано на обробку");}
  catch(e){showMessage("assignMessage","error",errorText(e));}finally{setBusy("saveAssignButton",false,"Призначити");}
}

function openComment(r){if(!r)return;STATE.selectedRequest=r;setText("commentTitle",`Коментар — ${r.id}`);setInput("requestComment","");clearMessage("commentMessage");openModal("commentModal");}
async function saveComment(){const r=STATE.selectedRequest;if(!r)return;const comment=value("requestComment");if(!comment)return showMessage("commentMessage","error","Введіть коментар.");try{setBusy("saveCommentButton",true,"Додаємо...");await sendCommand({action:"COMMENT",requestId:r.id,comment,expectedVersion:r.version});showMessage("commentMessage","success","Коментар передано на обробку.");toast("Коментар передано на обробку");}catch(e){showMessage("commentMessage","error",errorText(e));}finally{setBusy("saveCommentButton",false,"Додати коментар");}}

async function sendCommand(p) {
  if (STATE.mode === "unknown") throw new Error("Не визначено кабінет користувача");
  const tableName = STATE.mode === "manager" ? TABLES.managerInput : `tblRBDInput${String(STATE.employeeId).replace(/-/g,"")}`;
  const row = {
    "ID команди": createId("CMD"), "Дія": p.action || "", "ID заявки": p.requestId || "", "Тема заявки": p.subject || "", "Опис заявки": p.description || "",
    "Виконавець": p.executor || "", "Статус": p.status || "", "Місто": p.city || "", "Локація": p.location || "", "Категорія": p.category || "",
    "Замовник": p.customer || "", "Пріоритет": p.priority || "", "Коментар": p.comment || "", "Очікувана версія": p.expectedVersion ?? "", "Стан обробки": "Нова", "Результат": ""
  };

  await Excel.run(async context => {
    const table = context.workbook.tables.getItemOrNullObject(tableName);
    table.load("isNullObject"); await context.sync();
    if (table.isNullObject) throw new Error(`Не знайдено таблицю команд ${tableName}. Запустіть RBD_ACTIVITY_SETUP.`);
    const header=table.getHeaderRowRange();header.load("values");await context.sync();
    const headers=header.values[0];
    table.rows.add(null,[headers.map(h=>row[String(h)] ?? "")]);
    await context.sync();
  });
}

function allowedNextStatuses(currentName) {
  const current = STATE.statuses.find(x => x["Назва статусу"] === currentName);
  if (!current) return [];
  const fromId = current["ID статусу"];
  const ids = new Set(STATE.transitions.filter(x => x["Із статусу (ID)"] === fromId).map(x=>x["До статусу (ID)"]));
  return STATE.statuses.filter(x=>ids.has(x["ID статусу"])).map(x=>x["Назва статусу"]);
}

function getStatus(name){return STATE.statuses.find(x=>x["Назва статусу"]===name);}
function isFinalStatus(name){const s=getStatus(name);return s?isYes(s["Завершальний"]):["Закрита","Скасована","Виконана"].includes(name);}
function currentStatusHours(r){if(isFinalStatus(r.status)||r.status==="Призупинена")return null;const ms=dateMs(r.statusSince);if(!ms)return null;return Math.max(0,Math.round(((Date.now()-ms)/3600000)*10)/10);}
function isOverdue(r){if(isFinalStatus(r.status)||r.status==="Призупинена")return false;const sla=num(getStatus(r.status)?.["SLA, год"]);const h=currentStatusHours(r);return sla>0&&h!==null&&h>sla;}

function renderEmployees(){const body=byId("employeesBody");if(!body)return;let users=STATE.users;if(STATE.mode==="employee")users=users.filter(x=>sameName(x["ПІБ"],STATE.actor));if(!users.length){body.innerHTML=`<tr><td colspan="6"><div class="empty">Немає даних</div></td></tr>`;return;}const thirty=Date.now()-30*86400000;body.innerHTML=users.map(u=>{const n=u["ПІБ"];const req=STATE.requestsAll.filter(r=>sameName(r.executor,n));const active=req.filter(r=>!isFinalStatus(r.status));const work=active.filter(r=>!["Нова","Розподілена","Призупинена"].includes(r.status));const overdue=active.filter(isOverdue);const closed=req.filter(r=>isFinalStatus(r.status)&&dateMs(r.completedAt)>=thirty);return `<tr><td class="bold">${esc(n)}</td><td>${esc(u["Роль"]||"—")}</td><td>${active.length}</td><td>${work.length}</td><td class="${overdue.length?"overdue":""}">${overdue.length}</td><td>${closed.length}</td></tr>`;}).join("");}

function renderAnalytics(){const req=STATE.requests;setText("analyticsTotal",req.length);setText("analyticsUnassigned",req.filter(r=>!r.executor).length);setText("analyticsOverdue",req.filter(isOverdue).length);renderBars("statusAnalytics",groupCount(req,r=>r.status||"Без статусу"),STATE.statuses.length?STATE.statuses.map(x=>x["Назва статусу"]):[]);renderBars("priorityAnalytics",groupCount(req,r=>r.priority||"Без пріоритету"),STATE.priorities.map(x=>x["Назва"]));}
function renderBars(id,map,order){const el=byId(id);if(!el)return;const entries=[];for(const key of order)if(map[key])entries.push([key,map[key]]);for(const [k,v] of Object.entries(map))if(!entries.some(x=>x[0]===k))entries.push([k,v]);const max=Math.max(1,...entries.map(x=>x[1]));el.innerHTML=entries.length?entries.map(([k,v])=>`<div class="bar-row"><div class="ellipsis" title="${esc(k)}">${esc(k)}</div><div class="bar-track"><div class="bar-fill" style="width:${Math.round(v/max*100)}%"></div></div><div class="right bold">${v}</div></div>`).join(""):`<div class="empty">Немає даних</div>`;}
function groupCount(items,keyFn){const m={};for(const x of items){const k=keyFn(x);m[k]=(m[k]||0)+1;}return m;}

function renderHistoryActionOptions(){const actions=[...new Set(STATE.events.map(x=>clean(x["Дія"])).filter(Boolean))].sort();fillSelect("historyAction",actions,"Усі дії");}
function renderHistory(){const body=byId("historyBody");if(!body)return;const search=norm(value("historySearch"));const action=value("historyAction");const ids=new Set(STATE.requests.map(r=>r.id));let events=STATE.events.filter(e=>STATE.mode==="manager"?true:ids.has(clean(e["ID заявки"])));if(action)events=events.filter(e=>clean(e["Дія"])===action);if(search)events=events.filter(e=>norm([e["ID заявки"],e["Користувач"],e["Дія"],e["Змінене поле"],e["Було"],e["Стало"],e["Коментар"]].join(" ")).includes(search));events.sort((a,b)=>dateMs(b["Дата події"])-dateMs(a["Дата події"]));setText("historySummary",`${events.length} подій`);body.innerHTML=events.length?events.slice(0,500).map(e=>`<tr><td>${esc(fmtDateTime(e["Дата події"]))}</td><td class="bold">${esc(e["ID заявки"]||"—")}</td><td>${esc(e["Користувач"]||"—")}</td><td>${esc(e["Дія"]||"—")}</td><td>${esc(e["Змінене поле"]||"—")}</td><td><div class="ellipsis">${esc(e["Було"]||"—")}</div></td><td><div class="ellipsis">${esc(e["Стало"]||"—")}</div></td><td><div class="ellipsis" title="${esc(e["Коментар"]||"")}">${esc(e["Коментар"]||"—")}</div></td></tr>`).join(""):`<tr><td colspan="8"><div class="empty">Подій не знайдено</div></td></tr>`;}

function renderDictionaries(){renderChipList("citiesList",STATE.cities.map(x=>x["Назва"]));renderChipList("categoriesList",STATE.categories.map(x=>x["Назва"]));renderChipList("customersList",STATE.customers.map(x=>x["Назва"]));renderChipList("locationsList",STATE.locations.map(x=>`${x["Місто"] ? x["Місто"]+" — " : ""}${x["Назва"]}`));}
function renderChipList(id,values){const el=byId(id);if(!el)return;el.innerHTML=values.length?`<div style="display:flex;flex-wrap:wrap;gap:6px">${values.map(v=>`<span style="padding:5px 8px;border-radius:14px;background:#eef5fa;border:1px solid #dbe6ee;font-size:8.8px">${esc(v)}</span>`).join("")}</div>`:`<div class="muted" style="font-size:9px">Довідник порожній</div>`;}

async function addDictionaryValue(kind){if(STATE.mode!=="manager")return toast("Додавати значення може лише керівник");let payload={};if(kind==="city"){const v=value("newCity");if(!v)return;payload={action:"ADD_CITY",city:v};}
else if(kind==="category"){const v=value("newCategory");if(!v)return;payload={action:"ADD_CATEGORY",category:v};}
else if(kind==="customer"){const v=value("newCustomer");if(!v)return;payload={action:"ADD_CUSTOMER",customer:v};}
else if(kind==="location"){const city=value("newLocationCity"),name=value("newLocationName"),address=value("newLocationAddress");if(!city||!name)return toast("Для локації виберіть місто та вкажіть назву");payload={action:"ADD_LOCATION",city,location:name,description:address};}
try{await sendCommand(payload);toast("Значення передано на додавання. Після виконання сценарію натисніть «Оновити».");}catch(e){toast(`Помилка: ${errorText(e)}`);}}

function switchView(name){STATE.currentView=name;document.querySelectorAll(".view").forEach(x=>x.classList.remove("active"));byId(`view-${name}`)?.classList.add("active");document.querySelectorAll("[data-view]").forEach(x=>x.classList.toggle("active",x.dataset.view===name));if(name==="history")renderHistory();if(name==="employees")renderEmployees();if(name==="analytics")renderAnalytics();if(name==="dictionaries")renderDictionaries();}

async function onErpFileSelected(event){const file=event.target.files?.[0];if(!file)return;STATE.erpFileName=file.name;setText("erpFileName",file.name);try{const buffer=await file.arrayBuffer();const wb=XLSX.read(buffer,{type:"array",cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];const raw=XLSX.utils.sheet_to_json(ws,{defval:"",raw:false});STATE.erpPreview=raw.map(mapErpRow).filter(x=>Object.values(x).some(Boolean));renderErpPreview();}catch(e){clearErpPreview();toast(`Не вдалося прочитати файл: ${errorText(e)}`);}}

function mapErpRow(row){const get=(...aliases)=>{for(const alias of aliases){const key=Object.keys(row).find(k=>norm(k)===norm(alias));if(key!==undefined&&clean(row[key]))return clean(row[key]);}return "";};return {
  erpId:get("ID ERP","ERP ID","Зовнішній ID","External ID","ID","Номер заявки","Заявка"),
  subject:get("Тема заявки","Тема","Назва","Найменування","Предмет","Subject"),
  description:get("Опис заявки","Опис","Description","Коментар"),
  customer:get("Замовник","Customer","Ініціатор","Автор"),
  city:get("Місто","City"),location:get("Локація","Location","Адреса","Склад","Об'єкт","Об’єкт"),
  category:get("Категорія","Category"),priority:get("Пріоритет","Priority") || "Середній"
};}

function renderErpPreview(){const body=byId("erpPreviewBody");const existing=new Set(STATE.requestsAll.map(r=>norm(r.erpId)).filter(Boolean));let dup=0,err=0,newCount=0;for(const r of STATE.erpPreview){if(!r.erpId)err++;else if(existing.has(norm(r.erpId)))dup++;else newCount++;}setText("erpRowsCount",STATE.erpPreview.length);setText("erpNewCount",newCount);setText("erpDuplicateCount",dup);setText("erpErrorCount",err);setText("erpPreviewSub",`${STATE.erpPreview.length} рядків • ${newCount} нових • ${dup} вже існує`);if(byId("erpImportButton"))byId("erpImportButton").disabled=STATE.mode!=="manager"||STATE.erpPreview.length===0;if(byId("erpClearButton"))byId("erpClearButton").disabled=STATE.erpPreview.length===0;body.innerHTML=STATE.erpPreview.length?STATE.erpPreview.slice(0,100).map(r=>`<tr><td class="bold">${esc(r.erpId||"—")}</td><td><div class="ellipsis">${esc(r.subject||"—")}</div></td><td><div class="ellipsis">${esc(r.description||"—")}</div></td><td>${esc(r.customer||"—")}</td><td>${esc(r.city||"—")}</td><td><div class="ellipsis">${esc(r.location||"—")}</div></td><td>${esc(r.category||"—")}</td><td>${esc(r.priority||"—")}</td></tr>`).join(""):`<tr><td colspan="8"><div class="empty">Файл ще не вибрано</div></td></tr>`;}
function clearErpPreview(){STATE.erpPreview=[];STATE.erpFileName="";if(byId("erpFileInput"))byId("erpFileInput").value="";setText("erpFileName","Файл не вибрано");renderErpPreview();}

async function importErpPreview(){if(STATE.mode!=="manager")return toast("Імпорт ERP доступний лише керівнику");if(!STATE.erpPreview.length)return;try{setBusy("erpImportButton",true,"Імпорт...");const result=await Excel.run(async context=>{
    const reqTable=context.workbook.tables.getItem(TABLES.requests),stageTable=context.workbook.tables.getItem(TABLES.erpStage),impTable=context.workbook.tables.getItem(TABLES.erpImports),eventTable=context.workbook.tables.getItem(TABLES.events);
    const tables=[reqTable,stageTable,impTable,eventTable];for(const t of tables){t.getHeaderRowRange().load("values");t.rows.load("items");}await context.sync();
    const headers={req:reqTable.getHeaderRowRange().values[0],stage:stageTable.getHeaderRowRange().values[0],imp:impTable.getHeaderRowRange().values[0],event:eventTable.getHeaderRowRange().values[0]};
    let reqRows=[];if(reqTable.rows.items.length){const body=reqTable.getDataBodyRange();body.load("values");await context.sync();reqRows=body.values;}
    const idxId=headers.req.indexOf("ID заявки"),idxErp=headers.req.indexOf("ID ERP");let max=0;const existing=new Map();for(const row of reqRows){const id=clean(row[idxId]);const m=/^RBD-(\d+)$/.exec(id);if(m)max=Math.max(max,Number(m[1]));const erp=norm(row[idxErp]);if(erp)existing.set(erp,id);}
    const importId=createId("IMP"),timestamp=localTimestamp();const newReq=[],stageRows=[],eventRows=[];let added=0,duplicates=0,errors=0;
    for(let i=0;i<STATE.erpPreview.length;i++){const r=STATE.erpPreview[i];let resultText="",rbdId="";if(!r.erpId){errors++;resultText="Помилка: відсутній ID ERP";}else if(existing.has(norm(r.erpId))){duplicates++;rbdId=existing.get(norm(r.erpId));resultText="Вже існує";}else{max++;rbdId=`RBD-${String(max).padStart(6,"0")}`;existing.set(norm(r.erpId),rbdId);const requestObj={"ID заявки":rbdId,"ID ERP":r.erpId,"Тема заявки":r.subject||`ERP ${r.erpId}`,"Опис заявки":r.description,"Замовник":r.customer,"Місто":r.city,"Локація":r.location,"Категорія":r.category,"Пріоритет":r.priority||"Середній","Статус":"Нова","Виконавець":"","Створив":STATE.actor,"Дата створення":timestamp,"Дата оновлення":timestamp,"Версія":1,"Джерело":"ERP","Дата зміни статусу":timestamp,"Дата завершення":""};newReq.push(headers.req.map(h=>requestObj[String(h)]??""));const ev={"ID події":createId("EV"),"ID команди":importId,"ID заявки":rbdId,"Дата події":timestamp,"Користувач":STATE.actor,"Дія":"Імпорт ERP","Змінене поле":"ID ERP","Було":"","Стало":r.erpId,"Коментар":STATE.erpFileName,"Джерело":"ERP"};eventRows.push(headers.event.map(h=>ev[String(h)]??""));added++;resultText="Імпортовано";}
      const stage={"ID імпорту":importId,"Зовнішній ID":r.erpId,"Джерело":STATE.erpFileName||"ERP","Тема заявки":r.subject,"Опис заявки":r.description,"Результат перевірки":resultText,"ID заявки РБД":rbdId};stageRows.push(headers.stage.map(h=>stage[String(h)]??""));}
    if(newReq.length)reqTable.rows.add(null,newReq);if(stageRows.length)stageTable.rows.add(null,stageRows);if(eventRows.length)eventTable.rows.add(null,eventRows);
    const imp={"ID імпорту":importId,"Дата імпорту":timestamp,"Імпортував":STATE.actor,"Назва файлу":STATE.erpFileName,"Нових":added,"Оновлених":0,"Помилок":errors};impTable.rows.add(null,[headers.imp.map(h=>imp[String(h)]??"")]);await context.sync();return{added,duplicates,errors};
  });toast(`ERP імпорт завершено: нових ${result.added}, вже існує ${result.duplicates}, помилок ${result.errors}`);await loadBusinessData();clearErpPreview();renderAll();}
  catch(e){console.error(e);toast(`Помилка ERP імпорту: ${errorText(e)}`);}finally{setBusy("erpImportButton",false,"Імпортувати в РБД");configureMode();}}

function getPriorityStyle(name){switch(name){case"Низький":return{bg:"#D5F2DD",fg:"#17653B"};case"Середній":return{bg:"#FFF0B8",fg:"#785300"};case"Високий":return{bg:"#FFD4D7",fg:"#9C2330"};case"Критичний":return{bg:"#F5ABB2",fg:"#7D1420"};default:return{bg:"#edf3f7",fg:"#4d6477"};}}
function findRequest(id){return STATE.requests.find(r=>r.id===id)||STATE.requestsAll.find(r=>r.id===id);}
function fillSelect(id,values,placeholder){const el=byId(id);if(!el)return;const old=el.value;el.innerHTML="";const first=document.createElement("option");first.value="";first.textContent=placeholder;el.appendChild(first);for(const v of values){const o=document.createElement("option");o.value=v;o.textContent=v;el.appendChild(o);}if(values.includes(old))el.value=old;}
function openModal(id){byId(id)?.classList.add("open");}function closeModal(id){byId(id)?.classList.remove("open");}
function showMessage(id,type,text){const el=byId(id);if(!el)return;el.className=`message ${type}`;el.textContent=text;}function clearMessage(id){const el=byId(id);if(!el)return;el.className="message";el.textContent="";}
function setBusy(id,busy,text){const el=byId(id);if(!el)return;el.disabled=busy;el.textContent=text;}
function showRequestsError(message){const body=byId("requestsBody");if(body)body.innerHTML=`<tr><td colspan="13"><div class="empty overdue">Помилка: ${esc(message)}</div></td></tr>`;}
function updateTimestamp(){const d=new Date();setText("lastUpdated",`↻ Оновлено ${pad(d.getDate())}.${pad(d.getMonth()+1)}.${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`);}
function toast(text){const el=byId("toast");if(!el)return;el.textContent=text;el.classList.add("show");setTimeout(()=>el.classList.remove("show"),3200);}
function byId(id){return document.getElementById(id);}function value(id){return clean(byId(id)?.value);}function setInput(id,v){const el=byId(id);if(el)el.value=v??"";}function setText(id,v){const el=byId(id);if(el)el.textContent=String(v??"");}function show(el,yes){if(el)el.style.display=yes?"":"none";}
function clean(v){return String(v??"").trim();}function norm(v){return clean(v).replace(/\u00A0/g," ").replace(/\s+/g," ").toLocaleLowerCase();}function sameText(a,b){return norm(a)===norm(b);}function surname(name){return clean(name).split(/\s+/)[0]||"";}function sameName(a,b){if(!a||!b)return false;if(norm(a)===norm(b))return true;return norm(surname(a))===norm(surname(b));}function num(v){const n=Number(String(v??"").replace(/\s/g,"").replace(",","."));return Number.isFinite(n)?n:0;}function isYes(v){return ["так","yes","true","1"].includes(norm(v));}
function dateMs(v){if(!v)return 0;if(typeof v==="number")return(v-25569)*86400000;const s=clean(v);const m=s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/);if(m)return new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),Number(m[4]||0),Number(m[5]||0),Number(m[6]||0)).getTime();const p=Date.parse(s);return Number.isFinite(p)?p:0;}
function fmtDateTime(v){const ms=dateMs(v);if(!ms)return"—";const d=new Date(ms);return`${pad(d.getDate())}.${pad(d.getMonth()+1)}.${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;}
function localTimestamp(){const d=new Date();return`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;}
function createId(prefix){return window.crypto?.randomUUID?`${prefix}-${crypto.randomUUID()}`:`${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;}function initials(name){const p=clean(name).replace(/\./g,"").split(/\s+/);return((p[0]?.[0]||"Р")+(p[1]?.[0]||"")).toUpperCase();}function pad(v){return String(v).padStart(2,"0");}function domId(v){return clean(v).replace(/[^a-zA-Z0-9_-]/g,"_");}
function hexToRgba(hex,a){const m=/^#([0-9a-f]{6})$/i.exec(hex||"");if(!m)return`rgba(234,244,255,${a})`;const n=parseInt(m[1],16);return`rgba(${(n>>16)&255},${(n>>8)&255},${n&255},${a})`;}
function esc(v){return String(v??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#039;");}function errorText(e){return e?.message||String(e);}
