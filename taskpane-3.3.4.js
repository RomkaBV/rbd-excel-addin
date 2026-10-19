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
  assignments: "tblRBDAssignments2026",
  messages: "tblRBDMessages2026",
  reminders: "tblRBDReminders2026"
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
  messages: [],
  reminders: [],
  historyFilter: "all",
  selectedRequest: null,
  expandedStatusId: "",
  expandedAssignId: "",
  expandedFactId: "",
  dictReturnTarget: "",
  currentView: "requests",
  erpPreview: [],
  erpFileName: "",
  quickStatus: "",
  kpiFilter: "",
  distributionSelected: [],
  distributionExecutor: "",
  distributionShowAll: false,
  distributionSearch: ""
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
  byId("distributionButton")?.addEventListener("click", openDistributionModal);
  byId("distributionSearch")?.addEventListener("input", e => { STATE.distributionSearch=clean(e.target.value); renderDistribution(); });
  byId("distributionShowAll")?.addEventListener("change", e => { STATE.distributionShowAll=!!e.target.checked; STATE.distributionSelected=[]; renderDistribution(); });
  byId("distributionAssignButton")?.addEventListener("click", assignDistributionSelection);
  byId("saveCreateButton")?.addEventListener("click", saveCreate);
  document.querySelectorAll("[data-create-dict]").forEach(btn => btn.addEventListener("click", () => {
    STATE.dictReturnTarget = btn.dataset.target || "";
    openDictionaryModal(btn.dataset.createDict, true);
  }));
  byId("createCity")?.addEventListener("change", () => fillLocationSelect("createLocation", value("createCity"), "Оберіть локацію"));
  byId("saveAssignButton")?.addEventListener("click", saveAssignment);
  byId("chatSendButton")?.addEventListener("click", sendChatComment);
  byId("chatInput")?.addEventListener("focus", () => {
    document.querySelectorAll(".drawer-inline-form").forEach(x=>x.classList.remove("open"));
  });
  byId("reminderAddButton")?.addEventListener("click", addReminderFromDrawer);
  byId("messageAddButton")?.addEventListener("click", addCorrespondenceFromDrawer);
  byId("reminderToggleButton")?.addEventListener("click", () => setDrawerComposerMode("reminder"));
  byId("messageToggleButton")?.addEventListener("click", () => setDrawerComposerMode("message"));
  document.querySelectorAll("[data-drawer-form-close]").forEach(b => b.addEventListener("click", () => closeDrawerForm(b.dataset.drawerFormClose)));
  byId("erpChooseButton")?.addEventListener("click", () => byId("erpFile")?.click());
  byId("erpFile")?.addEventListener("change", onErpFileSelected);
  byId("erpImportButton")?.addEventListener("click", importErpPreview);
  document.querySelectorAll("[data-kpi-filter]").forEach(el => el.addEventListener("click", () => toggleKpiFilter(el.dataset.kpiFilter || "")));
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
    setText("pageTitle", `РБД 3.3.4 — ${STATE.actor}`);
    setText("pageSubtitle", "Кабінет керівника • усі заявки • ERP • контроль");
    setText("currentUserName", STATE.actor);
    setText("currentUserRole", STATE.role || "Керівник");
    setText("userAvatar", initials(STATE.actor));
    show(byId("filterExecutorField"), true);
    show(byId("createExecutorField"), true);
    show(byId("distributionButton"), true);
    show(byId("erpManagerNotice"), false);
    show(byId("dictionaryManagerNotice"), false);
    document.querySelectorAll(".dict-add").forEach(x => x.disabled = false);
    if (notice) notice.style.display = "none";
    if (byId("erpImportButton")) byId("erpImportButton").disabled = STATE.erpPreview.filter(x => x.state === "new").length === 0;
  } else if (STATE.mode === "employee") {
    setText("pageTitle", `РБД 3.3.4 — ${STATE.actor}`);
    setText("pageSubtitle", "Мої заявки • опрацювання • хронологія");
    setText("currentUserName", STATE.actor);
    setText("currentUserRole", STATE.role || "Виконавець");
    setText("userAvatar", initials(STATE.actor));
    show(byId("filterExecutorField"), false);
    show(byId("createExecutorField"), false);
    show(byId("distributionButton"), false);
    show(byId("erpManagerNotice"), true);
    show(byId("dictionaryManagerNotice"), true);
    document.querySelectorAll(".dict-add").forEach(x => x.disabled = true);
    if (byId("erpImportButton")) byId("erpImportButton").disabled = true;
    if (notice) notice.style.display = "none";
  } else {
    setText("pageTitle", "РБД 3.3.4 — Центр заявок");
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
    assignments: await readTable(context, TABLES.assignments, false),
    messages: await readTable(context, TABLES.messages, false),
    reminders: await readTable(context, TABLES.reminders, false)
  }));

  STATE.requestsAll = rowsToObjects(data.requests).map(requestFromObject).filter(x => x.id);
  STATE.timeline = rowsToObjects(data.timeline);
  STATE.notifications = rowsToObjects(data.notifications);
  STATE.legacyEvents = rowsToObjects(data.legacyEvents);
  STATE.assignments = rowsToObjects(data.assignments);
  STATE.messages = rowsToObjects(data.messages);
  STATE.reminders = rowsToObjects(data.reminders);

  if (STATE.mode === "manager") STATE.requests = STATE.requestsAll.slice();
  else if (STATE.mode === "employee") STATE.requests = STATE.requestsAll.filter(r => sameName(r.executor, STATE.actor));
  else STATE.requests = [];

  STATE.expandedStatusId = "";
  STATE.expandedAssignId = "";
  STATE.expandedFactId = "";
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
    customer: clean(x["Замовник"]), city: clean(x["Місто"]), location: clean(x["Локація"]), category: clean(x["Категорія"]), priority: clean(x["Пріоритет"]), budget: num(x["Бюджет заявки, грн"]), actualSpent: num(x["Фактично використано, грн"]), executionFact: clean(x["Факт виконання"]), factDate: x["Дата внесення факту"], factBy: clean(x["Факт вніс"]),
    status: clean(x["Статус"]), executor: clean(x["Виконавець"]), createdBy: clean(x["Створив"]), createdAt: x["Дата створення"], updatedAt: x["Дата оновлення"],
    version: num(x["Версія"]), source: clean(x["Джерело"]), statusSince: x["Дата зміни статусу"], completedAt: x["Дата завершення"],
    erpStatus: clean(x["Статус ERP"]), erpExecutor: clean(x["Виконавець ERP"]), erpSite: clean(x["Сайт ERP"]),
    erpLevel: clean(x["Рівень ERP"]), erpCloseCode: clean(x["Код закриття ERP"]), erpUpdatedAt: x["Дата оновлення ERP"]
  };
}

function renderAll() {
  updateKpis();
  renderQuickStatuses();
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
  fillSelect("reminderRecipient", STATE.users.map(x => clean(x["ПІБ"])), "Отримувач");
  fillLocationSelect("createLocation", value("createCity"), "Оберіть локацію");
}

function fillLocationSelect(id, city, placeholder) {
  const list = STATE.locations.filter(x => !city || sameName(x["Місто"], city)).map(x => clean(x["Назва"]));
  fillSelect(id, list, placeholder);
}

function applyFilters() {
  const search = norm(value("searchInput"));
  const status = STATE.quickStatus || value("filterStatus");
  const executor = value("filterExecutor");
  const city = value("filterCity");
  const priority = value("filterPriority");

  STATE.filtered = STATE.requests.filter(r => {
    if (status && r.status !== status) return false;
    if (STATE.kpiFilter === "active" && isFinalStatus(r.status)) return false;
    if (STATE.kpiFilter === "overdue" && !isOverdue(r)) return false;
    if (STATE.kpiFilter === "work" && (isFinalStatus(r.status) || ["Нова","Призупинена"].includes(r.status))) return false;
    if (STATE.kpiFilter === "closed30") {
      const cutoff = Date.now() - 30*86400000;
      if (!(isFinalStatus(r.status) && dateMs(r.completedAt) >= cutoff)) return false;
    }
    if (STATE.mode === "manager" && executor && !sameName(r.executor, executor)) return false;
    if (city && r.city !== city) return false;
    if (priority && r.priority !== priority) return false;
    if (search && !norm([r.id,r.erpId,r.subject,r.description,r.customer,r.city,r.location,r.category,r.priority,r.status,r.executor,r.budget,r.actualSpent,r.executionFact].join(" ")).includes(search)) return false;
    return true;
  });
  renderRequests();
  const budget = sumBudget(STATE.filtered), actual = sumActual(STATE.filtered);
  setText("requestsSummary", `Показано ${STATE.filtered.length} із ${STATE.requests.length} заявок • бюджет ${fmtMoney(budget)} • факт ${fmtMoney(actual)}`);
  setText("filteredBudget", fmtMoney(budget));
  setText("filteredActual", fmtMoney(actual));
  document.querySelectorAll("[data-kpi-filter]").forEach(el => el.classList.toggle("active", el.dataset.kpiFilter === STATE.kpiFilter));
}

function updateKpis() {
  const active = STATE.requests.filter(r => !isFinalStatus(r.status));
  const overdue = active.filter(isOverdue);
  const work = active.filter(r => !["Нова","Призупинена"].includes(r.status));
  const cutoff = Date.now() - 30*86400000;
  const closed = STATE.requests.filter(r => isFinalStatus(r.status) && dateMs(r.completedAt) >= cutoff);
  setText("kpiActive", active.length); setText("kpiActiveSub", `${fmtMoney(sumBudget(active))} • ${STATE.requests.filter(r=>!r.executor).length} без виконавця`);
  setText("kpiOverdue", overdue.length); setText("kpiOverdueSub", overdue.length ? fmtMoney(sumBudget(overdue)) : "Прострочень немає");
  setText("kpiWork", work.length); setText("kpiWorkSub", fmtMoney(sumBudget(work)));
  setText("kpiClosed", closed.length); setText("kpiClosedSub", fmtMoney(sumBudget(closed)));
}

function renderRequests() {
  const body = byId("requestsBody"); if (!body) return;
  if (!STATE.filtered.length) { body.innerHTML = `<tr><td colspan="14"><div class="empty">Заявок за обраними умовами немає</div></td></tr>`; return; }
  let html = "";
  for (const r of STATE.filtered) {
    const st = getStatus(r.status); const bg = validHex(st?.["Колір фону"]) ? st["Колір фону"] : "#ffffff"; const fg = validHex(st?.["Колір тексту"]) ? st["Колір тексту"] : "#173652";
    const hours = currentStatusHours(r); const days = currentStatusDays(r); const sla = num(st?.["SLA, год"]); const overdue = isOverdue(r);
    html += `<tr class="request-row" data-row-status="${esc(r.id)}" style="background:${overdue?"#ffb5ba":bg}22">
      <td><div class="request-id">${esc(r.id)}</div></td><td>${esc(r.erpId||"—")}</td><td><div class="ellipsis" title="${esc(r.subject)}">${esc(r.subject||"—")}</div></td>
      <td>${esc(r.category||"—")}</td><td>${esc(r.city||"—")}</td><td><div class="ellipsis" title="${esc(r.customer)}">${esc(r.customer||"—")}</div></td><td>${STATE.mode==="manager"?`<button class="executor-link" data-assign-inline="${esc(r.id)}">${esc(r.executor||"Без виконавця")}</button>`:esc(r.executor||"—")}</td>
      <td class="right bold">${esc(fmtMoney(r.budget))}</td><td class="right bold ${r.budget&&r.actualSpent>r.budget?"over-budget":""}">${esc(fmtMoney(r.actualSpent))}</td><td>${priorityPill(r.priority)}</td><td><button class="status-pill" data-status-open="${esc(r.id)}" style="background:${bg};color:${fg};border:0"><span class="dot"></span>${esc(r.status||"—")}</button></td>
      <td class="${overdue?"overdue":""}" title="${r.statusSince?`З ${esc(fmtDateTime(r.statusSince))} • ${hours} роб. год.`:""}">${days===null?"—":formatWorkDays(days)}</td><td>${sla?`${sla} год.`:"—"}</td>
      <td><div class="action-row"><button class="icon-btn" data-open-request="${esc(r.id)}" title="Деталі та історія">◉</button><button class="icon-btn" data-fact="${esc(r.id)}" title="Факт виконання">✓₴</button>${STATE.mode==="manager"?`<button class="icon-btn" data-assign="${esc(r.id)}" title="Призначити">♙</button>`:""}</div></td></tr>`;
    if (STATE.expandedStatusId === r.id) html += statusPanelHtml(r);
    if (STATE.expandedAssignId === r.id && STATE.mode==="manager") html += assignmentPanelHtml(r);
    if (STATE.expandedFactId === r.id) html += factPanelHtml(r);
  }
  body.innerHTML = html;
  bindRequestEvents();
}

function statusPanelHtml(r) {
  const names=uniqueStatusNames();
  const buttons=names.map(name=>{
    const st=getStatus(name);
    const bg=validHex(st?.["Колір фону"])?st["Колір фону"]:"#eef3f7";
    const fg=validHex(st?.["Колір тексту"])?st["Колір тексту"]:"#29445b";
    const current=name===r.status;
    if(current) return `<button class="status-option current" disabled style="background:${bg};color:${fg}">✓ ${esc(name)}</button>`;
    return `<button class="status-option" data-status-change="${esc(name)}" data-request-id="${esc(r.id)}" style="background:${bg};color:${fg}">${esc(name)}</button>`;
  }).join("");
  return `<tr class="status-panel-row"><td colspan="13"><div class="status-panel">
    <div class="inline-panel-title">Поточний статус: <strong>${esc(r.status)}</strong></div>
    <textarea class="status-comment" id="statusComment-${esc(r.id)}" placeholder="Коментар до зміни статусу — необов’язково"></textarea>
    <div class="status-options">${buttons}<button class="status-option close-inline" data-status-close="1">× Закрити</button></div>
  </div></td></tr>`;
}

function assignmentPanelHtml(r) {
  const options=STATE.users.map(u=>clean(u["ПІБ"])).filter(Boolean).map(n=>`<option value="${esc(n)}" ${sameName(n,r.executor)?"selected":""}>${esc(n)}</option>`).join("");
  return `<tr class="assign-panel-row"><td colspan="13"><div class="assign-inline-panel">
    <div class="inline-panel-title">Відповідальний — ${esc(r.id)}</div>
    <select id="assignInlineExecutor-${esc(r.id)}" class="assign-inline-select">${options}</select>
    <input id="assignInlineComment-${esc(r.id)}" class="assign-inline-comment" placeholder="Коментар — необов’язково">
    <button class="btn primary" data-assign-inline-save="${esc(r.id)}">Призначити</button>
    <button class="btn" data-assign-inline-close="1">×</button>
  </div></td></tr>`;
}

function factPanelHtml(r){
  return `<tr class="fact-panel-row"><td colspan="14"><div class="fact-inline-panel">
    <div class="fact-inline-head">
      <div><div class="inline-panel-title">Факт виконання — ${esc(r.id)}</div><div class="fact-inline-sub">Бюджет: <strong>${esc(fmtMoney(r.budget))}</strong></div></div>
      <button class="btn" data-fact-close="1">×</button>
    </div>
    <div class="fact-inline-grid">
      <div class="field"><label>Фактично використано, грн</label><input id="factAmount-${esc(r.id)}" type="number" min="0" step="0.01" value="${num(r.actualSpent)||""}" placeholder="0"></div>
      <div class="field fact-text-field"><label>Факт виконання / результат</label><textarea id="factText-${esc(r.id)}" placeholder="Що фактично виконано?">${esc(r.executionFact||"")}</textarea></div>
      <button class="btn primary fact-save-btn" data-fact-save="${esc(r.id)}">Зберегти факт</button>
    </div>
  </div></td></tr>`;
}

function toggleFactPanel(id){
  const r=findRequest(id); if(!r)return;
  if(STATE.mode!=="manager" && !sameName(r.executor,STATE.actor)) return toast("Факт можна вносити лише по власній заявці");
  STATE.expandedStatusId="";
  STATE.expandedAssignId="";
  STATE.expandedFactId=STATE.expandedFactId===id?"":id;
  renderRequests();
}

function toggleInlineAssignment(id){
  if(STATE.mode!=="manager") return;
  STATE.expandedStatusId="";
  STATE.expandedFactId="";
  STATE.expandedAssignId=STATE.expandedAssignId===id?"":id;
  renderRequests();
}
function bindRequestEvents() {
  document.querySelectorAll("[data-open-request]").forEach(b => b.addEventListener("click", e => { e.stopPropagation(); openRequest(findRequest(b.dataset.openRequest)); }));
  document.querySelectorAll("[data-assign]").forEach(b => b.addEventListener("click", e => { e.stopPropagation(); toggleInlineAssignment(b.dataset.assign); }));
  document.querySelectorAll("[data-assign-inline]").forEach(b => b.addEventListener("click", e => { e.stopPropagation(); toggleInlineAssignment(b.dataset.assignInline); }));
  document.querySelectorAll("[data-fact]").forEach(b => b.addEventListener("click", e => { e.stopPropagation(); toggleFactPanel(b.dataset.fact); }));
  document.querySelectorAll("[data-fact-close]").forEach(b => b.addEventListener("click", e => { e.stopPropagation(); STATE.expandedFactId=""; renderRequests(); }));
  document.querySelectorAll("[data-fact-save]").forEach(b => b.addEventListener("click", async e => {
    e.stopPropagation();
    const r=findRequest(b.dataset.factSave); if(!r)return;
    const amount=num(value(`factAmount-${r.id}`));
    const factText=value(`factText-${r.id}`);
    if(!factText && amount===0) return toast("Вкажіть факт виконання або фактичну суму");
    try{
      b.disabled=true;b.textContent="Зберігаємо...";
      await directSaveFact(r,amount,factText);
      STATE.expandedFactId="";
      await loadBusinessData(); renderAll();
      toast("Факт виконання збережено");
    }catch(err){toast("Помилка: "+errorText(err));}
    finally{b.disabled=false;b.textContent="Зберегти факт";}
  }));
  document.querySelectorAll("[data-status-open]").forEach(b => b.addEventListener("click", e => { e.stopPropagation(); const id=b.dataset.statusOpen; STATE.expandedAssignId=""; STATE.expandedFactId=""; STATE.expandedStatusId = STATE.expandedStatusId===id ? "" : id; renderRequests(); }));
  document.querySelectorAll("[data-row-status]").forEach(row => row.addEventListener("click", e => {
    if (e.target.closest("button,input,select,textarea,a")) return;
    const id=row.dataset.rowStatus;
    STATE.expandedAssignId="";
    STATE.expandedFactId="";
    STATE.expandedStatusId = STATE.expandedStatusId===id ? "" : id;
    renderRequests();
  }));
  document.querySelectorAll("[data-status-close]").forEach(b => b.addEventListener("click", () => { STATE.expandedStatusId=""; renderRequests(); }));
  document.querySelectorAll("[data-assign-inline-close]").forEach(b => b.addEventListener("click", e => { e.stopPropagation(); STATE.expandedAssignId=""; renderRequests(); }));
  document.querySelectorAll("[data-assign-inline-save]").forEach(b => b.addEventListener("click", async e => {
    e.stopPropagation(); const r=findRequest(b.dataset.assignInlineSave); if(!r)return;
    const executor=value(`assignInlineExecutor-${r.id}`); const comment=value(`assignInlineComment-${r.id}`);
    if(!executor) return toast("Оберіть відповідального");
    try { b.disabled=true; b.textContent="..."; await directAssign(r,executor,comment); STATE.expandedAssignId=""; await loadBusinessData(); renderAll(); toast("Відповідального змінено"); }
    catch(err){ toast("Помилка: "+errorText(err)); }
    finally { b.disabled=false; b.textContent="Призначити"; }
  }));
  document.querySelectorAll("[data-status-change]").forEach(b => b.addEventListener("click", async () => {
    const r=findRequest(b.dataset.requestId); const next=b.dataset.statusChange; if(!r)return;
    const comment = clean(byId(`statusComment-${r.id}`)?.value || "");
    if (next === "Закрита" && !window.confirm(`Закрити заявку ${r.id}?`)) return;
    try {
      b.disabled = true;
      await directStatusChange(r,next,comment);
      STATE.expandedStatusId="";
      await loadBusinessData();
      renderAll();
      toast(`Статус змінено: ${next}`);
    } catch(e){toast(`Помилка: ${errorText(e)}`);} finally { b.disabled=false; }
  }));
}


function openDistributionModal(){
  if(STATE.mode!=="manager") return;
  STATE.distributionSelected=[];
  STATE.distributionExecutor="";
  STATE.distributionShowAll=false;
  STATE.distributionSearch="";
  setInput("distributionSearch","");
  const showAll=byId("distributionShowAll"); if(showAll) showAll.checked=false;
  renderDistribution();
  openModal("distributionModal");
}

function distributionRequests(){
  const q=norm(STATE.distributionSearch);
  return STATE.requestsAll.filter(r=>{
    if(isFinalStatus(r.status)) return false;
    if(!STATE.distributionShowAll && r.executor) return false;
    if(q && !norm([r.id,r.erpId,r.subject,r.customer,r.city,r.category,r.executor].join(" ")).includes(q)) return false;
    return true;
  });
}

function renderDistribution(){
  const body=byId("distributionRequests"); const people=byId("distributionExecutors");
  if(!body||!people) return;
  const rows=distributionRequests();
  setText("distributionCount", `${rows.length} заявок`);
  body.innerHTML=rows.length?rows.map(r=>{
    const checked=STATE.distributionSelected.includes(r.id);
    return `<label class="distribution-request ${checked?"selected":""}">
      <input type="checkbox" data-distribution-request="${esc(r.id)}" ${checked?"checked":""}>
      <div class="distribution-request-main">
        <div class="distribution-request-top"><strong>${esc(r.id)}</strong><span>${esc(r.status||"—")}</span></div>
        <div class="distribution-request-title">${esc(r.subject||"Без теми")}</div>
        <div class="distribution-request-meta">${esc(r.city||"—")} • ${esc(r.category||"—")} • ${esc(r.customer||"—")}</div>
        ${r.executor?`<div class="distribution-current">Зараз: ${esc(r.executor)}</div>`:""}
      </div>
    </label>`;
  }).join(""):`<div class="distribution-empty">Немає заявок для розподілу</div>`;

  const activeBy=n=>STATE.requestsAll.filter(r=>sameName(r.executor,n)&&!isFinalStatus(r.status));
  people.innerHTML=STATE.users.map(u=>{
    const n=clean(u["ПІБ"]); const active=activeBy(n); const overdue=active.filter(isOverdue).length;
    const selected=sameName(STATE.distributionExecutor,n);
    return `<button class="distribution-person ${selected?"selected":""}" data-distribution-executor="${esc(n)}">
      <div class="distribution-avatar">${esc(initials(n))}</div>
      <div class="distribution-person-text">
        <strong>${esc(n)}</strong>
        <span>${active.length} активних • ${fmtMoney(sumBudget(active))}${overdue?` • прострочено ${overdue}`:""}</span>
      </div>
    </button>`;
  }).join("");

  body.querySelectorAll("[data-distribution-request]").forEach(ch=>ch.addEventListener("change",()=>{
    const id=ch.dataset.distributionRequest;
    if(ch.checked){ if(!STATE.distributionSelected.includes(id)) STATE.distributionSelected.push(id); }
    else STATE.distributionSelected=STATE.distributionSelected.filter(x=>x!==id);
    renderDistribution();
  }));
  people.querySelectorAll("[data-distribution-executor]").forEach(btn=>btn.addEventListener("click",()=>{
    STATE.distributionExecutor=btn.dataset.distributionExecutor||"";
    renderDistribution();
  }));

  const assign=byId("distributionAssignButton");
  const n=STATE.distributionSelected.length;
  if(assign){
    assign.disabled=!n||!STATE.distributionExecutor;
    assign.textContent=n?`Призначити вибрані (${n})`:"Призначити вибрані";
  }
  setText("distributionSelectedExecutor", STATE.distributionExecutor || "Оберіть відповідального");
}

async function assignDistributionSelection(){
  if(STATE.mode!=="manager") return;
  const ids=STATE.distributionSelected.slice();
  const executor=STATE.distributionExecutor;
  if(!ids.length||!executor) return;
  const btn=byId("distributionAssignButton");
  try{
    if(btn){btn.disabled=true;btn.textContent="Розподіляємо...";}
    let done=0, skipped=0;
    for(const id of ids){
      const r=findRequest(id); if(!r) continue;
      if(sameName(r.executor,executor)){ skipped++; continue; }
      await directAssign(r,executor,"Розподіл заявок керівником");
      done++;
    }
    await loadBusinessData(); renderAll();
    STATE.distributionSelected=[];
    renderDistribution();
    toast(`Розподілено: ${done}${skipped?`, без змін: ${skipped}`:""}`);
  }catch(e){ toast("Помилка розподілу: "+errorText(e)); }
  finally{
    if(btn){btn.disabled=false;btn.textContent="Призначити вибрані";}
  }
}

function openCreateModal() {
  if (STATE.mode === "unknown") return toast("Спочатку відкрийте персональну вкладку працівника");
  ["createSubject","createDescription","createComment","createBudget"].forEach(id => setInput(id,""));
  setInput("createCity",""); fillLocationSelect("createLocation","","Оберіть локацію"); setInput("createCategory",""); setInput("createCustomer","");
  setInput("createPriority", STATE.priorities.find(x=>clean(x["Назва"])==="Середній") ? "Середній" : "");
  setInput("createExecutor", STATE.mode === "employee" ? STATE.actor : ""); clearMessage("createMessage"); openModal("createModal");
}

async function saveCreate() {
  const subject=value("createSubject"); if(!subject) return showMessage("createMessage","error","Вкажіть тему заявки.");
  const executor=STATE.mode==="employee"?STATE.actor:value("createExecutor");
  try {
    setBusy("saveCreateButton",true,"Створюємо...");
    const id = await directCreateRequest({
      subject, description:value("createDescription"), executor,
      city:value("createCity"), location:value("createLocation"), category:value("createCategory"),
      customer:value("createCustomer"), priority:value("createPriority"), comment:value("createComment"),
      budget:num(value("createBudget"))
    });
    closeModal("createModal");
    await loadBusinessData(); renderAll();
    toast(`Створено заявку ${id}`);
  } catch(e){showMessage("createMessage","error",errorText(e));} finally{setBusy("saveCreateButton",false,"Створити");}
}

function openAssignment(r) {
  if (!r) return; if (STATE.mode !== "manager") return toast("Призначення іншого виконавця доступне керівнику");
  STATE.selectedRequest=r; setText("assignTitle",`Призначення — ${r.id}`); setInput("assignExecutor",r.executor||""); setInput("assignComment",""); clearMessage("assignMessage"); openModal("assignModal");
}

async function saveAssignment() {
  const r=STATE.selectedRequest;if(!r)return;const executor=value("assignExecutor");if(!executor)return showMessage("assignMessage","error","Оберіть виконавця.");
  try{
    setBusy("saveAssignButton",true,"Призначаємо...");
    await directAssign(r,executor,value("assignComment"));
    closeModal("assignModal");
    await loadBusinessData(); renderAll();
    toast("Виконавця змінено");
  } catch(e){showMessage("assignMessage","error",errorText(e));} finally{setBusy("saveAssignButton",false,"Призначити");}
}


function normalizeSurname(value) {
  return clean(value).replace(/[«»"'\`]/g,"").replace(/\u00A0/g," ").trim().split(/\s+/)[0].replace(/[.,;:]+$/g,"").toLocaleLowerCase("uk-UA");
}
function fmtMoney(v){ return new Intl.NumberFormat("uk-UA",{maximumFractionDigits:2}).format(num(v))+" грн"; }
function sumBudget(rows){ return rows.reduce((a,r)=>a+num(r.budget),0); }
function sumActual(rows){ return rows.reduce((a,r)=>a+num(r.actualSpent),0); }
function varianceMoney(r){ return num(r.budget)-num(r.actualSpent); }
function uniqueStatusNames(){
  const canonical=["Нова","Прийнята в роботу","Пошук підрядника","Погодження кошторису","Укладення договору","Погодження бюджету","Виконання робіт","Приймання робіт","Призупинена","Закрита"];
  const raw=[...canonical,...STATE.statuses.map(x=>clean(x["Назва статусу"])).filter(Boolean)];
  const seen=new Set();
  const result=[];
  for(const name of raw){
    const key=norm(name).replace(/[\u200B-\u200D\uFEFF]/g,"").replace(/\s+/g," ");
    if(!key || seen.has(key)) continue;
    seen.add(key);
    result.push(clean(name).replace(/[\u200B-\u200D\uFEFF]/g,""));
  }
  return result;
}

function renderQuickStatuses(){
  const el=byId("quickStatuses"); if(!el)return;
  const names=uniqueStatusNames();
  const count=n=>STATE.requests.filter(r=>r.status===n).length;

  const allButton=`<button class="status-filter status-filter-all ${STATE.quickStatus===""?"active":""}" data-quick-status="">
    <span class="status-filter-label">Усі</span>
    <span class="status-filter-count">${STATE.requests.length}</span>
  </button>`;

  const buttons=names.map(n=>{
    const st=getStatus(n);
    const bg=validHex(st?.["Колір фону"])?st["Колір фону"]:"#eef3f7";
    const fg=validHex(st?.["Колір тексту"])?st["Колір тексту"]:"#29445b";
    return `<button class="status-filter ${STATE.quickStatus===n?"active":""}" data-quick-status="${esc(n)}" style="background:${bg};color:${fg};border-color:${fg}55">
      <span class="status-filter-dot" style="background:${fg}"></span>
      <span class="status-filter-label">${esc(n)}</span>
      <span class="status-filter-count">${count(n)}</span>
    </button>`;
  }).join("");

  el.innerHTML=allButton+buttons;

  el.querySelectorAll("[data-quick-status]").forEach(b=>b.addEventListener("click",()=>{
    STATE.quickStatus=b.dataset.quickStatus||"";
    if(STATE.quickStatus) setInput("filterStatus","");
    renderQuickStatuses();
    applyFilters();
  }));
}
function toggleKpiFilter(name){ STATE.kpiFilter = STATE.kpiFilter===name ? "" : name; applyFilters(); }

async function getTableMeta(context,name){
  const t=context.workbook.tables.getItem(name); const h=t.getHeaderRowRange(); h.load("values"); t.rows.load("items"); await context.sync();
  return {table:t,headers:h.values[0].map(String)};
}
function mapRow(headers,obj){ return headers.map(h=>obj[h]??""); }
function nextRequestIdFromRows(headers,rows){
  const idx=headers.indexOf("ID заявки"); let max=0;
  for(const row of rows){ const m=/^RBD-(\d+)$/.exec(clean(row[idx])); if(m) max=Math.max(max,Number(m[1])); }
  return "RBD-"+String(max+1).padStart(6,"0");
}
async function directCreateRequest(p){
  if(STATE.mode==="unknown") throw new Error("Не визначено кабінет користувача");
  return Excel.run(async context=>{
    const req=await getTableMeta(context,TABLES.requests);
    const tl=await getTableMeta(context,TABLES.timeline);
    const asg=await getTableMeta(context,TABLES.assignments);
    const nt=await getTableMeta(context,TABLES.notifications);
    let rows=[]; if(req.table.rows.items.length){ const b=req.table.getDataBodyRange(); b.load("values"); await context.sync(); rows=b.values; }
    if(!req.headers.includes("Бюджет заявки, грн")) throw new Error("У таблиці Заявки немає колонки «Бюджет заявки, грн». Запустіть 01_RBD_Setup_New.");
    const id=nextRequestIdFromRows(req.headers,rows), now=localTimestamp(), owner=p.executor||"";
    const obj={
      "ID заявки":id,"ID ERP":"","Тема заявки":p.subject,"Опис заявки":p.description||"",
      "Замовник":p.customer||"","Місто":p.city||"","Локація":p.location||"","Категорія":p.category||"",
      "Пріоритет":p.priority||"Середній","Бюджет заявки, грн":num(p.budget),"Фактично використано, грн":0,"Факт виконання":"","Дата внесення факту":"","Факт вніс":"","Статус":"Нова","Виконавець":owner,
      "Створив":STATE.actor,"Дата створення":now,"Дата оновлення":now,"Версія":1,"Джерело":"РБД",
      "Дата зміни статусу":now,"Дата завершення":""
    };
    req.table.rows.add(null,[mapRow(req.headers,obj)]);
    tl.table.rows.add(null,[mapRow(tl.headers,{"ID події":createId("EV"),"ID заявки":id,"Дата події":now,"Автор":STATE.actor,"Тип":"Створення","Дія":"Створено заявку","Текст":p.comment||p.subject,"Було":"","Стало":"Нова","Джерело":"Надбудова"})]);
    if(owner){
      asg.table.rows.add(null,[mapRow(asg.headers,{"ID призначення":createId("AS"),"ID заявки":id,"Попередній виконавець":"","Новий виконавець":owner,"Призначив":STATE.actor,"Дата призначення":now,"Причина":"Створення заявки"})]);
      const email=clean(STATE.users.find(u=>sameName(u["ПІБ"],owner))?.["Електронна пошта"]||"");
      nt.table.rows.add(null,[mapRow(nt.headers,{"ID сповіщення":createId("NTF"),"ID заявки":id,"ID події":"","Отримувач":owner,"Email":email,"Канал":"Email","Тема":`Нова заявка ${id}`,"Текст":`${p.subject}\nВас призначено виконавцем.`,"Дата створення":now,"Статус":email?"Очікує":"Немає email","Спроби":0})]);
    }
    await context.sync(); return id;
  });
}
async function directAssign(r,executor,comment){
  if(STATE.mode!=="manager") throw new Error("Призначати виконавця може лише керівник");
  return Excel.run(async context=>{
    const req=await getTableMeta(context,TABLES.requests), tl=await getTableMeta(context,TABLES.timeline), asg=await getTableMeta(context,TABLES.assignments), nt=await getTableMeta(context,TABLES.notifications);
    const body=req.table.getDataBodyRange(); body.load("values"); await context.sync();
    const rows=body.values, idIdx=req.headers.indexOf("ID заявки"), rowIdx=rows.findIndex(x=>clean(x[idIdx])===r.id); if(rowIdx<0) throw new Error("Заявку не знайдено");
    const old=clean(rows[rowIdx][req.headers.indexOf("Виконавець")]); if(sameName(old,executor)) throw new Error("Цей виконавець уже призначений");
    const now=localTimestamp(), range=body.getRow(rowIdx);
    range.values=[req.headers.map((h,i)=> h==="Виконавець"?executor : h==="Дата оновлення"?now : h==="Версія"?num(rows[rowIdx][i])+1 : rows[rowIdx][i])];
    asg.table.rows.add(null,[mapRow(asg.headers,{"ID призначення":createId("AS"),"ID заявки":r.id,"Попередній виконавець":old,"Новий виконавець":executor,"Призначив":STATE.actor,"Дата призначення":now,"Причина":comment||""})]);
    const ev=createId("EV");
    tl.table.rows.add(null,[mapRow(tl.headers,{"ID події":ev,"ID заявки":r.id,"Дата події":now,"Автор":STATE.actor,"Тип":"Призначення","Дія":"Змінено виконавця","Текст":comment||"","Було":old,"Стало":executor,"Джерело":"Надбудова"})]);
    const email=clean(STATE.users.find(u=>sameName(u["ПІБ"],executor))?.["Електронна пошта"]||"");
    nt.table.rows.add(null,[mapRow(nt.headers,{"ID сповіщення":createId("NTF"),"ID заявки":r.id,"ID події":ev,"Отримувач":executor,"Email":email,"Канал":"Email","Тема":`Призначено заявку ${r.id}`,"Текст":`Вас призначено виконавцем заявки «${r.subject}».`,"Дата створення":now,"Статус":email?"Очікує":"Немає email","Спроби":0})]);
    await context.sync();
  });
}
async function directStatusChange(r,next,comment){
  if(STATE.mode!=="manager" && !sameName(r.executor,STATE.actor)) throw new Error("Можна змінювати лише власні заявки");
  return Excel.run(async context=>{
    const req=await getTableMeta(context,TABLES.requests), tl=await getTableMeta(context,TABLES.timeline), nt=await getTableMeta(context,TABLES.notifications);
    const body=req.table.getDataBodyRange(); body.load("values"); await context.sync();
    const rows=body.values, idIdx=req.headers.indexOf("ID заявки"), rowIdx=rows.findIndex(x=>clean(x[idIdx])===r.id); if(rowIdx<0) throw new Error("Заявку не знайдено");
    const now=localTimestamp(), range=body.getRow(rowIdx), completed=isFinalStatus(next);
    range.values=[req.headers.map((h,i)=> h==="Статус"?next : h==="Дата зміни статусу"?now : h==="Дата завершення"?(completed?now:"") : h==="Дата оновлення"?now : h==="Версія"?num(rows[rowIdx][i])+1 : rows[rowIdx][i])];
    const ev=createId("EV");
    const statusStart=r.statusSince||"";
    const statusHours=workingHoursBetween(statusStart,now);
    const statusDays=Math.round((statusHours/9)*10)/10;
    tl.table.rows.add(null,[mapRow(tl.headers,{"ID події":ev,"ID заявки":r.id,"Дата події":now,"Автор":STATE.actor,"Тип":"Статус","Дія":"Змінено статус","Текст":comment||"","Було":r.status,"Стало":next,"Початок статусу":statusStart,"Кінець статусу":now,"Тривалість статусу, год":statusHours,"Тривалість статусу, роб. днів":statusDays,"Джерело":"Надбудова"})]);
    if(r.executor && !sameName(r.executor,STATE.actor)){
      const email=clean(STATE.users.find(u=>sameName(u["ПІБ"],r.executor))?.["Електронна пошта"]||"");
      nt.table.rows.add(null,[mapRow(nt.headers,{"ID сповіщення":createId("NTF"),"ID заявки":r.id,"ID події":ev,"Отримувач":r.executor,"Email":email,"Канал":"Email","Тема":`Статус заявки ${r.id}`,"Текст":`${r.status} → ${next}${comment?"\n"+comment:""}`,"Дата створення":now,"Статус":email?"Очікує":"Немає email","Спроби":0})]);
    }
    await context.sync();
  });
}
async function directComment(r,comment){
  if(STATE.mode!=="manager" && !sameName(r.executor,STATE.actor)) throw new Error("Не можна коментувати чужу заявку");
  return Excel.run(async context=>{
    const req=await getTableMeta(context,TABLES.requests), tl=await getTableMeta(context,TABLES.timeline);
    const body=req.table.getDataBodyRange(); body.load("values"); await context.sync();
    const rows=body.values,idIdx=req.headers.indexOf("ID заявки"),rowIdx=rows.findIndex(x=>clean(x[idIdx])===r.id); if(rowIdx<0) throw new Error("Заявку не знайдено");
    const now=localTimestamp(), range=body.getRow(rowIdx);
    range.values=[req.headers.map((h,i)=> h==="Дата оновлення"?now : h==="Версія"?num(rows[rowIdx][i])+1 : rows[rowIdx][i])];
    tl.table.rows.add(null,[mapRow(tl.headers,{"ID події":createId("EV"),"ID заявки":r.id,"Дата події":now,"Автор":STATE.actor,"Тип":"Коментар","Дія":"Коментар","Текст":comment,"Джерело":"Надбудова"})]);
    await context.sync();
  });
}

async function directSaveFact(r,amount,factText){
  if(STATE.mode!=="manager" && !sameName(r.executor,STATE.actor)) throw new Error("Можна вносити факт лише по власній заявці");
  return Excel.run(async context=>{
    const req=await getTableMeta(context,TABLES.requests);
    const tl=await getTableMeta(context,TABLES.timeline);
    const required=["Фактично використано, грн","Факт виконання","Дата внесення факту","Факт вніс"];
    const missing=required.filter(h=>!req.headers.includes(h));
    if(missing.length) throw new Error("В таблиці Заявки відсутні колонки: "+missing.join(", ")+". Запустіть 01_RBD_Setup_New.");
    const body=req.table.getDataBodyRange(); body.load("values"); await context.sync();
    const rows=body.values, idIdx=req.headers.indexOf("ID заявки");
    const rowIdx=rows.findIndex(x=>clean(x[idIdx])===r.id);
    if(rowIdx<0) throw new Error("Заявку не знайдено");
    const now=localTimestamp(), range=body.getRow(rowIdx);
    const oldAmount=num(rows[rowIdx][req.headers.indexOf("Фактично використано, грн")]);
    const oldFact=clean(rows[rowIdx][req.headers.indexOf("Факт виконання")]);
    range.values=[req.headers.map((h,i)=>
      h==="Фактично використано, грн"?amount:
      h==="Факт виконання"?factText:
      h==="Дата внесення факту"?now:
      h==="Факт вніс"?STATE.actor:
      h==="Дата оновлення"?now:
      h==="Версія"?num(rows[rowIdx][i])+1:
      rows[rowIdx][i]
    )];
    tl.table.rows.add(null,[mapRow(tl.headers,{
      "ID події":createId("EV"),"ID заявки":r.id,"Дата події":now,"Автор":STATE.actor,
      "Тип":"Факт виконання","Дія":"Оновлено факт виконання",
      "Текст":factText||("Фактично використано "+fmtMoney(amount)),
      "Було":fmtMoney(oldAmount)+(oldFact?" • "+oldFact:""),
      "Стало":fmtMoney(amount)+(factText?" • "+factText:""),
      "Джерело":"Надбудова"
    })]);
    await context.sync();
  });
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
function isFinalStatus(name){const s=getStatus(name);return s?isYes(s["Завершальний"]):name==="Закрита";}
function workingHoursBetween(startValue,endValue){
  const startMs=dateMs(startValue), endMs=dateMs(endValue);
  if(!startMs||!endMs||endMs<=startMs) return 0;
  const start=new Date(startMs), end=new Date(endMs);
  let cursor=new Date(start.getFullYear(),start.getMonth(),start.getDate(),9,0,0,0);
  let totalMs=0;
  while(cursor<end){
    const day=cursor.getDay();
    if(day!==0&&day!==6){
      const workStart=new Date(cursor.getFullYear(),cursor.getMonth(),cursor.getDate(),9,0,0,0);
      const workEnd=new Date(cursor.getFullYear(),cursor.getMonth(),cursor.getDate(),18,0,0,0);
      const from=Math.max(startMs,workStart.getTime());
      const to=Math.min(endMs,workEnd.getTime());
      if(to>from) totalMs+=to-from;
    }
    cursor.setDate(cursor.getDate()+1);
  }
  return Math.round((totalMs/3600000)*10)/10;
}
function workingDaysBetween(startValue,endValue){
  return Math.round((workingHoursBetween(startValue,endValue)/9)*10)/10;
}
function currentStatusHours(r){
  if(!r.statusSince) return null;
  return workingHoursBetween(r.statusSince,localTimestamp());
}
function currentStatusDays(r){
  const h=currentStatusHours(r);
  return h===null?null:Math.round((h/9)*10)/10;
}
function totalRequestWorkHours(r){
  if(!r?.createdAt) return null;
  const end=r.completedAt || localTimestamp();
  return workingHoursBetween(r.createdAt,end);
}
function totalRequestWorkDays(r){
  const h=totalRequestWorkHours(r);
  return h===null?null:Math.round((h/9)*10)/10;
}
function formatWorkDays(days){
  if(days===null||days===undefined||Number.isNaN(days)) return "—";
  return new Intl.NumberFormat("uk-UA",{maximumFractionDigits:1}).format(days)+" дн.";
}
function isOverdue(r){if(isFinalStatus(r.status)||r.status==="Призупинена")return false;const sla=num(getStatus(r.status)?.["SLA, год"]);const h=currentStatusHours(r);return sla>0&&h!==null&&h>sla;}

function openRequest(r) { if(!r)return; STATE.selectedRequest=r; STATE.historyFilter="all"; renderRequestModal(r); setDrawerComposerMode("note",false); openModal("requestModal"); }
function renderRequestModal(r) {
  setText("requestModalTitle", `${r.id} — ${r.subject||"Заявка"}`);
  const h=currentStatusHours(r); const d=currentStatusDays(r); const totalDays=totalRequestWorkDays(r); const totalHours=totalRequestWorkHours(r); const sla=num(getStatus(r.status)?.["SLA, год"]);
  setText("requestTotalWorkTime", totalDays===null ? "—" : formatWorkDays(totalDays));
  const totalTimeEl=byId("requestTotalWorkTime");
  if(totalTimeEl) totalTimeEl.title=`Робочий час 09:00–18:00, Пн–Пт • ${totalHours??0} роб. год.`;
  const rows=[["ID ERP",r.erpId],["Тема",r.subject],["Опис",r.description],["Замовник",r.customer],["Місто",r.city],["Локація",r.location],["Категорія",r.category],["Бюджет заявки",fmtMoney(r.budget)],["Фактично використано",fmtMoney(r.actualSpent)],["Факт виконання",r.executionFact],["Факт вніс",r.factBy],["Дата факту",fmtDateTime(r.factDate)],["Пріоритет",r.priority],["Статус",r.status],["Виконавець",r.executor],["Створив",r.createdBy],["Створено",fmtDateTime(r.createdAt)],["Загальний час у роботі",totalDays===null?"—":formatWorkDays(totalDays)],["Днів у статусі",d===null?"—":formatWorkDays(d)],["SLA",sla?`${sla} год.`:"—"],["Джерело",r.source],["Завершено",fmtDateTime(r.completedAt)]];
  byId("requestSummary").innerHTML=rows.map(x=>`<div class="summary-label">${esc(x[0])}</div><div class="summary-value">${esc(String(x[1]||"—"))}</div>`).join("");
  renderHistoryTabs(); renderChat(r.id); setInput("chatInput",""); clearMessage("chatMessage");
}

function buildConversation(requestId) {
  const items=[];
  for (const x of STATE.timeline.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:clean(x["ID події"]),date:x["Дата події"],author:clean(x["Автор"]),type:clean(x["Тип"]),action:clean(x["Дія"]),text:clean(x["Текст"]),before:clean(x["Було"]),after:clean(x["Стало"]),statusStart:x["Початок статусу"],statusEnd:x["Кінець статусу"],statusHours:num(x["Тривалість статусу, год"]),statusDays:num(x["Тривалість статусу, роб. днів"]),channel:clean(x["Канал"]),recipient:clean(x["Отримувач"]),delivery:clean(x["Статус доставки"]),source:clean(x["Джерело"])});
  for (const x of STATE.legacyEvents.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:`LEG-${clean(x["ID події"]||x["ID команди"])}`,date:x["Дата події"],author:clean(x["Користувач"]),type:"Подія",action:clean(x["Дія"]),text:clean(x["Коментар"]),before:clean(x["Було"]),after:clean(x["Стало"]),source:clean(x["Джерело"])});
  for (const x of STATE.assignments.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:`AS-${clean(x["ID призначення"])}`,date:x["Дата призначення"],author:clean(x["Призначив"]),type:"Призначення",action:"Зміна виконавця",text:clean(x["Причина"]),before:clean(x["Попередній виконавець"]),after:clean(x["Новий виконавець"]),source:"Excel"});
  for (const x of STATE.messages.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:`MSG-${clean(x["ID листа"])}`,date:x["Дата відправлення"],author:clean(x["Відправник"]||STATE.actor),type:"Листування",action:"Листування",text:clean(x["Текст листа"]||x["Тема листа"]),recipient:clean(x["Отримувач"]),delivery:clean(x["Статус"]),channel:"Email",source:"Excel"});
  for (const x of STATE.reminders.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:`REM-${clean(x["ID нагадування"])}`,date:x["Дата створення"]||x["Дата нагадування"],author:clean(x["Створив"]||STATE.actor),type:"Нагадування",action:"Нагадування",text:clean(x["Опис"]),recipient:clean(x["Отримувач"]),delivery:clean(x["Статус"]),due:x["Дата нагадування"],source:"Excel"});
  for (const x of STATE.notifications.filter(e=>clean(e["ID заявки"])===requestId)) items.push({id:`NT-${clean(x["ID сповіщення"])}`,date:x["Дата відправки"]||x["Дата створення"],author:"Система",type:"Сповіщення",action:clean(x["Канал"]||"Повідомлення"),text:clean(x["Текст"]||x["Тема"]),channel:clean(x["Канал"]),recipient:clean(x["Отримувач"]),delivery:clean(x["Статус"]),source:"Система"});
  const seen=new Set();
  const sorted=items.filter(x=>{const k=x.id||`${x.date}|${x.author}|${x.action}|${x.text}`;if(seen.has(k))return false;seen.add(k);return true;}).sort((a,b)=>dateMs(a.date)-dateMs(b.date));

  const req=findRequest(requestId);
  let previousStatusMoment=req?.createdAt||"";
  for(const item of sorted){
    if(item.type!=="Статус" && item.action!=="Змінено статус") continue;
    if(!item.statusStart) item.statusStart=previousStatusMoment;
    if(!item.statusEnd) item.statusEnd=item.date;
    if(!item.statusHours && item.statusStart && item.statusEnd) item.statusHours=workingHoursBetween(item.statusStart,item.statusEnd);
    if(!item.statusDays && item.statusHours!==undefined) item.statusDays=Math.round((num(item.statusHours)/9)*10)/10;
    previousStatusMoment=item.date;
  }
  return sorted;
}

function renderHistoryTabs(){
  document.querySelectorAll("[data-history-filter]").forEach(b=>{
    b.classList.toggle("active",b.dataset.historyFilter===STATE.historyFilter);
    b.onclick=()=>{
      STATE.historyFilter=b.dataset.historyFilter||"all";
      setDrawerComposerMode("note",false);
      renderHistoryTabs();
      if(STATE.selectedRequest) renderChat(STATE.selectedRequest.id);
    };
  });
}
function historyItemMatches(x){
  if(STATE.historyFilter==="all") return true;
  if(STATE.historyFilter==="notes") return x.type==="Коментар";
  if(STATE.historyFilter==="reminders") return x.type==="Нагадування";
  if(STATE.historyFilter==="messages") return x.type==="Листування" || x.type==="Сповіщення";
  if(STATE.historyFilter==="status") return x.type==="Статус" || x.action==="Змінено статус";
  return true;
}
function historyStatusPill(name){
  const st=getStatus(name);
  const bg=validHex(st?.["Колір фону"])?st["Колір фону"]:"#eef3f7";
  const fg=validHex(st?.["Колір тексту"])?st["Колір тексту"]:"#29445b";
  return `<span class="history-status-pill" style="background:${bg};color:${fg}">${esc(name||"—")}</span>`;
}
function renderChat(requestId) {
  const list=byId("chatList"); if(!list)return; const items=buildConversation(requestId).filter(historyItemMatches);
  if(!items.length){list.innerHTML=`<div class="empty">У цій категорії записів ще немає</div>`;return;}
  list.innerHTML=items.map(x=>{
    const isComment=x.type==="Коментар"; const mine=sameName(x.author,STATE.actor);
    const change=(x.before||x.after)?`<div class="chat-status history-status-change">${historyStatusPill(x.before)}<span class="history-status-arrow">→</span>${historyStatusPill(x.after)}</div>`:"";
    const timing=(x.type==="Статус"||x.action==="Змінено статус")&&x.statusStart?`<div class="chat-status history-status-duration">Перебування: <strong>${formatWorkDays(x.statusDays ?? Math.round((workingHoursBetween(x.statusStart,x.statusEnd||x.date)/9)*10)/10)}</strong><br><span>${esc(fmtDateTime(x.statusStart))} → ${esc(fmtDateTime(x.statusEnd||x.date))}</span></div>`:"";
    const due=x.type==="Нагадування"&&x.due?`<div class="chat-status">На: ${esc(fmtDateTime(x.due))}</div>`:"";
    const delivery=(x.type==="Сповіщення"||x.type==="Листування"||x.type==="Нагадування")?`<div class="chat-status">${esc(x.channel||x.type)} • ${esc(x.recipient||"—")}${x.delivery?" • "+esc(x.delivery):""}</div>`:"";
    if(!isComment) return `<div class="chat-system"><strong>${esc(x.action||x.type||"Подія")}</strong>${x.author?` • ${esc(x.author)}`:""}${x.text?`<br>${esc(x.text)}`:""}${change}${timing}${due}${delivery}<div style="margin-top:3px">${esc(fmtDateTime(x.date))}</div></div>`;
    return `<div class="chat-item ${mine?"me":""}"><div class="bubble"><div class="chat-meta">${esc(x.author||"Користувач")} • ${esc(fmtDateTime(x.date))}</div><div class="chat-text">${esc(x.text||x.after||"Коментар")}</div></div></div>`;
  }).join("");
  list.scrollTop=list.scrollHeight;
}

function setDrawerComposerMode(mode, focus=true){
  document.querySelectorAll(".drawer-inline-form").forEach(x=>x.classList.remove("open"));
  byId("chatSendButton")?.classList.toggle("composer-active", mode==="note");
  byId("reminderToggleButton")?.classList.toggle("composer-active", mode==="reminder");
  byId("messageToggleButton")?.classList.toggle("composer-active", mode==="message");

  if(mode==="reminder") byId("reminderForm")?.classList.add("open");
  if(mode==="message") byId("messageForm")?.classList.add("open");
  if(mode==="note" && focus) byId("chatInput")?.focus();
}
function closeDrawerForm(id){
  byId(id)?.classList.remove("open");
  byId("reminderToggleButton")?.classList.remove("composer-active");
  byId("messageToggleButton")?.classList.remove("composer-active");
  byId("chatSendButton")?.classList.add("composer-active");
}

async function getOrCreateTableMeta(context, tableName, sheetName, headers){
  let table=context.workbook.tables.getItemOrNullObject(tableName);
  table.load("isNullObject");
  await context.sync();

  if(table.isNullObject){
    let ws=context.workbook.worksheets.getItemOrNullObject(sheetName);
    ws.load("isNullObject");
    await context.sync();
    if(ws.isNullObject) ws=context.workbook.worksheets.add(sheetName);

    const headerRange=ws.getRangeByIndexes(0,0,1,headers.length);
    headerRange.values=[headers];
    table=ws.tables.add(headerRange,true);
    table.name=tableName;
    await context.sync();
  }

  const h=table.getHeaderRowRange();
  h.load("values");
  table.rows.load("items");
  await context.sync();
  return {table,headers:h.values[0].map(String)};
}

async function addReminderFromDrawer(){
  const r=STATE.selectedRequest;if(!r)return;
  const text=value("reminderText");
  const due=value("reminderDate");
  const recipient=value("reminderRecipient")||r.executor||STATE.actor;
  if(!recipient) return toast("Оберіть отримувача");
  if(!due) return toast("Вкажіть дату та час нагадування");
  if(!text) return toast("Вкажіть текст нагадування");

  const btn=byId("reminderAddButton");
  try{
    if(btn){btn.disabled=true;btn.textContent="Зберігаємо...";}
    await Excel.run(async context=>{
      const t=await getOrCreateTableMeta(
        context,
        TABLES.reminders,
        "Нагадування",
        ["ID нагадування","ID заявки","Отримувач","Дата нагадування","Опис","Статус","Створив","Дата створення"]
      );
      t.table.rows.add(null,[mapRow(t.headers,{
        "ID нагадування":createId("REM"),
        "ID заявки":r.id,
        "Отримувач":recipient,
        "Дата нагадування":due,
        "Опис":text,
        "Статус":"Активне",
        "Створив":STATE.actor,
        "Дата створення":localTimestamp()
      })]);
      await context.sync();
    });

    setInput("reminderText","");
    setInput("reminderDate","");
    await loadBusinessData();
    STATE.selectedRequest=findRequest(r.id);
    STATE.historyFilter="reminders";
    renderRequestModal(STATE.selectedRequest);
    setDrawerComposerMode("note",false);
    toast("Нагадування збережено");
  }catch(e){
    toast("Помилка нагадування: "+errorText(e));
  }finally{
    if(btn){btn.disabled=false;btn.textContent="Зберегти";}
  }
}
async function addCorrespondenceFromDrawer(){
  const r=STATE.selectedRequest;if(!r)return;
  const recipient=value("messageRecipient"); const subject=value("messageSubject"); const text=value("messageText");
  if(!recipient) return toast("Вкажіть отримувача");
  if(!text) return toast("Вкажіть текст листування");
  try{
    const now=localTimestamp(), msgId=createId("MSG");
    await Excel.run(async context=>{
      const m=await getOrCreateTableMeta(context,TABLES.messages,"Листування",["ID листа","ID заявки","Відправник","Отримувач","Тема листа","Текст листа","Дата відправлення","ID листування","Статус"]);
      const n=await getTableMeta(context,TABLES.notifications);
      m.table.rows.add(null,[mapRow(m.headers,{"ID листа":msgId,"ID заявки":r.id,"Відправник":STATE.actor,"Отримувач":recipient,"Тема листа":subject||`Заявка ${r.id}`,"Текст листа":text,"Дата відправлення":now,"ID листування":r.id,"Статус":"Очікує відправки"})]);
      const email=clean(STATE.users.find(u=>sameName(u["ПІБ"],recipient))?.["Електронна пошта"]||recipient);
      n.table.rows.add(null,[mapRow(n.headers,{"ID сповіщення":createId("NTF"),"ID заявки":r.id,"ID події":msgId,"Отримувач":recipient,"Email":email,"Канал":"Email","Тема":subject||`Заявка ${r.id}`,"Текст":text,"Дата створення":now,"Статус":email?"Очікує":"Немає email","Спроби":0})]);
      await context.sync();
    });
    setInput("messageRecipient","");setInput("messageSubject","");setInput("messageText",""); closeDrawerForm("messageForm");
    await loadBusinessData();STATE.selectedRequest=findRequest(r.id);renderRequestModal(STATE.selectedRequest);toast("Листування додано в чергу");
  }catch(e){toast("Помилка листування: "+errorText(e));}
}

async function sendChatComment() {
  const r=STATE.selectedRequest;
  if(!r) return;

  const reminderOpen=byId("reminderForm")?.classList.contains("open");
  const messageOpen=byId("messageForm")?.classList.contains("open");
  setDrawerComposerMode("note",true);

  const comment=value("chatInput");
  if(!comment){
    // When switching from reminder/correspondence to note, do not show an error.
    if(reminderOpen || messageOpen) return;
    toast("Введіть текст примітки");
    return;
  }

  try{
    setBusy("chatSendButton",true,"...");
    await directComment(r,comment);
    setInput("chatInput","");
    await loadBusinessData();
    STATE.selectedRequest=findRequest(r.id);
    STATE.historyFilter="notes";
    renderRequestModal(STATE.selectedRequest);
    setDrawerComposerMode("note",false);
    toast("Примітку додано");
  } catch(e){
    showMessage("chatMessage","error",errorText(e));
    toast("Помилка примітки: "+errorText(e));
  } finally{
    setBusy("chatSendButton",false,"+ Примітка");
  }
}

function renderEmployees(){const body=byId("employeesBody");if(!body)return;let users=STATE.users;if(STATE.mode==="employee")users=users.filter(x=>sameName(x["ПІБ"],STATE.actor));const cutoff=Date.now()-30*86400000;body.innerHTML=users.length?users.map(u=>{const n=clean(u["ПІБ"]);const req=STATE.requestsAll.filter(r=>sameName(r.executor,n));const active=req.filter(r=>!isFinalStatus(r.status));const work=active.filter(r=>!["Нова","Призупинена"].includes(r.status));const overdue=active.filter(isOverdue);const closed=req.filter(r=>isFinalStatus(r.status)&&dateMs(r.completedAt)>=cutoff);return `<tr><td class="bold">${esc(n)}</td><td>${esc(u["Роль"]||"—")}</td><td>${active.length}</td><td>${work.length}</td><td class="${overdue.length?"overdue":""}">${overdue.length}</td><td>${closed.length}</td></tr>`;}).join(""):`<tr><td colspan="6"><div class="empty">Немає даних</div></td></tr>`;}
function renderAnalytics(){const req=STATE.requests;setText("analyticsTotal",req.length);setText("analyticsBudget",fmtMoney(sumBudget(req)));setText("analyticsActual",fmtMoney(sumActual(req)));setText("analyticsUnassigned",req.filter(r=>!r.executor).length);setText("analyticsOverdue",req.filter(isOverdue).length);setText("analyticsErp",req.filter(r=>norm(r.source)==="erp"||r.erpId).length);renderBars("statusAnalytics",groupCount(req,r=>r.status||"Без статусу"),STATE.statuses.map(x=>clean(x["Назва статусу"])));renderBars("priorityAnalytics",groupCount(req,r=>r.priority||"Без пріоритету"),STATE.priorities.map(x=>clean(x["Назва"])));}
function renderBars(id,map,order){const el=byId(id);if(!el)return;const entries=[];for(const key of order)if(map[key])entries.push([key,map[key]]);for(const [k,v] of Object.entries(map))if(!entries.some(x=>x[0]===k))entries.push([k,v]);const max=Math.max(1,...entries.map(x=>x[1]));el.innerHTML=entries.length?entries.map(([k,v])=>`<div class="bar-row"><div class="ellipsis" title="${esc(k)}">${esc(k)}</div><div class="bar-track"><div class="bar-fill" style="width:${Math.round(v/max*100)}%"></div></div><div class="right bold">${v}</div></div>`).join(""):`<div class="empty">Немає даних</div>`;}

function renderDictionaries(){renderSimpleDictionary("citiesBody",STATE.cities,["ID міста","Назва","Активний"]);renderSimpleDictionary("categoriesBody",STATE.categories,["ID категорії","Назва","Активний"]);renderSimpleDictionary("customersBody",STATE.customers,["ID замовника","Назва","Електронна пошта"]);renderSimpleDictionary("locationsBody",STATE.locations,["ID локації","Місто","Назва","Адреса"]);}
function renderSimpleDictionary(id,rows,fields){const body=byId(id);if(!body)return;body.innerHTML=rows.length?rows.map(r=>`<tr>${fields.map(f=>`<td><div class="ellipsis">${esc(r[f]||"—")}</div></td>`).join("")}</tr>`).join(""):`<tr><td colspan="${fields.length}"><div class="empty">Немає даних</div></td></tr>`;}

let DICT_TYPE="";
function openDictionaryModal(type, fromCreate=false){
  if(STATE.mode==="unknown") return toast("Не визначено кабінет користувача");
  if(!fromCreate && STATE.mode!=="manager") return toast("Додавати довідники з розділу «Довідники» може лише керівник");

  DICT_TYPE=type;
  if(!fromCreate) STATE.dictReturnTarget="";
  clearMessage("dictMessage");

  const titles={city:"Нове місто",category:"Нова категорія",customer:"Новий замовник",location:"Нова локація"};
  setText("dictTitle",titles[type]||"Додати значення");

  const box=byId("dictFields");
  if(type==="location"){
    box.innerHTML=`<div class="field"><label>Місто</label><select id="dictCity"></select></div>
      <div class="field" style="margin-top:10px"><label>Назва локації</label><input id="dictName"></div>
      <div class="field" style="margin-top:10px"><label>Адреса</label><input id="dictAddress"></div>`;
    fillSelect("dictCity",STATE.cities.map(x=>clean(x["Назва"])),"Оберіть місто");
    const currentCity=value("createCity");
    if(currentCity) setInput("dictCity",currentCity);
  } else if(type==="customer"){
    box.innerHTML=`<div class="field"><label>Назва замовника</label><input id="dictName"></div>
      <div class="field" style="margin-top:10px"><label>Email</label><input id="dictEmail"></div>`;
  } else {
    box.innerHTML=`<div class="field"><label>Назва</label><input id="dictName"></div>`;
  }
  openModal("dictModal");
}

async function directAddDictionary(type,data){
  const config={
    city:{table:TABLES.cities,id:"ID міста",name:"Назва",prefix:"CITY"},
    category:{table:TABLES.categories,id:"ID категорії",name:"Назва",prefix:"CAT"},
    customer:{table:TABLES.customers,id:"ID замовника",name:"Назва",prefix:"CUS"},
    location:{table:TABLES.locations,id:"ID локації",name:"Назва",prefix:"LOC"}
  }[type];
  if(!config) throw new Error("Невідомий тип довідника");

  return Excel.run(async context=>{
    const t=await getTableMeta(context,config.table);
    let rows=[];
    if(t.table.rows.items.length){
      const body=t.table.getDataBodyRange(); body.load("values"); await context.sync(); rows=body.values;
    }
    const nameIdx=t.headers.indexOf(config.name);
    const cityIdx=t.headers.indexOf("Місто");
    const duplicate=rows.some(row=>{
      if(norm(row[nameIdx])!==norm(data.name)) return false;
      if(type==="location") return norm(row[cityIdx])===norm(data.city);
      return true;
    });
    if(duplicate) throw new Error("Таке значення вже існує");

    const now=localTimestamp();
    let obj={};
    if(type==="city") obj={"ID міста":createId("CITY"),"Назва":data.name,"Активний":"Так"};
    if(type==="category") obj={"ID категорії":createId("CAT"),"Назва":data.name,"Активний":"Так"};
    if(type==="customer") obj={"ID замовника":createId("CUS"),"Назва":data.name,"Електронна пошта":data.email||"","Активний":"Так"};
    if(type==="location") obj={"ID локації":createId("LOC"),"Місто":data.city,"Назва":data.name,"Адреса":data.address||"","Додав":STATE.actor,"Дата додавання":now,"Активний":"Так"};

    t.table.rows.add(null,[mapRow(t.headers,obj)]);
    await context.sync();
  });
}

async function saveDictionaryValue(){
  const name=value("dictName");
  if(!name) return showMessage("dictMessage","error","Вкажіть назву.");

  const data={name,email:value("dictEmail"),city:value("dictCity"),address:value("dictAddress")};
  if(DICT_TYPE==="location"&&!data.city) return showMessage("dictMessage","error","Оберіть місто.");

  try{
    setBusy("saveDictButton",true,"Додаємо...");
    await directAddDictionary(DICT_TYPE,data);
    closeModal("dictModal");

    await loadReferenceData();

    if(STATE.dictReturnTarget){
      if(DICT_TYPE==="city"){
        setInput("createCity",name);
        fillLocationSelect("createLocation",name,"Оберіть локацію");
      } else if(DICT_TYPE==="location"){
        const city=data.city||value("createCity");
        if(city) setInput("createCity",city);
        fillLocationSelect("createLocation",city,"Оберіть локацію");
        setInput("createLocation",name);
      } else if(DICT_TYPE==="customer"){
        setInput("createCustomer",name);
      } else if(DICT_TYPE==="category"){
        setInput("createCategory",name);
      }
    }

    STATE.dictReturnTarget="";
    toast("Додано: "+name);
  }catch(e){
    showMessage("dictMessage","error",errorText(e));
  }finally{
    setBusy("saveDictButton",false,"Додати");
  }
}

async function onErpFileSelected(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  STATE.erpFileName = file.name;
  setText("erpFileName", file.name);
  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, {type:"array", cellDates:false, raw:true});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:"", raw:true});
    STATE.erpPreview = mapErpRows(rows);
    renderErpPreview();
    configureMode();
  } catch (e) {
    STATE.erpPreview = [];
    renderErpPreview();
    toast(`Не вдалося прочитати файл: ${errorText(e)}`);
  }
}

function cleanErpValue(v) {
  const s = clean(v);
  return norm(s) === "не призначено" ? "" : s;
}

function erpDate(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number" && Number.isFinite(v) && v > 20000) {
    const ms = Math.round((v - 25569) * 86400000);
    const d = new Date(ms);
    const p = n => String(n).padStart(2,"0");
    return `${d.getUTCFullYear()}-${p(d.getUTCMonth()+1)}-${p(d.getUTCDate())}T${p(d.getUTCHours())}:${p(d.getUTCMinutes())}:${p(d.getUTCSeconds())}`;
  }
  const s = cleanErpValue(v);
  if (!s) return "";
  const ua = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (ua) {
    const p = n => String(n).padStart(2,"0");
    return `${ua[3]}-${p(ua[2])}-${p(ua[1])}T${p(ua[4]||0)}:${p(ua[5]||0)}:${p(ua[6]||0)}`;
  }
  return s;
}

function detectErpHeaderRow(rows) {
  const limit = Math.min(rows?.length || 0, 40);
  for (let i=0; i<limit; i++) {
    const h = (rows[i] || []).map(x => normHeader(x));
    const hasId = h.includes(normHeader("Ідентифікатор запиту")) || h.includes(normHeader("ID ERP"));
    const hasSubject = h.includes(normHeader("Тема")) || h.includes(normHeader("Тема заявки"));
    const hasStatus = h.includes(normHeader("Статус запиту")) || h.includes(normHeader("Статус"));
    if (hasId && hasSubject && hasStatus) return i;
  }
  return -1;
}

function findErpUser(rawName) {
  const raw = cleanErpValue(rawName);
  if (!raw) return {name:"", warning:""};
  const erpSurname = normalizeSurname(raw);
  const candidates = STATE.users.filter(x => normalizeSurname(x["ПІБ"]) === erpSurname);
  if (candidates.length === 1) return {name:clean(candidates[0]["ПІБ"]), warning:""};
  if (candidates.length > 1) return {name:"", warning:`Знайдено кількох працівників з прізвищем «${surname(raw)}»`};
  return {name:"", warning:`Прізвище «${surname(raw)}» не знайдено у Працівниках`};
}

function inferErpCity(site) {
  const raw = cleanErpValue(site);
  if (!raw) return "";
  const text = norm(raw);

  const cities = STATE.cities
    .map(x => clean(x["Назва"]))
    .filter(Boolean)
    .sort((a,b) => b.length-a.length);

  for (const city of cities) {
    if (text.includes(norm(city))) return city;
  }

  const patterns = [
    ["Вінниця", ["вінниця","вінницьк"]],
    ["Київ", ["київ","київськ"]],
    ["Львів", ["львів","львівськ"]],
    ["Харків", ["харків","харківськ"]],
    ["Суми", ["суми","сумськ"]],
    ["Одеса", ["одеса","одеськ"]],
    ["Чернігів", ["чернігів","чернігівськ"]],
    ["Житомир", ["житомир","житомирськ"]],
    ["Миколаїв", ["миколаїв","миколаївськ"]],
    ["Полтава", ["полтава","полтавськ"]],
    ["Кропивницький", ["кропивницьк","кіровоградськ"]],
    ["Тернопіль", ["тернопіл","тернопільськ"]],
    ["Рівне", ["рівне","рівненськ"]],
    ["Луцьк", ["луцьк","волинськ"]],
    ["Черкаси", ["черкаси","черкаськ"]],
    ["Ужгород", ["ужгород","закарпатськ"]],
    ["Мукачево", ["мукачев"]],
    ["Хмельницький", ["хмельницьк"]],
    ["Чернівці", ["чернівц","чернівецьк"]],
    ["Івано-Франківськ", ["івано-франківськ","івано франківськ","франківськ"]],
    ["Шостка", ["шостка","шосткинськ"]],
    ["Запоріжжя", ["запоріж"]],
    ["Дніпро", ["дніпро","дніпропетровськ"]]
  ];
  for (const [city, aliases] of patterns) {
    if (aliases.some(a => text.includes(a))) return city;
  }
  return "";
}

function mapErpPriority(raw) {
  const s = norm(cleanErpValue(raw));
  if (!s) return "Середній";
  if (s.includes("критич" ) || /^1\b/.test(s)) return "Критичний";
  if (s.includes("висок") || /^2\b/.test(s)) return "Високий";
  if (s.includes("стандарт") || s.includes("серед") || /^3\b/.test(s)) return "Середній";
  if (s.includes("низьк") || /^4\b/.test(s)) return "Низький";
  return "Середній";
}

function mapErpStatus(rawStatus, closeCode) {
  const status = norm(cleanErpValue(rawStatus));
  const close = norm(cleanErpValue(closeCode));

  if (close.includes("скас")) return {status:"Закрита", warning:"ERP-заявка скасована — у РБД відображено як «Закрита»"};
  if (!status) return {status:"Нова", warning:"ERP-статус порожній — встановлено «Нова»"};

  if (status.includes("закрит")) return {status:"Закрита", warning:""};
  if (status.includes("прийнята в роботу") || status.includes("взята в роботу") || status.includes("розподілен"))
    return {status:"Прийнята в роботу", warning:""};
  if (status.includes("пошук підряд")) return {status:"Пошук підрядника", warning:""};
  if (status.includes("погодження коштор")) return {status:"Погодження кошторису", warning:""};
  if (status.includes("укладення договор") || status.includes("контрактац")) return {status:"Укладення договору", warning:""};
  if (status.includes("погодження бюдж")) return {status:"Погодження бюджету", warning:""};
  if (status.includes("приймання роб")) return {status:"Приймання робіт", warning:""};
  if (status.includes("виконання роб") || status.includes("в робот")) return {status:"Виконання робіт", warning:""};
  if (status.includes("призуп") || status.includes("заявка не за адресою"))
    return {status:"Призупинена", warning:status.includes("заявка не за адресою")?"ERP-статус «Заявка не за адресою» відображено як «Призупинена»":""};
  if (status.includes("погоджен")) return {status:"Погодження кошторису", warning:"ERP-статус погодження відображено як «Погодження кошторису»"};
  if (status.includes("викон")) return {status:"Виконання робіт", warning:""};

  return {status:"Нова", warning:`Невідомий ERP-статус «${cleanErpValue(rawStatus)}» — встановлено «Нова»`};
}

function erpNeedsUpdate(r, existing) {
  if (!existing) return false;
  const checks = [
    [existing.subject, r.subject], [existing.customer, r.customer], [existing.city, r.city],
    [existing.location, r.location], [existing.category, r.category], [existing.priority, r.priority],
    [existing.status, r.status], [existing.erpStatus, r.erpStatus], [existing.erpExecutor, r.erpExecutor],
    [existing.erpSite, r.erpSite], [existing.erpLevel, r.erpLevel], [existing.erpCloseCode, r.erpCloseCode],
    [existing.erpUpdatedAt, r.erpUpdatedAt]
  ];
  if (r.executor) checks.push([existing.executor, r.executor]);
  return checks.some(([a,b]) => norm(a) !== norm(b));
}

function mapErpRows(rows) {
  if (!rows?.length) return [];
  const headerRow = detectErpHeaderRow(rows);
  if (headerRow < 0) throw new Error("Не знайдено рядок заголовків ERP. Очікується колонка «Ідентифікатор запиту».");

  const sourceHeaders = (rows[headerRow] || []).map(x => clean(x));
  const normalizedHeaders = sourceHeaders.map(x => normHeader(x));
  const getIndex = names => {
    for (const n of names) {
      const i = normalizedHeaders.indexOf(normHeader(n));
      if (i >= 0) return i;
    }
    return -1;
  };
  const val = (row, names) => {
    const i = getIndex(names);
    return i >= 0 ? row[i] : "";
  };

  const localIds = new Set();
  const out = [];

  for (let i=headerRow+1; i<rows.length; i++) {
    const row = rows[i] || [];
    if (!row.some(x => clean(x))) continue;

    const raw = {};
    sourceHeaders.forEach((h,idx) => { if (h) raw[h] = row[idx] ?? ""; });

    const erpId = cleanErpValue(val(row,["Ідентифікатор запиту","ID ERP","ERP ID","Зовнішній ID","ID"]));
    const subject = cleanErpValue(val(row,["Тема","Тема заявки","Назва","Предмет"]));
    const erpExecutor = cleanErpValue(val(row,["Технічний спеціаліст","Виконавець"]));
    const erpSite = cleanErpValue(val(row,["Сайт","Локація","Адреса","Склад","Об'єкт"]));
    const erpStatus = cleanErpValue(val(row,["Статус запиту","Статус"]));
    const erpLevel = cleanErpValue(val(row,["Рівень"]));
    const erpCloseCode = cleanErpValue(val(row,["Код закриття запиту"]));
    const category = cleanErpValue(val(row,["Тип запита","Категорія"]));
    const customer = cleanErpValue(val(row,["Заявник","Замовник","Ініціатор","Автор"]));
    const priority = mapErpPriority(val(row,["Пріоритет"]));
    const erpCreatedAt = erpDate(val(row,["Час створення","Дата створення"]));
    const erpUpdatedAt = erpDate(val(row,["Час останнього оновлення","Дата оновлення"]));
    const erpCompletedAt = erpDate(val(row,["Час завершення","Дата завершення"]));
    const city = inferErpCity(erpSite);
    const location = erpSite;
    const userMap = findErpUser(erpExecutor);
    const statusMap = mapErpStatus(erpStatus, erpCloseCode);
    const warnings = [userMap.warning, statusMap.warning, (!city && location ? "Місто не визначено автоматично" : "")].filter(Boolean);

    let state = "new";
    let result = "Нова";
    let existingId = "";

    if (!erpId) {
      state = "error";
      result = "Немає ID ERP";
    } else if (localIds.has(norm(erpId))) {
      state = "error";
      result = "Дубль ID ERP у файлі";
    } else {
      localIds.add(norm(erpId));
      const existing = STATE.requestsAll.find(x => norm(x.erpId) === norm(erpId));
      if (existing) {
        existingId = existing.id;
        if (erpNeedsUpdate({subject,customer,city,location,category,priority,status:statusMap.status,executor:userMap.name,erpStatus,erpExecutor,erpSite,erpLevel,erpCloseCode,erpUpdatedAt}, existing)) {
          state = "update";
          result = "Оновлення";
        } else {
          state = "same";
          result = "Без змін";
        }
      }
    }

    if (warnings.length && state !== "error") result += ` • ${warnings.join("; ")}`;

    out.push({
      erpId, subject:subject || `ERP ${erpId}`, description:"", customer, city, location, category, priority,
      status:statusMap.status, executor:userMap.name, erpStatus, erpExecutor, erpSite, erpLevel, erpCloseCode,
      erpCreatedAt, erpUpdatedAt, erpCompletedAt, raw, warnings, state, result, existingId
    });
  }
  return out;
}

function renderErpPreview() {
  const body = byId("erpPreviewBody");
  if (!body) return;
  const total = STATE.erpPreview.length;
  const newCount = STATE.erpPreview.filter(x=>x.state==="new").length;
  const updateCount = STATE.erpPreview.filter(x=>x.state==="update").length;
  const sameCount = STATE.erpPreview.filter(x=>x.state==="same").length;
  const err = STATE.erpPreview.filter(x=>x.state==="error").length;
  setText("erpRowsCount", total);
  setText("erpNewCount", newCount);
  setText("erpDuplicateCount", updateCount);
  setText("erpSameCount", sameCount);
  setText("erpErrorCount", err);
  setText("erpPreviewSub", STATE.erpFileName ? `${STATE.erpFileName} • ${total} рядків • заголовки знайдено автоматично` : "Оберіть файл");

  if (!total) {
    body.innerHTML = `<tr><td colspan="10"><div class="empty">Файл ще не вибрано</div></td></tr>`;
  } else {
    body.innerHTML = STATE.erpPreview.map(r => {
      const cls = r.state === "error" ? "overdue" : r.state === "same" ? "muted" : "bold";
      return `<tr>
        <td>${esc(r.erpId||"—")}</td>
        <td>${esc(fmtDateTime(r.erpCreatedAt))}</td>
        <td><div class="ellipsis" title="${esc(r.erpExecutor)}">${esc(r.executor || r.erpExecutor || "—")}</div></td>
        <td><div class="ellipsis" title="${esc(r.subject)}">${esc(r.subject||"—")}</div></td>
        <td><div class="ellipsis" title="${esc(r.erpStatus)}">${esc(r.erpStatus||"—")}</div></td>
        <td>${esc(r.status||"—")}</td>
        <td><div class="ellipsis" title="${esc(r.location)}">${esc(r.location||"—")}</div></td>
        <td>${esc(r.category||"—")}</td>
        <td>${esc(r.priority||"—")}</td>
        <td class="${cls}"><div class="ellipsis" title="${esc(r.result)}">${esc(r.result)}</div></td>
      </tr>`;
    }).join("");
  }

  const processable = newCount + updateCount;
  if (byId("erpImportButton")) byId("erpImportButton").disabled = STATE.mode!=="manager" || processable===0;
}

async function importErpPreview() {
  if (STATE.mode !== "manager") return toast("ERP-імпорт доступний лише керівнику");
  const rows = STATE.erpPreview;
  if (!rows.length) return;

  const importId = createId("IMP");
  const timestamp = localTimestamp();
  const newRows = rows.filter(x=>x.state==="new");
  const updateRows = rows.filter(x=>x.state==="update");
  const queueRows = rows.filter(x=>x.state==="new" || x.state==="update");

  try {
    setBusy("erpImportButton", true, "Імпорт...");
    await Excel.run(async context => {
      const stage = context.workbook.tables.getItem(TABLES.erpStage);
      const commands = context.workbook.tables.getItem(TABLES.commands);
      const imports = context.workbook.tables.getItem(TABLES.erpImports);
      const hs = stage.getHeaderRowRange(), hc = commands.getHeaderRowRange(), hi = imports.getHeaderRowRange();
      hs.load("values"); hc.load("values"); hi.load("values");
      await context.sync();
      const sh = hs.values[0], ch = hc.values[0], ih = hi.values[0];

      const stageRows = rows.map(r => {
        const o = {
          "ID імпорту":importId,
          "Зовнішній ID":r.erpId,
          "Джерело":STATE.erpFileName || "ManageEngine",
          "Тема заявки":r.subject,
          "Опис заявки":r.description,
          "Результат перевірки":r.state==="new" || r.state==="update" ? "Передано в чергу" : r.result,
          "ID заявки РБД":r.existingId || "",
          "Виконавець РБД":r.executor,
          "Місто РБД":r.city,
          "Статус РБД":r.status,
          "Пріоритет РБД":r.priority
        };
        Object.assign(o, r.raw || {});
        return sh.map(h => o[String(h)] ?? "");
      });
      if (stageRows.length) stage.rows.add(null, stageRows);

      const cmdRows = queueRows.map(r => {
        const o = {
          "ID команди":createId("CMD"),
          "Дія":r.state==="update" ? "ERP_SYNC" : "CREATE",
          "ID заявки":r.existingId || "",
          "ID ERP":r.erpId,
          "ID імпорту":importId,
          "Тема заявки":r.subject || `ERP ${r.erpId}`,
          "Опис заявки":r.description,
          "Виконавець":r.executor,
          "Статус":r.status,
          "Місто":r.city,
          "Локація":r.location,
          "Категорія":r.category,
          "Замовник":r.customer,
          "Пріоритет":r.priority || "Середній",
          "Коментар":`Імпорт ManageEngine з ${STATE.erpFileName}${r.warnings?.length ? `. Попередження: ${r.warnings.join("; ")}` : ""}`,
          "Очікувана версія":"",
          "Користувач":STATE.actor,
          "Джерело":"ERP",
          "Дата команди":timestamp,
          "Стан обробки":"Нова",
          "Результат":"",
          "Статус ERP":r.erpStatus,
          "Дата створення ERP":r.erpCreatedAt,
          "Дата оновлення ERP":r.erpUpdatedAt,
          "Дата завершення ERP":r.erpCompletedAt,
          "Виконавець ERP":r.erpExecutor,
          "Сайт ERP":r.erpSite,
          "Рівень ERP":r.erpLevel,
          "Код закриття ERP":r.erpCloseCode
        };
        return ch.map(h => o[String(h)] ?? "");
      });
      if (cmdRows.length) commands.rows.add(null, cmdRows);

      const log = {
        "ID імпорту":importId,
        "Дата імпорту":timestamp,
        "Імпортував":STATE.actor,
        "Назва файлу":STATE.erpFileName,
        "Нових":newRows.length,
        "Оновлених":updateRows.length,
        "Помилок":rows.filter(x=>x.state==="error").length
      };
      imports.rows.add(null, [ih.map(h => log[String(h)] ?? "")]);
      await context.sync();
    });

    toast(`ERP: нових ${newRows.length}, оновлень ${updateRows.length}. Передано в центральну чергу.`);
    clearErpPreview();
  } catch (e) {
    console.error(e);
    toast(`Помилка ERP-імпорту: ${errorText(e)}`);
  } finally {
    setBusy("erpImportButton", false, "Імпортувати в РБД");
    configureMode();
  }
}

function clearErpPreview() {
  STATE.erpPreview = [];
  STATE.erpFileName = "";
  if (byId("erpFile")) byId("erpFile").value = "";
  setText("erpFileName", "Файл не вибрано");
  renderErpPreview();
}

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
