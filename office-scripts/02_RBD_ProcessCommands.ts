type V = string | number | boolean;

function main(workbook: ExcelScript.Workbook): string {
  // ==========================================================
  // РБД 3.0 — ЦЕНТРАЛЬНИЙ ПРОЦЕСОР КОМАНД
  // Обробляє tblRBDCommands2026.
  // Усі події пише в tblRBDTimeline2026.
  // Повідомлення ставить в tblRBDNotifications2026.
  // ==========================================================

  const REQUESTS = [
    "ID заявки", "ID ERP", "Тема заявки", "Опис заявки",
    "Замовник", "Місто", "Локація", "Категорія", "Пріоритет",
    "Статус", "Виконавець", "Створив", "Дата створення",
    "Дата оновлення", "Версія", "Джерело",
    "Дата зміни статусу", "Дата завершення"
  ];
  const USERS = ["ID працівника", "ПІБ", "Електронна пошта", "Роль", "Активний"];
  const STATUS = ["ID статусу", "Назва статусу", "Колір фону", "Колір тексту", "Іконка", "Порядок", "Активний", "Завершальний", "SLA, год"];
  const TRANSITIONS = ["Із статусу (ID)", "До статусу (ID)", "Активний"];
  const COMMANDS = [
    "ID команди", "Дія", "ID заявки", "ID ERP", "ID імпорту",
    "Тема заявки", "Опис заявки", "Виконавець", "Статус",
    "Місто", "Локація", "Категорія", "Замовник", "Пріоритет",
    "Коментар", "Очікувана версія", "Користувач", "Джерело",
    "Дата команди", "Стан обробки", "Результат"
  ];
  const TIMELINE = [
    "ID події", "ID заявки", "Дата події", "Автор", "Тип",
    "Дія", "Текст", "Було", "Стало", "Канал", "Отримувач",
    "Статус доставки", "Джерело", "ID команди"
  ];
  const NOTIFICATIONS = [
    "ID сповіщення", "ID заявки", "ID події", "Отримувач", "Email",
    "Канал", "Тема", "Текст", "Дата створення", "Дата відправки",
    "Статус", "Спроби", "Помилка"
  ];
  const ASSIGNMENTS = ["ID призначення", "ID заявки", "Попередній виконавець", "Новий виконавець", "Призначив", "Дата призначення", "Причина"];

  function table(name: string): ExcelScript.Table {
    for (const ws of workbook.getWorksheets()) {
      for (const t of ws.getTables()) {
        if (t.getName() === name) return t;
      }
    }
    throw new Error(`Не знайдено таблицю ${name}`);
  }

  function optionalTable(name: string): ExcelScript.Table | undefined {
    for (const ws of workbook.getWorksheets()) {
      for (const t of ws.getTables()) {
        if (t.getName() === name) return t;
      }
    }
    return undefined;
  }

  const requests = table("tblRBDRequests2026");
  const users = table("tblRBDUsers2026");
  const statuses = table("tblRBDStatuses2026");
  const transitions = table("tblRBDTransitions2026");
  const commands = table("tblRBDCommands2026");
  const timeline = table("tblRBDTimeline2026");
  const notifications = table("tblRBDNotifications2026");
  const assignments = table("tblRBDAssignments2026");
  const cities = table("tblRBDCities2026");
  const locations = table("tblRBDLocations2026");
  const categories = table("tblRBDCategories2026");
  const customers = table("tblRBDCustomers2026");
  const priorities = table("tblRBDPriorities2026");
  const erpStage = optionalTable("tblRBDERPStage2026");

  function headers(t: ExcelScript.Table): string[] {
    return t.getHeaderRowRange().getTexts()[0];
  }

  function rows(t: ExcelScript.Table): V[][] {
    if (t.getRowCount() === 0) return [];
    return t.getRangeBetweenHeaderAndTotal().getValues() as V[][];
  }

  function get(row: V[], fields: string[], name: string): string {
    const i = fields.indexOf(name);
    return i < 0 ? "" : String(row[i] ?? "").trim();
  }

  function put(t: ExcelScript.Table, rowIndex: number, column: string, value: V): void {
    const hh = headers(t);
    const i = hh.indexOf(column);
    if (i < 0) throw new Error(`Не знайдено поле ${column}`);
    t.getRangeBetweenHeaderAndTotal().getCell(rowIndex, i).setValue(value);
  }

  function addObject(t: ExcelScript.Table, obj: { [key: string]: V }): void {
    const hh = headers(t);
    t.addRow(-1, hh.map(h => obj[h] ?? ""));
  }

  function yes(v: string): boolean {
    return v.trim().toLowerCase() === "так";
  }

  function norm(v: string): string {
    return v.trim().replace(/\u00A0/g, " ").replace(/\s+/g, " ").toLocaleLowerCase();
  }

  function now(): string {
    return new Date().toISOString();
  }

  function uid(prefix: string): string {
    return prefix + "-" + new Date().getTime().toString() + "-" + Math.floor(Math.random() * 1000000).toString();
  }

  const userRows = rows(users);
  const statusRows = rows(statuses);
  const transitionRows = rows(transitions);

  function userByName(name: string): V[] | undefined {
    return userRows.find(r => norm(get(r, USERS, "ПІБ")) === norm(name) && yes(get(r, USERS, "Активний")));
  }

  function isManager(name: string): boolean {
    const u = userByName(name);
    return !!u && get(u, USERS, "Роль").includes("Керівник");
  }

  function emailOf(name: string): string {
    const u = userByName(name);
    return u ? get(u, USERS, "Електронна пошта") : "";
  }

  function statusRowByName(name: string): V[] | undefined {
    return statusRows.find(r => get(r, STATUS, "Назва статусу") === name && yes(get(r, STATUS, "Активний")));
  }

  function finalStatus(name: string): boolean {
    const r = statusRowByName(name);
    return r ? yes(get(r, STATUS, "Завершальний")) : false;
  }

  function transitionAllowed(before: string, after: string): boolean {
    const a = statusRowByName(before);
    const b = statusRowByName(after);
    if (!a || !b) return false;
    const aId = get(a, STATUS, "ID статусу");
    const bId = get(b, STATUS, "ID статусу");
    return transitionRows.some(r =>
      get(r, TRANSITIONS, "Із статусу (ID)") === aId &&
      get(r, TRANSITIONS, "До статусу (ID)") === bId &&
      yes(get(r, TRANSITIONS, "Активний"))
    );
  }

  function dictionaryHas(t: ExcelScript.Table, name: string): boolean {
    if (!name) return true;
    const hh = headers(t);
    return rows(t).some(r => norm(get(r, hh, "Назва")) === norm(name) && yes(get(r, hh, "Активний")));
  }

  function locationExists(city: string, location: string): boolean {
    if (!location) return true;
    const hh = headers(locations);
    return rows(locations).some(r =>
      norm(get(r, hh, "Назва")) === norm(location) &&
      (!city || norm(get(r, hh, "Місто")) === norm(city)) &&
      yes(get(r, hh, "Активний"))
    );
  }

  function nextRequestId(): string {
    let max = 0;
    for (const r of rows(requests)) {
      const m = /^RBD-(\d+)$/.exec(get(r, REQUESTS, "ID заявки"));
      if (m) max = Math.max(max, Number(m[1]));
    }
    return "RBD-" + String(max + 1).padStart(6, "0");
  }

  function findRequest(id: string): number {
    return rows(requests).findIndex(r => get(r, REQUESTS, "ID заявки") === id);
  }

  function requestByIndex(index: number): V[] {
    return rows(requests)[index];
  }

  function timelineHasCommand(commandId: string): boolean {
    const hh = headers(timeline);
    return rows(timeline).some(r => get(r, hh, "ID команди") === commandId);
  }

  function addTimeline(
    requestId: string,
    author: string,
    type: string,
    action: string,
    text: string,
    before: string,
    after: string,
    source: string,
    commandId: string,
    channel: string = "",
    recipient: string = "",
    delivery: string = ""
  ): string {
    const eventId = uid("EV");
    addObject(timeline, {
      "ID події": eventId,
      "ID заявки": requestId,
      "Дата події": now(),
      "Автор": author,
      "Тип": type,
      "Дія": action,
      "Текст": text,
      "Було": before,
      "Стало": after,
      "Канал": channel,
      "Отримувач": recipient,
      "Статус доставки": delivery,
      "Джерело": source,
      "ID команди": commandId
    });
    return eventId;
  }

  function queueNotification(
    requestId: string,
    eventId: string,
    recipient: string,
    subject: string,
    text: string
  ): void {
    if (!recipient) return;
    const email = emailOf(recipient);
    const status = email ? "Очікує" : "Немає email";
    addObject(notifications, {
      "ID сповіщення": uid("NTF"),
      "ID заявки": requestId,
      "ID події": eventId,
      "Отримувач": recipient,
      "Email": email,
      "Канал": "Email",
      "Тема": subject,
      "Текст": text,
      "Дата створення": now(),
      "Дата відправки": "",
      "Статус": status,
      "Спроби": 0,
      "Помилка": ""
    });
  }

  function recordAssignment(
    requestId: string,
    oldOwner: string,
    newOwner: string,
    actor: string,
    reason: string
  ): void {
    addObject(assignments, {
      "ID призначення": uid("AS"),
      "ID заявки": requestId,
      "Попередній виконавець": oldOwner,
      "Новий виконавець": newOwner,
      "Призначив": actor,
      "Дата призначення": now(),
      "Причина": reason
    });
  }

  function updateErpStage(importId: string, erpId: string, result: string, rbdId: string): void {
    if (!erpStage || !importId || !erpId || erpStage.getRowCount() === 0) return;
    const hh = headers(erpStage);
    const rr = rows(erpStage);
    for (let i = 0; i < rr.length; i++) {
      if (get(rr[i], hh, "ID імпорту") === importId && get(rr[i], hh, "Зовнішній ID") === erpId) {
        put(erpStage, i, "Результат перевірки", result);
        put(erpStage, i, "ID заявки РБД", rbdId);
        return;
      }
    }
  }

  function nextDictionaryId(t: ExcelScript.Table, idField: string, prefix: string): string {
    const hh = headers(t);
    let max = 0;
    for (const r of rows(t)) {
      const id = get(r, hh, idField);
      const m = new RegExp("^" + prefix + "-(\\d+)$").exec(id);
      if (m) max = Math.max(max, Number(m[1]));
    }
    return prefix + "-" + String(max + 1).padStart(4, "0");
  }

  function addDictionary(
    action: string,
    actor: string,
    city: string,
    location: string,
    category: string,
    customer: string,
    description: string,
    comment: string,
    commandId: string
  ): string {
    if (!isManager(actor)) throw new Error("Додавати довідники може лише керівник");

    let target: ExcelScript.Table;
    let name = "";
    let idField = "";
    let prefix = "";

    if (action === "ADD_CITY") {
      target = cities; name = city; idField = "ID міста"; prefix = "CITY";
    } else if (action === "ADD_CATEGORY") {
      target = categories; name = category; idField = "ID категорії"; prefix = "CAT";
    } else if (action === "ADD_CUSTOMER") {
      target = customers; name = customer; idField = "ID замовника"; prefix = "CUS";
    } else {
      target = locations; name = location; idField = "ID локації"; prefix = "LOC";
      if (!city || !dictionaryHas(cities, city)) throw new Error("Для локації виберіть існуюче місто");
    }

    if (!name) throw new Error("Не вказано назву нового значення");
    const hh = headers(target);
    const duplicate = rows(target).some(r => {
      const same = norm(get(r, hh, "Назва")) === norm(name);
      if (action === "ADD_LOCATION") return same && norm(get(r, hh, "Місто")) === norm(city);
      return same;
    });
    if (duplicate) throw new Error("Таке значення вже існує");

    const id = nextDictionaryId(target, idField, prefix);
    const obj: { [key: string]: V } = {"Назва": name, "Активний": "Так"};
    obj[idField] = id;
    if (action === "ADD_CUSTOMER") obj["Електронна пошта"] = comment;
    if (action === "ADD_LOCATION") {
      obj["Місто"] = city;
      obj["Адреса"] = description;
      obj["Додав"] = actor;
      obj["Дата додавання"] = now();
    }
    addObject(target, obj);
    addTimeline("", actor, "Довідник", action, `Додано: ${name}`, "", name, "Надбудова", commandId);
    return `Додано: ${name}`;
  }

  const commandRows = rows(commands);
  if (commandRows.length === 0) return "Нових команд немає";

  let done = 0;
  let failed = 0;
  const commandHeaders = headers(commands);

  for (let i = 0; i < commandRows.length; i++) {
    const row = commandRows[i];
    const state = get(row, commandHeaders, "Стан обробки");
    if (state !== "Нова" && state !== "NEW") continue;

    const commandId = get(row, commandHeaders, "ID команди");
    const action = get(row, commandHeaders, "Дія").toUpperCase();
    const requestId = get(row, commandHeaders, "ID заявки");
    const erpId = get(row, commandHeaders, "ID ERP");
    const importId = get(row, commandHeaders, "ID імпорту");
    const subject = get(row, commandHeaders, "Тема заявки");
    const description = get(row, commandHeaders, "Опис заявки");
    const executor = get(row, commandHeaders, "Виконавець");
    const status = get(row, commandHeaders, "Статус");
    const city = get(row, commandHeaders, "Місто");
    const location = get(row, commandHeaders, "Локація");
    const category = get(row, commandHeaders, "Категорія");
    const customer = get(row, commandHeaders, "Замовник");
    const priority = get(row, commandHeaders, "Пріоритет");
    const comment = get(row, commandHeaders, "Коментар");
    const expected = get(row, commandHeaders, "Очікувана версія");
    const actor = get(row, commandHeaders, "Користувач");
    const source = get(row, commandHeaders, "Джерело") || "Надбудова";

    try {
      if (!commandId) throw new Error("Відсутній ID команди");
      if (!userByName(actor)) throw new Error("Користувач відсутній або неактивний");

      put(commands, i, "Стан обробки", "В обробці");

      if (timelineHasCommand(commandId)) {
        put(commands, i, "Стан обробки", "Виконано");
        put(commands, i, "Результат", "Команда вже була виконана");
        done++;
        continue;
      }

      if (["ADD_CITY", "ADD_LOCATION", "ADD_CATEGORY", "ADD_CUSTOMER"].includes(action)) {
        const result = addDictionary(action, actor, city, location, category, customer, description, comment, commandId);
        put(commands, i, "Стан обробки", "Виконано");
        put(commands, i, "Результат", result);
        done++;
        continue;
      }

      if (action === "CREATE") {
        if (!subject) throw new Error("Вкажіть тему заявки");

        // Ідемпотентність CREATE: якщо попередня спроба вже встигла
        // створити заявку і записати її ID у команду, повторно не створюємо.
        if (requestId && findRequest(requestId) >= 0) {
          put(commands, i, "Стан обробки", "Виконано");
          put(commands, i, "Результат", `Вже створено ${requestId}`);
          if (source === "ERP") updateErpStage(importId, erpId, "Імпортовано", requestId);
          done++;
          continue;
        }

        const manager = isManager(actor);
        const owner = manager ? executor : actor;
        if (owner && !userByName(owner)) throw new Error("Невідомий або неактивний виконавець");
        if (city && !dictionaryHas(cities, city)) throw new Error("Міста немає в довіднику");
        if (category && !dictionaryHas(categories, category)) throw new Error("Категорії немає в довіднику");
        if (customer && !dictionaryHas(customers, customer)) throw new Error("Замовника немає в довіднику");
        if (priority && !dictionaryHas(priorities, priority)) throw new Error("Пріоритету немає в довіднику");
        if (!locationExists(city, location)) throw new Error("Локація не відповідає місту");

        if (erpId) {
          const duplicate = rows(requests).some(r => norm(get(r, REQUESTS, "ID ERP")) === norm(erpId));
          if (duplicate) throw new Error(`ID ERP ${erpId} вже існує`);
        }

        const newId = requestId || nextRequestId();
        // Фіксуємо майбутній ID у команді ДО вставки в базу.
        // Це дозволяє безпечно повторити обробку після переривання.
        put(commands, i, "ID заявки", newId);
        const created = now();
        addObject(requests, {
          "ID заявки": newId,
          "ID ERP": erpId,
          "Тема заявки": subject,
          "Опис заявки": description,
          "Замовник": customer,
          "Місто": city,
          "Локація": location,
          "Категорія": category,
          "Пріоритет": priority || "Середній",
          "Статус": "Нова",
          "Виконавець": owner,
          "Створив": actor,
          "Дата створення": created,
          "Дата оновлення": created,
          "Версія": 1,
          "Джерело": source === "ERP" ? "ERP" : "РБД",
          "Дата зміни статусу": created,
          "Дата завершення": ""
        });

        const eventId = addTimeline(
          newId, actor, source === "ERP" ? "ERP" : "Створення", "Створено заявку",
          comment || subject, "", "Нова", source, commandId
        );

        if (owner) {
          recordAssignment(newId, "", owner, actor, "Створення заявки");
          const assignEvent = addTimeline(newId, actor, "Призначення", "Призначено виконавця", "", "", owner, source, commandId + "-ASSIGN");
          queueNotification(newId, assignEvent, owner, `Нова заявка ${newId}`, `${subject}\nВас призначено виконавцем.`);
        }

        if (source === "ERP") updateErpStage(importId, erpId, "Імпортовано", newId);
        put(commands, i, "ID заявки", newId);
        put(commands, i, "Стан обробки", "Виконано");
        put(commands, i, "Результат", `Створено ${newId}`);
        done++;
        continue;
      }

      const ri = findRequest(requestId);
      if (ri < 0) throw new Error(`Заявку не знайдено: ${requestId}`);
      const current = requestByIndex(ri);
      const owner = get(current, REQUESTS, "Виконавець");
      const oldStatus = get(current, REQUESTS, "Статус");
      const version = Number(get(current, REQUESTS, "Версія") || "0");
      const manager = isManager(actor);

      if (expected && Number(expected) !== version) throw new Error(`Конфлікт версій: очікувалась ${expected}, поточна ${version}`);

      if (action === "ASSIGN") {
        if (!manager) throw new Error("Призначати іншого виконавця може лише керівник");
        if (!executor || !userByName(executor)) throw new Error("Оберіть активного виконавця");
        if (finalStatus(oldStatus)) throw new Error("Заявка вже завершена");
        if (norm(owner) === norm(executor)) throw new Error("Цей виконавець уже призначений");

        put(requests, ri, "Виконавець", executor);
        recordAssignment(requestId, owner, executor, actor, comment);
        const eventId = addTimeline(requestId, actor, "Призначення", "Змінено виконавця", comment, owner, executor, source, commandId);
        queueNotification(requestId, eventId, executor, `Призначено заявку ${requestId}`, `Вас призначено виконавцем заявки «${get(current, REQUESTS, "Тема заявки")}».`);
      }
      else if (action === "TAKE") {
        if (owner && norm(owner) !== norm(actor) && !manager) throw new Error("Заявка призначена іншому виконавцю");
        if (!transitionAllowed(oldStatus, "Взята в роботу")) throw new Error(`Недозволений перехід: ${oldStatus} → Взята в роботу`);
        if (!owner) {
          put(requests, ri, "Виконавець", actor);
          recordAssignment(requestId, "", actor, actor, "Самостійно взято в роботу");
          addTimeline(requestId, actor, "Призначення", "Взято заявку на себе", "", "", actor, source, commandId + "-ASSIGN");
        }
        put(requests, ri, "Статус", "Взята в роботу");
        put(requests, ri, "Дата зміни статусу", now());
        addTimeline(requestId, actor, "Статус", "Змінено статус", comment, oldStatus, "Взята в роботу", source, commandId);
      }
      else if (action === "STATUS_CHANGE" || action === "CLOSE" || action === "CANCEL") {
        if (!manager && norm(owner) !== norm(actor)) throw new Error("Можна змінювати лише власні заявки");
        const next = action === "CLOSE" ? "Закрита" : action === "CANCEL" ? "Скасована" : status;
        if (!statusRowByName(next)) throw new Error("Статус відсутній або неактивний");
        if (!transitionAllowed(oldStatus, next)) throw new Error(`Недозволений перехід: ${oldStatus} → ${next}`);
        if (next === "Призупинена" && !comment) throw new Error("Для призупинення потрібна причина");

        put(requests, ri, "Статус", next);
        put(requests, ri, "Дата зміни статусу", now());
        if (finalStatus(next) || next === "Виконана") put(requests, ri, "Дата завершення", now());
        const eventId = addTimeline(requestId, actor, "Статус", "Змінено статус", comment, oldStatus, next, source, commandId);
        if (owner && norm(owner) !== norm(actor)) queueNotification(requestId, eventId, owner, `Статус заявки ${requestId}`, `${oldStatus} → ${next}${comment ? "\n" + comment : ""}`);
      }
      else if (action === "COMMENT") {
        if (!manager && norm(owner) !== norm(actor)) throw new Error("Не можна коментувати чужу заявку");
        if (!comment) throw new Error("Введіть коментар");
        addTimeline(requestId, actor, "Коментар", "Коментар", comment, "", "", source, commandId);
      }
      else {
        throw new Error(`Невідома дія: ${action}`);
      }

      put(requests, ri, "Дата оновлення", now());
      put(requests, ri, "Версія", version + 1);
      put(commands, i, "Стан обробки", "Виконано");
      put(commands, i, "Результат", "Виконано");
      done++;
    }
    catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      try {
        put(commands, i, "Стан обробки", "Помилка");
        put(commands, i, "Результат", message);
        if (importId && erpId) updateErpStage(importId, erpId, "Помилка: " + message, "");
        if (requestId) addTimeline(requestId, "Система", "Система", "Помилка команди", message, "", "", source || "Система", commandId);
      } catch (_) {
        // Не перекриваємо первинну помилку.
      }
      failed++;
    }
  }

  return `Обробка завершена. Виконано: ${done}. Помилок: ${failed}.`;
}
