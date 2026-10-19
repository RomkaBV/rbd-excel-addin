type V = string | number | boolean;

function main(workbook: ExcelScript.Workbook): string {
  // ==========================================================
  // РБД 3.0 — ПІДГОТОВКА НОВОЇ АРХІТЕКТУРИ
  // ----------------------------------------------------------
  // 1) Зберігає існуючі центральні дані.
  // 2) Створює одну центральну таблицю команд.
  // 3) Створює єдину хронологію заявки та журнал сповіщень.
  // 4) Очищає персональні вкладки працівників від старих
  //    Excel-кабінетів. Вкладка = прізвище працівника.
  // 5) НЕ створює старі tblRBDInputUSR... / FILTER-реєстри.
  // ==========================================================

  const CLEAN_PERSONAL_SHEETS = false;

  const REQUESTS = [
    "ID заявки", "ID ERP", "Тема заявки", "Опис заявки",
    "Замовник", "Місто", "Локація", "Категорія", "Пріоритет",
    "Бюджет заявки, грн", "Фактично використано, грн", "Факт виконання", "Дата внесення факту", "Факт вніс",
    "Статус", "Виконавець", "Створив", "Дата створення",
    "Дата оновлення", "Версія", "Джерело",
    "Дата зміни статусу", "Дата завершення",
    "Статус ERP", "Виконавець ERP", "Сайт ERP", "Рівень ERP",
    "Код закриття ERP", "Дата оновлення ERP"
  ];

  const USERS = [
    "ID працівника", "ПІБ", "Електронна пошта", "Роль", "Активний"
  ];

  const STATUSES = [
    "ID статусу", "Назва статусу", "Колір фону", "Колір тексту",
    "Іконка", "Порядок", "Активний", "Завершальний", "SLA, год"
  ];

  const TRANSITIONS = [
    "Із статусу (ID)", "До статусу (ID)", "Активний"
  ];

  const ASSIGNMENTS = [
    "ID призначення", "ID заявки", "Попередній виконавець",
    "Новий виконавець", "Призначив", "Дата призначення", "Причина"
  ];

  const COMMANDS = [
    "ID команди", "Дія", "ID заявки", "ID ERP", "ID імпорту",
    "Тема заявки", "Опис заявки", "Виконавець", "Статус",
    "Місто", "Локація", "Категорія", "Замовник", "Пріоритет",
    "Бюджет заявки, грн", "Коментар", "Очікувана версія", "Користувач", "Джерело",
    "Дата команди", "Стан обробки", "Результат",
    "Статус ERP", "Дата створення ERP", "Дата оновлення ERP",
    "Дата завершення ERP", "Виконавець ERP", "Сайт ERP",
    "Рівень ERP", "Код закриття ERP"
  ];

  const TIMELINE = [
    "ID події", "ID заявки", "Дата події", "Автор", "Тип",
    "Дія", "Текст", "Було", "Стало",
    "Початок статусу", "Кінець статусу", "Тривалість статусу, год", "Тривалість статусу, роб. днів",
    "Канал", "Отримувач", "Статус доставки", "Джерело", "ID команди"
  ];

  const NOTIFICATIONS = [
    "ID сповіщення", "ID заявки", "ID події", "Отримувач", "Email",
    "Канал", "Тема", "Текст", "Дата створення", "Дата відправки",
    "Статус", "Спроби", "Помилка"
  ];

  const MESSAGES = [
    "ID листа", "ID заявки", "Відправник", "Отримувач",
    "Тема листа", "Текст листа", "Дата відправлення",
    "ID листування", "Статус"
  ];

  const REMINDERS = [
    "ID нагадування", "ID заявки", "Отримувач", "Дата нагадування",
    "Опис", "Статус", "Створив", "Дата створення"
  ];

  const ERP_STAGE = [
    "ID імпорту", "Зовнішній ID", "Джерело", "Тема заявки",
    "Опис заявки", "Результат перевірки", "ID заявки РБД",
    "Виконавець РБД", "Місто РБД", "Статус РБД", "Пріоритет РБД",
    "Ідентифікатор запиту", "Час створення", "Технічний спеціаліст",
    "Рівень", "Сайт", "Тема", "Статус запиту", "Тип запита",
    "Назва посади ініціатора запиту", "Код закриття запиту",
    "Винесення пропозицій та отримання рішення Тендерного Комітету",
    "Відділ", "Група", "Дата відповіді", "Заявник", "Керівник:",
    "Моніторинг ринку Пошук кращих моделей. Погодження ТЗ",
    "Назва угоди про рівень обслуговування", "Отримання товару",
    "Перевірка постачальника. Оформлення договору",
    "Погодження конкретної моделі по критеріям ТЗ з замовником",
    "Пріоритет", "Підбір постачальників по обраній моделі",
    "Час вирішення SLA", "Час відповіді минув",
    "Час останнього оновлення", "Час терміну виконання",
    "Статус очікування", "Час завершення"
  ];

  const ERP_IMPORTS = [
    "ID імпорту", "Дата імпорту", "Імпортував", "Назва файлу",
    "Нових", "Оновлених", "Помилок"
  ];

  const CITIES = ["ID міста", "Назва", "Активний"];
  const LOCATIONS = [
    "ID локації", "Місто", "Назва", "Адреса",
    "Додав", "Дата додавання", "Активний"
  ];
  const CATEGORIES = ["ID категорії", "Назва", "Активний"];
  const CUSTOMERS = ["ID замовника", "Назва", "Електронна пошта", "Активний"];
  const PRIORITIES = ["ID пріоритету", "Назва", "Порядок", "Активний"];

  const INITIAL_USERS: V[][] = [
    ["USR-001", "Іванченко В.М.", "", "Керівник + Виконавець", "Так"],
    ["USR-002", "Войцехівський Г.В.", "", "Виконавець", "Так"],
    ["USR-003", "Ридванський П.С.", "", "Виконавець", "Так"],
    ["USR-004", "Галько А.І.", "", "Виконавець", "Так"],
    ["USR-005", "Трунов Ю.О.", "", "Виконавець", "Так"],
    ["USR-006", "Желясков Д.О.", "", "Виконавець", "Так"],
    ["USR-007", "Слепущенко О.О.", "", "Виконавець", "Так"],
    ["USR-008", "Сергеєв П.А.", "", "Виконавець", "Так"],
    ["USR-009", "Туровський В.О.", "", "Виконавець", "Так"]
  ];

  const INITIAL_STATUSES: V[][] = [
    ["ST-001", "Нова", "#BFE2FF", "#145A96", "plus", 1, "Так", "Ні", 0],
    ["ST-002", "Прийнята в роботу", "#FFE6A8", "#6D4A00", "play", 2, "Так", "Ні", 0],
    ["ST-003", "Пошук підрядника", "#FFC487", "#A93B00", "search", 3, "Так", "Ні", 0],
    ["ST-004", "Погодження кошторису", "#FFE7A8", "#7A5200", "check", 4, "Так", "Ні", 0],
    ["ST-005", "Укладення договору", "#DCCEFF", "#4520A8", "file", 5, "Так", "Ні", 0],
    ["ST-006", "Погодження бюджету", "#F7C4E0", "#6E174F", "money", 6, "Так", "Ні", 0],
    ["ST-007", "Виконання робіт", "#BFEACF", "#0E6A45", "tools", 7, "Так", "Ні", 0],
    ["ST-008", "Приймання робіт", "#C3EFF5", "#116079", "check", 8, "Так", "Ні", 0],
    ["ST-009", "Призупинена", "#DCE5F1", "#23466C", "pause", 9, "Так", "Ні", 0],
    ["ST-010", "Закрита", "#C6EBD3", "#11623D", "lock", 10, "Так", "Так", 0]
  ];

  const INITIAL_TRANSITIONS: V[][] = [
    ["ST-001", "ST-002", "Так"],
    ["ST-002", "ST-003", "Так"],
    ["ST-003", "ST-004", "Так"],
    ["ST-004", "ST-005", "Так"],
    ["ST-005", "ST-006", "Так"],
    ["ST-006", "ST-007", "Так"],
    ["ST-007", "ST-008", "Так"],
    ["ST-008", "ST-010", "Так"],
    ["ST-001", "ST-009", "Так"],
    ["ST-002", "ST-009", "Так"],
    ["ST-003", "ST-009", "Так"],
    ["ST-004", "ST-009", "Так"],
    ["ST-005", "ST-009", "Так"],
    ["ST-006", "ST-009", "Так"],
    ["ST-007", "ST-009", "Так"],
    ["ST-008", "ST-009", "Так"],
    ["ST-009", "ST-003", "Так"]
  ];

  const INITIAL_PRIORITIES: V[][] = [
    ["PR-001", "Низький", 1, "Так"],
    ["PR-002", "Середній", 2, "Так"],
    ["PR-003", "Високий", 3, "Так"],
    ["PR-004", "Критичний", 4, "Так"]
  ];

  function sheet(name: string): ExcelScript.Worksheet {
    let ws = workbook.getWorksheet(name);
    if (!ws) ws = workbook.addWorksheet(name);
    return ws;
  }

  function findTable(name: string): ExcelScript.Table | undefined {
    for (const ws of workbook.getWorksheets()) {
      for (const t of ws.getTables()) {
        if (t.getName() === name) return t;
      }
    }
    return undefined;
  }

  function tableHeaders(t: ExcelScript.Table): string[] {
    return t.getHeaderRowRange().getTexts()[0];
  }

  function ensureTable(
    sheetName: string,
    tableName: string,
    headers: string[]
  ): ExcelScript.Table {
    const existing = findTable(tableName);
    if (existing) {
      // Міграція без втрати даних: існуючі колонки не чіпаємо,
      // а нові для ERP/РБД 3.1 просто додаємо в кінець таблиці.
      let actual = tableHeaders(existing);
      for (const header of headers) {
        if (!actual.includes(header)) {
          existing.addColumn(-1, undefined, header);
          actual = tableHeaders(existing);
        }
      }
      return existing;
    }

    const ws = sheet(sheetName);
    const start = ws.getRangeByIndexes(0, 0, 1, headers.length);
    start.setValues([headers]);
    const t = ws.addTable(start, true);
    t.setName(tableName);
    return t;
  }

  function seed(t: ExcelScript.Table, values: V[][]): void {
    if (t.getRowCount() === 0 && values.length > 0) t.addRows(-1, values);
  }
  function syncRowsById(t: ExcelScript.Table, headers: string[], values: V[][], idField: string): void {
    const idIdx = headers.indexOf(idField);
    if (idIdx < 0) return;

    const existingRows = rows(t);
    const existingMap = new Map<string, number>();
    existingRows.forEach((r, i) => existingMap.set(String(r[idIdx] ?? "").trim(), i));

    for (const source of values) {
      const id = String(source[idIdx] ?? "").trim();
      if (!id) continue;
      const rowIndex = existingMap.get(id);
      if (rowIndex === undefined) {
        t.addRow(-1, source);
      } else {
        t.getRangeBetweenHeaderAndTotal().getRow(rowIndex).setValues([source]);
      }
    }
  }

  function migrateStatusName(value: V): V {
    const text = String(value ?? "").trim();
    const map: { [key: string]: string } = {
      "Розподілена": "Прийнята в роботу",
      "Взята в роботу": "Прийнята в роботу",
      "В роботі": "Виконання робіт",
      "На погодженні": "Погодження кошторису",
      "Виконана": "Приймання робіт",
      "Скасована": "Закрита"
    };
    return map[text] ?? value;
  }

  function migrateStatusColumn(t: ExcelScript.Table, field: string): void {
    const headers = tableHeaders(t);
    const idx = headers.indexOf(field);
    if (idx < 0 || t.getRowCount() === 0) return;
    const range = t.getRangeBetweenHeaderAndTotal();
    const values = range.getValues() as V[][];
    let changed = false;
    for (const row of values) {
      const next = migrateStatusName(row[idx]);
      if (next !== row[idx]) {
        row[idx] = next;
        changed = true;
      }
    }
    if (changed) range.setValues(values);
  }


  function rows(t: ExcelScript.Table): V[][] {
    if (t.getRowCount() === 0) return [];
    return t.getRangeBetweenHeaderAndTotal().getValues() as V[][];
  }

  function get(row: V[], fields: string[], name: string): string {
    const i = fields.indexOf(name);
    return i < 0 ? "" : String(row[i] ?? "").trim();
  }

  function yes(v: string): boolean {
    return v.trim().toLowerCase() === "так";
  }

  function surname(person: string): string {
    return person.trim().split(/\s+/)[0] || "";
  }

  function styleTable(t: ExcelScript.Table): void {
    t.setPredefinedTableStyle("TableStyleMedium2");
    t.setShowBandedRows(true);
    t.setShowFilterButton(true);
    const head = t.getHeaderRowRange();
    head.getFormat().getFill().setColor("#1769AA");
    head.getFormat().getFont().setColor("#FFFFFF");
    head.getFormat().getFont().setBold(true);
    head.getFormat().setRowHeight(34);
  }

  function clearPersonalSheet(ws: ExcelScript.Worksheet): void {
    for (const t of ws.getTables()) t.delete();
    const used = ws.getUsedRange();
    if (used) used.clear(ExcelScript.ClearApplyTo.all);
  }

  // ----------------------------------------------------------
  // Центральні таблиці
  // ----------------------------------------------------------
  const requests = ensureTable("Заявки", "tblRBDRequests2026", REQUESTS);
  const users = ensureTable("Працівники", "tblRBDUsers2026", USERS);
  const statuses = ensureTable("Статуси", "tblRBDStatuses2026", STATUSES);
  const transitions = ensureTable("Переходи статусів", "tblRBDTransitions2026", TRANSITIONS);
  const assignments = ensureTable("Призначення", "tblRBDAssignments2026", ASSIGNMENTS);
  const commands = ensureTable("Команди", "tblRBDCommands2026", COMMANDS);
  const timeline = ensureTable("Хронологія", "tblRBDTimeline2026", TIMELINE);
  const notifications = ensureTable("Сповіщення", "tblRBDNotifications2026", NOTIFICATIONS);
  const messages = ensureTable("Листування", "tblRBDMessages2026", MESSAGES);
  const reminders = ensureTable("Нагадування", "tblRBDReminders2026", REMINDERS);
  const cities = ensureTable("Міста", "tblRBDCities2026", CITIES);
  const locations = ensureTable("Локації", "tblRBDLocations2026", LOCATIONS);
  const categories = ensureTable("Категорії", "tblRBDCategories2026", CATEGORIES);
  const customers = ensureTable("Замовники", "tblRBDCustomers2026", CUSTOMERS);
  const priorities = ensureTable("Пріоритети", "tblRBDPriorities2026", PRIORITIES);
  const erpStage = ensureTable("ERP — вхідні", "tblRBDERPStage2026", ERP_STAGE);
  const erpImports = ensureTable("Журнал імпортів", "tblRBDERPImports2026", ERP_IMPORTS);

  seed(users, INITIAL_USERS);
  syncRowsById(statuses, STATUSES, INITIAL_STATUSES, "ID статусу");
  syncRowsById(transitions, TRANSITIONS, INITIAL_TRANSITIONS, "Із статусу (ID)");
  seed(priorities, INITIAL_PRIORITIES);

  migrateStatusColumn(requests, "Статус");
  migrateStatusColumn(timeline, "Було");
  migrateStatusColumn(timeline, "Стало");

  // ----------------------------------------------------------
  // Персональні вкладки: одна вкладка = один працівник.
  // Вони більше не містять Excel-таблиць кабінету.
  // ----------------------------------------------------------
  const userRows = rows(users);
  const activeNames = userRows
    .filter(r => yes(get(r, USERS, "Активний")))
    .map(r => get(r, USERS, "ПІБ"))
    .filter(n => n !== "");

  const personalTabs: string[] = [];
  for (const person of activeNames) {
    const tabName = surname(person);
    if (!tabName) continue;
    const ws = sheet(tabName);
    if (CLEAN_PERSONAL_SHEETS) clearPersonalSheet(ws);
    personalTabs.push(tabName);
  }

  // Оформлюємо технічні таблиці. Дані не очищаються.
  [
    requests, users, statuses, transitions, assignments,
    commands, timeline, notifications, messages, reminders, cities, locations,
    categories, customers, priorities, erpStage, erpImports
  ].forEach(t => styleTable(t));

  // Корисна службова інформація на окремій вкладці.
  const info = sheet("Система");
  info.getRange("A1:B8").clear(ExcelScript.ClearApplyTo.all);
  info.getRange("A1:B8").setValues([
    ["РБД", "3.3.2"],
    ["Архітектура", "Одна центральна черга команд"],
    ["Таблиця команд", "tblRBDCommands2026"],
    ["Хронологія", "tblRBDTimeline2026"],
    ["Сповіщення", "tblRBDNotifications2026"],
    ["Кабінет", "Визначається за назвою активної вкладки"],
    ["Роль", "Береться з tblRBDUsers2026"],
    ["Персональні вкладки", personalTabs.join(", ")]
  ]);
  info.getRange("A1:A8").getFormat().getFont().setBold(true);
  info.getRange("A:B").getFormat().setColumnWidth(240);

  return `РБД 3.3.2 готово. Персональні вкладки: ${personalTabs.join(", ")}. Старий «Кабінет керівника» та старі службові вкладки не видалялися автоматично — після перевірки їх можна прибрати вручну.`;
}
