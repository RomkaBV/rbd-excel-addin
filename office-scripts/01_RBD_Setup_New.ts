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
    "Коментар", "Очікувана версія", "Користувач", "Джерело",
    "Дата команди", "Стан обробки", "Результат",
    "Статус ERP", "Дата створення ERP", "Дата оновлення ERP",
    "Дата завершення ERP", "Виконавець ERP", "Сайт ERP",
    "Рівень ERP", "Код закриття ERP"
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
    ["ST-001", "Нова", "#32A8FF", "#FFFFFF", "plus", 1, "Так", "Ні", 0],
    ["ST-002", "Розподілена", "#80C7FF", "#17382B", "user", 2, "Так", "Ні", 0],
    ["ST-003", "Взята в роботу", "#FFCA28", "#202020", "play", 3, "Так", "Ні", 0],
    ["ST-004", "В роботі", "#88D8AA", "#17382B", "clock", 4, "Так", "Ні", 0],
    ["ST-005", "Призупинена", "#FFB4BA", "#81202B", "pause", 5, "Так", "Ні", 0],
    ["ST-006", "На погодженні", "#CFB6F7", "#352052", "check", 6, "Так", "Ні", 0],
    ["ST-007", "Виконана", "#90E4B8", "#145B3C", "check", 7, "Так", "Ні", 0],
    ["ST-008", "Закрита", "#D4DEDA", "#34453D", "lock", 8, "Так", "Так", 0],
    ["ST-009", "Скасована", "#FFCCD0", "#892632", "x", 9, "Так", "Так", 0]
  ];

  const INITIAL_TRANSITIONS: V[][] = [
    ["ST-001", "ST-002", "Так"],
    ["ST-001", "ST-003", "Так"],
    ["ST-001", "ST-009", "Так"],
    ["ST-002", "ST-003", "Так"],
    ["ST-003", "ST-004", "Так"],
    ["ST-004", "ST-005", "Так"],
    ["ST-005", "ST-004", "Так"],
    ["ST-004", "ST-006", "Так"],
    ["ST-006", "ST-004", "Так"],
    ["ST-004", "ST-007", "Так"],
    ["ST-004", "ST-009", "Так"],
    ["ST-007", "ST-008", "Так"]
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
  const cities = ensureTable("Міста", "tblRBDCities2026", CITIES);
  const locations = ensureTable("Локації", "tblRBDLocations2026", LOCATIONS);
  const categories = ensureTable("Категорії", "tblRBDCategories2026", CATEGORIES);
  const customers = ensureTable("Замовники", "tblRBDCustomers2026", CUSTOMERS);
  const priorities = ensureTable("Пріоритети", "tblRBDPriorities2026", PRIORITIES);
  const erpStage = ensureTable("ERP — вхідні", "tblRBDERPStage2026", ERP_STAGE);
  const erpImports = ensureTable("Журнал імпортів", "tblRBDERPImports2026", ERP_IMPORTS);

  seed(users, INITIAL_USERS);
  seed(statuses, INITIAL_STATUSES);
  seed(transitions, INITIAL_TRANSITIONS);
  seed(priorities, INITIAL_PRIORITIES);

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
    commands, timeline, notifications, cities, locations,
    categories, customers, priorities, erpStage, erpImports
  ].forEach(t => styleTable(t));

  // Корисна службова інформація на окремій вкладці.
  const info = sheet("Система");
  info.getRange("A1:B8").clear(ExcelScript.ClearApplyTo.all);
  info.getRange("A1:B8").setValues([
    ["РБД", "3.1"],
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

  return `РБД 3.1 готово. Персональні вкладки: ${personalTabs.join(", ")}. Старий «Кабінет керівника» та старі службові вкладки не видалялися автоматично — після перевірки їх можна прибрати вручну.`;
}
