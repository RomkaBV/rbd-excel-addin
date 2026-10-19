type V = string | number | boolean;

function main(
  workbook: ExcelScript.Workbook,
  notificationId: string,
  status: string,
  errorText: string = "",
  sentAt: string = ""
): string {
  function findTable(name: string): ExcelScript.Table {
    for (const ws of workbook.getWorksheets()) {
      for (const t of ws.getTables()) {
        if (t.getName() === name) return t;
      }
    }
    throw new Error(`Не знайдено таблицю ${name}`);
  }

  function headers(t: ExcelScript.Table): string[] {
    return t.getHeaderRowRange().getTexts()[0];
  }

  function rows(t: ExcelScript.Table): V[][] {
    if (t.getRowCount() === 0) return [];
    return t.getRangeBetweenHeaderAndTotal().getValues() as V[][];
  }

  function get(row: V[], hh: string[], name: string): string {
    const i = hh.indexOf(name);
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

  function now(): string {
    return new Date().toISOString();
  }

  function uid(prefix: string): string {
    return prefix + "-" + new Date().getTime().toString() + "-" + Math.floor(Math.random() * 1000000).toString();
  }

  const notifications = findTable("tblRBDNotifications2026");
  const timeline = findTable("tblRBDTimeline2026");
  const hh = headers(notifications);
  const rr = rows(notifications);
  const index = rr.findIndex(r => get(r, hh, "ID сповіщення") === notificationId);
  if (index < 0) throw new Error(`Сповіщення ${notificationId} не знайдено`);

  const row = rr[index];
  const requestId = get(row, hh, "ID заявки");
  const recipient = get(row, hh, "Отримувач");
  const channel = get(row, hh, "Канал");
  const subject = get(row, hh, "Тема");
  const attempts = Number(get(row, hh, "Спроби") || "0") + 1;
  const finalSentAt = sentAt || (status.toLowerCase().includes("надісл") ? now() : "");

  put(notifications, index, "Статус", status);
  put(notifications, index, "Спроби", attempts);
  put(notifications, index, "Помилка", errorText);
  if (finalSentAt) put(notifications, index, "Дата відправки", finalSentAt);

  addObject(timeline, {
    "ID події": uid("EV"),
    "ID заявки": requestId,
    "Дата події": finalSentAt || now(),
    "Автор": "Система",
    "Тип": "Сповіщення",
    "Дія": "Результат відправки",
    "Текст": subject,
    "Було": "Очікує",
    "Стало": status,
    "Канал": channel,
    "Отримувач": recipient,
    "Статус доставки": status,
    "Джерело": "Power Automate",
    "ID команди": notificationId
  });

  return `Оновлено ${notificationId}: ${status}`;
}
