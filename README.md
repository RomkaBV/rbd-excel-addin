# РБД 2.0 — чистий GitHub Pages комплект

Це актуальний комплект для `https://romkabv.github.io/rbd-excel-addin/`.

## Файли
- `taskpane.html` — єдиний інтерфейс.
- `taskpane.js` — єдина JS-логіка надбудови.
- `manifest-content.xml` — Content Add-in для вставки в аркуш Excel.
- `manifest.xml` — Task Pane manifest.
- `icon-32.png`, `icon-64.png`, `icon-80.png` — іконки.
- `RBD_ACTIVITY_SETUP.txt` — актуальний Office Script RBD 2.0 для структури книги та обробки команд.
- `index.html` — перенаправлення на `taskpane.html`.

Старі `adapter.js`, `taskpane-v3.js`, `test-seed.json` не потрібні.

## Таблиці, які використовує надбудова
`tblRBDRequests2026`, `tblRBDEvents2026`, `tblRBDAssignments2026`, `tblRBDUsers2026`, `tblRBDStatuses2026`, `tblRBDTransitions2026`, `tblRBDCities2026`, `tblRBDLocations2026`, `tblRBDCategories2026`, `tblRBDCustomers2026`, `tblRBDPriorities2026`, `tblRBDERPStage2026`, `tblRBDERPImports2026`, `tblRBDInputManager2026`, а також персональні `tblRBDInputUSRxxx`.

## Ролі
- Аркуш `Кабінет керівника` — режим керівника; бачить усі заявки.
- Персональні аркуші названі прізвищами (`Іванченко`, `Войцехівський` тощо) — режим виконавця; показуються лише власні заявки.

## Команди
Створення, призначення, зміна статусу, коментарі та закриття записуються в персональну таблицю команд зі станом `Нова`. Їх обробляє Office Script `RBD_ACTIVITY_SETUP` при наступному запуску/автоматизації.

## ERP
ERP імпорт доступний керівнику. Підтримуються `.xlsx`, `.xls`, `.csv`. Надбудова:
1. читає перший аркуш файлу;
2. автоматично шукає типові заголовки (`ID ERP`, `Тема`, `Опис`, `Замовник`, `Місто`, `Локація`, `Категорія`, `Пріоритет`);
3. перевіряє дублікати за `ID ERP`;
4. додає нові заявки зі статусом `Нова`, джерелом `ERP` і порожнім виконавцем;
5. пише staging у `tblRBDERPStage2026`, журнал у `tblRBDERPImports2026` та подію в `tblRBDEvents2026`.

Планова дата закриття не використовується.
