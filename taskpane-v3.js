const APP_JS_VERSION = "20260925-activity1";
console.log("RBD JS VERSION:", APP_JS_VERSION);

const STATE = {
  activeSheet: "",
  employeeMode: false,
  employee: "",
  actor: "",
  allRequests: [],
  requests: [],
  filtered: [],
  archive: [],
  categories: [],
  cities: [],
  employees: [],
  statuses: [],
  selectedRequest: null,
  expandedStatusId: "",
  activities: [],
  activityRequestId: "",
  activityFilter: "ALL",
  activityRefreshTimer: null
};


// ============================================================
// РОЛІ
// ============================================================

const MANAGER_NAME =
  "Іванченко В.М.";


const MANAGER_SHEETS = [
  "Іванченко В.М."
];


const EMPLOYEE_SHEETS = {

  "Войцехівський Г.В.":
    "Войцехівський Г.В.",

  "Ридванський П.С.":
    "Ридванський П.С.",

  "Галько А.І.":
    "Галько А.І.",

  "Трунов Ю.О.":
    "Трунов Ю.О.",

  "Желясков Д.О.":
    "Желясков Д.О.",

  "Слепущенко О.О.":
    "Слепущенко О.О.",

  "Сергеєв П.А.":
    "Сергеєв П.А.",

  "Туровський В.О.":
    "Туровський В.О."

};


// ============================================================
// РЕЗЕРВНІ ДОВІДНИКИ
// ============================================================

const FALLBACK_EMPLOYEES = [

  "Войцехівський Г.В.",
  "Ридванський П.С.",
  "Галько А.І.",
  "Трунов Ю.О.",
  "Желясков Д.О.",
  "Слепущенко О.О.",
  "Сергеєв П.А.",
  "Туровський В.О."

];


const FALLBACK_STATUSES = [

  "Нова",
  "Прийнята в роботу",
  "Пошук підрядника",
  "Погодження кошторису",
  "Укладання договору",
  "Погодження бюджету",
  "Виконання робіт",
  "Прийняття робіт",
  "Призупинена",
  "Закрита"

];


const FALLBACK_CATEGORIES = [

  "Будівельні роботи",
  "Ремонтні роботи",
  "Електромонтажні роботи",
  "Сантехнічні роботи",
  "Інженерні мережі",
  "Покрівельні роботи",
  "Оздоблювальні роботи",
  "Аварійні роботи",
  "Обслуговування",
  "Інше"

];


const FALLBACK_CITIES = [

  "Київ",
  "Чернігів",
  "Ніжин",
  "Суми",
  "Шостка",
  "Харків",
  "Полтава",
  "Дніпро",
  "Житомир",
  "Запоріжжя",
  "Кам'янське",
  "Кривий Ріг",
  "Кропивницький",
  "Миколаїв",
  "Одеса",
  "Павлоград",
  "Біла Церква",
  "Вінниця",
  "Львів",
  "Стрий",
  "Луцьк",
  "Рівне",
  "Тернопіль",
  "Хмельницький",
  "Чернівці",
  "Івано-Франківськ",
  "Мукачево",
  "Ковель",
  "Тячів",
  "Інше"

];


// ============================================================
// START
// ============================================================

Office.onReady(async info => {

  initEvents();
  initRequiredFields();


  if (
    info.host !==
    Office.HostType.Excel
  ) {

    return;

  }


  try {

    await detectMode();

    await loadDictionaries();

    configureMode();

    await loadDashboard();

  }

  catch (error) {

    console.error(error);

    showTableError(
      getErrorText(error)
    );

  }

});


// ============================================================
// EVENTS
// ============================================================

function initEvents() {


  document
    .getElementById(
      "refreshButton"
    )
    ?.addEventListener(
      "click",
      async () => {

        await detectMode();

        configureMode();

        await loadDashboard();

      }
    );


  document
    .getElementById(
      "newRequestButton"
    )
    ?.addEventListener(
      "click",
      openCreateModal
    );


  [
    "filterStatus",
    "filterExecutor",
    "filterCity",
    "filterCategory"
  ]
    .forEach(id => {

      document
        .getElementById(id)
        ?.addEventListener(
          "change",
          applyFilters
        );

    });


  document
    .getElementById(
      "searchInput"
    )
    ?.addEventListener(
      "input",
      applyFilters
    );


  document
    .getElementById(
      "saveCreateButton"
    )
    ?.addEventListener(
      "click",
      createRequest
    );


  document
    .getElementById(
      "saveEditButton"
    )
    ?.addEventListener(
      "click",
      saveEdit
    );


  document
    .querySelectorAll(
      "[data-close]"
    )
    .forEach(button => {

      button.addEventListener(
        "click",
        () => {

          closeModal(
            button.dataset.close
          );

        }
      );

    });


  document
    .getElementById(
      "clearCreateButton"
    )
    ?.addEventListener(
      "click",
      () => {

        clearCreateForm();

        if (STATE.employeeMode) {
          const executor = document.getElementById("createExecutor");
          if (executor) {
            executor.disabled = false;
            executor.value = STATE.employee;
            executor.disabled = true;
          }
        }

        refreshRequiredStates();

      }
    );


  document
    .getElementById(
      "activityCloseButton"
    )
    ?.addEventListener(
      "click",
      closeActivityDrawer
    );


  document
    .getElementById(
      "addNoteButton"
    )
    ?.addEventListener(
      "click",
      addNoteFromDrawer
    );


  document
    .getElementById(
      "addReminderButton"
    )
    ?.addEventListener(
      "click",
      addReminderFromDrawer
    );


  document
    .querySelectorAll(
      "[data-activity-filter]"
    )
    .forEach(button => {

      button.addEventListener(
        "click",
        () => {

          STATE.activityFilter =
            button.dataset.activityFilter ||
            "ALL";

          document
            .querySelectorAll(
              "[data-activity-filter]"
            )
            .forEach(item => {
              item.classList.toggle(
                "active",
                item === button
              );
            });

          renderActivity();

        }
      );

    });

}


// ============================================================
// ВИЗНАЧЕННЯ КАБІНЕТУ
// ============================================================

async function detectMode() {

  await Excel.run(
    async context => {

      const sheet =
        context.workbook
          .worksheets
          .getActiveWorksheet();


      sheet.load(
        "name"
      );


      await context.sync();


      STATE.activeSheet =
        cleanText(
          sheet.name
        );

    }
  );


  // ----------------------------------------------------------
  // ІВАНЧЕНКО — КЕРІВНИК
  // ----------------------------------------------------------

  if (
    MANAGER_SHEETS.includes(
      STATE.activeSheet
    )
  ) {

    STATE.employeeMode =
      false;


    STATE.employee =
      "";


    STATE.actor =
      MANAGER_NAME;


    return;

  }


  // ----------------------------------------------------------
  // ПЕРСОНАЛЬНИЙ КАБІНЕТ
  // ----------------------------------------------------------

  const employee =
    EMPLOYEE_SHEETS[
      STATE.activeSheet
    ];


  if (employee) {

    STATE.employeeMode =
      true;


    STATE.employee =
      employee;


    STATE.actor =
      employee;


    return;

  }


  // ----------------------------------------------------------
  // ТЕХНІЧНІ ВКЛАДКИ *. Таб. НЕ ВИЗНАЧАЮТЬ РОЛЬ
  // ----------------------------------------------------------

  if (
    STATE.activeSheet.endsWith(
      " Таб."
    )
  ) {

    if (
      STATE.actor
    ) {

      return;

    }


    throw new Error(
      "Відкрийте вкладку кабінету керівника або виконавця."
    );

  }


  // ----------------------------------------------------------
  // ІНШІ ВКЛАДКИ ТАКОЖ НЕ ДАЮТЬ РОЛЬ КЕРІВНИКА
  // ----------------------------------------------------------

  if (
    STATE.actor
  ) {

    return;

  }


  throw new Error(
    "Не вдалося визначити кабінет за активною вкладкою: " +
    STATE.activeSheet
  );

}


// ============================================================
// НАЛАШТУВАННЯ ІНТЕРФЕЙСУ
// ============================================================

function configureMode() {

  const executorFilter =
    document.getElementById(
      "filterExecutorField"
    );


  const createExecutor =
    document.getElementById(
      "createExecutor"
    );


  // ----------------------------------------------------------
  // ВИКОНАВЕЦЬ
  // ----------------------------------------------------------

  if (
    STATE.employeeMode
  ) {

    setText(
      "pageTitle",
      "РБД — " +
      STATE.employee
    );


    setText(
      "pageSubtitle",
      "Мої заявки • Опрацювання • Контроль"
    );


    setText(
      "currentUserName",
      STATE.employee
    );


    setText(
      "currentUserRole",
      "Виконавець"
    );


    setText(
      "userAvatar",
      getInitials(
        STATE.employee
      )
    );


    if (
      executorFilter
    ) {

      executorFilter.style.display =
        "none";

    }


    if (
      createExecutor
    ) {

      createExecutor.disabled =
        false;


      createExecutor.value =
        STATE.employee;


      createExecutor.disabled =
        true;

    }


    return;

  }


  // ----------------------------------------------------------
  // КЕРІВНИК
  // ----------------------------------------------------------

  setText(
    "pageTitle",
    "РБД — Кабінет керівника"
  );


  setText(
    "pageSubtitle",
    "Усі заявки • Усі виконавці • Контроль • Результат"
  );


  setText(
    "currentUserName",
    MANAGER_NAME
  );


  setText(
    "currentUserRole",
    "Керівник РБД"
  );


  setText(
    "userAvatar",
    getInitials(
      MANAGER_NAME
    )
  );


  if (
    executorFilter
  ) {

    executorFilter.style.display =
      "";

  }


  if (
    createExecutor
  ) {

    createExecutor.disabled =
      false;

  }

}


// ============================================================
// ДОВІДНИКИ
// ============================================================

async function loadDictionaries() {

  const data =
    await Excel.run(
      async context => {


        const categories =
          await readFirstColumnSafe(
            context,
            "tbl_RBD_Categories"
          );


        const cities =
          await readFirstColumnSafe(
            context,
            "tbl_RBD_Cities"
          );


        const employees =
          await readFirstColumnSafe(
            context,
            "tbl_RBD_Employees"
          );


        const statuses =
          await readFirstColumnSafe(
            context,
            "tbl_SLA"
          );


        return {

          categories,
          cities,
          employees,
          statuses

        };

      }
    );


  STATE.categories =
    data.categories.length
      ? data.categories
      : FALLBACK_CATEGORIES.slice();


  STATE.cities =
    data.cities.length
      ? data.cities
      : FALLBACK_CITIES.slice();


  const employeeSource =
    data.employees.length
      ? data.employees
      : FALLBACK_EMPLOYEES.slice();


  STATE.employees =
    employeeSource.filter(
      employee =>
        normalizeName(
          employee
        ) !==
        normalizeName(
          MANAGER_NAME
        )
    );


  STATE.statuses =
    data.statuses.length
      ? data.statuses
      : FALLBACK_STATUSES.slice();


  fillSelect(
    "filterStatus",
    STATE.statuses,
    "Усі статуси"
  );


  fillSelect(
    "filterExecutor",
    STATE.employees,
    "Усі виконавці"
  );


  fillSelect(
    "filterCity",
    STATE.cities,
    "Усі міста"
  );


  fillSelect(
    "filterCategory",
    STATE.categories,
    "Усі категорії"
  );


  fillSelect(
    "createCategory",
    STATE.categories,
    "Оберіть категорію"
  );


  fillSelect(
    "createCity",
    STATE.cities,
    "Оберіть місто"
  );


  fillSelect(
    "createExecutor",
    STATE.employees,
    "Оберіть виконавця"
  );


  fillSelect(
    "editCategory",
    STATE.categories,
    "Оберіть категорію"
  );


  fillSelect(
    "editCity",
    STATE.cities,
    "Оберіть місто"
  );

}


// ============================================================
// БЕЗПЕЧНЕ ЧИТАННЯ ДОВІДНИКА
// ============================================================

async function readFirstColumnSafe(
  context,
  tableName
) {

  const table =
    context.workbook
      .tables
      .getItemOrNullObject(
        tableName
      );


  table.load(
    "isNullObject"
  );


  await context.sync();


  if (
    table.isNullObject
  ) {

    return [];

  }


  table.rows.load(
    "items"
  );


  await context.sync();


  if (
    table.rows.items.length ===
    0
  ) {

    return [];

  }


  const body =
    table.getDataBodyRange();


  body.load(
    "values"
  );


  await context.sync();


  return body.values

    .map(
      row =>
        String(
          row[0] ?? ""
        ).trim()
    )

    .filter(
      value =>
        value !== ""
    );

}


// ============================================================
// ЧИТАННЯ ТАБЛИЦІ
// ============================================================

async function readTable(
  context,
  tableName,
  required = true
) {

  const table =
    context.workbook
      .tables
      .getItemOrNullObject(
        tableName
      );


  table.load(
    "isNullObject"
  );


  await context.sync();


  if (
    table.isNullObject
  ) {

    if (
      required
    ) {

      throw new Error(
        "Не знайдено таблицю " +
        tableName
      );

    }


    return {

      headers: [],
      rows: []

    };

  }


  const header =
    table.getHeaderRowRange();


  header.load(
    "values"
  );


  table.rows.load(
    "items"
  );


  await context.sync();


  if (
    table.rows.items.length ===
    0
  ) {

    return {

      headers:
        header.values[0],

      rows: []

    };

  }


  const body =
    table.getDataBodyRange();


  body.load(
    "values"
  );


  await context.sync();


  return {

    headers:
      header.values[0],

    rows:
      body.values

  };

}


// ============================================================
// ЗАВАНТАЖЕННЯ ЗАЯВОК
// ============================================================

async function loadDashboard() {

  showLoading();


  try {

    const result =
      await Excel.run(
        async context => {


          const base =
            await readTable(
              context,
              "tbl_RBD_Base",
              true
            );


          const archive =
            await readTable(
              context,
              "tbl_RBD_Archive",
              false
            );


          return {

            base,
            archive

          };

        }
      );


    // --------------------------------------------------------
    // ВСЯ БАЗА
    // --------------------------------------------------------

    STATE.allRequests =
      result.base.rows

        .map(
          row =>
            rowToRequest(
              result.base.headers,
              row
            )
        )

        .filter(
          request =>
            request.id !== ""
        );


    STATE.archive =
      result.archive.rows

        .map(
          row =>
            rowToRequest(
              result.archive.headers,
              row
            )
        )

        .filter(
          request =>
            request.id !== ""
        );


    // --------------------------------------------------------
    // ВИКОНАВЕЦЬ / КЕРІВНИК
    // --------------------------------------------------------

    if (
      STATE.employeeMode
    ) {

      STATE.requests =
        STATE.allRequests.filter(
          request =>

            normalizeName(
              request.executor
            ) ===

            normalizeName(
              STATE.employee
            )

        );

    }

    else {

      STATE.requests =
        STATE.allRequests.slice();

    }


    STATE.expandedStatusId =
      "";


    updateKpis();

    applyFilters();

    updateLastUpdated();

  }

  catch (error) {

    console.error(error);


    showTableError(
      getErrorText(
        error
      )
    );

  }

}


// ============================================================
// EXCEL ROW → REQUEST
// ============================================================

function rowToRequest(
  headers,
  row
) {

  function get(
    name
  ) {

    const index =
      headers.indexOf(
        name
      );


    if (
      index < 0
    ) {

      return "";

    }


    return row[index] ?? "";

  }


  return {

    sd:
      cleanText(
        get(
          "Номер заявки SD"
        )
      ),

    category:
      cleanText(
        get(
          "Категорія"
        )
      ),

    description:
      cleanText(
        get(
          "Опис"
        )
      ),

    city:
      cleanText(
        get(
          "Місто"
        )
      ),

    address:
      cleanText(
        get(
          "Адреса / локація"
        )
      ),

    created:
      get(
        "Дата створення"
      ),

    amount:
      toNumber(
        get(
          "Сума витрат"
        )
      ),

    customer:
      cleanText(
        get(
          "Замовник"
        )
      ),

    status:
      cleanText(
        get(
          "Статус"
        )
      ),

    executor:
      cleanText(
        get(
          "Виконавець"
        )
      ),

    id:
      cleanText(
        get(
          "ID"
        )
      ),

    source:
      cleanText(
        get(
          "Джерело заявки"
        )
      ),

    statusSince:
      get(
        "Поточний статус з"
      ),

    previousStatus:
      cleanText(
        get(
          "Попередній статус"
        )
      ),

    sla:
      toNumber(
        get(
          "SLA, год"
        )
      ),

    storedTimeInStatus:
      toNumber(
        get(
          "Час у статусі, год"
        )
      ),

    slaOverdue:
      cleanText(
        get(
          "Прострочено SLA"
        )
      ),

    plannedDate:
      get(
        "Планова дата завершення"
      ),

    dateOverdue:
      cleanText(
        get(
          "Прострочено по даті"
        )
      ),

    pauseReason:
      cleanText(
        get(
          "Причина призупинення"
        )
      ),

    comment:
      cleanText(
        get(
          "Коментар"
        )
      ),

    closedDate:
      get(
        "Дата закриття"
      )

  };

}


// ============================================================
// ФІЛЬТРИ
// ============================================================

function applyFilters() {

  const status =
    valueOf(
      "filterStatus"
    );


  const executor =
    valueOf(
      "filterExecutor"
    );


  const city =
    valueOf(
      "filterCity"
    );


  const category =
    valueOf(
      "filterCategory"
    );


  const search =
    valueOf(
      "searchInput"
    )
      .toLowerCase();


  STATE.filtered =
    STATE.requests.filter(
      request => {


        if (
          status &&
          request.status !==
            status
        ) {

          return false;

        }


        if (
          !STATE.employeeMode &&
          executor &&

          normalizeName(
            request.executor
          ) !==

          normalizeName(
            executor
          )
        ) {

          return false;

        }


        if (
          city &&
          request.city !==
            city
        ) {

          return false;

        }


        if (
          category &&
          request.category !==
            category
        ) {

          return false;

        }


        if (search) {

          const searchable =
            [

              request.id,
              request.sd,
              request.category,
              request.description,
              request.city,
              request.customer,
              request.executor,
              request.status,
              request.comment

            ]
              .join(" ")
              .toLowerCase();


          if (
            !searchable.includes(
              search
            )
          ) {

            return false;

          }

        }


        return true;

      }
    );


  renderRequests();

  updateSummary();

}


// ============================================================
// KPI
// ============================================================

function updateKpis() {

  const active =
    STATE.requests;


  const overdue =
    active.filter(
      request =>
        isOverdue(
          request
        )
    );


  const work =
    active.filter(
      request =>

        request.status !==
          "Нова" &&

        request.status !==
          "Призупинена" &&

        request.status !==
          "Закрита"

    );


  let archiveScope =
    STATE.archive;


  if (
    STATE.employeeMode
  ) {

    archiveScope =
      archiveScope.filter(
        request =>

          normalizeName(
            request.executor
          ) ===

          normalizeName(
            STATE.employee
          )

      );

  }


  const now =
    Date.now();


  const thirtyDays =
    30 *
    24 *
    60 *
    60 *
    1000;


  const closed =
    archiveScope.filter(
      request => {


        const date =
          dateToMilliseconds(
            request.closedDate
          );


        return (
          date &&
          now - date <=
          thirtyDays
        );

      }
    );


  setText(
    "kpiActive",
    active.length
  );


  setText(
    "kpiActiveSum",
    money(
      sumRequests(
        active
      )
    )
  );


  setText(
    "kpiOverdue",
    overdue.length
  );


  setText(
    "kpiOverdueSum",
    money(
      sumRequests(
        overdue
      )
    )
  );


  setText(
    "kpiWork",
    work.length
  );


  setText(
    "kpiWorkSum",
    money(
      sumRequests(
        work
      )
    )
  );


  setText(
    "kpiClosed",
    closed.length
  );


  setText(
    "kpiClosedSum",
    money(
      sumRequests(
        closed
      )
    )
  );

}


// ============================================================
// ПІДСУМКИ
// ============================================================

function updateSummary() {

  const allSum =
    sumRequests(
      STATE.requests
    );


  const visibleSum =
    sumRequests(
      STATE.filtered
    );


  setText(
    "summaryAllCount",
    STATE.requests.length
  );


  setText(
    "summaryAllSum",
    money(
      allSum
    )
  );


  setText(
    "summaryVisibleCount",
    STATE.filtered.length
  );


  setText(
    "summaryVisibleSum",
    money(
      visibleSum
    )
  );


  setText(

    "tableHeaderSummary",

    "Показано " +
    STATE.filtered.length +

    " із " +
    STATE.requests.length +

    " заявок • " +

    money(
      visibleSum
    )

  );

}


// ============================================================
// СУМА
// ============================================================

function sumRequests(
  requests
) {

  return requests.reduce(
    (
      sum,
      request
    ) =>

      sum +
      toNumber(
        request.amount
      ),

    0
  );

}


// ============================================================
// ПРОСТРОЧЕННЯ
// ============================================================

function isOverdue(
  request
) {

  if (
    request.status ===
      "Призупинена" ||

    request.status ===
      "Закрита"
  ) {

    return false;

  }


  if (
    isSlaOverdue(
      request
    )
  ) {

    return true;

  }


  if (
    request.dateOverdue ===
      "Так"
  ) {

    return true;

  }


  const planned =
    dateToMilliseconds(
      request.plannedDate
    );


  if (
    planned &&
    planned <
      startOfToday()
  ) {

    return true;

  }


  return false;

}


// ============================================================
// SLA OVERDUE
// ============================================================

function isSlaOverdue(
  request
) {

  if (
    request.status ===
      "Призупинена" ||

    request.status ===
      "Закрита"
  ) {

    return false;

  }


  if (
    request.slaOverdue ===
      "Так"
  ) {

    return true;

  }


  const elapsed =
    currentStatusHours(
      request
    );


  return (

    request.sla > 0 &&

    elapsed !== null &&

    elapsed >
      request.sla

  );

}


// ============================================================
// ВІДОБРАЖЕННЯ
// ============================================================

function renderRequests() {

  const body =
    document.getElementById(
      "requestsBody"
    );


  if (!body) {

    return;

  }


  if (
    STATE.filtered.length ===
    0
  ) {

    body.innerHTML = `

      <tr>

        <td colspan="13">

          <div class="empty-state">

            Заявок за обраними умовами немає

          </div>

        </td>

      </tr>

    `;


    return;

  }


  let html =
    "";


  STATE.filtered.forEach(
    request => {


      html +=
        buildRequestRow(
          request
        );


      if (
        STATE.expandedStatusId ===
        request.id
      ) {

        html +=
          buildStatusPanel(
            request
          );

      }

    }
  );


  body.innerHTML =
    html;


  bindRequestActions();

  bindStatusActions();

}


// ============================================================
// РЯДОК ЗАЯВКИ
// ============================================================

function buildRequestRow(
  request
) {

  const rowClass =
    getRowClass(
      request
    );


  const statusClass =
    getStatusClass(
      request.status
    );


  const elapsed =
    currentStatusHours(
      request
    );


  return `

    <tr class="request-row ${rowClass}">


      <td>

        <div class="request-id">

          ${escapeHtml(
            request.id
          )}

        </div>


        ${
          request.sd
            ? `

              <div
                style="
                  margin-top:3px;
                  font-size:8px;
                  opacity:.72;
                "
              >

                ${escapeHtml(
                  request.sd
                )}

              </div>

            `
            : ""
        }

      </td>


      <td>

        <div
          class="ellipsis"
          title="${escapeHtml(
            request.category
          )}"
        >

          ${escapeHtml(
            request.category
          )}

        </div>

      </td>


      <td>

        <div
          class="ellipsis"
          title="${escapeHtml(
            request.description
          )}"
        >

          ${escapeHtml(
            request.description
          )}

        </div>

      </td>


      <td>

        ${escapeHtml(
          request.city
        )}

      </td>


      <td>

        <div
          class="ellipsis"
          title="${escapeHtml(
            request.customer
          )}"
        >

          ${escapeHtml(
            request.customer
          )}

        </div>

      </td>


      <td>

        <div
          class="ellipsis"
          title="${escapeHtml(
            request.executor
          )}"
        >

          ${escapeHtml(
            request.executor ||
            "—"
          )}

        </div>

      </td>


      <td>

        <button
          class="
            status-button
            ${statusClass}
          "
          data-status-open="${escapeHtml(
            request.id
          )}"
          type="button"
          title="Натисніть для зміни статусу"
        >

          <span class="status-dot"></span>

          ${escapeHtml(
            request.status
          )}

        </button>

      </td>


      <td
        style="
          text-align:right;
          font-weight:650;
        "
      >

        ${moneyCompact(
          request.amount
        )}

      </td>


      <td
        class="${
          isOverdue(
            request
          )
            ? "overdue-text"
            : ""
        }"
      >

        ${formatExcelDate(
          request.plannedDate
        )}

      </td>


      <td>

        ${
          elapsed !== null
            ? elapsed +
              " год."
            : "—"
        }

      </td>


      <td>

        ${
          request.sla > 0
            ? request.sla +
              " год."
            : "—"
        }

      </td>


      <td>

        <div
          class="ellipsis"
          title="${escapeHtml(
            request.comment
          )}"
        >

          ${escapeHtml(
            request.comment ||
            "—"
          )}

        </div>

      </td>


      <td>

        <div class="actions">

          <button
            class="action-button"
            data-view="${escapeHtml(
              request.id
            )}"
            type="button"
            title="Деталі"
          >
            ◉
          </button>

          <button
            class="action-button edit"
            data-edit="${escapeHtml(
              request.id
            )}"
            type="button"
            title="Редагувати"
          >
            ✎
          </button>

          <button
            class="action-button activity"
            data-activity="${escapeHtml(
              request.id
            )}"
            type="button"
            title="Історія та примітки"
          >
            ☷
          </button>

        </div>

      </td>


    </tr>

  `;

}


// ============================================================
// ПАНЕЛЬ СТАТУСУ
// ============================================================

function buildStatusPanel(
  request
) {

  const buttons =
    STATE.statuses

      .map(
        status => {


          const current =
            status ===
            request.status;


          return `

            <button
              class="
                status-option
                ${getStatusButtonClass(
                  status
                )}
                ${
                  current
                    ? "current"
                    : ""
                }
              "
              data-inline-status="${escapeHtml(
                status
              )}"
              data-request-id="${escapeHtml(
                request.id
              )}"
              type="button"
              ${
                current
                  ? "disabled"
                  : ""
              }
            >

              ${escapeHtml(
                status
              )}

            </button>

          `;

        }
      )

      .join("");


  const safeId =
    domSafe(
      request.id
    );


  return `

    <tr class="status-panel-row">

      <td colspan="13">

        <div class="status-panel">


          <div class="status-panel-header">

            <div class="status-panel-title">

              ${escapeHtml(
                request.id
              )}

              • Поточний статус:

              ${escapeHtml(
                request.status
              )}

            </div>


            <button
              class="status-panel-close"
              data-status-close="1"
              type="button"
            >
              ×
            </button>

          </div>


          <textarea
            id="inlineStatusComment"
            class="status-comment"
            placeholder="Коментар до зміни статусу — необов'язково"
          ></textarea>


          <div class="status-options">

            ${buttons}

          </div>


          <div
            id="specialStatusArea_${safeId}"
            class="special-status-area"
          ></div>


          <div
            id="inlineStatusMessage_${safeId}"
            class="inline-message"
          ></div>


        </div>

      </td>

    </tr>

  `;

}


// ============================================================
// ДІЇ
// ============================================================

function bindRequestActions() {


  document
    .querySelectorAll(
      "[data-status-open]"
    )
    .forEach(button => {

      button.addEventListener(
        "click",
        () => {


          const requestId =
            button.dataset.statusOpen;


          STATE.expandedStatusId =

            STATE.expandedStatusId ===
              requestId

              ? ""

              : requestId;


          renderRequests();

        }
      );

    });


  document
    .querySelectorAll(
      "[data-view]"
    )
    .forEach(button => {

      button.addEventListener(
        "click",
        () => {


          const request =
            findRequest(
              button.dataset.view
            );


          if (
            request
          ) {

            openDetails(
              request
            );

          }

        }
      );

    });


  document
    .querySelectorAll(
      "[data-edit]"
    )
    .forEach(button => {

      button.addEventListener(
        "click",
        () => {


          const request =
            findRequest(
              button.dataset.edit
            );


          if (
            request
          ) {

            openEdit(
              request
            );

          }

        }
      );

    });


  document
    .querySelectorAll(
      "[data-activity]"
    )
    .forEach(button => {

      button.addEventListener(
        "click",
        () => {

          const request =
            findRequest(
              button.dataset.activity
            );

          if (request) {
            openActivityDrawer(request);
          }

        }
      );

    });

}


// ============================================================
// INLINE STATUS ACTIONS
// ============================================================

function bindStatusActions() {


  document
    .querySelectorAll(
      "[data-status-close]"
    )
    .forEach(button => {

      button.addEventListener(
        "click",
        () => {

          STATE.expandedStatusId =
            "";

          renderRequests();

        }
      );

    });


  document
    .querySelectorAll(
      "[data-inline-status]"
    )
    .forEach(button => {

      button.addEventListener(
        "click",
        async () => {


          const request =
            findRequest(
              button.dataset.requestId
            );


          const newStatus =
            button.dataset.inlineStatus;


          if (
            !request ||
            !newStatus
          ) {

            return;

          }


          if (
            newStatus ===
            "Призупинена"
          ) {

            showPauseBox(
              request
            );

            return;

          }


          if (
            newStatus ===
            "Закрита"
          ) {

            showCloseBox(
              request
            );

            return;

          }


          const comment =
            valueOf(
              "inlineStatusComment"
            );


          await queueStatusChange(

            request,

            newStatus,

            comment,

            ""

          );

        }
      );

    });

}


// ============================================================
// ПРИЗУПИНЕННЯ
// ============================================================

function showPauseBox(
  request
) {

  const area =
    document.getElementById(

      "specialStatusArea_" +
      domSafe(
        request.id
      )

    );


  if (!area) {

    return;

  }


  area.innerHTML = `

    <div class="special-box">

      <input
        id="pauseReasonInline"
        type="text"
        placeholder="Причина призупинення — обов'язково"
      >

      <button
        id="confirmPauseButton"
        class="inline-confirm"
        type="button"
      >
        Призупинити
      </button>

      <button
        id="cancelSpecialButton"
        class="inline-cancel"
        type="button"
      >
        Скасувати
      </button>

    </div>

  `;


  document
    .getElementById(
      "confirmPauseButton"
    )
    ?.addEventListener(
      "click",
      async () => {


        const reason =
          valueOf(
            "pauseReasonInline"
          );


        if (
          !reason
        ) {

          showInlineError(
            request.id,
            "Вкажіть причину призупинення."
          );

          return;

        }


        const comment =
          valueOf(
            "inlineStatusComment"
          );


        await queueStatusChange(

          request,

          "Призупинена",

          comment,

          reason

        );

      }
    );


  document
    .getElementById(
      "cancelSpecialButton"
    )
    ?.addEventListener(
      "click",
      () => {

        area.innerHTML =
          "";

      }
    );

}


// ============================================================
// ЗАКРИТТЯ
// ============================================================

function showCloseBox(
  request
) {

  const area =
    document.getElementById(

      "specialStatusArea_" +
      domSafe(
        request.id
      )

    );


  if (!area) {

    return;

  }


  area.innerHTML = `

    <div class="special-box">

      <div
        style="
          flex:1;
          font-size:10px;
          color:#445d72;
        "
      >

        Закрити заявку

        <strong>
          ${escapeHtml(
            request.id
          )}
        </strong>?

      </div>


      <button
        id="confirmCloseButton"
        class="inline-confirm"
        type="button"
      >
        Так, закрити
      </button>


      <button
        id="cancelSpecialButton"
        class="inline-cancel"
        type="button"
      >
        Скасувати
      </button>

    </div>

  `;


  document
    .getElementById(
      "confirmCloseButton"
    )
    ?.addEventListener(
      "click",
      async () => {


        const comment =
          valueOf(
            "inlineStatusComment"
          );


        await queueStatusChange(

          request,

          "Закрита",

          comment,

          ""

        );

      }
    );


  document
    .getElementById(
      "cancelSpecialButton"
    )
    ?.addEventListener(
      "click",
      () => {

        area.innerHTML =
          "";

      }
    );

}


// ============================================================
// ЗМІНА СТАТУСУ → ЧЕРГА
// ============================================================

async function queueStatusChange(
  request,
  newStatus,
  comment,
  pauseReason
) {

  try {

    setInlineButtonsDisabled(
      true
    );


    showInlineSuccess(
      request.id,
      "Передаємо зміну в чергу..."
    );


    await enqueue({

      eventId:
        createId(
          "EVT"
        ),

      requestKey:
        "",

      requestId:
        request.id,

      operation:
        newStatus === "Закрита"
          ? "CLOSE"
          : "STATUS_CHANGE",

      actor:
        STATE.actor,

      source:
        STATE.employeeMode
          ? "Виконавець"
          : "Керівник",

      executor:
        request.executor,

      expectedStatus:
        request.status,

      newStatus,

      sd: "",

      category: "",

      description: "",

      city: "",

      address: "",

      amount: "",

      customer: "",

      plannedDate: "",

      comment,

      pauseReason,

      newExecutor:
        ""

    });


    // Після успішного запису в чергу одразу закриваємо панель статусів.
    STATE.expandedStatusId = "";

    renderRequests();


    // Не підміняємо статус локально. Чекаємо, поки 03_ProcessQueue
    // реально змінить tbl_RBD_Base / перенесе закриту заявку в архів.
    waitForStatusUpdate(
      request.id,
      newStatus,
      request.status,
      comment
    ).catch(error => {
      console.warn(
        "Автооновлення статусу:",
        error
      );
    });

  }

  catch (error) {

    showInlineError(
      request.id,
      getErrorText(
        error
      )
    );

  }

  finally {

    setInlineButtonsDisabled(
      false
    );

  }

}


// ============================================================
// ОЧІКУВАННЯ ОБРОБКИ ЧЕРГИ
// ============================================================

async function waitForStatusUpdate(
  requestId,
  expectedStatus,
  previousStatus = "",
  comment = ""
) {

  const maxAttempts = 30;
  const delayMs = 1500;


  for (
    let attempt = 1;
    attempt <= maxAttempts;
    attempt++
  ) {

    await sleep(
      delayMs
    );


    try {

      const result =
        await Excel.run(
          async context => {

            const base =
              await readTable(
                context,
                "tbl_RBD_Base",
                true
              );


            const request =
              base.rows
                .map(
                  row =>
                    rowToRequest(
                      base.headers,
                      row
                    )
                )
                .find(
                  item =>
                    item.id ===
                    requestId
                );


            return request
              ? {
                  exists: true,
                  status: request.status
                }
              : {
                  exists: false,
                  status: ""
                };

          }
        );


      // Звичайна зміна статусу: чекаємо появи нового статусу в Base.
      if (
        result.exists &&
        result.status ===
          expectedStatus
      ) {

        await appendActivitySafe(
          requestId,
          "STATUS",
          buildStatusActivityText(
            previousStatus,
            expectedStatus,
            comment
          )
        );

        await loadDashboard();

        return true;

      }


      // Закрита заявка може вже бути перенесена 03_ProcessQueue
      // з Base до Archive. Відсутність у Base після CLOSE вважаємо успіхом.
      if (
        expectedStatus ===
          "Закрита" &&
        !result.exists
      ) {

        await appendActivitySafe(
          requestId,
          "STATUS",
          buildStatusActivityText(
            previousStatus,
            expectedStatus,
            comment
          )
        );

        await loadDashboard();

        return true;

      }

    }

    catch (error) {

      console.warn(
        "Очікування обробки черги:",
        error
      );

    }

  }


  // Якщо Office Script працював довше, робимо фінальне перечитування.
  await loadDashboard();

  return false;

}


// ============================================================
// ЗАТРИМКА
// ============================================================

function sleep(
  ms
) {

  return new Promise(
    resolve =>
      setTimeout(
        resolve,
        ms
      )
  );

}


// ============================================================

function setInlineButtonsDisabled(
  disabled
) {

  document
    .querySelectorAll(
      "[data-inline-status]"
    )
    .forEach(button => {


      if (
        !button.classList.contains(
          "current"
        )
      ) {

        button.disabled =
          disabled;

      }

    });

}


// ============================================================
// INLINE MESSAGES
// ============================================================

function showInlineSuccess(
  requestId,
  message
) {

  const element =
    document.getElementById(

      "inlineStatusMessage_" +
      domSafe(
        requestId
      )

    );


  if (!element) {

    return;

  }


  element.className =
    "inline-message success";


  element.textContent =
    message;

}


// ============================================================

function showInlineError(
  requestId,
  message
) {

  const element =
    document.getElementById(

      "inlineStatusMessage_" +
      domSafe(
        requestId
      )

    );


  if (!element) {

    return;

  }


  element.className =
    "inline-message error";


  element.textContent =
    "Помилка: " +
    message;

}


// ============================================================
// CREATE FORM
// ============================================================

function openCreateModal() {

  clearCreateForm();


  const executor =
    document.getElementById(
      "createExecutor"
    );


  if (
    STATE.employeeMode
  ) {

    if (
      executor
    ) {

      executor.disabled =
        false;


      executor.value =
        STATE.employee;


      executor.disabled =
        true;

    }

  }

  else {

    if (
      executor
    ) {

      executor.disabled =
        false;


      executor.value =
        "";

    }

  }


  openModal(
    "createModal"
  );

  refreshRequiredStates();

}


// ============================================================
// CREATE REQUEST
// ============================================================

async function createRequest() {

  const category =
    valueOf(
      "createCategory"
    );


  const description =
    valueOf(
      "createDescription"
    );


  const city =
    valueOf(
      "createCity"
    );


  const customer =
    valueOf(
      "createCustomer"
    );


  const executor =

    STATE.employeeMode

      ? STATE.employee

      : valueOf(
          "createExecutor"
        );


  markRequiredErrors();


  if (
    !category ||
    !description ||
    !city ||
    !customer ||
    !executor
  ) {

    showMessageError(

      "createMessage",

      "Заповніть усі обов'язкові поля."

    );


    return;

  }


  setButtonBusy(

    "saveCreateButton",

    true,

    "Створюємо..."

  );


  try {

    await enqueue({

      eventId:
        createId(
          "EVT"
        ),

      requestKey:
        createId(
          "REQ"
        ),

      requestId:
        "",

      operation:
        "CREATE",

      actor:
        STATE.actor,

      source:
        STATE.employeeMode
          ? "Виконавець"
          : "Керівник",

      executor,

      expectedStatus:
        "",

      newStatus:
        "Нова",

      sd:
        valueOf(
          "createSD"
        ),

      category,

      description,

      city,

      address:
        valueOf(
          "createAddress"
        ),

      amount:
        numberOf(
          "createAmount"
        ),

      customer,

      plannedDate:
        valueOf(
          "createPlannedDate"
        ),

      comment:
        valueOf(
          "createComment"
        ),

      pauseReason:
        "",

      newExecutor:
        ""

    });


    closeModal(
      "createModal"
    );

    clearCreateForm();

    scheduleDashboardRefresh();

  }

  catch (error) {

    showMessageError(

      "createMessage",

      getErrorText(
        error
      )

    );

  }

  finally {

    setButtonBusy(

      "saveCreateButton",

      false,

      "Створити заявку"

    );

  }

}


// ============================================================
// ОБОВ'ЯЗКОВІ ПОЛЯ CREATE
// ============================================================

function initRequiredFields() {

  [
    "createCategory",
    "createDescription",
    "createCity",
    "createCustomer",
    "createExecutor"
  ].forEach(id => {

    const element =
      document.getElementById(id);

    if (!element) {
      return;
    }

    const handler = () => {
      updateRequiredFieldState(element);
    };

    element.addEventListener("input", handler);
    element.addEventListener("change", handler);

  });

}


function updateRequiredFieldState(element) {

  if (!element) {
    return;
  }

  const filled =
    String(element.value ?? "")
      .trim() !== "";

  element.classList.toggle(
    "required-filled",
    filled
  );

  if (filled) {
    element.classList.remove(
      "required-error"
    );
  }

}


function refreshRequiredStates() {

  [
    "createCategory",
    "createDescription",
    "createCity",
    "createCustomer",
    "createExecutor"
  ].forEach(id => {
    updateRequiredFieldState(
      document.getElementById(id)
    );
  });

}


function markRequiredErrors() {

  [
    "createCategory",
    "createDescription",
    "createCity",
    "createCustomer",
    "createExecutor"
  ].forEach(id => {

    const element =
      document.getElementById(id);

    if (!element) {
      return;
    }

    const empty =
      String(element.value ?? "")
        .trim() === "";

    element.classList.toggle(
      "required-error",
      empty
    );

  });

}


function scheduleDashboardRefresh() {

  [2500, 6000].forEach(delay => {

    setTimeout(
      () => {
        loadDashboard()
          .catch(error => {
            console.warn(
              "Автооновлення після CREATE:",
              error
            );
          });
      },
      delay
    );

  });

}


// ============================================================
// DETAILS
// ============================================================

function openDetails(
  request
) {

  STATE.selectedRequest =
    request;


  setText(
    "detailsTitle",
    request.id
  );


  const elapsed =
    currentStatusHours(
      request
    );


  const rows = [

    [
      "Номер SD",
      request.sd
    ],

    [
      "Категорія",
      request.category
    ],

    [
      "Опис",
      request.description
    ],

    [
      "Місто",
      request.city
    ],

    [
      "Адреса",
      request.address
    ],

    [
      "Дата створення",
      formatExcelDateTime(
        request.created
      )
    ],

    [
      "Сума",
      money(
        request.amount
      )
    ],

    [
      "Замовник",
      request.customer
    ],

    [
      "Виконавець",
      request.executor
    ],

    [
      "Статус",
      request.status
    ],

    [
      "Попередній статус",
      request.previousStatus
    ],

    [
      "Поточний статус з",
      formatExcelDateTime(
        request.statusSince
      )
    ],

    [
      "Час у статусі",

      elapsed !== null

        ? elapsed +
          " год."

        : "—"
    ],

    [
      "SLA",

      request.sla > 0

        ? request.sla +
          " год."

        : "—"
    ],

    [
      "Планова дата",

      formatExcelDate(
        request.plannedDate
      )
    ],

    [
      "Прострочено SLA",

      isSlaOverdue(
        request
      )
        ? "Так"
        : "Ні"
    ],

    [
      "Причина призупинення",
      request.pauseReason
    ],

    [
      "Коментар",
      request.comment
    ]

  ];


  const content =
    document.getElementById(
      "detailsContent"
    );


  if (
    content
  ) {

    content.innerHTML =
      rows
        .map(
          row => `

            <div class="details-label">

              ${escapeHtml(
                row[0]
              )}

            </div>

            <div class="details-value">

              ${escapeHtml(
                String(
                  row[1] ||
                  "—"
                )
              )}

            </div>

          `
        )
        .join("");

  }


  openModal(
    "detailsModal"
  );

}


// ============================================================
// EDIT
// ============================================================

function openEdit(
  request
) {

  STATE.selectedRequest =
    request;


  setText(
    "editTitle",
    "Редагування " +
    request.id
  );


  setInputValue(
    "editSD",
    request.sd
  );


  setInputValue(
    "editCategory",
    request.category
  );


  setInputValue(
    "editDescription",
    request.description
  );


  setInputValue(
    "editCity",
    request.city
  );


  setInputValue(
    "editAddress",
    request.address
  );


  setInputValue(
    "editCustomer",
    request.customer
  );


  setInputValue(
    "editAmount",
    request.amount ||
    ""
  );


  setInputValue(

    "editPlannedDate",

    excelSerialToInputDate(
      request.plannedDate
    )

  );


  setInputValue(
    "editComment",
    request.comment
  );


  clearMessage(
    "editMessage"
  );


  openModal(
    "editModal"
  );

}


// ============================================================
// SAVE EDIT
// ============================================================

async function saveEdit() {

  const request =
    STATE.selectedRequest;


  if (
    !request
  ) {

    return;

  }


  setButtonBusy(

    "saveEditButton",

    true,

    "Зберігаємо..."

  );


  try {

    await enqueue({

      eventId:
        createId(
          "EVT"
        ),

      requestKey:
        "",

      requestId:
        request.id,

      operation:
        "UPDATE",

      actor:
        STATE.actor,

      source:
        STATE.employeeMode
          ? "Виконавець"
          : "Керівник",

      executor:
        request.executor,

      expectedStatus:
        "",

      newStatus:
        "",

      sd:
        valueOf(
          "editSD"
        ),

      category:
        valueOf(
          "editCategory"
        ),

      description:
        valueOf(
          "editDescription"
        ),

      city:
        valueOf(
          "editCity"
        ),

      address:
        valueOf(
          "editAddress"
        ),

      amount:
        numberOf(
          "editAmount"
        ),

      customer:
        valueOf(
          "editCustomer"
        ),

      plannedDate:
        valueOf(
          "editPlannedDate"
        ),

      comment:
        valueOf(
          "editComment"
        ),

      pauseReason:
        "",

      newExecutor:
        ""

    });


    showMessageSuccess(

      "editMessage",

      "Зміни передано в чергу."

    );

  }

  catch (error) {

    showMessageError(

      "editMessage",

      getErrorText(
        error
      )

    );

  }

  finally {

    setButtonBusy(

      "saveEditButton",

      false,

      "Зберегти"

    );

  }

}


// ============================================================
// QUEUE
// ============================================================

async function enqueue(
  event
) {

  await Excel.run(
    async context => {


      const queue =
        context.workbook.tables
          .getItem(
            "tbl_RBD_Queue"
          );


      queue.rows.add(
        null,
        [[

          event.eventId,
          event.requestKey,
          event.requestId,
          event.operation,

          localTimestamp(),

          event.actor,
          event.source,
          event.executor,
          event.expectedStatus,
          event.newStatus,

          event.sd,
          event.category,
          event.description,
          event.city,
          event.address,
          event.amount,
          event.customer,
          event.plannedDate,
          event.comment,
          event.pauseReason,
          event.newExecutor,

          "NEW",
          0,
          "",
          "",
          "",
          ""

        ]]
      );


      await context.sync();

    }
  );

}


// ============================================================
// ACTIVITY / NOTES / REMINDERS
// ============================================================

async function openActivityDrawer(
  request
) {

  STATE.activityRequestId =
    request.id;

  STATE.activityFilter =
    "ALL";


  setText(
    "activitySubtitle",
    request.id +
    " • " +
    request.description +
    " • " +
    request.city
  );


  document
    .querySelectorAll(
      "[data-activity-filter]"
    )
    .forEach(button => {
      button.classList.toggle(
        "active",
        button.dataset.activityFilter ===
          "ALL"
      );
    });


  const drawer =
    document.getElementById(
      "activityDrawer"
    );


  if (drawer) {
    drawer.classList.add("open");
    drawer.setAttribute(
      "aria-hidden",
      "false"
    );
  }


  clearActivityMessage();

  await loadActivity(
    request.id
  );

  startActivityAutoRefresh();

}


function closeActivityDrawer() {

  const drawer =
    document.getElementById(
      "activityDrawer"
    );


  if (drawer) {
    drawer.classList.remove("open");
    drawer.setAttribute(
      "aria-hidden",
      "true"
    );
  }


  stopActivityAutoRefresh();

}


function startActivityAutoRefresh() {

  stopActivityAutoRefresh();


  STATE.activityRefreshTimer =
    setInterval(
      () => {

        if (
          STATE.activityRequestId
        ) {

          loadActivity(
            STATE.activityRequestId,
            false
          );

        }

      },
      20000
    );

}


function stopActivityAutoRefresh() {

  if (
    STATE.activityRefreshTimer
  ) {

    clearInterval(
      STATE.activityRefreshTimer
    );

    STATE.activityRefreshTimer =
      null;

  }

}


async function loadActivity(
  requestId,
  showLoading = true
) {

  const body =
    document.getElementById(
      "activityBody"
    );


  if (
    showLoading &&
    body
  ) {

    body.innerHTML =
      '<div class="activity-empty">Завантаження історії...</div>';

  }


  try {

    const data =
      await Excel.run(
        async context => {

          const table =
            context.workbook.tables
              .getItemOrNullObject(
                "tbl_RBD_Activity"
              );


          table.load(
            "isNullObject"
          );


          await context.sync();


          if (
            table.isNullObject
          ) {

            return {
              exists: false,
              headers: [],
              rows: []
            };

          }


          const header =
            table.getHeaderRowRange();

          header.load("values");

          table.rows.load("items");

          await context.sync();


          if (
            table.rows.items.length ===
            0
          ) {

            return {
              exists: true,
              headers: header.values[0],
              rows: []
            };

          }


          const range =
            table.getDataBodyRange();

          range.load("values");

          await context.sync();


          return {
            exists: true,
            headers: header.values[0],
            rows: range.values
          };

        }
      );


    if (
      !data.exists
    ) {

      STATE.activities = [];

      if (body) {
        body.innerHTML = `
          <div class="activity-empty">
            Не знайдено таблицю <strong>tbl_RBD_Activity</strong>.<br><br>
            Створіть її один раз за структурою, яку я надам нижче.
          </div>
        `;
      }

      return;

    }


    STATE.activities =
      data.rows
        .map(row =>
          rowToActivity(
            data.headers,
            row
          )
        )
        .filter(activity =>
          normalizeName(
            activity.requestId
          ) ===
          normalizeName(
            requestId
          )
        )
        .sort(
          (a, b) =>
            dateToMilliseconds(
              b.eventDate
            ) -
            dateToMilliseconds(
              a.eventDate
            )
        );


    renderActivity();

  }

  catch (error) {

    console.error(
      "Помилка історії:",
      error
    );

    if (body) {
      body.innerHTML = `
        <div class="activity-empty">
          Помилка завантаження історії:<br>
          ${escapeHtml(
            getErrorText(error)
          )}
        </div>
      `;
    }

  }

}


function rowToActivity(
  headers,
  row
) {

  function get(name) {

    const index =
      headers.indexOf(name);

    if (index < 0) {
      return "";
    }

    return row[index] ?? "";

  }


  return {

    id:
      cleanText(
        get("Activity ID")
      ),

    requestId:
      cleanText(
        get("Request ID")
      ),

    eventDate:
      get("Дата події"),

    author:
      cleanText(
        get("Автор")
      ),

    type:
      cleanText(
        get("Тип")
      ).toUpperCase(),

    text:
      cleanText(
        get("Текст")
      ),

    reminderDate:
      get("Дата нагадування"),

    reminderStatus:
      cleanText(
        get("Статус нагадування")
      ).toUpperCase(),

    completedDate:
      get("Дата виконання")

  };

}


function renderActivity() {

  const body =
    document.getElementById(
      "activityBody"
    );


  if (!body) {
    return;
  }


  const source =
    STATE.activityFilter ===
      "ALL"
      ? STATE.activities
      : STATE.activities.filter(
          activity =>
            activity.type ===
            STATE.activityFilter
        );


  if (
    source.length === 0
  ) {

    body.innerHTML =
      '<div class="activity-empty">Історія за обраним типом поки порожня.</div>';

    return;

  }


  let html = "";
  let lastDate = "";


  source.forEach(activity => {

    const dateLabel =
      activityDateLabel(
        activity.eventDate
      );


    if (
      dateLabel !==
      lastDate
    ) {

      html += `
        <div class="activity-date-group">
          ${escapeHtml(dateLabel)}
        </div>
      `;

      lastDate =
        dateLabel;

    }


    const typeInfo =
      getActivityTypeInfo(
        activity.type
      );


    const reminderDate =
      activity.reminderDate
        ? formatActivityDateTime(
            activity.reminderDate
          )
        : "";


    const isDone =
      activity.reminderStatus ===
        "DONE";


    html += `
      <div class="activity-item">

        <div class="activity-time">
          ${escapeHtml(
            activityTimeLabel(
              activity.eventDate
            )
          )}
        </div>

        <div class="activity-type-icon ${typeInfo.className}">
          ${typeInfo.icon}
        </div>

        <div>

          <div class="activity-item-title ${typeInfo.className}">
            ${escapeHtml(
              typeInfo.title
            )}
          </div>

          <div class="activity-item-text">
            ${escapeHtml(
              activity.text || "—"
            )}
          </div>

          ${
            activity.type === "REMINDER"
              ? `
                <div class="activity-reminder-line">
                  ${
                    reminderDate
                      ? `<span class="activity-pill">⏰ ${escapeHtml(reminderDate)}</span>`
                      : ""
                  }

                  <span class="activity-pill ${isDone ? "done" : "active"}">
                    ${isDone ? "Виконано" : "Активне"}
                  </span>

                  ${
                    !isDone
                      ? `
                        <button
                          class="activity-done-button"
                          type="button"
                          data-reminder-done="${escapeHtml(activity.id)}"
                        >
                          ✓ Виконано
                        </button>
                      `
                      : ""
                  }
                </div>
              `
              : ""
          }

          <div class="activity-item-meta">
            ${escapeHtml(
              activity.author || "Система"
            )}
          </div>

        </div>

      </div>
    `;

  });


  body.innerHTML =
    html;


  body
    .querySelectorAll(
      "[data-reminder-done]"
    )
    .forEach(button => {

      button.addEventListener(
        "click",
        async () => {

          await markReminderDone(
            button.dataset.reminderDone
          );

        }
      );

    });

}


function getActivityTypeInfo(
  type
) {

  switch (type) {

    case "REMINDER":
      return {
        title: "Нагадування",
        icon: "⏰",
        className: "reminder"
      };

    case "STATUS":
      return {
        title: "Зміна статусу",
        icon: "↻",
        className: "status"
      };

    default:
      return {
        title: "Примітка",
        icon: "▤",
        className: ""
      };

  }

}


async function addNoteFromDrawer() {

  const text =
    valueOf(
      "activityText"
    );


  if (!text) {

    showActivityMessage(
      "Введіть текст примітки.",
      true
    );

    return;

  }


  await addActivityRecord({
    type: "NOTE",
    text,
    reminderDate: "",
    reminderStatus: ""
  });

}


async function addReminderFromDrawer() {

  const text =
    valueOf(
      "activityText"
    );

  const reminderDate =
    valueOf(
      "activityReminderAt"
    );


  if (!text) {

    showActivityMessage(
      "Введіть текст нагадування.",
      true
    );

    return;

  }


  if (!reminderDate) {

    showActivityMessage(
      "Вкажіть дату та час нагадування.",
      true
    );

    return;

  }


  await addActivityRecord({
    type: "REMINDER",
    text,
    reminderDate,
    reminderStatus: "NEW"
  });

}


async function addActivityRecord({
  type,
  text,
  reminderDate,
  reminderStatus
}) {

  if (
    !STATE.activityRequestId
  ) {

    showActivityMessage(
      "Не визначено заявку.",
      true
    );

    return;

  }


  try {

    await appendActivity(
      STATE.activityRequestId,
      type,
      text,
      reminderDate,
      reminderStatus
    );


    setInputValue(
      "activityText",
      ""
    );

    if (
      type === "REMINDER"
    ) {
      setInputValue(
        "activityReminderAt",
        ""
      );
    }


    showActivityMessage(
      type === "REMINDER"
        ? "Нагадування додано."
        : "Примітку додано.",
      false
    );


    await loadActivity(
      STATE.activityRequestId,
      false
    );

  }

  catch (error) {

    showActivityMessage(
      getErrorText(error),
      true
    );

  }

}


async function appendActivity(
  requestId,
  type,
  text,
  reminderDate = "",
  reminderStatus = ""
) {

  await Excel.run(
    async context => {

      const table =
        context.workbook.tables
          .getItemOrNullObject(
            "tbl_RBD_Activity"
          );


      table.load(
        "isNullObject"
      );

      await context.sync();


      if (
        table.isNullObject
      ) {

        throw new Error(
          "Не знайдено таблицю tbl_RBD_Activity."
        );

      }


      table.rows.add(
        null,
        [[
          createId("ACT"),
          requestId,
          localTimestamp(),
          STATE.actor,
          type,
          text,
          reminderDate,
          reminderStatus,
          ""
        ]]
      );


      await context.sync();

    }
  );

}


async function appendActivitySafe(
  requestId,
  type,
  text
) {

  try {

    await appendActivity(
      requestId,
      type,
      text,
      "",
      ""
    );


    if (
      STATE.activityRequestId ===
      requestId
    ) {

      await loadActivity(
        requestId,
        false
      );

    }

  }

  catch (error) {

    console.warn(
      "Історію не записано:",
      error
    );

  }

}


async function markReminderDone(
  activityId
) {

  try {

    await Excel.run(
      async context => {

        const table =
          context.workbook.tables
            .getItemOrNullObject(
              "tbl_RBD_Activity"
            );


        table.load(
          "isNullObject"
        );

        await context.sync();


        if (
          table.isNullObject
        ) {
          throw new Error(
            "Не знайдено таблицю tbl_RBD_Activity."
          );
        }


        const body =
          table.getDataBodyRange();

        body.load("values");

        await context.sync();


        const index =
          body.values.findIndex(
            row =>
              cleanText(row[0]) ===
              cleanText(activityId)
          );


        if (
          index < 0
        ) {
          throw new Error(
            "Нагадування не знайдено."
          );
        }


        body.getCell(
          index,
          7
        ).values = [["DONE"]];

        body.getCell(
          index,
          8
        ).values = [[localTimestamp()]];


        await context.sync();

      }
    );


    await loadActivity(
      STATE.activityRequestId,
      false
    );

  }

  catch (error) {

    showActivityMessage(
      getErrorText(error),
      true
    );

  }

}


function buildStatusActivityText(
  previousStatus,
  newStatus,
  comment
) {

  let text =
    (previousStatus || "—") +
    " → " +
    (newStatus || "—");


  if (comment) {
    text +=
      "\nКоментар: " +
      comment;
  }


  return text;

}


function activityDateLabel(
  value
) {

  const ms =
    dateToMilliseconds(value);


  if (!ms) {
    return "Без дати";
  }


  const date =
    new Date(ms);

  const now =
    new Date();

  const today =
    new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate()
    ).getTime();

  const target =
    new Date(
      date.getFullYear(),
      date.getMonth(),
      date.getDate()
    ).getTime();


  if (
    target === today
  ) {
    return "Сьогодні, " +
      formatExcelDate(value);
  }


  if (
    target ===
    today - 86400000
  ) {
    return "Вчора, " +
      formatExcelDate(value);
  }


  return formatExcelDate(value);

}


function activityTimeLabel(
  value
) {

  const ms =
    dateToMilliseconds(value);


  if (!ms) {
    return "";
  }


  const date =
    new Date(ms);


  return (
    pad2(
      date.getHours()
    ) +
    ":" +
    pad2(
      date.getMinutes()
    )
  );

}


function formatActivityDateTime(
  value
) {

  const ms =
    dateToMilliseconds(value);


  if (!ms) {
    return "—";
  }


  const date =
    new Date(ms);


  return (
    pad2(date.getDate()) +
    "." +
    pad2(date.getMonth() + 1) +
    "." +
    date.getFullYear() +
    " " +
    pad2(date.getHours()) +
    ":" +
    pad2(date.getMinutes())
  );

}


function showActivityMessage(
  message,
  isError
) {

  const element =
    document.getElementById(
      "activityMessage"
    );


  if (!element) {
    return;
  }


  element.className =
    "activity-message " +
    (isError
      ? "error"
      : "success");

  element.textContent =
    message;

}


function clearActivityMessage() {

  const element =
    document.getElementById(
      "activityMessage"
    );


  if (!element) {
    return;
  }


  element.className =
    "activity-message";

  element.textContent =
    "";

}


// ============================================================
// COLORS
// ============================================================

function getRowClass(
  request
) {

  if (
    isOverdue(
      request
    )
  ) {

    return "row-overdue";

  }


  switch (
    request.status
  ) {

    case "Нова":
      return "row-new";

    case "Прийнята в роботу":
      return "row-accepted";

    case "Пошук підрядника":
      return "row-contractor";

    case "Погодження кошторису":
      return "row-estimate";

    case "Укладання договору":
      return "row-contract";

    case "Погодження бюджету":
      return "row-budget";

    case "Виконання робіт":
      return "row-execution";

    case "Прийняття робіт":
      return "row-acceptance";

    case "Призупинена":
      return "row-paused";

    case "Закрита":
      return "row-closed";

    default:
      return "row-new";

  }

}


// ============================================================

function getStatusClass(
  status
) {

  switch (
    status
  ) {

    case "Нова":
      return "status-new";

    case "Прийнята в роботу":
      return "status-accepted";

    case "Пошук підрядника":
      return "status-contractor";

    case "Погодження кошторису":
      return "status-estimate";

    case "Укладання договору":
      return "status-contract";

    case "Погодження бюджету":
      return "status-budget";

    case "Виконання робіт":
      return "status-execution";

    case "Прийняття робіт":
      return "status-acceptance";

    case "Призупинена":
      return "status-paused";

    case "Закрита":
      return "status-closed";

    default:
      return "status-new";

  }

}


// ============================================================

function getStatusButtonClass(
  status
) {

  switch (
    status
  ) {

    case "Нова":
      return "btn-new";

    case "Прийнята в роботу":
      return "btn-accepted";

    case "Пошук підрядника":
      return "btn-contractor";

    case "Погодження кошторису":
      return "btn-estimate";

    case "Укладання договору":
      return "btn-contract";

    case "Погодження бюджету":
      return "btn-budget";

    case "Виконання робіт":
      return "btn-execution";

    case "Прийняття робіт":
      return "btn-acceptance";

    case "Призупинена":
      return "btn-paused";

    case "Закрита":
      return "btn-closed";

    default:
      return "btn-new";

  }

}


// ============================================================
// TIME IN STATUS
// ============================================================

function currentStatusHours(
  request
) {

  if (
    request.status ===
      "Призупинена" ||

    request.status ===
      "Закрита"
  ) {

    return null;

  }


  const start =
    dateToMilliseconds(
      request.statusSince
    );


  if (
    !start
  ) {

    return (
      request.storedTimeInStatus ||
      null
    );

  }


  const hours =
    (
      Date.now() -
      start
    ) /
    3600000;


  if (
    hours < 0
  ) {

    return 0;

  }


  return Math.round(
    hours * 10
  ) / 10;

}


// ============================================================
// FIND REQUEST
// ============================================================

function findRequest(
  requestId
) {

  return STATE.requests.find(
    request =>
      request.id ===
      requestId
  );

}


// ============================================================
// SELECT
// ============================================================

function fillSelect(
  id,
  values,
  placeholder
) {

  const select =
    document.getElementById(
      id
    );


  if (
    !select
  ) {

    return;

  }


  const oldValue =
    select.value;


  select.innerHTML =
    "";


  const first =
    document.createElement(
      "option"
    );


  first.value =
    "";


  first.textContent =
    placeholder;


  select.appendChild(
    first
  );


  values.forEach(
    value => {


      const option =
        document.createElement(
          "option"
        );


      option.value =
        value;


      option.textContent =
        value;


      select.appendChild(
        option
      );

    }
  );


  if (
    values.includes(
      oldValue
    )
  ) {

    select.value =
      oldValue;

  }

}


// ============================================================
// MODAL
// ============================================================

function openModal(
  id
) {

  document
    .getElementById(id)
    ?.classList.add(
      "open"
    );

}


function closeModal(
  id
) {

  document
    .getElementById(id)
    ?.classList.remove(
      "open"
    );

}


// ============================================================
// CLEAR CREATE
// ============================================================

function clearCreateForm() {

  [

    "createSD",
    "createCategory",
    "createDescription",
    "createCity",
    "createAddress",
    "createCustomer",
    "createExecutor",
    "createAmount",
    "createPlannedDate",
    "createComment"

  ]
    .forEach(
      id => {


        const element =
          document.getElementById(
            id
          );


        if (
          element
        ) {

          element.value =
            "";

        }

      }
    );


  clearMessage(
    "createMessage"
  );

  refreshRequiredStates();

}


// ============================================================
// LOADING
// ============================================================

function showLoading() {

  const body =
    document.getElementById(
      "requestsBody"
    );


  if (!body) {

    return;

  }


  body.innerHTML = `

    <tr>

      <td colspan="13">

        <div class="empty-state">

          Завантаження заявок...

        </div>

      </td>

    </tr>

  `;

}


// ============================================================
// ERROR
// ============================================================

function showTableError(
  message
) {

  const body =
    document.getElementById(
      "requestsBody"
    );


  if (!body) {

    return;

  }


  body.innerHTML = `

    <tr>

      <td colspan="13">

        <div class="empty-state">

          Помилка завантаження:
          ${escapeHtml(
            message
          )}

        </div>

      </td>

    </tr>

  `;

}


// ============================================================
// UPDATED
// ============================================================

function updateLastUpdated() {

  const now =
    new Date();


  setText(

    "lastUpdated",

    "↻ Дані оновлено: " +

    pad2(
      now.getDate()
    ) +

    "." +

    pad2(
      now.getMonth() + 1
    ) +

    "." +

    now.getFullYear() +

    " " +

    pad2(
      now.getHours()
    ) +

    ":" +

    pad2(
      now.getMinutes()
    )

  );

}


// ============================================================
// DATE
// ============================================================

function formatExcelDate(
  value
) {

  if (
    value === "" ||
    value === null ||
    value === undefined
  ) {

    return "—";

  }


  const text =
    String(
      value
    ).trim();


  // Excel Online інколи повертає серійну дату як число,
  // а інколи як текстове число, наприклад "46290".
  // Обидва варіанти треба трактувати як Excel serial date.
  if (
    typeof value === "number" ||
    /^\d{4,6}(?:[.,]\d+)?$/.test(text)
  ) {

    const serial =
      Number(
        text.replace(",", ".")
      );


    if (
      Number.isFinite(serial) &&
      serial > 20000 &&
      serial < 100000
    ) {

      const date =
        excelSerialToDate(
          serial
        );


      return (

        pad2(
          date.getUTCDate()
        ) +

        "." +

        pad2(
          date.getUTCMonth() + 1
        ) +

        "." +

        date.getUTCFullYear()

      );

    }

  }


  const isoMatch =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})/
    );


  if (
    isoMatch
  ) {

    return (
      isoMatch[3] +
      "." +
      isoMatch[2] +
      "." +
      isoMatch[1]
    );

  }


  const uaMatch =
    text.match(
      /^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/
    );


  if (
    uaMatch
  ) {

    return (
      pad2(uaMatch[1]) +
      "." +
      pad2(uaMatch[2]) +
      "." +
      uaMatch[3]
    );

  }


  return text;

}


// ============================================================

function formatExcelDateTime(
  value
) {

  if (
    typeof value !==
    "number"
  ) {

    if (
      !value
    ) {

      return "—";

    }


    const text =
      String(
        value
      );


    const match =
      text.match(
        /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/
      );


    if (
      match
    ) {

      return (

        match[3] +
        "." +

        match[2] +
        "." +

        match[1] +
        " " +

        match[4] +
        ":" +

        match[5]

      );

    }


    return text;

  }


  const date =
    excelSerialToDate(
      value
    );


  return (

    pad2(
      date.getUTCDate()
    ) +

    "." +

    pad2(
      date.getUTCMonth() + 1
    ) +

    "." +

    date.getUTCFullYear() +

    " " +

    pad2(
      date.getUTCHours()
    ) +

    ":" +

    pad2(
      date.getUTCMinutes()
    )

  );

}


// ============================================================
// EXCEL SERIAL
// ============================================================

function excelSerialToDate(
  serial
) {

  return new Date(

    Math.round(
      (
        Number(
          serial
        ) -
        25569
      ) *
      86400000
    )

  );

}


// ============================================================

function excelSerialToInputDate(
  value
) {

  if (
    value === "" ||
    value === null ||
    value === undefined
  ) {

    return "";

  }


  const text =
    String(value).trim();


  const isoMatch =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})/
    );


  if (isoMatch) {

    return (
      isoMatch[1] +
      "-" +
      isoMatch[2] +
      "-" +
      isoMatch[3]
    );

  }


  const serial =
    Number(
      text.replace(",", ".")
    );


  if (
    !Number.isFinite(serial) ||
    serial <= 20000 ||
    serial >= 100000
  ) {

    return "";

  }


  const date =
    excelSerialToDate(
      serial
    );


  return (

    date.getUTCFullYear() +

    "-" +

    pad2(
      date.getUTCMonth() + 1
    ) +

    "-" +

    pad2(
      date.getUTCDate()
    )

  );

}


// ============================================================
// DATE TO MILLISECONDS
// ============================================================

function dateToMilliseconds(
  value
) {

  if (
    typeof value ===
    "number"
  ) {

    return (
      value -
      25569
    ) *
    86400000;

  }


  if (
    !value
  ) {

    return 0;

  }


  const text =
    String(
      value
    ).trim();


  // Серійна Excel-дата може прийти як текст, наприклад "46290".
  if (
    /^\d{4,6}(?:[.,]\d+)?$/.test(text)
  ) {

    const serial =
      Number(
        text.replace(",", ".")
      );


    if (
      Number.isFinite(serial) &&
      serial > 20000 &&
      serial < 100000
    ) {

      return (
        serial -
        25569
      ) *
      86400000;

    }

  }


  const localMatch =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/
    );


  if (
    localMatch
  ) {

    return new Date(

      Number(
        localMatch[1]
      ),

      Number(
        localMatch[2]
      ) - 1,

      Number(
        localMatch[3]
      ),

      Number(
        localMatch[4]
      ),

      Number(
        localMatch[5]
      ),

      Number(
        localMatch[6]
      )

    ).getTime();

  }


  const dateOnlyMatch =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})$/
    );


  if (
    dateOnlyMatch
  ) {

    return new Date(

      Number(
        dateOnlyMatch[1]
      ),

      Number(
        dateOnlyMatch[2]
      ) - 1,

      Number(
        dateOnlyMatch[3]
      )

    ).getTime();

  }


  const parsed =
    Date.parse(
      text
    );


  return Number.isFinite(
    parsed
  )
    ? parsed
    : 0;

}


// ============================================================
// TODAY
// ============================================================

function startOfToday() {

  const now =
    new Date();


  return new Date(

    now.getFullYear(),

    now.getMonth(),

    now.getDate()

  ).getTime();

}


// ============================================================
// MONEY
// ============================================================

function money(
  value
) {

  const number =
    toNumber(
      value
    );


  return (

    number.toLocaleString(
      "uk-UA",
      {

        minimumFractionDigits:
          0,

        maximumFractionDigits:
          2

      }
    ) +

    " грн"

  );

}


// ============================================================

function moneyCompact(
  value
) {

  const number =
    toNumber(
      value
    );


  if (
    number === 0
  ) {

    return "—";

  }


  return number.toLocaleString(
    "uk-UA",
    {

      minimumFractionDigits:
        0,

      maximumFractionDigits:
        2

    }
  );

}


// ============================================================
// NUMBER
// ============================================================

function toNumber(
  value
) {

  if (
    typeof value ===
    "number"
  ) {

    return Number.isFinite(
      value
    )
      ? value
      : 0;

  }


  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return 0;

  }


  const cleaned =
    String(
      value
    )
      .replace(
        /\s/g,
        ""
      )
      .replace(
        /\u00A0/g,
        ""
      )
      .replace(
        ",",
        "."
      );


  const result =
    Number(
      cleaned
    );


  return Number.isFinite(
    result
  )
    ? result
    : 0;

}


// ============================================================
// NAME NORMALIZATION
// ============================================================

function normalizeName(
  value
) {

  return String(
    value ?? ""
  )
    .replace(
      /\s+/g,
      " "
    )
    .trim()
    .toLowerCase();

}


// ============================================================
// VALUES
// ============================================================

function valueOf(
  id
) {

  return String(

    document
      .getElementById(
        id
      )
      ?.value ?? ""

  ).trim();

}


function numberOf(
  id
) {

  return toNumber(
    valueOf(
      id
    )
  );

}


function cleanText(
  value
) {

  return String(
    value ?? ""
  ).trim();

}


// ============================================================
// SET INPUT
// ============================================================

function setInputValue(
  id,
  value
) {

  const element =
    document.getElementById(
      id
    );


  if (
    element
  ) {

    element.value =
      value ?? "";

  }

}


// ============================================================
// ID
// ============================================================

function createId(
  prefix
) {

  if (
    window.crypto &&
    typeof crypto.randomUUID ===
      "function"
  ) {

    return (

      prefix +
      "-" +
      crypto.randomUUID()

    );

  }


  return (

    prefix +
    "-" +
    Date.now() +
    "-" +
    Math.random()
      .toString(16)
      .slice(2)

  );

}


// ============================================================
// TIMESTAMP
// ============================================================

function localTimestamp() {

  const date =
    new Date();


  return (

    date.getFullYear() +

    "-" +

    pad2(
      date.getMonth() + 1
    ) +

    "-" +

    pad2(
      date.getDate()
    ) +

    "T" +

    pad2(
      date.getHours()
    ) +

    ":" +

    pad2(
      date.getMinutes()
    ) +

    ":" +

    pad2(
      date.getSeconds()
    )

  );

}


// ============================================================
// UI HELPERS
// ============================================================

function setText(
  id,
  value
) {

  const element =
    document.getElementById(
      id
    );


  if (
    element
  ) {

    element.textContent =
      String(
        value
      );

  }

}


// ============================================================

function clearMessage(
  id
) {

  const element =
    document.getElementById(
      id
    );


  if (
    !element
  ) {

    return;

  }


  element.className =
    "message";


  element.textContent =
    "";

}


// ============================================================

function showMessageSuccess(
  id,
  message
) {

  const element =
    document.getElementById(
      id
    );


  if (
    !element
  ) {

    return;

  }


  element.className =
    "message success";


  element.textContent =
    "✓ " +
    message;

}


// ============================================================

function showMessageError(
  id,
  message
) {

  const element =
    document.getElementById(
      id
    );


  if (
    !element
  ) {

    return;

  }


  element.className =
    "message error";


  element.textContent =
    "Помилка: " +
    message;

}


// ============================================================

function setButtonBusy(
  id,
  busy,
  caption
) {

  const button =
    document.getElementById(
      id
    );


  if (
    !button
  ) {

    return;

  }


  button.disabled =
    busy;


  button.textContent =
    caption;

}


// ============================================================
// INITIALS
// ============================================================

function getInitials(
  name
) {

  const parts =
    String(
      name ||
      ""
    )
      .replace(
        /\./g,
        ""
      )
      .trim()
      .split(
        /\s+/
      );


  if (
    !parts.length
  ) {

    return "Р";

  }


  let result =
    parts[0]
      .charAt(0)
      .toUpperCase();


  if (
    parts.length > 1
  ) {

    result +=
      parts[1]
        .charAt(0)
        .toUpperCase();

  }


  return result;

}


// ============================================================
// DOM SAFE
// ============================================================

function domSafe(
  value
) {

  return String(
    value
  ).replace(
    /[^a-zA-Z0-9_-]/g,
    "_"
  );

}


// ============================================================
// PAD
// ============================================================

function pad2(
  value
) {

  return String(
    value
  ).padStart(
    2,
    "0"
  );

}


// ============================================================
// ERROR
// ============================================================

function getErrorText(
  error
) {

  if (
    error &&
    typeof error.message ===
      "string"
  ) {

    return error.message;

  }


  return String(
    error
  );

}


// ============================================================
// ESCAPE HTML
// ============================================================

function escapeHtml(
  value
) {

  return String(
    value ?? ""
  )

    .replace(
      /&/g,
      "&amp;"
    )

    .replace(
      /</g,
      "&lt;"
    )

    .replace(
      />/g,
      "&gt;"
    )

    .replace(
      /"/g,
      "&quot;"
    )

    .replace(
      /'/g,
      "&#039;"
    );

}
