// ============================================================
// RBD — TASKPANE V3
// Керівник: Іванченко В.М.
// ============================================================


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

  expandedStatusId: ""

};


// ============================================================
// КЕРІВНИК
// ============================================================

const MANAGER_NAME =
  "Іванченко В.М.";


// На цих вкладках працює повний режим керівника.
// Іванченко бачить ВСІ заявки.
const MANAGER_SHEETS = [

  "Іванченко В.М.",
  "Аркуш1",
  "ЗВЕДЕНА",
  "КАБІНЕТ_КЕРІВНИКА"

];


// ============================================================
// ПЕРСОНАЛЬНІ ВКЛАДКИ ВИКОНАВЦІВ
// ============================================================

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
// START
// ============================================================

Office.onReady(async info => {

  initEvents();


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

}


// ============================================================
// ВИЗНАЧАЄМО ПОТОЧНУ ВКЛАДКУ
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
        sheet.name;

    }
  );


  // ==========================================================
  // КЕРІВНИК
  // ==========================================================

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


  // ==========================================================
  // ВИКОНАВЕЦЬ
  // ==========================================================

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


  // ==========================================================
  // ІНШІ ВКЛАДКИ
  // ==========================================================

  STATE.employeeMode =
    false;


  STATE.employee =
    "";


  STATE.actor =
    MANAGER_NAME;

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


  // ==========================================================
  // ПЕРСОНАЛЬНИЙ КАБІНЕТ
  // ==========================================================

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


    // Виконавець не бачить фільтр
    // по інших виконавцях.
    if (
      executorFilter
    ) {

      executorFilter.style.display =
        "none";

    }


    // Нова заявка автоматично
    // створюється на нього.
    if (
      createExecutor
    ) {

      createExecutor.disabled =
        true;


      createExecutor.value =
        STATE.employee;

    }


    return;

  }


  // ==========================================================
  // КАБІНЕТ КЕРІВНИКА
  // ==========================================================

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


  // Керівник бачить
  // фільтр по виконавцях.
  if (
    executorFilter
  ) {

    executorFilter.style.display =
      "";

  }


  // Керівник може створити
  // заявку на будь-кого.
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
          await readSimpleTable(
            context,
            "tbl_RBD_Categories"
          );


        const cities =
          await readSimpleTable(
            context,
            "tbl_RBD_Cities"
          );


        const employees =
          await readSimpleTable(
            context,
            "tbl_RBD_Employees"
          );


        const statuses =
          await readSimpleTable(
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
    data.categories;


  STATE.cities =
    data.cities;


  STATE.employees =
    data.employees;


  STATE.statuses =
    data.statuses;


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
// ПРОСТИЙ ДОВІДНИК
// ============================================================

async function readSimpleTable(
  context,
  tableName
) {

  const table =
    context.workbook
      .tables
      .getItem(
        tableName
      );


  table.rows.load(
    "items"
  );


  await context.sync();


  if (
    table.rows.items.length === 0
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
  tableName
) {

  const table =
    context.workbook
      .tables
      .getItem(
        tableName
      );


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
    table.rows.items.length === 0
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
// ЗАВАНТАЖЕННЯ КАБІНЕТУ
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
              "tbl_RBD_Base"
            );


          const archive =
            await readTable(
              context,
              "tbl_RBD_Archive"
            );


          return {

            base,

            archive

          };

        }
      );


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


    // ========================================================
    // КЕРІВНИК — ВСІ ЗАЯВКИ
    // ВИКОНАВЕЦЬ — ЛИШЕ СВОЇ
    // ========================================================

    if (
      STATE.employeeMode
    ) {

      STATE.requests =
        STATE.allRequests.filter(
          request =>
            request.executor ===
            STATE.employee
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
      getErrorText(error)
    );

  }

}


// ============================================================
// РЯДОК EXCEL → ОБ'ЄКТ
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
          request.executor !==
          executor
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
          request.executor ===
          STATE.employee
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


        if (!date) {

          return false;

        }


        return (
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
// ЗАГАЛЬНИЙ ПІДСУМОК
// ============================================================

function updateSummary() {

  const totalSum =
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
      totalSum
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
// СУМА ЗАЯВОК
// ============================================================

function sumRequests(
  requests
) {

  return requests.reduce(
    (
      total,
      request
    ) => {

      return (
        total +
        toNumber(
          request.amount
        )
      );

    },
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
    request.slaOverdue ===
      "Так" ||

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
      startOfToday() &&
    request.status !==
      "Закрита"
  ) {

    return true;

  }


  return false;

}


// ============================================================
// ВІДОБРАЖЕННЯ ЗАЯВОК
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

    <tr
      class="
        request-row
        ${rowClass}
      "
    >


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

          <span
            class="status-dot"
          ></span>

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
          isOverdue(request)
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
            class="
              action-button
              edit
            "
            data-edit="${escapeHtml(
              request.id
            )}"
            type="button"
            title="Редагувати"
          >

            ✎

          </button>


        </div>

      </td>


    </tr>

  `;

}


// ============================================================
// ПАНЕЛЬ ЗМІНИ СТАТУСУ
// ============================================================

function buildStatusPanel(
  request
) {

  const buttons =
    STATE.statuses
      .map(
        status => {


          const isCurrent =
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
                  isCurrent
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
                isCurrent
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
// ДІЇ У РЯДКУ
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


          if (
            STATE.expandedStatusId ===
            requestId
          ) {

            STATE.expandedStatusId =
              "";

          }

          else {

            STATE.expandedStatusId =
              requestId;

          }


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


          if (request) {

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


          if (request) {

            openEdit(
              request
            );

          }

        }
      );

    });

}


// ============================================================
// КНОПКИ СТАТУСІВ
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


          const requestId =
            button.dataset.requestId;


          const newStatus =
            button.dataset.inlineStatus;


          const request =
            findRequest(
              requestId
            );


          if (!request) {

            return;

          }


          // Призупинення
          if (
            newStatus ===
            "Призупинена"
          ) {

            showPauseBox(
              request
            );

            return;

          }


          // Закриття
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

  const safeId =
    domSafe(
      request.id
    );


  const area =
    document.getElementById(
      "specialStatusArea_" +
      safeId
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


        if (!reason) {

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

  const safeId =
    domSafe(
      request.id
    );


  const area =
    document.getElementById(
      "specialStatusArea_" +
      safeId
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
        "STATUS_CHANGE",

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


    showInlineSuccess(
      request.id,

      "✓ " +
      request.status +
      " → " +
      newStatus +
      ". Передано в чергу."
    );

  }

  catch (error) {

    showInlineError(
      request.id,
      getErrorText(
        error
      )
    );

  }

}


// ============================================================
// INLINE MESSAGE
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
// НОВА ЗАЯВКА
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

    if (executor) {

      executor.disabled =
        false;


      executor.value =
        STATE.employee;


      executor.disabled =
        true;

    }

  }

  else {

    if (executor) {

      executor.disabled =
        false;


      executor.value =
        "";

    }

  }


  openModal(
    "createModal"
  );

}


// ============================================================
// CREATE
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


    showMessageSuccess(
      "createMessage",
      "Заявку передано в чергу."
    );

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
// ДЕТАЛІ
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
      request.slaOverdue ||
      "Ні"
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


  document
    .getElementById(
      "detailsContent"
    )
    .innerHTML =
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


  openModal(
    "detailsModal"
  );

}


// ============================================================
// РЕДАГУВАННЯ
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


  if (!request) {

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
// ЗАПИС У ЧЕРГУ
// ============================================================

async function enqueue(
  event
) {

  await Excel.run(
    async context => {


      const queue =
        context.workbook.tables.getItem(
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
// КОЛІР РЯДКА
// ============================================================

function getRowClass(
  request
) {

  // Прострочення має пріоритет.
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
// КОЛІР СТАТУСУ
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
// КОЛІР КНОПКИ СТАТУСУ
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
// ЧАС У ПОТОЧНОМУ СТАТУСІ
// ============================================================

function currentStatusHours(
  request
) {

  // У призупиненому статусі
  // SLA не рахуємо.
  if (
    request.status ===
    "Призупинена"
  ) {

    return null;

  }


  const start =
    dateToMilliseconds(
      request.statusSince
    );


  if (!start) {

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
// ЗНАЙТИ ЗАЯВКУ
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


  if (!select) {

    return;

  }


  const oldValue =
    select.value;


  select.innerHTML =
    "";


  const empty =
    document.createElement(
      "option"
    );


  empty.value =
    "";


  empty.textContent =
    placeholder;


  select.appendChild(
    empty
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
    .getElementById(
      id
    )
    ?.classList.add(
      "open"
    );

}


// ============================================================

function closeModal(
  id
) {

  document
    .getElementById(
      id
    )
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


        if (element) {

          element.value =
            "";

        }

      }
    );


  clearMessage(
    "createMessage"
  );

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
// ERROR TABLE
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
// DATE FORMAT
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


  if (
    typeof value ===
    "number"
  ) {

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

      date.getUTCFullYear()

    );

  }


  const text =
    String(
      value
    );


  const iso =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})/
    );


  if (iso) {

    return (
      iso[3] +
      "." +
      iso[2] +
      "." +
      iso[1]
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

    return formatExcelDate(
      value
    );

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
// EXCEL SERIAL → DATE
// ============================================================

function excelSerialToDate(
  serial
) {

  return new Date(

    Math.round(
      (
        Number(serial) -
        25569
      ) *
      86400000
    )

  );

}


// ============================================================
// EXCEL SERIAL → INPUT DATE
// ============================================================

function excelSerialToInputDate(
  value
) {

  if (
    typeof value !==
    "number"
  ) {

    return "";

  }


  const date =
    excelSerialToDate(
      value
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
// DATE → MILLISECONDS
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


  if (!value) {

    return 0;

  }


  const parsed =
    Date.parse(
      String(value)
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
// MONEY COMPACT
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
// SAFE NUMBER
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
    String(value)
      .replace(
        /\s/g,
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


// ============================================================

function numberOf(
  id
) {

  return toNumber(
    valueOf(
      id
    )
  );

}


// ============================================================

function cleanText(
  value
) {

  return String(
    value ?? ""
  ).trim();

}


// ============================================================
// INPUT VALUE
// ============================================================

function setInputValue(
  id,
  value
) {

  const element =
    document.getElementById(
      id
    );


  if (element) {

    element.value =
      value ?? "";

  }

}


// ============================================================
// UUID
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
// LOCAL TIMESTAMP
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
// SET TEXT
// ============================================================

function setText(
  id,
  value
) {

  const element =
    document.getElementById(
      id
    );


  if (element) {

    element.textContent =
      String(
        value
      );

  }

}


// ============================================================
// MESSAGE
// ============================================================

function clearMessage(
  id
) {

  const element =
    document.getElementById(
      id
    );


  if (!element) {

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


  if (!element) {

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


  if (!element) {

    return;

  }


  element.className =
    "message error";


  element.textContent =
    "Помилка: " +
    message;

}


// ============================================================
// BUTTON
// ============================================================

function setButtonBusy(
  id,
  busy,
  text
) {

  const button =
    document.getElementById(
      id
    );


  if (!button) {

    return;

  }


  button.disabled =
    busy;


  button.textContent =
    text;

}


// ============================================================
// INITIALS
// ============================================================

function getInitials(
  name
) {

  const parts =
    String(
      name || ""
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
    parts.length === 0
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
// PAD 2
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
// ERROR TEXT
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
// HTML ESCAPE
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
