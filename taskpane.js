const STATE = {

  requests: [],

  archive: [],

  filtered: [],

  categories: [],

  cities: [],

  employees: [],

  statuses: [],

  selectedRequest: null

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

    await loadDictionaries();

    await loadDashboard();

  }

  catch (error) {

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
      loadDashboard
    );


  document
    .getElementById(
      "newRequestButton"
    )
    ?.addEventListener(
      "click",
      openCreateModal
    );


  document
    .getElementById(
      "filterStatus"
    )
    ?.addEventListener(
      "change",
      applyFilters
    );


  document
    .getElementById(
      "filterExecutor"
    )
    ?.addEventListener(
      "change",
      applyFilters
    );


  document
    .getElementById(
      "filterCity"
    )
    ?.addEventListener(
      "change",
      applyFilters
    );


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
      "saveStatusButton"
    )
    ?.addEventListener(
      "click",
      saveStatusChange
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
    .getElementById(
      "statusNew"
    )
    ?.addEventListener(
      "change",
      handlePauseReason
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
// DICTIONARIES
// ============================================================

async function loadDictionaries() {

  await Excel.run(async context => {


    const categoryTable =
      context.workbook.tables.getItem(
        "tbl_RBD_Categories"
      );


    const cityTable =
      context.workbook.tables.getItem(
        "tbl_RBD_Cities"
      );


    const employeeTable =
      context.workbook.tables.getItem(
        "tbl_RBD_Employees"
      );


    const slaTable =
      context.workbook.tables.getItem(
        "tbl_SLA"
      );


    const categoryRange =
      categoryTable.getDataBodyRange();


    const cityRange =
      cityTable.getDataBodyRange();


    const employeeRange =
      employeeTable.getDataBodyRange();


    const slaRange =
      slaTable.getDataBodyRange();


    categoryRange.load(
      "values"
    );


    cityRange.load(
      "values"
    );


    employeeRange.load(
      "values"
    );


    slaRange.load(
      "values"
    );


    await context.sync();


    STATE.categories =
      firstColumn(
        categoryRange.values
      );


    STATE.cities =
      firstColumn(
        cityRange.values
      );


    STATE.employees =
      firstColumn(
        employeeRange.values
      );


    STATE.statuses =
      firstColumn(
        slaRange.values
      );

  });


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


  fillSelect(
    "statusNew",
    STATE.statuses,
    "Оберіть статус"
  );

}


// ============================================================
// FIRST COLUMN
// ============================================================

function firstColumn(
  rows
) {

  return rows
    .map(row =>
      String(
        row[0] ?? ""
      ).trim()
    )
    .filter(value =>
      value !== ""
    );

}


// ============================================================
// LOAD DASHBOARD
// ============================================================

async function loadDashboard() {

  showLoading();


  try {

    const result =
      await Excel.run(
        async context => {


          const base =
            context.workbook.tables.getItem(
              "tbl_RBD_Base"
            );


          const archive =
            context.workbook.tables.getItem(
              "tbl_RBD_Archive"
            );


          base.rows.load(
            "items"
          );


          archive.rows.load(
            "items"
          );


          const baseHeaders =
            base.getHeaderRowRange();


          const archiveHeaders =
            archive.getHeaderRowRange();


          baseHeaders.load(
            "values"
          );


          archiveHeaders.load(
            "values"
          );


          await context.sync();


          let baseValues = [];

          let archiveValues = [];


          if (
            base.rows.items.length > 0
          ) {

            const baseBody =
              base.getDataBodyRange();


            baseBody.load(
              "values"
            );


            await context.sync();


            baseValues =
              baseBody.values;

          }


          if (
            archive.rows.items.length > 0
          ) {

            const archiveBody =
              archive.getDataBodyRange();


            archiveBody.load(
              "values"
            );


            await context.sync();


            archiveValues =
              archiveBody.values;

          }


          return {

            baseHeaders:
              baseHeaders.values[0],

            baseValues,

            archiveHeaders:
              archiveHeaders.values[0],

            archiveValues

          };

        }
      );


    STATE.requests =
      result.baseValues.map(
        row =>
          rowToRequest(
            result.baseHeaders,
            row
          )
      );


    STATE.archive =
      result.archiveValues.map(
        row =>
          rowToRequest(
            result.archiveHeaders,
            row
          )
      );


    updateKpis();

    applyFilters();

    updateLastUpdated();

  }

  catch (error) {

    showTableError(
      getErrorText(error)
    );

  }

}


// ============================================================
// ROW TO REQUEST
// ============================================================

function rowToRequest(
  headers,
  row
) {

  function get(
    columnName
  ) {

    const index =
      headers.indexOf(
        columnName
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
      get(
        "Сума витрат"
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
      Number(
        get(
          "SLA, год"
        )
      ) || 0,

    timeInStatus:
      Number(
        get(
          "Час у статусі, год"
        )
      ) || 0,

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
// FILTERS
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
          request.status !== status
        ) {

          return false;

        }


        if (
          executor &&
          request.executor !== executor
        ) {

          return false;

        }


        if (
          city &&
          request.city !== city
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

}


// ============================================================
// KPI
// ============================================================

function updateKpis() {

  const active =
    STATE.requests.length;


  const overdue =
    STATE.requests.filter(
      isOverdue
    ).length;


  const work =
    STATE.requests.filter(
      request => {

        return (
          request.status !== "Нова" &&
          request.status !== "Призупинена" &&
          request.status !== "Закрита"
        );

      }
    ).length;


  const now =
    Date.now();


  const thirtyDays =
    30 *
    24 *
    60 *
    60 *
    1000;


  const closed =
    STATE.archive.filter(
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
    ).length;


  setText(
    "kpiActive",
    active
  );


  setText(
    "kpiOverdue",
    overdue
  );


  setText(
    "kpiWork",
    work
  );


  setText(
    "kpiClosed",
    closed
  );

}


// ============================================================
// OVERDUE
// ============================================================

function isOverdue(
  request
) {

  return (

    request.slaOverdue ===
      "Так" ||

    request.dateOverdue ===
      "Так"

  );

}


// ============================================================
// RENDER
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
    STATE.filtered.length === 0
  ) {

    body.innerHTML = `

      <tr>

        <td colspan="11">

          <div class="empty-state">

            Заявок за обраними умовами немає

          </div>

        </td>

      </tr>

    `;


    updateSummary();

    return;

  }


  body.innerHTML =
    STATE.filtered
      .map(
        request =>
          buildRow(
            request
          )
      )
      .join("");


  bindRowActions();

  updateSummary();

}


// ============================================================
// BUILD ROW
// ============================================================

function buildRow(
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


  const slaText =
    request.sla > 0
      ? request.sla +
        " год."
      : "—";


  const slaClass =
    request.slaOverdue ===
      "Так"
      ? "sla-overdue"
      : "";


  const plannedClass =
    request.dateOverdue ===
      "Так"
      ? "planned-date"
      : "";


  return `

    <tr class="${rowClass}">


      <td>

        <div class="request-id">

          ${escapeHtml(request.id)}

        </div>

        ${
          request.sd
            ? `
              <div style="
                margin-top:3px;
                font-size:8px;
                opacity:.7;
              ">
                ${escapeHtml(request.sd)}
              </div>
            `
            : ""
        }

      </td>


      <td>

        <div
          class="text-ellipsis"
          title="${escapeHtml(request.category)}"
        >
          ${escapeHtml(request.category)}
        </div>

      </td>


      <td>

        <div
          class="text-ellipsis"
          title="${escapeHtml(request.description)}"
        >
          ${escapeHtml(request.description)}
        </div>

      </td>


      <td>

        ${escapeHtml(request.city)}

      </td>


      <td>

        <div
          class="text-ellipsis"
          title="${escapeHtml(request.customer)}"
        >
          ${escapeHtml(request.customer)}
        </div>

      </td>


      <td>

        <div
          class="text-ellipsis"
          title="${escapeHtml(request.executor)}"
        >
          ${
            request.executor
              ? escapeHtml(
                  request.executor
                )
              : "—"
          }
        </div>

      </td>


      <td>

        <div
          class="status-badge ${statusClass}"
        >

          <span class="status-dot"></span>

          ${escapeHtml(request.status)}

        </div>

      </td>


      <td class="${plannedClass}">

        ${formatExcelDate(request.plannedDate)}

      </td>


      <td class="${slaClass}">

        ${slaText}

      </td>


      <td>

        <div
          class="text-ellipsis"
          title="${escapeHtml(request.comment)}"
        >
          ${escapeHtml(request.comment || "—")}
        </div>

      </td>


      <td>

        <div class="row-actions">


          <button
            class="action-button"
            data-action="view"
            data-id="${escapeHtml(request.id)}"
            title="Переглянути"
            type="button"
          >
            ◉
          </button>


          <button
            class="action-button"
            data-action="edit"
            data-id="${escapeHtml(request.id)}"
            title="Редагувати"
            type="button"
          >
            ✎
          </button>


          <button
            class="action-button process"
            data-action="process"
            data-id="${escapeHtml(request.id)}"
            title="Опрацювати"
            type="button"
          >
            ▶
          </button>


        </div>

      </td>


    </tr>

  `;

}


// ============================================================
// ROW COLOR
// ============================================================

function getRowClass(
  request
) {

  if (
    isOverdue(request)
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
// STATUS CLASS
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
// ACTIONS
// ============================================================

function bindRowActions() {

  document
    .querySelectorAll(
      "[data-action]"
    )
    .forEach(
      button => {

        button.addEventListener(
          "click",
          () => {


            const request =
              STATE.requests.find(
                item =>
                  item.id ===
                  button.dataset.id
              );


            if (!request) {
              return;
            }


            const action =
              button.dataset.action;


            if (
              action === "view"
            ) {

              openDetails(
                request
              );

            }


            if (
              action === "edit"
            ) {

              openEdit(
                request
              );

            }


            if (
              action === "process"
            ) {

              openStatus(
                request
              );

            }

          }
        );

      }
    );

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
      "Адреса / локація",
      request.address
    ],

    [
      "Дата створення",
      formatExcelDateTime(
        request.created
      )
    ],

    [
      "Сума витрат",
      formatMoney(
        request.amount
      )
    ],

    [
      "Замовник",
      request.customer
    ],

    [
      "Статус",
      request.status
    ],

    [
      "Виконавець",
      request.executor
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
      "SLA",
      request.sla
        ? request.sla +
          " год."
        : "—"
    ],

    [
      "Прострочено SLA",
      request.slaOverdue ||
      "Ні"
    ],

    [
      "Планова дата",
      formatExcelDate(
        request.plannedDate
      )
    ],

    [
      "Прострочено по даті",
      request.dateOverdue ||
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

              ${escapeHtml(row[0])}

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
// STATUS MODAL
// ============================================================

function openStatus(
  request
) {

  STATE.selectedRequest =
    request;


  setText(
    "statusTitle",
    "Опрацювання " +
    request.id
  );


  document
    .getElementById(
      "statusCurrent"
    )
    .value =
      request.status;


  document
    .getElementById(
      "statusNew"
    )
    .value =
      "";


  document
    .getElementById(
      "pauseReason"
    )
    .value =
      "";


  document
    .getElementById(
      "statusComment"
    )
    .value =
      "";


  clearMessage(
    "statusMessage"
  );


  handlePauseReason();


  openModal(
    "statusModal"
  );

}


// ============================================================
// PAUSE
// ============================================================

function handlePauseReason() {

  const status =
    valueOf(
      "statusNew"
    );


  const block =
    document.getElementById(
      "pauseReasonBlock"
    );


  if (!block) {
    return;
  }


  block.style.display =
    status ===
      "Призупинена"
      ? "block"
      : "none";

}


// ============================================================
// SAVE STATUS
// ============================================================

async function saveStatusChange() {

  const request =
    STATE.selectedRequest;


  if (!request) {
    return;
  }


  const newStatus =
    valueOf(
      "statusNew"
    );


  const comment =
    valueOf(
      "statusComment"
    );


  const pauseReason =
    valueOf(
      "pauseReason"
    );


  if (!newStatus) {

    showError(
      "statusMessage",
      "Оберіть новий статус."
    );

    return;

  }


  if (
    newStatus ===
    request.status
  ) {

    showError(
      "statusMessage",
      "Новий статус збігається з поточним."
    );

    return;

  }


  if (
    newStatus ===
      "Призупинена" &&
    !pauseReason
  ) {

    showError(
      "statusMessage",
      "Вкажіть причину призупинення."
    );

    return;

  }


  setButtonBusy(
    "saveStatusButton",
    true,
    "Передаємо..."
  );


  try {

    await enqueue({

      eventId:
        createId(
          "EVT"
        ),

      requestKey: "",

      requestId:
        request.id,

      operation:
        "STATUS_CHANGE",

      actor:
        "Керівник",

      source:
        "Керівник",

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

      newExecutor: ""

    });


    showSuccess(
      "statusMessage",
      "Зміну статусу передано в чергу."
    );

  }

  catch (error) {

    showError(
      "statusMessage",
      getErrorText(
        error
      )
    );

  }

  finally {

    setButtonBusy(
      "saveStatusButton",
      false,
      "Змінити статус"
    );

  }

}


// ============================================================
// CREATE
// ============================================================

function openCreateModal() {

  clearCreateForm();

  openModal(
    "createModal"
  );

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
    valueOf(
      "createExecutor"
    );


  if (
    !category ||
    !description ||
    !city ||
    !customer ||
    !executor
  ) {

    showError(
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

      requestId: "",

      operation:
        "CREATE",

      actor:
        "Керівник",

      source:
        "Керівник",

      executor,

      expectedStatus: "",

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

      pauseReason: "",

      newExecutor: ""

    });


    showSuccess(
      "createMessage",
      "Заявку передано в чергу."
    );

  }

  catch (error) {

    showError(
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


  document
    .getElementById(
      "editSD"
    )
    .value =
      request.sd;


  document
    .getElementById(
      "editCategory"
    )
    .value =
      request.category;


  document
    .getElementById(
      "editDescription"
    )
    .value =
      request.description;


  document
    .getElementById(
      "editCity"
    )
    .value =
      request.city;


  document
    .getElementById(
      "editAddress"
    )
    .value =
      request.address;


  document
    .getElementById(
      "editCustomer"
    )
    .value =
      request.customer;


  document
    .getElementById(
      "editAmount"
    )
    .value =
      request.amount ||
      "";


  document
    .getElementById(
      "editPlannedDate"
    )
    .value =
      excelSerialToInputDate(
        request.plannedDate
      );


  document
    .getElementById(
      "editComment"
    )
    .value =
      request.comment;


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

      requestKey: "",

      requestId:
        request.id,

      operation:
        "UPDATE",

      actor:
        "Керівник",

      source:
        "Керівник",

      executor:
        request.executor,

      expectedStatus: "",

      newStatus: "",

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

      pauseReason: "",

      newExecutor: ""

    });


    showSuccess(
      "editMessage",
      "Зміни передано в чергу."
    );

  }

  catch (error) {

    showError(
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
      "Зберегти зміни"
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
        context.workbook.tables.getItem(
          "tbl_RBD_Queue"
        );


      const row = [

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

      ];


      queue.rows.add(
        null,
        [row]
      );


      await context.sync();

    }
  );

}


// ============================================================
// SUMMARY
// ============================================================

function updateSummary() {

  const text =
    "Показано " +
    STATE.filtered.length +
    " із " +
    STATE.requests.length +
    " заявок";


  setText(
    "tableSummary",
    text
  );


  setText(
    "headerSummary",
    text
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

      <td colspan="11">

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

      <td colspan="11">

        <div class="empty-state">

          Помилка: ${escapeHtml(message)}

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


  const text =
    "↻ Дані оновлено: " +
    String(
      now.getDate()
    ).padStart(
      2,
      "0"
    ) +
    "." +
    String(
      now.getMonth() + 1
    ).padStart(
      2,
      "0"
    ) +
    "." +
    now.getFullYear() +
    " " +
    String(
      now.getHours()
    ).padStart(
      2,
      "0"
    ) +
    ":" +
    String(
      now.getMinutes()
    ).padStart(
      2,
      "0"
    );


  setText(
    "lastUpdated",
    text
  );

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

  const value =
    valueOf(id)
      .replace(
        ",",
        "."
      );


  if (!value) {

    return 0;

  }


  const number =
    Number(
      value
    );


  return Number.isFinite(
    number
  )
    ? number
    : 0;

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
// CREATE ID
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
// LOCAL TIME
// ============================================================

function localTimestamp() {

  const date =
    new Date();


  const pad =
    value =>
      String(
        value
      ).padStart(
        2,
        "0"
      );


  return (

    date.getFullYear() +
    "-" +

    pad(
      date.getMonth() + 1
    ) +
    "-" +

    pad(
      date.getDate()
    ) +
    "T" +

    pad(
      date.getHours()
    ) +
    ":" +

    pad(
      date.getMinutes()
    ) +
    ":" +

    pad(
      date.getSeconds()
    )

  );

}


// ============================================================
// EXCEL DATE
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

      String(
        date.getUTCDate()
      ).padStart(
        2,
        "0"
      ) +

      "." +

      String(
        date.getUTCMonth() + 1
      ).padStart(
        2,
        "0"
      ) +

      "." +

      date.getUTCFullYear()

    );

  }


  const stringValue =
    String(value);


  const match =
    stringValue.match(
      /^(\d{4})-(\d{2})-(\d{2})/
    );


  if (match) {

    return (
      match[3] +
      "." +
      match[2] +
      "." +
      match[1]
    );

  }


  return stringValue;

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

    String(
      date.getUTCDate()
    ).padStart(
      2,
      "0"
    ) +

    "." +

    String(
      date.getUTCMonth() + 1
    ).padStart(
      2,
      "0"
    ) +

    "." +

    date.getUTCFullYear() +

    " " +

    String(
      date.getUTCHours()
    ).padStart(
      2,
      "0"
    ) +

    ":" +

    String(
      date.getUTCMinutes()
    ).padStart(
      2,
      "0"
    )

  );

}


// ============================================================

function excelSerialToDate(
  serial
) {

  return new Date(

    Math.round(
      (
        serial -
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

    String(
      date.getUTCMonth() + 1
    ).padStart(
      2,
      "0"
    ) +

    "-" +

    String(
      date.getUTCDate()
    ).padStart(
      2,
      "0"
    )

  );

}


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
// MONEY
// ============================================================

function formatMoney(
  value
) {

  const number =
    Number(
      value
    );


  if (
    !Number.isFinite(
      number
    )
  ) {

    return "—";

  }


  return (
    number.toLocaleString(
      "uk-UA",
      {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      }
    ) +
    " грн"
  );

}


// ============================================================
// TEXT
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
      String(value);

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

function showSuccess(
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

function showError(
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
  caption
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
    caption;

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
